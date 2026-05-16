const express = require("express");
const session = require("express-session");
const rateLimit = require("express-rate-limit");
const multer = require("multer");
const initSqlJs = require("sql.js");
const bcrypt = require("bcryptjs");
const XLSX = require("xlsx");
const { v4: uuidv4 } = require("uuid");
const AdmZip = require("adm-zip");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const zlib = require("zlib");

const app = express();
const PORT = process.env.PORT || 5000;
const DB_PATH = path.join(__dirname, "database.db");
const MODEL_PATH = path.join(__dirname, "modele_tva.xlsx");

const jobs = new Map();

const SECRET_KEY_FILE = path.join(__dirname, ".secret_key");
let SECRET_KEY;
if (fs.existsSync(SECRET_KEY_FILE)) {
  SECRET_KEY = fs.readFileSync(SECRET_KEY_FILE, "utf8").trim();
} else {
  SECRET_KEY = crypto.randomBytes(32).toString("hex");
  fs.writeFileSync(SECRET_KEY_FILE, SECRET_KEY);
}

const GA_ID = process.env.GA_ID || "";
const SITE_URL = "https://tva2xml.ma";

// --- Express setup ---
app.set("trust proxy", 1);
app.use(express.urlencoded({ extended: true }));
const FileStore = require("session-file-store")(session);
app.use(
  session({
    store: new FileStore({
      path: path.join(__dirname, "sessions"),
      reapInterval: 3600,
      ttl: 86400,
    }),
    secret: SECRET_KEY,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: process.env.SESSION_COOKIE_SECURE === "1",
      maxAge: 24 * 60 * 60 * 1000,
    },
  })
);

const limiter = rateLimit({
  windowMs: 60 * 1000,
  max: 10,
  keyGenerator: (req) => req.ip,
  skip: () => process.env.NODE_ENV === "test",
});

const upload = multer({
  dest: path.join(__dirname, "tmp_uploads"),
  fileFilter: (req, file) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === ".xls" || ext === ".xlsx") return true;
    throw new Error("Format accepté: .xls ou .xlsx");
  },
  limits: { fileSize: 10 * 1024 * 1024 },
});

if (!fs.existsSync(path.join(__dirname, "sessions")))
  fs.mkdirSync(path.join(__dirname, "sessions"), { recursive: true });
if (!fs.existsSync(path.join(__dirname, "tmp_uploads")))
  fs.mkdirSync(path.join(__dirname, "tmp_uploads"));
if (!fs.existsSync(path.join(__dirname, "tmp_xml")))
  fs.mkdirSync(path.join(__dirname, "tmp_xml"));

// --- OG image generator ---
function makeOGPng() {
  const W = 1200,
    H = 630;
  const raw = Buffer.alloc(W * H * 3 + H);
  let off = 0;
  for (let y = 0; y < H; y++) {
    raw[off++] = 0;
    const t2 = (y / H) * 0.18;
    for (let x = 0; x < W; x++) {
      const t = x / W;
      const r = Math.round((0x66 + (0x76 - 0x66) * t) * (1 - t2));
      const g = Math.round((0x7e + (0x4b - 0x7e) * t) * (1 - t2));
      const b = Math.round((0xea + (0xa2 - 0xea) * t) * (1 - t2));
      raw[off++] = r;
      raw[off++] = g;
      raw[off++] = b;
    }
  }
  const crc32 = (b) => {
    let c = 0xffffffff;
    for (let i = 0; i < b.length; i++) {
      c ^= b[i];
      for (let j = 0; j < 8; j++) c = c >>> 1 ^ (c & 1 ? 0xedb88320 : 0);
    }
    return (c ^ 0xffffffff) >>> 0;
  };
  const ch = (t, d) => {
    const l = Buffer.alloc(4);
    l.writeUInt32BE(d.length);
    const tb = Buffer.from(t, "ascii");
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(Buffer.concat([tb, d])));
    return Buffer.concat([l, tb, d, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(W, 0);
  ihdr.writeUInt32BE(H, 4);
  ihdr[8] = 8; ihdr[9] = 2; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    ch("IHDR", ihdr),
    ch("IDAT", zlib.deflateSync(raw)),
    ch("IEND", Buffer.alloc(0)),
  ]);
}
const OG_PNG = makeOGPng();

// --- DB init (sql.js wrapper) ---
let db;
let _dbReady = false;

function dbGet(sql, params) {
  const stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  const row = stmt.step() ? stmt.getAsObject() : null;
  stmt.free();
  return row;
}

function dbAll(sql, params) {
  const stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
}

let _dbDirty = false;
let _dbFlushTimer = null;

function dbRun(sql, params) {
  const stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  stmt.step();
  stmt.free();
  _dbDirty = true;
  if (!_dbFlushTimer) _dbFlushTimer = setTimeout(flushDb, 3000);
}

function flushDb() {
  _dbFlushTimer = null;
  if (!_dbDirty) return;
  _dbDirty = false;
  const data = db.export();
  fs.writeFileSync(DB_PATH, Buffer.from(data));
}

function dbRunSync(sql, params) {
  if (_dbFlushTimer) { clearTimeout(_dbFlushTimer); _dbFlushTimer = null; }
  const stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  stmt.step();
  stmt.free();
  const data = db.export();
  fs.writeFileSync(DB_PATH, Buffer.from(data));
  _dbDirty = false;
}

let _dbInit = null;

async function initDb() {
  if (_dbInit) return _dbInit;
  _dbInit = initDbInner();
  return _dbInit;
}

async function initDbInner() {
  const SQL = await initSqlJs({ locateFile: file => path.join(__dirname, "node_modules", "sql.js", "dist", file) });
  if (fs.existsSync(DB_PATH)) {
    const buf = fs.readFileSync(DB_PATH);
    db = new SQL.Database(buf);
  } else {
    db = new SQL.Database();
  }
  // db.run("PRAGMA journal_mode=WAL"); // non supporté par sql.js
  db.run(`CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    name TEXT DEFAULT '',
    subscription_start TEXT,
    subscription_end TEXT,
    active INTEGER DEFAULT 1,
    is_admin INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now'))
  )`);
  db.run(`CREATE TABLE IF NOT EXISTS downloads (
    file_id TEXT PRIMARY KEY,
    zip_path TEXT NOT NULL,
    zip_name TEXT NOT NULL,
    user_id INTEGER,
    created_at TEXT DEFAULT (datetime('now'))
  )`);
  db.run("CREATE INDEX IF NOT EXISTS idx_users_email ON users(email)");
  db.run("CREATE INDEX IF NOT EXISTS idx_downloads_user_id ON downloads(user_id)");
  const row = dbGet("SELECT COUNT(*) as c FROM users");
  if (row.c === 0) {
    const hashed = bcrypt.hashSync("admin", 10);
    const today = new Date().toISOString().split("T")[0];
    const future = new Date(Date.now() + 365 * 10 * 86400000).toISOString().split("T")[0];
    dbRunSync("INSERT INTO users (email,password,name,is_admin,subscription_start,subscription_end,active) VALUES (?,?,?,1,?,?,1)",
      ["admin@tva2xml.ma", hashed, "Admin", today, future]);
  }
  _dbReady = true;
}

// --- Helpers ---
function getUser(id) {
  const r = dbGet("SELECT * FROM users WHERE id=?", [id]);
  if (!r) return null;
  return {
    id: r.id,
    email: r.email,
    password: r.password,
    name: r.name,
    subscription_start: r.subscription_start,
    subscription_end: r.subscription_end,
    active: r.active,
    is_admin: r.is_admin,
  };
}

function getUserByEmail(email) {
  const r = dbGet("SELECT * FROM users WHERE email=?", [email]);
  if (!r) return null;
  return {
    id: r.id,
    email: r.email,
    password: r.password,
    name: r.name,
    subscription_start: r.subscription_start,
    subscription_end: r.subscription_end,
    active: r.active,
    is_admin: r.is_admin,
  };
}

async function createUser(email, password, name) {
  const hashed = await bcrypt.hash(password, 10);
  dbRunSync("INSERT INTO users (email,password,name) VALUES (?,?,?)", [email, hashed, name]);
}

function subscriptionActive(user) {
  if (!user || !user.active) return false;
  if (!user.subscription_end) return false;
  const today = new Date().toISOString().split("T")[0];
  return user.subscription_end >= today;
}

function getAllUsers(page = 1, perPage = 50) {
  const offset = (page - 1) * perPage;
  return dbAll(
    "SELECT * FROM users ORDER BY created_at DESC LIMIT ? OFFSET ?",
    [perPage, offset]
  ).map((r) => ({
      id: r.id,
      email: r.email,
      name: r.name,
      subscription_start: r.subscription_start,
      subscription_end: r.subscription_end,
      active: r.active,
      is_admin: r.is_admin,
      created_at: r.created_at,
    }));
}

function countUsers() {
  return dbGet("SELECT COUNT(*) as c FROM users").c;
}

function isSafeUrl(target) {
  if (!target) return false;
  try {
    const url = new URL(target, SITE_URL);
    return url.hostname === "localhost" || url.hostname === "127.0.0.1" || !url.hostname;
  } catch {
    return false;
  }
}

function safeNext(target, def) {
  if (target && isSafeUrl(target)) return target;
  return def;
}

// --- CSRF ---
const csrfTokens = new Map();
setInterval(() => {
  const expiry = Date.now() - 3600000;
  for (const [k, v] of csrfTokens) if (v < expiry) csrfTokens.delete(k);
}, 600000);

function generateCsrf(req, res, next) {
  if (!req.session) req.session = {};
  if (!req.session.csrfToken) {
    req.session.csrfToken = crypto.randomBytes(32).toString("hex");
  }
  csrfTokens.set(req.session.csrfToken, Date.now());
  next();
}

function validateCsrf(req, res, next) {
  if (req.method === "POST") {
    const token = req.body && req.body.csrf_token;
    const valid = token && csrfTokens.has(token);
    if (!valid) {
      return res.status(403).send("CSRF invalide");
    }
    csrfTokens.delete(token);
  }
  next();
}

app.use(generateCsrf);

// --- Constants ---
const EXPECTED_HEADERS = [
  "NUMERO","DATEFACT","NUMFACT","IDENTIF","FOURNISSEUR","NATURE",
  "HT","TAUX","PRORATA","TVA","TTC","DATE_REG","MODE_REG","ICE",
];
const NUM_COLS = [6, 7, 8, 9, 10];
const MODE_MAP = {
  ESPECE: 1, ESPECES: 1, CHEQUE: 2,
  PRELEVEMENT: 3, VIREMENT: 4,
  EFFET: 5, COMPENSATION: 6, AUTRES: 7,
};
const MONTH_NAMES = [
  "JANVIER","FEVRIER","MARS","AVRIL","MAI","JUIN",
  "JUILLET","AOUT","SEPTEMBRE","OCTOBRE","NOVEMBRE","DECEMBRE",
];
const TRIM_NAMES = [
  "1ER_TRIMESTRE","2E_TRIMESTRE","3E_TRIMESTRE","4E_TRIMESTRE",
];

// --- Excel helpers ---
function excelSerialToDate(val) {
  if (val instanceof Date && !isNaN(val))
    return val.toISOString().split("T")[0];
  if (typeof val === "number" && val >= 0)
    return new Date((val - 25569) * 86400000).toISOString().split("T")[0];
  if (typeof val === "string" && val.trim()) {
    const f = [
      /^(\d{2})\/(\d{2})\/(\d{4})$/,
      /^(\d{4})-(\d{2})-(\d{2})$/,
      /^(\d{2})\/(\d{2})\/(\d{2})$/,
    ];
    for (const re of f) {
      const m = val.trim().match(re);
      if (m) {
        let y = m[3];
        if (y.length === 2) y = "20" + y;
        return `${y}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
      }
    }
  }
  return "";
}

function openWorkbook(filepath) {
  const wb = XLSX.read(fs.readFileSync(filepath), { type: "buffer", cellDates: true, cellText: false });
  const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1,
    defval: null,
  });
  return data;
}

function txt(v) {
  if (v == null) return "";
  if (typeof v === "number")
    return Number.isInteger(v) ? String(v) : String(v);
  return String(v).trim();
}

function xmlsafe(v) {
  return txt(v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function safeFloat(v, def) {
  const d = def || 0.0;
  try {
    const val = parseFloat(v || 0);
    return isNaN(val) || !isFinite(val) ? d : val;
  } catch {
    return d;
  }
}

// --- Process & Validate ---
function processFile(filepath, societe, idfiscal, annee, periode, regime) {
  const rows = openWorkbook(filepath);
  const lines = [];
  lines.push('<?xml version="1.0" encoding="UTF-8"?>');
  lines.push(
    `<DeclarationReleveDeduction> <identifiantFiscal>${idfiscal}</identifiantFiscal><annee>${annee}</annee><periode>${periode}</periode><regime>${regime}</regime><releveDeductions>`
  );
  for (let i = 1; i < rows.length; i++) {
    const rd = rows[i];
    let ord = 0;
    try {
      ord = parseInt(rd[0]);
    } catch {
      ord = 0;
    }
    if (isNaN(ord)) ord = 0;
    const datefac = excelSerialToDate(rd[1]);
    const numfact = txt(rd[2]);
    const identif = txt(rd[3]);
    const nom = xmlsafe(rd[4]).toUpperCase();
    const des = xmlsafe(rd[5]).toUpperCase();
    const ht = safeFloat(rd[6]);
    let taux = safeFloat(rd[7]);
    if (taux <= 1) taux *= 100;
    let prorata = safeFloat(rd[8]);
    if (prorata <= 1) prorata *= 100;
    const tva = safeFloat(rd[9]);
    const ttc = safeFloat(rd[10]);
    const dpai = excelSerialToDate(rd[11]);
    const modeRaw = txt(rd[12])
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase();
    const mpId = MODE_MAP[modeRaw] || 8;
    const ice = txt(rd[13]);
    lines.push(
      ` <rd> <ord>${ord}</ord><num>${numfact}</num>` +
        `<des>${des}</des>` +
        `<mht>${ht.toFixed(2).padStart(10)}</mht><tva>${tva.toFixed(2).padStart(10)}</tva><ttc>${ttc.toFixed(2).padStart(10)}</ttc>` +
        `<refF> <if>${identif}</if><nom>${nom}</nom><ice>${ice}</ice></refF>` +
        `<tx>${(taux * 100).toFixed(2).padStart(10)}</tx><prorata>${prorata.toFixed(2).padStart(10)}</prorata>` +
        `<mp> <id>${mpId}</id></mp><dpai>${dpai}</dpai><dfac>${datefac}</dfac></rd>`
    );
  }
  lines.push("</releveDeductions></DeclarationReleveDeduction>");
  const outPath = path.join(
    __dirname,
    "tmp_xml",
    uuidv4() + ".xml"
  );
  fs.writeFileSync(outPath, lines.join("\n"), "utf8");
  return outPath;
}

function validate(filepath) {
  const errors = [];
  const warns = [];
  let rows;
  try {
    rows = openWorkbook(filepath);
  } catch (e) {
    return [[`Fichier invalide: ${e.message}`], []];
  }
  if (rows.length < 2) return [["Le fichier ne contient aucune ligne de données"], []];
  if (rows[0].length < 14)
    return [[`La première feuille a ${rows[0].length} colonnes au lieu de 14`], []];
  const hdrs = [];
  for (let c = 0; c < 14; c++) hdrs.push(txt(rows[0][c]).toUpperCase());
  for (let i = 0; i < 14; i++) {
    if (hdrs[i] !== EXPECTED_HEADERS[i])
      errors.push(
        `En-tête colonne ${i + 1}: attendu «${EXPECTED_HEADERS[i]}», trouvé «${hdrs[i]}»`
      );
  }
  if (errors.length) return [errors, []];
  for (let r = 1; r < rows.length; r++) {
    const rn = r + 1;
    for (let c = 0; c < 14; c++) {
      const val = rows[r][c];
      const cn = EXPECTED_HEADERS[c];
      if (NUM_COLS.includes(c)) {
        if (val == null || (typeof val === "string" && val.trim() === ""))
          errors.push(`Ligne ${rn}, colonne ${cn}: vide`);
        else if (typeof val === "string") {
          if (isNaN(parseFloat(val.replace(",", "."))))
            errors.push(`Ligne ${rn}, colonne ${cn}: «${val}» n'est pas un nombre`);
        }
      } else if ([0, 1, 11].includes(c)) {
        if (val == null || (typeof val === "string" && val.trim() === ""))
          errors.push(`Ligne ${rn}, ${cn}: vide`);
      } else if ([2, 3, 4, 5, 12, 13].includes(c)) {
        if (val == null || (typeof val === "string" && val.trim() === ""))
          errors.push(`Ligne ${rn}, ${cn}: vide`);
      }
    }
  }
  for (let r = 1; r < rows.length; r++) {
    const rn = r + 1;
    for (const c of [2, 3, 4, 5, 13]) {
      const val = txt(rows[r][c]);
      if (val.includes("&"))
        warns.push(`Ligne ${rn}, ${EXPECTED_HEADERS[c]}: contient «&» (sera converti en &amp;)`);
      if (val.includes("<"))
        warns.push(`Ligne ${rn}, ${EXPECTED_HEADERS[c]}: contient «<» (sera converti en &lt;)`);
      if (val.includes(">"))
        warns.push(`Ligne ${rn}, ${EXPECTED_HEADERS[c]}: contient «>» (sera converti en &gt;)`);
    }
  }
  return [errors, warns];
}

// --- Auth middleware ---
function loginRequired(req, res, next) {
  if (!req.session.user_id) return res.redirect("/login?next=" + req.path);
  next();
}

function subscriptionRequired(req, res, next) {
  if (!req.user) return res.redirect("/login?next=" + req.path);
  if (!subscriptionActive(req.user)) return res.redirect("/profile?msg=expired");
  next();
}

function adminRequired(req, res, next) {
  if (!req.user || !req.user.is_admin) return res.redirect("/");
  next();
}

// --- Render helpers ---
function gaBlock() {
  if (!GA_ID) return "";
  return (
    `<link rel="dns-prefetch" href="https://www.googletagmanager.com">` +
    `<link rel="preconnect" href="https://www.googletagmanager.com">` +
    `<script async src="https://www.googletagmanager.com/gtag/js?id=${GA_ID}"></script>` +
    `<script>window.dataLayer=window.dataLayer||[];function gtag(){dataLayer.push(arguments)};gtag('js',new Date());gtag('config','${GA_ID}');</script>`
  );
}

function userBar(user) {
  if (user) {
    let html = `<span style="color:#6b7280;margin-right:4px">${user.email}</span>
<a href="/profile" style="color:#4a6cf7">Mon compte</a>`;
    if (user.is_admin)
      html += `<a href="/admin" style="color:#059669">Admin</a>`;
    html += `<a href="/logout" style="color:#ef4444">Déconnexion</a>`;
    return html;
  }
  return `<a class="login-link" href="/login">Se connecter</a>`;
}

function userBarStyle() {
  return `position:absolute;top:20px;right:28px;display:flex;align-items:center;gap:4px;font-size:.82rem`;
}

function userBarLinkStyle() {
  return `text-decoration:none;font-weight:600`;
}

// --- Security headers ---
app.use((req, res, next) => {
  res.setHeader(
    "Strict-Transport-Security",
    "max-age=31536000; includeSubDomains"
  );
  next();
});

const LOADING_HTML = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Chargement...</title><style>body{margin:0;font-family:Inter,system-ui,sans-serif;background:linear-gradient(135deg,#0f0c29,#302b63,#24243e);min-height:100vh;display:flex;align-items:center;justify-content:center;color:#fff;text-align:center}.card{background:rgba(255,255,255,.97);border-radius:20px;padding:48px;max-width:440px;box-shadow:0 25px 60px rgba(0,0,0,.5)}.spinner{width:48px;height:48px;border:4px solid #e5e7eb;border-top-color:#667eea;border-radius:50%;animation:s 1s linear infinite;margin:0 auto 20px}@keyframes s{to{transform:rotate(360deg)}}p{color:#6b7280;font-size:.9rem;line-height:1.5;margin:0}</style></head><body><div class="card"><div class="spinner"></div><p>L'application démarre, veuillez patienter quelques instants...</p></div></body></html>`;

// --- Load user middleware ---
app.use((req, res, next) => {
  if (!_dbReady) {
    if (['/og-image.png','/robots.txt','/sitemap.xml'].includes(req.path))
      return next();
    return res.status(503).send(LOADING_HTML);
  }
  req.user = req.session.user_id ? getUser(req.session.user_id) : null;
  next();
});

// --- Routes ---
app.get("/", (req, res) => {
  const user = req.user;
  res.send(renderIndex(user, null, null, null, req.session.csrfToken || ""));
});

app.get("/job/:id", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ status: "notfound" });
  res.json(job);
});

app.post("/upload", upload.single("file"), validateCsrf, (req, res) => {
  const user = req.user;
  const ct = req.session.csrfToken || "";
  try {
    if (!req.file) return res.send(renderIndex(user, "Aucun fichier sélectionné", "err", null, ct));
    const filepath = req.file.path;
    const societe = (req.body.societe || "").trim().toUpperCase();
    const idfiscal = (req.body.idfiscal || "").trim();
    const annee = (req.body.annee || "2025").trim();
    const periode = (req.body.periode || "1").trim();
    const perType = (req.body.per_type || "mensuel").trim();
    if (!["mensuel", "trimestriel"].includes(perType))
      return respondError(res, user, "Type de période invalide", ct);
    const regime = perType === "mensuel" ? "1" : "2";
    if (!societe) return respondError(res, user, "Veuillez saisir le nom de la société", ct);
    if (!idfiscal) return respondError(res, user, "Veuillez saisir l'identifiant fiscal", ct);
    if (!/^\d{4}$/.test(annee))
      return respondError(res, user, "Année invalide", ct);
    const anneeI = parseInt(annee);
    if (anneeI < 2020 || anneeI > 2030)
      return respondError(res, user, "Année hors plage (2020-2030)", ct);
    const perI = parseInt(periode);
    if (isNaN(perI))
      return respondError(res, user, "Période invalide", ct);
    if (perType === "trimestriel") {
      if (perI < 1 || perI > 4)
        return respondError(res, user, "Trimestre invalide (1-4)", ct);
    } else {
      if (perI < 1 || perI > 12)
        return respondError(res, user, "Mois invalide (1-12)", ct);
    }
    const [errs, warns] = validate(filepath);
    if (errs.length) {
      try { fs.unlinkSync(filepath); } catch {}
      return res.send(renderIndex(user, null, null, errs, ct));
    }
    
    const jobId = uuidv4();
    jobs.set(jobId, { status: 'queued', progress: 0, createdAt: Date.now() });
    res.send(renderProcessing(user, jobId, societe));
    
    setImmediate(() => {
      try {
        jobs.set(jobId, { status: 'processing', progress: 30 });
        const outPath = processFile(filepath, societe, idfiscal, annee, periode, regime);
        const label = perType === "trimestriel" ? TRIM_NAMES[perI - 1] : MONTH_NAMES[perI - 1];
        const safeName = societe.replace(/[^a-zA-Z0-9_\- ]/g, "").trim().replace(/ /g, "_");
        const xmlContent = fs.readFileSync(outPath, "utf8");
        const nrows = (xmlContent.match(/<rd>/g) || []).length;
        const xmlName = `TVA_${safeName}_${label}_${annee}.xml`;
        const zipName = `TVA_${safeName}_${label}_${annee}.zip`;
        const zipPath = path.join(__dirname, "tmp_xml", uuidv4() + ".zip");
        makeZip(zipPath, xmlName, outPath);
        fs.unlinkSync(outPath);
        const fileId = uuidv4();
        dbRun("INSERT INTO downloads (file_id,zip_path,zip_name,user_id) VALUES (?,?,?,?)",
          [fileId, zipPath, zipName, req.session.user_id || null]);
        jobs.set(jobId, { status: 'done', progress: 100, fileId, zipName, nrows, warns, doneAt: Date.now() });
      } catch (e) {
        console.error("Job error:", e);
        jobs.set(jobId, { status: 'error', error: e.message });
      } finally {
        try { fs.unlinkSync(filepath); } catch {}
      }
    });
  } catch (e) {
    console.error("Upload error:", e);
    if (req.file) try { fs.unlinkSync(req.file.path); } catch {}
    res.send(renderIndex(user, "Une erreur est survenue", "err", null, ct));
  }
}); rateLimit = require("express-rate-limit");
const multer = require("multer");
const initSqlJs = require("sql.js");
const bcrypt = require("bcryptjs");
const XLSX = require("xlsx");
const { v4: uuidv4 } = require("uuid");
const AdmZip = require("adm-zip");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const zlib = require("zlib");

const app = express();
const PORT = process.env.PORT || 5000;
const DB_PATH = path.join(__dirname, "database.db");
const MODEL_PATH = path.join(__dirname, "modele_tva.xlsx");

const jobs = new Map();

const SECRET_KEY_FILE = path.join(__dirname, ".secret_key");
let SECRET_KEY;
if (fs.existsSync(SECRET_KEY_FILE)) {
  SECRET_KEY = fs.readFileSync(SECRET_KEY_FILE, "utf8").trim();
} else {
  SECRET_KEY = crypto.randomBytes(32).toString("hex");
  fs.writeFileSync(SECRET_KEY_FILE, SECRET_KEY);
}

const GA_ID = process.env.GA_ID || "";
const SITE_URL = "https://tva2xml.ma";

// --- Express setup ---
app.set("trust proxy", 1);
app.use(express.urlencoded({ extended: true }));
const FileStore = require("session-file-store")(session);
app.use(
  session({
    store: new FileStore({
      path: path.join(__dirname, "sessions"),
      reapInterval: 3600,
      ttl: 86400,
    }),
    secret: SECRET_KEY,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: process.env.SESSION_COOKIE_SECURE === "1",
      maxAge: 24 * 60 * 60 * 1000,
    },
  })
);

const limiter = rateLimit({
  windowMs: 60 * 1000,
  max: 10,
  keyGenerator: (req) => req.ip,
  skip: () => process.env.NODE_ENV === "test",
});

const upload = multer({
  dest: path.join(__dirname, "tmp_uploads"),
  fileFilter: (req, file) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === ".xls" || ext === ".xlsx") return true;
    throw new Error("Format accepté: .xls ou .xlsx");
  },
  limits: { fileSize: 10 * 1024 * 1024 },
});

if (!fs.existsSync(path.join(__dirname, "sessions")))
  fs.mkdirSync(path.join(__dirname, "sessions"), { recursive: true });
if (!fs.existsSync(path.join(__dirname, "tmp_uploads")))
  fs.mkdirSync(path.join(__dirname, "tmp_uploads"));
if (!fs.existsSync(path.join(__dirname, "tmp_xml")))
  fs.mkdirSync(path.join(__dirname, "tmp_xml"));

// --- OG image generator ---
function makeOGPng() {
  const W = 1200,
    H = 630;
  const raw = Buffer.alloc(W * H * 3 + H);
  let off = 0;
  for (let y = 0; y < H; y++) {
    raw[off++] = 0;
    const t2 = (y / H) * 0.18;
    for (let x = 0; x < W; x++) {
      const t = x / W;
      const r = Math.round((0x66 + (0x76 - 0x66) * t) * (1 - t2));
      const g = Math.round((0x7e + (0x4b - 0x7e) * t) * (1 - t2));
      const b = Math.round((0xea + (0xa2 - 0xea) * t) * (1 - t2));
      raw[off++] = r;
      raw[off++] = g;
      raw[off++] = b;
    }
  }
  const crc32 = (b) => {
    let c = 0xffffffff;
    for (let i = 0; i < b.length; i++) {
      c ^= b[i];
      for (let j = 0; j < 8; j++) c = c >>> 1 ^ (c & 1 ? 0xedb88320 : 0);
    }
    return (c ^ 0xffffffff) >>> 0;
  };
  const ch = (t, d) => {
    const l = Buffer.alloc(4);
    l.writeUInt32BE(d.length);
    const tb = Buffer.from(t, "ascii");
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(Buffer.concat([tb, d])));
    return Buffer.concat([l, tb, d, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(W, 0);
  ihdr.writeUInt32BE(H, 4);
  ihdr[8] = 8; ihdr[9] = 2; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    ch("IHDR", ihdr),
    ch("IDAT", zlib.deflateSync(raw)),
    ch("IEND", Buffer.alloc(0)),
  ]);
}
const OG_PNG = makeOGPng();

// --- DB init (sql.js wrapper) ---
let db;
let _dbReady = false;

function dbGet(sql, params) {
  const stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  const row = stmt.step() ? stmt.getAsObject() : null;
  stmt.free();
  return row;
}

function dbAll(sql, params) {
  const stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
}

let _dbDirty = false;
let _dbFlushTimer = null;

function dbRun(sql, params) {
  const stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  stmt.step();
  stmt.free();
  _dbDirty = true;
  if (!_dbFlushTimer) _dbFlushTimer = setTimeout(flushDb, 3000);
}

function flushDb() {
  _dbFlushTimer = null;
  if (!_dbDirty) return;
  _dbDirty = false;
  const data = db.export();
  fs.writeFileSync(DB_PATH, Buffer.from(data));
}

function dbRunSync(sql, params) {
  if (_dbFlushTimer) { clearTimeout(_dbFlushTimer); _dbFlushTimer = null; }
  const stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  stmt.step();
  stmt.free();
  const data = db.export();
  fs.writeFileSync(DB_PATH, Buffer.from(data));
  _dbDirty = false;
}

let _dbInit = null;

async function initDb() {
  if (_dbInit) return _dbInit;
  _dbInit = initDbInner();
  return _dbInit;
}

async function initDbInner() {
  const SQL = await initSqlJs({ locateFile: file => path.join(__dirname, "node_modules", "sql.js", "dist", file) });
  if (fs.existsSync(DB_PATH)) {
    const buf = fs.readFileSync(DB_PATH);
    db = new SQL.Database(buf);
  } else {
    db = new SQL.Database();
  }
  // db.run("PRAGMA journal_mode=WAL"); // non supporté par sql.js
  db.run(`CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    name TEXT DEFAULT '',
    subscription_start TEXT,
    subscription_end TEXT,
    active INTEGER DEFAULT 1,
    is_admin INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now'))
  )`);
  db.run(`CREATE TABLE IF NOT EXISTS downloads (
    file_id TEXT PRIMARY KEY,
    zip_path TEXT NOT NULL,
    zip_name TEXT NOT NULL,
    user_id INTEGER,
    created_at TEXT DEFAULT (datetime('now'))
  )`);
  db.run("CREATE INDEX IF NOT EXISTS idx_users_email ON users(email)");
  db.run("CREATE INDEX IF NOT EXISTS idx_downloads_user_id ON downloads(user_id)");
  const row = dbGet("SELECT COUNT(*) as c FROM users");
  if (row.c === 0) {
    const hashed = bcrypt.hashSync("admin", 10);
    const today = new Date().toISOString().split("T")[0];
    const future = new Date(Date.now() + 365 * 10 * 86400000).toISOString().split("T")[0];
    dbRunSync("INSERT INTO users (email,password,name,is_admin,subscription_start,subscription_end,active) VALUES (?,?,?,1,?,?,1)",
      ["admin@tva2xml.ma", hashed, "Admin", today, future]);
  }
  _dbReady = true;
}

// --- Helpers ---
function getUser(id) {
  const r = dbGet("SELECT * FROM users WHERE id=?", [id]);
  if (!r) return null;
  return {
    id: r.id,
    email: r.email,
    password: r.password,
    name: r.name,
    subscription_start: r.subscription_start,
    subscription_end: r.subscription_end,
    active: r.active,
    is_admin: r.is_admin,
  };
}

function getUserByEmail(email) {
  const r = dbGet("SELECT * FROM users WHERE email=?", [email]);
  if (!r) return null;
  return {
    id: r.id,
    email: r.email,
    password: r.password,
    name: r.name,
    subscription_start: r.subscription_start,
    subscription_end: r.subscription_end,
    active: r.active,
    is_admin: r.is_admin,
  };
}

async function createUser(email, password, name) {
  const hashed = await bcrypt.hash(password, 10);
  dbRunSync("INSERT INTO users (email,password,name) VALUES (?,?,?)", [email, hashed, name]);
}

function subscriptionActive(user) {
  if (!user || !user.active) return false;
  if (!user.subscription_end) return false;
  const today = new Date().toISOString().split("T")[0];
  return user.subscription_end >= today;
}

function getAllUsers(page = 1, perPage = 50) {
  const offset = (page - 1) * perPage;
  return dbAll(
    "SELECT * FROM users ORDER BY created_at DESC LIMIT ? OFFSET ?",
    [perPage, offset]
  ).map((r) => ({
      id: r.id,
      email: r.email,
      name: r.name,
      subscription_start: r.subscription_start,
      subscription_end: r.subscription_end,
      active: r.active,
      is_admin: r.is_admin,
      created_at: r.created_at,
    }));
}

function countUsers() {
  return dbGet("SELECT COUNT(*) as c FROM users").c;
}

function isSafeUrl(target) {
  if (!target) return false;
  try {
    const url = new URL(target, SITE_URL);
    return url.hostname === "localhost" || url.hostname === "127.0.0.1" || !url.hostname;
  } catch {
    return false;
  }
}

function safeNext(target, def) {
  if (target && isSafeUrl(target)) return target;
  return def;
}

// --- CSRF ---
const csrfTokens = new Map();
setInterval(() => {
  const expiry = Date.now() - 3600000;
  for (const [k, v] of csrfTokens) if (v < expiry) csrfTokens.delete(k);
}, 600000);

function generateCsrf(req, res, next) {
  if (!req.session) req.session = {};
  if (!req.session.csrfToken) {
    req.session.csrfToken = crypto.randomBytes(32).toString("hex");
  }
  csrfTokens.set(req.session.csrfToken, Date.now());
  next();
}

function validateCsrf(req, res, next) {
  if (req.method === "POST") {
    const token = req.body && req.body.csrf_token;
    const valid = token && csrfTokens.has(token);
    if (!valid) {
      return res.status(403).send("CSRF invalide");
    }
    csrfTokens.delete(token);
  }
  next();
}

app.use(generateCsrf);

// --- Constants ---
const EXPECTED_HEADERS = [
  "NUMERO","DATEFACT","NUMFACT","IDENTIF","FOURNISSEUR","NATURE",
  "HT","TAUX","PRORATA","TVA","TTC","DATE_REG","MODE_REG","ICE",
];
const NUM_COLS = [6, 7, 8, 9, 10];
const MODE_MAP = {
  ESPECE: 1, ESPECES: 1, CHEQUE: 2,
  PRELEVEMENT: 3, VIREMENT: 4,
  EFFET: 5, COMPENSATION: 6, AUTRES: 7,
};
const MONTH_NAMES = [
  "JANVIER","FEVRIER","MARS","AVRIL","MAI","JUIN",
  "JUILLET","AOUT","SEPTEMBRE","OCTOBRE","NOVEMBRE","DECEMBRE",
];
const TRIM_NAMES = [
  "1ER_TRIMESTRE","2E_TRIMESTRE","3E_TRIMESTRE","4E_TRIMESTRE",
];

// --- Excel helpers ---
function excelSerialToDate(val) {
  if (val instanceof Date && !isNaN(val))
    return val.toISOString().split("T")[0];
  if (typeof val === "number" && val >= 0)
    return new Date((val - 25569) * 86400000).toISOString().split("T")[0];
  if (typeof val === "string" && val.trim()) {
    const f = [
      /^(\d{2})\/(\d{2})\/(\d{4})$/,
      /^(\d{4})-(\d{2})-(\d{2})$/,
      /^(\d{2})\/(\d{2})\/(\d{2})$/,
    ];
    for (const re of f) {
      const m = val.trim().match(re);
      if (m) {
        let y = m[3];
        if (y.length === 2) y = "20" + y;
        return `${y}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
      }
    }
  }
  return "";
}

function openWorkbook(filepath) {
  const wb = XLSX.read(fs.readFileSync(filepath), { type: "buffer", cellDates: true, cellText: false });
  const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1,
    defval: null,
  });
  return data;
}

function txt(v) {
  if (v == null) return "";
  if (typeof v === "number")
    return Number.isInteger(v) ? String(v) : String(v);
  return String(v).trim();
}

function xmlsafe(v) {
  return txt(v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function safeFloat(v, def) {
  const d = def || 0.0;
  try {
    const val = parseFloat(v || 0);
    return isNaN(val) || !isFinite(val) ? d : val;
  } catch {
    return d;
  }
}

// --- Process & Validate ---
function processFile(filepath, societe, idfiscal, annee, periode, regime) {
  const rows = openWorkbook(filepath);
  const lines = [];
  lines.push('<?xml version="1.0" encoding="UTF-8"?>');
  lines.push(
    `<DeclarationReleveDeduction> <identifiantFiscal>${idfiscal}</identifiantFiscal><annee>${annee}</annee><periode>${periode}</periode><regime>${regime}</regime><releveDeductions>`
  );
  for (let i = 1; i < rows.length; i++) {
    const rd = rows[i];
    let ord = 0;
    try {
      ord = parseInt(rd[0]);
    } catch {
      ord = 0;
    }
    if (isNaN(ord)) ord = 0;
    const datefac = excelSerialToDate(rd[1]);
    const numfact = txt(rd[2]);
    const identif = txt(rd[3]);
    const nom = xmlsafe(rd[4]).toUpperCase();
    const des = xmlsafe(rd[5]).toUpperCase();
    const ht = safeFloat(rd[6]);
    let taux = safeFloat(rd[7]);
    if (taux <= 1) taux *= 100;
    let prorata = safeFloat(rd[8]);
    if (prorata <= 1) prorata *= 100;
    const tva = safeFloat(rd[9]);
    const ttc = safeFloat(rd[10]);
    const dpai = excelSerialToDate(rd[11]);
    const modeRaw = txt(rd[12])
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase();
    const mpId = MODE_MAP[modeRaw] || 8;
    const ice = txt(rd[13]);
    lines.push(
      ` <rd> <ord>${ord}</ord><num>${numfact}</num>` +
        `<des>${des}</des>` +
        `<mht>${ht.toFixed(2).padStart(10)}</mht><tva>${tva.toFixed(2).padStart(10)}</tva><ttc>${ttc.toFixed(2).padStart(10)}</ttc>` +
        `<refF> <if>${identif}</if><nom>${nom}</nom><ice>${ice}</ice></refF>` +
        `<tx>${(taux * 100).toFixed(2).padStart(10)}</tx><prorata>${prorata.toFixed(2).padStart(10)}</prorata>` +
        `<mp> <id>${mpId}</id></mp><dpai>${dpai}</dpai><dfac>${datefac}</dfac></rd>`
    );
  }
  lines.push("</releveDeductions></DeclarationReleveDeduction>");
  const outPath = path.join(
    __dirname,
    "tmp_xml",
    uuidv4() + ".xml"
  );
  fs.writeFileSync(outPath, lines.join("\n"), "utf8");
  return outPath;
}

function validate(filepath) {
  const errors = [];
  const warns = [];
  let rows;
  try {
    rows = openWorkbook(filepath);
  } catch (e) {
    return [[`Fichier invalide: ${e.message}`], []];
  }
  if (rows.length < 2) return [["Le fichier ne contient aucune ligne de données"], []];
  if (rows[0].length < 14)
    return [[`La première feuille a ${rows[0].length} colonnes au lieu de 14`], []];
  const hdrs = [];
  for (let c = 0; c < 14; c++) hdrs.push(txt(rows[0][c]).toUpperCase());
  for (let i = 0; i < 14; i++) {
    if (hdrs[i] !== EXPECTED_HEADERS[i])
      errors.push(
        `En-tête colonne ${i + 1}: attendu «${EXPECTED_HEADERS[i]}», trouvé «${hdrs[i]}»`
      );
  }
  if (errors.length) return [errors, []];
  for (let r = 1; r < rows.length; r++) {
    const rn = r + 1;
    for (let c = 0; c < 14; c++) {
      const val = rows[r][c];
      const cn = EXPECTED_HEADERS[c];
      if (NUM_COLS.includes(c)) {
        if (val == null || (typeof val === "string" && val.trim() === ""))
          errors.push(`Ligne ${rn}, colonne ${cn}: vide`);
        else if (typeof val === "string") {
          if (isNaN(parseFloat(val.replace(",", "."))))
            errors.push(`Ligne ${rn}, colonne ${cn}: «${val}» n'est pas un nombre`);
        }
      } else if ([0, 1, 11].includes(c)) {
        if (val == null || (typeof val === "string" && val.trim() === ""))
          errors.push(`Ligne ${rn}, ${cn}: vide`);
      } else if ([2, 3, 4, 5, 12, 13].includes(c)) {
        if (val == null || (typeof val === "string" && val.trim() === ""))
          errors.push(`Ligne ${rn}, ${cn}: vide`);
      }
    }
  }
  for (let r = 1; r < rows.length; r++) {
    const rn = r + 1;
    for (const c of [2, 3, 4, 5, 13]) {
      const val = txt(rows[r][c]);
      if (val.includes("&"))
        warns.push(`Ligne ${rn}, ${EXPECTED_HEADERS[c]}: contient «&» (sera converti en &amp;)`);
      if (val.includes("<"))
        warns.push(`Ligne ${rn}, ${EXPECTED_HEADERS[c]}: contient «<» (sera converti en &lt;)`);
      if (val.includes(">"))
        warns.push(`Ligne ${rn}, ${EXPECTED_HEADERS[c]}: contient «>» (sera converti en &gt;)`);
    }
  }
  return [errors, warns];
}

// --- Auth middleware ---
function loginRequired(req, res, next) {
  if (!req.session.user_id) return res.redirect("/login?next=" + req.path);
  next();
}

function subscriptionRequired(req, res, next) {
  if (!req.user) return res.redirect("/login?next=" + req.path);
  if (!subscriptionActive(req.user)) return res.redirect("/profile?msg=expired");
  next();
}

function adminRequired(req, res, next) {
  if (!req.user || !req.user.is_admin) return res.redirect("/");
  next();
}

// --- Render helpers ---
function gaBlock() {
  if (!GA_ID) return "";
  return (
    `<link rel="dns-prefetch" href="https://www.googletagmanager.com">` +
    `<link rel="preconnect" href="https://www.googletagmanager.com">` +
    `<script async src="https://www.googletagmanager.com/gtag/js?id=${GA_ID}"></script>` +
    `<script>window.dataLayer=window.dataLayer||[];function gtag(){dataLayer.push(arguments)};gtag('js',new Date());gtag('config','${GA_ID}');</script>`
  );
}

function userBar(user) {
  if (user) {
    let html = `<span style="color:#6b7280;margin-right:4px">${user.email}</span>
<a href="/profile" style="color:#4a6cf7">Mon compte</a>`;
    if (user.is_admin)
      html += `<a href="/admin" style="color:#059669">Admin</a>`;
    html += `<a href="/logout" style="color:#ef4444">Déconnexion</a>`;
    return html;
  }
  return `<a class="login-link" href="/login">Se connecter</a>`;
}

function userBarStyle() {
  return `position:absolute;top:20px;right:28px;display:flex;align-items:center;gap:4px;font-size:.82rem`;
}

function userBarLinkStyle() {
  return `text-decoration:none;font-weight:600`;
}

// --- Security headers ---
app.use((req, res, next) => {
  res.setHeader(
    "Strict-Transport-Security",
    "max-age=31536000; includeSubDomains"
  );
  next();
});

const LOADING_HTML = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Chargement...</title><style>body{margin:0;font-family:Inter,system-ui,sans-serif;background:linear-gradient(135deg,#0f0c29,#302b63,#24243e);min-height:100vh;display:flex;align-items:center;justify-content:center;color:#fff;text-align:center}.card{background:rgba(255,255,255,.97);border-radius:20px;padding:48px;max-width:440px;box-shadow:0 25px 60px rgba(0,0,0,.5)}.spinner{width:48px;height:48px;border:4px solid #e5e7eb;border-top-color:#667eea;border-radius:50%;animation:s 1s linear infinite;margin:0 auto 20px}@keyframes s{to{transform:rotate(360deg)}}p{color:#6b7280;font-size:.9rem;line-height:1.5;margin:0}</style></head><body><div class="card"><div class="spinner"></div><p>L'application démarre, veuillez patienter quelques instants...</p></div></body></html>`;

// --- Load user middleware ---
app.use((req, res, next) => {
  if (!_dbReady) {
    if (['/og-image.png','/robots.txt','/sitemap.xml'].includes(req.path))
      return next();
    return res.status(503).send(LOADING_HTML);
  }
  req.user = req.session.user_id ? getUser(req.session.user_id) : null;
  next();
});

// --- Routes ---
app.get("/", (req, res) => {
  const user = req.user;
  res.send(renderIndex(user, null, null, null, req.session.csrfToken || ""));
});

app.get("/job/:id", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ status: "notfound" });
  res.json(job);
});

app.post("/upload", upload.single("file"), validateCsrf, (req, res) => {
  const user = req.user;
  const ct = req.session.csrfToken || "";
  try {
    if (!req.file) return res.send(renderIndex(user, "Aucun fichier sélectionné", "err", null, ct));
    const filepath = req.file.path;
    const societe = (req.body.societe || "").trim().toUpperCase();
    const idfiscal = (req.body.idfiscal || "").trim();
    const annee = (req.body.annee || "2025").trim();
    const periode = (req.body.periode || "1").trim();
    const perType = (req.body.per_type || "mensuel").trim();
    if (!["mensuel", "trimestriel"].includes(perType))
      return respondError(res, user, "Type de période invalide", ct);
    const regime = perType === "mensuel" ? "1" : "2";
    if (!societe) return respondError(res, user, "Veuillez saisir le nom de la société", ct);
    if (!idfiscal) return respondError(res, user, "Veuillez saisir l'identifiant fiscal", ct);
    if (!/^\d{4}$/.test(annee))
      return respondError(res, user, "Année invalide", ct);
    const anneeI = parseInt(annee);
    if (anneeI < 2020 || anneeI > 2030)
      return respondError(res, user, "Année hors plage (2020-2030)", ct);
    const perI = parseInt(periode);
    if (isNaN(perI))
      return respondError(res, user, "Période invalide", ct);
    if (perType === "trimestriel") {
      if (perI < 1 || perI > 4)
        return respondError(res, user, "Trimestre invalide (1-4)", ct);
    } else {
      if (perI < 1 || perI > 12)
        return respondError(res, user, "Mois invalide (1-12)", ct);
    }

    // Créer un job asynchrone
    const jobId = uuidv4();
    const userId = req.session.user_id || null;
    jobs.set(jobId, { status: "queued", progress: 0, error: null, fileId: null, zipName: null, nrows: 0, warns: [] });

    // Répondre immédiatement avec page d'attente
    res.send(`<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Conversion en cours...</title><style>body{font-family:Inter,system-ui,sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);min-height:100vh;display:flex;align-items:center;justify-content:center;margin:0;padding:20px} .card{background:#fff;border-radius:20px;padding:40px;max-width:480px;width:100%;text-align:center;box-shadow:0 25px 60px rgba(0,0,0,.3)} .spinner{width:48px;height:48px;border:4px solid #e5e7eb;border-top-color:#667eea;border-radius:50%;animation:spin 1s linear infinite;margin:0 auto 20px} @keyframes spin{to{transform:rotate(360deg)}} h1{font-size:1.3rem;color:#1a1a2e;margin:0 0 8px} p{color:#6b7280;font-size:.95rem;margin:0 0 20px} .progress{height:6px;background:#f3f4f6;border-radius:3px;overflow:hidden} .bar{height:100%;width:0%;background:linear-gradient(90deg,#667eea,#764ba2);transition:width .3s} .err{color:#dc2626;margin-top:16px}</style></head><body><div class="card"><div class="spinner"></div><h1>Conversion en cours...</h1><p id="msg">Analyse du fichier Excel</p><div class="progress"><div class="bar" id="bar"></div></div><div id="err" class="err"></div></div><script>const jobId="${jobId}";let p=10;const bar=document.getElementById('bar'),msg=document.getElementById('msg'),err=document.getElementById('err');function poll(){fetch('/job/'+jobId).then(r=>r.json()).then(j=>{if(j.status==='processing'||j.status==='queued'){p=Math.min(90,p+5);bar.style.width=p+'%';if(j.progress)msg.textContent=j.progress;setTimeout(poll,1500)}else if(j.status==='done'){bar.style.width='100%';msg.textContent='Terminé ! Redirection...';location.href='/download/'+j.fileId}else if(j.status==='error'){err.textContent=j.error||'Erreur';msg.textContent='Échec'}else{setTimeout(poll,1500)}}).catch(()=>setTimeout(poll,2000))}poll();</script></body></html>`);

    // Lancer le traitement en arrière-plan
    setImmediate(() => {
      const job = jobs.get(jobId);
      try {
        job.status = "processing";
        job.progress = "Validation du fichier...";
        const [errs, warns] = validate(filepath);
        if (errs.length) {
          job.status = "error";
          job.error = errs.slice(0,3).join(" | ");
          try { fs.unlinkSync(filepath); } catch {}
          return;
        }
        job.progress = "Génération XML...";
        job.warns = warns;
        const outPath = processFile(filepath, societe, idfiscal, annee, periode, regime);
        const label = perType === "trimestriel" ? TRIM_NAMES[perI - 1] : MONTH_NAMES[perI - 1];
        const safeName = societe.replace(/[^a-zA-Z0-9_\- ]/g, "").trim().replace(/ /g, "_");
        const xmlName = `TVA_${safeName}_${label}_${annee}.xml`;
        const zipName = `TVA_${safeName}_${label}_${annee}.zip`;
        const zipPath = path.join(__dirname, "tmp_xml", uuidv4() + ".zip");
        const xmlContent = fs.readFileSync(outPath, "utf8");
        const nrows = (xmlContent.match(/<rd>/g) || []).length;
        makeZip(zipPath, xmlName, outPath);
        fs.unlinkSync(outPath);
        const fileId = uuidv4();
        dbRun("INSERT INTO downloads (file_id,zip_path,zip_name,user_id) VALUES (?,?,?,?)", [fileId, zipPath, zipName, userId]);
        job.status = "done";
        job.fileId = fileId;
        job.zipName = zipName;
        job.nrows = nrows;
      } catch (e) {
        console.error("Job error:", e);
        job.status = "error";
        job.error = "Une erreur est survenue lors de la conversion";
      } finally {
        try { fs.unlinkSync(filepath); } catch {}
        // nettoyage après 10 min
        setTimeout(() => jobs.delete(jobId), 10 * 60 * 1000);
      }
    });

  } catch (e) {
    console.error("Upload error:", e);
    res.send(renderIndex(user, "Une erreur est survenue", "err", null, ct));
    if (req.file) try { fs.unlinkSync(req.file.path); } catch {}
  }
});

function makeZip(zipPath, xmlName, xmlPath) {
  const zip = new AdmZip();
  zip.addLocalFile(xmlPath, "", xmlName);
  zip.writeZip(zipPath);
}

function respondError(res, user, msg, csrfToken) {
  res.send(renderIndex(user, msg, "err", null, csrfToken));
}

app.get("/download/:fileId", subscriptionRequired, (req, res) => {
  const row = dbGet("SELECT zip_path, zip_name, user_id FROM downloads WHERE file_id=?", [req.params.fileId]);
  if (!row) return res.redirect("/");
  if (row.user_id && row.user_id !== req.session.user_id) return res.redirect("/");
  dbRunSync("DELETE FROM downloads WHERE file_id=?", [req.params.fileId]);
  res.download(row.zip_path, row.zip_name, (err) => {
    if (err) console.error("Download error:", err);
    try { fs.unlinkSync(row.zip_path); } catch {}
  });
});

app.get("/login", (req, res) => {
  const nextUrl = req.query.next || "/";
  res.send(renderLogin(nextUrl, false, req.session.csrfToken || ""));
});

app.post("/login", limiter, validateCsrf, async (req, res) => {
  const email = (req.body.email || "").trim();
  const pwd = req.body.pwd || "";
  const user = getUserByEmail(email);
  if (user && await bcrypt.compare(pwd, user.password)) {
    req.session.user_id = user.id;
    return res.redirect(safeNext(req.body.next || req.query.next || "", "/"));
  }
  res.send(renderLogin(req.body.next || "/", true, req.session.csrfToken || ""));
});

app.get("/logout", (req, res) => {
  req.session.destroy();
  res.redirect("/");
});

app.get("/register", (req, res) => {
  res.send(renderRegister(null, req.session.csrfToken || ""));
});

app.post("/register", validateCsrf, async (req, res) => {
  const email = (req.body.email || "").trim();
  const pwd = req.body.pwd || "";
  const name = (req.body.name || "").trim();
  const ct = req.session.csrfToken || "";
  if (!email || !pwd) return res.send(renderRegister("Email et mot de passe requis", ct));
  if (getUserByEmail(email)) return res.send(renderRegister("Cet email est déjà utilisé", ct));
  if (pwd.length < 6) return res.send(renderRegister("Mot de passe trop court (min 6 caractères)", ct));
  await createUser(email, pwd, name);
  res.redirect("/login?msg=registered");
});

app.get("/profile", (req, res) => {
  const user = req.user;
  if (!user) return res.redirect("/login?next=/profile");
  const expired = req.query.msg === "expired";
  const pwdType = req.query.pwd || "";
  const pwdMsg = {
    ok: { type: "ok", text: "Mot de passe changé avec succès." },
    wrong: { type: "err", text: "Mot de passe actuel incorrect." },
    mismatch: { type: "err", text: "Les nouveaux mots de passe ne correspondent pas." },
    short: { type: "err", text: "Le nouveau mot de passe doit contenir au moins 6 caractères." },
  }[pwdType] || null;
  res.send(renderProfile(user, expired, pwdMsg, req.session.csrfToken || ""));
});

app.post("/change_password", validateCsrf, async (req, res) => {
  const user = req.user;
  if (!user) return res.redirect("/login");
  const current = req.body.current_pwd || "";
  const newPwd = req.body.new_pwd || "";
  const confirm = req.body.confirm_pwd || "";
  if (!await bcrypt.compare(current, user.password))
    return res.redirect("/profile?pwd=wrong");
  if (newPwd !== confirm) return res.redirect("/profile?pwd=mismatch");
  if (newPwd.length < 6) return res.redirect("/profile?pwd=short");
  const hashed = await bcrypt.hash(newPwd, 10);
  dbRunSync("UPDATE users SET password=? WHERE id=?", [hashed, user.id]);
  res.redirect("/profile?pwd=ok");
});

app.get("/admin", adminRequired, (req, res) => {
  const user = req.user;
  const page = Math.max(1, parseInt(req.query.page) || 1);
  const users = getAllUsers(page, 50);
  const total = countUsers();
  const maxPage = Math.max(1, Math.ceil(total / 50));
  res.send(renderAdmin(user, users, page, maxPage, total, req.session.csrfToken || ""));
});

app.post("/admin/set_sub", adminRequired, validateCsrf, (req, res) => {
  const targetId = req.body.user_id;
  const start = req.body.start;
  const end = req.body.end;
  dbRunSync("UPDATE users SET subscription_start=?, subscription_end=?, active=1 WHERE id=?", [start, end, targetId]);
  res.redirect("/admin");
});

app.get("/aide", (req, res) => res.send(HELP_HTML));
app.get("/apropos", (req, res) => res.send(ABOUT_HTML));
app.get("/contact", (req, res) => res.send(CONTACT_HTML));
app.get("/confidentialite", (req, res) => res.send(PRIVACY_HTML));

app.get("/modele", (req, res) => {
  if (fs.existsSync(MODEL_PATH))
    return res.download(MODEL_PATH, "modele_tva.xlsx");
  res.redirect("/");
});

app.get("/og-image.png", (req, res) => {
  res.set({
    "Content-Type": "image/png",
    "Cache-Control": "public, max-age=86400",
  });
  res.send(OG_PNG);
});

app.get("/robots.txt", (req, res) => {
  res.set({ "Content-Type": "text/plain", "Cache-Control": "public, max-age=86400" });
  res.send(`User-agent: *
Allow: /
Sitemap: ${SITE_URL}/sitemap.xml
`);
});

app.get("/sitemap.xml", (req, res) => {
  res.set({ "Content-Type": "application/xml", "Cache-Control": "public, max-age=86400" });
  res.send(`<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
<url><loc>${SITE_URL}/</loc><priority>1.0</priority></url>
<url><loc>${SITE_URL}/aide</loc><priority>0.8</priority></url>
<url><loc>${SITE_URL}/apropos</loc><priority>0.7</priority></url>
<url><loc>${SITE_URL}/contact</loc><priority>0.6</priority></url>
<url><loc>${SITE_URL}/confidentialite</loc><priority>0.5</priority></url>
</urlset>`);
});

app.use((req, res) => {
  res.status(404).send(NOT_FOUND_HTML);
});

// ====================== TEMPLATES ======================

const HEAD_META = `<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="Convertissez vos relevés Excel en XML de déclaration TVA conforme à la Direction Générale des Impôts. Générateur simple, rapide et sécurisé. Conforme aux normes fiscales marocaines.">
<meta name="keywords" content="TVA, XML, Excel, déclaration, Direction Générale des Impôts, Maroc, relevé de déductions, conversion, facture">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='32' height='32'%3E%3Crect width='32' height='32' rx='8' fill='%23667eea'/%3E%3Ctext x='16' y='22' text-anchor='middle' fill='%23fff' font-size='16' font-weight='700' font-family='Arial'%3ET%3C/text%3E%3C/svg%3E">
<link rel="apple-touch-icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='180' height='180'%3E%3Crect width='180' height='180' rx='20' fill='%23667eea'/%3E%3Ctext x='90' y='125' text-anchor='middle' fill='%23fff' font-size='100' font-weight='700' font-family='Arial'%3ET%3C/text%3E%3C/svg%3E">
<meta property="og:title" content="Générateur de déclaration TVA - Excel vers XML">
<meta property="og:description" content="Convertissez vos relevés Excel en XML de déclaration TVA. Conforme aux normes fiscales marocaines.">
<meta property="og:type" content="website">
<meta property="og:url" content="${SITE_URL}">
<meta property="og:image" content="${SITE_URL}/og-image.png">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta name="twitter:card" content="summary_large_image">
<meta name="twitter:title" content="Générateur TVA - Excel vers XML">
<meta name="twitter:description" content="Convertissez vos relevés Excel en XML de déclaration TVA conforme aux normes fiscales marocaines.">
<link rel="canonical" href="${SITE_URL}/">
<link rel="alternate" hreflang="fr" href="${SITE_URL}/">
<script type="application/ld+json">
{"@context":"https://schema.org","@type":"WebApplication","name":"G\u00e9n\u00e9rateur TVA","url":"${SITE_URL}","description":"Convertissez vos relev\u00e9s Excel en XML de d\u00e9claration TVA","applicationCategory":"BusinessApplication","operatingSystem":"All","offers":{"@type":"Offer","price":"50","priceCurrency":"MAD"}}
</script>
<script type="application/ld+json">
{"@context":"https://schema.org","@type":"BreadcrumbList","itemListElement":[{"@type":"ListItem","position":1,"name":"Accueil","item":"${SITE_URL}/"}]}
</script>`;

const COMMON_CSS = `*,:after,:before{margin:0;padding:0;box-sizing:border-box}html{scroll-behavior:smooth}
body{font-family:Inter,'Segoe UI',system-ui,sans-serif;background:linear-gradient(135deg,#0f0c29,#302b63,#24243e);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px}`;

function renderProcessing(user, jobId, societe) {
  return `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Conversion en cours</title><style>${COMMON_CSS}body{margin:0;font-family:Inter,system-ui,sans-serif;background:linear-gradient(135deg,#0f0c29,#302b63,#24243e);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}.card{background:rgba(255,255,255,.97);border-radius:20px;padding:48px;width:520px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);text-align:center;position:relative;overflow:hidden}.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}.spinner{width:64px;height:64px;border:5px solid #e5e7eb;border-top-color:#667eea;border-radius:50%;animation:s 1s linear infinite;margin:0 auto 24px}@keyframes s{to{transform:rotate(360deg)}}h1{font-size:1.4rem;color:#1a1a2e;margin-bottom:12px}p{color:#6b7280;font-size:.95rem;line-height:1.6;margin-bottom:8px}.progress{height:6px;background:#e5e7eb;border-radius:3px;overflow:hidden;margin:24px 0}.bar{height:100%;width:0%;background:linear-gradient(90deg,#667eea,#764ba2);transition:width .4s}.societe{font-weight:700;color:#4a6cf7}</style></head><body><div class="card"><div class="spinner" id="spinner"></div><h1>Conversion en cours...</h1><p>Traitement du fichier pour <span class="societe">${societe}</span></p><p id="status">Initialisation...</p><div class="progress"><div class="bar" id="bar"></div></div><p style="font-size:.8rem;color:#9ca3af">Ne fermez pas cette page</p></div><script>const jobId="${jobId}";let tries=0;async function check(){tries++;try{const r=await fetch('/status/'+jobId);const d=await r.json();if(d.status==='processing'||d.status==='queued'){document.getElementById('status').textContent='Traitement en cours...';document.getElementById('bar').style.width=(d.progress||30)+'%';setTimeout(check,1000);}else if(d.status==='done'){document.getElementById('spinner').style.display='none';document.getElementById('status').innerHTML='Terminé ! Redirection...';document.getElementById('bar').style.width='100%';setTimeout(()=>{window.location='/success/'+d.fileId},800);}else if(d.status==='error'){document.getElementById('status').innerHTML='<span style=color:#ef4444>Erreur: '+(d.error||'inconnue')+'</span>';}}catch(e){if(tries<60)setTimeout(check,2000);}}check();</script></body></html>`;
}

app.get("/status/:jobId", (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) return res.json({ status: 'notfound' });
  res.json(job);
});

app.get("/success/:fileId", (req, res) => {
  if (!req.session.user_id) return res.redirect("/login?next=/success/"+req.params.fileId);
  const row = dbGet("SELECT zip_name FROM downloads WHERE file_id=?", [req.params.fileId]);
  if (!row) return res.redirect("/");
  const job = Array.from(jobs.values()).find(j => j.fileId === req.params.fileId);
  res.send(renderSuccess(req.user, req.params.fileId, row.zip_name, job?.nrows || 0, job?.warns || []));
});

// Nettoyage jobs
setInterval(() => {
  const now = Date.now();
  for (const [id, job] of jobs) {
    if (job.createdAt && now - job.createdAt > 3600000) jobs.delete(id);
  }
}, 600000);


const CARD_CSS = `.card{background:rgba(255,255,255,.97);border-radius:20px;padding:48px;width:600px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before,.user-bar,.login-link,.footer-links{display:none!important}}`;

const FOOTER_LINKS = `<div class="footer-links">
<a href="/">Accueil</a> &middot; <a href="/apropos">À propos</a> &middot; <a href="/confidentialite">Confidentialité</a> &middot; <a href="/contact">Contact</a>
</div>`;

const SUCCESS_CSS = `*,:after,:before{margin:0;padding:0;box-sizing:border-box}html{scroll-behavior:smooth}
body{font-family:Inter,'Segoe UI',system-ui,sans-serif;background:linear-gradient(135deg,#0f0c29,#302b63,#24243e);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:40px;width:480px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);text-align:center;position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before,.user-bar,.login-link,.footer-links{display:none!important}}
.check-circle{width:72px;height:72px;margin:0 auto 20px;background:linear-gradient(135deg,#ecfdf5,#d1fae5);border-radius:50%;display:flex;align-items:center;justify-content:center}
h1{font-size:1.3rem;color:#1a1a2e;margin-bottom:6px;letter-spacing:-.02em}
.info{color:#6b7280;font-size:.88rem;margin-bottom:8px;line-height:1.5}
.badge{display:inline-block;background:#eef1ff;color:#667eea;font-size:.8rem;font-weight:600;padding:4px 14px;border-radius:20px;margin-bottom:24px}
.actions{display:flex;flex-direction:column;gap:10px}
.btn{display:flex;align-items:center;justify-content:center;gap:8px;padding:13px 24px;border-radius:10px;font-size:.9rem;font-weight:600;text-decoration:none;cursor:pointer}
.btn-primary{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff}
.btn-primary:hover{box-shadow:0 8px 24px rgba(102,126,234,.35)}
.btn-secondary{background:#f3f4f6;color:#374151;border:2px solid #e5e7eb}
.fname{font-size:.78rem;color:#9ca3af;margin-top:16px}
.footer-links{margin-top:20px;padding-top:14px;border-top:1px solid #f0f0f5;font-size:.78rem;color:#9ca3af}
.footer-links a{color:#4a6cf7;font-weight:600;text-decoration:none}
.login-link{position:absolute;top:20px;right:24px;font-size:.82rem;color:#4a6cf7;text-decoration:underline;font-weight:600;white-space:nowrap}
.user-bar{position:absolute;top:20px;right:24px;display:flex;align-items:center;gap:4px;font-size:.82rem}
.user-bar a{text-decoration:none;font-weight:600}
@media(max-width:480px){.user-bar,.login-link{font-size:.72rem;right:16px;top:14px}}`;

// --- Render functions ---
function renderIndex(user, msg, cat, errors, csrfToken) {
  const errorBlock = errors
    ? `<div class="msg err"><strong>Fichier invalide</strong><ul>${errors
        .map((e) => `<li>${e}</li>`)
        .join("")}</ul></div>`
    : ``;
  return `<!DOCTYPE html>
<html lang="fr">
<head>${HEAD_META}
${gaBlock()}
<title>Générateur TVA - Convertir Excel en XML de déclaration</title>
<style>${COMMON_CSS}
${CARD_CSS}
.login-link{position:absolute;top:20px;right:28px;font-size:.82rem;color:#4a6cf7;text-decoration:underline;font-weight:600;white-space:nowrap}
.login-link:hover{opacity:.8}
.header{display:flex;align-items:center;gap:12px;margin-bottom:20px}
.header svg{flex-shrink:0}
.header h1{font-size:1.3rem;font-weight:700;color:#1a1a2e;letter-spacing:-.02em}
.header .badge{background:#eef1ff;color:#667eea;font-size:.7rem;font-weight:600;padding:3px 10px;border-radius:20px;margin-left:auto}
.desc{color:#6b7280;font-size:.9rem;line-height:1.65;margin-bottom:32px;padding-bottom:24px;border-bottom:1px solid #f0f0f5}
.desc a{color:#4a6cf7;font-weight:600;text-decoration:none}
.step{display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:.8rem;font-weight:600;color:#1a1a2e;text-transform:uppercase;letter-spacing:.04em}
.step span{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;width:20px;height:20px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:.7rem;font-weight:700}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.form-grid .full{grid-column:1/-1}
.form-group{display:flex;flex-direction:column;gap:6px}
.form-group label{font-size:.82rem;font-weight:600;color:#374151}
.form-group input,.form-group select{width:100%;padding:12px 14px;border:2px solid #e5e7eb;border-radius:10px;font-size:.9rem;outline:0;background:#f9fafb;color:#111827}
.form-group input:focus,.form-group select:focus{border-color:#667eea;box-shadow:0 0 0 3px rgba(102,126,234,.15)}
.radio-group{display:flex;gap:16px;margin-top:4px}
.radio-group label{display:flex;align-items:center;gap:6px;font-size:.88rem;font-weight:500;color:#374151;cursor:pointer}
.radio-group input[type=radio]{width:16px;height:16px;accent-color:#667eea}
.upload-zone{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;padding:36px 20px;border:2px dashed #d1d5db;border-radius:14px;background:#fafafe;cursor:pointer;margin-top:4px}
.upload-zone:hover,.upload-zone.dragover{border-color:#667eea;background:#eef1ff}
.upload-zone input{display:none}
.upload-zone .icon svg{display:block}
.upload-zone .txt{font-size:.88rem;color:#6b7280}
.upload-zone .txt strong{color:#4a6cf7}
#fname{font-size:.8rem;color:#059669;font-weight:600;margin-top:6px;text-align:center}
.btn-wrap{margin-top:24px}
.btn-wrap button{width:100%;padding:14px;border:none;border-radius:12px;font-size:1rem;font-weight:700;color:#fff;background:linear-gradient(135deg,#667eea,#764ba2);cursor:pointer}
.btn-wrap button:hover{box-shadow:0 8px 24px rgba(102,126,234,.35)}
.btn-wrap button:disabled{opacity:.6;cursor:not-allowed}
.msg{margin-top:16px;padding:12px 16px;border-radius:10px;font-size:.85rem;display:none;line-height:1.5}
.msg.err{display:block;background:#fef2f2;border:1px solid #fecaca;color:#991b1b}
.msg.ok{display:block;background:#ecfdf5;border:1px solid #a7f3d0;color:#065f46}
.spinner{display:inline-block;width:16px;height:16px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:s 1s linear infinite;margin-right:8px;vertical-align:middle}
@keyframes s{to{transform:rotate(360deg)}}
.footer-links{margin-top:24px;padding-top:16px;border-top:1px solid #f0f0f5;font-size:.78rem;color:#9ca3af;text-align:center}
.footer-links a{color:#4a6cf7;font-weight:600;text-decoration:none}
@media(max-width:640px){.form-grid{grid-template-columns:1fr!important}.user-bar,.login-link{font-size:.75rem;right:16px;top:14px}.header h1{font-size:1.1rem}.radio-group{flex-wrap:wrap;gap:10px}}
@media(max-width:480px){.header h1{font-size:1rem}.user-bar{font-size:.7rem;right:12px;top:12px}.login-link{right:12px;top:12px}}
</style>
</head>
<body>
<div class="card">
<div class="user-bar" style="${userBarStyle()}">
${userBar(user)}
</div>
<div class="header">
<svg width="36" height="36" viewBox="0 0 36 36" fill="none" aria-label="Logo générateur TVA" role="img"><rect width="36" height="36" rx="10" fill="url(#g)"/><path d="M10 12h16v2H10zm0 5h16v2H10zm0 5h16v2H10z" fill="#fff" opacity=".8"/><text x="18" y="9" text-anchor="middle" fill="#fff" font-size="6" font-weight="700" font-family="Arial">XML</text><defs><linearGradient id="g" x1="0" y1="0" x2="36" y2="36"><stop stop-color="#667eea"/><stop offset="1" stop-color="#764ba2"/></linearGradient></defs></svg>
<h1>Générateur TVA</h1>
<span class="badge">Excel → XML</span>
</div>
<div class="desc">
Convertissez votre relevé Excel en fichier XML de déclaration de TVA.<br>
<a href="/modele">Télécharger le modèle Excel</a> &middot; <a href="/aide">Comment ça marche ?</a>
</div>
<div class="step"><span>1</span> Informations société</div>
<form id="form" method="post" enctype="multipart/form-data" action="/upload">
<input type="hidden" name="csrf_token" value="${csrfToken}">
<div class="form-grid">
<div class="full">
<div class="form-group">
<label for="societe">Société</label>
<input type="text" id="societe" name="societe" placeholder="Raison sociale" required>
</div>
</div>
<div class="form-group">
<label for="idfiscal">Identifiant fiscal</label>
<input type="text" id="idfiscal" name="idfiscal" placeholder="Ex: 1234567" required>
</div>
<div class="form-group">
<label for="annee">Année</label>
<input type="number" id="annee" name="annee" value="2025" min="2020" max="2030" required>
</div>
<div class="full">
<div class="form-group">
<label>Période</label>
<div class="radio-group">
<label><input type="radio" name="per_type" value="mensuel" checked onchange="togglePeriod()"><span>Mensuel</span></label>
<label><input type="radio" name="per_type" value="trimestriel" onchange="togglePeriod()"><span>Trimestriel</span></label>
</div>
<select name="periode" id="periode">
<option value="1">Janvier</option>
<option value="2">Février</option>
<option value="3">Mars</option>
<option value="4">Avril</option>
<option value="5">Mai</option>
<option value="6">Juin</option>
<option value="7">Juillet</option>
<option value="8">Août</option>
<option value="9">Septembre</option>
<option value="10">Octobre</option>
<option value="11">Novembre</option>
<option value="12">Décembre</option>
</select>
</div>
</div>
</div>
<div class="step"><span>2</span> Fichier Excel</div>
<label class="upload-zone" id="dropzone">
<div class="icon"><svg width="48" height="56" viewBox="0 0 48 56" fill="none"><rect x="4" y="2" width="40" height="52" rx="6" fill="#1D6F42" stroke="#145530" stroke-width="1.5"/><rect x="10" y="26" width="28" height="2" rx="1" fill="#fff" opacity=".6"/><rect x="10" y="31" width="28" height="2" rx="1" fill="#fff" opacity=".6"/><rect x="10" y="36" width="28" height="2" rx="1" fill="#fff" opacity=".6"/><rect x="10" y="41" width="28" height="2" rx="1" fill="#fff" opacity=".6"/><text x="24" y="19" text-anchor="middle" fill="#fff" font-size="10" font-weight="700" font-family="Arial,sans-serif">XLS</text></svg></div>
<div class="txt">Glissez-déposez votre fichier .xls ou .xlsx ici, ou <strong>parcourez</strong></div>
<input type="file" name="file" accept=".xls,.xlsx" required>
</label>
<div id="fname"></div>
<div class="btn-wrap"><button type="submit" id="btn">Lancer la conversion</button></div>
</form>
<div class="msg" id="msg">${msg ? (cat === 'err' ? `<strong>Erreur</strong> ${msg}` : msg) : ''}</div>
${errorBlock}
${FOOTER_LINKS}
</div>
<script>
function togglePeriod(){
const sel=document.getElementById('periode');
if(document.querySelector('input[name="per_type"]:checked').value==='trimestriel'){
sel.innerHTML='<option value="1">1er Trimestre</option><option value="2">2e Trimestre</option><option value="3">3e Trimestre</option><option value="4">4e Trimestre</option>';
}else{
sel.innerHTML='<option value="1">Janvier</option><option value="2">Février</option><option value="3">Mars</option><option value="4">Avril</option><option value="5">Mai</option><option value="6">Juin</option><option value="7">Juillet</option><option value="8">Août</option><option value="9">Septembre</option><option value="10">Octobre</option><option value="11">Novembre</option><option value="12">Décembre</option>';
}
}
const dz=document.getElementById('dropzone'),fi=dz.querySelector('input'),fn=document.getElementById('fname'),btn=document.getElementById('btn'),msg=document.getElementById('msg'),form=document.getElementById('form');
fi.addEventListener('change',()=>{const f=fi.files[0];fn.textContent=f?f.name:'';msg.style.display='none'});
dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('dragover')});
dz.addEventListener('dragleave',()=>dz.classList.remove('dragover'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('dragover');if(e.dataTransfer.files.length){fi.files=e.dataTransfer.files;fn.textContent=e.dataTransfer.files[0].name;msg.style.display='none'}});
form.addEventListener('submit',e=>{if(!fi.files.length){e.preventDefault();msg.innerHTML='Veuillez sélectionner un fichier';msg.className='msg err';return}btn.disabled=true;btn.innerHTML='<span class="spinner"></span> Conversion en cours...'});
</script>
</body>
</html>`;
}

function renderSuccess(user, fileId, fname, nrows, warns) {
  const warnBlock = warns && warns.length
    ? `<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:10px;padding:12px 14px;margin-bottom:16px;text-align:left;font-size:.8rem;color:#92400e">
<strong>&#9888; Remarques</strong>
<ul style="margin:6px 0 0 16px">${warns.map(w => `<li>${w}</li>`).join('')}</ul>
</div>`
    : '';
  return `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex">
<link rel="canonical" href="${SITE_URL}/">
<title>Conversion réussie - Générateur TVA</title>
${gaBlock()}
<style>${SUCCESS_CSS}</style>
<body>
<div class="card">
<div class="user-bar" style="${userBarStyle()}">
${userBar(user)}
</div>
<div class="check-circle"><svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="4,13 9,18 20,6"/></svg></div>
<h1>Déclaration générée</h1>
<p class="info">${nrows} ligne${nrows > 1 ? 's' : ''} de relevé convertie${nrows > 1 ? 's' : ''} avec succès</p>
${warnBlock}
<div class="badge">Prêt à télécharger</div>
<div class="actions">
<a class="btn btn-primary" href="/download/${fileId}"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.5" stroke-linecap="round"><polyline points="8,17 12,21 16,17"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.88 18.09A5 5 0 0018 9h-1.26A8 8 0 103 16.29"/></svg> Télécharger le fichier XML</a>
<a class="btn btn-secondary" href="/">&larr; Nouvelle conversion</a>
</div>
<p class="fname">${fname}</p>
<div class="footer-links">
<a href="/">Accueil</a> &middot; <a href="/apropos">À propos</a> &middot; <a href="/confidentialite">Confidentialité</a> &middot; <a href="/contact">Contact</a>
</div>
</div>
</body>
</html>`;
}

function renderLogin(nextUrl, err, csrfToken) {
  return `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Connexion - Déclaration TVA</title>
<style>${COMMON_CSS}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:40px;width:400px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);text-align:center;position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before{display:none!important}}
.lock{width:56px;height:56px;margin:0 auto 16px;background:#eef1ff;border-radius:50%;display:flex;align-items:center;justify-content:center}
h1{font-size:1.2rem;color:#1a1a2e;margin-bottom:4px}
.sub{color:#6b7280;font-size:.85rem;margin-bottom:24px}
.form-group{margin-bottom:16px;text-align:left}
.form-group label{display:block;font-size:.8rem;font-weight:600;color:#374151;margin-bottom:5px}
.form-group input{width:100%;padding:11px 14px;border:2px solid #e5e7eb;border-radius:10px;font-size:.9rem;outline:0;transition:all .2s;background:#f9fafb}
.form-group input:focus{border-color:#667eea;background:#fff;box-shadow:0 0 0 4px rgba(102,126,234,.12)}
.btn{width:100%;padding:13px;border:none;border-radius:10px;font-size:.95rem;font-weight:600;cursor:pointer;color:#fff;background:linear-gradient(135deg,#667eea,#764ba2)}
.btn:hover{transform:translateY(-1px);box-shadow:0 8px 24px rgba(102,126,234,.35)}
.register-link{display:block;margin-top:16px;font-size:.82rem;color:#6b7280}
.register-link a{color:#4a6cf7;font-weight:600;text-decoration:underline}
</style>
</head>
<body>
<div class="card">
${err ? `<script>alert("Email ou mot de passe incorrect");window.location.href="/";</script>` : ''}
<div class="lock"><svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2.5" stroke-linecap="round"><rect x="5" y="11" width="14" height="11" rx="2"/><path d="M8 11V7a4 4 0 118 0v4"/></svg></div>
<h1>Connexion</h1>
<p class="sub">Connectez-vous pour télécharger vos fichiers</p>
<form method="post" action="/login">
<input type="hidden" name="csrf_token" value="${csrfToken}">
<div class="form-group"><label for="email">Email</label><input type="email" id="email" name="email" required autofocus></div>
<div class="form-group"><label for="pwd">Mot de passe</label><input type="password" id="pwd" name="pwd" required></div>
<input type="hidden" name="next" value="${nextUrl}">
<button type="submit" class="btn">Se connecter</button>
</form>
<div class="register-link">Pas encore de compte ? <a href="/register">Créer un compte</a></div>
</div></body></html>`;
}

function renderRegister(err, csrfToken) {
  return `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Inscription - Déclaration TVA</title>
<style>${COMMON_CSS}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:40px;width:400px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);text-align:center;position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before{display:none!important}}
h1{font-size:1.2rem;color:#1a1a2e;margin-bottom:4px}
.sub{color:#6b7280;font-size:.85rem;margin-bottom:24px}
.form-group{margin-bottom:16px;text-align:left}
.form-group label{display:block;font-size:.8rem;font-weight:600;color:#374151;margin-bottom:5px}
.form-group input{width:100%;padding:11px 14px;border:2px solid #e5e7eb;border-radius:10px;font-size:.9rem;outline:0;background:#f9fafb}
.form-group input:focus{border-color:#667eea;background:#fff;box-shadow:0 0 0 4px rgba(102,126,234,.12)}
.btn{width:100%;padding:13px;border:none;border-radius:10px;font-size:.95rem;font-weight:600;cursor:pointer;color:#fff;background:linear-gradient(135deg,#667eea,#764ba2)}
.btn:hover{transform:translateY(-1px);box-shadow:0 8px 24px rgba(102,126,234,.35)}
.msg{color:#ef4444;font-size:.82rem;margin-bottom:12px}
.login-link{display:block;margin-top:16px;font-size:.82rem;color:#6b7280}
.login-link a{color:#4a6cf7;font-weight:600;text-decoration:underline}
</style>
</head>
<body>
<div class="card">
<h1>Créer un compte</h1>
<p class="sub">Inscrivez-vous pour accéder au service</p>
${err ? `<div class="msg">${err}</div>` : ''}
<form method="post" action="/register">
<input type="hidden" name="csrf_token" value="${csrfToken}">
<div class="form-group"><label for="name">Nom (optionnel)</label><input type="text" id="name" name="name"></div>
<div class="form-group"><label for="email">Email</label><input type="email" id="email" name="email" required autofocus></div>
<div class="form-group"><label for="pwd">Mot de passe</label><input type="password" id="pwd" name="pwd" required></div>
<button type="submit" class="btn">Créer mon compte</button>
</form>
<div class="login-link">Déjà un compte ? <a href="/login">Se connecter</a></div>
</div></body></html>`;
}

function renderProfile(user, expired, pwdMsg, csrfToken) {
  const today = new Date().toISOString().split("T")[0];
  const active =
    user.subscription_end && user.subscription_end >= today;
  return `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mon compte - Déclaration TVA</title>
<style>${COMMON_CSS}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:40px;width:480px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);text-align:center;position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before,.back,.footer-links{display:none!important}}
.back{position:absolute;top:20px;right:24px;font-size:.82rem;color:#4a6cf7;text-decoration:underline;font-weight:600}
h1{font-size:1.2rem;color:#1a1a2e;margin-bottom:16px}
.info-row{display:flex;justify-content:space-between;padding:12px 0;border-bottom:1px solid #f0f0f5;font-size:.88rem}
.info-row .label{color:#6b7280}
.info-row .value{color:#1a1a2e;font-weight:600}
.badge{display:inline-block;padding:3px 12px;border-radius:20px;font-size:.78rem;font-weight:600}
.badge.active{background:#ecfdf5;color:#059669}
.badge.inactive{background:#fef2f2;color:#ef4444}
.alert{margin-bottom:16px;padding:12px;background:#fef2f2;border:1px solid #fecaca;border-radius:10px;font-size:.82rem;color:#991b1b}
.actions{margin-top:24px;display:flex;flex-direction:column;gap:10px}
.btn{display:block;padding:12px;border-radius:10px;font-size:.9rem;font-weight:600;text-decoration:none}
.btn-primary{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff}
.btn-secondary{background:#f3f4f6;color:#374151;border:2px solid #e5e7eb}
.pwd-section{margin-top:20px;padding-top:16px;border-top:1px solid #f0f0f5;text-align:left}
.pwd-section h2{font-size:.95rem;color:#1a1a2e;margin-bottom:10px}
.pwd-section .form-group{margin-bottom:10px}
.pwd-section .form-group label{display:block;font-size:.78rem;font-weight:600;color:#374151;margin-bottom:3px}
.pwd-section .form-group input{width:100%;padding:9px 12px;border:2px solid #e5e7eb;border-radius:8px;font-size:.85rem;outline:0;background:#f9fafb}
.pwd-section .form-group input:focus{border-color:#667eea}
.pwd-section .btn-small{padding:9px 16px;border:none;border-radius:8px;font-size:.82rem;font-weight:600;cursor:pointer;color:#fff;background:linear-gradient(135deg,#667eea,#764ba2)}
.pwd-msg{font-size:.8rem;margin-top:6px}
.pwd-msg.ok{color:#059669}
.pwd-msg.err{color:#ef4444}
</style>
</head>
<body>
<div class="card">
<a class="back" href="/">&larr; Accueil</a>
<h1>Mon compte</h1>
${expired ? '<div class="alert">Votre abonnement a expiré. Contactez l\'administrateur pour le renouveler.</div>' : ''}
<div class="info-row"><span class="label">Email</span><span class="value">${user.email}</span></div>
<div class="info-row"><span class="label">Nom</span><span class="value">${user.name || '—'}</span></div>
<div class="info-row"><span class="label">Abonnement</span><span class="value">${user.subscription_end || 'Non défini'}</span></div>
<div class="info-row"><span class="label">Statut</span><span class="value">${active ? '<span class="badge active">Actif</span>' : '<span class="badge inactive">Inactif</span>'}</span></div>
<div class="pwd-section">
<h2>Changer le mot de passe</h2>
<form method="post" action="/change_password">
<input type="hidden" name="csrf_token" value="${csrfToken}">
<div class="form-group"><label for="current_pwd">Mot de passe actuel</label><input type="password" id="current_pwd" name="current_pwd" required></div>
<div class="form-group"><label for="new_pwd">Nouveau mot de passe</label><input type="password" id="new_pwd" name="new_pwd" required minlength="4"></div>
<div class="form-group"><label for="confirm_pwd">Confirmer</label><input type="password" id="confirm_pwd" name="confirm_pwd" required></div>
<button type="submit" class="btn-small">Changer</button>
</form>
${pwdMsg ? `<div class="pwd-msg ${pwdMsg.type}">${pwdMsg.text}</div>` : ''}
</div>
<div class="actions"><a class="btn btn-secondary" href="/logout">Déconnexion</a></div>
</div></body></html>`;
}

function renderAdmin(user, users, page, maxPage, total, csrfToken) {
  const today = new Date().toISOString().split("T")[0];
  const rows = users
    .map(
      (u) => `<tr>
<td>${u.email}${u.is_admin ? ' <span class="badge admin">Admin</span>' : ''}</td>
<td>${u.name || '—'}</td>
<td>${u.created_at ? u.created_at.substring(0, 10) : '—'}</td>
<td>${u.subscription_start || '—'}</td>
<td>${u.subscription_end || '—'}</td>
<td>${u.active && u.subscription_end && u.subscription_end >= today ? '<span class="badge active">Actif</span>' : '<span class="badge inactive">Inactif</span>'}</td>
<td>
<form class="sub-form" method="post" action="/admin/set_sub">
<input type="hidden" name="csrf_token" value="${csrfToken}">
<input type="hidden" name="user_id" value="${u.id}">
<input type="date" name="start" value="${u.subscription_start || ''}" required>
<input type="date" name="end" value="${u.subscription_end || ''}" required>
<button type="submit">Définir</button>
</form>
</td>
</tr>`
    )
    .join("");
  const pagination =
    maxPage > 1
      ? `<div style="display:flex;justify-content:center;align-items:center;gap:12px;margin-top:20px;font-size:.85rem">
${page > 1 ? `<a href="/admin?page=${page - 1}" style="color:#4a6cf7;font-weight:600;text-decoration:none">&larr; Précédent</a>` : ''}
<span style="color:#6b7280">Page ${page} / ${maxPage} (${total} utilisateurs)</span>
${page < maxPage ? `<a href="/admin?page=${page + 1}" style="color:#4a6cf7;font-weight:600;text-decoration:none">Suivant &rarr;</a>` : ''}
</div>`
      : "";
  return `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Administration - Déclaration TVA</title>
<style>${COMMON_CSS}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:40px;width:800px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before,.back{display:none!important}}
.back{position:absolute;top:20px;right:24px;font-size:.82rem;color:#4a6cf7;text-decoration:underline;font-weight:600}
h1{font-size:1.2rem;color:#1a1a2e;margin-bottom:16px}
.table-wrap{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}
table{width:100%;border-collapse:collapse;font-size:.82rem;min-width:600px}
th,td{padding:10px 8px;text-align:left;border-bottom:1px solid #f0f0f5;white-space:nowrap}
th{color:#6b7280;font-weight:600;font-size:.75rem;text-transform:uppercase}
td{color:#1a1a2e}
tr:hover td{background:#fafafe}
.sub-form{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.sub-form input{padding:6px 10px;border:1px solid #e5e7eb;border-radius:6px;font-size:.8rem;outline:0}
.sub-form input:focus{border-color:#667eea}
.sub-form button{padding:6px 12px;border:none;border-radius:6px;font-size:.8rem;font-weight:600;cursor:pointer;color:#fff;background:#059669}
.badge{display:inline-block;padding:2px 8px;border-radius:12px;font-size:.75rem;font-weight:600}
.badge.active{background:#ecfdf5;color:#059669}
.badge.inactive{background:#fef2f2;color:#ef4444}
.badge.admin{background:#eef1ff;color:#667eea}
</style>
</head>
<body>
<div class="card">
<a class="back" href="/">&larr; Accueil</a>
<h1>Administration &mdash; Gestion des abonnements</h1>
<div class="table-wrap"><table>
<thead><tr><th>Email</th><th>Nom</th><th>Inscription</th><th>Début abon.</th><th>Fin abon.</th><th>Statut</th><th>Action</th></tr></thead>
<tbody>${rows}</tbody>
</table></div>
${pagination}
</div></body></html>`;
}

// --- Static pages ---
const HELP_HTML = `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Comment ça marche - Déclaration TVA</title>
<style>${COMMON_CSS}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:40px;width:620px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before,.back{display:none!important}}
.back{position:absolute;top:20px;right:24px;font-size:.82rem;color:#4a6cf7;text-decoration:underline;font-weight:600}
h1{font-size:1.3rem;color:#1a1a2e;margin-bottom:16px;letter-spacing:-.02em}
.step{display:flex;gap:16px;margin-bottom:18px;align-items:flex-start}
.step-num{flex-shrink:0;width:32px;height:32px;background:linear-gradient(135deg,#667eea,#764ba2);border-radius:50%;color:#fff;display:flex;align-items:center;justify-content:center;font-size:.85rem;font-weight:700}
.step-content{flex:1}
.step-content h2{font-size:.95rem;color:#1a1a2e;margin-bottom:3px}
.step-content p{font-size:.82rem;color:#6b7280;line-height:1.5}
.step-content code{background:#f3f4f6;padding:1px 5px;border-radius:4px;font-size:.78rem}
.note{margin-top:16px;padding:14px;background:#fffbeb;border:1px solid #fde68a;border-radius:10px;font-size:.8rem;color:#92400e;line-height:1.5}
</style>
</head>
<body>
<div class="card">
<a class="back" href="/">&larr; Retour</a>
<h1>Comment ça marche</h1>
<div class="step"><div class="step-num">1</div><div class="step-content"><h2>Téléchargez le modèle</h2><p>Depuis la page d'accueil, cliquez sur <strong>"Télécharger le modèle Excel"</strong> pour obtenir un fichier vierge avec les bonnes colonnes et un exemple.</p></div></div>
<div class="step"><div class="step-num">2</div><div class="step-content"><h2>Remplissez votre relevé</h2><p>Ouvrez le fichier dans Excel ou LibreOffice et saisissez vos données : numéro de facture, fournisseur, montant HT, taux de TVA, etc. La colonne <strong>TAUX</strong> accepte la valeur en décimal (ex: <code>0.20</code> pour 20%) ou en pourcentage (ex: <code>20</code> pour 20%).</p></div></div>
<div class="step"><div class="step-num">3</div><div class="step-content"><h2>Renseignez la société</h2><p>Sur la page d'accueil, saisissez le <strong>nom de la société</strong>, l'<strong>identifiant fiscal</strong>, l'<strong>année</strong> et la <strong>période</strong> (mensuelle ou trimestrielle).</p></div></div>
<div class="step"><div class="step-num">4</div><div class="step-content"><h2>Importez le fichier</h2><p>Glissez-déposez votre fichier Excel rempli dans la zone prévue, ou cliquez pour le sélectionner.</p></div></div>
<div class="step"><div class="step-num">5</div><div class="step-content"><h2>Générez la déclaration</h2><p>Cliquez sur <strong>"Lancer la conversion"</strong>. Le fichier est analysé, validé, puis converti au format XML de déclaration TVA.</p></div></div>
<div class="step"><div class="step-num">6</div><div class="step-content"><h2>Téléchargez le résultat</h2><p>Une page de confirmation affiche le nombre de lignes converties. Cliquez sur <strong>"Télécharger"</strong> pour récupérer le fichier <code>.zip</code> contenant le XML. Une <strong>authentification</strong> vous sera demandée si ce n'est pas fait au début.</p></div></div>
<div class="note"><strong>Modes de règlement acceptés :</strong> ESPECE/ESPECES (1), CHEQUE (2), PRELEVEMENT (3), VIREMENT (4), EFFET (5), COMPENSATION (6), AUTRES (7). Tout autre mode non reconnu &rarr; 8.</div>
</div></body></html>`;

const ABOUT_HTML = `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>À propos - Déclaration TVA</title>
<style>${COMMON_CSS}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:40px;width:620px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before,.back,.footer-links{display:none!important}}
.back{position:absolute;top:20px;right:24px;font-size:.82rem;color:#4a6cf7;text-decoration:underline;font-weight:600}
h1{font-size:1.3rem;color:#1a1a2e;margin-bottom:16px;letter-spacing:-.02em}
p{font-size:.88rem;color:#6b7280;line-height:1.7;margin-bottom:12px}
.footer-links{margin-top:24px;padding-top:16px;border-top:1px solid #f0f0f5;font-size:.8rem;color:#9ca3af;text-align:center}
.footer-links a{color:#4a6cf7;text-decoration:underline;font-weight:600}
</style>
</head>
<body>
<div class="card">
<a class="back" href="/">&larr; Retour</a>
<h1>À propos</h1>
<p>Cette application permet de convertir un relevé de déductions TVA au format Excel (.<strong>xls</strong> ou .<strong>xlsx</strong>) en un fichier XML conforme à la structure attendue par l'administration fiscale.</p>
<p>Développée en Node.js avec le framework <strong>Express</strong>, elle utilise la bibliothèque <strong>SheetJS</strong> pour la lecture des fichiers Excel, et génère un fichier XML structuré selon le modèle <code>DeclarationReleveDeduction</code>.</p>
<p>Le résultat est livré sous forme d'archive <code>.zip</code> contenant le fichier XML prêt à être transmis.</p>
<div style="margin-top:18px;padding:16px;background:#eef1ff;border-left:4px solid #667eea;border-radius:10px;font-size:.85rem;color:#1a1a2e;line-height:1.6">
<strong style="font-size:.9rem">Notre engagement</strong><br>
Ce programme est conforme au cahier de charges de la <strong>Direction Générale des Impôts</strong>. Nous nous engageons à le maintenir à jour pour garantir sa conformité avec les évolutions fiscales.
</div>
<div style="margin-top:18px;padding:16px;background:linear-gradient(135deg,#667eea,#764ba2);border-radius:10px;font-size:.88rem;color:#fff;line-height:1.6;text-align:center">
Fini les complications&thinsp;! Avec notre application cloud, transformez vos données TVA en XML en un clic. Profitez d'une liberté totale&thinsp;: aucune installation, aucune restriction, juste la simplicité et la rapidité. Essayez dès aujourd'hui et découvrez une nouvelle façon de gérer vos obligations fiscales.
</div>
<div style="margin-top:18px;padding:20px 16px;background:linear-gradient(135deg,#0f0c29,#302b63,#24243e);border-radius:10px;font-size:.9rem;color:#e0e0ff;line-height:1.7;text-align:center">
<strong style="font-size:1rem;color:#fff">Où que vous soyez, votre TVA vous suit.</strong><br>
La liberté numérique au service de vos obligations fiscales.<br>
Transformez la contrainte en liberté&#160;: TVA en XML, sans effort.
</div>
<div class="footer-links"><a href="/">Accueil</a> &middot; <a href="/aide">Comment ça marche</a> &middot; <a href="/contact">Contact</a></div>
</div></body></html>`;

const CONTACT_HTML = `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Contact - Déclaration TVA</title>
<style>${COMMON_CSS}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:40px;width:540px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before,.back,.footer-links{display:none!important}}
.back{position:absolute;top:20px;right:24px;font-size:.82rem;color:#4a6cf7;text-decoration:underline;font-weight:600}
h1{font-size:1.3rem;color:#1a1a2e;margin-bottom:16px;letter-spacing:-.02em}
.contact-item{display:flex;align-items:center;gap:12px;padding:14px 16px;background:#f9fafb;border-radius:10px;margin-bottom:10px}
.contact-item .label{font-size:.82rem;font-weight:600;color:#374151;min-width:80px}
.contact-item .value{font-size:.88rem;color:#4a6cf7;font-weight:500}
.footer-links{margin-top:24px;padding-top:16px;border-top:1px solid #f0f0f5;font-size:.8rem;color:#9ca3af;text-align:center}
.footer-links a{color:#4a6cf7;text-decoration:underline;font-weight:600}
</style>
</head>
<body>
<div class="card">
<a class="back" href="/">&larr; Retour</a>
<h1>Contact</h1>
<p style="font-size:.88rem;color:#6b7280;line-height:1.6;margin-bottom:6px">Pour activer votre compte, veuillez nous contacter par email pour convenir des modalités de paiement.</p>
<p style="font-size:1.05rem;color:#1a1a2e;font-weight:700;margin-bottom:16px">Abonnement mensuel : 50 DH seulement</p>
<div class="contact-item"><span class="label">Email</span><span class="value">contact@tva2xml.ma</span></div>
<div class="contact-item"><span class="label">Téléphone</span><span class="value"> - </span></div>
<div class="contact-item"><span class="label">Adresse</span><span class="value">Agadir, Maroc</span></div>
<div class="footer-links"><a href="/">Accueil</a> &middot; <a href="/aide">Comment ça marche</a> &middot; <a href="/apropos">À propos</a> &middot; <a href="/confidentialite">Confidentialité</a></div>
</div></body></html>`;

const PRIVACY_HTML = `<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Confidentialité - Déclaration TVA</title>
<style>${COMMON_CSS}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:40px;width:680px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media(max-width:640px){.card{width:100%!important}}@media(max-width:480px){.card{padding:28px 20px}}@media(orientation:landscape) and (max-height:500px){body{padding:8px}.card{padding:20px 16px!important;width:100%!important}}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before,.back,.footer-links{display:none!important}}
.back{position:absolute;top:20px;right:24px;font-size:.82rem;color:#4a6cf7;text-decoration:underline;font-weight:600}
h1{font-size:1.3rem;color:#1a1a2e;margin-bottom:16px;letter-spacing:-.02em}
h2{font-size:1rem;color:#1a1a2e;margin:18px 0 6px}
p{font-size:.85rem;color:#6b7280;line-height:1.7;margin-bottom:8px}
ul{margin:4px 0 12px 20px}
li{font-size:.85rem;color:#6b7280;line-height:1.7}
.footer-links{margin-top:24px;padding-top:16px;border-top:1px solid #f0f0f5;font-size:.8rem;color:#9ca3af;text-align:center}
.footer-links a{color:#4a6cf7;text-decoration:underline;font-weight:600}
</style>
</head>
<body>
<div class="card">
<a class="back" href="/">&larr; Accueil</a>
<h1>Politique de confidentialité</h1>
<p><strong>Dernière mise à jour :</strong> mai 2026</p>
<h2>1. Données collectées</h2>
<p>Cette application collecte les données suivantes lors de l'inscription et de l'utilisation :</p>
<ul>
<li>Adresse email (obligatoire pour la création du compte)</li>
<li>Nom (optionnel)</li>
<li>Fichiers Excel importés pour la conversion (traités en mémoire, supprimés après conversion)</li>
</ul>
<h2>2. Finalité du traitement</h2>
<p>Les données sont utilisées uniquement pour :</p>
<ul>
<li>Gérer votre accès et votre abonnement au service</li>
<li>Convertir vos fichiers Excel en XML de déclaration TVA</li>
</ul>
<h2>3. Stockage et sécurité</h2>
<p>Les mots de passe sont hachés avec l'algorithme <strong>bcrypt</strong>. Aucun mot de passe n'est stocké en clair. Les fichiers temporaires sont supprimés immédiatement après la conversion.</p>
<h2>4. Partage des données</h2>
<p>Aucune donnée personnelle n'est transmise à des tiers. Les fichiers convertis ne sont pas conservés après téléchargement.</p>
<h2>5. Vos droits</h2>
<p>Conformément au RGPD, vous disposez d'un droit d'accès, de rectification et de suppression de vos données. Pour l'exercer, contactez-nous via la page <a href="/contact">Contact</a>.</p>
<h2>6. Cookies</h2>
<p>Cette application utilise uniquement des cookies de session nécessaires au fonctionnement (authentification). Aucun cookie tiers ou de suivi n'est utilisé.</p>
<div class="footer-links"><a href="/">Accueil</a> &middot; <a href="/contact">Contact</a> &middot; <a href="/apropos">À propos</a></div>
</div></body></html>`;

const NOT_FOUND_HTML = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>404 - Page non trouvée</title><style>${COMMON_CSS}
.card{background:rgba(255,255,255,.97);border-radius:20px;padding:48px;width:480px;max-width:100%;box-shadow:0 25px 60px rgba(0,0,0,.5);position:relative;overflow:hidden;text-align:center}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#667eea,#764ba2)}
@media print{body{background:#fff!important;padding:0!important}.card{box-shadow:none!important;border:1px solid #ddd!important}.card::before{display:none!important}}
h1{font-size:3rem;color:#667eea;margin-bottom:4px}
p{color:#6b7280;font-size:.95rem;margin-bottom:20px;line-height:1.5}
.btn{display:inline-block;padding:12px 28px;border-radius:10px;font-size:.9rem;font-weight:600;text-decoration:none;color:#fff;background:linear-gradient(135deg,#667eea,#764ba2)}
</style></head><body><div class="card"><h1>404</h1><p>La page que vous cherchez n'existe pas ou a été déplacée.</p><a class="btn" href="/">Retour à l'accueil</a></div></body></html>`;

process.on("exit", flushDb);
process.on("SIGINT", () => { flushDb(); process.exit(0); });
process.on("SIGTERM", () => { flushDb(); process.exit(0); });

process.on("uncaughtException", err => {
  console.error("Uncaught Exception:", err);
});
process.on("unhandledRejection", err => {
  console.error("Unhandled Rejection:", err);
});

if (require.main === module) {
  initDb().then(() => {
    app.listen(PORT, "0.0.0.0", () => {
      console.log(`Serveur démarré sur http://0.0.0.0:${PORT}`);
      console.log(`Mode: ${process.env.NODE_ENV || 'development'}`);
    });
  }).catch(e => {
    console.error("DB init failed:", e);
    process.exit(1);
  });
}

// initDb géré au démarrage

module.exports = { app, initDb };
