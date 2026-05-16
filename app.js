const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const initSqlJs = require('sql.js');

const app = express();
const upload = multer({ dest: 'uploads/' });
const PORT = process.env.PORT || 5000;

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Créer dossiers
['uploads','public'].forEach(d => { if (!fs.existsSync(d)) fs.mkdirSync(d); });

// Page d'accueil
app.get('/', (req, res) => {
  res.send(`
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Convertisseur TVA Maroc</title>
<style>
body{font-family:Arial;background:#f5f7fa;margin:0;padding:40px}
.container{max-width:800px;margin:auto;background:white;padding:30px;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.1)}
h1{color:#2c3e50;text-align:center}
.upload{border:2px dashed #3498db;padding:40px;text-align:center;margin:20px 0;border-radius:8px}
button{background:#27ae60;color:white;padding:12px 30px;border:none;border-radius:6px;font-size:16px;cursor:pointer}
button:hover{background:#219a52}
.result{margin-top:20px;padding:15px;background:#e8f5e9;border-radius:6px;display:none}
</style>
</head>
<body>
<div class="container">
<h1>🇲🇦 Convertisseur TVA Excel → XML</h1>
<div class="upload">
<input type="file" id="file" accept=".xlsx,.xls" style="display:none">
<button onclick="document.getElementById('file').click()">Choisir fichier Excel</button>
<p id="filename"></p>
</div>
<button id="convert" style="display:none;width:100%" onclick="convert()">Convertir en XML</button>
<div id="result" class="result"></div>
</div>
<script>
const fileInput=document.getElementById('file');
const filename=document.getElementById('filename');
const convertBtn=document.getElementById('convert');
const result=document.getElementById('result');
let selectedFile;

fileInput.onchange=e=>{selectedFile=e.target.files[0];filename.textContent=selectedFile.name;convertBtn.style.display='block';};

async function convert(){
  const form=new FormData();
  form.append('excel',selectedFile);
  convertBtn.textContent='Conversion...';
  const res=await fetch('/convert',{method:'POST',body:form});
  const data=await res.json();
  if(data.success){
    result.style.display='block';
    result.innerHTML='<h3>✅ Fichier prêt</h3><a href="'+data.downloadUrl+'" download><button>Télécharger '+data.filename+'</button></a>';
  }else{
    result.style.display='block';
    result.innerHTML='<h3>❌ Erreur</h3><p>'+data.error+'</p>';
  }
  convertBtn.textContent='Convertir en XML';
}
</script>
</body>
</html>
  `);
});

// Conversion
app.post('/convert', upload.single('excel'), async (req, res) => {
  try {
    if (!req.file) throw new Error('Aucun fichier');

    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Chercher période
    let periode = '';
    for(let i=0;i<10;i++){
      for(let j=0;j<5;j++){
        const val = data[i]?.[j];
        if(val && /\\d{2}\\/\\d{4}/.test(val.toString())){
          periode = val.toString();
          break;
        }
      }
    }

    // Générer XML simplifié
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<DeclarationReleveDeduction>
  <periode>${periode}</periode>
  <identifiantFiscal>12345678</identifiantFiscal>
  <raisonSociale>SOCIETE EXEMPLE</raisonSociale>
  <factures>
    ${data.slice(5).map(row => row[0]? `<facture><num>${row[0]}</num><montant>${row[5]||0}</montant></facture>` : '').join('')}
  </factures>
</DeclarationReleveDeduction>`;

    const filename = `TVA_${Date.now()}.xml`;
    const filepath = path.join('public', filename);
    fs.writeFileSync(filepath, xml);

    fs.unlinkSync(req.file.path);

    res.json({ success: true, filename, downloadUrl: `/${filename}` });
  } catch (e) {
    res.json({ success: false, error: e.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Serveur démarré sur le port ${PORT}`);
});
