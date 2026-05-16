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
app.use(express.json());

// ... (le reste de ton code)
