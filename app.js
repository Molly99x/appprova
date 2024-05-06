const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const moment = require('moment');

const app = express();

// Configurazione di multer con storage personalizzato per mantenere le estensioni dei file
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, './uploads/')
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, uniqueSuffix + path.extname(file.originalname));  // mantenere l'estensione del file originale
    }
});
const upload = multer({ storage: storage });

// Creazione della directory 'uploads' se non esiste
if (!fs.existsSync('./uploads')) {
    fs.mkdirSync('./uploads', { recursive: true });
}

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));
app.use(bodyParser.urlencoded({ extended: true }));

// Route per la homepage
app.get('/', (req, res) => {
    res.render('index');
});

// Route per visualizzare i dettagli dei turni
app.get('/turno/:tipo', async (req, res) => {
    const tipo = req.params.tipo;
    const workbook = new ExcelJS.Workbook();
    const controlli = [];

    try {
        await workbook.xlsx.readFile(path.join(__dirname, 'data', 'controllidaeseguire.xlsx'));
        const worksheet = workbook.getWorksheet(1);

        let isHeader = true;  // Saltare la prima riga (intestazioni)
        worksheet.eachRow((row) => {
            if (isHeader) {
                isHeader = false;
            } else {
                if (row.getCell(1).text.toUpperCase() === tipo.toUpperCase()) {
                    controlli.push({
                        turno_id: row.getCell(1).text,
                        controllo: row.getCell(2).text,
                        descrizione: row.getCell(3).text
                    });
                }
            }
        });

        res.render('turno', { tipo, controlli });
    } catch (error) {
        console.error("Errore durante la lettura del foglio di lavoro:", error);
        res.status(500).send("Errore durante la lettura del file Excel: " + error.message);
    }
});


app.get('/turno/:tipo', async (req, res) => {
    try {
        const tipo = req.params.tipo;
        const workbook = new ExcelJS.Workbook();
        const controlli = [];

        await workbook.xlsx.readFile(path.join(__dirname, 'data', 'controllidaeseguire.xlsx'));
        const worksheet = workbook.getWorksheet(1);

        let isHeader = true;  // Saltare la prima riga (intestazioni)
        worksheet.eachRow((row) => {
            if (isHeader) {
                isHeader = false;
            } else {
                if (row.getCell(1).text.toUpperCase() === tipo.toUpperCase()) {
                    controlli.push({
                        turno_id: row.getCell(1).text,
                        controllo: row.getCell(2).text,
                        descrizione: row.getCell(3).text
                    });
                }
            }
        });

        res.render('turno', { tipo, controlli });
    } catch (error) {
        console.error("Errore durante la lettura del foglio di lavoro:", error);
        res.status(500).send("Errore durante la lettura del file Excel: " + error.message);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server in ascolto sulla porta ${PORT}`);
});
