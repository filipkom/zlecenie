const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

// Przykładowe dane firm
const firmy = [
    { nazwa: 'Firma A', adres: 'ul. Przykładowa 1, 00-001 Miasto', nip: '123-456-32-18' },
    { nazwa: 'Firma B', adres: 'ul. Testowa 2, 00-002 Miasto', nip: '987-654-32-18' },
    // Dodaj więcej firm
];

app.use(cors());
app.use(bodyParser.json());

// Endpoint do pobierania danych firmy na podstawie nazwy
app.post('/api/firmy', (req, res) => {
    const { nazwa } = req.body;
    const firma = firmy.find(f => f.nazwa.toLowerCase() === nazwa.toLowerCase());

    if (firma) {
        res.json(firma);
    } else {
        res.status(404).json({ message: 'Nie znaleziono firmy' });
    }
});

// Endpoint do generowania pliku Excel
app.post('/api/generuj-excel', (req, res) => {
    const { nazwa } = req.body;
    const firma = firmy.find(f => f.nazwa.toLowerCase() === nazwa.toLowerCase());

    if (firma) {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet([firma]);
        XLSX.utils.book_append_sheet(wb, ws, 'Firmy');

        const filePath = path.join(__dirname, `${firma.nazwa}.xlsx`);
        XLSX.writeFile(wb, filePath);

        // Usuń plik po przegraniu
        fs.unlink(filePath, (err) => {
            if (err) console.error(err);
        });

        res.download(filePath, `${firma.nazwa}.xlsx`);
    } else {
        res.status(404).json({ message: 'Nie znaleziono firmy' });
    }
});

app.listen(port, () => {
    console.log(`Server działa na http://localhost:${port}`);
});