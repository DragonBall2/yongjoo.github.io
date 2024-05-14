const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
app.use(bodyParser.json());

app.post('/submit', (req, res) => {
    const { name, email } = req.body;

    // Create a new workbook and a new worksheet
    let workbook;
    let worksheet;
    const filePath = './data.xlsx';

    if (fs.existsSync(filePath)) {
        workbook = XLSX.readFile(filePath);
        worksheet = workbook.Sheets['Sheet1'];
    } else {
        workbook = XLSX.utils.book_new();
        worksheet = XLSX.utils.aoa_to_sheet([['Name', 'Email']]);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    }

    // Find the next empty row
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const nextRow = range.e.r + 1;

    // Add the new data
    worksheet[`A${nextRow + 1}`] = { v: name };
    worksheet[`B${nextRow + 1}`] = { v: email };
    worksheet['!ref'] = XLSX.utils.encode_range(range.s, { c: range.e.c, r: nextRow });

    // Write to the file
    XLSX.writeFile(workbook, filePath);

    res.sendStatus(200);
});

app.listen(3000, () => {
    console.log('Server is running on port 3000');
});
