const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
app.use(bodyParser.json());

app.post('/submit', (req, res) => {
    const { name, email } = req.body;

    // 파일 경로 설정
    const filePath = './data.xlsx';

    // 워크북과 워크시트를 설정
    let workbook;
    let worksheet;

    // data.xlsx 파일이 이미 존재하는지 확인
    if (fs.existsSync(filePath)) {
        // 파일이 존재하면 읽어들임
        workbook = XLSX.readFile(filePath);
        worksheet = workbook.Sheets['Sheet1'];
    } else {
        // 파일이 존재하지 않으면 새 워크북과 워크시트 생성
        workbook = XLSX.utils.book_new();
        worksheet = XLSX.utils.aoa_to_sheet([['Name', 'Email']]);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    }

    // 워크시트의 범위 설정
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const nextRow = range.e.r + 1;

    // 새로운 데이터 추가
    worksheet[`A${nextRow + 1}`] = { v: name };
    worksheet[`B${nextRow + 1}`] = { v: email };
    worksheet['!ref'] = XLSX.utils.encode_range(range.s, { c: range.e.c, r: nextRow });

    // 파일 쓰기
    XLSX.writeFile(workbook, filePath);

    res.sendStatus(200);
});

app.listen(3000, () => {
    console.log('Server is running on port 3000');
});