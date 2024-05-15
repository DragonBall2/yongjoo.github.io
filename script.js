// AWS.config.update({
//     accessKeyId: 'AKIAZQ3DT756WD7MDMGT',
//     secretAccessKey: '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',
//     region: 'ap-northeast-2'
// });

// const s3 = new AWS.S3();
// const BUCKET_NAME = 'yongjoo-bucket';
// const FILE_NAME = 'data.xlsx';

// document.getElementById('userForm').addEventListener('submit', function(event) {
//     event.preventDefault();

//     const name = document.getElementById('name').value;
//     const email = document.getElementById('email').value;

//     s3.getObject({ Bucket: BUCKET_NAME, Key: FILE_NAME }, function(err, data) {
//         if (err) {
//             console.log('Error getting object:', err);
//         } else {
//             console.log('Object retrieved:', data);
//         }

//         let workbook;
//         if (err && err.code === 'NoSuchKey') {
//             workbook = XLSX.utils.book_new();
//             workbook.SheetNames.push("Sheet1");
//             const worksheet_data = [["ID", "Name", "Email"]];
//             const worksheet = XLSX.utils.aoa_to_sheet(worksheet_data);
//             workbook.Sheets["Sheet1"] = worksheet;
//         } else if (data) {
//             const arrayBuffer = data.Body.buffer;
//             workbook = XLSX.read(arrayBuffer, { type: 'array' });
//         } else {
//             console.log(err);
//             return;
//         }

//         const worksheet = workbook.Sheets["Sheet1"];
//         const newRow = [worksheet['!ref'].split(':')[1].match(/\d+/)[0], name, email];
//         XLSX.utils.sheet_add_aoa(worksheet, [newRow], { origin: -1 });

//         const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

//         s3.putObject({
//             Bucket: BUCKET_NAME,
//             Key: FILE_NAME,
//             Body: wbout,
//             ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
//         }, function(err, data) {
//             if (err) {
//                 console.log('Error putting object:', err);
//                 alert('Failed to submit data.');
//             } else {
//                 console.log('Object put successfully:', data);
//                 alert('Data submitted successfully!');
//             }
//         });
//     });
// });


document.addEventListener('DOMContentLoaded', function() {
    AWS.config.update({
        accessKeyId: 'AKIAZQ3DT756WD7MDMGT',
        secretAccessKey: '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',
        region: 'ap-northeast-2'
    });

    const s3 = new AWS.S3();
    const BUCKET_NAME = 'yongjoo-bucket';
    const FILE_NAME = 'data.xlsx';

    document.getElementById('userForm').addEventListener('submit', function(event) {
        event.preventDefault();

        const name = document.getElementById('name').value;
        const email = document.getElementById('email').value;

        s3.getObject({ Bucket: BUCKET_NAME, Key: FILE_NAME }, function(err, data) {
            if (err) {
                console.log('Error getting object:', err);
                return;
            }

            let workbook;
            if (!data || !data.Body) {
                workbook = XLSX.utils.book_new();
                workbook.SheetNames.push("Sheet1");
                const worksheet_data = [["ID", "Name", "Email"]];
                const worksheet = XLSX.utils.aoa_to_sheet(worksheet_data);
                workbook.Sheets["Sheet1"] = worksheet;
            } else {
                const arrayBuffer = data.Body.buffer;
                workbook = XLSX.read(arrayBuffer, { type: 'array' });
            }

            const worksheet = workbook.Sheets["Sheet1"];
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const maxRow = range.e.r + 1;
            const newRow = [maxRow, name, email];
            XLSX.utils.sheet_add_aoa(worksheet, [newRow], { origin: -1 });

            const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

            s3.putObject({
                Bucket: BUCKET_NAME,
                Key: FILE_NAME,
                Body: wbout,
                ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }, function(err, data) {
                if (err) {
                    console.log('Error putting object:', err);
                    alert('Failed to submit data.');
                } else {
                    console.log('Object put successfully:', data);
                    alert('Data submitted successfully!');
                }
            });
        });
    });
});
