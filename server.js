const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 3000;

// Body-parser middleware for parsing form data
app.use(bodyParser.urlencoded({ extended: true }));

// Serve the static HTML file
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.post('/save_data', (req, res) => {
    const { 
        name, 
        class: studentClass, 
        barcode_id, 
        destiny, 
        email, 
        phone, 
        parents_phone, 
        current_duty, 
        school_index_number, 
        prefect_unique_id, 
        section 
    } = req.body;

    const newStudent = {
        name, 
        class: studentClass, 
        barcode_id, 
        destiny, 
        email, 
        phone, 
        parents_phone, 
        current_duty, 
        school_index_number, 
        prefect_unique_id, 
        section
    };

    const excelFile = 'students.xlsx';
    
    // Excel තීරු සඳහා අවශ්‍ය නිශ්චිත නම් මෙහි යොදන්න.
    const headers = [
        'name', 'class', 'barcode_id', 'destiny', 'email', 'phone', 
        'parents_phone', 'current_duty', 'school_index_number', 
        'prefect_unique_id', 'section'
    ];

    let data = [];

    // ගොනුව පවතින්නේදැයි පරීක්ෂා කර පවතින දත්ත කියවන්න
    if (fs.existsSync(excelFile)) {
        const workbook = xlsx.readFile(excelFile);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        data = xlsx.utils.sheet_to_json(worksheet);
    } else {
        // ගොනුව නොපවතී නම්, හිස් array එකක් සාදන්න.
        data = [];
    }

    // නව ශිෂ්‍යයාගේ දත්ත JSON object එකක් ලෙස array එකට එකතු කරන්න.
    data.push(newStudent);
    
    // JSON data එක worksheet එකක් බවට හරවන්න.
    const newWorksheet = xlsx.utils.json_to_sheet(data, { header: headers });
    const newWorkbook = xlsx.utils.book_new();
    
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Students');
    xlsx.writeFile(newWorkbook, excelFile);

    res.send('තොරතුරු සාර්ථකව සුරකින ලදී!');
});
// Start the server
app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}`);
});