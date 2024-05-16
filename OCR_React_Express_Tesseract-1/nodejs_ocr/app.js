const Tesseract = require('tesseract.js');
const express = require('express');
const cors = require('cors');
const fileUpload = require('express-fileupload');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const storageDir = path.join(__dirname, 'storage');

app.use(cors());
app.use(fileUpload());
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use('/img', express.static(storageDir));

// Start the server
const PORT = 5000;
app.listen(PORT, () => {
    console.log(`Server active @ port ${PORT}!`);
});

app.post('/upload', async (req, res) => {
    if (req.files && req.files.file) {
        const uploadedFile = req.files.file;
        const fileName = uploadedFile.name;
        const filePath = path.join(storageDir, fileName);

        // Ensure the storage directory exists
        if (!fs.existsSync(storageDir)) {
            fs.mkdirSync(storageDir, { recursive: true });
        }

        uploadedFile.mv(filePath, async (err) => {
            if (err) {
                console.error('Error moving file:', err);
                return res.status(500).send(err);
            }

            console.log('File successfully moved to:', filePath);
            try {
                // Perform OCR using Tesseract.js
                const { data: { text } } = await Tesseract.recognize(
                    filePath,
                    'eng',
                    { logger: m => console.log(m) }
                );

                console.log('OCR Text:', text);

                // Extract relevant data from the OCR text
                const manufacturingDate = extractManufacturingDate(text);
                const batchNumber = extractBatchNumber(text);
                const expiryDate = extractExpiryDate(text);
                const mrp = extractMRP(text);

                // Append data to Excel file
                await appendDataToExcel([manufacturingDate, batchNumber, expiryDate, mrp]);

                res.send({
                    image: `http://localhost:5000/img/${fileName}`,
                    path: `http://localhost:5000/img/${fileName}`,
                    data: { manufacturingDate, batchNumber, expiryDate, mrp }
                });
            } catch (error) {
                console.error('Error during OCR:', error);
                res.status(500).send(error);
            }
        });
    } else {
        console.log('No files uploaded');
        res.status(400).send('No file uploaded.');
    }
});

// Helper functions to extract relevant data from the OCR text
function extractManufacturingDate(text) {
    const dateRegex = /(?:MFG|Manufacturing\s*Date|Mfg\.?\s*Date)[:\-\s]*([0-9]{1,2}[\/\-][0-9]{1,2}[\/\-][0-9]{2,4})/i;
    const match = text.match(dateRegex);
    
    if (match) {
        const dateString = match[1];
        const dateParts = dateString.split(/[\/\-]/);
        
        // Normalize the date format to yyyy/mm/dd
        let year = dateParts[2].length === 2 ? `20${dateParts[2]}` : dateParts[2];
        let month = dateParts[0].padStart(2, '0');
        let day = dateParts[1].padStart(2, '0');
        
        return `${year}/${month}/${day}`;
    }
    
    return 'Not Found';
}

function extractBatchNumber(text) {
    const regex = /\b(?:Batch\s*No\.?|Batch\s*Number|Batch)[:\-\s]*([A-Z0-9\-]+)\b/i;
    const match = text.match(regex);
    return match ? match[1] : 'Not Found';
}

function extractExpiryDate(text) {
    const regex = /\b(?:EXP|Expiry\s*Date|Exp\.?\s*Date)[:\-\s]*([0-9]{2}[\/\-][0-9]{2}[\/\-][0-9]{4})\b/i;
    const match = text.match(regex);
    return match ? match[1] : 'Not Found';
}

function extractMRP(text) {
    const regex = /\b(?:MRP|Price)[:\-\s]*₹?([0-9]+\.[0-9]{2})\b/i;
    const match = text.match(regex);
    return match ? match[1] : 'Not Found';
}
// Function to append data to Excel file
async function appendDataToExcel(data) {
    const excelFilePath = path.join(storageDir, 'ocr_data.xlsx');

    try {
        let workbook;
        let worksheet;
        if (fs.existsSync(excelFilePath)) {
            // If the file exists, load the workbook
            workbook = xlsx.readFile(excelFilePath);
            worksheet = workbook.Sheets[workbook.SheetNames[0]];

            // Find the next empty row in the worksheet
            const range = xlsx.utils.decode_range(worksheet['!ref']);
            const nextRow = range.e.r + 1;

            // Check for date-like structures in the extracted text and populate the "Manufacturing Date" column
            const manufacturingDate = findManufacturingDate(data);
            worksheet[xlsx.utils.encode_cell({ r: nextRow, c: 0 })] = { v: manufacturingDate };

            // Add data to the remaining columns
            data.forEach((value, index) => {
                if (index !== 0) {
                    const cellAddress = xlsx.utils.encode_cell({ r: nextRow, c: index });
                    worksheet[cellAddress] = { v: value };
                }
            });

            // Update the range reference to include the new data
            worksheet['!ref'] = xlsx.utils.encode_range({
                s: { r: 0, c: 0 },
                e: { r: nextRow, c: data.length - 1 }
            });
        } else {
            // If the file doesn't exist, create a new workbook with headers and data
            workbook = xlsx.utils.book_new();
            worksheet = xlsx.utils.aoa_to_sheet([
                ['Manufacturing Date', 'Batch Number', 'Expiry Date', 'MRP'],
                data
            ]);

            xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        }

        // Save the updated workbook to the Excel file
        xlsx.writeFile(workbook, excelFilePath);
    } catch (error) {
        console.error('Error appending data to Excel file:', error);
    }
}

function findManufacturingDate(data) {
    for (const value of data) {
        if (isDateLike(value)) {
            return value;
        }
    }
    return 'Not Found';
}

function isDateLike(value) {
    const dateRegex = /(\d{1,2}\/\d{1,2}\/\d{4})/; // Date format: dd/mm/yyyy
    return dateRegex.test(value);
}