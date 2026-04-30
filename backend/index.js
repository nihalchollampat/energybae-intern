require('dotenv').config();
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const Tesseract = require('tesseract.js');
const Groq = require('groq-sdk');

const app = express();
const port = process.env.PORT || 5001;

app.use(cors());
app.use(express.json());

const upload = multer({ dest: 'uploads/' });

const apiKey = process.env.GROQ_API_KEY;
const groq = new Groq({ apiKey: apiKey || 'INVALID_KEY' });

async function extractDataWithGroq(filePath) {
    console.log(`Starting OCR for ${filePath}...`);
    const { data: { text } } = await Tesseract.recognize(filePath, 'eng');
    
    const prompt = `
        Analyze this MSEDCL bill OCR text. You MUST extract the "Consumption History" table (usually contains 12 months of data).
        Return ONLY a JSON object:
        {
            "consumerName": "Full name",
            "consumerNumber": "12 digit ID",
            "sanctionedLoad": number,
            "fixedCharges": number,
            "connectionType": "string",
            "monthlyConsumption": [
                { "month": "Month Year", "units": number, "billAmount": number, "unitCost": number }
            ]
        }
        
        OCR TEXT:
        ${text}
    `;

    const chatCompletion = await groq.chat.completions.create({
        messages: [{ role: 'user', content: prompt }],
        model: 'llama-3.3-70b-versatile',
        response_format: { type: 'json_object' }
    });

    return JSON.parse(chatCompletion.choices[0].message.content);
}

// Handle multiple files
app.post('/api/upload', upload.array('bills', 2), async (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'No files uploaded' });
        }

        const templatePath = path.join(__dirname, 'Solar_Load_Template.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        const sheet = workbook.getWorksheet(1);

        const extractedResults = [];

        for (let i = 0; i < req.files.length; i++) {
            const file = req.files[i];
            const data = await extractDataWithGroq(file.path);
            extractedResults.push(data);

            // Mapping: Bill 1 -> Column D, Bill 2 -> Column H
            const colData = i === 0 ? 'D' : 'H';
            const colUnits = i === 0 ? 'D' : 'H';
            const colMonth = i === 0 ? 'C' : 'G';
            const colAmount = i === 0 ? 'E' : 'I';
            const colCost = i === 0 ? 'F' : 'J';

            sheet.getCell(`${colData}1`).value = data.consumerName;
            sheet.getCell(`${colData}2`).value = data.consumerNumber;
            sheet.getCell(`${colData}3`).value = parseFloat(data.fixedCharges) || 0;
            sheet.getCell(`${colData}4`).value = parseFloat(data.sanctionedLoad) || 0;
            sheet.getCell(`${colData}5`).value = data.connectionType;

            if (data.monthlyConsumption) {
                data.monthlyConsumption.forEach((item, idx) => {
                    const row = 9 + idx;
                    if (row <= 20) {
                        sheet.getCell(`${colMonth}${row}`).value = item.month;
                        sheet.getCell(`${colUnits}${row}`).value = parseFloat(item.units) || 0;
                        sheet.getCell(`${colAmount}${row}`).value = parseFloat(item.billAmount) || 0;
                        sheet.getCell(`${colCost}${row}`).value = parseFloat(item.unitCost) || 0;
                    }
                });
            }
            fs.unlinkSync(file.path);
        }

        const outputFileName = `Multi_Solar_Load_${Date.now()}.xlsx`;
        const outputPath = path.join(__dirname, 'uploads', outputFileName);
        await workbook.xlsx.writeFile(outputPath);

        res.json({
            message: `Successfully processed ${req.files.length} bills`,
            data: extractedResults,
            downloadUrl: `/api/download/${outputFileName}`
        });

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Failed to process bills', details: error.message });
    }
});

app.get('/api/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(filePath)) res.download(filePath);
    else res.status(404).send('File not found');
});

app.listen(port, () => {
    console.log(`Backend (Multi-Bill) running at http://localhost:${port}`);
});
