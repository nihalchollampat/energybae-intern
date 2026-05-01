require('dotenv').config();
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const port = process.env.PORT || 5001;

app.use(cors());
app.use(express.json());

const upload = multer({ dest: 'uploads/' });

// Gemini Setup
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// Supabase Setup
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);

app.get('/', (req, res) => {
    res.send('Energybae Solar Data Feeder is Live! ☀️🔋');
});

async function extractDataWithGemini(filePath) {
    console.log(`Step 1: Processing file with Gemini 1.5 Flash...`);
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    const fileBuffer = fs.readFileSync(filePath);
    // Basic MIME type detection
    let mimeType = 'image/jpeg';
    if (filePath.toLowerCase().endsWith('.pdf')) mimeType = 'application/pdf';
    else if (filePath.toLowerCase().endsWith('.png')) mimeType = 'image/png';

    const prompt = `
        You are a high-precision Data Entry Assistant for an Energy Solar company.
        Extract the consumer details and the 12-month Consumption History table from this MSEDCL bill.
        
        GOAL: Reconstruct the consumer details and the 12-month Consumption History table accurately.
        
        FIELDS TO EXTRACT:
        1. Consumer Name: (Full name mentioned in the bill)
        2. Consumer Number: (12-digit unique number)
        3. Fixed Charges: (Amount in Rupees)
        4. Sanctioned Load: (Load in kW)
        5. Connection Type: (Tariff/Connection category)
        
        TABLE EXTRACTION:
        Find the "Consumption History" table (usually last 12 months).
        Columns: Month, Units, Bill Amount.
        Reconstruct all 12 rows. Use null for missing values.
        
        OUTPUT: Return ONLY a valid JSON object (no markdown, no preamble):
        {
            "name": "string",
            "consumer_no": "string",
            "fixed_charges": number,
            "sanctioned_load": number,
            "connection_type": "string",
            "monthly_units": [
                { "month": "string", "units": number, "amount": number }
            ]
        }
    `;

    const result = await model.generateContent([
        {
            inlineData: {
                data: fileBuffer.toString("base64"),
                mimeType: mimeType
            }
        },
        prompt
    ]);

    const response = await result.response;
    const text = response.text();
    
    // Clean up the response in case Gemini adds markdown code blocks
    const cleanedText = text.replace(/```json|```/g, '').trim();
    return JSON.parse(cleanedText);
}

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
            const data = await extractDataWithGemini(file.path);
            extractedResults.push(data);

            // Lead Management
            await supabase.from('bill_analysis').insert([{
                consumer_name: data.name,
                consumer_number: data.consumer_no,
                sanctioned_load: parseFloat(data.sanctioned_load) || 0,
                fixed_charges: parseFloat(data.fixed_charges) || 0,
                connection_type: data.connection_type,
                monthly_consumption: data.monthly_units
            }]);

            const colPrefix = i === 0 ? 'D' : 'H';
            const colMonth = i === 0 ? 'B' : 'G';
            
            sheet.getCell(`${colPrefix}1`).value = data.name;
            sheet.getCell(`${colPrefix}2`).value = data.consumer_no;
            sheet.getCell(`${colPrefix}3`).value = parseFloat(data.fixed_charges) || 0;
            sheet.getCell(`${colPrefix}4`).value = parseFloat(data.sanctioned_load) || 0;
            sheet.getCell(`${colPrefix}5`).value = data.connection_type;

            // Wipe raw unit rows (9-20)
            for (let r = 9; r <= 20; r++) {
                ['C','D','E','F','G','H','I','J'].forEach(c => sheet.getCell(`${c}${r}`).value = null);
            }

            if (data.monthly_units) {
                data.monthly_units.forEach((item, idx) => {
                    const row = 9 + idx;
                    if (row <= 20) {
                        sheet.getCell(`${colMonth}${row}`).value = item.month;
                        sheet.getCell(`${colPrefix}${row}`).value = parseFloat(item.units) || 0;
                        const amountCol = i === 0 ? 'E' : 'I';
                        const costCol = i === 0 ? 'F' : 'J';
                        sheet.getCell(`${amountCol}${row}`).value = parseFloat(item.amount) || 0;
                        if (item.units > 0) {
                           sheet.getCell(`${costCol}${row}`).value = (item.amount / item.units).toFixed(2);
                        }
                    }
                });
            }
            fs.unlinkSync(file.path);
        }
        
        // If only 1 bill was uploaded, clear the 2nd user's side to keep it clean
        if (req.files.length === 1) {
            // Clear Top Details (H1:H5)
            for (let r = 1; r <= 5; r++) sheet.getCell(`H${r}`).value = null;
            // Clear Headers (G8:J8)
            ['G','H','I','J'].forEach(c => sheet.getCell(`${c}8`).value = null);
            // Clear Table (G9:J20)
            for (let r = 9; r <= 20; r++) {
                ['G','H','I','J'].forEach(c => sheet.getCell(`${c}${r}`).value = null);
            }
            // Clear Result Labels/Formulas for User 2 (G24:J28)
            for (let r = 24; r <= 28; r++) {
                ['G','H','I','J'].forEach(c => sheet.getCell(`${c}${r}`).value = null);
            }
            // Clear Totals at the bottom
            sheet.getCell('B29').value = null;
            sheet.getCell('D29').value = null;
            sheet.getCell('B30').value = null;
            sheet.getCell('D30').value = null;
        }

        const outputFileName = `Energybae_Analysis_${Date.now()}.xlsx`;
        const outputPath = path.join(__dirname, 'uploads', outputFileName);
        await workbook.xlsx.writeFile(outputPath);

        res.json({ message: 'Success', data: extractedResults, downloadUrl: `/api/download/${outputFileName}` });

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Extraction failed', details: error.message });
    }
});

app.get('/api/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(filePath)) res.download(filePath);
    else res.status(404).send('File not found');
});

app.listen(port, () => console.log(`Energybae Data Feeder running at http://localhost:${port}`));
