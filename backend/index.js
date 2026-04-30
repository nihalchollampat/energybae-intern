require('dotenv').config();
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const Tesseract = require('tesseract.js');
const Groq = require('groq-sdk');
const { createClient } = require('@supabase/supabase-js');
const sharp = require('sharp');

const app = express();
const port = process.env.PORT || 5001;

app.use(cors());
app.use(express.json());

const upload = multer({ dest: 'uploads/' });

// Groq Setup (Llama 3.3 for intelligent cleaning)
const apiKey = process.env.GROQ_API_KEY;
const groq = new Groq({ apiKey: apiKey || 'INVALID_KEY' });

// Supabase Setup
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);

app.get('/', (req, res) => {
    res.send('Energybae Solar Data Feeder is Live! ☀️🔋');
});

async function extractDataWithGroq(filePath) {
    console.log(`Step 1: Preprocessing Image ${filePath}...`);
    const processedPath = filePath + '_processed.jpg';
    
    // Grayscale, Normalize, and Sharpen for better OCR
    await sharp(filePath)
        .grayscale()
        .normalize()
        .sharpen()
        .toFile(processedPath);

    console.log(`Step 2: OCR Extraction...`);
    const { data: { text } } = await Tesseract.recognize(processedPath, 'eng');
    
    // Cleanup processed file
    if (fs.existsSync(processedPath)) fs.unlinkSync(processedPath);
    
    console.log('Step 3: Intelligent Data Cleaning with Llama 3.3...');
    const prompt = `
        You are a high-precision Data Entry Assistant for an Energy Solar company.
        I will provide you with messy OCR text from an MSEDCL electricity bill.
        
        GOAL: Reconstruct the consumer details and the 12-month Consumption History table.
        
        IDENTIFY THESE FIELDS:
        1. Consumer Name: (Look for "Name of Consumer" or similar)
        2. Consumer Number: (12 digits, look for "Consumer No.")
        3. Fixed Charges: (Look for "Fixed Charges" in the bill summary)
        4. Sanctioned Load: (Look for "Sanctioned Load" or "Connected Load" in kW)
        5. Connection Type: (Look for "Connection Type" or "Tariff Category")
        
        TABLE EXTRACTION (CRITICAL):
        Find the table titled "Consumption History" or "Last 12 Months Consumption".
        It usually has columns like: Month | Units | Bill Amount.
        Reconstruct all 12 rows. If a value is missing, use null.
        
        OUTPUT: Return ONLY a valid JSON object:
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
