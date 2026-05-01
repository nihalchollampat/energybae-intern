require('dotenv').config();
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const Groq = require('groq-sdk');
const Tesseract = require('tesseract.js');
const sharp = require('sharp');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const port = process.env.PORT || 5001;

app.use(cors());
app.use(express.json());

// Ensure uploads directory exists (crucial for Render deployments!)
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
}
const upload = multer({ dest: 'uploads/' });

// Groq Setup
const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });

// Supabase Setup
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);

app.get('/', (req, res) => {
    res.send('Energybae Solar Data Feeder is Live! ☀️🔋');
});

// Health check endpoint for deployment platforms (like Render)
app.head('/', (req, res) => {
    res.sendStatus(200);
});

async function extractDataWithGroq(filePath) {
    try {
        console.log(`[OCR] Preprocessing: ${filePath}`);
        const processedPath = filePath + '_processed.png';
        const isPdf = filePath.toLowerCase().endsWith('.pdf');
        
        if (!isPdf) {
            await sharp(filePath)
                .resize(2000) // Upscale for better OCR
                .grayscale()
                .linear(1.5, -0.2) // Increase contrast
                .toFile(processedPath);
        }

        console.log(`[OCR] Starting Tesseract...`);
        const ocrTarget = isPdf ? filePath : (fs.existsSync(processedPath) ? processedPath : filePath);
        
        // Use local traineddata if possible (Tesseract.js automatic lookup in CWD)
        const { data: { text } } = await Tesseract.recognize(ocrTarget, 'eng', {
            logger: m => console.log(`[Tesseract] ${m.status}: ${Math.round(m.progress * 100)}%`)
        });
        
        if (fs.existsSync(processedPath)) fs.unlinkSync(processedPath);
        
        if (!text || text.trim().length < 20) {
            throw new Error('OCR failed to extract enough text. Please ensure the bill is clear and readable.');
        }

        console.log('[AI] Sending to Groq (Llama 3.3)...');
        const prompt = `
            You are a high-precision Data Entry Assistant. 
            I will provide messy OCR text from an MSEDCL electricity bill.
            
            GOAL: Extract the consumer details and the 12-month Consumption History table.
            
            FIELDS TO IDENTIFY:
            1. Consumer Name
            2. Consumer Number (12 digits)
            3. Fixed Charges
            4. Sanctioned Load (kW)
            5. Connection Type
            
            TABLE DATA:
            Reconstruct the 12-month "Consumption History" (Month, Units, Bill Amount).
            
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

        const content = chatCompletion.choices[0].message.content;
        return JSON.parse(content);
    } catch (err) {
        console.error('Error in AI extraction:', err);
        throw new Error(`AI Extraction failed: ${err.message}`);
    }
}

app.post('/api/upload', upload.array('bills', 2), async (req, res) => {
    try {
        console.log('[Upload] Request received');
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'No files uploaded' });
        }

        const templatePath = path.join(__dirname, 'Solar_Load_Template.xlsx');
        if (!fs.existsSync(templatePath)) {
            throw new Error('Excel template not found.');
        }

        // Process sequentially to save memory on free tiers (prevents 502 OOM crashes)
        const extractedResultsData = [];
        for (let i = 0; i < req.files.length; i++) {
            const file = req.files[i];
            console.log(`[Process] Starting extraction for bill ${i + 1}/${req.files.length}...`);
            const data = await extractDataWithGroq(file.path);
            extractedResultsData.push({ file, data, i });
        }
        
        const extractedResults = extractedResultsData.map(res => res.data);

        // --- 2. Save extracted data to Database First ---
        console.log('[DB] Saving extracted data to Supabase database...');
        for (const { data } of extractedResultsData) {
            try {
                const { error } = await supabase.from('bill_analysis').insert([{
                    consumer_name: data.name,
                    consumer_number: data.consumer_no,
                    sanctioned_load: parseFloat(data.sanctioned_load) || 0,
                    fixed_charges: parseFloat(data.fixed_charges) || 0,
                    connection_type: data.connection_type,
                    monthly_consumption: data.monthly_units
                }]);
                if (error) throw error;
                console.log(`[DB] Successfully saved data for consumer: ${data.name || 'Unknown'}`);
            } catch (dbErr) {
                console.error(`[DB] Error saving data to Supabase:`, dbErr.message);
                // If database save is critical, you could throw the error here to stop execution
            }
        }

        // --- 3. Generate Excel File with the result ---
        console.log('[Excel] Loading Excel template and generating report...');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        const sheet = workbook.getWorksheet(1);

        for (const { file, data, i } of extractedResultsData) {
            const colPrefix = i === 0 ? 'D' : 'H';
            const colMonth = i === 0 ? 'B' : 'G';
            
            sheet.getCell(`${colPrefix}1`).value = data.name;
            sheet.getCell(`${colPrefix}2`).value = data.consumer_no;
            sheet.getCell(`${colPrefix}3`).value = parseFloat(data.fixed_charges) || 0;
            sheet.getCell(`${colPrefix}4`).value = parseFloat(data.sanctioned_load) || 0;
            sheet.getCell(`${colPrefix}5`).value = data.connection_type;

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
            if (fs.existsSync(file.path)) fs.unlinkSync(file.path);
        }

        // Cleanup for single file uploads
        if (req.files.length === 1) {
            for (let r = 1; r <= 5; r++) sheet.getCell(`H${r}`).value = null;
            ['G','H','I','J'].forEach(c => sheet.getCell(`${c}8`).value = null);
            for (let r = 9; r <= 20; r++) {
                ['G','H','I','J'].forEach(c => sheet.getCell(`${c}${r}`).value = null);
            }
        }

        const outputFileName = `Energybae_Analysis_${Date.now()}.xlsx`;
        const outputPath = path.join(__dirname, 'uploads', outputFileName);
        
        await workbook.xlsx.writeFile(outputPath);
        console.log(`[Excel] Report saved successfully at ${outputPath}`);

        res.json({ message: 'Success', data: extractedResults, downloadUrl: `/api/download/${outputFileName}` });

    } catch (error) {
        console.error('Final Error:', error);
        res.status(500).json({ error: 'Extraction failed', details: error.message });
    }
});

app.get('/api/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(filePath)) res.download(filePath);
    else res.status(404).send('File not found');
});

app.listen(port, '0.0.0.0', () => console.log(`Energybae Data Feeder running at http://localhost:${port}`));
