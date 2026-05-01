const Tesseract = require('tesseract.js');
const path = require('path');

async function testOCR() {
    console.log('Testing Tesseract.js OCR...');
    // We'll use a simple image URL or a local file if we had one.
    // Since we don't have a bill image, let's just see if the engine initializes.
    try {
        const worker = await Tesseract.createWorker('eng');
        console.log('Tesseract Worker initialized successfully!');
        await worker.terminate();
        console.log('OCR Engine is healthy.');
    } catch (error) {
        console.error('OCR Error:', error.message);
        process.exit(1);
    }
}

testOCR();
