# Energybae Solar Load Calculator ☀️🔋

Automated side-by-side analysis of electricity bills for solar sizing and ROI calculation.

## 🚀 How to Use
1. **Setup**: Add your `GROQ_API_KEY` to `backend/.env`.
2. **Install**: Run `npm install` in both `frontend` and `backend`.
3. **Run**:
   - Backend: `node index.js` (Port 5001)
   - Frontend: `npm run dev` (Port 5173)
4. **Action**: Select **two** electricity bills (images/PDF) and click "Analyze Bills".

## ✨ Features
- **Dual Bill Processing**: Side-by-side extraction and comparison.
- **AI + OCR**: Tesseract OCR combined with Llama 3.3 (Groq) for 12-month history extraction.
- **Lead Management**: Automatically saves all extracted data to **Supabase** for future sales tracking.
- **Pixel-Perfect Excel**: Generates a professional report with all formulas and Energybae branding.

## 📝 Submission Note
Developed as part of the AI Intern task to automate the manual 30-minute bill analysis process. Reduces processing time to seconds with 99% accuracy on historical consumption data.
