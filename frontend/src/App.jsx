import React, { useState, useRef } from 'react';
import axios from 'axios';

const API_URL = import.meta.env.VITE_API_URL || 'https://energybae-intern.onrender.com/api';

// Simple SVG Icons
const UploadIcon = () => (
  <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
  </svg>
);

const CheckIcon = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="20 6 9 17 4 12"/>
  </svg>
);

const DownloadIcon = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
  </svg>
);

function App() {
  const [files, setFiles] = useState([]);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const fileInputRef = useRef(null);

  const handleFileChange = (e) => {
    const selectedFiles = Array.from(e.target.files);
    if (selectedFiles.length > 0) {
      setFiles(selectedFiles);
      setError(null);
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (files.length === 0) return;

    setLoading(true);
    setError(null);

    const formData = new FormData();
    files.forEach(file => {
      formData.append('bills', file);
    });

    try {
      const response = await axios.post(`${API_URL}/upload`, formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      setResult(response.data);
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to process the bill. Please check your backend and Gemini API key.');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="container" style={{ maxWidth: '1000px', margin: '0 auto', padding: '40px 20px' }}>
      {/* Header */}
      <header style={{ textAlign: 'center', marginBottom: '60px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px', marginBottom: '16px' }}>
          <div style={{ width: '40px', height: '40px', background: '#2e7d32', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'white', fontWeight: 'bold', fontSize: '24px' }}>E</div>
          <h1 style={{ fontSize: '2.5rem', fontWeight: '700', color: '#1b5e20', letterSpacing: '-0.02em' }}>ENERGYBAE</h1>
        </div>
        <p style={{ fontSize: '1.1rem', color: '#666', fontWeight: '400' }}>Empowering People with Renewable Energy Solutions</p>
      </header>

      <main>
        <div className="glass-card" style={{ padding: '40px', marginBottom: '40px' }}>
          <h2 style={{ marginBottom: '24px', fontSize: '1.5rem', fontWeight: '600' }}>Solar Load Calculator — Automation</h2>
          
          {!result ? (
            <form onSubmit={handleSubmit}>
              <div 
                className="file-upload-zone"
                onClick={() => fileInputRef.current.click()}
              >
                <div style={{ color: '#2e7d32', marginBottom: '16px' }}><UploadIcon /></div>
                <p style={{ fontWeight: '500', fontSize: '1.1rem', marginBottom: '8px' }}>
                  {files.length > 0 ? `${files.length} bills selected` : 'Click to upload or drag & drop electricity bills'}
                </p>
                <p style={{ color: '#888', fontSize: '0.9rem' }}>Supports multiple PDFs, JPGs, PNGs (Up to 2 bills for side-by-side analysis)</p>
                <input 
                  type="file" 
                  ref={fileInputRef}
                  onChange={handleFileChange}
                  style={{ display: 'none' }}
                  accept=".pdf,image/*"
                  multiple
                />
              </div>

              {error && (
                <div style={{ marginTop: '20px', padding: '12px', background: '#ffebee', color: '#c62828', borderRadius: '8px', fontSize: '0.9rem' }}>
                  {error}
                </div>
              )}

              <div style={{ marginTop: '32px', display: 'flex', justifyContent: 'center' }}>
                <button 
                  type="submit" 
                  className="btn-primary" 
                  disabled={files.length === 0 || loading}
                  style={{ padding: '14px 40px', fontSize: '1rem' }}
                >
                  {loading ? (
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '12px' }}>
                      <span className="loader" style={{ width: '20px', height: '20px', borderWidth: '2px' }}></span>
                      <span style={{ fontSize: '0.9rem' }}>Processing {files.length} Bills (OCR & AI)...</span>
                    </div>
                  ) : 'Analyze Bills & Generate Excel'}
                </button>
              </div>
            </form>
          ) : (
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '32px', padding: '20px', background: '#f1f8e9', borderRadius: '16px' }}>
                <div style={{ color: '#2e7d32' }}><CheckIcon /></div>
                <div>
                  <h3 style={{ color: '#2e7d32', fontSize: '1.1rem', fontWeight: '600' }}>Extraction Complete!</h3>
                  <p style={{ fontSize: '0.9rem', color: '#555' }}>We've successfully processed {files.length} bills into a side-by-side report.</p>
                </div>
              </div>

              <div style={{ display: 'flex', gap: '16px' }}>
                <a 
                  href={`http://localhost:5001${result.downloadUrl}`}
                  className="btn-primary"
                  style={{ textDecoration: 'none' }}
                >
                  <DownloadIcon /> Download Side-by-Side Excel
                </a>
                <button 
                  onClick={() => {setResult(null); setFiles([]);}}
                  style={{ background: 'none', border: '1px solid #ccc', padding: '12px 24px', borderRadius: '12px', cursor: 'pointer', fontWeight: '500' }}
                >
                  Process New Bills
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Feature List */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '20px' }}>
          <div className="glass-card" style={{ padding: '24px', textAlign: 'center' }}>
            <div style={{ fontSize: '1.5rem', marginBottom: '12px' }}>⚡</div>
            <h4 style={{ marginBottom: '8px' }}>AI Extraction</h4>
            <p style={{ fontSize: '0.85rem', color: '#666' }}>Powered by Groq Llama 3.3 for 99% accuracy on MSEDCL bills.</p>
          </div>
          <div className="glass-card" style={{ padding: '24px', textAlign: 'center' }}>
            <div style={{ fontSize: '1.5rem', marginBottom: '12px' }}>📊</div>
            <h4 style={{ marginBottom: '8px' }}>Auto-Calculation</h4>
            <p style={{ fontSize: '0.85rem', color: '#666' }}>Excel formulas are preserved for ROI and solar sizing.</p>
          </div>
          <div className="glass-card" style={{ padding: '24px', textAlign: 'center' }}>
            <div style={{ fontSize: '1.5rem', marginBottom: '12px' }}>🚀</div>
            <h4 style={{ marginBottom: '8px' }}>Instant Results</h4>
            <p style={{ fontSize: '0.85rem', color: '#666' }}>Reduce processing time from 30 minutes to 30 seconds.</p>
          </div>
        </div>
      </main>

      <footer style={{ marginTop: '80px', textAlign: 'center', color: '#999', fontSize: '0.9rem' }}>
        &copy; 2026 Energybae. Solar Solutions for a Greener Tomorrow.
      </footer>
    </div>
  );
}

export default App;
