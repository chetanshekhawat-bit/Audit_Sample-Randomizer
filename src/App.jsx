import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';

function App() {
  const [processType, setProcessType] = useState('voice');
  const [dragActive, setDragActive] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [resultCsv, setResultCsv] = useState(null);
  const [error, setError] = useState(null);
  const inputRef = useRef(null);

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e) => {
    e.preventDefault();
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  };

  const handleFile = (file) => {
    setError(null);
    setProcessing(true);
    setResultCsv(null);

    const targetColumn = processType === 'voice' ? 'phone_number' : 'Offer Id';

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Use the first sheet or the 'Data' sheet if it exists
        const sheetName = workbook.SheetNames.includes('Data') ? 'Data' : workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to JSON
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        // Process data
        const results = processRandomization(json, targetColumn);
        
        if (results.length === 0) {
          throw new Error(`No valid data found. Ensure your file has 'Emp Name' and '${targetColumn}' columns.`);
        }
        
        // Convert back to CSV
        const worksheetOut = XLSX.utils.json_to_sheet(results);
        const csvOut = XLSX.utils.sheet_to_csv(worksheetOut);
        
        setResultCsv(csvOut);
      } catch (err) {
        console.error(err);
        setError(err.message || "An error occurred while processing the file.");
      } finally {
        setProcessing(false);
      }
    };
    
    reader.readAsArrayBuffer(file);
  };

  const processRandomization = (data, targetColumn) => {
    // Clean column keys to handle trailing spaces just in case
    const cleanedData = data.map(row => {
      const cleanRow = {};
      Object.keys(row).forEach(key => {
        cleanRow[String(key).trim()] = row[key];
      });
      return cleanRow;
    });

    // Group by Emp Name
    const groups = {};
    cleanedData.forEach(row => {
      const empName = row['Emp Name'];
      const targetValue = row[targetColumn];
      
      if (empName) {
        if (!groups[empName]) groups[empName] = new Set();
        if (targetValue) groups[empName].add(targetValue);
      }
    });

    // Randomize and select 5
    const results = [];
    Object.keys(groups).forEach(empName => {
      const targetValues = Array.from(groups[empName]);
      const selected = [];
      const tempValues = [...targetValues];
      const limit = 5;

      for (let i = 0; i < limit; i++) {
        if (tempValues.length > 0) {
          const randIdx = Math.floor(Math.random() * tempValues.length);
          selected.push(tempValues.splice(randIdx, 1)[0]);
        } else {
          selected.push("");
        }
      }

      const rowResult = { 'Emp Name': empName };
      for (let i = 0; i < limit; i++) {
        rowResult[`${targetColumn}${i + 1}`] = selected[i];
      }
      results.push(rowResult);
    });

    return results;
  };

  const downloadCsv = () => {
    if (!resultCsv) return;
    const blob = new Blob([resultCsv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", `randomized_${processType}_output.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const reset = () => {
    setResultCsv(null);
    setError(null);
    if (inputRef.current) inputRef.current.value = "";
  };

  return (
    <div className="app-container">
      <div className="glass-card">
        <div className="header">
          <h1>Data Randomizer</h1>
          <p>Randomly select 5 records per employee for your process.</p>
        </div>

        {!resultCsv && !processing && (
          <>
            <div className="process-selector" style={{display: 'flex', justifyContent: 'center', gap: '1rem', marginBottom: '2rem'}}>
              <button 
                className={`btn ${processType === 'voice' ? '' : 'reset-btn'}`}
                style={{margin: 0, padding: '0.8rem 2rem', fontSize: '1rem'}}
                onClick={() => setProcessType('voice')}
              >
                Voice Process
              </button>
              <button 
                className={`btn ${processType === 'non-voice' ? '' : 'reset-btn'}`}
                style={{margin: 0, padding: '0.8rem 2rem', fontSize: '1rem'}}
                onClick={() => setProcessType('non-voice')}
              >
                Non-Voice Process
              </button>
            </div>

            <form 
              className={`upload-zone ${dragActive ? "drag-active" : ""}`}
              onDragEnter={handleDrag}
              onDragLeave={handleDrag}
              onDragOver={handleDrag}
              onDrop={handleDrop}
              onClick={() => inputRef.current.click()}
            >
              <input 
                ref={inputRef}
                type="file" 
                className="hidden-input" 
                accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
                onChange={handleChange}
              />
              <span className="upload-icon">📄</span>
              <div className="upload-text">Click to upload or drag and drop</div>
              <div className="upload-hint">Upload {processType === 'voice' ? 'Voice' : 'Non-Voice'} File (XLSX, XLS, CSV)</div>
              {error && <div style={{color: '#ef4444', marginTop: '1rem'}}>{error}</div>}
            </form>
          </>
        )}

        {processing && (
          <div className="result-container" style={{padding: '3rem 0'}}>
            <div className="upload-icon" style={{animation: 'spin 2s linear infinite'}}>⚙️</div>
            <h2>Processing Data...</h2>
          </div>
        )}

        {resultCsv && !processing && (
          <div className="result-container">
            <span className="success-icon">✨</span>
            <h2>Randomization Complete!</h2>
            <p style={{color: 'var(--text-secondary)', marginBottom: '2rem'}}>
              Your {processType === 'voice' ? 'Voice' : 'Non-Voice'} process data has been successfully generated.
            </p>
            
            <button className="btn" onClick={downloadCsv}>
              Download CSV
            </button>
            <button className="btn reset-btn" onClick={reset}>
              Process Another
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
