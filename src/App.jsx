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

  const handleProcessChange = (type) => {
    setProcessType(type);
    setError(null);
    setResultCsv(null);
    if (inputRef.current) inputRef.current.value = "";
  };

  const findColumnKey = (keys, candidates) => {
    const cleanedKeys = keys.map(k => String(k).trim());
    
    // 1. Exact match (case-insensitive)
    for (const cand of candidates) {
      const match = cleanedKeys.find(k => k.toLowerCase() === cand.toLowerCase());
      if (match) return match;
    }
    
    // 2. Normalized match (ignoring spaces, underscores, hyphens, dots)
    const normalize = str => str.toLowerCase().replace(/[\s_\-\.]+/g, '');
    for (const cand of candidates) {
      const normCand = normalize(cand);
      const match = cleanedKeys.find(k => normalize(k) === normCand);
      if (match) return match;
    }
    
    // 3. Substring match
    for (const cand of candidates) {
      const normCand = normalize(cand);
      const match = cleanedKeys.find(k => normalize(k).includes(normCand));
      if (match) return match;
    }
    
    return null;
  };

  const handleFile = (file) => {
    setError(null);
    setProcessing(true);
    setResultCsv(null);

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
        const results = processRandomization(json, processType);
        
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

  const processRandomization = (data, currentProcessType) => {
    if (!data || data.length === 0) {
      throw new Error("The uploaded file contains no data.");
    }

    // Clean column keys to handle trailing spaces just in case
    const cleanedData = data.map(row => {
      const cleanRow = {};
      Object.keys(row).forEach(key => {
        cleanRow[String(key).trim()] = row[key];
      });
      return cleanRow;
    });

    // Gather all unique keys across rows
    const allKeys = Array.from(new Set(cleanedData.flatMap(row => Object.keys(row))));

    // Detect Emp Name column
    const empCandidates = ['Emp Name', 'Employee Name', 'EmpName', 'Agent Name', 'AgentName', 'Agent'];
    const empKey = findColumnKey(allKeys, empCandidates);
    if (!empKey) {
      throw new Error("Could not find the employee name column. Please ensure your file has an 'Emp Name' column.");
    }

    // Detect Target column based on processType
    let targetKey = null;
    let outputPrefix = '';

    if (currentProcessType === 'voice') {
      const voiceCandidates = [
        'Mobile Number',
        'Mobile No',
        'Mobile No.',
        'Mobile',
        'MobileNumber',
        'Mobile_Number',
        'phone_number',
        'Phone Number',
        'Phone No',
        'Phone No.',
        'Phone',
        'Contact Number',
        'Contact No',
        'Contact'
      ];
      targetKey = findColumnKey(allKeys, voiceCandidates);
      if (!targetKey) {
        throw new Error("Could not find mobile number column in file. For Voice Process, ensure your file contains a 'Mobile Number' (or Mobile No, phone_number) column.");
      }
      outputPrefix = 'Mobile Number';
    } else {
      const nonVoiceCandidates = [
        'Offer Id',
        'Offer ID',
        'Offer id',
        'OfferId',
        'Offer_Id',
        'Offer'
      ];
      targetKey = findColumnKey(allKeys, nonVoiceCandidates);
      if (!targetKey) {
        throw new Error("Could not find offer ID column in file. For Non-Voice Process, ensure your file contains an 'Offer Id' column.");
      }
      outputPrefix = 'Offer Id';
    }

    // Group unique target values by Emp Name
    const groups = {};
    cleanedData.forEach(row => {
      const rawEmp = row[empKey];
      const rawTarget = row[targetKey];
      
      if (rawEmp !== undefined && rawEmp !== null && String(rawEmp).trim() !== "") {
        const empName = String(rawEmp).trim();
        if (!groups[empName]) groups[empName] = new Set();
        if (rawTarget !== undefined && rawTarget !== null && String(rawTarget).trim() !== "") {
          groups[empName].add(String(rawTarget).trim());
        }
      }
    });

    const empNames = Object.keys(groups);
    if (empNames.length === 0) {
      const expectedCol = currentProcessType === 'voice' ? 'Mobile Number' : 'Offer Id';
      throw new Error(`No valid data found. Ensure your file has valid entries for 'Emp Name' and '${expectedCol}' columns.`);
    }

    // Randomize and select 5 records per employee
    const results = [];
    empNames.forEach(empName => {
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
        rowResult[`${outputPrefix}${i + 1}`] = selected[i];
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

  const downloadSampleFile = (type = 'xlsx') => {
    const isVoice = processType === 'voice';
    const targetCol = isVoice ? 'Mobile Number' : 'Offer Id';

    const sampleData = isVoice ? [
      { 'Emp Name': 'Rahul Sharma', [targetCol]: '9876543210' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: '9876543211' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: '9876543212' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: '9876543213' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: '9876543214' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: '9876543215' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: '9123456780' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: '9123456781' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: '9123456782' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: '9123456783' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: '9123456784' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: '9123456785' },
    ] : [
      { 'Emp Name': 'Rahul Sharma', [targetCol]: 'OFFER1001' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: 'OFFER1002' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: 'OFFER1003' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: 'OFFER1004' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: 'OFFER1005' },
      { 'Emp Name': 'Rahul Sharma', [targetCol]: 'OFFER1006' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: 'OFFER2001' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: 'OFFER2002' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: 'OFFER2003' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: 'OFFER2004' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: 'OFFER2005' },
      { 'Emp Name': 'Pooja Verma', [targetCol]: 'OFFER2006' },
    ];

    const ws = XLSX.utils.json_to_sheet(sampleData);
    ws['!cols'] = [{ wch: 18 }, { wch: 20 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');

    if (type === 'csv') {
      const csvOut = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob([csvOut], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement("a");
      const url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.setAttribute("download", `sample_${processType}_template.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } else {
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement("a");
      const url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.setAttribute("download", `sample_${processType}_template.xlsx`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }
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
                type="button"
                className={`btn ${processType === 'voice' ? '' : 'reset-btn'}`}
                style={{margin: 0, padding: '0.8rem 2rem', fontSize: '1rem'}}
                onClick={() => handleProcessChange('voice')}
              >
                Voice Process
              </button>
              <button 
                type="button"
                className={`btn ${processType === 'non-voice' ? '' : 'reset-btn'}`}
                style={{margin: 0, padding: '0.8rem 2rem', fontSize: '1rem'}}
                onClick={() => handleProcessChange('non-voice')}
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
              <div className="upload-hint">
                Upload {processType === 'voice' ? 'Voice' : 'Non-Voice'} File (XLSX, XLS, CSV)
              </div>
              <div style={{fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '0.5rem'}}>
                Required columns: <strong>Emp Name</strong> &amp; <strong>{processType === 'voice' ? 'Mobile Number' : 'Offer Id'}</strong>
              </div>
              {error && <div style={{color: '#ef4444', marginTop: '1rem', fontWeight: '500'}}>{error}</div>}
            </form>

            <div className="sample-template-box">
              <div className="sample-template-info">
                <span className="sample-template-icon">📥</span>
                <div>
                  <div style={{ color: 'var(--text-primary)', fontWeight: '600' }}>
                    Download Sample {processType === 'voice' ? 'Voice' : 'Non-Voice'} Template
                  </div>
                  <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
                    Paste your data into this template ({processType === 'voice' ? 'Emp Name & Mobile Number' : 'Emp Name & Offer Id'}), then upload above.
                  </div>
                </div>
              </div>
              <div className="sample-btn-group">
                <button 
                  type="button" 
                  className="sample-download-btn"
                  onClick={() => downloadSampleFile('xlsx')}
                  title="Download Excel (.xlsx) Template"
                >
                  <span>📊</span> Excel (.xlsx)
                </button>
                <button 
                  type="button" 
                  className="sample-download-btn"
                  onClick={() => downloadSampleFile('csv')}
                  title="Download CSV (.csv) Template"
                >
                  <span>📝</span> CSV (.csv)
                </button>
              </div>
            </div>
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
              Your {processType === 'voice' ? 'Voice' : 'Non-Voice'} process data has been successfully generated ({processType === 'voice' ? '5 mobile numbers' : '5 offer IDs'} per employee).
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
