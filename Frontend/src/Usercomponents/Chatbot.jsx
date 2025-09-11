import React, { useState, useRef } from 'react';
import './App3.css';

function Chatbot() {
  const [file, setFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [headerOption, setHeaderOption] = useState('Row 1 (0)');
  const [removeUnnamed, setRemoveUnnamed] = useState(true);
  const [removeLastRow, setRemoveLastRow] = useState(true);
  const [maxRowsDisplay, setMaxRowsDisplay] = useState(10);
  const [dataInfo, setDataInfo] = useState(null);
  const [userQuestion, setUserQuestion] = useState('');
  const [chatHistory, setChatHistory] = useState([]);
  const [showCode, setShowCode] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const fileInputRef = useRef(null);

  const sampleQuestions = [
    "What's the total sum of [column_name]?",
    "Show me the top 5 rows by [column_name]",
    "What's the average value in [column_name]?",
    "How many unique values are in [column_name]?",
    "Filter rows where [column_name] > 100",
    "What's the correlation between [col1] and [col2]?"
  ];

  const handleFileUpload = async (e) => {
    const uploadedFile = e.target.files[0];
    if (!uploadedFile) return;

    setFile(uploadedFile);
    setError(null);

    const formData = new FormData();
    formData.append('file', uploadedFile);

    try {
      const response = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to process file');
      }

      const data = await response.json();
      if (data.success) {
        setSheetNames(data.sheet_names);
        setSelectedSheet(data.sheet_names[0]);
      }
    } catch (err) {
      setError(err.message || 'Failed to upload file');
      console.error('Upload error:', err);
    }
  };

  const loadData = async () => {
    if (!file) return;

    setIsLoading(true);
    setError(null);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('selected_sheet', selectedSheet);
    formData.append('header_option', headerOption);
    formData.append('remove_unnamed', removeUnnamed);
    formData.append('remove_last_row', removeLastRow);

    try {
      const response = await fetch('/api/analyze', {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to analyze data');
      }

      const data = await response.json();
      setDataInfo(data.data_info);
    } catch (err) {
      setError(err.message || 'Failed to analyze data');
      console.error('Analysis error:', err);
    } finally {
      setIsLoading(false);
    }
  };

  const askQuestion = async () => {
    if (!userQuestion || !file) return;

    setIsLoading(true);
    setError(null);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('selected_sheet', selectedSheet);
    formData.append('header_option', headerOption);
    formData.append('remove_unnamed', removeUnnamed);
    formData.append('remove_last_row', removeLastRow);
    formData.append('user_question', userQuestion);

    try {
      const response = await fetch('/api/ask', {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to get answer');
      }

      const data = await response.json();
      setChatHistory([...chatHistory, {
        question: data.question,
        answer: data.result,
        code: data.python_code
      }]);
      setUserQuestion('');
    } catch (err) {
      setError(err.message || 'Failed to get answer');
      console.error('Question error:', err);
    } finally {
      setIsLoading(false);
    }
  };

  const handleSampleQuestion = (question) => {
    setUserQuestion(question);
  };

  const clearChatHistory = () => {
    setChatHistory([]);
  };

  return (
    <div className="app">
      <header className="app-header">
        <h1>📊 Chat with Your Excel File</h1>
      </header>

      <div className="app-container">
        <div className="sidebar">
          <div className="sidebar-section">
            <h2>⚙️ Options</h2>
            <label>
              <input
                type="checkbox"
                checked={showCode}
                onChange={() => setShowCode(!showCode)}
              />
              Always show generated code
            </label>
            <div className="slider-container">
              <label>Max rows to display: {maxRowsDisplay}</label>
              <input
                type="range"
                min="5"
                max="50"
                value={maxRowsDisplay}
                onChange={(e) => setMaxRowsDisplay(parseInt(e.target.value))}
              />
            </div>
          </div>

          <div className="sidebar-section">
            <h2>📤 Upload File</h2>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls"
              style={{ display: 'none' }}
            />
            <button onClick={() => fileInputRef.current.click()}>
              {file ? file.name : 'Choose Excel File'}
            </button>
            {file && (
              <>
                <select
                  value={selectedSheet}
                  onChange={(e) => setSelectedSheet(e.target.value)}
                >
                  {sheetNames.map((sheet) => (
                    <option key={sheet} value={sheet}>{sheet}</option>
                  ))}
                </select>
                <div className="radio-group">
                  <label>Header row:</label>
                  <label>
                    <input
                      type="radio"
                      value="Row 1 (0)"
                      checked={headerOption === 'Row 1 (0)'}
                      onChange={() => setHeaderOption('Row 1 (0)')}
                    />
                    Row 1 (0)
                  </label>
                  <label>
                    <input
                      type="radio"
                      value="Row 2 (1)"
                      checked={headerOption === 'Row 2 (1)'}
                      onChange={() => setHeaderOption('Row 2 (1)')}
                    />
                    Row 2 (1)
                  </label>
                  <label>
                    <input
                      type="radio"
                      value="No header"
                      checked={headerOption === 'No header'}
                      onChange={() => setHeaderOption('No header')}
                    />
                    No header
                  </label>
                </div>
                <div className="checkbox-group">
                  <label>
                    <input
                      type="checkbox"
                      checked={removeUnnamed}
                      onChange={() => setRemoveUnnamed(!removeUnnamed)}
                    />
                    Remove unnamed columns
                  </label>
                  <label>
                    <input
                      type="checkbox"
                      checked={removeLastRow}
                      onChange={() => setRemoveLastRow(!removeLastRow)}
                    />
                    Remove last row (if summary)
                  </label>
                </div>
                <button onClick={loadData} disabled={isLoading}>
                  {isLoading ? 'Loading...' : 'Load Data'}
                </button>
              </>
            )}
          </div>
        </div>

        <div className="main-content">
          {error && <div className="error-message">{error}</div>}

          {!file ? (
            <div className="welcome-message">
              <h2>👆 Please upload an Excel file to get started!</h2>
              <div className="info-box">
                <h3>ℹ️ How to use this app</h3>
                <ol>
                  <li><strong>Upload</strong> your Excel file using the file uploader</li>
                  <li><strong>Select</strong> the sheet you want to analyze</li>
                  <li><strong>Adjust</strong> data cleaning options if needed</li>
                  <li><strong>Ask questions</strong> in natural language about your data</li>
                  <li><strong>View</strong> the results and generated Python code</li>
                </ol>
                <h4>Example questions:</h4>
                <ul>
                  <li>"What's the total revenue?"</li>
                  <li>"Show me the top 10 customers by sales"</li>
                  <li>"What's the average order value?"</li>
                  <li>"How many unique products do we have?"</li>
                </ul>
              </div>
            </div>
          ) : (
            <>
              {dataInfo && (
                <>
                  <div className="success-message">
                    Uploaded: {file.name} | Sheet: {selectedSheet} | Shape: {dataInfo.shape[0]} rows × {dataInfo.shape[1]} columns
                  </div>

                  <div className="data-preview">
                    <h2>📋 Data Preview</h2>
                    <div className="table-container">
                      <table>
                        <thead>
                          <tr>
                            {dataInfo.columns.map((col) => (
                              <th key={col}>{col}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {dataInfo.preview.map((row, i) => (
                            <tr key={i}>
                              {dataInfo.columns.map((col) => (
                                <td key={col}>{row[col]}</td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <div className="data-info">
                    <div className="info-column">
                      <h3>📊 Columns</h3>
                      <ul>
                        {dataInfo.columns.map((col) => (
                          <li key={col}>{col}</li>
                        ))}
                      </ul>
                    </div>
                    <div className="info-column">
                      <h3>📈 Data Types</h3>
                      <ul>
                        {Object.entries(dataInfo.dtypes).map(([col, dtype]) => (
                          <li key={col}>{col}: {dtype}</li>
                        ))}
                      </ul>
                    </div>
                  </div>
                </>
              )}

              {chatHistory.length > 0 && (
                <div className="chat-history">
                  <h2>💬 Chat History</h2>
                  {chatHistory.map((chat, i) => (
                    <div key={i} className="chat-item">
                      <details>
                        <summary>Question {i + 1}: {chat.question.substring(0, 50)}...</summary>
                        <div className="chat-content">
                          <p><strong>You:</strong> {chat.question}</p>
                          {chat.answer.type === 'dataframe' ? (
                            <>
                              <p><strong>Bot:</strong></p>
                              <div className="table-container">
                                <table>
                                  <thead>
                                    <tr>
                                      {Object.keys(chat.answer.data[0]).map((col) => (
                                        <th key={col}>{col}</th>
                                      ))}
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {chat.answer.data.map((row, i) => (
                                      <tr key={i}>
                                        {Object.values(row).map((val, j) => (
                                          <td key={j}>{val}</td>
                                        ))}
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            </>
                          ) : (
                            <p><strong>Bot:</strong> {chat.answer.data}</p>
                          )}
                          {showCode && (
                            <div className="code-block">
                              <pre>{chat.code}</pre>
                            </div>
                          )}
                        </div>
                      </details>
                    </div>
                  ))}
                  <button onClick={clearChatHistory}>🗑️ Clear Chat History</button>
                </div>
              )}

              <div className="question-section">
                <h2>❓ Ask Your Question</h2>
                <div className="question-input">
                  <input
                    type="text"
                    value={userQuestion}
                    onChange={(e) => setUserQuestion(e.target.value)}
                    placeholder="e.g., What's the average sales? Show top 5 customers by revenue"
                    onKeyPress={(e) => e.key === 'Enter' && askQuestion()}
                  />
                  <button onClick={askQuestion} disabled={isLoading || !userQuestion}>
                    {isLoading ? 'Analyzing...' : 'Ask'}
                  </button>
                </div>

                <div className="sample-questions">
                  <h3>💡 Sample Questions</h3>
                  <div className="sample-buttons">
                    {sampleQuestions.map((question, i) => (
                      <button
                        key={i}
                        onClick={() => handleSampleQuestion(question)}
                        className="sample-btn"
                      >
                        {question}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      </div>

      <footer className="app-footer">
        Made with ❤️ using Flask and React
      </footer>
    </div>
  );
}

export default Chatbot;
