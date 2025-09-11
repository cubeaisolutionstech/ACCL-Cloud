import React, { useState } from 'react';

const months = [
  "April", "May", "June", "July", 
  "August", "September", "October", "November",
  "December", "January", "February", "March"
];

const API_BASE_URL = '';

const App = () => {
  const [files, setFiles] = useState({});
  const [skipFirstRow, setSkipFirstRow] = useState(false);
  const [preview, setPreview] = useState(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState(null);
  const [previewGenerated, setPreviewGenerated] = useState(false);

  const handleFileChange = (month, e) => {
    if (e.target.files[0]) {
      const newFiles = { ...files, [month]: e.target.files[0] };
      setFiles(newFiles);
      setPreviewGenerated(false);
      setMessage(null);
    }
  };

  const handleDelete = (month) => {
    const newFiles = { ...files };
    delete newFiles[month];
    setFiles(newFiles);
    setPreviewGenerated(false);
    setPreview(null);
    setMessage(null);
  };

  const handleProcess = async () => {
    if (Object.keys(files).length === 0) {
      setMessage('Please select at least one file to process.');
      return;
    }

    setLoading(true);
    setMessage(null);
    setPreview(null);

    try {
      const formData = new FormData();
      
      Object.entries(files).forEach(([month, file]) => {
        formData.append(month, file);
      });
      
      formData.append('skipFirstRow', skipFirstRow.toString());

      const response = await fetch(`${API_BASE_URL}/api/process`, {
        method: 'POST',
        body: formData,
        credentials: 'include',
      });

      if (!response.ok) {
        const contentType = response.headers.get('content-type') || '';
        
        if (contentType.includes('application/json')) {
          const errorData = await response.json();
          throw new Error(errorData.message || errorData.error || `Processing failed`);
        } else {
          throw new Error(`Server error. Please check if the server is running.`);
        }
      }

      const data = await response.json();

      if (!data.success) {
        throw new Error(data.message || 'Processing failed');
      }

      const cleanPreview = Array.isArray(data.preview) 
        ? data.preview.map(row => {
            const cleanRow = {};
            Object.entries(row).forEach(([key, value]) => {
              cleanRow[key] = (value === null || value === undefined || value === 'NaN') ? '' : value;
            });
            return cleanRow;
          })
        : [];

      setPreview(cleanPreview);
      setPreviewGenerated(true);
      
      let successMessage = data.message || 'Files processed successfully';
      if (data.total_rows) {
        successMessage += ` (${data.total_rows} rows processed)`;
      }
      
      setMessage(successMessage);

      if (data.warnings?.length > 0) {
        setMessage(prev => `${prev} - ${data.warnings.length} warning(s) encountered`);
      }

    } catch (err) {
      setMessage(`Error: ${err.message}`);
      setPreviewGenerated(false);
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async () => {
    if (Object.keys(files).length === 0) {
      setMessage('Please select at least one file to download.');
      return;
    }

    setLoading(true);
    setMessage(null);
    
    try {
      const formData = new FormData();
      
      Object.entries(files).forEach(([month, file]) => {
        formData.append(month, file);
      });
      
      formData.append('skipFirstRow', skipFirstRow.toString());

      const response = await fetch(`${API_BASE_URL}/api/download`, {
        method: 'POST',
        body: formData,
        credentials: 'include',
      });
      
      if (!response.ok) {
        const contentType = response.headers.get('content-type') || '';
        
        if (contentType.includes('application/json')) {
          const errorData = await response.json();
          throw new Error(errorData.message || errorData.error || 'Download failed');
        } else {
          throw new Error('Download failed. Please try again.');
        }
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Combined_Sales_Report.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      
      setMessage('File downloaded successfully!');
      
    } catch (err) {
      setMessage(`Download Error: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const isProcessingDisabled = loading || Object.keys(files).length === 0;
  const isDownloadDisabled = loading || Object.keys(files).length === 0;

  return (
    <div style={{ 
      minHeight: '100vh', 
      padding: '20px',
      backgroundColor: '#f8f9fa',
      fontFamily: 'Arial, sans-serif'
    }}>
      <div style={{ 
        maxWidth: '1200px', 
        margin: '0 auto',
        backgroundColor: 'white',
        padding: '30px',
        borderRadius: '12px',
        boxShadow: '0 4px 6px rgba(0,0,0,0.1)'
      }}>
        <h1 style={{ 
          textAlign: 'center',
          marginBottom: '40px',
          color: '#2c3e50',
          fontSize: '2.2rem',
          fontWeight: '600',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: '12px'
        }}>
          <span style={{ fontSize: '2rem' }}>📊</span> 
          Monthly Sales Data Processor
        </h1>
        
        {/* Settings Section */}
        <div style={{ 
          marginBottom: '30px',
          padding: '20px',
          backgroundColor: '#f8f9fa',
          borderRadius: '8px',
          border: '1px solid #e9ecef'
        }}>
          <label style={{ 
            display: 'flex', 
            alignItems: 'center',
            fontSize: '16px',
            fontWeight: '500',
            color: '#495057'
          }}>
            <input
              type="checkbox"
              checked={skipFirstRow}
              onChange={() => setSkipFirstRow(!skipFirstRow)}
              style={{ 
                marginRight: '12px',
                width: '20px',
                height: '20px',
                accentColor: '#007bff'
              }}
            />
            <span>Skip first row (use second row as headers)</span>
          </label>
        </div>
        
        {/* File Upload Section */}
        <div style={{ 
          marginBottom: '30px',
          padding: '25px',
          border: '2px solid #dee2e6',
          borderRadius: '10px',
          backgroundColor: '#fdfdfd'
        }}>
          <h2 style={{ 
            marginBottom: '20px',
            color: '#2c3e50',
            fontSize: '1.4rem',
            fontWeight: '600'
          }}>
            Upload Monthly Files
            {Object.keys(files).length > 0 && (
              <span style={{
                marginLeft: '15px',
                padding: '4px 12px',
                backgroundColor: '#28a745',
                color: 'white',
                borderRadius: '20px',
                fontSize: '0.9rem',
                fontWeight: '500'
              }}>
                {Object.keys(files).length} selected
              </span>
            )}
          </h2>
          
          <div style={{ 
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))',
            gap: '18px'
          }}>
            {months.map(month => (
              <div 
                key={month}
                style={{ 
                  padding: '20px',
                  border: `2px solid ${files[month] ? '#28a745' : '#ced4da'}`,
                  borderRadius: '10px',
                  backgroundColor: files[month] ? '#f8fff9' : 'white',
                  transition: 'all 0.2s ease-in-out',
                  position: 'relative'
                }}
              >
                <div style={{ 
                  fontWeight: '600',
                  marginBottom: '15px',
                  color: '#2c3e50',
                  fontSize: '1.1rem',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between'
                }}>
                  {month}
                  {files[month] && (
                    <span style={{
                      color: '#28a745',
                      fontSize: '1.2rem'
                    }}>✓</span>
                  )}
                </div>
                
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) => handleFileChange(month, e)}
                  style={{ display: 'none' }}
                  id={`file-${month}`}
                />
                
                <label
                  htmlFor={`file-${month}`}
                  style={{
                    display: 'block',
                    padding: '12px 16px',
                    backgroundColor: files[month] ? '#6c757d' : '#007bff',
                    color: 'white',
                    borderRadius: '6px',
                    cursor: 'pointer',
                    textAlign: 'center',
                    fontSize: '14px',
                    fontWeight: '500',
                    transition: 'background-color 0.2s ease',
                    border: 'none'
                  }}
                  onMouseEnter={(e) => {
                    e.target.style.backgroundColor = files[month] ? '#5a6268' : '#0056b3';
                  }}
                  onMouseLeave={(e) => {
                    e.target.style.backgroundColor = files[month] ? '#6c757d' : '#007bff';
                  }}
                >
                  {files[month] ? 'Change File' : 'Select File'}
                </label>
                
                {files[month] && (
                  <>
                    <div style={{ 
                      marginTop: '12px',
                      fontSize: '13px',
                      color: '#6c757d',
                      wordBreak: 'break-word',
                      backgroundColor: '#f8f9fa',
                      padding: '8px',
                      borderRadius: '4px'
                    }}>
                      <div style={{ fontWeight: '500' }}>{files[month].name}</div>
                      <div style={{ marginTop: '4px' }}>
                        {(files[month].size / 1024).toFixed(1)} KB
                      </div>
                    </div>
                    <button
                      onClick={() => handleDelete(month)}
                      style={{
                        marginTop: '10px',
                        padding: '6px 12px',
                        backgroundColor: '#dc3545',
                        color: 'white',
                        border: 'none',
                        borderRadius: '4px',
                        cursor: 'pointer',
                        fontSize: '12px',
                        fontWeight: '500',
                        width: '100%',
                        transition: 'background-color 0.2s ease'
                      }}
                      onMouseEnter={(e) => {
                        e.target.style.backgroundColor = '#c82333';
                      }}
                      onMouseLeave={(e) => {
                        e.target.style.backgroundColor = '#dc3545';
                      }}
                    >
                      Remove File
                    </button>
                  </>
                )}
              </div>
            ))}
          </div>
        </div>
        
        {/* Action Buttons */}
        <div style={{ 
          display: 'flex',
          justifyContent: 'center',
          gap: '20px',
          marginBottom: '30px',
          flexWrap: 'wrap'
        }}>
          <button
            onClick={handleProcess}
            disabled={isProcessingDisabled}
            style={{
              padding: '15px 30px',
              backgroundColor: isProcessingDisabled ? '#6c757d' : '#007bff',
              color: 'white',
              border: 'none',
              borderRadius: '8px',
              cursor: isProcessingDisabled ? 'not-allowed' : 'pointer',
              fontSize: '16px',
              fontWeight: '600',
              minWidth: '200px',
              transition: 'all 0.2s ease',
              boxShadow: isProcessingDisabled ? 'none' : '0 2px 4px rgba(0,123,255,0.3)'
            }}
            onMouseEnter={(e) => {
              if (!isProcessingDisabled) {
                e.target.style.backgroundColor = '#0056b3';
                e.target.style.transform = 'translateY(-1px)';
              }
            }}
            onMouseLeave={(e) => {
              if (!isProcessingDisabled) {
                e.target.style.backgroundColor = '#007bff';
                e.target.style.transform = 'translateY(0)';
              }
            }}
          >
            {loading ? 'Processing...' : 'Process Files'}
          </button>
          
          <button
            onClick={handleDownload}
            disabled={isDownloadDisabled}
            style={{
              padding: '15px 30px',
              backgroundColor: isDownloadDisabled ? '#6c757d' : '#28a745',
              color: 'white',
              border: 'none',
              borderRadius: '8px',
              cursor: isDownloadDisabled ? 'not-allowed' : 'pointer',
              fontSize: '16px',
              fontWeight: '600',
              minWidth: '200px',
              transition: 'all 0.2s ease',
              boxShadow: isDownloadDisabled ? 'none' : '0 2px 4px rgba(40,167,69,0.3)'
            }}
            onMouseEnter={(e) => {
              if (!isDownloadDisabled) {
                e.target.style.backgroundColor = '#218838';
                e.target.style.transform = 'translateY(-1px)';
              }
            }}
            onMouseLeave={(e) => {
              if (!isDownloadDisabled) {
                e.target.style.backgroundColor = '#28a745';
                e.target.style.transform = 'translateY(0)';
              }
            }}
          >
            {loading ? 'Preparing...' : 'Download Excel'}
          </button>
        </div>
        
        {/* Status Messages */}
        {message && (
          <div style={{ 
            padding: '18px 24px',
            marginBottom: '25px',
            backgroundColor: message.includes('Error') ? '#fff3cd' : '#d1edff',
            color: message.includes('Error') ? '#856404' : '#0c5460',
            borderRadius: '8px',
            border: `1px solid ${message.includes('Error') ? '#ffeaa7' : '#bee5eb'}`,
            fontSize: '15px',
            fontWeight: '500',
            display: 'flex',
            alignItems: 'center',
            gap: '10px'
          }}>
            <span style={{ fontSize: '1.2rem' }}>
              {message.includes('Error') ? '⚠️' : '✅'}
            </span>
            {message}
          </div>
        )}
        
        {/* Data Preview */}
        {preview && preview.length > 0 && (
          <div style={{ 
            padding: '25px',
            border: '1px solid #dee2e6',
            borderRadius: '10px',
            backgroundColor: '#fdfdfd'
          }}>
            <h2 style={{ 
              marginBottom: '20px',
              display: 'flex',
              alignItems: 'center',
              gap: '12px',
              color: '#2c3e50',
              fontSize: '1.4rem',
              fontWeight: '600'
            }}>
              <span style={{ fontSize: '1.5rem' }}>📋</span> 
              Data Preview
              <span style={{
                fontSize: '0.9rem',
                color: '#6c757d',
                fontWeight: '400'
              }}>
                (First 5 rows)
              </span>
            </h2>
            
            <div style={{ 
              overflowX: 'auto',
              border: '1px solid #dee2e6',
              borderRadius: '6px',
              boxShadow: '0 2px 4px rgba(0,0,0,0.05)'
            }}>
              <table style={{ 
                width: '100%',
                borderCollapse: 'collapse',
                fontSize: '14px',
                backgroundColor: 'white'
              }}>
                <thead>
                  <tr style={{ 
                    backgroundColor: '#f8f9fa',
                    borderBottom: '2px solid #dee2e6'
                  }}>
                    {Object.keys(preview[0]).map(header => (
                      <th 
                        key={header}
                        style={{ 
                          padding: '12px 15px',
                          textAlign: 'left',
                          border: '1px solid #dee2e6',
                          fontWeight: '600',
                          color: '#2c3e50',
                          fontSize: '13px',
                          textTransform: 'uppercase',
                          letterSpacing: '0.5px'
                        }}
                      >
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                
                <tbody>
                  {preview.map((row, rowIndex) => (
                    <tr 
                      key={rowIndex}
                      style={{ 
                        backgroundColor: rowIndex % 2 === 0 ? 'white' : '#f8f9fa',
                        transition: 'background-color 0.2s ease'
                      }}
                      onMouseEnter={(e) => {
                        e.target.style.backgroundColor = '#e3f2fd';
                      }}
                      onMouseLeave={(e) => {
                        e.target.style.backgroundColor = rowIndex % 2 === 0 ? 'white' : '#f8f9fa';
                      }}
                    >
                      {Object.values(row).map((cell, cellIndex) => (
                        <td 
                          key={cellIndex}
                          style={{ 
                            padding: '10px 15px',
                            border: '1px solid #dee2e6',
                            whiteSpace: 'nowrap',
                            color: '#495057'
                          }}
                        >
                          {cell !== null && cell !== undefined && cell !== '' ? cell.toString() : ''}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default App;
