import React, { useState, useEffect } from 'react';
import api from "../api/axios";

const months = [
  "April", "May", "June", "July",
  "August", "September", "October", "November",
  "December", "January", "February", "March"
];

const Cumulative = () => {
  const [files, setFiles] = useState({});
  const [skipFirstRow, setSkipFirstRow] = useState(false);
  const [preview, setPreview] = useState(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState(null);
  const [previewGenerated, setPreviewGenerated] = useState(false);
  const [availableFiles, setAvailableFiles] = useState([]);
  const [selectedMonths, setSelectedMonths] = useState([]);
  const [dbLoading, setDbLoading] = useState(false);
  const [dbError, setDbError] = useState(null);

  const fetchAvailableFiles = async () => {
    try {
      setDbLoading(true);
      setDbError(null);
      
      const response = await api.get('/cumulative/api/files');
      
      if (response.data.success) {
        setAvailableFiles(response.data.files);
      } else {
        throw new Error(response.data.message || 'Invalid response from server');
      }
    } catch (err) {
      console.error('Error loading files:', err);
      setDbError(err.message);
    } finally {
      setDbLoading(false);
    }
  };

  useEffect(() => {
    fetchAvailableFiles();
  }, []);

  const handleFileChange = (month, e) => {
    if (e.target.files[0]) {
      const newFiles = { ...files, [month]: e.target.files[0] };
      setFiles(newFiles);
      setPreviewGenerated(false);
      setSelectedMonths(prev => prev.filter(m => m !== month));
    }
  };

  const handleProcess = async () => {
    setLoading(true);
    setMessage(null);
    setPreview(null);
    
    try {
      const formData = new FormData();
      
      Object.entries(files).forEach(([month, file]) => {
        formData.append(month, file);
      });
      
      selectedMonths.forEach(month => {
        formData.append('months[]', month);
      });
      
      formData.append('skipFirstRow', skipFirstRow);

      const response = await api.post('/cumulative/api/process', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });
      
      if (response.data.success) {
        setPreview(response.data.preview);
        setMessage(response.data.message);
        setPreviewGenerated(true);
        fetchAvailableFiles();
        
        if (response.data.warnings?.length > 0) {
          setMessage(prev => `${prev} (with ${response.data.warnings.length} warnings)`);
        }
      } else {
        throw new Error(response.data.message || 'Failed to process files');
      }
    } catch (err) {
      console.error('Processing error:', err);
      setMessage(err.message);
      setPreviewGenerated(false);
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async () => {
    setLoading(true);
    setMessage(null);
    
    try {
      const formData = new FormData();
      
      Object.entries(files).forEach(([month, file]) => {
        formData.append(month, file);
      });
      
      selectedMonths.forEach(month => {
        formData.append('months[]', month);
      });
      
      formData.append('skipFirstRow', skipFirstRow);

      const response = await api.post('/cumulative/api/download', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob',
      });
      
      const blob = new Blob([response.data]);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Combined_Sales_Report.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Download error:', err);
      setMessage(err.message);
    } finally {
      setLoading(false);
    }
  };

  const toggleMonthSelection = (month) => {
    setSelectedMonths(prev => 
      prev.includes(month) 
        ? prev.filter(m => m !== month) 
        : [...prev, month]
    );
    
    if (files[month]) {
      setFiles(prev => {
        const newFiles = { ...prev };
        delete newFiles[month];
        return newFiles;
      });
    }
  };

  const isProcessingDisabled = loading || 
    (Object.keys(files).length === 0 && selectedMonths.length === 0);

  const isDownloadDisabled = loading || 
    (!previewGenerated && selectedMonths.length === 0);

  return (
    <div className="p-6 bg-gray-50 min-h-screen">
      <div className="max-w-7xl mx-auto bg-white rounded-lg shadow-sm p-6">
        <h1 className="text-3xl font-bold text-center mb-8 text-gray-800 flex items-center justify-center gap-3">
          <span className="text-2xl">📊</span> 
          Monthly Sales Data Processor
        </h1>
        
        {/* Settings Section */}
        <div className="mb-6 p-4 bg-blue-50 rounded-lg">
          <label className="flex items-center cursor-pointer">
            <input
              type="checkbox"
              checked={skipFirstRow}
              onChange={() => setSkipFirstRow(!skipFirstRow)}
              className="mr-3 w-5 h-5 text-blue-600 rounded focus:ring-blue-500"
            />
            <span className="text-gray-700">Skip first row (use second row as headers)</span>
          </label>
        </div>
        
        {/* File Upload Section */}
        <div className="mb-6 p-4 border border-gray-200 rounded-lg">
          <h2 className="text-xl font-semibold mb-4 text-gray-800">Upload New Files</h2>
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            {months.map(month => (
              <div 
                key={month}
                className={`p-4 border-2 rounded-lg transition-colors ${
                  files[month] 
                    ? 'border-green-500 bg-green-50' 
                    : 'border-gray-300 bg-white'
                } ${selectedMonths.includes(month) ? 'opacity-50' : ''}`}
              >
                <div className="font-semibold text-gray-700 mb-3">
                  {month}
                </div>
                
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) => handleFileChange(month, e)}
                  className="hidden"
                  id={`file-${month}`}
                  disabled={selectedMonths.includes(month)}
                />
                
                <label
                  htmlFor={`file-${month}`}
                  className={`block w-full text-center py-2 px-3 rounded-md cursor-pointer transition-colors ${
                    selectedMonths.includes(month)
                      ? 'bg-gray-300 text-gray-500 cursor-not-allowed'
                      : 'bg-blue-600 text-white hover:bg-blue-700'
                  }`}
                >
                  {files[month] ? 'Change File' : 'Select File'}
                </label>
                
                {files[month] && (
                  <div className="mt-2 text-xs text-gray-600 break-words">
                    {files[month].name}
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
        
        {/* Database Files Section */}
        <div className="mb-6 p-4 border border-gray-200 rounded-lg">
          <h2 className="text-xl font-semibold mb-4 text-gray-800">Available Files in Database</h2>
          
          {dbLoading ? (
            <div className="text-center py-8 text-gray-500">
              Loading available files...
            </div>
          ) : dbError ? (
            <div className="p-4 bg-yellow-50 border border-yellow-200 rounded-lg mb-4">
              <div className="text-yellow-800 mb-2">
                Error loading files: {dbError}
              </div>
              <button 
                onClick={fetchAvailableFiles}
                className="px-3 py-1 bg-yellow-500 text-white rounded hover:bg-yellow-600 transition-colors"
              >
                Retry
              </button>
            </div>
          ) : availableFiles.length === 0 ? (
            <div className="text-center py-8 text-gray-500">
              No files found in database. Upload files above to get started.
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {availableFiles.map(file => (
                <div 
                  key={file.id}
                  onClick={() => toggleMonthSelection(file.month)}
                  className={`p-4 border-2 rounded-lg cursor-pointer transition-all hover:shadow-md ${
                    selectedMonths.includes(file.month)
                      ? 'border-blue-500 bg-blue-50'
                      : 'border-gray-200 bg-white hover:border-gray-300'
                  }`}
                >
                  <div className="flex justify-between items-center mb-2">
                    <span className="font-semibold text-gray-800">
                      {file.month}
                    </span>
                    {selectedMonths.includes(file.month) && (
                      <span className="text-blue-600 text-sm font-medium">Selected</span>
                    )}
                  </div>
                  
                  <div className="text-sm text-gray-600 mb-2 break-words">
                    {file.filename}
                  </div>
                  
                  <div className="text-xs text-gray-500">
                    Uploaded: {new Date(file.upload_date).toLocaleString()}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
        
        {/* Action Buttons */}
        <div className="flex flex-col sm:flex-row justify-center gap-4 mb-6">
          <button
            onClick={handleProcess}
            disabled={isProcessingDisabled}
            className={`px-6 py-3 rounded-lg font-semibold text-white transition-colors ${
              isProcessingDisabled
                ? 'bg-gray-400 cursor-not-allowed'
                : 'bg-blue-600 hover:bg-blue-700'
            }`}
          >
            {loading ? 'Processing...' : 'Process Files'}
          </button>
          
          <button
            onClick={handleDownload}
            disabled={isDownloadDisabled}
            className={`px-6 py-3 rounded-lg font-semibold text-white transition-colors ${
              isDownloadDisabled
                ? 'bg-gray-400 cursor-not-allowed'
                : 'bg-green-600 hover:bg-green-700'
            }`}
          >
            {loading ? 'Preparing...' : 'Download Excel'}
          </button>
        </div>
        
        {/* Status Messages */}
        {message && (
          <div className={`p-4 rounded-lg mb-6 ${
            message.includes('Error') 
              ? 'bg-yellow-50 border border-yellow-200 text-yellow-800'
              : 'bg-green-50 border border-green-200 text-green-800'
          }`}>
            {message}
          </div>
        )}
        
        {/* Data Preview */}
        {preview && preview.length > 0 && (
          <div className="p-4 border border-gray-200 rounded-lg">
            <h2 className="text-xl font-semibold mb-4 text-gray-800 flex items-center gap-2">
              <span>📋</span> Data Preview (First 5 Rows)
            </h2>
            
            <div className="overflow-x-auto border border-gray-200 rounded-lg">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gray-50">
                    {Object.keys(preview[0]).map(header => (
                      <th 
                        key={header}
                        className="px-4 py-3 text-left font-semibold text-gray-700 border-b border-gray-200"
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
                      className={`${
                        rowIndex % 2 === 0 ? 'bg-white' : 'bg-gray-50'
                      } hover:bg-gray-100`}
                    >
                      {Object.values(row).map((cell, cellIndex) => (
                        <td 
                          key={cellIndex}
                          className="px-4 py-3 border-b border-gray-200 whitespace-nowrap"
                        >
                          {cell !== null && cell !== undefined ? cell.toString() : ''}
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

export default Cumulative; 
