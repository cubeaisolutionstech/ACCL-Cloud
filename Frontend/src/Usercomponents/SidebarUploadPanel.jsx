import React, { useState, useRef } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';

const SidebarUploadPanel = ({ mode, onLogout }) => {
  const { selectedFiles, setSelectedFiles } = useExcelData();
  const [dragActive, setDragActive] = useState({});
  const [uploadProgress, setUploadProgress] = useState({});
  const [uploadingFiles, setUploadingFiles] = useState({});
  const fileInputRefs = useRef({});

  const mapTypeForInternal = (type) => {
    const typeMap = {
      osJan: 'osPrev',
      osFeb: 'osCurr',
    };
    return typeMap[type] || type;
  };

  const uploadFile = async (file, type) => {
    const mappedType = mapTypeForInternal(type);
    const formData = new FormData();
    formData.append('file', file);

    // Set uploading state
    setUploadingFiles(prev => ({ ...prev, [type]: true }));
    setUploadProgress(prev => ({ ...prev, [type]: 0 }));

    try {
      const response = await axios.post(
        `http://localhost:5000/api/${mode}/upload`, 
        formData,
        {
          headers: {
            'Content-Type': 'multipart/form-data',
          },
          onUploadProgress: (progressEvent) => {
            const percentCompleted = Math.round(
              (progressEvent.loaded * 100) / progressEvent.total
            );
            setUploadProgress(prev => ({ ...prev, [type]: percentCompleted }));
          },
        }
      );
      
      const filename = response.data.filename;
      
      setSelectedFiles((prev) => ({
        ...prev,
        [`${mappedType}File`]: filename
      }));

      // Clear progress after success
      setTimeout(() => {
        setUploadProgress(prev => {
          const newProgress = { ...prev };
          delete newProgress[type];
          return newProgress;
        });
      }, 1000);

    } catch (err) {
      console.error(`Error uploading ${type} file:`, err);
      alert(`Failed to upload ${type} file. Please try again.`);
    } finally {
      // Clear uploading state
      setUploadingFiles(prev => {
        const newState = { ...prev };
        delete newState[type];
        return newState;
      });
    }
  };

  const handleFileChange = (e, type) => {
    const file = e.target.files[0];
    if (file && !uploadingFiles[type]) {
      uploadFile(file, type);
    }
  };

  const handleDrag = (e, type) => {
    e.preventDefault();
    e.stopPropagation();
    
    // Don't allow drag events during upload
    if (uploadingFiles[type]) return;
    
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(prev => ({ ...prev, [type]: true }));
    } else if (e.type === "dragleave") {
      setDragActive(prev => ({ ...prev, [type]: false }));
    }
  };

  const handleDrop = (e, type) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(prev => ({ ...prev, [type]: false }));
    
    // Don't allow drop during upload
    if (uploadingFiles[type]) return;
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      uploadFile(file, type);
    }
  };

  const handleClick = (type) => {
    // Don't allow click during upload
    if (uploadingFiles[type]) return;
    
    if (fileInputRefs.current[type]) {
      fileInputRefs.current[type].click();
    }
  };

  const removeFile = (type) => {
    // Don't allow removal during upload
    if (uploadingFiles[type]) return;
    
    setSelectedFiles((prev) => {
      const newState = { ...prev };
      delete newState[`${type}File`];
      return newState;
    });
  };

  const fileInputs = [
    { label: 'Sales File', type: 'sales' },
    { label: 'Budget / Executive Target File', type: 'budget' },
    { label: 'OS Previous File', type: 'osPrev' },
    { label: 'OS Current File', type: 'osCurr' },
    { label: 'Last Year Sales File', type: 'lastYearSales' },
  ];

  const getFileIcon = (type) => {
    return '📄';
  };

  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  };

  return (
    <div className="w-72 bg-blue-900 text-white p-3 space-y-2 shadow-lg h-screen flex flex-col">
      <h2 className="text-lg font-bold mb-3 capitalize text-white bg-blue-800/50 px-3 py-2 rounded-lg border border-blue-600">{mode} Uploads</h2>
      
      <div className="flex-1 space-y-2 overflow-y-auto scrollbar-hide">
        {fileInputs.map(({ label, type }) => (
          <div key={type} className="space-y-1">
            <label className="block text-sm font-bold text-white mb-1 bg-blue-800/30 px-2 py-1 rounded">{label}</label>
            
            {/* Hidden file input */}
            <input 
              ref={(el) => fileInputRefs.current[type] = el}
              type="file" 
              accept=".xlsx,.xls,.jpg,.png" 
              onChange={(e) => handleFileChange(e, type)} 
              className="hidden"
              disabled={uploadingFiles[type]}
            />
            
            {/* Drag and drop area */}
            <div
              className={`relative border-2 border-dashed rounded p-1.5 text-center transition-all duration-300 ${
                uploadingFiles[type] 
                  ? 'border-yellow-400 bg-yellow-800/30 cursor-not-allowed opacity-75' 
                  : dragActive[type] 
                    ? 'border-blue-400 bg-blue-800/50 scale-105 cursor-pointer' 
                    : 'border-blue-300 hover:border-blue-400 hover:bg-blue-800/30 cursor-pointer'
              }`}
              onDragEnter={(e) => handleDrag(e, type)}
              onDragLeave={(e) => handleDrag(e, type)}
              onDragOver={(e) => handleDrag(e, type)}
              onDrop={(e) => handleDrop(e, type)}
              onClick={() => handleClick(type)}
            >
              {/* Upload Progress Overlay */}
              {uploadingFiles[type] && (
                <div className="absolute inset-0 bg-blue-800/80 rounded flex flex-col items-center justify-center">
                  <div className="w-8 h-8 mb-2">
                    <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-white"></div>
                  </div>
                  <div className="text-xs font-medium text-white">
                    Uploading... {uploadProgress[type] || 0}%
                  </div>
                  <div className="w-20 bg-gray-200 rounded-full h-1.5 mt-1">
                    <div 
                      className="bg-blue-600 h-1.5 rounded-full transition-all duration-300" 
                      style={{ width: `${uploadProgress[type] || 0}%` }}
                    ></div>
                  </div>
                </div>
              )}

              <div className="space-y-0.5">
                <div className="text-sm mb-0.5">📁</div>
                <div className="text-xs font-medium">
                  {uploadingFiles[type] ? 'Uploading...' : 'Drag and drop file here'}
                </div>
                <div className="text-xs text-blue-200">Limit 200MB • XLSX, XLS</div>
                <button 
                  type="button"
                  className={`mt-0.5 px-2 py-0.5 rounded font-medium transition-colors duration-200 text-xs ${
                    uploadingFiles[type] 
                      ? 'bg-gray-600 cursor-not-allowed' 
                      : 'bg-blue-600 hover:bg-blue-700'
                  }`}
                  onClick={(e) => {
                    e.stopPropagation();
                    handleClick(type);
                  }}
                  disabled={uploadingFiles[type]}
                >
                  {uploadingFiles[type] ? 'Uploading...' : 'Browse files'}
                </button>
              </div>
            </div>

            {/* Uploaded file display */}
            {selectedFiles[`${type}File`] && !uploadingFiles[type] && (
              <div className="bg-green-50/20 border border-green-400/50 rounded p-1.5 flex items-center justify-between">
                <div className="flex items-center space-x-1.5">
                  <span className="text-xs text-green-300">✓</span>
                  <div className="flex-1 min-w-0">
                    <p className="text-xs font-medium text-white truncate">
                      {selectedFiles[`${type}File`]}
                    </p>
                    <p className="text-xs text-white">Uploaded successfully</p>
                  </div>
                </div>
                <button
                  onClick={() => removeFile(type)}
                  className="text-red-400 hover:text-red-300 transition-colors duration-200 text-xs"
                  title="Remove file"
                  disabled={uploadingFiles[type]}
                >
                  ✕
                </button>
              </div>
            )}
          </div>
        ))}
      </div>

      {/* Logout Button */}
      <div className="pt-2 border-t border-blue-700 flex-shrink-0">
        <button
          onClick={onLogout}
          className="w-full bg-red-600 text-white py-1.5 rounded hover:bg-red-700 text-sm font-semibold transition-all duration-200"
          disabled={Object.keys(uploadingFiles).length > 0}
        >
          {Object.keys(uploadingFiles).length > 0 ? 'Uploading...' : 'Logout'}
        </button>
      </div>
    </div>
  );
};

export default SidebarUploadPanel;
