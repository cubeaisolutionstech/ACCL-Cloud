import React, { useState, useRef } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';

const SidebarUploadPanel = ({ mode, onLogout }) => {
  const { selectedFiles, setSelectedFiles } = useExcelData();
  const [dragActive, setDragActive] = useState({});
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

    try {
      const response = await axios.post(`/api/${mode}/upload`, formData);
      
      const filename = response.data.filename;
      
      setSelectedFiles((prev) => ({
        ...prev,
        [`${mappedType}File`]: filename
      }));
    } catch (err) {
      console.error(`Error uploading ${type} file:`, err);
      alert(`Failed to upload ${type} file.`);
    }
  };

  const handleFileChange = (e, type) => {
    const file = e.target.files[0];
    if (file) {
      uploadFile(file, type);
    }
  };

  const handleDrag = (e, type) => {
    e.preventDefault();
    e.stopPropagation();
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
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      uploadFile(file, type);
    }
  };

  const handleClick = (type) => {
    if (fileInputRefs.current[type]) {
      fileInputRefs.current[type].click();
    }
  };

  const removeFile = (type) => {
    setSelectedFiles((prev) => {
      const newState = { ...prev };
      delete newState[`${type}File`];
      return newState;
    });
  };

  const fileInputs = [
    { label: 'Current Month Sales Excel File', type: 'sales' },
    { label: 'Current Year Budget/Executive Target File', type: 'budget' },
    { label: 'OS Previous File', type: 'osPrev' },        // Changed from 'osPrev' to 'osJan'
    { label: 'OS Current File', type: 'osCurr' },         // Changed from 'osCurr' to 'osFeb'
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
      <h2 className="text-base font-bold mb-2 capitalize">{mode} Uploads</h2>
      
      <div className="flex-1 space-y-2 overflow-y-auto scrollbar-hide">
        {fileInputs.map(({ label, type }) => (
          <div key={type} className="space-y-1">
            <label className="block text-xs font-semibold text-blue-100">{label}</label>
            
            {/* Hidden file input */}
            <input 
              ref={(el) => fileInputRefs.current[type] = el}
              type="file" 
              accept=".xlsx,.xls,.jpg,.png" 
              onChange={(e) => handleFileChange(e, type)} 
              className="hidden" 
            />
            
            {/* Drag and drop area */}
            <div
              className={`relative border-2 border-dashed rounded p-1.5 text-center transition-all duration-300 cursor-pointer ${
                dragActive[type] 
                  ? 'border-blue-400 bg-blue-800/50 scale-105' 
                  : 'border-blue-300 hover:border-blue-400 hover:bg-blue-800/30'
              }`}
              onDragEnter={(e) => handleDrag(e, type)}
              onDragLeave={(e) => handleDrag(e, type)}
              onDragOver={(e) => handleDrag(e, type)}
              onDrop={(e) => handleDrop(e, type)}
              onClick={() => handleClick(type)}
            >
              <div className="space-y-0.5">
                <div className="text-sm mb-0.5">📁</div>
                <div className="text-xs font-medium">Drag and drop file here</div>
                <div className="text-xs text-blue-200">Limit 200MB • XLSX, XLS</div>
                <button 
                  type="button"
                  className="mt-0.5 px-2 py-0.5 bg-blue-600 hover:bg-blue-700 rounded font-medium transition-colors duration-200 text-xs"
                  onClick={(e) => {
                    e.stopPropagation();
                    handleClick(type);
                  }}
                >
                  Browse files
                </button>
              </div>
            </div>

            {/* Uploaded file display */}
            {selectedFiles[`${type}File`] && (
              <div className="bg-green-900/50 border border-green-500 rounded p-1.5 flex items-center justify-between">
                <div className="flex items-center space-x-1.5">
                  <span className="text-xs">{getFileIcon(type)}</span>
                  <div className="flex-1 min-w-0">
                    <p className="text-xs font-medium text-green-200 truncate">
                      {selectedFiles[`${type}File`]}
                    </p>
                    <p className="text-xs text-green-300">Uploaded successfully</p>
                  </div>
                </div>
                <button
                  onClick={() => removeFile(type)}
                  className="text-red-400 hover:text-red-300 transition-colors duration-200 text-xs"
                  title="Remove file"
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
        >
          Logout
        </button>
      </div>
    </div>
  );
};

export default SidebarUploadPanel;