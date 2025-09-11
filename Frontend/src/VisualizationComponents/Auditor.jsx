import React, { useState, useCallback, useEffect } from 'react';
import { 
  FileSpreadsheet, 
  Download, 
  Merge, 
  AlertCircle, 
  CheckCircle, 
  Info, 
  X, 
  Search,
  RefreshCw,
  BarChart3,
  FileText,
  Trash2,
  Settings,
  Database,
  MapPin,
  Package,
  Building
} from 'lucide-react';
import './App1.css';
import SalesFormat from './SalesFormat';
import AuditorFormat from './AuditorFormat';
import RegionAnalysis from './RegionAnalysis';
import ProductAnalysis from './ProductAnalysis';
import TsPwAnalysis from './TsPwAnalysis';
import EroPwAnalysis from './EroPwAnalysis';
import SalesAnalysisMonthwise from './SalesAnalysisMonthwise';
import CombinedExcelManager from './CombinedExcelManager';

const API_BASE_URL = process.env.REACT_APP_API_URL || 'http://localhost:5000/api';

function Auditor() {
// File state
  const [files, setFiles] = useState({
    budget: null,
    sales: null,
    totalSales: null,
    auditor: null
  });
  
  // Core application state
  const [uploadedFiles, setUploadedFiles] = useState({});
  const [processedData, setProcessedData] = useState({});
  const [selectedSheets, setSelectedSheets] = useState({});
  const [loading, setLoading] = useState(false);
  const [messages, setMessages] = useState([]);
  const [activeTab, setActiveTab] = useState('auditor');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedTable, setSelectedTable] = useState('Table 2: SALES in Value');
  const [availableTables, setAvailableTables] = useState([]);
  
  // Merge configuration
  const [mergeConfig, setMergeConfig] = useState({
    datasets: [],
    mergeColumn: '',
    mergeType: 'left'
  });

  // Product analysis data
  const [productMtData, setProductMtData] = useState(null);
  const [productValueData, setProductValueData] = useState(null);
  const [sessionTotals, setSessionTotals] = useState(null);
  const [sessionTotalsHistory, setSessionTotalsHistory] = useState([]);

  // Region analysis data
  const [regionData, setRegionData] = useState({
    mt: null,
    value: null
  });
  
  // Fiscal information
  const [fiscalInfo, setFiscalInfo] = useState({
    fiscal_year_str: '2025-26'
  });

  // Stored Files Management
  const [storedFiles, setStoredFiles] = useState([]);

  // Load saved data on component mount
  useEffect(() => {
    const savedMessages = localStorage.getItem('acl_messages');
    if (savedMessages) {
      try {
        setMessages(JSON.parse(savedMessages));
      } catch (e) {
        console.warn('Failed to parse saved messages');
      }
    }

    const savedProductMtData = localStorage.getItem('acl_product_mt_data');
    const savedProductValueData = localStorage.getItem('acl_product_value_data');
    
    if (savedProductMtData) {
      try {
        setProductMtData(JSON.parse(savedProductMtData));
      } catch (e) {
        console.warn('Failed to parse saved product MT data');
      }
    }
    
    if (savedProductValueData) {
      try {
        setProductValueData(JSON.parse(savedProductValueData));
      } catch (e) {
        console.warn('Failed to parse saved product Value data');
      }
    }

    // Load stored files from localStorage
    const savedStoredFiles = localStorage.getItem('acl_stored_files');
    if (savedStoredFiles) {
      try {
        const parsedFiles = JSON.parse(savedStoredFiles);
        // Recreate blob URLs for stored files (note: this won't work across browser sessions)
        const filesWithoutUrls = parsedFiles.map(file => ({
          ...file,
          url: null // URLs can't be persisted across sessions
        }));
        setStoredFiles(filesWithoutUrls);
      } catch (e) {
        console.warn('Failed to parse saved stored files');
      }
    }

    // Load session totals from localStorage
    const savedSessionTotals = localStorage.getItem('acl_session_totals');
    if (savedSessionTotals) {
      try {
        setSessionTotals(JSON.parse(savedSessionTotals));
      } catch (e) {
        console.warn('Failed to parse saved session totals');
      }
    }
  }, []);

  // Save messages to localStorage
  useEffect(() => {
    localStorage.setItem('acl_messages', JSON.stringify(messages));
  }, [messages]);

  // Save product MT data to localStorage
  useEffect(() => {
    if (productMtData) {
      localStorage.setItem('acl_product_mt_data', JSON.stringify(productMtData));
    } else {
      localStorage.removeItem('acl_product_mt_data');
    }
  }, [productMtData]);

  // Save product Value data to localStorage
  useEffect(() => {
    if (productValueData) {
      localStorage.setItem('acl_product_value_data', JSON.stringify(productValueData));
    } else {
      localStorage.removeItem('acl_product_value_data');
    }
  }, [productValueData]);

  // Save session totals to localStorage
  useEffect(() => {
    if (sessionTotals) {
      localStorage.setItem('acl_session_totals', JSON.stringify(sessionTotals));
    } else {
      localStorage.removeItem('acl_session_totals');
    }
  }, [sessionTotals]);

  // Save stored files to localStorage (without blob URLs)
  useEffect(() => {
    const filesToSave = storedFiles.map(file => ({
      ...file,
      url: null, // Don't save blob URLs
      blob: null // Don't save blob data
    }));
    localStorage.setItem('acl_stored_files', JSON.stringify(filesToSave));
  }, [storedFiles]);

  // Add message function
  const addMessage = useCallback((message, type = 'info') => {
    const newMessage = {
      id: Date.now(),
      text: message,
      type,
      timestamp: new Date().toLocaleTimeString()
    };
    setMessages(prev => [...prev, newMessage].slice(-50));
  }, []);

  // Clear messages function
  const clearMessages = () => {
    setMessages([]);
    localStorage.removeItem('acl_messages');
  };

  // ENHANCED: Stored Files Management Functions with automatic storage for Sales Analysis
  const addStoredFile = useCallback((fileData) => {
    // Enhanced logging for sales analysis files
    if (fileData.type === 'sales-analysis-excel') {
      console.log('📊 Sales Analysis Excel file being stored:', {
        name: fileData.name,
        size: fileData.size,
        hasMetadata: !!fileData.metadata,
        metadata: fileData.metadata
      });
      
      addMessage(`📊 Sales Analysis Excel file automatically stored: ${fileData.name} (${formatFileSize(fileData.size)})`, 'success');
    } else {
      addMessage(`📁 File stored: ${fileData.name}`, 'info');
    }
    
    setStoredFiles(prev => [fileData, ...prev]);
  }, []);

  const removeStoredFile = useCallback((fileId) => {
    setStoredFiles(prev => {
      const updatedFiles = prev.filter(file => file.id !== fileId);
      const fileToDelete = prev.find(file => file.id === fileId);
      if (fileToDelete && fileToDelete.url) {
        URL.revokeObjectURL(fileToDelete.url);
      }
      
      // Enhanced logging for sales analysis files
      if (fileToDelete && fileToDelete.type === 'sales-analysis-excel') {
        addMessage(`📊 Sales Analysis Excel file removed: ${fileToDelete.name}`, 'info');
      } else {
        addMessage('🗑️ File removed from storage', 'info');
      }
      
      return updatedFiles;
    });
  }, []);

  const clearAllStoredFiles = useCallback(() => {
    const salesAnalysisFiles = storedFiles.filter(file => file.type === 'sales-analysis-excel');
    
    storedFiles.forEach(file => {
      if (file.url) {
        URL.revokeObjectURL(file.url);
      }
    });
    
    setStoredFiles([]);
    localStorage.removeItem('acl_stored_files');
    
    if (salesAnalysisFiles.length > 0) {
      addMessage(`🗑️ All stored files cleared (including ${salesAnalysisFiles.length} Sales Analysis Excel files)`, 'info');
    } else {
      addMessage('🗑️ All stored files cleared', 'info');
    }
  }, [storedFiles]);

  // Helper function for file size formatting
  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  // ENHANCED: Auto-storage handler specifically for Sales Analysis Excel files
  const handleSalesAnalysisFileGenerated = useCallback((fileData) => {
    console.log('🎯 Sales Analysis Excel file generated, auto-storing:', fileData.name);
    
    // Ensure the file has the correct type for identification
    const enhancedFileData = {
      ...fileData,
      type: 'sales-analysis-excel',
      category: 'sales-analysis',
      autoStored: true,
      generatedAt: new Date().toISOString()
    };
    
    // Automatically add to stored files
    addStoredFile(enhancedFileData);
    
    // Log the successful auto-storage
    console.log('✅ Sales Analysis Excel file automatically stored in React state');
    
    return enhancedFileData;
  }, [addStoredFile]);

  // Extract session totals from data - IMPROVED VERSION
  const extractSessionTotals = useCallback((mtData, valueData) => {
    try {
      console.log('🔍 Extracting session totals...');
      console.log('MT Data available:', !!mtData);
      console.log('Value Data available:', !!valueData);
      
      const sessionTotals = {};
      
      if (mtData && mtData.data && mtData.data.length > 0) {
        console.log('Processing MT data, rows:', mtData.data.length);
        
        // Look for totals row (case insensitive, multiple patterns)
        const totalsRow = mtData.data.find(row => {
          const firstColumnValue = String(row[mtData.columns[0]] || '').toUpperCase();
          return firstColumnValue.includes('TOTAL') || 
                 firstColumnValue.includes('GRAND') ||
                 firstColumnValue.includes('SUM') ||
                 firstColumnValue.includes('AGGREGATE') ||
                 firstColumnValue.includes('ALL');
        });
        
        if (totalsRow) {
          sessionTotals.tonnage = totalsRow;
          console.log('✅ Found MT totals row:', Object.keys(totalsRow).length, 'columns');
        } else {
          // Use last row as fallback
          const lastRow = mtData.data[mtData.data.length - 1];
          sessionTotals.tonnage = lastRow;
          console.log('⚠️ Using last MT row as totals (no explicit totals found)');
        }
      }
      
      if (valueData && valueData.data && valueData.data.length > 0) {
        console.log('Processing Value data, rows:', valueData.data.length);
        
        // Look for totals row (case insensitive, multiple patterns)
        const totalsRow = valueData.data.find(row => {
          const firstColumnValue = String(row[valueData.columns[0]] || '').toUpperCase();
          return firstColumnValue.includes('TOTAL') || 
                 firstColumnValue.includes('GRAND') ||
                 firstColumnValue.includes('SUM') ||
                 firstColumnValue.includes('AGGREGATE') ||
                 firstColumnValue.includes('ALL');
        });
        
        if (totalsRow) {
          sessionTotals.value = totalsRow;
          console.log('✅ Found Value totals row:', Object.keys(totalsRow).length, 'columns');
        } else {
          // Use last row as fallback
          const lastRow = valueData.data[valueData.data.length - 1];
          sessionTotals.value = lastRow;
          console.log('⚠️ Using last Value row as totals (no explicit totals found)');
        }
      }
      
      console.log('🎯 Final session totals extracted:', Object.keys(sessionTotals));
      return sessionTotals;
    } catch (error) {
      console.error('❌ Error extracting session totals:', error);
      return {};
    }
  }, []);

  // Auto-extract session totals when product data changes
  useEffect(() => {
    if ((productMtData || productValueData) && !sessionTotals) {
      console.log('🔄 Product data available, auto-extracting session totals...');
      const extracted = extractSessionTotals(productMtData, productValueData);
      if (Object.keys(extracted).length > 0) {
        setSessionTotals(extracted);
        addMessage('✅ Session totals automatically extracted from product data', 'success');
      } else {
        console.log('⚠️ No session totals could be extracted');
      }
    }
  }, [productMtData, productValueData, sessionTotals, extractSessionTotals, addMessage]);

  // Handle product analysis completion
  const handleProductAnalysisComplete = useCallback((analysisResult) => {
    console.log('📦 Product analysis completed:', analysisResult);
    
    const mtResult = analysisResult?.mtData;
    const valueResult = analysisResult?.valueData;
    const extractedTotals = analysisResult?.sessionTotals;

    if (mtResult || valueResult) {
      setProductMtData(mtResult);
      setProductValueData(valueResult);
      addMessage('✅ Product analysis completed - Sales integration now available!', 'success');
    }

    // Handle provided session totals or extract them
    if (extractedTotals && Object.keys(extractedTotals).length > 0) {
      setSessionTotals(extractedTotals);
      setSessionTotalsHistory(prev => [
        {
          id: Date.now(),
          timestamp: new Date().toISOString(),
          totals: extractedTotals,
          rowCounts: analysisResult?.rowCounts || {
            mt: mtResult?.data?.length || 0,
            value: valueResult?.data?.length || 0
          }
        },
        ...prev.slice(0, 4)
      ]);
    } else if (mtResult || valueResult) {
      // Extract session totals if not provided
      const extracted = extractSessionTotals(mtResult, valueResult);
      if (Object.keys(extracted).length > 0) {
        setSessionTotals(extracted);
      }
    }
  }, [addMessage, extractSessionTotals]);

  // Handle region analysis completion
  const handleRegionAnalysisComplete = useCallback((analysisResult) => {
    if (analysisResult?.mt_data) {
      setRegionData(prev => ({
        ...prev,
        mt: {
          data: analysisResult.mt_data,
          columns: analysisResult.columns?.mt_columns || []
        }
      }));
    }

    if (analysisResult?.value_data) {
      setRegionData(prev => ({
        ...prev,
        value: {
          data: analysisResult.value_data,
          columns: analysisResult.columns?.value_columns || []
        }
      }));
    }

    if (analysisResult?.fiscal_year) {
      setFiscalInfo({
        fiscal_year_str: analysisResult.fiscal_year
      });
    }

    addMessage('✅ Region analysis completed - Files can now be stored in Combined Excel Manager!', 'success');
  }, [addMessage]);

  // Clear session totals
  const clearSessionTotals = useCallback(() => {
    setSessionTotals(null);
    setSessionTotalsHistory([]);
    localStorage.removeItem('acl_session_totals');
    addMessage('Session totals cleared', 'info');
  }, [addMessage]);

  // Clear product data
  const clearProductData = useCallback(() => {
    setProductMtData(null);
    setProductValueData(null);
    localStorage.removeItem('acl_product_mt_data');
    localStorage.removeItem('acl_product_value_data');
    clearSessionTotals();
    addMessage('Product data and session totals cleared - re-run product analysis for integration', 'info');
  }, [addMessage, clearSessionTotals]);

  // Load available tables when auditor file and sheet are selected
  useEffect(() => {
    if (uploadedFiles.auditor && selectedSheets.auditor) {
      loadAvailableTables();
    } else {
      setAvailableTables([]);
      setSelectedTable('Table 2: SALES in Value');
    }
  }, [uploadedFiles.auditor, selectedSheets.auditor]);

  // Load available tables function
  const loadAvailableTables = async () => {
    setLoading(true);
    try {
      const response = await fetch(`${API_BASE_URL}/get-available-tables`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: uploadedFiles.auditor.filepath,
          sheet_name: selectedSheets.auditor
        })
      });

      const result = await response.json();

      if (result.success) {
        setAvailableTables(result.available_tables);
        
        if (result.default_table && result.available_tables.includes(result.default_table)) {
          setSelectedTable(result.default_table);
        } else if (result.available_tables.length > 0) {
          setSelectedTable(result.available_tables[0]);
        }
      } else {
        addMessage(result.error || 'Failed to load available tables', 'error');
      }
    } catch (error) {
      addMessage(`Error loading available tables: ${error.message}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  // Handle file upload
  const handleFileUpload = async (fileType, file) => {
    if (!file) return;

    setLoading(true);
    const formData = new FormData();
    formData.append('file', file);
    formData.append('type', fileType);

    try {
      const response = await fetch(`${API_BASE_URL}/upload`, {
        method: 'POST',
        body: formData
      });

      const result = await response.json();

      if (result.success) {
        setUploadedFiles(prev => ({
          ...prev,
          [fileType]: result
        }));
        addMessage(`${fileType} file uploaded successfully: ${result.filename}`, 'success');
        
        if (['budget', 'sales', 'totalSales'].includes(fileType)) {
          if (productMtData || productValueData || sessionTotals) {
            clearProductData();
          }
          // Clear region data when new files are uploaded
          setRegionData({ mt: null, value: null });
        }
      } else {
        addMessage(result.error || 'Upload failed', 'error');
      }
    } catch (error) {
      addMessage(`Upload error: ${error.message}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  // Process sheet function
  const processSheet = async (fileType, sheetName, processingType = 'general') => {
    const fileInfo = uploadedFiles[fileType];
    if (!fileInfo) {
      addMessage('No file uploaded for this type', 'error');
      return;
    }

    setLoading(true);
    try {
      const response = await fetch(`${API_BASE_URL}/process-sheet`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: fileInfo.filepath,
          sheet_name: sheetName,
          processing_type: processingType
        })
      });

      const result = await response.json();

      if (result.success) {
        const dataKey = `${fileType}_${sheetName}`;
        setProcessedData(prev => ({
          ...prev,
          [dataKey]: result
        }));
        addMessage(`Sheet "${sheetName}" processed successfully (${result.shape[0]} rows, ${result.shape[1]} columns)`, 'success');
      } else {
        addMessage(result.error || 'Processing failed', 'error');
      }
    } catch (error) {
      addMessage(`Processing error: ${error.message}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  // Delete file function
  const deleteFile = async (fileType) => {
    const fileInfo = uploadedFiles[fileType];
    if (!fileInfo) return;

    try {
      const response = await fetch(`${API_BASE_URL}/delete-file`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: fileInfo.filepath
        })
      });

      const result = await response.json();
      if (result.success) {
        setUploadedFiles(prev => {
          const newFiles = { ...prev };
          delete newFiles[fileType];
          return newFiles;
        });
        addMessage(`File deleted successfully`, 'success');
        
        if (['budget', 'sales', 'totalSales'].includes(fileType)) {
          if (productMtData || productValueData || sessionTotals) {
            clearProductData();
          }
          // Clear region data when files are deleted
          setRegionData({ mt: null, value: null });
        }
      } else {
        addMessage(result.error || 'Delete failed', 'error');
      }
    } catch (error) {
      addMessage(`Delete error: ${error.message}`, 'error');
    }
  };

  // FIXED: Handle tab change with proper logging and session totals handling
  const handleTabChange = (newTab) => {
    console.log('🔄 Tab clicked:', newTab);
    console.log('Current state:', {
      activeTab,
      sessionTotals: !!sessionTotals,
      productMtData: !!productMtData,
      productValueData: !!productValueData,
      auditorFile: !!uploadedFiles.auditor,
      auditorSheet: !!selectedSheets.auditor
    });
    
    if (newTab === 'sales-analysis') {
      console.log('📊 Sales Analysis tab clicked');
      
      // Check requirements
      if (!uploadedFiles.auditor) {
        addMessage('❌ Auditor file required for Sales Analysis', 'error');
        return;
      }
      
      if (!selectedSheets.auditor) {
        addMessage('❌ Please select auditor sheet for Sales Analysis', 'error');
        return;
      }
      
      // If we don't have session totals but have product data, extract them
      if (!sessionTotals && (productMtData || productValueData)) {
        console.log('🔍 Extracting session totals from existing product data...');
        const extracted = extractSessionTotals(productMtData, productValueData);
        console.log('Extracted totals keys:', Object.keys(extracted));
        
        if (Object.keys(extracted).length > 0) {
          setSessionTotals(extracted);
          addMessage('✅ Session totals extracted from product data for Sales Analysis', 'info');
        } else {
          addMessage('⚠️ Could not extract session totals - please re-run Product Analysis', 'warning');
        }
      }
      
      // If still no session totals and no product data
      if (!sessionTotals && !productMtData && !productValueData) {
        addMessage('ℹ️ No session totals available. Run Product Analysis first for full integration.', 'info');
      }
      
      // Always allow tab switch - let component handle the rest
      setActiveTab(newTab);
      console.log('✅ Switching to Sales Analysis tab');
    } else {
      // For all other tabs, just switch
      setActiveTab(newTab);
    }
  };

  // Data preview component
  const DataPreview = ({ data, title, dataKey }) => {
    if (!data || !data.data) return null;

    const filteredData = searchTerm 
      ? data.data.filter(row => 
          Object.values(row).some(value => 
            String(value).toLowerCase().includes(searchTerm.toLowerCase())
          )
        )
      : data.data;

    return (
      <div className="data-preview">
        <div className="preview-header">
          <h4>{title}</h4>
          <div className="preview-actions">
            <button
              onClick={() => {
                const csvData = data.data.map(row => 
                  data.columns.map(col => row[col] || '').join(',')
                ).join('\n');
                const csvContent = [data.columns.join(','), csvData].join('\n');
                const blob = new Blob([csvContent], { type: 'text/csv' });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `${title.replace(/[^a-z0-9]/gi, '_')}.csv`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
              }}
              className="btn btn-secondary btn-small"
            >
              <Download size={14} />
              Export CSV
            </button>
          </div>
        </div>
        
        <div className="table-info">
          <span>Shape: {data.shape[0]} rows × {data.shape[1]} columns</span>
          {data.type && <span className="data-type">Type: {data.type}</span>}
          {data.table_name && <span className="data-type">Table: {data.table_name}</span>}
          {searchTerm && <span>Filtered: {filteredData.length} of {data.data.length} rows</span>}
        </div>
        
        <div className="search-container">
          <div className="search-input-wrapper">
            <Search size={16} className="search-icon" />
            <input
              type="text"
              placeholder="Search in data..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="search-input"
            />
            {searchTerm && (
              <button
                onClick={() => setSearchTerm('')}
                className="clear-search"
              >
                <X size={16} />
              </button>
            )}
          </div>
        </div>
        
        <div className="table-container">
          <table className="data-table">
            <thead>
              <tr>
                {data.columns.map((col, index) => (
                  <th key={index}>{col}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredData.slice(0, 100).map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {data.columns.map((col, colIndex) => (
                    <td key={colIndex}>{row[col] || ''}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        
        {filteredData.length > 100 && (
          <div className="table-note">
            Showing first 100 rows of {filteredData.length} filtered rows
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="app">
      <header className="app-header">
        <h1>📊 ACL Extraction</h1>
      </header>

      <div className="app-layout">
        <aside className="sidebar">
          <div className="sidebar-section">
            <h3 className="sidebar-title">📁 File Uploads</h3>
            
            <div className="file-upload-item">
              <label>Upload Budget Dataset</label>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => e.target.files[0] && handleFileUpload('budget', e.target.files[0])}
                className="file-input-sidebar"
              />
              {uploadedFiles.budget && (
                <div className="file-status">
                  <CheckCircle size={14} />
                  <span>{uploadedFiles.budget.filename}</span>
                  <button onClick={() => deleteFile('budget')} className="delete-btn">
                    <Trash2 size={12} />
                  </button>
                </div>
              )}
            </div>

            <div className="file-upload-item">
              <label>Upload Sales Dataset</label>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => e.target.files[0] && handleFileUpload('sales', e.target.files[0])}
                className="file-input-sidebar"
              />
              {uploadedFiles.sales && (
                <div className="file-status">
                  <CheckCircle size={14} />
                  <span>{uploadedFiles.sales.filename}</span>
                  <button onClick={() => deleteFile('sales')} className="delete-btn">
                    <Trash2 size={12} />
                  </button>
                </div>
              )}
            </div>

            <div className="file-upload-item">
              <label>Upload Total Sales Dataset</label>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => e.target.files[0] && handleFileUpload('totalSales', e.target.files[0])}
                className="file-input-sidebar"
              />
              {uploadedFiles.totalSales && (
                <div className="file-status">
                  <CheckCircle size={14} />
                  <span>{uploadedFiles.totalSales.filename}</span>
                  <button onClick={() => deleteFile('totalSales')} className="delete-btn">
                    <Trash2 size={12} />
                  </button>
                </div>
              )}
            </div>

            <div className="file-upload-item">
              <label>Upload Auditor Format File</label>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => e.target.files[0] && handleFileUpload('auditor', e.target.files[0])}
                className="file-input-sidebar"
              />
              {uploadedFiles.auditor && (
                <div className="file-status">
                  <CheckCircle size={14} />
                  <span>{uploadedFiles.auditor.filename}</span>
                  <button onClick={() => deleteFile('auditor')} className="delete-btn">
                    <Trash2 size={12} />
                  </button>
                </div>
              )}
            </div>
          </div>

          {uploadedFiles.auditor && (
            <div className="sidebar-section">
              <h3 className="sidebar-title">📄 Auditor Sheet Selection</h3>
              <select 
                className="sheet-select"
                onChange={(e) => setSelectedSheets(prev => ({ ...prev, auditor: e.target.value }))}
                value={selectedSheets.auditor || ''}
              >
                <option value="">Select Auditor Sheet</option>
                {uploadedFiles.auditor.sheet_names.map(sheet => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>
            </div>
          )}

          {selectedSheets.auditor && availableTables.length > 0 && (
            <div className="sidebar-section">
              <h3 className="sidebar-title">📊 Select Table to Display</h3>
              <div className="table-selection">
                {availableTables.map(tableOption => (
                  <label key={tableOption} className="radio-label">
                    <input
                      type="radio"
                      name="tableType"
                      value={tableOption}
                      checked={selectedTable === tableOption}
                      onChange={(e) => setSelectedTable(e.target.value)}
                    />
                    <span>{tableOption}</span>
                  </label>
                ))}
              </div>
            </div>
          )}

          {uploadedFiles.budget && (
            <div className="sidebar-section">
              <h3 className="sidebar-title">📄 Budget Sheet Selection</h3>
              <select 
                className="sheet-select"
                onChange={(e) => setSelectedSheets(prev => ({ ...prev, budget: e.target.value }))}
                value={selectedSheets.budget || ''}
              >
                <option value="">Select Budget Sheet</option>
                {uploadedFiles.budget.sheet_names.map(sheet => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>
            </div>
          )}

          {uploadedFiles.sales && (
            <div className="sidebar-section">
              <h3 className="sidebar-title">📄 Sales Sheet Selection</h3>
              <select 
                className="sheet-select"
                onChange={(e) => setSelectedSheets(prev => ({ ...prev, sales: e.target.value }))}
                value={selectedSheets.sales || ''}
              >
                <option value="">Select Sales Sheet</option>
                {uploadedFiles.sales.sheet_names.map(sheet => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>
            </div>
          )}

          {uploadedFiles.totalSales && (
            <div className="sidebar-section">
              <h3 className="sidebar-title">📄 Total Sales Selection</h3>
              <select 
                className="sheet-select"
                onChange={(e) => setSelectedSheets(prev => ({ ...prev, totalSales: e.target.value }))}
                value={selectedSheets.totalSales || ''}
              >
                <option value="">Select Total Sales Sheet</option>
                {uploadedFiles.totalSales.sheet_names.map(sheet => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>
            </div>
          )}

          {/* ENHANCED: Show stored files count in sidebar */}
          {storedFiles.length > 0 && (
            <div className="sidebar-section">
              <h3 className="sidebar-title">📁 Stored Files ({storedFiles.length})</h3>
              <div className="stored-files-summary">
                {storedFiles.filter(f => f.type === 'sales-analysis-excel').length > 0 && (
                  <div className="file-type-count">
                    📊 Sales Analysis: {storedFiles.filter(f => f.type === 'sales-analysis-excel').length}
                  </div>
                )}
                {storedFiles.filter(f => f.type && f.type.includes('region')).length > 0 && (
                  <div className="file-type-count">
                    🌍 Region Analysis: {storedFiles.filter(f => f.type && f.type.includes('region')).length}
                  </div>
                )}
                {storedFiles.filter(f => f.type && f.type.includes('product')).length > 0 && (
                  <div className="file-type-count">
                    📦 Product Analysis: {storedFiles.filter(f => f.type && f.type.includes('product')).length}
                  </div>
                )}
                {storedFiles.filter(f => f.type && (f.type.includes('tspw') || f.type.includes('eropw'))).length > 0 && (
                  <div className="file-type-count">
                    🔧 Other Analysis: {storedFiles.filter(f => f.type && (f.type.includes('tspw') || f.type.includes('eropw'))).length}
                  </div>
                )}
              </div>
              <button 
                onClick={clearAllStoredFiles}
                className="btn btn-secondary btn-small"
                style={{ marginTop: '8px', width: '100%' }}
              >
                <Trash2 size={12} />
                Clear All Files
              </button>
            </div>
          )}
        </aside>

        <main className="main-content">
          <nav className="tab-navigation">
            <button
              className={`tab ${activeTab === 'auditor' ? 'active' : ''}`}
              onClick={() => handleTabChange('auditor')}
            >
              Auditor Format
              {messages.filter(msg => msg.type === 'error').length > 0 && (
                <span className="message-count">
                  {messages.filter(msg => msg.type === 'error').length}
                </span>
              )}
            </button>
            <button
              className={`tab ${activeTab === 'sales-budget' ? 'active' : ''}`}
              onClick={() => handleTabChange('sales-budget')}
            >
              Sales and Budget Dataset
            </button>
            <button
              className={`tab ${activeTab === 'region-analysis' ? 'active' : ''}`}
              onClick={() => handleTabChange('region-analysis')}
            >
              Region Month-wise Analysis
              {(regionData.mt || regionData.value) && (
                <span className="integration-indicator">🌍</span>
              )}
              {storedFiles.filter(f => f.type && f.type.includes('region')).length > 0 && (
                <span className="integration-indicator">📁</span>
              )}
            </button>
            <button
              className={`tab ${activeTab === 'product-analysis' ? 'active' : ''}`}
              onClick={() => handleTabChange('product-analysis')}
            >
              Product-wise Analysis
              {(productMtData || productValueData) && (
                <span className="integration-indicator">📦</span>
              )}
              {sessionTotals && (
                <span className="integration-indicator">🔗</span>
              )}
              {storedFiles.filter(f => f.type && f.type.includes('product')).length > 0 && (
                <span className="integration-indicator">📁{storedFiles.filter(f => f.type && f.type.includes('product')).length}</span>
              )}
            </button>
            <button
              className={`tab ${activeTab === 'ts-pw-analysis' ? 'active' : ''}`}
              onClick={() => handleTabChange('ts-pw-analysis')}
            >
              TS-PW Analysis
              {storedFiles.filter(f => f.type && f.type.includes('tspw')).length > 0 && (
                <span className="integration-indicator">📁{storedFiles.filter(f => f.type && f.type.includes('tspw')).length}</span>
              )}
            </button>
            <button
              className={`tab ${activeTab === 'ero-pw-analysis' ? 'active' : ''}`}
              onClick={() => handleTabChange('ero-pw-analysis')}
            >
              ERO-PW Analysis
              {storedFiles.filter(f => f.type && f.type.includes('eropw')).length > 0 && (
                <span className="integration-indicator">📁{storedFiles.filter(f => f.type && f.type.includes('eropw')).length}</span>
              )}
            </button>
            <button
              className={`tab ${activeTab === 'sales-analysis' ? 'active' : ''}`}
              onClick={() => handleTabChange('sales-analysis')}
            >
              Sales Analysis Month-wise
              {/* ENHANCED: Better indicators for Sales Analysis tab */}
              {sessionTotals && (
                <span className="integration-indicator session-ready">🔗</span>
              )}
              {!sessionTotals && (productMtData || productValueData) && (
                <span className="integration-indicator session-pending">⚠️</span>
              )}
              {storedFiles.filter(f => f.type === 'sales-analysis-excel').length > 0 && (
                <span className="integration-indicator">📁{storedFiles.filter(f => f.type === 'sales-analysis-excel').length}</span>
              )}
            </button>
            <button
              className={`tab ${activeTab === 'combined-data' ? 'active' : ''}`}
              onClick={() => handleTabChange('combined-data')}
            >
              Combined Data Export
              {(regionData.mt || regionData.value) && (
                <span className="integration-indicator">📊</span>
              )}
              {storedFiles.length > 0 && (
                <span className="integration-indicator">📁{storedFiles.length}</span>
              )}
            </button>  
   
          </nav>

          {activeTab === 'auditor' && (
            <AuditorFormat 
              uploadedFiles={uploadedFiles}
              selectedSheets={selectedSheets}
              selectedTable={selectedTable}
              addMessage={addMessage}
              loading={loading}
              setLoading={setLoading}
            />
          )}

          {activeTab === 'sales-budget' && (
            <SalesFormat 
              uploadedFiles={uploadedFiles}
              selectedSheets={selectedSheets}
              addMessage={addMessage}
              loading={loading}
              setLoading={setLoading}
            />
          )}

          {activeTab === 'region-analysis' && (
            <RegionAnalysis 
              uploadedFiles={uploadedFiles}
              selectedSheets={selectedSheets}
              addMessage={addMessage}
              loading={loading}
              setLoading={setLoading}
              onAnalysisComplete={handleRegionAnalysisComplete}
              storedFiles={storedFiles}
              setStoredFiles={setStoredFiles}
            />
          )}

          {activeTab === 'product-analysis' && (
            <ProductAnalysis 
              uploadedFiles={uploadedFiles}
              selectedSheets={selectedSheets}
              addMessage={addMessage}
              loading={loading}
              setLoading={setLoading}
              onAnalysisComplete={handleProductAnalysisComplete}
              storedFiles={storedFiles}
              setStoredFiles={setStoredFiles}
              onFileAdd={addStoredFile}
            />
          )}

          {activeTab === 'ts-pw-analysis' && (
            <TsPwAnalysis 
              uploadedFiles={uploadedFiles}
              selectedSheets={selectedSheets}
              addMessage={addMessage}
              loading={loading}
              setLoading={setLoading}
              storedFiles={storedFiles}
              setStoredFiles={setStoredFiles}
              onFileAdd={addStoredFile}
            />
          )}

          {activeTab === 'ero-pw-analysis' && (
            <EroPwAnalysis 
              uploadedFiles={uploadedFiles}
              selectedSheets={selectedSheets}
              addMessage={addMessage}
              loading={loading}
              setLoading={setLoading}
              storedFiles={storedFiles}
              setStoredFiles={setStoredFiles}
              onFileAdd={addStoredFile}
            />
          )}

          {/* ENHANCED: Sales Analysis with automatic file storage */}
          {activeTab === 'sales-analysis' && (
            <SalesAnalysisMonthwise 
              uploadedFiles={uploadedFiles}
              selectedSheets={selectedSheets}
              addMessage={addMessage}
              loading={loading}
              setLoading={setLoading}
              productMtData={productMtData}
              productValueData={productValueData}
              sessionTotals={sessionTotals}
              setSessionTotals={setSessionTotals}
              // ENHANCED: Props for automatic file storage
              storedFiles={storedFiles}
              setStoredFiles={setStoredFiles}
              onFileAdd={handleSalesAnalysisFileGenerated}
              onFileRemove={removeStoredFile}
            />
          )}

          {activeTab === 'combined-data' && (
            <CombinedExcelManager 
              regionData={regionData}
              fiscalInfo={fiscalInfo}
              uploadedFiles={uploadedFiles}
              selectedSheets={selectedSheets}
              addMessage={addMessage}
              loading={loading}
              setLoading={setLoading}
              storedFiles={storedFiles}
              setStoredFiles={setStoredFiles}
              onFileAdd={addStoredFile}
              onFileRemove={removeStoredFile}
              onFilesClear={clearAllStoredFiles}
            />
          )}





        </main>
      </div>

      {loading && (
        <div className="loading-overlay">
          <div className="loading-spinner">
            <div className="spinner">
              <RefreshCw className="spinner-icon" />
            </div>
            <p>Processing...</p>
          </div>
        </div>
      )}
    </div>
  );
}

export default Auditor;
