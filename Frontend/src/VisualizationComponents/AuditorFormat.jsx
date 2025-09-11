import { useState, useEffect } from 'react';
import {
  FileSpreadsheet,
  Download,
  AlertCircle,
  Info,
  X,
  BarChart3,
  RefreshCw,
  Eye,
  Calendar,
  Database,
  Trash2
} from 'lucide-react';

const API_BASE_URL = process.env.REACT_APP_API_URL || '/api';

const AuditorFormat = ({ 
  uploadedFiles, 
  selectedSheets, 
  selectedTable,
  addMessage, 
  loading, 
  setLoading 
}) => {
  const [processedData, setProcessedData] = useState(null);
  const [analysisType, setAnalysisType] = useState(null);

  // Process auditor table whenever the selectedTable changes from sidebar
  useEffect(() => {
    if (uploadedFiles.auditor && selectedSheets.auditor && selectedTable) {
      processAuditorTable();
    } else {
      // Clear data when no file/sheet/table selected
      setProcessedData(null);
      setAnalysisType(null);
    }
  }, [uploadedFiles.auditor, selectedSheets.auditor, selectedTable]);

  const processAuditorTable = async (forceRefresh = false) => {
    if (!uploadedFiles.auditor || !selectedSheets.auditor || !selectedTable) return;

    setLoading(true);
    try {
      const response = await fetch(`${API_BASE_URL}/process-auditor-auto`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: uploadedFiles.auditor.filepath,
          sheet_name: selectedSheets.auditor,
          table_choice: selectedTable,
          force_refresh: forceRefresh
        })
      });

      const result = await response.json();

      if (result.success) {
        setProcessedData(result);
        setAnalysisType(result.analysis_type);
        
        const columnInfo = result.column_info;
        const message = `Processed ${result.table_name}: ${result.shape[0]} rows × ${result.shape[1]} columns (${columnInfo.act_columns.length} Act, ${columnInfo.gr_columns.length} Gr, ${columnInfo.ach_columns.length} Ach columns)`;
        addMessage(message, 'success');
      } else {
        addMessage(result.error || 'Failed to process auditor table', 'error');
        setProcessedData(null);
      }
    } catch (error) {
      addMessage(`Error processing auditor table: ${error.message}`, 'error');
      setProcessedData(null);
    } finally {
      setLoading(false);
    }
  };

  const exportData = async (format = 'csv') => {
    if (!processedData) return;

    setLoading(true);
    try {
      const response = await fetch(`${API_BASE_URL}/export-auditor-data`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: uploadedFiles.auditor.filepath,
          sheet_name: selectedSheets.auditor,
          table_choice: selectedTable,
          format: format
        })
      });

      const result = await response.json();

      if (result.success && format === 'csv') {
        // Create and download CSV
        const blob = new Blob([result.data], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = result.filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        addMessage(`Data exported successfully as ${format.toUpperCase()}`, 'success');
      } else {
        addMessage(result.error || 'Export failed', 'error');
      }
    } catch (error) {
      addMessage(`Export error: ${error.message}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  const refreshData = () => {
    processAuditorTable(true);
  };

  const DataPreview = ({ data, title }) => {
    if (!data || !data.data) return null;

    return (
      <div className="data-preview">
        <div className="preview-header">
          <div>
            <h4 className="text-xl font-semibold flex items-center gap-2">
              <FileSpreadsheet size={20} />
              {title}
            </h4>
            <p className="text-sm text-gray-600 mt-1">
              Processed with DataProcessor
            </p>
          </div>
          <div className="preview-actions">
            <button
              onClick={refreshData}
              className="btn btn-secondary btn-small"
              disabled={loading}
            >
              <RefreshCw size={14} className={loading ? 'animate-spin' : ''} />
              Refresh
            </button>
            <button
              onClick={() => exportData('csv')}
              className="btn btn-secondary btn-small"
              disabled={loading}
            >
              <Download size={14} />
              Export CSV
            </button>
          </div>
        </div>
        
        <div className="table-info">
          <span className="flex items-center gap-1">
            <Eye size={14} />
            Shape: {data.shape[0]} rows × {data.shape[1]} columns
          </span>
          <span className="data-type">Auditor Format</span>
        </div>
        
        <div className="table-container">
          <table className="data-table">
            <thead>
              <tr className="bg-violet-600 text-white">
                {data.columns.map((col, index) => (
                  <th key={index}>
                    {col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.data.slice(0, 100).map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {data.columns.map((col, colIndex) => {
                    const value = row[col];
                    // Format numeric values for better display
                    const displayValue = typeof value === 'number' ? 
                      (value % 1 === 0 ? value.toLocaleString() : value.toFixed(2)) : 
                      (value || '');
                    
                    return (
                      <td key={colIndex}>{displayValue}</td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        
        {data.data.length > 100 && (
          <div className="table-note">
            Showing first 100 rows of {data.data.length} total rows
          </div>
        )}
      </div>
    );
  };

  const LoadingState = () => (
    <div className="empty-state">
      <div className="flex items-center justify-center mb-4">
        <RefreshCw size={48} className="animate-spin text-blue-500" />
      </div>
      <h3>Processing auditor data...</h3>
      
    </div>
  );

  return (
    <div className="auditor-section">
      <div className="section-header">
        <h2 className="text-2xl font-bold mb-2 flex items-center gap-3">
          <Calendar size={28} className="text-blue-600" />
          📊 Auditor Format - Data Tables
        </h2>
        <p className="text-gray-600">
          Automatically processes and displays auditor tables with cleaned data including Gr/Ach columns based on Act columns.
        </p>
      </div>
      
      {uploadedFiles.auditor && selectedSheets.auditor ? (
        <div className="auditor-content">
          {loading ? (
            <LoadingState />
          ) : processedData ? (
            <div className="data-tables">
              <DataPreview
                data={processedData}
                title={processedData.table_name}
              />
            </div>
          ) : (
            <div className="empty-state">
              <BarChart3 size={48} />
              <h3>Select a table to view</h3>
              <p>Choose from the available tables in the sidebar to process and display the data</p>
            </div>
          )}
        </div>
      ) : (
        <div className="empty-state">
          <FileSpreadsheet size={48} />
          <h3>No auditor data selected</h3>
          <p>Upload an auditor format file and select a sheet from the sidebar to display the cleaned table with Gr/Ach columns</p>
        </div>
      )}
    </div>
  );
};

export default AuditorFormat;