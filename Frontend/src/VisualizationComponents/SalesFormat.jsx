import { useState, useEffect } from 'react';
import {
  FileSpreadsheet,
  Download,
  AlertCircle,
  Info,
  Search,
  X,
  BarChart3,
  RefreshCw,
  Eye,
  TrendingUp,
  Calendar,
  ChevronDown,
  ChevronUp
} from 'lucide-react';

const API_BASE_URL = process.env.REACT_APP_API_URL || '/api';

const SalesFormat = ({ 
  uploadedFiles, 
  selectedSheets, 
  addMessage, 
  loading, 
  setLoading 
}) => {
  const [processedData, setProcessedData] = useState({
    sales: {},
    budget: null,
    lastYear: null
  });
  const [searchTerm, setSearchTerm] = useState('');
  const [expandedSections, setExpandedSections] = useState({
    sales: true,
    budget: false,
    lastYear: false
  });

  // Debug logging
  console.log('SalesFormat State:', {
    uploadedFiles: {
      sales: !!uploadedFiles.sales,
      budget: !!uploadedFiles.budget,
      totalSales: !!uploadedFiles.totalSales
    },
    selectedSheets,
    processedData,
    loading
  });

  // Auto-process when files and sheets are selected
  useEffect(() => {
    const processSequentially = async () => {
      // Step 1: Process Sales sheets first
      if (uploadedFiles.sales && selectedSheets.sales) {
        setExpandedSections({ sales: true, budget: false, lastYear: false });
        
        if (Array.isArray(selectedSheets.sales)) {
          for (const sheet of selectedSheets.sales) {
            await processSalesSheet(sheet);
          }
        } else {
          await processSalesSheet(selectedSheets.sales);
        }
        
        // Wait a bit before moving to next section
        await new Promise(resolve => setTimeout(resolve, 1000));
      }

      // Step 2: Process Budget sheet
      if (uploadedFiles.budget && selectedSheets.budget) {
        setExpandedSections({ sales: false, budget: true, lastYear: false });
        await processBudgetSheet(selectedSheets.budget);
        
        // Wait a bit before moving to next section
        await new Promise(resolve => setTimeout(resolve, 1000));
      }

      // Step 3: Process Last Year sheet
      if (uploadedFiles.totalSales && selectedSheets.totalSales) {
        setExpandedSections({ sales: false, budget: false, lastYear: true });
        await processLastYearSheet(selectedSheets.totalSales);
      }
    };

    processSequentially();
  }, [uploadedFiles, selectedSheets]);

  const processSalesSheet = async (sheetName) => {
    if (!uploadedFiles.sales || !sheetName) return;

    console.log('Processing sales sheet:', sheetName);
    setLoading(true);
    try {
      const response = await fetch(`${API_BASE_URL}/process-sales-sheet`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: uploadedFiles.sales.filepath,
          sheet_name: sheetName
        })
      });

      const result = await response.json();
      console.log('Sales processing result:', result);

      if (result.success) {
        setProcessedData(prev => ({
          ...prev,
          sales: {
            ...prev.sales,
            [sheetName]: result
          }
        }));
        addMessage(`Sales sheet "${sheetName}" processed successfully (${result.shape[0]} rows)`, 'success');
      } else {
        addMessage(result.error || 'Failed to process sales sheet', 'error');
      }
    } catch (error) {
      addMessage(`Error processing sales sheet: ${error.message}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  const processBudgetSheet = async (sheetName) => {
    if (!uploadedFiles.budget || !sheetName) return;

    console.log('Getting budget sheet preview:', sheetName);
    setLoading(true);
    try {
      const response = await fetch(`${API_BASE_URL}/get-sheet-preview`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: uploadedFiles.budget.filepath,
          sheet_name: sheetName,
          preview_rows: 6
        })
      });

      const result = await response.json();
      console.log('Budget preview result:', result);

      if (result.success) {
        setProcessedData(prev => ({
          ...prev,
          budget: {
            ...result,
            data_type: 'budget'
          }
        }));
        addMessage(`Budget sheet "${sheetName}" preview loaded (${result.total_rows} total rows)`, 'success');
      } else {
        addMessage(result.error || 'Failed to load budget sheet preview', 'error');
      }
    } catch (error) {
      addMessage(`Error loading budget sheet preview: ${error.message}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  const processLastYearSheet = async (sheetName) => {
    if (!uploadedFiles.totalSales || !sheetName) return;

    console.log('Getting last year sheet preview:', sheetName);
    setLoading(true);
    try {
      const response = await fetch(`${API_BASE_URL}/get-sheet-preview`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: uploadedFiles.totalSales.filepath,
          sheet_name: sheetName,
          preview_rows: 6
        })
      });

      const result = await response.json();
      console.log('Last year preview result:', result);

      if (result.success) {
        setProcessedData(prev => ({
          ...prev,
          lastYear: {
            ...result,
            data_type: 'last_year'
          }
        }));
        addMessage(`Last year sheet "${sheetName}" preview loaded (${result.total_rows} total rows)`, 'success');
      } else {
        addMessage(result.error || 'Failed to load last year sheet preview', 'error');
      }
    } catch (error) {
      addMessage(`Error loading last year sheet preview: ${error.message}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  const exportData = async (dataType, sheetName, format = 'csv') => {
    const fileMapping = {
      sales: uploadedFiles.sales,
      budget: uploadedFiles.budget,
      lastYear: uploadedFiles.totalSales
    };

    const file = fileMapping[dataType];
    if (!file) return;

    setLoading(true);
    try {
      const response = await fetch(`${API_BASE_URL}/export-sales-data`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          filepath: file.filepath,
          sheet_name: sheetName,
          data_type: dataType === 'lastYear' ? 'last_year' : dataType,
          format: format
        })
      });

      const result = await response.json();

      if (result.success && format === 'csv') {
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

  const toggleSection = (section) => {
    setExpandedSections(prev => ({
      ...prev,
      [section]: !prev[section]
    }));
  };

  const DataPreview = ({ data, title, dataType, sheetName, isExpanded, onToggle }) => {
    if (!data || !data.data) return null;

    const filteredData = searchTerm 
      ? data.data.filter(row => 
          Object.values(row).some(value => 
            String(value).toLowerCase().includes(searchTerm.toLowerCase())
          )
        )
      : data.data;

    // Show only 6 rows initially
    const displayRows = isExpanded ? filteredData : filteredData.slice(0, 6);
    const hasMoreRows = filteredData.length > 6;

    return (
      <div className="data-preview mb-6">
        <div className="preview-header bg-white border rounded-lg shadow-sm">
          <div className="flex justify-between items-center p-4 border-b">
            <div className="flex items-center gap-3">
              <button
                onClick={onToggle}
                className="flex items-center gap-2 text-lg font-semibold hover:text-blue-600 transition-colors"
              >
                {dataType === 'sales' && <BarChart3 size={20} className="text-blue-600" />}
                {dataType === 'budget' && <Calendar size={20} className="text-green-600" />}
                {dataType === 'lastYear' && <Calendar size={20} className="text-orange-600" />}
                {title}
                {isExpanded ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
              </button>
              <span className="text-sm bg-gray-100 px-2 py-1 rounded">
                {data.shape[0]} rows × {data.shape[1]} columns
              </span>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={() => exportData(dataType, sheetName, 'csv')}
                className="btn btn-secondary btn-small flex items-center gap-1"
              >
                <Download size={14} />
                Export CSV
              </button>
            </div>
          </div>
          
          {isExpanded && (
            <>
              <div className="p-4 bg-gray-50 border-b">
                <div className="search-input-wrapper max-w-md">
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
                {searchTerm && (
                  <p className="text-sm text-gray-600 mt-2">
                    Showing {filteredData.length} of {data.data.length} rows
                  </p>
                )}
              </div>
              
              <div className="table-container max-h-96 overflow-auto">
                <table className="data-table w-full">
                  <thead className="sticky top-0 bg-white border-b">
                    <tr>
                      {data.columns.map((col, index) => {
                        let headerClass = 'px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider';
                        if (col.startsWith('Budget-')) headerClass += ' bg-blue-50 text-blue-800';
                        else if (col.startsWith('LY-')) headerClass += ' bg-orange-50 text-orange-800';
                        else if (col.includes('Value')) headerClass += ' bg-green-50 text-green-800';
                        else if (col.includes('MT')) headerClass += ' bg-purple-50 text-purple-800';
                        
                        return (
                          <th key={index} className={headerClass}>
                            {col}
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {displayRows.map((row, rowIndex) => (
                      <tr key={rowIndex} className="hover:bg-gray-50">
                        {data.columns.map((col, colIndex) => {
                          const value = row[col];
                          const displayValue = typeof value === 'number' ? 
                            (value % 1 === 0 ? value.toLocaleString() : value.toFixed(2)) : 
                            (value || '');
                          
                          return (
                            <td key={colIndex} className="px-3 py-2 text-sm text-gray-900 whitespace-nowrap">
                              {displayValue}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              
              {hasMoreRows && !isExpanded && (
                <div className="p-3 text-center bg-gray-50 border-t">
                  <button
                    onClick={onToggle}
                    className="text-blue-600 hover:text-blue-800 text-sm font-medium"
                  >
                    Show all {filteredData.length} rows
                  </button>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    );
  };

  const LoadingState = () => (
    <div className="empty-state">
      <div className="flex items-center justify-center mb-4">
        <RefreshCw size={48} className="animate-spin text-blue-500" />
      </div>
      <h3>Processing data...</h3>
      <p>Loading sales, budget, and last year data</p>
    </div>
  );

  const hasAnyData = () => {
    return Object.keys(processedData.sales).length > 0 || 
           processedData.budget || 
           processedData.lastYear;
  };

  const hasAnyFiles = () => {
    return uploadedFiles.sales || uploadedFiles.budget || uploadedFiles.totalSales;
  };

  return (
    <div className="sales-section">
      <div className="section-header mb-6">
        <h2 className="text-2xl font-bold mb-2 flex items-center gap-3">
          <TrendingUp size={28} className="text-green-600" />
          📊 Sales, Budget, and Last Year Dataset
        </h2>
        <p className="text-gray-600">
          Process and view sales data across multiple sheets, budget information, and last year comparisons
        </p>
      </div>
      
      {!hasAnyFiles() ? (
        <div className="empty-state">
          <FileSpreadsheet size={48} />
          <h3>No files uploaded</h3>
          <p>Upload Sales, Budget, or Last Year files using the sidebar to view data</p>
        </div>
      ) : loading ? (
        <LoadingState />
      ) : !hasAnyData() ? (
        <div className="empty-state">
          <AlertCircle size={48} />
          <h3>Select sheets to process</h3>
          <p>Select sheets from the sidebar dropdowns to process and view the data</p>
        </div>
      ) : (
        <div className="sales-content">
          {/* Step 1: Sales Data */}
          {Object.keys(processedData.sales).length > 0 && (
            <div className="step-section">
              <div className="step-header mb-3">
                <span className="step-number">1</span>
                <h3 className="text-lg font-semibold text-blue-600">Sales Data</h3>
              </div>
              {Object.entries(processedData.sales).map(([sheetName, data]) => (
                <DataPreview
                  key={sheetName}
                  data={data}
                  title={`Sales Data - ${sheetName}`}
                  dataType="sales"
                  sheetName={sheetName}
                  isExpanded={expandedSections.sales}
                  onToggle={() => toggleSection('sales')}
                />
              ))}
            </div>
          )}

          {/* Step 2: Budget Data Preview */}
          {processedData.budget && (
            <div className="step-section">
              <div className="step-header mb-3">
                <span className="step-number">2</span>
                <h3 className="text-lg font-semibold text-green-600">Budget Data Preview</h3>
              </div>
              <DataPreview
                data={processedData.budget}
                title={`Budget Preview - ${processedData.budget.sheet_name}`}
                dataType="budget"
                sheetName={processedData.budget.sheet_name}
                isExpanded={expandedSections.budget}
                onToggle={() => toggleSection('budget')}
              />
            </div>
          )}

          {/* Step 3: Last Year Data Preview */}
          {processedData.lastYear && (
            <div className="step-section">
              <div className="step-header mb-3">
                <span className="step-number">3</span>
                <h3 className="text-lg font-semibold text-orange-600">Last Year Data Preview</h3>
              </div>
              <DataPreview
                data={processedData.lastYear}
                title={`Last Year Preview - ${processedData.lastYear.sheet_name}`}
                dataType="lastYear"
                sheetName={processedData.lastYear.sheet_name}
                isExpanded={expandedSections.lastYear}
                onToggle={() => toggleSection('lastYear')}
              />
            </div>
          )}
        </div>
      )}
      
      <style jsx>{`
        .step-section {
          margin-bottom: 2rem;
        }
        
        .step-header {
          display: flex;
          align-items: center;
          gap: 12px;
        }
        
        .step-number {
          display: flex;
          align-items: center;
          justify-content: center;
          width: 32px;
          height: 32px;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          border-radius: 50%;
          font-weight: bold;
          font-size: 14px;
        }
        
        .data-table {
          border-collapse: collapse;
        }
        
        .search-input-wrapper {
          position: relative;
          display: flex;
          align-items: center;
        }
        
        .search-icon {
          position: absolute;
          left: 12px;
          color: #6b7280;
        }
        
        .search-input {
          width: 100%;
          padding: 8px 12px 8px 40px;
          border: 1px solid #d1d5db;
          border-radius: 6px;
          font-size: 14px;
        }
        
        .search-input:focus {
          outline: none;
          border-color: #3b82f6;
          box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
        }
        
        .clear-search {
          position: absolute;
          right: 12px;
          color: #6b7280;
          hover: #374151;
        }
        
        .btn {
          padding: 8px 16px;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s;
          border: none;
        }
        
        .btn-secondary {
          background: #f3f4f6;
          color: #374151;
        }
        
        .btn-secondary:hover {
          background: #e5e7eb;
        }
        
        .btn-small {
          padding: 6px 12px;
          font-size: 12px;
        }
        
        .empty-state {
          text-align: center;
          padding: 4rem 2rem;
          color: #6b7280;
        }
        
        .empty-state h3 {
          margin: 1rem 0 0.5rem;
          font-size: 1.5rem;
          color: #374151;
        }
      `}</style>
    </div>
  );
};

export default SalesFormat;