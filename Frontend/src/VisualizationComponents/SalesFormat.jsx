import { useState, useEffect } from 'react';
import {
  FileSpreadsheet,
  Download,
  AlertCircle,
  BarChart3,
  RefreshCw,
  Calendar,
  ChevronDown,
  ChevronUp
} from 'lucide-react';

const API_BASE_URL = process.env.REACT_APP_API_URL || 'http://localhost:5000/api';

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
  const [expandedSections, setExpandedSections] = useState({
    sales: true,
    budget: true,
    lastYear: true
  });

  // Auto-process when files and sheets are selected
  useEffect(() => {
    const processSequentially = async () => {
      if (uploadedFiles.sales && selectedSheets.sales) {
        if (Array.isArray(selectedSheets.sales)) {
          for (const sheet of selectedSheets.sales) {
            await processSalesSheet(sheet);
          }
        } else {
          await processSalesSheet(selectedSheets.sales);
        }
      }

      if (uploadedFiles.budget && selectedSheets.budget) {
        await processBudgetSheet(selectedSheets.budget);
      }

      if (uploadedFiles.totalSales && selectedSheets.totalSales) {
        await processLastYearSheet(selectedSheets.totalSales);
      }
    };

    processSequentially();
  }, [uploadedFiles, selectedSheets]);

  const processSalesSheet = async (sheetName) => {
    if (!uploadedFiles.sales || !sheetName) return;

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

    const displayRows = isExpanded ? data.data.slice(0, 10) : data.data.slice(0, 5);

    return (
      <div className="mb-6 border border-gray-200 rounded-lg bg-white">
        <div className="flex justify-between items-center p-4 border-b border-gray-200">
          <button
            onClick={onToggle}
            className="flex items-center gap-2 text-lg font-medium hover:text-blue-600"
          >
            {dataType === 'sales' && <BarChart3 size={20} />}
            {dataType === 'budget' && <Calendar size={20} />}
            {dataType === 'lastYear' && <Calendar size={20} />}
            {title}
            {isExpanded ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
          </button>
          <button
            onClick={() => exportData(dataType, sheetName, 'csv')}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium"
          >
            <Download size={16} />
            Export CSV
          </button>
        </div>
        
        {isExpanded && (
          <>            
            <div className="overflow-auto max-h-96">
              <table className="w-full">
                <thead className="bg-gray-50 sticky top-0">
                  <tr>
                    {data.columns.map((col, index) => (
                      <th key={index} className="px-4 py-2 text-left text-xs font-medium text-gray-700 uppercase">
                        {col}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {displayRows.map((row, rowIndex) => (
                    <tr key={rowIndex} className="border-t border-gray-200 hover:bg-gray-50">
                      {data.columns.map((col, colIndex) => {
                        const value = row[col];
                        const displayValue = typeof value === 'number' ? 
                          (value % 1 === 0 ? value.toLocaleString() : value.toFixed(2)) : 
                          (value || '');
                        
                        return (
                          <td key={colIndex} className="px-4 py-2 text-sm text-gray-900">
                            {displayValue}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            
            {data.data.length > displayRows.length && (
              <div className="p-3 text-center bg-gray-50 border-t border-gray-200">
                <span className="text-sm text-gray-600">
                  Showing {displayRows.length} of {data.data.length} rows
                </span>
              </div>
            )}
          </>
        )}
      </div>
    );
  };

  const hasAnyData = () => {
    return Object.keys(processedData.sales).length > 0 || 
           processedData.budget || 
           processedData.lastYear;
  };

  const hasAnyFiles = () => {
    return uploadedFiles.sales || uploadedFiles.budget || uploadedFiles.totalSales;
  };

  if (loading) {
    return (
      <div className="text-center py-12">
        <RefreshCw size={40} className="animate-spin text-blue-500 mx-auto mb-4" />
        <h3 className="text-lg font-medium mb-2">Processing data...</h3>
        <p className="text-gray-600">Loading sales, budget, and last year data</p>
      </div>
    );
  }

  if (!hasAnyFiles()) {
    return (
      <div className="text-center py-12">
        <FileSpreadsheet size={40} className="text-gray-400 mx-auto mb-4" />
        <h3 className="text-lg font-medium mb-2">No files uploaded</h3>
        <p className="text-gray-600">Upload Sales, Budget, or Last Year files using the sidebar to view data</p>
      </div>
    );
  }

  if (!hasAnyData()) {
    return (
      <div className="text-center py-12">
        <AlertCircle size={40} className="text-gray-400 mx-auto mb-4" />
        <h3 className="text-lg font-medium mb-2">Select sheets to process</h3>
        <p className="text-gray-600">Select sheets from the sidebar dropdowns to process and view the data</p>
      </div>
    );
  }

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold mb-2 flex items-center gap-2">
          <BarChart3 size={28} className="text-blue-600" />
          Sales Data Overview
        </h2>
        <p className="text-gray-600">
          View and export your sales, budget, and last year data
        </p>
      </div>

      {/* Sales Data */}
      {Object.keys(processedData.sales).length > 0 && (
        <div className="mb-8">
          <h3 className="text-lg font-semibold mb-4 text-blue-600">Sales Data</h3>
          {Object.entries(processedData.sales).map(([sheetName, data]) => (
            <DataPreview
              key={sheetName}
              data={data}
              title={`${sheetName}`}
              dataType="sales"
              sheetName={sheetName}
              isExpanded={expandedSections.sales}
              onToggle={() => toggleSection('sales')}
            />
          ))}
        </div>
      )}

      {/* Budget Data */}
      {processedData.budget && (
        <div className="mb-8">
          <h3 className="text-lg font-semibold mb-4 text-green-600">Budget Data</h3>
          <DataPreview
            data={processedData.budget}
            title={processedData.budget.sheet_name}
            dataType="budget"
            sheetName={processedData.budget.sheet_name}
            isExpanded={expandedSections.budget}
            onToggle={() => toggleSection('budget')}
          />
        </div>
      )}

      {/* Last Year Data */}
      {processedData.lastYear && (
        <div className="mb-8">
          <h3 className="text-lg font-semibold mb-4 text-orange-600">Last Year Data</h3>
          <DataPreview
            data={processedData.lastYear}
            title={processedData.lastYear.sheet_name}
            dataType="lastYear"
            sheetName={processedData.lastYear.sheet_name}
            isExpanded={expandedSections.lastYear}
            onToggle={() => toggleSection('lastYear')}
          />
        </div>
      )}
    </div>
  );
};

export default SalesFormat;
