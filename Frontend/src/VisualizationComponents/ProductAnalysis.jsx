import React, { useState, useEffect, useCallback } from 'react';
import { 
  Download,
  AlertCircle, 
  CheckCircle, 
  Info, 
  BarChart3,
  RefreshCw,
  Eye,
  TrendingUp,
  Package,
  Send,
  Database,
  Save,
  FileSpreadsheet
} from 'lucide-react';

const API_BASE_URL = process.env.REACT_APP_API_URL || 'http://localhost:5000/api';

const ProductAnalysis = ({ 
  uploadedFiles, 
  selectedSheets, 
  addMessage, 
  loading, 
  setLoading,
  onProductDataReceived,
  productDataIntegration,
  onAnalysisComplete,
  storedFiles = [],
  setStoredFiles = () => {},
  onFileAdd = () => {},
  salesMonthly
}) => {
  // State declarations
  const [activeSubTab, setActiveSubTab] = useState('tables'); // Default to tables tab
  const [productData, setProductData] = useState({
    mt: null,
    value: null
  });
  const [processing, setProcessing] = useState(false);
  const [fiscalInfo, setFiscalInfo] = useState({});
  const [integrationSent, setIntegrationSent] = useState(false);
  const [sessionTotalsExtracted, setSessionTotalsExtracted] = useState(false);
  const [autoExportEnabled, setAutoExportEnabled] = useState(true);
  const [hasCheckedSession, setHasCheckedSession] = useState(false);

  // Initialize from session storage
  useEffect(() => {
    const loadSessionData = () => {
      const storedData = sessionStorage.getItem('productAnalysisData');
      if (storedData) {
        try {
          const parsedData = JSON.parse(storedData);
          if (parsedData) {
            setProductData(parsedData.productData || { mt: null, value: null });
            setFiscalInfo(parsedData.fiscalInfo || {});
            setIntegrationSent(parsedData.integrationSent || false);
            setSessionTotalsExtracted(parsedData.sessionTotalsExtracted || false);
            addMessage('Loaded product analysis data from session', 'info');
          }
        } catch (error) {
          console.error('Error parsing session storage data:', error);
          addMessage('Failed to load session data', 'error');
        }
      }
    };

    loadSessionData();
  }, [addMessage]);

  // Save to session storage
  useEffect(() => {
    const saveSessionData = () => {
      const dataToStore = {
        productData,
        fiscalInfo,
        integrationSent,
        sessionTotalsExtracted,
        salesMonthly
      };
      sessionStorage.setItem('productAnalysisData', JSON.stringify(dataToStore));
    };

    saveSessionData();
  }, [productData, fiscalInfo, integrationSent, sessionTotalsExtracted, salesMonthly]);

  // Helper functions
  const canProcess = useCallback(() => {
    return uploadedFiles.budget && 
           selectedSheets.budget && 
           uploadedFiles.sales && 
           selectedSheets.sales;
  }, [uploadedFiles.budget, uploadedFiles.sales, selectedSheets.budget, selectedSheets.sales]);

  const extractSessionTotals = useCallback((mtData, valueData) => {
    try {
      const sessionTotals = {};
      
      if (mtData && mtData.data && mtData.data.length > 0) {
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
        } else {
          const lastRow = mtData.data[mtData.data.length - 1];
          sessionTotals.tonnage = lastRow;
        }
      }
      
      if (valueData && valueData.data && valueData.data.length > 0) {
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
        } else {
          const lastRow = valueData.data[valueData.data.length - 1];
          sessionTotals.value = lastRow;
        }
      }
      
      return sessionTotals;
    } catch (error) {
      console.error('Error extracting session totals:', error);
      return {};
    }
  }, []);

  const storeFileInSession = useCallback(async (fileBlob, fileName, analysisType, fileDescription) => {
    try {
      const timestamp = new Date().toISOString();
      const fileUrl = URL.createObjectURL(fileBlob);
      
      const storedFileData = {
        id: `product_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
        name: fileName,
        blob: fileBlob,
        size: fileBlob.size,
        url: fileUrl,
        createdAt: timestamp,
        fiscalYear: fiscalInfo.current_year || 'N/A',
        type: `product-${analysisType}`,
        source: 'Product Analysis',
        description: fileDescription,
        mtRecords: analysisType === 'mt' || analysisType.includes('combined') ? 
          (productData.mt?.data?.length || 0) : 0,
        valueRecords: analysisType === 'value' || analysisType.includes('combined') ? 
          (productData.value?.data?.length || 0) : 0,
        sheets: analysisType.includes('combined') ? 
          ['Combined Product Analysis'] :
          [`Product ${analysisType.toUpperCase()} Analysis`],
        analysisType: analysisType,
        metadata: {
          singleSheet: analysisType.includes('single-sheet'),
          bothTables: analysisType.includes('combined'),
          columnOrdered: true,
          fiscalYear: fiscalInfo.current_year,
          analysisType: 'product',
          autoGenerated: analysisType.includes('auto'),
          integration: {
            combined_data_ready: true,
            sync_status: 'ready'
          }
        },
        tags: ['product', 'analysis', analysisType.includes('combined') ? 'combined' : 'individual']
      };

      if (typeof setStoredFiles === 'function') {
        setStoredFiles(prev => {
          const oldProductFiles = prev.filter(f => f.type && f.type.includes('product'));
          oldProductFiles.forEach(file => {
            if (file.url) {
              URL.revokeObjectURL(file.url);
            }
          });
          
          const otherFiles = prev.filter(f => !f.type || !f.type.includes('product'));
          return [storedFileData, ...otherFiles];
        });
        
        addMessage(`💾 ${fileName} stored (previous Product file replaced)`, 'success');
        return storedFileData;
      }
      
      if (typeof onFileAdd === 'function') {
        onFileAdd(storedFileData);
        addMessage(`💾 ${fileName} stored in Combined Excel Manager`, 'success');
        return storedFileData;
      }

      addMessage('⚠️ File generated but could not be stored', 'warning');
      return null;
    } catch (error) {
      console.error('Error storing Product file:', error);
      addMessage(`❌ Error storing file: ${error.message}`, 'error');
      return null;
    }
  }, [fiscalInfo, productData, addMessage, onFileAdd, setStoredFiles]);

  const fixColumnOrdering = useCallback((data, dataType = 'product') => {
    if (!data || !data.data || !data.columns) return data;
    
    const correctOrder = [
      'Product', 'Product Name',
      'Budget',
      'LY', 'Last Year', 'LY Total',
      'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar',
      'YTD-25-26 (Apr to Mar) Budget',
      'YTD-24-25 (Apr to Mar)LY',
      'YTD', 'Act-YTD-25-26 (Apr to Mar)', 'Act', 'Actual', 'Total',
      'Gr-YTD-25-26 (Apr to Mar)', 'Gr', 'Growth', 'Gr%', 'Growth%',
      'Ach-YTD-25-26 (Apr to Mar)', 'Ach', 'Ach%', 'Achievement', 'Achievement%'
    ];
    
    const sortColumns = (columns) => {
      return [...columns].sort((a, b) => {
        const ytdBudgetPattern = /YTD.*Budget/i;
        const ytdLYPattern = /YTD.*LY/i;
        const ytdActPattern = /Act.*YTD/i;
        const ytdGrPattern = /Gr.*YTD/i;
        const ytdAchPattern = /Ach.*YTD/i;
        
        if (ytdBudgetPattern.test(a) && ytdLYPattern.test(b)) return -1;
        if (ytdLYPattern.test(a) && ytdBudgetPattern.test(b)) return 1;
        if (ytdBudgetPattern.test(a) && ytdActPattern.test(b)) return -1;
        if (ytdActPattern.test(a) && ytdBudgetPattern.test(b)) return 1;
        if (ytdLYPattern.test(a) && ytdActPattern.test(b)) return -1;
        if (ytdActPattern.test(a) && ytdLYPattern.test(b)) return 1;
        if (ytdActPattern.test(a) && ytdGrPattern.test(b)) return -1;
        if (ytdGrPattern.test(a) && ytdActPattern.test(b)) return 1;
        if (ytdGrPattern.test(a) && ytdAchPattern.test(b)) return -1;
        if (ytdAchPattern.test(a) && ytdGrPattern.test(b)) return 1;
        
        let indexA = correctOrder.findIndex(pattern => 
          a.toLowerCase().includes(pattern.toLowerCase()) || 
          pattern.toLowerCase().includes(a.toLowerCase()) ||
          a === pattern
        );
        let indexB = correctOrder.findIndex(pattern => 
          b.toLowerCase().includes(pattern.toLowerCase()) || 
          pattern.toLowerCase().includes(b.toLowerCase()) ||
          b === pattern
        );
        
        if (indexA === -1) indexA = 999;
        if (indexB === -1) indexB = 999;
        
        return indexA - indexB;
      });
    };
    
    const sortedColumns = sortColumns(data.columns);
    const reorderedData = data.data.map(row => {
      const newRow = {};
      sortedColumns.forEach(col => {
        newRow[col] = row[col];
      });
      return newRow;
    });
    
    return {
      ...data,
      columns: sortedColumns,
      data: reorderedData
    };
  }, []);

  const autoGenerateProductReport = useCallback(async (mtData, valueData, fiscalYear) => {
    if (!autoExportEnabled) return;
    
    const hasMtData = mtData && mtData.data && mtData.data.length > 0;
    const hasValueData = valueData && valueData.data && valueData.data.length > 0;
    
    if (!hasMtData && !hasValueData) return;

    try {
      const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
      const fileName = `product_auto_combined_${fiscalYear || 'report'}_${timestamp}.xlsx`;

      const fixedMtData = hasMtData ? fixColumnOrdering(mtData, 'mt') : null;
      const fixedValueData = hasValueData ? fixColumnOrdering(valueData, 'value') : null;

      const response = await fetch(`${API_BASE_URL}/product/export/combined-single-sheet`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          mt_data: fixedMtData?.data || [],
          value_data: fixedValueData?.data || [],
          mt_columns: fixedMtData?.columns || [],
          value_columns: fixedValueData?.columns || [],
          fiscal_year: fiscalYear || '',
          include_both_tables: true,
          single_sheet: true,
          column_order_fixed: true,
          auto_generated: true
        })
      });

      if (response.ok) {
        const blob = await response.blob();
        const fileDescription = `Auto-generated combined product analysis with MT (${hasMtData ? mtData.data.length : 0} records) and Value (${hasValueData ? valueData.data.length : 0} records) for fiscal year ${fiscalYear}`;
        await storeFileInSession(blob, fileName, 'combined-auto', fileDescription);
        addMessage(`🤖 Auto-generated: ${fileName}`, 'success');
      } else {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to generate Excel file');
      }
    } catch (error) {
      console.error('Auto-generation error:', error);
      addMessage(`❌ Auto-generation error: ${error.message}`, 'error');
    }
  }, [autoExportEnabled, storeFileInSession, addMessage, fixColumnOrdering]);

  // Auto-generate when data changes
  useEffect(() => {
    if ((productData.mt || productData.value) && autoExportEnabled && !processing) {
      const timer = setTimeout(() => {
        autoGenerateProductReport(
          productData.mt,
          productData.value,
          fiscalInfo.current_year
        );
      }, 2000);
      
      return () => clearTimeout(timer);
    }
  }, [productData.mt, productData.value, autoExportEnabled, processing, autoGenerateProductReport, fiscalInfo]);

  // Main processing function
  const processProductAnalysis = useCallback(async (analysisType = 'both') => {
    if (!canProcess()) {
      addMessage('Budget and Sales files are required for product analysis', 'error');
      return;
    }

    setProcessing(true);
    setLoading(true);
    setIntegrationSent(false);
    setSessionTotalsExtracted(false);
    
    try {
      const salesFiles = [];
      if (uploadedFiles.sales && selectedSheets.sales) {
        salesFiles.push({
          filepath: uploadedFiles.sales.filepath,
          sheet_name: selectedSheets.sales
        });
      }

      let lastYearFile = null;
      if (uploadedFiles.totalSales && selectedSheets.totalSales) {
        lastYearFile = {
          filepath: uploadedFiles.totalSales.filepath,
          sheet_name: selectedSheets.totalSales
        };
      }

      const requestData = {
        budget_filepath: uploadedFiles.budget.filepath,
        budget_sheet: selectedSheets.budget,
        sales_files: salesFiles,
        last_year_file: lastYearFile,
        analysis_type: analysisType
      };

      addMessage('🔄 Processing Product-wise analysis...', 'info');

      const response = await fetch(`${API_BASE_URL}/product/process`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestData)
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result = await response.json();

      if (result.success) {
        let mtData = null;
        let valueData = null;
        let resultFiscalInfo = null;
        
        if (result.type === 'merge_analysis') {
          mtData = result.mt_data;
          valueData = result.value_data;
          resultFiscalInfo = result.fiscal_info;
        } else {
          if (analysisType === 'both' || analysisType === 'merge') {
            mtData = result.mt_data || (analysisType === 'mt' ? result : null);
            valueData = result.value_data || (analysisType === 'value' ? result : null);
          } else if (analysisType === 'mt') {
            mtData = result;
          } else if (analysisType === 'value') {
            valueData = result;
          }
          resultFiscalInfo = result.fiscal_info;
        }
        
        setProductData(prev => ({
          ...prev,
          mt: mtData,
          value: valueData
        }));
        
        if (resultFiscalInfo) {
          setFiscalInfo(resultFiscalInfo);
        }

        // Auto-switch to available tab (like TS-PW)
        if (valueData && valueData.data && valueData.data.length > 0) {
          setActiveSubTab('value');
        } else if (mtData && mtData.data && mtData.data.length > 0) {
          setActiveSubTab('mt');
        }

        // Create complete analysis result including salesMonthly
        const analysisResult = {
          mtData,
          valueData,
          fiscalInfo: resultFiscalInfo,
          salesMonthly,
          timestamp: new Date().toISOString()
        };

        if (onProductDataReceived) {
          onProductDataReceived(analysisResult);
          setIntegrationSent(true);
          addMessage('📤 Product data sent to Sales Module for ACCLLP integration', 'success');
        }

        if (onAnalysisComplete) {
          const sessionTotals = extractSessionTotals(mtData, valueData);
          
          if (Object.keys(sessionTotals).length > 0) {
            setSessionTotalsExtracted(true);
            
            onAnalysisComplete({
              ...analysisResult,
              sessionTotals,
              rowCounts: {
                mt: mtData?.data?.length || 0,
                value: valueData?.data?.length || 0
              }
            });
            addMessage('🔗 Session totals extracted for Sales Analysis integration!', 'success');
          }
        }
        
        // Enhanced success messages based on data availability
        const mtRows = mtData?.data?.length || 0;
        const valueRows = valueData?.data?.length || 0;
        
        if (mtRows > 0 && valueRows > 0) {
          addMessage(`✅ Product analysis completed successfully - MT: ${mtRows} rows, Value: ${valueRows} rows`, 'success');
        } else if (mtRows > 0) {
          addMessage(`✅ Product analysis completed - MT: ${mtRows} rows (Value data not available)`, 'success');
        } else if (valueRows > 0) {
          addMessage(`✅ Product analysis completed - Value: ${valueRows} rows (MT data not available)`, 'success');
        } else {
          addMessage('⚠️ Product analysis completed but no data was generated. Please check your input files.', 'warning');
        }
      } else {
        addMessage(result.error || 'Product analysis failed', 'error');
      }
    } catch (error) {
      addMessage(`❌ Product analysis error: ${error.message}`, 'error');
    } finally {
      setLoading(false);
      setProcessing(false);
    }
  }, [canProcess, uploadedFiles, selectedSheets, addMessage, setLoading, 
      onProductDataReceived, onAnalysisComplete, extractSessionTotals, salesMonthly]);

  // Auto-process on data availability (moved after function definition)
  useEffect(() => {
    const timer = setTimeout(() => {
      if (canProcess() && !processing && !productData.mt && !productData.value) {
        console.log('🔄 Auto-processing Product analysis...');
        processProductAnalysis('both');
      }
    }, 1000);

    return () => clearTimeout(timer);
  }, [uploadedFiles, selectedSheets, canProcess, processProductAnalysis, processing, productData]);

  const exportMergedTables = async () => {
    const hasMtData = productData.mt && productData.mt.data && productData.mt.data.length > 0;
    const hasValueData = productData.value && productData.value.data && productData.value.data.length > 0;
    
    if (!hasMtData && !hasValueData) {
      addMessage('No product data available to export', 'error');
      return;
    }

    setLoading(true);
    
    try {
      const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
      const fileName = `product_combined_manual_${fiscalInfo.current_year || 'report'}_${timestamp}.xlsx`;

      const fixedMtData = hasMtData ? fixColumnOrdering(productData.mt, 'mt') : null;
      const fixedValueData = hasValueData ? fixColumnOrdering(productData.value, 'value') : null;

      const response = await fetch(`${API_BASE_URL}/product/export/combined-single-sheet`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          mt_data: fixedMtData?.data || [],
          value_data: fixedValueData?.data || [],
          mt_columns: fixedMtData?.columns || [],
          value_columns: fixedValueData?.columns || [],
          fiscal_year: fiscalInfo.current_year || '',
          include_both_tables: true,
          single_sheet: true,
          column_order_fixed: true,
          column_order_priority: [
            'Product', 'Budget', 'LY', 
            'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar',
            'YTD-25-26 (Apr to Mar) Budget', 'YTD-24-25 (Apr to Mar)LY', 'Act-YTD-25-26 (Apr to Mar)', 
            'Gr-YTD-25-26 (Apr to Mar)', 'Ach-YTD-25-26 (Apr to Mar)'
          ]
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      addMessage('📁 Product analysis single-sheet exported successfully', 'success');
    } catch (error) {
      addMessage(`❌ Export error: ${error.message}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  const formatNumber = (value) => {
    if (typeof value === 'number') {
      return value.toLocaleString('en-IN', { 
        minimumFractionDigits: 2, 
        maximumFractionDigits: 2 
      });
    }
    return value;
  };

  const getProductFiles = () => {
    return storedFiles.filter(f => f.type && f.type.includes('product'));
  };

  const DataTable = ({ data, title, analysisType }) => {
    if (!data || !data.data) return null;

    const columns = data.columns || [];

    return (
      <div className="data-preview">
        <div className="preview-header">
          <h4>{title}</h4>
        </div>
        
        <div className="table-container">
          <table className="data-table">
            <thead>
              <tr>
                {columns.map((col, index) => (
                  <th key={index} className={index === 0 ? 'sticky-col product-header' : 'data-header'}>{col}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.data.slice(0, 100).map((row, rowIndex) => {
                const isTotal = row[columns[0]]?.toString().toUpperCase().includes('TOTAL');
                return (
                  <tr key={rowIndex} className={isTotal ? 'total-row' : ''}>
                    {columns.map((col, colIndex) => {
                      const value = row[col];
                      const isNumeric = typeof value === 'number' && !isNaN(value);
                      const formattedValue = isNumeric ? formatNumber(value) : (value || '');
                      
                      return (
                        <td 
                          key={colIndex} 
                          className={`${colIndex === 0 ? 'sticky-col product-cell' : 'data-cell'} ${isNumeric ? 'numeric' : ''}`}
                        >
                          {formattedValue}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
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

  const TablesPreviewTab = () => {
    const productFiles = getProductFiles();
    
    return (
      <div className="tables-preview-section">
        {/* Removed the violet/purple gradient preview-intro section */}

        <div className="table-section">
          <div className="table-section-header">
            <h4> Product-wise SALES in Tonage Table</h4>
            <div className="table-status">
              {productData.mt ? (
                <span className="status-badge available">✅ Available ({productData.mt?.data?.length || 0} products)</span>
              ) : (
                <span className="status-badge unavailable">❌ Not Available</span>
              )}
            </div>
          </div>

          {productData.mt ? (
            <DataTable
              data={productData.mt}
              title={`Product-wise Budget and Actual Tonage (Month-wise) [${fiscalInfo.fiscal_year_str || '25-26'}]`}
              analysisType="mt"
            />
          ) : (
            <div className="table-empty-state">
              <Package size={48} />
              <h4>No Product Tonage Data Available</h4>
              <p>
                {!canProcess() ? 
                  "Upload Budget and Sales files to start analysis" :
                  "Click 'Generate Analysis' to generate MT data"
                }
              </p>
              {canProcess() && (
                <button
                  onClick={() => processProductAnalysis('mt')}
                  className="btn btn-primary"
                  disabled={loading || processing}
                >
                  <Package size={16} />
                  {processing ? 'Processing...' : 'Generate MT Analysis'}
                </button>
              )}
            </div>
          )}
        </div>

        <div className="table-separator"></div>

        <div className="table-section">
          <div className="table-section-header">
            <h4>Product-wise SALES in Value Table</h4>
            <div className="table-status">
              {productData.value ? (
                <span className="status-badge available">✅ Available ({productData.value?.data?.length || 0} products)</span>
              ) : (
                <span className="status-badge unavailable">❌ Not Available</span>
              )}
            </div>
          </div>

          {productData.value ? (
            <DataTable
              data={productData.value}
              title={`Product-wise Budget and Actual Value (Month-wise) [${fiscalInfo.fiscal_year_str || '25-26'}]`}
              analysisType="value"
            />
          ) : (
            <div className="table-empty-state">
              <Package size={48} />
              <h4>No Product Value Data Available</h4>
              <p>
                {!canProcess() ? 
                  "Upload Budget and Sales files to start analysis" :
                  "Click 'Generate Analysis' to generate Value data"
                }
              </p>
              {canProcess() && (
                <button
                  onClick={() => processProductAnalysis('value')}
                  className="btn btn-primary"
                  disabled={loading || processing}
                >
                  <Package size={16} />
                  {processing ? 'Processing...' : 'Generate Value Analysis'}
                </button>
                )}
            </div>
          )}
        </div>

        <div className="tables-summary">
          <h4> Product Analysis Export</h4>
          {(productData.mt || productData.value) ? (
            <div className="export-merged-section">
              <div className="export-actions">
                <button
                  onClick={exportMergedTables}
                  className="btn btn-secondary"
                  disabled={loading}
                  title="Download product report"
                >
                  <Download size={16} />
                  {loading ? 'Processing...' : 'Download Product Report'}
                </button>
              </div>
            </div>
          ) : (
            <div className="export-empty-state">
              <p>No product data available to export. Please generate analysis data first.</p>
              {canProcess() && (
                <button
                  onClick={() => processProductAnalysis('both')}
                  className="btn btn-primary"
                  disabled={loading || processing}
                >
                  <BarChart3 size={16} />
                  {processing ? 'Processing...' : 'Generate Both Tables'}
                </button>
              )}
            </div>
          )}
        </div>
      </div>
    );
  };

  // Enhanced Initial State Component
  const InitialStateComponent = () => (
    <div className="initial-state">
      <div className="initial-state-header">
        <Package size={64} />
        <h3>Welcome to Product-wise Analysis</h3>
        <p>Analyze your products' sales performance across tonage and value metrics with detailed month-wise breakdown and fiscal year comparisons.</p>
      </div>

                <div className="requirements-checklist">
        <h4>📋 Requirements Checklist</h4>
        <div className="checklist-items">
          <div className={`checklist-item ${uploadedFiles.budget ? 'completed' : 'pending'}`}>
            <div className="checklist-icon">
              {uploadedFiles.budget ? <CheckCircle size={20} /> : <AlertCircle size={20} />}
            </div>
            <div className="checklist-content">
              <span className="checklist-title">Budget File</span>
              <span className="checklist-status">
                {uploadedFiles.budget ? `✅ ${uploadedFiles.budget.name}` : '❌ Required'}
              </span>
            </div>
          </div>

          <div className={`checklist-item ${selectedSheets.budget ? 'completed' : 'pending'}`}>
            <div className="checklist-icon">
              {selectedSheets.budget ? <CheckCircle size={20} /> : <AlertCircle size={20} />}
            </div>
            <div className="checklist-content">
              <span className="checklist-title">Budget Sheet Selected</span>
              <span className="checklist-status">
                {selectedSheets.budget ? `✅ ${selectedSheets.budget}` : '❌ Required'}
              </span>
            </div>
          </div>

          <div className={`checklist-item ${uploadedFiles.sales ? 'completed' : 'pending'}`}>
            <div className="checklist-icon">
              {uploadedFiles.sales ? <CheckCircle size={20} /> : <AlertCircle size={20} />}
            </div>
            <div className="checklist-content">
              <span className="checklist-title">Sales File</span>
              <span className="checklist-status">
                {uploadedFiles.sales ? `✅ ${uploadedFiles.sales.name}` : '❌ Required'}
              </span>
            </div>
          </div>

          <div className={`checklist-item ${selectedSheets.sales ? 'completed' : 'pending'}`}>
            <div className="checklist-icon">
              {selectedSheets.sales ? <CheckCircle size={20} /> : <AlertCircle size={20} />}
            </div>
            <div className="checklist-content">
              <span className="checklist-title">Sales Sheet Selected</span>
              <span className="checklist-status">
                {selectedSheets.sales ? `✅ ${selectedSheets.sales}` : '❌ Required'}
              </span>
            </div>
          </div>

          <div className={`checklist-item ${uploadedFiles.totalSales ? 'completed' : 'optional'}`}>
            <div className="checklist-icon">
              {uploadedFiles.totalSales ? <CheckCircle size={20} /> : <Info size={20} />}
            </div>
            <div className="checklist-content">
              <span className="checklist-title">Last Year File (Optional)</span>
              <span className="checklist-status">
                {uploadedFiles.totalSales ? `✅ ${uploadedFiles.totalSales.name}` : '📋 Optional for LY comparison'}
              </span>
            </div>
          </div>
        </div>
      </div>

      <div className="getting-started">
        <h4>🚀 Getting Started</h4>
        <div className="steps">
          <div className="step">
            <span className="step-number">1</span>
            <div className="step-content">
              <h5>Upload Required Files</h5>
              <p>Upload your Budget and Sales Excel files using the File Upload section.</p>
            </div>
          </div>
          <div className="step">
            <span className="step-number">2</span>
            <div className="step-content">
              <h5>Select Sheets</h5>
              <p>Choose the appropriate sheets from your uploaded files that contain the data.</p>
            </div>
          </div>
          <div className="step">
            <span className="step-number">3</span>
            <div className="step-content">
              <h5>Generate Analysis</h5>
              <p>Click the "Generate Analysis" button to process your product-wise data.</p>
            </div>
          </div>
        </div>
      </div>

      <div className="action-section">
        {canProcess() ? (
          <button
            onClick={() => processProductAnalysis('both')}
            className="btn btn-primary btn-large"
            disabled={loading || processing}
          >
            <BarChart3 size={20} />
            {processing ? 'Processing Analysis...' : 'Generate Product Analysis'}
          </button>
        ) : (
          <div className="action-disabled">
            <p>Complete the requirements above to start product analysis</p>
          </div>
        )}
      </div>
    </div>
  );

  return (
    <div className="product-analysis-section">
      <div className="section-header">
        <h2> Product-wise Analysis</h2>
        <div className="header-actions">
          <button
            onClick={() => processProductAnalysis('both')}
            className="btn btn-primary"
            disabled={!canProcess() || loading || processing}
          >
            {processing ? <RefreshCw size={16} /> : <BarChart3 size={16} />}
            {processing ? 'Processing...' : 'Generate Analysis'}
          </button>

          {getProductFiles().length > 0 && (
            <div className="integration-status-badge">
              <Package size={14} />
              <span>Latest file stored</span>
              <span className="latest-indicator">Latest Only</span>
            </div>
          )}
        </div>
      </div>

      <div className="stats-overview">
        <div className="stat-card">
          <Package className="stat-icon" />
          <div>
            <span className="stat-number">{productData.mt?.data?.length || 0}</span>
            <span className="stat-label">Product Tonage Records</span>
          </div>
        </div>
        <div className="stat-card">
          <Package className="stat-icon" />
          <div>
            <span className="stat-number">{productData.value?.data?.length || 0}</span>
            <span className="stat-label">Product Value Records</span>
          </div>
        </div>
        <div className="stat-card">
          <Database className="stat-icon" />
          <div>
            <span className="stat-number">
              {getProductFiles().length > 0 ? '1' : '0'}
            </span>
            <span className="stat-label">Latest File</span>
          </div>
        </div>
      </div>

      {/* Always show navigation and content when files are available */}
      <nav className="sub-tab-navigation">
        <button
          className={`sub-tab ${activeSubTab === 'mt' ? 'active' : ''}`}
          onClick={() => setActiveSubTab('mt')}
        >
          <Package size={16} />
          SALES in Tonage
          {productData.mt && <span className="data-indicator"></span>}
        </button>
        <button
          className={`sub-tab ${activeSubTab === 'value' ? 'active' : ''}`}
          onClick={() => setActiveSubTab('value')}
        >
          <Package size={16} />
          SALES in Value
          {productData.value && <span className="data-indicator"></span>}
        </button>
        <button
          className={`sub-tab ${activeSubTab === 'tables' ? 'active' : ''}`}
          onClick={() => setActiveSubTab('tables')}
        >
          <Eye size={16} />
          Tables Preview
          {getProductFiles().length > 0 && <span className="data-indicator"></span>}
        </button>
      </nav>

      <div className="sub-tab-content">
        {activeSubTab === 'mt' && (
          <div className="mt-analysis">
            {productData.mt ? (
              <DataTable
                data={productData.mt}
                title={`Product-wise Budget and Actual Tonage (Month-wise) [${fiscalInfo.current_year || '25-26'}]`}
                analysisType="mt"
              />
            ) : (
              <div className="empty-state">
                <Package size={48} />
                <h3>No MT analysis data</h3>
                <p>
                  {!canProcess() ? 
                    "Upload Budget and Sales files to start analysis" :
                    processing ? "Processing MT data..." : "Waiting for data to be processed"
                  }
                </p>
                {canProcess() && !processing && (
                  <button
                    onClick={() => processProductAnalysis('mt')}
                    className="btn btn-primary"
                    disabled={loading || processing}
                  >
                    <Package size={16} />
                    Generate MT Analysis
                  </button>
                )}
              </div>
            )}
          </div>
        )}

        {activeSubTab === 'value' && (
          <div className="value-analysis">
            {productData.value ? (
              <DataTable
                data={productData.value}
                title={`Product-wise Budget and Actual Value (Month-wise) [${fiscalInfo.current_year || '25-26'}]`}
                analysisType="value"
              />
            ) : (
              <div className="empty-state">
                <Package size={48} />
                <h3>No Value analysis data</h3>
                <p>
                  {!canProcess() ? 
                    "Upload Budget and Sales files to start analysis" :
                    processing ? "Processing Value data..." : "Waiting for data to be processed"
                  }
                </p>
                {canProcess() && !processing && (
                  <button
                    onClick={() => processProductAnalysis('value')}
                    className="btn btn-primary"
                    disabled={loading || processing}
                  >
                    <Package size={16} />
                    Generate Value Analysis
                  </button>
                  )}
              </div>
            )}
          </div>
        )}

        {activeSubTab === 'tables' && <TablesPreviewTab />}
      </div>

      {processing && (
        <div className="processing-indicator">
          <RefreshCw size={24} />
          <span>Processing Product-wise analysis...</span>
        </div>
      )}

      <style jsx>{`
        .product-analysis-section {
          padding: 20px;
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Oxygen', 'Ubuntu', 'Cantarell', 'Fira Sans', 'Droid Sans', 'Helvetica Neue', sans-serif;
          background-color: #f8f9fa;
          min-height: 100vh;
        }

        .section-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 20px;
          background: white;
          padding: 20px;
          border-radius: 8px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .section-header h2 {
          margin: 0;
          color: #333;
          font-size: 24px;
          font-weight: 600;
        }

        .header-actions {
          display: flex;
          gap: 10px;
          align-items: center;
        }

        .integration-status-badge {
          display: flex;
          align-items: center;
          gap: 6px;
          background: linear-gradient(135deg, #28a745, #20c997);
          color: white;
          padding: 8px 12px;
          border-radius: 16px;
          font-size: 12px;
          font-weight: 500;
          animation: pulse 2s infinite;
        }

        .latest-indicator {
          background: rgba(255, 255, 255, 0.2);
          padding: 2px 6px;
          border-radius: 8px;
          font-size: 10px;
          margin-left: 4px;
        }

        @keyframes pulse {
          0% { transform: scale(1); }
          50% { transform: scale(1.05); }
          100% { transform: scale(1); }
        }

        .stats-overview {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
          gap: 15px;
          margin-bottom: 20px;
        }

        .stat-card {
          background: white;
          padding: 20px;
          border-radius: 8px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          display: flex;
          align-items: center;
          gap: 15px;
          transition: transform 0.2s, box-shadow 0.2s;
        }

        .stat-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }

        .stat-icon {
          width: 40px;
          height: 40px;
          color: #667eea;
          flex-shrink: 0;
        }

        .stat-number {
          display: block;
          font-size: 24px;
          font-weight: bold;
          color: #333;
          line-height: 1.2;
        }

        .stat-label {
          display: block;
          font-size: 12px;
          color: #666;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }

        /* Initial State Styles */
        .initial-state {
          background: white;
          border-radius: 12px;
          padding: 40px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
          margin-top: 20px;
        }

        .initial-state-header {
          text-align: center;
          margin-bottom: 40px;
        }

        .initial-state-header svg {
          color: #667eea;
          margin-bottom: 20px;
        }

        .initial-state-header h3 {
          margin: 0 0 16px 0;
          color: #333;
          font-size: 28px;
          font-weight: 600;
        }

        .initial-state-header p {
          margin: 0;
          color: #666;
          font-size: 16px;
          line-height: 1.6;
          max-width: 600px;
          margin: 0 auto;
        }

        .requirements-checklist {
          margin-bottom: 40px;
        }

        .requirements-checklist h4 {
          margin: 0 0 20px 0;
          color: #333;
          font-size: 20px;
          font-weight: 600;
        }

        .checklist-items {
          display: grid;
          gap: 16px;
        }

        .checklist-item {
          display: flex;
          align-items: center;
          gap: 16px;
          padding: 16px;
          border-radius: 8px;
          border: 2px solid;
          transition: all 0.3s ease;
        }

        .checklist-item.completed {
          background: #d4edda;
          border-color: #c3e6cb;
        }

        .checklist-item.pending {
          background: #f8d7da;
          border-color: #f5c6cb;
        }

        .checklist-item.optional {
          background: #d1ecf1;
          border-color: #bee5eb;
        }

        .checklist-icon {
          flex-shrink: 0;
        }

        .checklist-item.completed .checklist-icon {
          color: #28a745;
        }

        .checklist-item.pending .checklist-icon {
          color: #dc3545;
        }

        .checklist-item.optional .checklist-icon {
          color: #17a2b8;
        }

        .checklist-content {
          flex: 1;
        }

        .checklist-title {
          display: block;
          font-weight: 600;
          color: #333;
          margin-bottom: 4px;
        }

        .checklist-status {
          display: block;
          font-size: 14px;
          color: #666;
        }

        .getting-started {
          margin-bottom: 40px;
        }

        .getting-started h4 {
          margin: 0 0 24px 0;
          color: #333;
          font-size: 20px;
          font-weight: 600;
        }

        .steps {
          display: grid;
          gap: 24px;
        }

        .step {
          display: flex;
          align-items: flex-start;
          gap: 16px;
          padding: 20px;
          background: #f8f9fa;
          border-radius: 8px;
          border-left: 4px solid #667eea;
        }

        .step-number {
          display: flex;
          align-items: center;
          justify-content: center;
          width: 32px;
          height: 32px;
          background: #667eea;
          color: white;
          border-radius: 50%;
          font-weight: 600;
          font-size: 16px;
          flex-shrink: 0;
        }

        .step-content h5 {
          margin: 0 0 8px 0;
          color: #333;
          font-size: 16px;
          font-weight: 600;
        }

        .step-content p {
          margin: 0;
          color: #666;
          font-size: 14px;
          line-height: 1.5;
        }

        .action-section {
          text-align: center;
          padding: 30px;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          border-radius: 12px;
          color: white;
        }

        .action-disabled p {
          margin: 0;
          font-size: 16px;
          opacity: 0.9;
        }

        .sub-tab-navigation {
          display: flex;
          gap: 2px;
          margin-bottom: 20px;
          background: white;
          border-radius: 8px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          overflow: hidden;
        }

        .sub-tab {
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 16px 24px;
          background: none;
          border: none;
          border-bottom: 3px solid transparent;
          cursor: pointer;
          font-size: 14px;
          font-weight: 500;
          color: #6c757d;
          transition: all 0.2s ease;
          position: relative;
          flex: 1;
          justify-content: center;
        }

        .sub-tab:hover:not(:disabled) {
          color: #495057;
          background: rgba(102, 126, 234, 0.1);
        }

        .sub-tab.active {
          color: #667eea;
          border-bottom-color: #667eea;
          font-weight: 600;
          background: white;
        }

        .sub-tab:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .data-indicator {
          color: #28a745;
          font-size: 18px;
          line-height: 1;
        }

        .sub-tab-content {
          min-height: 400px;
        }

        .tables-preview-section {
          padding: 0;
        }

        .table-section {
          margin-bottom: 32px;
          border: 1px solid #e0e0e0;
          border-radius: 12px;
          overflow: hidden;
          background: white;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }

        .table-section-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 20px 24px;
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
          border-bottom: 1px solid #e0e0e0;
        }

        .table-section-header h4 {
          margin: 0;
          color: #333;
          font-size: 18px;
          font-weight: 600;
        }

        .table-status {
          display: flex;
          align-items: center;
        }

        .status-badge {
          padding: 6px 12px;
          border-radius: 16px;
          font-size: 12px;
          font-weight: 600;
          border: 1px solid;
        }

        .status-badge.available {
          background: #d4edda;
          color: #155724;
          border-color: #c3e6cb;
        }

        .status-badge.unavailable {
          background: #f8d7da;
          color: #721c24;
          border-color: #f5c6cb;
        }

        .table-empty-state {
          text-align: center;
          padding: 80px 24px;
          color: #666;
        }

        .table-empty-state h4 {
          margin: 24px 0 16px;
          color: #333;
          font-size: 20px;
          font-weight: 600;
        }

        .table-empty-state p {
          color: #666;
          max-width: 500px;
          margin: 0 auto;
          line-height: 1.6;
          font-size: 16px;
        }

        .table-separator {
          height: 24px;
        }

        .tables-summary {
          background: white;
          border: 1px solid #e0e0e0;
          border-radius: 12px;
          padding: 24px;
          margin-top: 32px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }

        .tables-summary h4 {
          margin: 0 0 20px 0;
          color: #333;
          font-size: 20px;
          font-weight: 600;
        }

        .export-merged-section {
          margin-top: 25px;
          padding: 20px;
          background: white;
          border: 2px dashed #667eea;
          border-radius: 8px;
          text-align: center;
        }

        .export-empty-state {
          text-align: center;
          padding: 40px;
          background: #f8f9fa;
          border-radius: 8px;
          border: 2px dashed #dee2e6;
        }

        .export-empty-state p {
          margin: 0 0 20px 0;
          color: #666;
          font-size: 16px;
        }

        .export-actions {
          display: flex;
          gap: 12px;
          justify-content: center;
          flex-wrap: wrap;
        }

        .empty-state {
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          padding: 60px 20px;
          text-align: center;
          color: #6c757d;
        }

        .empty-state svg {
          color: #dee2e6;
          margin-bottom: 20px;
        }

        .empty-state h3 {
          margin: 0 0 10px 0;
          font-size: 18px;
          color: #495057;
        }

        .empty-state p {
          margin: 0 0 20px 0;
          font-size: 14px;
          max-width: 400px;
        }

        .data-preview {
          background: white;
          border-radius: 0;
          overflow: hidden;
        }

        .preview-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 15px 20px;
          background: #f8f9fa;
          border-bottom: 1px solid #dee2e6;
        }

        .preview-header h4 {
          margin: 0;
          font-size: 16px;
          font-weight: 600;
        }

        .table-container {
          overflow-x: auto;
          max-height: 600px;
          overflow-y: auto;
        }

        .data-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 13px;
        }

        .data-table th {
          background: #f8fafc;
          color: #374151;
          padding: 10px 8px;
          text-align: left;
          font-weight: 600;
          position: sticky;
          top: 0;
          z-index: 10;
          border-bottom: 2px solid #e2e8f0;
          border-right: 1px solid #e2e8f0;
        }

        .data-table th.product-header {
          background: #667eea !important;
          color: white !important;
          min-width: 250px;
          position: sticky;
          left: 0;
          z-index: 11;
        }

        .data-table th.data-header {
          text-align: center !important;
          min-width: 100px;
        }

        .data-table td {
          padding: 8px;
          border-bottom: 1px solid #f1f5f9;
          border-right: 1px solid #f1f5f9;
          white-space: nowrap;
        }

        .data-table td.product-cell {
          font-weight: 500;
          color: #1e293b;
          background: #f8fafc;
          position: sticky;
          left: 0;
          z-index: 5;
          border-right: 2px solid #667eea;
          min-width: 250px;
          word-wrap: break-word;
        }

        .data-table td.data-cell {
          font-family: 'Courier New', monospace;
          font-size: 13px;
          text-align: right;
        }

        .data-table td.numeric {
          text-align: right;
        }

        .data-table tr.total-row {
          background: #e2efda !important;
          font-weight: 600;
        }

        .data-table tr.total-row td.product-cell {
          background: #c3e6cb !important;
          color: #155724 !important;
          font-weight: 700;
        }

        .data-table tr:hover {
          background: #f8f9fa;
        }

        .data-table tr:hover td.product-cell {
          background: #e9ecef;
        }

        .data-table tr.total-row:hover {
          background: #d5e8d4 !important;
        }

        .data-table tr.total-row:hover td.product-cell {
          background: #b3d7b8 !important;
        }

        .table-note {
          padding: 10px 20px;
          background: #f8fafc;
          border-top: 1px solid #dee2e6;
          font-size: 12px;
          color: #6c757d;
          text-align: center;
        }

        .processing-indicator {
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 10px;
          padding: 40px;
          color: #666;
          background: white;
          border-radius: 8px;
          margin: 20px 0;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .btn {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 10px 16px;
          border: none;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s ease;
          text-decoration: none;
          font-family: inherit;
        }

        .btn:disabled {
          opacity: 0.6;
          cursor: not-allowed;
          transform: none !important;
        }

        .btn-primary {
          background: #667eea;
          color: white;
          box-shadow: 0 2px 4px rgba(102, 126, 234, 0.3);
        }

        .btn-primary:hover:not(:disabled) {
          background: #5a67d8;
          transform: translateY(-1px);
          box-shadow: 0 4px 8px rgba(102, 126, 234, 0.4);
        }

        .btn-secondary {
          background: #6c757d;
          color: white;
          box-shadow: 0 2px 4px rgba(108, 117, 125, 0.3);
        }

        .btn-secondary:hover:not(:disabled) {
          background: #545b62;
          transform: translateY(-1px);
          box-shadow: 0 4px 8px rgba(108, 117, 125, 0.4);
        }

        .btn-large {
          padding: 14px 28px;
          font-size: 16px;
          font-weight: 600;
        }

        @media (max-width: 1024px) {
          .product-analysis-section {
            padding: 16px;
          }

          .export-actions {
            flex-direction: column;
            align-items: stretch;
          }

          .steps {
            grid-template-columns: 1fr;
          }
        }

        @media (max-width: 768px) {
          .product-analysis-section {
            padding: 12px;
          }

          .section-header {
            flex-direction: column;
            gap: 16px;
            align-items: stretch;
            text-align: center;
          }

          .stats-overview {
            grid-template-columns: 1fr;
          }

          .sub-tab-navigation {
            flex-direction: column;
          }

          .sub-tab {
            justify-content: flex-start;
            padding: 14px 20px;
          }

          .data-table {
            font-size: 12px;
          }

          .data-table th,
          .data-table td {
            padding: 8px 10px;
          }

          .data-table th.product-header,
          .data-table td.product-cell {
            min-width: 200px;
          }

          .initial-state {
            padding: 24px;
          }

          .initial-state-header h3 {
            font-size: 24px;
          }

          .step {
            flex-direction: column;
            text-align: center;
          }

          .checklist-items {
            grid-template-columns: 1fr;
          }
        }

        @media (max-width: 480px) {
          .product-analysis-section {
            padding: 8px;
          }

          .section-header {
            padding: 16px;
          }

          .section-header h2 {
            font-size: 20px;
          }

          .stat-card {
            padding: 16px;
          }

          .stat-icon {
            width: 32px;
            height: 32px;
          }

          .stat-number {
            font-size: 20px;
          }

          .sub-tab {
            padding: 12px 16px;
            font-size: 13px;
          }

          .data-table {
            font-size: 11px;
          }

          .data-table th,
          .data-table td {
            padding: 6px 8px;
          }

          .btn {
            padding: 8px 12px;
            font-size: 13px;
          }

          .initial-state {
            padding: 20px;
          }

          .initial-state-header svg {
            width: 48px;
            height: 48px;
          }

          .initial-state-header h3 {
            font-size: 20px;
          }

          .step-number {
            width: 28px;
            height: 28px;
            font-size: 14px;
          }

          .checklist-item {
            padding: 12px;
          }

          .checklist-icon svg {
            width: 16px;
            height: 16px;
          }
        }
      `}</style>
    </div>
  );
};

export default ProductAnalysis;
