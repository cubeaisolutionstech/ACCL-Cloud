import React, { useState, useCallback } from 'react';
import { 
  Download, 
  FileSpreadsheet,
  Database,
  Layers,
  BarChart3,
  RefreshCw,
  MapPin,
  TrendingUp,
  Package,
  Merge,
  Activity
} from 'lucide-react';

const CombinedExcelManager = ({ 
  regionData, 
  fiscalInfo, 
  addMessage, 
  loading, 
  setLoading,
  storedFiles = [],
  setStoredFiles = () => {}
}) => {
  // State Management
  const [processing, setProcessing] = useState(false);

  const API_BASE_URL = process.env.REACT_APP_API_URL || 'http://localhost:5000/api';

  // Region Sheet Column Definitions
  const regionMtColumns = [
    { header: "Region", key: "region", type: "text" },
    { header: "Product", key: "product", type: "text" },
    { header: "Month", key: "month", type: "text" },
    { header: "Fiscal Year", key: "fiscalYear", type: "text" },
    { header: "Quantity (MT)", key: "quantity", type: "number", format: "0.00" },
    { header: "Growth %", key: "growthPercentage", type: "number", format: "0.00" },
    { header: "Market Share %", key: "marketShare", type: "number", format: "0.00" },
    { header: "Avg Price", key: "avgPrice", type: "number", format: "0.00" },
    { header: "Previous Year Qty", key: "prevYearQty", type: "number", format: "0.00" },
    { header: "YTD Qty", key: "ytdQty", type: "number", format: "0.00" },
    { header: "YTD Growth %", key: "ytdGrowth", type: "number", format: "0.00" },
    { header: "Notes", key: "notes", type: "text" }
  ];

  const regionValueColumns = [
    { header: "Region", key: "region", type: "text" },
    { header: "Product", key: "product", type: "text" },
    { header: "Month", key: "month", type: "text" },
    { header: "Fiscal Year", key: "fiscalYear", type: "text" },
    { header: "Value (₹)", key: "value", type: "number", format: "0.00" },
    { header: "Growth %", key: "growthPercentage", type: "number", format: "0.00" },
    { header: "Market Share %", key: "marketShare", type: "number", format: "0.00" },
    { header: "Avg Price", key: "avgPrice", type: "number", format: "0.00" },
    { header: "Previous Year Value", key: "prevYearValue", type: "number", format: "0.00" },
    { header: "YTD Value", key: "ytdValue", type: "number", format: "0.00" },
    { header: "YTD Growth %", key: "ytdGrowth", type: "number", format: "0.00" },
    { header: "Notes", key: "notes", type: "text" }
  ];

  // Enhanced file categorization function
  const categorizeFiles = useCallback((files) => {
    const categories = {
      sales: [],
      region: [],
      product: [],
      ts_pw: [],
      ero_pw: [],
      combined: []
    };

    files.forEach(file => {
      const fileType = file.type?.toLowerCase() || '';
      const fileName = file.name?.toLowerCase() || '';
      const source = file.source?.toLowerCase() || '';

      if (fileType.includes('sales') || fileName.includes('sales') || source.includes('sales') || fileType === 'sales-analysis-excel') {
        categories.sales.push(file);
      } else if (fileType.includes('region') || fileName.includes('region') || source.includes('region')) {
        categories.region.push(file);
      } else if (fileType.includes('product') || fileName.includes('product') || source.includes('product') || file.metadata?.analysisType === 'product') {
        categories.product.push(file);
      } else if (fileType.includes('tspw') || fileType.includes('ts-pw') || fileName.includes('ts')) {
        categories.ts_pw.push(file);
      } else if (fileType.includes('eropw') || fileType.includes('ero-pw') || fileName.includes('ero')) {
        categories.ero_pw.push(file);
      } else if (['combined', 'master-combined', 'selected-combined', 'master-combined-enhanced', 'selected-combined-enhanced', 'category-based-combined'].includes(fileType)) {
        categories.combined.push(file);
      }
    });

    return categories;
  }, []);

  // Utility functions
  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  // Computed values
  const fileGroups = categorizeFiles(storedFiles);

  const categorySummary = Object.entries(fileGroups).map(([category, files]) => ({
    name: category,
    count: files.length,
    icon: {
      sales: <TrendingUp size={24} />,
      region: <MapPin size={24} />,
      product: <Package size={24} />,
      ts_pw: <Activity size={24} />,
      ero_pw: <BarChart3 size={24} />,
      combined: <Merge size={24} />
    }[category] || <FileSpreadsheet size={24} />,
    color: {
      sales: '#28a745',
      region: '#007bff',
      product: '#dc3545',
      ts_pw: '#6c757d',
      ero_pw: '#17a2b8',
      combined: '#ffc107'
    }[category] || '#6c757d'
  })).filter(cat => cat.count > 0);

  // Get overall statistics
  const getOverallStats = useCallback(() => {
    const totalFiles = storedFiles.length;
    const totalSize = storedFiles.reduce((sum, file) => sum + (file.size || 0), 0);
    const totalSheets = storedFiles.reduce((sum, file) => 
      sum + (file.sheets?.length || 1), 0
    );
    const totalRecords = storedFiles.reduce((sum, file) => 
      sum + (file.mtRecords || 0) + (file.valueRecords || 0) + 
      (file.metadata?.mtRows || 0) + (file.metadata?.valueRows || 0), 0
    );

    return { totalFiles, totalSize, totalSheets, totalRecords };
  }, [storedFiles]);

  // Enhanced Combine all Excel files into one master file
  const combineAndDownloadExcel = useCallback(async () => {
    setProcessing(true);
    try {
      // Get all category files
      const categories = categorizeFiles(storedFiles);
      const filesToCombine = [];
      
      // Prepare files for combining
      for (const [category, files] of Object.entries(categories)) {
        if (files.length > 0) {
          const mostRecentFile = files.reduce((prev, current) => 
            (new Date(prev.createdAt) > new Date(current.createdAt)) ? prev : current
          );
          
          if (mostRecentFile.blob) {
            const arrayBuffer = await mostRecentFile.blob.arrayBuffer();
            const base64String = btoa(
              new Uint8Array(arrayBuffer).reduce(
                (data, byte) => data + String.fromCharCode(byte), ''
              )
            );
            
            filesToCombine.push({
              name: category,
              content: base64String,
              metadata: {
                ...mostRecentFile.metadata,
                decimalPlaces: mostRecentFile.metadata?.decimalPlaces || 2,
                category: category,
                sourceFile: mostRecentFile.name,
                centerTitles: true
              }
            });
          }
        }
      }

      if (filesToCombine.length === 0) {
        throw new Error('No files available to combine');
      }

      addMessage(`🔄 Combining all category files with centered titles and comprehensive formatting...`, 'info');
      
      const response = await fetch(`${API_BASE_URL}/combined/combine-excel-files`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          files: filesToCombine,
          region_mt_columns: regionMtColumns,
          region_value_columns: regionValueColumns,
          excel_formatting: {
            number_format: {
              decimal_places: 2,
              apply_to_all_numbers: true,
              format_pattern: '0.00',
              force_format: true,
              apply_to_sheets: 'all'
            },
            title_formatting: {
              center_all_titles: true,
              title_font_size: 16,
              title_font_bold: true,
              title_background_color: '#E3F2FD',
              title_font_color: '#1565C0',
              merge_title_cells: true,
              add_title_borders: true,
              title_row_height: 30,
              apply_to_all_sheets: true
            }
          }
        })
      });

      const result = await response.json();
      
      if (result.success) {
        // Convert base64 back to blob
        const byteCharacters = atob(result.file_data);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
        
        // Download combined file
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = result.file_name || `combined_auditor_format_${new Date().toISOString().slice(0,10)}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        // Add to stored files
        const newFile = {
          id: Date.now(),
          name: result.file_name,
          blob: blob,
          size: blob.size,
          createdAt: new Date().toISOString(),
          type: 'master-combined-auditor-format',
          source: 'Master Combined Auditor Format',
          url: URL.createObjectURL(blob),
          metadata: {
            ...result.metadata,
            totalSourceFiles: filesToCombine.length,
            categoriesIncluded: Object.keys(categories).filter(key => categories[key].length > 0),
            generationType: 'auditor_format'
          }
        };
        
        setStoredFiles(prev => [newFile, ...prev]);
        addMessage('✅ Successfully created Auditor Format Excel file', 'success');
        addMessage(`📊 Combined ${filesToCombine.length} category files with professional formatting`, 'success');
        
      } else {
        throw new Error(result.error || 'Failed to combine files');
      }
    } catch (error) {
      addMessage(`❌ Error creating Auditor Format file: ${error.message}`, 'error');
    } finally {
      setProcessing(false);
    }
  }, [storedFiles, addMessage, setStoredFiles, API_BASE_URL, categorizeFiles, regionMtColumns, regionValueColumns]);

  const stats = getOverallStats();

  return (
    <div className="simplified-excel-manager">
      {/* Header Section */}
      <div className="manager-header">
        <h3> Auditor Format Generator</h3>
        <p>Combine all analysis files into one professional Auditor Format Excel file</p>
      </div>



      {/* Category Overview */}
      {categorySummary.length > 0 && (
        <div className="category-overview">
          <h4> Available Categories</h4>
          <div className="category-grid">
            {categorySummary.map(category => (
              <div key={category.name} className="category-card" style={{borderLeftColor: category.color}}>
                <div className="category-icon" style={{color: category.color}}>
                  {category.icon}
                </div>
                <div className="category-content">
                  <span className="category-name">{category.name.replace('_', ' ').toUpperCase()}</span>
                  <span className="category-count">{category.count} files</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Main Action */}
      <div className="main-action">
        <button
          onClick={combineAndDownloadExcel}
          className="btn btn-primary btn-large"
          disabled={loading || processing || storedFiles.length === 0}
        >
          {processing ? <RefreshCw size={20} className="spinning" /> : <Merge size={20} />}
          {processing ? 'Creating Auditor Format...' : 'Generate Auditor Format Excel'}
        </button>
        
      </div>

      {/* Empty State */}
      {storedFiles.length === 0 && (
        <div className="empty-state">
          <FileSpreadsheet size={48} />
          <h4>No Files Available</h4>
          <p>Generate files from analysis modules to create an Auditor Format Excel file</p>
        </div>
      )}

      {/* Styles */}
      <style jsx>{`
        .simplified-excel-manager {
          background: white;
          border-radius: 12px;
          padding: 24px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.1);
          border: 1px solid #e0e0e0;
        }

        .manager-header {
          text-align: center;
          margin-bottom: 32px;
          padding-bottom: 20px;
          border-bottom: 2px solid #f0f0f0;
        }

        .manager-header h3 {
          margin: 0 0 8px 0;
          color: #333;
          font-size: 24px;
          font-weight: 700;
        }

        .manager-header p {
          margin: 0;
          color: #666;
          font-size: 16px;
          line-height: 1.5;
        }



        .category-overview {
          background: linear-gradient(135deg, #f0f8ff 0%, #e6f3ff 100%);
          border-radius: 12px;
          padding: 24px;
          margin-bottom: 32px;
          border: 1px solid #2196f3;
        }

        .category-overview h4 {
          margin: 0 0 20px 0;
          color: #1565c0;
          font-size: 18px;
          font-weight: 600;
        }

        .category-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
          gap: 16px;
        }

        .category-card {
          background: white;
          border-radius: 8px;
          padding: 16px;
          display: flex;
          align-items: center;
          gap: 12px;
          border: 1px solid #e9ecef;
          border-left: 4px solid #007bff;
          transition: all 0.2s ease;
        }

        .category-card:hover {
          box-shadow: 0 4px 8px rgba(0,0,0,0.1);
          transform: translateY(-1px);
        }

        .category-icon {
          padding: 8px;
          border-radius: 8px;
          background: rgba(0,123,255,0.1);
        }

        .category-content {
          flex: 1;
        }

        .category-name {
          display: block;
          font-weight: 600;
          color: #333;
          font-size: 14px;
          margin-bottom: 2px;
        }

        .category-count {
          display: block;
          color: #666;
          font-size: 12px;
        }

        .main-action {
          text-align: center;
          padding: 32px;
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
          border-radius: 12px;
          border: 2px dashed #007bff;
          margin-bottom: 32px;
        }

        .btn {
          display: inline-flex;
          align-items: center;
          gap: 12px;
          padding: 16px 32px;
          border: none;
          border-radius: 8px;
          font-size: 16px;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.3s ease;
          text-decoration: none;
          position: relative;
          overflow: hidden;
        }

        .btn:disabled {
          opacity: 0.6;
          cursor: not-allowed;
          transform: none !important;
        }

        .btn-primary {
          background: linear-gradient(135deg, #007bff, #0056b3);
          color: white;
          box-shadow: 0 4px 12px rgba(0,123,255,0.3);
        }

        .btn-primary:hover:not(:disabled) {
          transform: translateY(-2px);
          box-shadow: 0 6px 20px rgba(0,123,255,0.4);
        }

        .btn-large {
          padding: 18px 36px;
          font-size: 18px;
          font-weight: 700;
        }

        .action-description {
          margin: 16px 0 0;
          font-size: 14px;
          color: #666;
          line-height: 1.6;
        }

        .empty-state {
          text-align: center;
          padding: 60px 20px;
          color: #666;
          border: 2px dashed #dee2e6;
          border-radius: 12px;
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        }

        .empty-state svg {
          color: #dee2e6;
          margin-bottom: 20px;
        }

        .empty-state h4 {
          margin: 0 0 12px 0;
          color: #495057;
          font-size: 20px;
          font-weight: 600;
        }

        .empty-state p {
          margin: 0;
          font-size: 16px;
          line-height: 1.5;
        }

        .spinning {
          animation: spin 1s linear infinite;
        }

        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }

        @media (max-width: 768px) {
          .simplified-excel-manager {
            padding: 16px;
          }

          .category-grid {
            grid-template-columns: 1fr;
          }

          .main-action {
            padding: 24px;
          }

          .btn-large {
            padding: 14px 28px;
            font-size: 16px;
          }

          .manager-header h3 {
            font-size: 20px;
          }

          .manager-header p {
            font-size: 14px;
          }
        }

        @media (max-width: 480px) {
          .main-action {
            padding: 20px;
          }

          .action-description {
            font-size: 13px;
          }
        }
      `}</style>
    </div>
  );
};

export default CombinedExcelManager;
