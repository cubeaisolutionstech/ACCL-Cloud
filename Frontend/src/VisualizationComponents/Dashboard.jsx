import React, { useState, useEffect, useRef, useMemo } from 'react';
import axios from 'axios';
import Plot from 'react-plotly.js';
import * as XLSX from 'xlsx';
import './App.css';

// Constants
const BRANCH_EXCLUDE_TERMS = ['CHN Total', 'ERD SALES', 'WEST SALES', 'GROUP COMPANIES',
  'GRAND TOTAL', 'NORTH TOTAL', 'SOUTH TOTAL', 'EAST TOTAL', 'WEST TOTAL',
  'TOTAL', 'SUMMARY', 'REGION TOTAL'];
const PRODUCT_EXCLUDE_TERMS = ['PRODUCT WISE SALES - VALUE DATA', 'PRODUCT WISE SALES', 'SUMMARY', 'TOTAL', 'GRAND TOTAL', 'OVERALL TOTAL',
  'TOTAL SALES'];
const MONTH_ORDER = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar'];

// API service functions
const api = {
  testConnection: async () => {
    try {
      const response = await axios.get("http://localhost:5003/test", {
        timeout: 5000,
      });
      return response.data;
    } catch (error) {
      console.error("Connection test failed:", error);
      throw error;
    }
  },
  uploadFile: async (file) => {
    try {
      const formData = new FormData();
      formData.append("file", file);
      const response = await axios.post("http://localhost:5003/upload", formData, {
        headers: {
          "Content-Type": "multipart/form-data",
        },
        timeout: 60000,
      });
      return response.data;
    } catch (error) {
      console.error("Upload failed:", error);
      if (error.code === "ECONNABORTED") {
        throw new Error("Upload timeout - file may be too large");
      } else if (error.response) {
        throw new Error(error.response.data.error || "Server error");
      } else if (error.request) {
        throw new Error("Cannot connect to server. Make sure Flask backend is running on http://localhost:5003");
      } else {
        throw new Error("Upload failed: " + error.message);
      }
    }
  },
  processSheet: async (data) => {
    try {
      const response = await axios.post("http://localhost:5003/process-sheet", data, {
        timeout: 60000,
      });
      return response.data;
    } catch (error) {
      console.error("Process sheet failed:", error);
      if (error.response) {
        throw new Error(error.response.data.error || "Server error");
      } else if (error.request) {
        throw new Error("Cannot connect to server");
      } else {
        throw new Error("Processing failed: " + error.message);
      }
    }
  },
  generateVisualizations: async (data) => {
    try {
      const response = await axios.post("http://localhost:5003/visualizations", data, {
        timeout: 60000,
      });
      return response.data;
    } catch (error) {
      console.error("Visualization generation failed:", error);
      if (error.response) {
        throw new Error(error.response.data.error || "Server error");
      } else if (error.request) {
        throw new Error("Cannot connect to server");
      } else {
        throw new Error("Visualization failed: " + error.message);
      }
    }
  },
  downloadCSV: async (data) => {
    try {
      const response = await axios.post("http://localhost:5003/download-csv", data, {
        responseType: 'blob',
        headers: {
          'Content-Type': 'application/json',
        },
        timeout: 30000,
      });
      
      if (response.status >= 400) {
        const errorData = await response.data.text();
        throw new Error(errorData || 'CSV generation failed');
      }
      
      return response;
    } catch (error) {
      console.error("CSV download failed:", {
        error: error,
        response: error.response?.data,
        stack: error.stack
      });
      
      if (error.response?.data instanceof Blob) {
        try {
          const errorText = await error.response.data.text();
          error.message = errorText || error.message;
        } catch (e) {
          console.error("Couldn't parse error blob:", e);
        }
      }
      
      throw error;
    }
  },
  generatePPT: async (data) => {
    try {
      const response = await axios.post("http://localhost:5003/download-ppt", data, {
        responseType: 'blob',
        headers: {
          'Content-Type': 'application/json'
        },
        timeout: 60000,
      });
      return response.data;
    } catch (error) {
      console.error("PPT generation failed:", error);
      if (error.code === "ECONNABORTED") {
        throw new Error("PPT generation timeout - process took too long");
      } else if (error.response) {
        throw new Error(error.response.data.error || "Server error");
      } else if (error.request) {
        throw new Error("Cannot connect to server. Make sure Flask backend is running on http://localhost:5003");
      } else {
        throw new Error("PPT generation failed: " + error.message);
      }
    }
  },
  generateMasterPPT: async (data) => {
    try {
      const response = await axios.post("http://localhost:5003/download-master-ppt", data, {
        responseType: 'blob',
        headers: {
          'Content-Type': 'application/json'
        },
        timeout: 120000,
      });
      return response.data;
    } catch (error) {
      console.error("Master PPT generation failed:", error);
      if (error.code === "ECONNABORTED") {
        throw new Error("Master PPT generation timeout - process took too long");
      } else if (error.response) {
        throw new Error(error.response.data.error || "Server error");
      } else if (error.request) {
        throw new Error("Cannot connect to server. Make sure Flask backend is running on http://localhost:5003");
      } else {
        throw new Error("Master PPT generation failed: " + error.message);
      }
    }
  }
};

function Dashboard() {
    // State management
      const [uploadedFile, setUploadedFile] = useState(null);
      const [fileData, setFileData] = useState(null);
      const [sheetNames, setSheetNames] = useState([]);
      const [selectedSheet, setSelectedSheet] = useState('');
      const [tableData, setTableData] = useState(null);
      const [tableName, setTableName] = useState('');
      const [visualType, setVisualType] = useState('Bar Chart');
      const [selectedMonth, setSelectedMonth] = useState('Select All');
      const [selectedYear, setSelectedYear] = useState('Select All');
      const [selectedBranch, setSelectedBranch] = useState('Select All');
      const [selectedProduct, setSelectedProduct] = useState('Select All');
      const [isLoading, setIsLoading] = useState(false);
      const [isProcessing, setIsProcessing] = useState(false);
      const [processingProgress, setProcessingProgress] = useState(0);
      const [error, setError] = useState(null);
      const [activeTab, setActiveTab] = useState('Budget vs Actual');
      const [months, setMonths] = useState(['Select All']);
      const [years, setYears] = useState(['Select All']);
      const [branches, setBranches] = useState(['Select All']);
      const [products, setProducts] = useState(['Select All']);
      const [isFirstSheet, setIsFirstSheet] = useState(false);
      const [isBranchAnalysis, setIsBranchAnalysis] = useState(false);
      const [isProductAnalysis, setIsProductAnalysis] = useState(false);
      const [tableOptions, setTableOptions] = useState([]);
      const [tableChoice, setTableChoice] = useState('');
      const [connectionStatus, setConnectionStatus] = useState('checking');
      const [isTableLoading, setIsTableLoading] = useState(false);
      const [tableAvailability, setTableAvailability] = useState({});
      const [sheetType, setSheetType] = useState('');
    
      // Visualization data
      const [visualizationData, setVisualizationData] = useState({
        'Budget vs Actual': null,
        'Budget': null,
        'LY': null,
        'Act': null,
        'Gr': null,
        'Ach': null,
        'YTD Budget': null,
        'YTD LY': null,
        'YTD Act': null,
        'YTD Gr': null,
        'YTD Ach': null,
        'Branch Performance': null,
        'Branch Monthwise': null,
        'Product Performance': null,
        'Product Monthwise': null
      });
    
      const fileInputRef = useRef(null);
    
      // Helper to detect if current sheet is Sales Analysis Month-wise
      const isSalesAnalysisMonthwise = useMemo(() => {
        const sheetLower = selectedSheet?.toLowerCase().trim();
        return [
          'sales analysis month wise',
          'sales analysis month-wise',
          'month-wise sales'
        ].includes(sheetLower);
      }, [selectedSheet]);
    
      const isRegionWiseAnalysis = useMemo(() => {
        return selectedSheet?.toLowerCase().includes('region wise analysis');
      }, [selectedSheet]);
    
      // Clean column headers based on sheet type
      const cleanColumnHeader = (header) => {
      if (isSalesAnalysisMonthwise) return header;
      
      // For YTD columns, show the full original name
      if (typeof header === 'string' && 
          (header.includes('YTD') || header.includes('Budget') || header.includes('LY'))) {
        return header;
      }
        if (isProductAnalysis) {
        return String(header).replace(/\s+/g, ' ').trim(); // Just normalize spaces
      }
      
      return header;
    };
    
    // UPDATED extractMonthYear function - this removes LY prefix from x-axis labels
    const extractMonthYear = (colName) => {
      const colStr = String(colName).trim();
      
      // Handle LY columns - extract month-year without LY prefix
      if (colStr.startsWith('LY-')) {
        const lyMatch = colStr.match(/LY-(\w{3})-(\d{2})/i);
        if (lyMatch) {
          return `${lyMatch[1]}-${lyMatch[2]}`; // Return just "Apr-24" instead of "LY-Apr-24"
        }
      }
      
      // Handle all YTD column formats
      const ytdMatch = colStr.match(/(Budget|LY|Act|Gr|Ach)-YTD-(\d{2})-(\d{2})\((Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar) to (Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar)\)/i);
      if (ytdMatch) {
        return `${ytdMatch[1]}-YTD-${ytdMatch[2]}-${ytdMatch[3]}(${ytdMatch[4]}-${ytdMatch[5]})`;
      }
      
      // Original month-year extraction for other columns
      const monthMatch = colStr.match(/(apr|may|jun|jul|aug|sep|oct|nov|dec|jan|feb|mar)/i);
      const yearMatch = colStr.match(/(\d{2,4})/);
      
      if (monthMatch && yearMatch) {
        return `${monthMatch[0].charAt(0).toUpperCase() + monthMatch[0].slice(1).toLowerCase()}-${yearMatch[0].slice(-2)}`;
      }
      return colStr;
    };
    
      const safeConvertValue = (x) => {
        if (x === null || x === undefined || x === '' || String(x).toLowerCase() === 'nan') {
          return null;
        }
        return String(x).trim();
      };
    
      const makeJsonlySerializable = (data) => {
        if (!data || data.length === 0) return data;
        return data.map(row => {
          const newRow = {};
          Object.keys(row).forEach(key => {
            newRow[key] = safeConvertValue(row[key]);
          });
          return newRow;
        });
      };
    
      const columnFilter = (col) => {
        const colStr = String(col).toLowerCase().replace(/[,–]/g, '-');
        
        const monthOk = !selectedMonth || selectedMonth === "Select All" || 
                       colStr.includes(selectedMonth.toLowerCase().slice(0, 3));
        const yearOk = !selectedYear || selectedYear === "Select All" || 
                       colStr.includes(String(selectedYear).slice(-2));
    
        return monthOk && yearOk;
      };
    
      // File upload handler
      const handleFileUpload = async (event) => {
      const file = event.target.files[0];
      if (!file) return;
    
      setIsLoading(true);
      setIsProcessing(true);
      setProcessingProgress(0);
      setUploadedFile(file);
      setError(null);
      setTableData(null);
      setTableOptions([]);
      setTableChoice('');
      setTableAvailability({});
      setSheetType('');
    
      try {
        setProcessingProgress(20);
        const result = await api.uploadFile(file);
        setProcessingProgress(60);
    
        if (!result.sheet_names || result.sheet_names.length === 0) {
          throw new Error("No sheets found in the uploaded file.");
        }
        if (!result.preview_data || result.preview_data.length === 0) {
          throw new Error("Preview data is empty. The sheet may not contain valid table data.");
        }
    
        setSheetNames(result.sheet_names);
        setSelectedSheet(result.sheet_names[0] || '');
    
        const sheetName = result.sheet_names[0].toLowerCase();
        let detectedSheetType = 'Unknown';
        if (sheetName.includes('sales analysis month-wise') || 
            sheetName.includes('sales analysis monthwise') || 
            sheetName.includes('month-wise sales')) {
          detectedSheetType = 'Sales Analysis Month-wise';
        } else if (sheetName.includes('region wise analysis')) {
          detectedSheetType = 'Region-wise Analysis';
        } else if (sheetName.includes('product') || 
                   sheetName.includes('ts-pw') || 
                   sheetName.includes('ero-pw')) {
          detectedSheetType = 'Product Analysis';
        }
        setSheetType(detectedSheetType);
    
        // Initialize FileReader
        const reader = new FileReader();
        reader.onload = (e) => {
          const arrayBuffer = e.target.result;
          const bytes = new Uint8Array(arrayBuffer);
          let binary = '';
          for (let i = 0; i < bytes.byteLength; i++) {
            binary += String.fromCharCode(bytes[i]);
          }
          const base64String = btoa(binary);
          setFileData(base64String);
          setProcessingProgress(100);
        };
        reader.readAsArrayBuffer(file);
      } catch (err) {
        setError(err.message);
      } finally {
        setIsLoading(false);
        setIsProcessing(false);
      }
    };
    
      // Process sheet data
      useEffect(() => {
        if (!fileData || !selectedSheet) return;
    
        const processSheet = async () => {
          setIsLoading(true);
          setIsProcessing(true);
          setIsTableLoading(true);
          setProcessingProgress(0);
          setError(null);
          setSelectedMonth('Select All');
          setSelectedYear('Select All');
          setSelectedBranch('Select All');
          setSelectedProduct('Select All');
    
          const sheetLower = selectedSheet.toLowerCase();
          let detectedSheetType = 'Unknown';
          if (sheetLower.includes('sales analysis month-wise') || 
              sheetLower.includes('sales analysis monthwise') || 
              sheetLower.includes('month-wise sales')) {
            detectedSheetType = 'Sales Analysis Month-wise';
          } else if (sheetLower.includes('region wise analysis')) {
            detectedSheetType = 'Region-wise Analysis';
          } else if (sheetLower.includes('product') || 
                     sheetLower.includes('ts-pw') || 
                     sheetLower.includes('ero-pw')) {
            detectedSheetType = 'Product Analysis';
          }
          setSheetType(detectedSheetType);
    
          try {
            setProcessingProgress(20);
            const result = await api.processSheet({
              file_data: fileData,
              sheet_name: selectedSheet,
              table_choice: tableChoice || undefined,
            });
    
            setProcessingProgress(60);
            setTableData(result);
            setTableName(result.csv_filename || selectedSheet);
    
            const firstSheet = sheetNames.indexOf(selectedSheet) === 0;
            setIsFirstSheet(firstSheet);
            setIsBranchAnalysis(sheetLower.includes('region wise analysis'));
            setIsProductAnalysis(
              sheetLower.includes('product') || 
              sheetLower.includes('ts-pw') || 
              sheetLower.includes('ero-pw')
            );
    
            const availableTables = result.tables || [];
            setTableAvailability(prev => ({
              ...prev,
              [selectedSheet]: availableTables
            }));
    
            if (availableTables.length > 0) {
              setTableOptions(availableTables);
              if (!tableChoice || !availableTables.includes(tableChoice)) {
                setTableChoice(availableTables[0] || '');
              }
            } else {
              setTableOptions([]);
              setTableChoice('');
            }
    
            if (result.months) setMonths(['Select All', ...result.months]);
            if (result.years) setYears(['Select All', ...result.years]);
            if (result.branches) setBranches(['Select All', ...result.branches]);
            if (result.products) setProducts(['Select All', ...result.products]);
    
            setProcessingProgress(90);
          } catch (err) {
            console.error('Process Sheet Error:', err);
            setError(err.message);
            if (err.message.includes('Selected table') && tableAvailability[selectedSheet]?.length > 0) {
              setTableChoice(tableAvailability[selectedSheet][0] || '');
            } else {
              setTableOptions([]);
              setTableChoice('');
            }
          } finally {
            setIsLoading(false);
            setIsProcessing(false);
            setIsTableLoading(false);
            setProcessingProgress(100);
          }
        };
    
        processSheet();
      }, [fileData, selectedSheet, tableChoice]);
    
      const processedData = useMemo(() => {
        if (!tableData || !tableData.data) return [];
    
        return tableData.data.map((row, index) => ({
          ...row,
          isTotalRow: isBranchAnalysis
            ? String(row[tableData.columns[0]]).toUpperCase().includes('GRAND TOTAL')
            : String(row[tableData.columns[0]]).toUpperCase().includes('TOTAL SALES')
        })).filter(row => 
          (!isBranchAnalysis || selectedBranch === 'Select All' || String(row[tableData.columns[0]]).trim() === selectedBranch) &&
          (!isProductAnalysis || selectedProduct === 'Select All' || String(row[tableData.columns[0]]).trim() === selectedProduct)
        );
      }, [tableData, selectedBranch, selectedProduct, isBranchAnalysis, isProductAnalysis]);
    
      const generateBudgetVsActualData = () => {
        if (!tableData || !tableData.columns) return null;
    
        const budgetCols = tableData.columns.filter(col => 
          col.toLowerCase().startsWith('budget') && !col.toLowerCase().includes('ytd') && columnFilter(col)
        );
        const actCols = tableData.columns.filter(col => 
          col.toLowerCase().startsWith('act') && !col.toLowerCase().includes('ytd') && columnFilter(col)
        );
    
        if (!budgetCols.length || !actCols.length) return null;
        
        const commonMonths = [];
        budgetCols.forEach(bCol => {
          const monthYear = extractMonthYear(bCol);
          if (actCols.some(aCol => extractMonthYear(aCol) === monthYear)) {
            commonMonths.push(monthYear);
          }
        });
    
        if (!commonMonths.length) return null;
    
        const filteredData = processedData;
    
        const budgetValues = [];
        const actValues = [];
        commonMonths.forEach(monthYear => {
          const budgetCol = budgetCols.find(col => extractMonthYear(col) === monthYear);
          const actCol = actCols.find(col => extractMonthYear(col) === monthYear);
          if (!budgetCol || !actCol) return;
    
          const budgetTotal = filteredData.reduce((sum, row) => sum + (parseFloat(String(row[budgetCol]).replace(',', '')) || 0), 0);
          const actTotal = filteredData.reduce((sum, row) => sum + (parseFloat(String(row[actCol]).replace(',', '')) || 0), 0);
    
          budgetValues.push(budgetTotal);
          actValues.push(actTotal);
        });
        // Handle pie chart case differently
      if (visualType.toLowerCase().includes('pie')) {
        // For pie chart, combine budget and actual into one pie
        const pieData = [
          {
            labels: ['Budget', 'Actual'],
            values: [
              budgetValues.reduce((a, b) => a + b, 0),
              actValues.reduce((a, b) => a + b, 0)
            ],
            type: 'pie',
            marker: { colors: ['#2E86AB', '#FF8C00'] },
            textinfo: 'percent+value',
            hoverinfo: 'label+percent+value',
            hole: 0,
            showlegend: true
          }
        ];
    
        return {
          data: pieData,
          layout: {
            title: { text: `Budget vs Actual (Total)`, x: 0.5, xanchor: 'center', font: { size: 16 } },
            height: 600,
            margin: { l: 20, r: 20, t: 80, b: 20 },
            annotations: [] // Remove center annotation (optional)
          }
        };
      }
    
        const chartData = [
          {
            x: commonMonths,
            y: budgetValues,
            name: 'Budget',
            type: visualType.toLowerCase().includes('bar') ? 'bar' : 
                  visualType.toLowerCase().includes('line') ? 'scatter' : 'pie',
            marker: { color: '#2E86AB' },
            hovertemplate: '<b>%{x}</b><br>Budget: %{y:,.0f}<extra></extra>',
            width: 0.4
            
          },
          {
            x: commonMonths,
            y: actValues,
            name: 'Act',
            type: visualType.toLowerCase().includes('bar') ? 'bar' : 
                  visualType.toLowerCase().includes('line') ? 'scatter' : 'pie',
            marker: { color: '#FF8C00' },
            hovertemplate: '<b>%{x}</b><br>Actual: %{y:,.0f}<extra></extra>',
            width: 0.4,
            mode: visualType.toLowerCase().includes('line') ? 'lines+markers' : undefined
          }
        ];
    
        chartData.sort((a, b) => {
          const getFiscalSortKey = month => {
            const [m, y] = month.split('-');
            const monthIdx = MONTH_ORDER.indexOf(m);
            const yearInt = parseInt(y);
            return (monthIdx >= 9 ? yearInt - 1 : yearInt) * 100 + monthIdx;
          };
          return getFiscalSortKey(a.x[0]) - getFiscalSortKey(b.x[0]);
        });
    
        return {
          data: chartData,
          layout: {
            title: { text: `Budget vs Actual`, x: 0.5, xanchor: 'center', font: { size: 16 } },
            xaxis: { 
              title: 'Month', 
              titlefont: { size: 14 }, 
              tickfont: { size: 12 }, 
              categoryorder: 'array', 
              categoryarray: commonMonths,
              tickangle: 0 ,
              automargin: true 
            },
            yaxis: { title: 'Value', titlefont: { size: 14 }, tickfont: { size: 12 } },
            barmode: 'group',
            bargap: 0.2,
            plot_bgcolor: 'white',
            paper_bgcolor: 'white',
            height: 600,
            margin: { l: 70, r: 70, t: 90, b: 70 },
            showlegend: true,
            hovermode: 'x unified',
            
          }
        };
      };
    
      const generateDefaultData = (label) => {
        if (!tableData || !tableData.columns) return null;
    
        const getColumnPatterns = (metric) => {
          const patterns = {
            'Gr': [
              /gr[-–\s]*\w{3}[-–\s]*\d{2,4}/i,
              /growth/i,
              /\bgr\b/i
            ],
            'Ach': [
              /ach[-–\s]*\w{3}[-–\s]*\d{2,4}/i,
              /achievement/i,
              /\bach\b/i
            ],
            'Act': [
              /act[-–\s]*\w{3}[-–\s]*\d{2,4}/i,
              /\bact\b/i
            ],
            'LY': [
            /^LY-\w{3}-\d{2}$/i,  // Exact match for LY-Apr-24
            /^LY\s+\w{3}\s+\d{2,4}$/i,
            /^Last\s+Year\s+\w{3}\s+\d{2,4}$/i,
            /^LY\b/i
          ]
            
          };
          return patterns[metric] || [];
        };
    
        let targetCols = [];
        if (label === 'Gr' || label === 'Ach' || label === 'Act' || label === 'LY') {
          const patterns = getColumnPatterns(label);
          targetCols = tableData.columns.filter(col => {
            const colStr = String(col).toLowerCase().replace(/[,–]/g, '-');
            return patterns.some(pattern => pattern.test(colStr)) && 
                   !colStr.includes('ytd') && 
                   columnFilter(col);
          });
        } else {
          targetCols = tableData.columns.filter(col => {
            const colStr = String(col).toLowerCase();
            if (label === 'Budget') {
              return colStr.includes('budget') && !colStr.includes('ytd');
            } else if (label === 'LY') {
              return colStr.includes('ly') || colStr.includes('last year');
            }
            return false;
          }).filter(col => columnFilter(col));
        }
    
        if (!targetCols.length) return null;
    
        const filteredData = processedData;
    
        const chartData = targetCols.map(col => {
          const monthYear = extractMonthYear(col);
          const value = filteredData.reduce((sum, row) => sum + (parseFloat(String(row[col]).replace(',', '')) || 0), 0);
          return { month: monthYear, value };
        }).filter(d => d.value !== null && !isNaN(d.value));
    
        if (!chartData.length) return null;
    
        chartData.sort((a, b) => {
          const getFiscalSortKey = month => {
            const [m, y] = month.split('-');
            const monthIdx = MONTH_ORDER.indexOf(m);
            const yearInt = parseInt(y);
            return (monthIdx >= 9 ? yearInt - 1 : yearInt) * 100 + monthIdx;
          };
          return getFiscalSortKey(a.month) - getFiscalSortKey(b.month);
        });
    
        const plotType = visualType.toLowerCase().includes('bar') ? 'bar' : 
                     visualType.toLowerCase().includes('line') ? 'scatter' : 'pie';
        const isPie = plotType === 'pie';
        return {
          data: [{
            x: isPie ? chartData.map(d => d.month) : chartData.map(d => d.month),
            y: isPie ? chartData.map(d => d.value) : chartData.map(d => d.value),
            values: isPie ? chartData.map(d => d.value) : undefined,
            labels: isPie ? chartData.map(d => d.month) : undefined,
            type: plotType,
            marker: { color: label === 'Act' ? '#FF8C00' : '#2E86AB' },
            mode: plotType === 'scatter' ? 'lines+markers' : undefined,
            hovertemplate: isPie ? 
              '<b>%{label}</b><br>Value: %{value:,.0f}<br>Percentage: %{percent:.1%}<extra></extra>' : 
              '<b>%{x}</b><br>Value: %{y:,.0f}<extra></extra>',
            textposition: isPie ? 'inside' : undefined,
            textinfo: isPie ? 'percent' : undefined
          }],
          layout: {
            title: { text: `${label}`, x: 0.5, xanchor: 'center', font: { size: 16 } },
            xaxis: { title: 'Month', titlefont: { size: 14 }, tickfont: { size: 12 },automargin: true , tickangle:  0 },
            yaxis: { title: 'Value', titlefont: { size: 14 }, tickfont: { size: 12 }, visible: !isPie },
            plot_bgcolor: 'white',
            paper_bgcolor: 'white',
            height: 500,
            margin: { l: 60, r: 60, t: 80, b: 60 },
            showlegend: isPie,
            hovermode: isPie ? 'closest' : 'x unified'
          }
        };
      };
    
      const generateYTDData = (label) => {
        if (!tableData || !tableData.columns) return null;
        const metric = label.replace('YTD ', '').trim();
        const ytdPatterns = [
      // Existing patterns
      new RegExp(`${metric}-YTD-\\d{2}-\\d{2}\\(.*?\\)`, 'i'),
      new RegExp(`YTD-\\d{2}-\\d{2}\\(.*?\\)\\s*${metric}`, 'i'),
      new RegExp(`${metric}-YTD\\s*\\d{4}\\(.*?\\)`, 'i'),
      new RegExp(`YTD\\s*\\d{4}\\(.*?\\)\\s*${metric}`, 'i'),
      // Add more flexible patterns
      new RegExp(`\\b${metric}\\b.*YTD`, 'i'),  // Matches "Budget YTD" or "LY YTD"
      new RegExp(`YTD.*\\b${metric}\\b`, 'i'),  // Matches "YTD Budget" or "YTD LY"
      new RegExp(`\\b${metric}\\b.*\\d{4}`, 'i') // Matches "Budget 2023" etc.
    ];
    
        const targetCols = tableData.columns.filter(col => {
          const colStr = String(col).replace(/[,–]/g, '-');
          return ytdPatterns.some(pattern => pattern.test(colStr)) && 
                 columnFilter(col, null, selectedYear);
        });
    
        if (!targetCols.length) return null;
    
        const filteredData = processedData;
    
        const ytdData = targetCols.map(col => {
          const colStr = String(col);
          const ytdMatch = colStr.match(/(\d{2,4})[-–](\d{2,4})\s*\((.*?)\)/i);
          const period = ytdMatch ? `YTD ${ytdMatch[1]}-${ytdMatch[2]} (${ytdMatch[3]})` : colStr;
          
          const value = filteredData.reduce((sum, row) => sum + (parseFloat(String(row[col]).replace(',', '')) || 0), 0);
          return { period, value };
        }).filter(d => d.value !== null && !isNaN(d.value));
    
        if (!ytdData.length) return null;
    
        // Sort YTD periods in fiscal year order (Apr-Mar)
        ytdData.sort((a, b) => {
          const monthOrder = {
            'Apr': 1, 'May': 2, 'Jun': 3, 'Jul': 4, 'Aug': 5, 'Sep': 6,
            'Oct': 7, 'Nov': 8, 'Dec': 9, 'Jan': 10, 'Feb': 11, 'Mar': 12
          };
          
          const getSortKey = (period) => {
            const match = period.match(/\((Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar)/i);
            return match ? monthOrder[match[1]] : 99;
          };
          
          return getSortKey(a.period) - getSortKey(b.period);
        });
    
        const plotType = visualType.toLowerCase().includes('bar') ? 'bar' : 
                         visualType.toLowerCase().includes('line') ? 'scatter' : 'pie';
        const isPie = plotType === 'pie';
        
        return {
          data: [{
            x: isPie ? ytdData.map(d => d.period) : ytdData.map(d => d.period),
            y: isPie ? ytdData.map(d => d.value) : ytdData.map(d => d.value),
            values: isPie ? ytdData.map(d => d.value) : undefined,
            labels: isPie ? ytdData.map(d => d.period) : undefined,
            type: plotType,
            marker: { color: metric === 'Act' ? '#FF8C00' : '#2E86AB' },
            mode: plotType === 'scatter' ? 'lines+markers' : undefined,
            hovertemplate: isPie ? 
              '<b>%{label}</b><br>Value: %{value:,.0f}<br>Percentage: %{percent:.1%}<extra></extra>' : 
              '<b>%{x}</b><br>Value: %{y:,.0f}<extra></extra>',
            textposition: isPie ? 'inside' : undefined,
            textinfo: isPie ? 'percent' : undefined
          }],
          layout: {
            title: { text: `${label}`, x: 0.5, xanchor: 'center', font: { size: 16 } },
            xaxis: { 
              title: 'Period', 
              titlefont: { size: 14 }, 
              tickfont: { size: 12 }, 
              tickangle: 0,
              automargin: true 
            },
            yaxis: { 
              title: 'Value', 
              titlefont: { size: 14 }, 
              tickfont: { size: 12 }, 
              visible: !isPie 
            },
            plot_bgcolor: 'white',
            paper_bgcolor: 'white',
            height: 500,
            margin: { l: 60, r: 60, t: 80, b: 60 },
            showlegend: isPie,
            hovermode: isPie ? 'closest' : 'x unified'
          }
        };
      };
    
      const generateBranchPerformanceData = () => {
        if (!isBranchAnalysis || !tableData || !tableData.columns) return null;
    
        const ytdActCol = tableData.columns.find(col => 
          col === "Act-YTD-25-26 (Apr to Mar)" || 
          col.match(/YTD[-–\s]*\d{2}[-–\s]*\d{2}\s*\([^)]*\)\s*Act/i)
        );
    
        if (!ytdActCol) return null;
    
        const firstCol = tableData.columns[0];
        const branchData = tableData.data
          .filter(row => row[firstCol] && !BRANCH_EXCLUDE_TERMS.some(term => String(row[firstCol]).toUpperCase().includes(term)))
          .filter(row => selectedBranch === 'Select All' || String(row[firstCol]).trim() === selectedBranch)
          .map(row => ({
            branch: String(row[firstCol]).trim(),
            performance: parseFloat(String(row[ytdActCol]).replace(',', '')) || 0
          }))
          .filter(d => !isNaN(d.performance))
          .sort((a, b) => b.performance - a.performance);
    
        if (!branchData.length) return null;
    
        const plotType = visualType.toLowerCase().includes('bar') ? 'bar' : 
                         visualType.toLowerCase().includes('line') ? 'scatter' : 'pie';
        const isPie = plotType === 'pie';
        return {
          data: [{
            x: isPie ? branchData.map(d => d.branch) : branchData.map(d => d.branch),
            y: isPie ? branchData.map(d => d.performance) : branchData.map(d => d.performance),
            values: isPie ? branchData.map(d => d.performance) : undefined,
            labels: isPie ? branchData.map(d => d.branch) : undefined,
            type: plotType,
            marker: { color: '#2E86AB' },
            mode: plotType === 'scatter' ? 'lines+markers' : undefined,
            hovertemplate: isPie ? 
              '<b>%{label}</b><br>Value: %{value:,.0f}<br>Percentage: %{percent:.1%}<extra></extra>' : 
              '<b>%{x}</b><br>Performance: %{y:,.0f}<extra></extra>',
            textposition: isPie ? 'inside' : undefined,
            textinfo: isPie ? 'percent' : undefined
          }],
          layout: {
            title: { text: `Branch Performance`, x: 0.5, xanchor: 'center', font: { size: 16 } },
            xaxis: { title: 'Branch', titlefont: { size: 14 }, tickfont: { size: 12 }, tickangle: 0 },
            yaxis: { title: 'Performance', titlefont: { size: 14 }, tickfont: { size: 12 }, visible: !isPie },
            plot_bgcolor: 'white',
            paper_bgcolor: 'white',
            height: 500,
            margin: { l: 60, r: 60, t: 80, b: 60 },
            showlegend: isPie,
            hovermode: isPie ? 'closest' : 'x unified'
          },
          metrics: {
            topPerformer: branchData[0],
            totalPerformance: branchData.reduce((sum, d) => sum + d.performance, 0),
            avgPerformance: branchData.reduce((sum, d) => sum + d.performance, 0) / branchData.length,
            top5: branchData.slice(0, 5),
            bottom5: branchData.slice(-5).reverse()
          }
        };
      };
    
    
      const generateBranchMonthwiseData = () => {
        if (!isBranchAnalysis || !tableData || !tableData.columns) return null;
    
        const actCols = tableData.columns.filter(col => 
          col.toLowerCase().match(/\bact\b.*(apr|may|jun|jul|aug|sep|oct|nov|dec|jan|feb|mar)[-\s]*\d{2}/i) && 
          !col.toLowerCase().includes('ytd') && 
          columnFilter(col)
        );
    
        if (!actCols.length) return null;
    
        const firstCol = tableData.columns[0];
        const monthwiseData = tableData.data
          .filter(row => row[firstCol] && !BRANCH_EXCLUDE_TERMS.some(term => String(row[firstCol]).toUpperCase().includes(term)))
          .filter(row => selectedBranch === 'Select All' || String(row[firstCol]).trim() === selectedBranch)
          .map(row => {
            const monthData = { branch: String(row[firstCol]).trim() };
            actCols.forEach(col => {
              const monthYear = extractMonthYear(col);
              monthData[monthYear] = parseFloat(String(row[col]).replace(',', '')) || 0;
            });
            return monthData;
          });
    
        if (!monthwiseData.length) return null;
    
        const months = actCols.map(col => extractMonthYear(col)).sort((a, b) => {
          const [monthA, yearA] = a.split('-');
          const [monthB, yearB] = b.split('-');
          const monthIdxA = MONTH_ORDER.indexOf(monthA);
          const monthIdxB = MONTH_ORDER.indexOf(monthB);
          const yearIntA = parseInt(yearA);
          const yearIntB = parseInt(yearB);
          return (monthIdxA >= 9 ? yearIntA - 1 : yearIntA) * 100 + monthIdxA - 
                 (monthIdxB >= 9 ? yearIntB - 1 : yearIntB) * 100 - monthIdxB;
        });
    
        const aggregatedData = months.map(month => ({
          month,
          value: monthwiseData.reduce((sum, row) => sum + (row[month] || 0), 0)
        }));
    
        const plotType = visualType.toLowerCase().includes('bar') ? 'bar' : 'scatter';
        return {
          data: [{
            x: aggregatedData.map(d => d.month),
            y: aggregatedData.map(d => d.value),
            type: plotType,
            marker: { color: '#2E86AB' },
            mode: plotType === 'scatter' ? 'lines+markers' : undefined,
            hovertemplate: '<b>%{x}</b><br>Value: %{y:,.0f}<extra></extra>'
          }],
          layout: {
            title: { text: `Branch Monthwise Performance`, x: 0.5, xanchor: 'center', font: { size: 16 } },
            xaxis: { title: 'Month', titlefont: { size: 14 }, tickfont: { size: 12 }, categoryorder: 'array', categoryarray: months,automargin: true  },
            yaxis: { title: 'Value', titlefont: { size: 14 }, tickfont: { size: 12 } },
            plot_bgcolor: 'white',
            paper_bgcolor: 'white',
            height: 500,
            margin: { l: 60, r: 60, t: 80, b: 60 },
            hovermode: 'x unified'
          },
          metrics: {
            bestMonth: aggregatedData.reduce((max, d) => d.value > (max.value || 0) ? d : max, {}),
            avgMonthly: aggregatedData.reduce((sum, d) => sum + d.value, 0) / aggregatedData.length,
            totalPerformance: aggregatedData.reduce((sum, d) => sum + d.value, 0)
          }
        };
      };
    
      const generateProductPerformanceData = () => {
      if (!isProductAnalysis || !tableData || !tableData.columns) return null;
    
      // Updated performance column detection logic to match Flask
      const performanceCol = (() => {
        // First look for Act YTD columns with highest year
        const actYtdCols = tableData.columns
          .filter(col => {
            const colStr = String(col).toUpperCase();
            return colStr.includes('ACT') && colStr.includes('YTD');
          })
          .map(col => {
            const colStr = String(col).toUpperCase();
            const ytdMatch = colStr.match(/YTD[-\s]*(\d{2})[-\s]*(\d{2})/);
            const startYear = ytdMatch ? parseInt(ytdMatch[1]) : 0;
            const endYear = ytdMatch ? parseInt(ytdMatch[2]) : 0;
            
            // Calculate sum for this column
            const colSum = processedData.reduce((sum, row) => {
              const valStr = String(row[col]).replace(',', '');
              const val = parseFloat(valStr) || 0;
              return sum + val;
            }, 0);
            
            return { endYear, startYear, colSum, col };
          })
          .sort((a, b) => b.endYear - a.endYear || b.colSum - a.colSum);
    
        // For products, prefer "Apr to Mar" fiscal year
        if (actYtdCols.length > 0 && isProductAnalysis) {
          const fiscalYearCol = actYtdCols.find(item => 
            String(item.col).toUpperCase().includes('(APR TO MAR)')
          );
          if (fiscalYearCol) return fiscalYearCol.col;
          return actYtdCols[0].col;
        }
    
        // Fallback to any YTD column if no Act-YTD found
        const ytdCols = tableData.columns
          .filter(col => String(col).toUpperCase().includes('YTD'))
          .map(col => {
            const colStr = String(col).toUpperCase();
            const ytdMatch = colStr.match(/YTD[-\s]*(\d{2})[-\s]*(\d{2})/);
            const startYear = ytdMatch ? parseInt(ytdMatch[1]) : 0;
            const endYear = ytdMatch ? parseInt(ytdMatch[2]) : 0;
            
            const colSum = processedData.reduce((sum, row) => {
              const valStr = String(row[col]).replace(',', '');
              const val = parseFloat(valStr) || 0;
              return sum + val;
            }, 0);
            
            return { endYear, startYear, colSum, col };
          })
          .sort((a, b) => b.endYear - a.endYear || b.colSum - a.colSum);
    
        if (ytdCols.length > 0) return ytdCols[0].col;
    
        // Final fallback - find first numeric column with highest sum
        const numericCols = tableData.columns
          .slice(1) // Skip first column (names)
          .map(col => {
            const colSum = processedData.reduce((sum, row) => {
              const valStr = String(row[col]).replace(',', '');
              const val = parseFloat(valStr) || 0;
              return sum + val;
            }, 0);
            return { colSum, col };
          })
          .filter(col => col.colSum > 0) // Only consider columns with positive values
          .sort((a, b) => b.colSum - a.colSum);
    
        if (numericCols.length > 0) return numericCols[0].col;
    
        return null;
      })(); // This is the closing parenthesis for the IIFE
    
      if (!performanceCol) {
        console.log('No valid performance column found for product analysis');
        return null;
      }
    
      const firstCol = tableData.columns[0];
      const productData = processedData
        .filter(row => row[firstCol] && 
          !PRODUCT_EXCLUDE_TERMS.some(term => 
            String(row[firstCol]).toUpperCase().includes(term.toUpperCase())))
        .map(row => {
          const valueStr = String(row[performanceCol]).replace(',', '');
          const value = parseFloat(valueStr) || 0;
          return {
            product: String(row[firstCol]).trim(),
            performance: value
          };
        })
        .filter(d => !isNaN(d.performance))
        .sort((a, b) => b.performance - a.performance);
    
      console.log('Product Data:', productData);
    
      if (!productData.length) return null;
    
      const plotType = visualType.toLowerCase().includes('bar') ? 'bar' : 
                       visualType.toLowerCase().includes('line') ? 'scatter' : 'pie';
      const isPie = plotType === 'pie';
      // Calculate bar width based on number of products
      const barWidth = Math.max(0.3, Math.min(0.8, 30 / productData.length));
      
      return {
        data: [{
          x: isPie ? productData.map(d => d.product) : productData.map(d => d.product),
          y: isPie ? productData.map(d => d.performance) : productData.map(d => d.performance),
          values: isPie ? productData.map(d => d.performance) : undefined,
          labels: isPie ? productData.map(d => d.product) : undefined,
          type: plotType,
          marker: { color: '#2E86AB' },
          mode: plotType === 'scatter' ? 'lines+markers' : undefined,
          textangle: 0,
          width: barWidth, // Dynamic width based on product count
          hoverinfo: 'x+y',
          hovertemplate: isPie ? 
            '<b>%{label}</b><br>Performance: %{value:,.0f}<br>Percentage: %{percent:.1%}<extra></extra>' : 
            '<b>%{x}</b><br>Performance: %{y:,.0f}<extra></extra>',
          textposition: isPie ? 'inside' : undefined,
          textinfo: isPie ? 'percent' : undefined
        }],
        layout: {
          title: { text: `Product Performance`, x: 0.5, xanchor: 'center', font: { size: 16 } },
          xaxis: { 
            title: 'Product', 
            titlefont: { size: 14 }, 
            tickfont: { size: 12 }, 
            tickangle: isPie ? 0 : 45,
            automargin: true, // Prevent label cutoff
            tickmode: 'array',
            tickvals: productData.map((_, i) => i), // Ensure all ticks are shown
            type: 'category' // Treat as categorical data
          },
          yaxis: { 
            title: 'Performance', 
            titlefont: { size: 14 }, 
            tickfont: { size: 12 }, 
            visible: !isPie 
          },
          plot_bgcolor: 'white',
          paper_bgcolor: 'white',
          height: 500,
          margin: { 
            l: 80, // Left margin
            r: 80, // Right margin
            t: 80, // Top margin
            b: productData.length > 10 ? 150 : 100 // Bottom margin (larger if many products)
          },
          showlegend: isPie,
          bargap: 0.2, // Gap between bars
          hovermode: isPie ? 'closest' : 'x unified'
        },
        metrics: {
          topPerformer: productData[0],
          totalPerformance: productData.reduce((sum, d) => sum + d.performance, 0),
          avgPerformance: productData.reduce((sum, d) => sum + d.performance, 0) / productData.length,
          top5: productData.slice(0, Math.min(5, productData.length)),
          bottom5: productData.slice(-5).reverse().slice(0, Math.min(5, productData.length))
        }
      };
    };
    
      const generateProductMonthwiseData = () => {
      if (!isProductAnalysis || !tableData || !tableData.columns) return null;
    
      const actCols = tableData.columns.filter(col => 
        col.toLowerCase().match(/\bact\b.*(apr|may|jun|jul|aug|sep|oct|nov|dec|jan|feb|mar)[-\s]*\d{2}/i) && 
        !col.toLowerCase().includes('ytd') && 
        columnFilter(col)
      );
    
      if (!actCols.length) return null;
    
      const firstCol = tableData.columns[0];
      const monthwiseData = processedData
        .filter(row => row[firstCol] && 
          !PRODUCT_EXCLUDE_TERMS.some(term => String(row[firstCol]).toUpperCase().includes(term.toUpperCase())))
        .map(row => {
          const monthData = { product: String(row[firstCol]).trim() };
          actCols.forEach(col => {
            const monthYear = extractMonthYear(col);
            monthData[monthYear] = parseFloat(String(row[col]).replace(',', '')) || 0;
          });
          return monthData;
        });
    
      if (!monthwiseData.length) return null;
    
      const months = actCols.map(col => extractMonthYear(col)).sort((a, b) => {
        const [monthA, yearA] = a.split('-');
        const [monthB, yearB] = b.split('-');
        const monthIdxA = MONTH_ORDER.indexOf(monthA);
        const monthIdxB = MONTH_ORDER.indexOf(monthB);
        const yearIntA = parseInt(yearA);
        const yearIntB = parseInt(yearB);
        return (monthIdxA >= 9 ? yearIntA - 1 : yearIntA) * 100 + monthIdxA - 
               (monthIdxB >= 9 ? yearIntB - 1 : yearIntB) * 100 - monthIdxB;
      });
    
      const aggregatedData = months.map(month => ({
        month,
        value: monthwiseData.reduce((sum, row) => sum + (row[month] || 0), 0)
      }));
    
      const plotType = visualType.toLowerCase().includes('bar') ? 'bar' : 
                       visualType.toLowerCase().includes('line') ? 'scatter' : 'pie';
      const isPie = plotType === 'pie';
    
      const chartData = isPie 
        ? aggregatedData.filter(d => d.value > 0)
        : aggregatedData;
    
      if (chartData.length === 0) return null;
    
      return {
        data: [{
          x: isPie ? chartData.map(d => d.month) : chartData.map(d => d.month),
          y: isPie ? chartData.map(d => d.value) : chartData.map(d => d.value),
          values: isPie ? chartData.map(d => d.value) : undefined,
          labels: isPie ? chartData.map(d => d.month) : undefined,
          type: plotType,
          marker: { 
            color: isPie ? ['#2E86AB', '#FF8C00', '#A23B72', '#F18F01', '#C73E1D'] : '#2E86AB',
            line: isPie ? { width: 1, color: '#FFFFFF' } : undefined
          },
          textposition: isPie ? 'inside' : undefined, // Removed 'auto' to hide bar values
          textinfo: isPie ? 'percent+label' : 'none', // Disabled text on bars
          insidetextorientation: isPie ? 'horizontal' : undefined,
          hovertemplate: isPie ? 
            '<b>%{label}</b><br>Value: %{value:,.0f}<br>Percentage: %{percent:.1%}<extra></extra>' : 
            '<b>%{x}</b><br>Value: %{y:,.0f}<extra></extra>',
          textangle: 0,
          width: plotType === 'bar' ? 0.6 : undefined
        }],
        layout: {
          title: { 
            text: `Product Monthwise Performance (${visualType})`, 
            x: 0.5, 
            xanchor: 'center', 
            font: { size: 16 } 
          },
          xaxis: { 
            title: 'Month', 
            titlefont: { size: 14 }, 
            tickfont: { size: 12 }, 
            categoryorder: 'array', 
            categoryarray: months,
            tickangle: 0, // Made x-axis labels straight (0° rotation)
            automargin: true
          },
          yaxis: { 
            title: 'Value', 
            titlefont: { size: 14 }, 
            tickfont: { size: 12 }, 
            visible: !isPie 
          },
          plot_bgcolor: 'white',
          paper_bgcolor: 'white',
          height: 500,
          margin: { l: 60, r: 60, t: 80, b: 100 },
          showlegend: isPie,
          hovermode: isPie ? 'closest' : 'x unified',
          uniformtext: isPie ? {
            minsize: 12,
            mode: 'hide'
          } : undefined
        },
        metrics: {
          bestMonth: chartData.reduce((max, d) => d.value > (max.value || 0) ? d : max, {}),
          avgMonthly: chartData.reduce((sum, d) => sum + d.value, 0) / chartData.length,
          totalPerformance: chartData.reduce((sum, d) => sum + d.value, 0)
        }
      };
    };
    
      // Visualization generation effect
      useEffect(() => {
        if (!tableData || !tableData.data) return;
    
        if (isSalesAnalysisMonthwise) {
          console.log("🚫 Visualizations blocked for Sales Month-wise sheet");
          setVisualizationData({});
          return;
        }
    
        const generateViz = async () => {
          setIsLoading(true);
          setIsProcessing(true);
          setProcessingProgress(0);
    
          try {
            setProcessingProgress(20);
    
            const visualizations = await Promise.all([
              generateBudgetVsActualData(),
              generateDefaultData('Budget'),
              generateDefaultData('LY'),
              generateDefaultData('Act'),
              generateDefaultData('Gr'),
              generateDefaultData('Ach'),
              generateYTDData('YTD Budget'),
              generateYTDData('YTD LY'),
              generateYTDData('YTD Act'),
              generateYTDData('YTD Gr'),
              generateYTDData('YTD Ach'),
              generateBranchPerformanceData(),
              generateBranchMonthwiseData(),
              generateProductPerformanceData(),
              generateProductMonthwiseData(),
            ]);
    
            setProcessingProgress(80);
    
            const vizData = {
              'Budget vs Actual': visualizations[0],
              'Budget': visualizations[1],
              'LY': visualizations[2],
              'Act': visualizations[3],
              'Gr': visualizations[4],
              'Ach': visualizations[5],
              'YTD Budget': visualizations[6],
              'YTD LY': visualizations[7],
              'YTD Act': visualizations[8],
              'YTD Gr': visualizations[9],
              'YTD Ach': visualizations[10],
              'Branch Performance': visualizations[11],
              'Branch Monthwise': visualizations[12],
              'Product Performance': visualizations[13],
              'Product Monthwise': visualizations[14],
            };
    
            setVisualizationData(vizData);
          } catch (err) {
            console.error("Visualization generation error:", err);
            setError(`Error generating visualizations: ${err.message}`);
          } finally {
            setIsLoading(false);
            setIsProcessing(false);
            setProcessingProgress(100);
          }
        };
    
        generateViz();
      }, [
        tableData,
        visualType,
        selectedMonth,
        selectedYear,
        selectedBranch,
        selectedProduct,
        isBranchAnalysis,
        isProductAnalysis,
        months,
        years,
        branches,
        products,
        isSalesAnalysisMonthwise,
      ]);
    
      const downloadPPT = async (tabName) => {
      try {
        if (!visualizationData[tabName] || !visualizationData[tabName].data) {
          throw new Error(`No visualization data available for ${tabName}`);
        }
    
        setIsLoading(true);
        setIsProcessing(true);
        setProcessingProgress(0);
    
        const visData = visualizationData[tabName];
        const isPie = visualType.toLowerCase().includes('pie');
        const isLine = visualType.toLowerCase().includes('line');
        const isGrouped = visData.layout?.barmode === 'group';
    
        // Determine if this is a product-related chart
        const isProductChart = tabName.includes('Product');
        // Special handling for Budget vs Actual pie chart
        if (tabName === 'Budget vs Actual' && isPie) {
          // Extract the pie chart data
          const pieData = visData.data[0];
          // Validate pie chart data
          if (!pieData.values || pieData.values.length < 2) {
            throw new Error('Invalid pie chart data format');
          }
          
          // Prepare payload for pie chart
          const payload = {
            vis_label: tabName,
            table_name: tableName || selectedSheet,
            chart_data: [
              {
                name: 'Budget',
                y: pieData.values[0] || 0  // First value is budget
              },
              {
                name: 'Actual',
                y: pieData.values[1] || 0  // Second value is actual
              }
            ],
            x_col: 'Category',
            y_col: 'Value',
            ppt_type: 'pie',
            selected_filter: getSelectedFilter(),
            is_budget_vs_actual: true,
            is_pie_chart: true
          };
    
          const response = await api.generatePPT(payload);
          handlePPTDownload(response, tabName);
          return;
        }
    
        // Special handling for Budget vs Actual to match the first image format
        if (tabName === 'Budget vs Actual') {
          // Enhanced trace detection with multiple possible name variations
          const budgetTrace = visData.data.find(trace => 
            trace.name && (
              trace.name.toLowerCase().includes('budget') || 
              trace.name.toLowerCase().includes('planned') ||
              trace.name.toLowerCase().includes('target')
            )
          );
    
          const actualTrace = visData.data.find(trace => 
            trace.name && (
              trace.name.toLowerCase().includes('actual') || 
              trace.name.toLowerCase().includes('real') ||
              trace.name.toLowerCase().includes('achieved') ||
              trace.name.toLowerCase() === 'act'
            )
          );
    
          if (!budgetTrace || !actualTrace) {
            const availableTraces = visData.data.map(t => t.name || 'Unnamed Trace');
            throw new Error(
              `Budget vs Actual requires both Budget and Actual data traces.\n` +
              `Available traces: ${availableTraces.join(', ')}`
            );
          }
    
          // Validate month formats
          const months = budgetTrace.x || actualTrace.x;
          if (!months || months.some(month => typeof month !== 'string' || !/^[A-Za-z]{3}-\d{2}$/.test(month))) {
            throw new Error('Invalid month format in Budget vs Actual data. Expected format like "Apr-23"');
          }
    
          // Ensure both traces have same months in same order
          if (JSON.stringify(budgetTrace.x) !== JSON.stringify(actualTrace.x)) {
            console.warn('Month mismatch between budget and actual traces. Aligning months...');
            const alignedData = alignTraceMonths(budgetTrace, actualTrace);
            budgetTrace.x = alignedData.months;
            budgetTrace.y = alignedData.budgetValues;
            actualTrace.x = alignedData.months;
            actualTrace.y = alignedData.actualValues;
          }
          
          const pptType = isLine ? 'line' : isPie ? 'pie' : 'bar';
          const payload = {
            vis_label: tabName,
            table_name: tableData?.tableName || selectedSheet,
            chart_data: [
              {
                name: 'Budget',
                x: budgetTrace.x,
                y: budgetTrace.y,
                type: pptType, // Use dynamic type
                marker: { color: '#2E86AB' }
              },
              {
                name: 'Actual',
                x: actualTrace.x,
                y: actualTrace.y,
                type: pptType, // Use dynamic type
                marker: { color: '#FF8C00' }
              }
            ],
            x_col: 'Month',
            y_col: 'Value',
            ppt_type: pptType, // Use dynamic type
            selected_filter: getSelectedFilter(),
            is_budget_vs_actual: true // Flag for special handling
          };
    
          const response = await api.generatePPT(payload);
          handlePPTDownload(response, tabName);
          return;
        }
    
        // Standard processing for other chart types
        let x_col, y_col;
        switch (tabName) {
          case 'Branch Performance':
            x_col = 'Branch';
            y_col = 'Performance';
            break;
          case 'Product Performance':
            x_col = 'Product';
            y_col = 'Performance';
            break;
          case 'Branch Monthwise':
          case 'Product Monthwise':
            x_col = 'Month';
            y_col = 'Value';
            break;
          default:
            x_col = tabName.startsWith('YTD') ? 'Period' : 'Month';
            y_col = tabName.startsWith('YTD') ? tabName.replace('YTD ', '') : 'Value';
        }
    
        const chartData = visData.data.reduce((acc, trace) => {
          if (!trace.x && !trace.labels) return acc;
          
          const values = trace.y || trace.values || [];
          const labels = trace.x || trace.labels || [];
    
          if (isGrouped) {
            labels.forEach((x, i) => {
              const value = parseFloat(values[i]) || 0;
              if (isPie && value <= 0) return;
              acc.push({
                [x_col]: String(x),
                [y_col]: value,
                Metric: trace.name || 'Unknown'
              });
            });
          } else {
            labels.forEach((x, i) => {
              const value = parseFloat(values[i]) || 0;
              if (isPie && value <= 0) return;
              acc.push({
                [x_col]: String(x),
                [y_col]: value
              });
            });
          }
          return acc;
        }, []);
    
        if (chartData.length === 0) {
          throw new Error(`No valid chart data for ${tabName}. Ensure data contains numeric values${isPie ? ' and positive values for pie charts' : ''}.`);
        }
    
        const payload = {
          vis_label: tabName,
          table_name: tableData?.tableName || selectedSheet,
          chart_data: makeJsonlySerializable(chartData),
          x_col,
          y_col,
          ppt_type: visualType.toLowerCase().replace(' chart', ''),
          color_override: visData.data[0]?.marker?.color || null,
          selected_filter: getSelectedFilter(),
          is_budget_vs_actual: false,
          is_product_chart: isProductChart  // Add this flag
        };
    
        const response = await api.generatePPT(payload);
        handlePPTDownload(response, tabName);
    
      } catch (err) {
        console.error('PPT generation failed:', {
          error: err,
          response: err.response?.data,
          stack: err.stack
        });
        let errorMessage = `Failed to generate PPT for ${tabName}: ${err.message}`;
        if (err.response?.data?.error) {
          errorMessage = `Server error: ${err.response.data.error}`;
        }
        setError(errorMessage);
      } finally {
        setIsLoading(false);
        setIsProcessing(false);
        setProcessingProgress(100);
      }
    };
    
    // Helper function to align months between budget and actual traces
    function alignTraceMonths(budgetTrace, actualTrace) {
      const budgetMonths = budgetTrace.x || [];
      const actualMonths = actualTrace.x || [];
      
      // Create union of all unique months
      const allMonths = [...new Set([...budgetMonths, ...actualMonths])];
      
      // Sort months chronologically
      allMonths.sort((a, b) => {
        const parseDate = (str) => {
          const [month, year] = str.split('-');
          return new Date(`20${year}`, monthToNumber(month));
        };
        return parseDate(a) - parseDate(b);
      });
    
      // Create new aligned arrays
      const budgetValues = [];
      const actualValues = [];
      
      allMonths.forEach(month => {
        const budgetIndex = budgetMonths.indexOf(month);
        budgetValues.push(budgetIndex >= 0 ? budgetTrace.y[budgetIndex] : null);
        
        const actualIndex = actualMonths.indexOf(month);
        actualValues.push(actualIndex >= 0 ? actualTrace.y[actualIndex] : null);
      });
    
      return {
        months: allMonths,
        budgetValues,
        actualValues
      };
    }
    
    // Helper to convert month abbreviations to numbers
    function monthToNumber(month) {
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                     'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return months.indexOf(month);
    }
    
    // Helper function to handle the download
    const handlePPTDownload = (response, tabName) => {
      setProcessingProgress(90);
      const url = window.URL.createObjectURL(new Blob([response]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `${tabName.toLowerCase().replace(/\s+/g, '_')}_analysis.pptx`);
      document.body.appendChild(link);
      link.click();
      setTimeout(() => {
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
      }, 100);
    };
    
      const downloadMasterPPT = async () => {
      try {
        setIsLoading(true);
        setIsProcessing(true);
        setProcessingProgress(0);
    
        const visualizationTypes = [
          'Budget vs Actual',
          'Budget',
          'LY',
          'Act',
          'Gr',
          'Ach',
          'YTD Budget',
          'YTD LY',
          'YTD Act',
          'YTD Gr',
          'YTD Ach',
          'Branch Performance',
          'Branch Monthwise',
          'Product Performance',
          'Product Monthwise'
        ];
    
        const allData = visualizationTypes.map(label => {
          const visData = visualizationData[label];
          if (!visData || !visData.data) return null;
    
          // Determine if this is a product-related chart
          const isProductChart = label.includes('Product');
    
          // Special handling for Budget vs Actual
          if (label === 'Budget vs Actual') {
            // Check if this is a pie chart
            const isPieChart = visData.data.some(trace => trace.type === 'pie');
            
            if (isPieChart) {
              const pieTrace = visData.data.find(trace => trace.type === 'pie');
              if (!pieTrace || !pieTrace.labels || !pieTrace.values) return null;
              
              // Find budget and actual values in the pie data
              const budgetIndex = pieTrace.labels.findIndex(l => 
                String(l).toLowerCase().includes('budget'));
              const actualIndex = pieTrace.labels.findIndex(l => 
                String(l).toLowerCase().includes('actual'));
                
              if (budgetIndex === -1 || actualIndex === -1) return null;
              
              return {
                label,
                data: [
                  { Metric: 'Budget', Value: pieTrace.values[budgetIndex] || 0 },
                  { Metric: 'Actual', Value: pieTrace.values[actualIndex] || 0 }
                ],
                x_col: 'Metric',
                y_col: 'Value',
                is_pie_chart: true
              };
            }
            
            // Handle bar/line chart version
            const budgetTrace = visData.data.find(trace => 
              trace.name && (
                trace.name.toLowerCase().includes('budget') || 
                trace.name.toLowerCase().includes('planned') ||
                trace.name.toLowerCase().includes('target')
              )
            );
    
            const actualTrace = visData.data.find(trace => 
              trace.name && (
                trace.name.toLowerCase().includes('actual') || 
                trace.name.toLowerCase().includes('real') ||
                trace.name.toLowerCase().includes('achieved') ||
                trace.name.toLowerCase() === 'act'
              )
            );
    
            if (!budgetTrace || !actualTrace) {
              console.warn(`Missing budget or actual trace for ${label}`);
              return null;
            }
    
            // Align months and values
            const alignedData = [];
            const months = budgetTrace.x || [];
            const budgetValues = budgetTrace.y || [];
            const actualValues = actualTrace.y || [];
    
            months.forEach((month, i) => {
              alignedData.push({
                Month: month,
                Value: budgetValues[i] || 0,
                Metric: 'Budget'
              });
              alignedData.push({
                Month: month,
                Value: actualValues[i] || 0,
                Metric: 'Actual'
              });
            });
    
            return {
              label,
              data: alignedData,
              x_col: 'Month',
              y_col: 'Value',
              is_budget_vs_actual: true
            };
          }
    
          // Determine chart type from the visualization data
          const isPie = visData.data.some(trace => trace.type === 'pie');
          const isGrouped = visData.layout?.barmode === 'group';
    
          let x_col, y_col;
          switch (label) {
            case 'Branch Performance':
              x_col = 'Branch';
              y_col = 'Performance';
              break;
            case 'Product Performance':
              x_col = 'Product';
              y_col = 'Performance';
              break;
            case 'Branch Monthwise':
            case 'Product Monthwise':
              x_col = 'Month';
              y_col = 'Value';
              break;
            default:
              x_col = label.startsWith('YTD') ? 'Period' : 'Month';
              y_col = label.startsWith('YTD') ? label.replace('YTD ', '') : label;
          }
    
          const chartData = [];
          const traces = Array.isArray(visData.data) ? visData.data : [];
          
          traces.forEach(trace => {
            if (!trace.x && !trace.labels) return;
            
            const values = trace.y || trace.values || [];
            const labels = trace.x || trace.labels || [];
    
            labels.forEach((x, i) => {
              const value = parseFloat(values[i]) || 0;
              if (isPie && value <= 0) return;
              
              const dataPoint = {
                [x_col]: String(x),
                [y_col]: value
              };
              
              // Add metric name for grouped charts
              if (isGrouped && trace.name) {
                dataPoint.Metric = trace.name;
              }
              
              chartData.push(dataPoint);
            });
          });
    
          if (chartData.length === 0) return null;
    
          return {
            label,
            data: chartData,
            x_col,
            y_col,
            is_product_chart: isProductChart
          };
        }).filter(item => item !== null);
    
        if (allData.length === 0) {
          throw new Error('No valid visualization data available for Master PPT');
        }
    
        const payload = {
          all_data: allData,
          table_name: tableData?.tableName || selectedSheet,
          selected_sheet: selectedSheet,
          visual_type: visualType.toLowerCase().replace(' chart', ''),
          selected_filter: getSelectedFilter()
        };
    
        console.log('Master PPT payload:', payload);
    
        const response = await api.generateMasterPPT(payload);
    
        setProcessingProgress(90);
        const url = window.URL.createObjectURL(new Blob([response]));
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', `complete_analysis_${selectedSheet.toLowerCase().replace(' ', '_')}.pptx`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      } catch (err) {
        console.error('Master PPT generation failed:', err);
        let errorMessage = `Error generating Master PPT: ${err.message}`;
        if (err.response?.data?.error) {
          errorMessage = `Server error: ${err.response.data.error}`;
        }
        setError(errorMessage);
      } finally {
        setIsLoading(false);
        setIsProcessing(false);
        setProcessingProgress(100);
      }
    };
    
      const getSelectedFilter = () => {
        if (isBranchAnalysis && selectedBranch !== 'Select All') return selectedBranch;
        if (isProductAnalysis && selectedProduct !== 'Select All') return selectedProduct;
        if (selectedMonth !== 'Select All') return selectedMonth;
        if (selectedYear !== 'Select All') return selectedYear;
        return null;
      };
    
      const downloadCSV = async () => {
      if (!tableData || !processedData) {
        setError("No data available to download");
        return;
      }
    
      try {
        setIsLoading(true);
        
        // Prepare data exactly as shown in the table
        const csvData = processedData.map(row => {
          const csvRow = {};
          tableData.columns.forEach(col => {
            // Preserve the exact formatting from the table
            const value = row[col];
            if (value === null || value === undefined || value === 'NaN') {
              csvRow[col] = '';
            } else {
              // Check if the value is numeric
              const numValue = parseFloat(String(value).replace(/,/g, ''));
              if (!isNaN(numValue)) {
                // Format numbers with commas and 2 decimal places if needed
                csvRow[col] = numValue % 1 === 0 ? 
                  numValue.toLocaleString() : 
                  numValue.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
              } else {
                // Keep strings as-is
                csvRow[col] = String(value);
              }
            }
          });
          return csvRow;
        });
    
        const payload = {
          table_data: csvData,
          columns: tableData.columns,
          filename: `${tableName || 'data'}_${new Date().toISOString().slice(0,10)}.csv`
        };
    
        const response = await api.downloadCSV(payload);
        
        // Handle the download
        const blob = new Blob([response.data], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', payload.filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
      } catch (err) {
        setError(`CSV Download Failed: ${err.message}`);
      } finally {
        setIsLoading(false);
      }
    };
    
      useEffect(() => {
        const testConnection = async () => {
          try {
            await api.testConnection();
            console.log("✅ Backend connection successful");
            setConnectionStatus("connected");
          } catch (error) {
            setError("❌ Cannot connect to backend server. Please make sure Flask is running on http://localhost:5003");
            console.error("❌ Backend connection failed:", error);
            setConnectionStatus("failed");
          }
        };
    
        testConnection();
      }, []);
    
      const renderTabContent = () => {
        if (isSalesAnalysisMonthwise) {
          return (
            <div className="tab-content">
              <div className="data-table-container">
                <h3>📋 Sales Month-wise Data</h3>
                <div className="table-wrapper">
                  <table>
                    <thead>
                      <tr>
                        {tableData.columns.map((header, idx) => (
                          <th key={idx}>{header}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {processedData.map((row, rowIdx) => (
                        <tr key={rowIdx}>
                          {tableData.columns.map((col, colIdx) => (
                            <td key={colIdx}>
                              {row[col] || '0.00'}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
              <button onClick={downloadCSV} className="download-btn">
                ⬇️ Download CSV
              </button>
            </div>
          );
        }
    
        const data = visualizationData[activeTab];
        if (!data || !data.data || !processedData || !tableData?.columns) {
          return (
            <div className="tab-content">
              <div className="no-data-message">
                No data available for {activeTab}.{' '}
                {tableChoice && tableChoice !== 'Default Table' && !tableAvailability[selectedSheet]?.includes(tableChoice) && (
                  <span>Selected table "{tableChoice}" is not available in "{selectedSheet}". Using default table.</span>
                )}
                {(!processedData || processedData.length === 0) && (
                  <span>
                    Table data is empty or rows are missing. 
                    {isBranchAnalysis && "The selected table may only contain summary rows (e.g., 'North Total' or 'Grand Total') that have been filtered out by the backend. "}
                    {isProductAnalysis && "The selected table may only contain summary rows (e.g., 'Total Sales' or 'PRODUCT WISE SALES - VALUE DATA') that have been filtered out. "}
                    Please check the Excel file or select a different sheet/table.
                  </span>
                )}
              </div>
            </div>
          );
        }
    
        return (
          <div className="tab-content">
            <h3>{activeTab}</h3>
            <div className="visualization-container">
              <Plot
                data={data.data}
                layout={data.layout}
                config={{
                  displayModeBar: true,
                  displaylogo: false,
                  modeBarButtonsToRemove: ['pan2d', 'lasso2d', 'select2d'],
                  toImageButtonOptions: {
                    format: 'png',
                    filename: 'chart',
                    height: 800,
                    width: 1200,
                    scale: 2
                  }
                }}
                useResizeHandler={true}
                style={{ width: '100%', height: '100%' }}
              />
            </div>
    
            {(activeTab === 'Branch Performance' || activeTab === 'Product Performance') && data.metrics && (
              <div className="metrics-container">
                <div className="metrics-row">
                  <div className="metric-card">
                    <h4>🏆 Top Performer</h4>
                    <p className="metric-value">{data.metrics.topPerformer.branch || data.metrics.topPerformer.product}</p>
                    <p className="metric-number">{data.metrics.topPerformer.performance.toLocaleString()}</p>
                  </div>
                  <div className="metric-card">
                    <h4>📊 Total Performance</h4>
                    <p className="metric-number">{data.metrics.totalPerformance.toLocaleString()}</p>
                  </div>
                  <div className="metric-card">
                    <h4>📈 Average Performance</h4>
                    <p className="metric-number">{Math.round(data.metrics.avgPerformance).toLocaleString()}</p>
                  </div>
                </div>
                <div className="metrics-row">
                  <div className="top-bottom-card">
                    <h4>🏆 Top 5</h4>
                    <div className="top-bottom-list">
                      {data.metrics.top5.map((item, idx) => (
                        <div key={idx} className="top-bottom-item">
                          <span>{item.branch || item.product}</span>
                          <span>{item.performance.toLocaleString()}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                  <div className="top-bottom-card">
                    <h4>📉 Bottom 5</h4>
                    <div className="top-bottom-list">
                      {data.metrics.bottom5.map((item, idx) => (
                        <div key={idx} className="top-bottom-item">
                          <span>{item.branch || item.product}</span>
                          <span>{item.performance.toLocaleString()}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            )}
    
            {(activeTab === 'Branch Monthwise' || activeTab === 'Product Monthwise') && data.metrics && (
              <div className="metrics-container">
                <div className="metrics-row">
                  <div className="metric-card">
                    <h4>🏆 Best Month</h4>
                    <p className="metric-value">{data.metrics.bestMonth.month}</p>
                    <p className="metric-number">{data.metrics.bestMonth.value.toLocaleString()}</p>
                  </div>
                  <div className="metric-card">
                    <h4>📊 Monthly Average</h4>
                    <p className="metric-number">{Math.round(data.metrics.avgMonthly).toLocaleString()}</p>
                  </div>
                  <div className="metric-card">
                    <h4>📈 Total Performance</h4>
                    <p className="metric-number">{data.metrics.totalPerformance.toLocaleString()}</p>
                  </div>
                </div>
              </div>
            )}
    
            <div className="data-table-container">
              <h3>Data Table</h3>
              <div className="table-wrapper">
                {isTableLoading ? (
                  <div className="loading-spinner">
                    <div className="spinner"></div>
                    <p>Loading table data...</p>
                  </div>
                ) : processedData.length === 0 ? (
                  <p>
                    No rows available. 
                    {isBranchAnalysis && "The selected table may only contain summary rows (e.g., 'North Total' or 'Grand Total') that have been filtered out by the backend. "}
                    {isProductAnalysis && "The selected table may only contain summary rows (e.g., 'Total Sales' or 'PRODUCT WISE SALES - VALUE DATA') that have been filtered out. "}
                    Check the Excel file or select a different sheet/table.
                  </p>
                ) : (
                  <table>
                    <thead>
                      <tr>
                        {tableData.columns.map((header, idx) => (
                          <th key={idx}>{cleanColumnHeader(header) || `Column ${idx + 1}`}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {processedData.map((row, rowIdx) => (
                        <tr key={rowIdx} className={row.isTotalRow ? 'total-row' : ''}>
                          {tableData.columns.map((col, colIdx) => (
                            <td key={colIdx}>
                              {row[col] !== null && !isNaN(parseFloat(String(row[col]).replace(',', ''))) ? 
                                parseFloat(String(row[col]).replace(',', '')).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : 
                                row[col] || '-'}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
    
            <button className="download-btn" onClick={() => downloadPPT(activeTab)}>
              ⬇️ Download {activeTab} PPT
            </button>
            <button className="download-btn" onClick={downloadCSV}>
              ⬇️ Download Filtered Data as CSV
            </button>
          </div>
        );
      };
    
      return (
        <div className="app-container">
          {isProcessing && (
            <div className="processing-overlay">
              <div className="processing-spinner"></div>
              <p>Processing {processingProgress}%</p>
            </div>
          )}
    
          <header className="app-header">
            <h1>📊 Excel Dashboard - Data Table & Visualizations</h1>
            {connectionStatus === 'failed' && <div className="connection-status error">❌ Backend Disconnected</div>}
            {connectionStatus === 'connected' && <div className="connection-status success">✅ Backend Connected</div>}
          </header>
    
          <div className="main-container">
            <div className="sidebar">
              <div className="sidebar-section">
                <h3>📂 File Upload</h3>
                <input 
                  type="file" 
                  ref={fileInputRef}
                  accept=".xlsx,.xls" 
                  onChange={handleFileUpload} 
                  disabled={isLoading}
                  style={{ display: 'none' }}
                />
                <button 
                  className="upload-btn"
                  onClick={() => fileInputRef.current.click()}
                  disabled={isLoading}
                >
                  {uploadedFile ? 'Change File' : 'Upload Excel File'}
                </button>
                {uploadedFile && (
                  <div className="file-info">
                    <p>{uploadedFile.name}</p>
                    <p>{Math.round(uploadedFile.size / 1024)} KB</p>
                    {sheetType && <p>Sheet Type: {sheetType}</p>}
                  </div>
                )}
              </div>
    
              {sheetNames.length > 0 && (
                <div className="sidebar-section">
                  <h3>📄 Select Sheet</h3>
                  <select 
                    value={selectedSheet} 
                    onChange={(e) => setSelectedSheet(e.target.value)}
                    disabled={isLoading}
                    className="filter-select"
                  >
                    {sheetNames.map((sheet, idx) => (
                      <option key={idx} value={sheet}>{sheet}</option>
                    ))}
                  </select>
                </div>
              )}
    
              {tableOptions.length > 0 && (
                <div className="sidebar-section">
                  <h3>📌 Select Table</h3>
                  <div className="table-radio-group">
                    {tableOptions.map(option => (
                      <label key={option} className={isTableLoading ? 'disabled' : ''}>
                        <input
                          type="radio"
                          value={option}
                          checked={tableChoice === option}
                          onChange={() => setTableChoice(option)}
                          disabled={isTableLoading}
                        />
                        {option}
                        {isTableLoading && tableChoice === option && (
                          <span className="table-loading"> (Loading...)</span>
                        )}
                      </label>
                    ))}
                  </div>
                  {error && error.includes('Selected table') && (
                    <p className="error-message">
                      {error}. Please select a different table or check the Excel file for the correct table headers.
                    </p>
                  )}
                  {tableOptions.length === 0 && (
                    <p className="info-text">
                      No table options available in "{selectedSheet}". Using default table.
                    </p>
                  )}
                </div>
              )}
    
              {tableData && (
                <>
                  <div className="sidebar-section">
                    <h3>📅 Filters</h3>
                    <div className="filter-container">
                      {months.length > 1 && (
                        <div className="filter-item">
                          <label htmlFor="month-filter">Month:</label>
                          <select
                            id="month-filter"
                            value={selectedMonth} 
                            onChange={(e) => setSelectedMonth(e.target.value)}
                            className="filter-select"
                          >
                            {months.map((month, idx) => (
                              <option key={idx} value={month}>{month}</option>
                            ))}
                          </select>
                        </div>
                      )}
                      
                      {years.length > 1 && (
                        <div className="filter-item">
                          <label htmlFor="year-filter">Year:</label>
                          <select
                            id="year-filter"
                            value={selectedYear} 
                            onChange={(e) => setSelectedYear(e.target.value)}
                            className="filter-select"
                          >
                            {years.map((year, idx) => (
                              <option key={idx} value={year}>{year}</option>
                            ))}
                          </select>
                        </div>
                      )}
                      
                      {isBranchAnalysis && branches.length > 1 && (
                        <div className="filter-item">
                          <label htmlFor="branch-filter">Branch:</label>
                          <select
                            id="branch-filter"
                            value={selectedBranch} 
                            onChange={(e) => setSelectedBranch(e.target.value)}
                            className="filter-select"
                          >
                            {branches.map((branch, idx) => (
                              <option key={idx} value={branch}>{branch}</option>
                            ))}
                          </select>
                        </div>
                      )}
                      
                      {isProductAnalysis && products.length > 1 && (
                        <div className="filter-item">
                          <label htmlFor="product-filter">Product:</label>
                          <select
                            id="product-filter"
                            value={selectedProduct} 
                            onChange={(e) => setSelectedProduct(e.target.value)}
                            className="filter-select"
                          >
                            {products.map((product, idx) => (
                              <option key={idx} value={product}>{product}</option>
                            ))}
                          </select>
                        </div>
                      )}
                    </div>
                  </div>
    
                  {!isSalesAnalysisMonthwise && (
                    <>
                      <div className="sidebar-section">
                        <h3>📊 Visualization Type</h3>
                        <div className="filter-item">
                          <label htmlFor="visual-type">Chart Type:</label>
                          <select
                            id="visual-type"
                            value={visualType} 
                            onChange={(e) => setVisualType(e.target.value)}
                            className="filter-select"
                          >
                            <option value="Bar Chart">Bar Chart</option>
                            <option value="Line Chart">Line Chart</option>
                            <option value="Pie Chart">Pie Chart</option>
                          </select>
                        </div>
                      </div>
    
                      <div className="sidebar-section">
                        <h3>📊 Download Options</h3>
                        <button 
                          className="download-btn master"
                          onClick={downloadMasterPPT}
                          disabled={isLoading}
                        >
                          🔄 Generate Master PPT
                        </button>
                        <p className="info-text">💡 All PPT downloads contain ONLY clean charts (no values, no tables, straight x-axis labels)</p>
                      </div>
                    </>
                  )}
                </>
              )}
            </div>
    
            <div className="content-area">
              {isLoading ? (
                <div className="loading-spinner">
                  <div className="spinner"></div>
                  <p>Processing data...</p>
                </div>
              ) : error ? (
                <div className="error-message">
                  <p>❌ Error: {error}</p>
                  <button onClick={() => window.location.reload()}>Try Again</button>
                </div>
              ) : !tableData ? (
                <div className="welcome-message">
                  <p>Please upload an Excel file to begin analysis</p>
                  <button 
                    className="upload-btn large"
                    onClick={() => fileInputRef.current.click()}
                  >
                    Upload Excel File
                  </button>
                </div>
              ) : (
                <>
                  <div className="tabs">
                    {isSalesAnalysisMonthwise ? (
                      <div className="single-tab">
                        <h3>📊 Sales Month-wise Data</h3>
                      </div>
                    ) : (
                      Object.keys(visualizationData).map((tabName) => (
                        <button
                          key={tabName}
                          className={`tab-btn ${activeTab === tabName ? 'active' : ''}`}
                          onClick={() => setActiveTab(tabName)}
                        >
                          {tabName.includes('Budget') ? '📊 ' : 
                           tabName.includes('Branch') ? '🌍 ' : 
                           tabName.includes('Product') ? '📦 ' : ''}
                          {tabName}
                        </button>
                      ))
                    )}
                  </div>
                  {renderTabContent()}
                </>
              )}
            </div>
          </div>
        </div>
      );
    }

export default Dashboard;
