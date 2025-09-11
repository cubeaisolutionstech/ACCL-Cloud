import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import { addReportToStorage } from '../utils/consolidatedStorage';

const ExecutiveODC = () => {
  const { selectedFiles } = useExcelData();

  // Sheet configurations
  const [osJanSheet, setOsJanSheet] = useState('Sheet1');
  const [osFebSheet, setOsFebSheet] = useState('Sheet1');
  const [salesSheet, setSalesSheet] = useState('Sheet1');
  const [osJanSheets, setOsJanSheets] = useState([]);
  const [osFebSheets, setOsFebSheets] = useState([]);
  const [salesSheets, setSalesSheets] = useState([]);
  const [osJanHeader, setOsJanHeader] = useState(1);
  const [osFebHeader, setOsFebHeader] = useState(1);
  const [salesHeader, setSalesHeader] = useState(1);

  // Column configurations
  const [osJanColumns, setOsJanColumns] = useState([]);
  const [osFebColumns, setOsFebColumns] = useState([]);
  const [salesColumns, setSalesColumns] = useState([]);
  const [autoMap, setAutoMap] = useState({});

  // Filters and options
  const [monthOptions, setMonthOptions] = useState([]);
  const [executiveOptions, setExecutiveOptions] = useState([]);
  const [branchOptions, setBranchOptions] = useState([]);
  const [filters, setFilters] = useState({
    selectedMonth: '',
    selectedExecutives: [],
    selectedBranches: []
  });

  // State management
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [downloadingPpt, setDownloadingPpt] = useState(false);
  const [error, setError] = useState(null);

  // Column selections
  const [columnSelections, setColumnSelections] = useState({
    osJan: {},
    osFeb: {},
    sales: {}
  });

  // Fetch available sheet names from backend
  const fetchSheets = async (filename, setter) => {
    try {
      const res = await axios.post('http://localhost:5000/api/branch/sheets', { filename });
      setter(res.data.sheets);
    } catch (error) {
      console.error('Error fetching sheets:', error);
      setError('Failed to load sheet names');
    }
  };

  useEffect(() => {
    if (selectedFiles.osPrevFile) fetchSheets(selectedFiles.osPrevFile, setOsJanSheets);
    if (selectedFiles.osCurrFile) fetchSheets(selectedFiles.osCurrFile, setOsFebSheets);
    if (selectedFiles.salesFile) fetchSheets(selectedFiles.salesFile, setSalesSheets);
  }, [selectedFiles]);

  // Fetch columns after user selects sheet + header
  const fetchColumns = async () => {
    if (!osJanSheet || !osFebSheet || !salesSheet) return;

    setLoading(true);
    setError(null);

    try {
      const getCols = async (filename, sheet_name, header) => {
        const res = await axios.post('http://localhost:5000/api/branch/get_columns', {
          filename,
          sheet_name,
          header
        });
        return res.data.columns || [];
      };

      const [osJanCols, osFebCols, salesCols] = await Promise.all([
        getCols(selectedFiles.osPrevFile, osJanSheet, osJanHeader),
        getCols(selectedFiles.osCurrFile, osFebSheet, osFebHeader),
        getCols(selectedFiles.salesFile, salesSheet, salesHeader)
      ]);

      setOsJanColumns(osJanCols);
      setOsFebColumns(osFebCols);
      setSalesColumns(salesCols);

      // Use OD auto-mapping endpoint
      const res = await axios.post('http://localhost:5000/api/executive/od_auto_map_columns', {
        os_jan_file_path: `uploads/${selectedFiles.osPrevFile}`,
        os_feb_file_path: `uploads/${selectedFiles.osCurrFile}`,
        sales_file_path: `uploads/${selectedFiles.salesFile}`
      });

      const mapping = res.data;

      console.log('OD auto-mapping response:', mapping);

      if (mapping?.os_jan_mapping && mapping?.os_feb_mapping && mapping?.sales_mapping) {
        setAutoMap(mapping);
        setColumnSelections({
          osJan: { ...mapping.os_jan_mapping },
          osFeb: { ...mapping.os_feb_mapping },
          sales: { ...mapping.sales_mapping }
        });

        await fetchExecAndBranches();
        await fetchMonths();
      } else {
        console.warn("❌ OD auto-mapping missing required fields:", mapping);
        setError("Auto-mapping failed: check if required columns were found.");
      }
    } catch (error) {
      console.error('Error in fetchColumns:', error);
      setError('Failed to load columns and auto-mapping');
    } finally {
      setLoading(false);
    }
  };

  const fetchExecAndBranches = async () => {
    try {
      const res = await axios.post('http://localhost:5000/api/executive/od_get_exec_branch_options', {
        os_jan_file_path: `uploads/${selectedFiles.osPrevFile}`,
        os_feb_file_path: `uploads/${selectedFiles.osCurrFile}`,
        sales_file_path: `uploads/${selectedFiles.salesFile}`
      });

      console.log('OD get_exec_branch_options response:', res.data);

      // Handle executives
      const executives = res.data.executives || [];
      setExecutiveOptions(executives);
      
      // Handle branches
      const branches = res.data.branches || [];
      setBranchOptions(branches);
      
      // Select all executives and branches by default
      setFilters(prev => ({
        ...prev,
        selectedExecutives: executives,
        selectedBranches: branches
      }));
      
      console.log('Processed OD executives and branches:', {
        executives,
        branches
      });
    } catch (error) {
      console.error('Error fetching OD executives and branches:', error);
      setError('Failed to load executives and branches');
    }
  };

  const fetchMonths = async () => {
    try {
      const res = await axios.post('http://localhost:5000/api/executive/od_get_available_months', {
        os_jan_file_path: `uploads/${selectedFiles.osPrevFile}`,
        os_feb_file_path: `uploads/${selectedFiles.osCurrFile}`,
        sales_file_path: `uploads/${selectedFiles.salesFile}`
      });

      console.log('Available OD months response:', res.data);
      
      const months = res.data.available_months || [];
      setMonthOptions(months);
      
      // Select the latest month by default
      if (months.length > 0) {
        setFilters(prev => ({
          ...prev,
          selectedMonth: months[months.length - 1]
        }));
      }
    } catch (error) {
      console.error('Error fetching OD months:', error);
      setError('Failed to load available months');
    }
  };

  // Function to add OD vs Collection results to consolidated storage
  const addODVsCollectionReportsToStorage = (resultsData) => {
    try {
      const odVsCollectionReports = [{
        df: resultsData.od_results || [],
        title: `OD TARGET vs COLLECTION - ${filters.selectedMonth}`,
        percent_cols: [3, 6] // Overall % Achieved and % Achieved (Selected Month)
      }];

      if (odVsCollectionReports.length > 0) {
        addReportToStorage(odVsCollectionReports, 'od_vs_results');
        console.log(`✅ Added ${odVsCollectionReports.length} OD vs Collection reports to consolidated storage`);
      }
    } catch (error) {
      console.error('Error adding OD vs Collection reports to consolidated storage:', error);
    }
  };

  const handleCalculate = async () => {
    if (!selectedFiles.osPrevFile || !selectedFiles.osCurrFile || !selectedFiles.salesFile) {
      setError('Please upload all three files (OS Jan, OS Feb, and Sales)');
      return;
    }

    // Validate required columns
    const requiredOsJanColumns = ['due_date', 'ref_date', 'net_value', 'executive', 'sl_code', 'area'];
    const requiredOsFebColumns = ['due_date', 'ref_date', 'net_value', 'executive', 'sl_code', 'area'];
    const requiredSalesColumns = ['bill_date', 'due_date', 'value', 'executive', 'sl_code', 'area'];
    
    const missingOsJanColumns = requiredOsJanColumns.filter(col => 
      !columnSelections.osJan[col] || columnSelections.osJan[col] === ''
    );
    const missingOsFebColumns = requiredOsFebColumns.filter(col => 
      !columnSelections.osFeb[col] || columnSelections.osFeb[col] === ''
    );
    const missingSalesColumns = requiredSalesColumns.filter(col => 
      !columnSelections.sales[col] || columnSelections.sales[col] === ''
    );
    
    if (missingOsJanColumns.length > 0) {
      setError(`Missing required OS Jan column mappings: ${missingOsJanColumns.join(', ')}`);
      return;
    }
    
    if (missingOsFebColumns.length > 0) {
      setError(`Missing required OS Feb column mappings: ${missingOsFebColumns.join(', ')}`);
      return;
    }
    
    if (missingSalesColumns.length > 0) {
      setError(`Missing required Sales column mappings: ${missingSalesColumns.join(', ')}`);
      return;
    }

    if (!filters.selectedMonth) {
      setError('Please select a month');
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const payload = {
        os_jan_file_path: `uploads/${selectedFiles.osPrevFile}`,
        os_feb_file_path: `uploads/${selectedFiles.osCurrFile}`,
        sales_file_path: `uploads/${selectedFiles.salesFile}`,
        
        // OS Jan column mappings
        os_jan_due_date: columnSelections.osJan.due_date,
        os_jan_ref_date: columnSelections.osJan.ref_date,
        os_jan_net_value: columnSelections.osJan.net_value,
        os_jan_executive: columnSelections.osJan.executive,
        os_jan_sl_code: columnSelections.osJan.sl_code,
        os_jan_area: columnSelections.osJan.area,
        
        // OS Feb column mappings
        os_feb_due_date: columnSelections.osFeb.due_date,
        os_feb_ref_date: columnSelections.osFeb.ref_date,
        os_feb_net_value: columnSelections.osFeb.net_value,
        os_feb_executive: columnSelections.osFeb.executive,
        os_feb_sl_code: columnSelections.osFeb.sl_code,
        os_feb_area: columnSelections.osFeb.area,
        
        // Sales column mappings
        sales_bill_date: columnSelections.sales.bill_date,
        sales_due_date: columnSelections.sales.due_date,
        sales_value: columnSelections.sales.value,
        sales_executive: columnSelections.sales.executive,
        sales_sl_code: columnSelections.sales.sl_code,
        sales_area: columnSelections.sales.area,
        
        // Filters
        selected_executives: filters.selectedExecutives,
        selected_month: filters.selectedMonth,
        selected_branches: filters.selectedBranches
      };

      console.log('OD vs Collection calculate payload:', payload);

      const res = await axios.post('http://localhost:5000/api/executive/calculate_od_vs_collection', payload);
      
      if (res.data.success) {
        setResults(res.data);
        setError(null);
        console.log('OD vs Collection calculation successful!');
        
        // 🎯 ADD TO CONSOLIDATED REPORTS - Just like Streamlit!
        addODVsCollectionReportsToStorage(res.data);
        
      } else {
        setError(res.data.error || 'Failed to calculate results');
      }
    } catch (error) {
      console.error('Error calculating OD vs Collection:', error);
      if (error.response?.data?.error) {
        setError(error.response.data.error);
      } else {
        setError('Error calculating results. Please check console for details.');
      }
    } finally {
      setLoading(false);
    }
  };

  const handleDownloadPpt = async () => {
    if (!results) {
      setError('No results available for PPT generation');
      return;
    }

    setDownloadingPpt(true);
    setError(null);

    try {
      const payload = {
        results_data: {
          od_results: results.od_results
        },
        month_title: filters.selectedMonth,
        logo_file: null // Add logo support later if needed
      };

      console.log('OD PPT generation payload:', payload);

      const response = await axios.post('http://localhost:5000/api/executive/generate_od_ppt', payload, {
        responseType: 'blob',
        headers: {
          'Content-Type': 'application/json',
        },
      });

      // Create blob and download
      const blob = new Blob([response.data], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      });
      
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `Executive_OD_Target_vs_Collection_${filters.selectedMonth.replace(/[^a-zA-Z0-9]/g, '_')}.pptx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      console.log('OD PPT downloaded successfully');
    } catch (error) {
      console.error('Error downloading OD PPT:', error);
      if (error.response?.data?.error) {
        setError(error.response.data.error);
      } else {
        setError('Failed to download PowerPoint presentation');
      }
    } finally {
      setDownloadingPpt(false);
    }
  };

  return (
    <div className="p-6">
      <h2 className="text-2xl font-bold text-blue-800 mb-6">Executive OD Target vs Collection Analysis</h2>

      {error && (
        <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
          {error}
        </div>
      )}

      {/* Sheet Selection */}
      <div className="bg-white p-4 rounded-lg shadow mb-6">
        <h3 className="text-lg font-semibold text-blue-700 mb-4">Sheet Configuration</h3>
        <div className="grid grid-cols-3 gap-6">
          <div>
            <label className="block font-semibold mb-2">OS-Previous Month Sheet</label>
            <select 
              className="w-full p-2 border border-gray-300 rounded" 
              value={osJanSheet} 
              onChange={e => setOsJanSheet(e.target.value)}
            >
              <option value="">Select Sheet</option>
              {osJanSheets.map(sheet => <option key={sheet} value={sheet}>{sheet}</option>)}
            </select>

            <label className="block mt-4 font-semibold mb-2">OS-Previous Header Row</label>
            <input 
              type="number" 
              className="w-full p-2 border border-gray-300 rounded" 
              min={1} 
              value={osJanHeader}
              onChange={e => setOsJanHeader(Number(e.target.value))} 
            />
          </div>

          <div>
            <label className="block font-semibold mb-2">OS-Current Month Sheet</label>
            <select 
              className="w-full p-2 border border-gray-300 rounded" 
              value={osFebSheet} 
              onChange={e => setOsFebSheet(e.target.value)}
            >
              <option value="">Select Sheet</option>
              {osFebSheets.map(sheet => <option key={sheet} value={sheet}>{sheet}</option>)}
            </select>

            <label className="block mt-4 font-semibold mb-2">OS-Current Header Row</label>
            <input 
              type="number" 
              className="w-full p-2 border border-gray-300 rounded" 
              min={1} 
              value={osFebHeader}
              onChange={e => setOsFebHeader(Number(e.target.value))} 
            />
          </div>

          <div>
            <label className="block font-semibold mb-2">Sales Sheet</label>
            <select 
              className="w-full p-2 border border-gray-300 rounded" 
              value={salesSheet} 
              onChange={e => setSalesSheet(e.target.value)}
            >
              <option value="">Select Sheet</option>
              {salesSheets.map(sheet => <option key={sheet} value={sheet}>{sheet}</option>)}
            </select>

            <label className="block mt-4 font-semibold mb-2">Sales Header Row</label>
            <input 
              type="number" 
              className="w-full p-2 border border-gray-300 rounded" 
              min={1} 
              value={salesHeader}
              onChange={e => setSalesHeader(Number(e.target.value))} 
            />
          </div>
        </div>

        <button
          className="bg-blue-600 text-white px-4 py-2 rounded disabled:bg-gray-400" 
          onClick={fetchColumns}
          disabled={!osJanSheet || !osFebSheet || !salesSheet || loading}
        >
          {loading ? 'Loading...' : 'Load Columns & Auto-Map'}
        </button>
      </div>

      {/* Column Mapping */}
      {osJanColumns.length > 0 && osFebColumns.length > 0 && salesColumns.length > 0 && (
        <div className="bg-white p-4 rounded-lg shadow mb-6">
          <h3 className="text-lg font-semibold text-blue-700 mb-4">Column Mapping</h3>
          
          {/* OS Jan Columns */}
          <div className="mb-6">
            <h4 className="text-md font-semibold text-gray-700 mb-3">OS-Previous Month Columns</h4>
            <div className="grid grid-cols-3 gap-4">
              {[
                { key: 'due_date', label: 'Due Date *', required: true },
                { key: 'ref_date', label: 'Reference Date *', required: true },
                { key: 'net_value', label: 'Net Value *', required: true },
                { key: 'executive', label: 'Executive *', required: true },
                { key: 'sl_code', label: 'SL Code *', required: true },
                { key: 'area', label: 'Branch/Area *', required: true }
              ].map(({ key, label, required }) => (
                <div key={key}>
                  <label className="block font-medium mb-1">
                    {label} {required && <span className="text-red-500">*</span>}
                  </label>
                  <select
                    className={`w-full p-2 border rounded ${required && !columnSelections.osJan[key] ? 'border-red-300' : 'border-gray-300'}`}
                    value={columnSelections.osJan[key] || ''}
                    onChange={(e) =>
                      setColumnSelections(prev => ({
                        ...prev,
                        osJan: { ...prev.osJan, [key]: e.target.value }
                      }))
                    }
                  >
                    <option value="">Select Column</option>
                    {osJanColumns.map(col => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>
          </div>

          {/* OS Feb Columns */}
          <div className="mb-6">
            <h4 className="text-md font-semibold text-gray-700 mb-3">OS-Current Month Columns</h4>
            <div className="grid grid-cols-3 gap-4">
              {[
                { key: 'due_date', label: 'Due Date *', required: true },
                { key: 'ref_date', label: 'Reference Date *', required: true },
                { key: 'net_value', label: 'Net Value *', required: true },
                { key: 'executive', label: 'Executive *', required: true },
                { key: 'sl_code', label: 'SL Code *', required: true },
                { key: 'area', label: 'Branch/Area *', required: true }
              ].map(({ key, label, required }) => (
                <div key={key}>
                  <label className="block font-medium mb-1">
                    {label} {required && <span className="text-red-500">*</span>}
                  </label>
                  <select
                    className={`w-full p-2 border rounded ${required && !columnSelections.osFeb[key] ? 'border-red-300' : 'border-gray-300'}`}
                    value={columnSelections.osFeb[key] || ''}
                    onChange={(e) =>
                      setColumnSelections(prev => ({
                        ...prev,
                        osFeb: { ...prev.osFeb, [key]: e.target.value }
                      }))
                    }
                  >
                    <option value="">Select Column</option>
                    {osFebColumns.map(col => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>
          </div>

          {/* Sales Columns */}
          <div>
            <h4 className="text-md font-semibold text-gray-700 mb-3">Sales Columns</h4>
            <div className="grid grid-cols-3 gap-4">
              {[
                { key: 'bill_date', label: 'Bill Date *', required: true },
                { key: 'due_date', label: 'Due Date *', required: true },
                { key: 'value', label: 'Value *', required: true },
                { key: 'executive', label: 'Executive *', required: true },
                { key: 'sl_code', label: 'SL Code *', required: true },
                { key: 'area', label: 'Branch/Area *', required: true }
              ].map(({ key, label, required }) => (
                <div key={key}>
                  <label className="block font-medium mb-1">
                    {label} {required && <span className="text-red-500">*</span>}
                  </label>
                  <select
                    className={`w-full p-2 border rounded ${required && !columnSelections.sales[key] ? 'border-red-300' : 'border-gray-300'}`}
                    value={columnSelections.sales[key] || ''}
                    onChange={(e) =>
                      setColumnSelections(prev => ({
                        ...prev,
                        sales: { ...prev.sales, [key]: e.target.value }
                      }))
                    }
                  >
                    <option value="">Select Column</option>
                    {salesColumns.map(col => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Filters */}
      {(executiveOptions.length > 0 || branchOptions.length > 0 || monthOptions.length > 0) && (
        <div className="bg-white p-4 rounded-lg shadow mb-6">
          <h3 className="text-lg font-semibold text-blue-700 mb-4">Filters</h3>
          
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {/* Month Filter */}
            {monthOptions.length > 0 && (
              <div>
                <label className="block font-semibold mb-3">
                  Select Month
                </label>
                <select
                  className="w-full p-2 border border-gray-300 rounded"
                  value={filters.selectedMonth}
                  onChange={(e) => {
                    setFilters(prev => ({ ...prev, selectedMonth: e.target.value }));
                  }}
                >
                  <option value="">Select Month</option>
                  {monthOptions.map(month => (
                    <option key={month} value={month}>{month}</option>
                  ))}
                </select>
              </div>
            )}

            {/* Executive Filter */}
            {executiveOptions.length > 0 && (
              <div>
                <label className="block font-semibold mb-3">
                  Executives ({filters.selectedExecutives.length} of {executiveOptions.length})
                </label>
                <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
                  <label className="flex items-center mb-2">
                    <input
                      type="checkbox"
                      checked={filters.selectedExecutives.length === executiveOptions.length}
                      onChange={(e) => {
                        if (e.target.checked) {
                          setFilters(prev => ({ ...prev, selectedExecutives: executiveOptions }));
                        } else {
                          setFilters(prev => ({ ...prev, selectedExecutives: [] }));
                        }
                      }}
                      className="mr-3"
                    />
                    <span className="font-medium text-xm">Select All</span>
                  </label>
                  {executiveOptions.map(exec => (
                    <label key={exec} className="flex items-center mb-2">
                      <input
                        type="checkbox"
                        checked={filters.selectedExecutives.includes(exec)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setFilters(prev => ({
                              ...prev,
                              selectedExecutives: [...prev.selectedExecutives, exec]
                            }));
                          } else {
                            setFilters(prev => ({
                              ...prev,
                              selectedExecutives: prev.selectedExecutives.filter(e => e !== exec)
                            }));
                          }
                        }}
                        className="mr-3"
                      />
                      <span className="text-xm">{exec}</span>
                    </label>
                  ))}
                </div>
              </div>
            )}

            {/* Branch Filter */}
            {branchOptions.length > 0 && (
              <div>
                <label className="block font-semibold mb-3">
                  Branches ({filters.selectedBranches.length} of {branchOptions.length})
                </label>
                <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
                  <label className="flex items-center mb-3">
                    <input
                      type="checkbox"
                      checked={filters.selectedBranches.length === branchOptions.length}
                      onChange={(e) => {
                        if (e.target.checked) {
                          setFilters(prev => ({ ...prev, selectedBranches: branchOptions }));
                        } else {
                          setFilters(prev => ({ ...prev, selectedBranches: [] }));
                        }
                      }}
                      className="mr-3"
                    />
                    <span className="font-medium text-xm">Select All</span>
                  </label>
                  {branchOptions.map(branch => (
                    <label key={branch} className="flex items-center mb-2">
                      <input
                        type="checkbox"
                        checked={filters.selectedBranches.includes(branch)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setFilters(prev => ({
                              ...prev,
                              selectedBranches: [...prev.selectedBranches, branch]
                            }));
                          } else {
                            setFilters(prev => ({
                              ...prev,
                              selectedBranches: prev.selectedBranches.filter(b => b !== branch)
                            }));
                          }
                        }}
                        className="mr-3"
                      />
                      <span className="text-xm">{branch}</span>
                    </label>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Calculate Button */}
      <div className="text-center mb-6">
        <button
          onClick={handleCalculate}
          disabled={loading || !osJanColumns.length || !osFebColumns.length || !salesColumns.length}
          className="bg-red-600 text-white px-4 py-2 rounded disabled:bg-gray-400"
        >
          {loading ? 'Calculating...' : 'Calculate OD Target vs Collection'}
        </button>
      </div>

      {/* Results */}
      {results && (
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-xl font-bold text-purple-700 mb-4">Results</h3>

          {/* Success Message */}
          <div className="bg-green-100 border border-green-400 text-green-700 px-4 py-3 rounded mb-4">
            ✅ Results calculated and automatically added to consolidated reports!
          </div>

          {/* Download PPT Button */}
          <div className="text-center mb-6">
            <button
              onClick={handleDownloadPpt}
              disabled={downloadingPpt || !results}
              className="bg-green-600 text-white px-5 py-2 rounded hover:bg-green-700"
            >
              {downloadingPpt ? 'Generating PPT...' : 'Download PowerPoint'}
            </button>
          </div>

          {/* Data Table */}
          {results.od_results && results.od_results.length > 0 && (
            <div className="mb-8">
              <h4 className="text-lg font-semibold text-gray-800 mb-3">
                OD Targetvs Collection - {filters.selectedMonth} (Value in Lakhs)
             </h4>
             <div className="overflow-x-auto">
               <table className="min-w-full table-auto border-collapse border border-gray-300">
                 <thead>
                   <tr className="bg-purple-600 text-white">
                     {[
                       'Executive',
                       'Due Target',
                       'Collection Achieved',
                       'Overall % Achieved',
                       'For the month Overdue',
                       'For the month Collection',
                       'For the month % Achieved'
                     ].map(col => (
                       <th key={col} className="border border-gray-300 px-4 py-2 text-left font-semibold">
                         {col}
                       </th>
                     ))}
                   </tr>
                 </thead>
                 <tbody>
                   {results.od_results.map((row, i) => (
                     <tr 
                       key={i} 
                       className={`
                         ${row.Executive === 'TOTAL' 
                           ? 'bg-gray-200 font-bold border-t-2 border-gray-400' 
                           : i % 2 === 0 
                             ? 'bg-gray-50' 
                             : 'bg-white'
                         } hover:bg-purple-50
                       `}
                     >
                       <td className="border border-gray-300 px-4 py-2">{row.Executive}</td>
                       <td className="border border-gray-300 px-4 py-2">
                         {typeof row['Due Target'] === 'number' ? row['Due Target'].toLocaleString() : row['Due Target']}
                       </td>
                       <td className="border border-gray-300 px-4 py-2">
                         {typeof row['Collection Achieved'] === 'number' ? row['Collection Achieved'].toLocaleString() : row['Collection Achieved']}
                       </td>
                       <td className="border border-gray-300 px-4 py-2">
                         {typeof row['Overall % Achieved'] === 'number' ? `${row['Overall % Achieved'].toFixed(2)}%` : row['Overall % Achieved']}
                       </td>
                       <td className="border border-gray-300 px-4 py-2">
                         {typeof row['For the month Overdue'] === 'number' ? row['For the month Overdue'].toLocaleString() : row['For the month Overdue']}
                       </td>
                       <td className="border border-gray-300 px-4 py-2">
                         {typeof row['For the month Collection'] === 'number' ? row['For the month Collection'].toLocaleString() : row['For the month Collection']}
                       </td>
                       <td className="border border-gray-300 px-4 py-2">
                         {typeof row['% Achieved (Selected Month)'] === 'number' ? `${row['% Achieved (Selected Month)'].toFixed(2)}%` : row['For the month % Achieved']}
                       </td>
                     </tr>
                   ))}
                 </tbody>
               </table>
             </div>
           </div>
         )}
       </div>
     )}
   </div>
 );
};

export default ExecutiveODC;
