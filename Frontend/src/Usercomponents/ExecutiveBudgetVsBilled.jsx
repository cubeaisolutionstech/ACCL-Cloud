import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
// Import from your consolidated storage module
import { addReportToStorage } from '../utils/consolidatedStorage';

const ExecutiveBudgetVsBilled = () => {
  const { selectedFiles } = useExcelData();

  const [salesSheet, setSalesSheet] = useState('Sheet1');
  const [targetSheet, setTargetSheet] = useState('Sheet1');
  const [salesSheets, setSalesSheets] = useState([]);
  const [targetSheets, setTargetSheets] = useState([]);
  const [salesHeader, setSalesHeader] = useState(1);
  const [targetHeader, setTargetHeader] = useState(1);
  const [salesColumns, setSalesColumns] = useState([]);
  const [targetColumns, setTargetColumns] = useState([]);
  const [autoMap, setAutoMap] = useState({});
  const [monthOptions, setMonthOptions] = useState([]);
  const [executiveOptions, setExecutiveOptions] = useState([]);
  const [branchOptions, setBranchOptions] = useState([]);
  const [filters, setFilters] = useState({
    selectedMonths: [],
    selectedExecutives: [],
    selectedBranches: []
  });
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [downloadingPpt, setDownloadingPpt] = useState(false);
  const [downloadingProof, setDownloadingProof] = useState(false);
  const [error, setError] = useState(null);

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
    if (selectedFiles.salesFile) fetchSheets(selectedFiles.salesFile, setSalesSheets);
    if (selectedFiles.budgetFile) fetchSheets(selectedFiles.budgetFile, setTargetSheets);
  }, [selectedFiles]);

  const [columnSelections, setColumnSelections] = useState({
    sales: {},
    budget: {}
  });

  // Fetch columns after user selects sheet + header
  const fetchColumns = async () => {
    if (!salesSheet || !targetSheet) return;

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

      const [salesCols, targetCols] = await Promise.all([
        getCols(selectedFiles.salesFile, salesSheet, salesHeader),
        getCols(selectedFiles.budgetFile, targetSheet, targetHeader)
      ]);

      setSalesColumns(salesCols);
      setTargetColumns(targetCols);

      // Use executive auto-mapping endpoint
      const res = await axios.post('http://localhost:5000/api/executive/auto_map_columns', {
        sales_file_path: `uploads/${selectedFiles.salesFile}`,
        budget_file_path: `uploads/${selectedFiles.budgetFile}`
      });

      const mapping = res.data;

      console.log('Executive auto-mapping response:', mapping);

      if (mapping?.sales_mapping && mapping?.budget_mapping) {
        setAutoMap(mapping);
        setColumnSelections({
          sales: { ...mapping.sales_mapping },
          budget: { ...mapping.budget_mapping }
        });

        await fetchExecAndBranches(mapping);
        await fetchMonths(mapping.sales_mapping.date);
      } else {
        console.warn("❌ Executive auto-mapping missing required fields:", mapping);
        setError("Auto-mapping failed: check if required columns were found.");
      }
    } catch (error) {
      console.error('Error in fetchColumns:', error);
      setError('Failed to load columns and auto-mapping');
    } finally {
      setLoading(false);
    }
  };

  const fetchExecAndBranches = async (autoMapData) => {
    try {
      const res = await axios.post('http://localhost:5000/api/executive/get_exec_branch_options', {
        sales_file_path: `uploads/${selectedFiles.salesFile}`,
        budget_file_path: `uploads/${selectedFiles.budgetFile}`
      });

      console.log('Executive get_exec_branch_options response:', res.data);

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
      
      console.log('Processed executives and branches:', {
        executives,
        branches
      });
    } catch (error) {
      console.error('Error fetching executives and branches:', error);
      setError('Failed to load executives and branches');
    }
  };

  const fetchMonths = async (dateCol) => {
    if (!dateCol) return;

    try {
      const res = await axios.post('http://localhost:5000/api/executive/get_available_months', {
        sales_file_path: `uploads/${selectedFiles.salesFile}`
      });

      console.log('Available months response:', res.data);
      
      const months = res.data.available_months || [];
      setMonthOptions(months);
      
      // Select all months by default
      setFilters(prev => ({
        ...prev,
        selectedMonths: months
      }));
    } catch (error) {
      console.error('Error fetching months:', error);
      setError('Failed to load available months');
    }
  };

  // 🎯 UPDATED: Enhanced function to add results to consolidated storage
  const addToConsolidatedReports = (resultsData) => {
    try {
      const monthTitle = filters.selectedMonths.length > 0 
        ? filters.selectedMonths.join(', ') 
        : 'All Months';

      // 📊 EXACT match to your Streamlit st.session_state.budget_results structure
      const executiveReports = [
        {
          df: resultsData.budget_vs_billed_qty || [],
          title: `TARGET AGAINST BILLED - QUANTITY - ${monthTitle}`,
          percent_cols: [3] // Column index for percentage
        },
        {
          df: resultsData.budget_vs_billed_value || [],
          title: `TARGET AGAINST BILLED - VALUE - ${monthTitle}`,
          percent_cols: [3] // Column index for percentage
        },
        {
          df: resultsData.overall_sales_qty || [],
          title: `OVERALL SALES - QUANTITY - ${monthTitle}`,
          percent_cols: [] // No percentage columns
        },
        {
          df: resultsData.overall_sales_value || [],
          title: `OVERALL SALES - VALUE - ${monthTitle}`,
          percent_cols: [] // No percentage columns
        }
      ];

      // 🎯 Use the consolidated storage function with validation
      addReportToStorage(executiveReports, 'budget_results');
      
      console.log(`✅ Added ${executiveReports.length} executive reports to consolidated storage with category: budget_results`);
    } catch (error) {
      console.error('❌ Error adding executive reports to consolidated storage:', error);
    }
  };

  const handleCalculate = async () => {
    if (!selectedFiles.salesFile || !selectedFiles.budgetFile) {
      setError('Please upload both sales and target files');
      return;
    }

    // Validate required columns
    const requiredSalesColumns = ['date', 'value', 'qty', 'exec'];
    const requiredTargetColumns = ['value', 'qty', 'exec'];
    
    const missingSalesColumns = requiredSalesColumns.filter(col => 
      !columnSelections.sales[col] || columnSelections.sales[col] === ''
    );
    const missingTargetColumns = requiredTargetColumns.filter(col => 
      !columnSelections.budget[col] || columnSelections.budget[col] === ''
    );
    
    if (missingSalesColumns.length > 0) {
      setError(`Missing required sales column mappings: ${missingSalesColumns.join(', ')}`);
      return;
    }
    
    if (missingTargetColumns.length > 0) {
      setError(`Missing required target column mappings: ${missingTargetColumns.join(', ')}`);
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const payload = {
        sales_file_path: `uploads/${selectedFiles.salesFile}`,
        budget_file_path: `uploads/${selectedFiles.budgetFile}`,
        
        // Sales column mappings
        sales_date: columnSelections.sales.date,
        sales_value: columnSelections.sales.value,
        sales_quantity: columnSelections.sales.qty,
        sales_product_group: columnSelections.sales.product_group,
        sales_sl_code: columnSelections.sales.sl_code,
        sales_executive: columnSelections.sales.exec,
        sales_area: columnSelections.sales.area,
        
        // Budget column mappings (backend still uses budget naming)
        budget_value: columnSelections.budget.value,
        budget_quantity: columnSelections.budget.qty,
        budget_product_group: columnSelections.budget.product_group,
        budget_sl_code: columnSelections.budget.sl_code,
        budget_executive: columnSelections.budget.exec,
        budget_exec_code: columnSelections.budget.exec_code, // NEW: Executive Code
        budget_area: columnSelections.budget.area,
        
        // Filters
        selected_executives: filters.selectedExecutives,
        selected_months: filters.selectedMonths,
        selected_branches: filters.selectedBranches
      };

      console.log('Executive calculate payload:', payload);

      const res = await axios.post('http://localhost:5000/api/executive/calculate_budget_vs_billed', payload);
      
      if (res.data.success) {
        setResults(res.data);
        setError(null);
        console.log('🎉 Executive calculation successful!');
        
        // 🎯 ADD TO CONSOLIDATED REPORTS - Enhanced with better error handling
        addToConsolidatedReports(res.data);
        
      } else {
        setError(res.data.error || 'Failed to calculate results');
      }
    } catch (error) {
      console.error('Error calculating executive target vs billed:', error);
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
      // Prepare month title from selected months
      const monthTitle = filters.selectedMonths.length > 0 
        ? filters.selectedMonths.join(', ') 
        : 'All Months';

      const payload = {
        results_data: {
          budget_vs_billed_qty: results.budget_vs_billed_qty,
          budget_vs_billed_value: results.budget_vs_billed_value,
          overall_sales_qty: results.overall_sales_qty,
          overall_sales_value: results.overall_sales_value
        },
        month_title: monthTitle,
        logo_file: null // Add logo support later if needed
      };

      console.log('PPT generation payload:', payload);

      const response = await axios.post('http://localhost:5000/api/executive/generate_ppt', payload, {
        responseType: 'blob', // Important for file download
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
      link.download = `Executive_Target_vs_Billed_${monthTitle.replace(/[^a-zA-Z0-9]/g, '_')}.pptx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      console.log('📊 PPT downloaded successfully');
    } catch (error) {
      console.error('Error downloading PPT:', error);
      if (error.response?.data?.error) {
        setError(error.response.data.error);
      } else {
        setError('Failed to download PowerPoint presentation');
      }
    } finally {
      setDownloadingPpt(false);
    }
  };

  // 🔥 UPDATED: Fixed proof of calculation function to match backend API
  const handleDownloadProof = async () => {
    if (!results) {
      setError('No results available for proof of calculation');
      return;
    }

    setDownloadingProof(true);
    setError(null);

    try {
      // Prepare month title from selected months
      const monthTitle = filters.selectedMonths.length > 0 
        ? filters.selectedMonths.join(', ') 
        : 'All Months';

      const payload = {
        sales_file_path: `uploads/${selectedFiles.salesFile}`,
        budget_file_path: `uploads/${selectedFiles.budgetFile}`,
        
        // Sales column mappings
        sales_date: columnSelections.sales.date,
        sales_value: columnSelections.sales.value,
        sales_quantity: columnSelections.sales.qty,
        sales_product_group: columnSelections.sales.product_group,
        sales_sl_code: columnSelections.sales.sl_code,
        sales_executive: columnSelections.sales.exec,
        sales_area: columnSelections.sales.area,
        
        // Budget column mappings (backend still uses budget naming)
        budget_value: columnSelections.budget.value,
        budget_quantity: columnSelections.budget.qty,
        budget_product_group: columnSelections.budget.product_group,
        budget_sl_code: columnSelections.budget.sl_code,
        budget_executive: columnSelections.budget.exec,
        budget_exec_code: columnSelections.budget.exec_code, // NEW: Executive Code
        budget_area: columnSelections.budget.area,
        
        // Filters
        selected_executives: filters.selectedExecutives,
        selected_months: filters.selectedMonths,
        selected_branches: filters.selectedBranches,
        
        // Month title for filename
        month_title: monthTitle
      };

      console.log('Proof of calculation payload:', payload);

      const response = await axios.post('http://localhost:5000/api/executive/generate_proof_of_calculation', payload, {
        responseType: 'blob', // Important for file download
        headers: {
          'Content-Type': 'application/json',
        },
      });

      // Create blob and download
      const blob = new Blob([response.data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      
      // Create clean filename
      const cleanMonthTitle = monthTitle.replace(/[^a-zA-Z0-9]/g, '_');
      const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
      link.download = `Executive_Target_vs_Billed_Proof_${cleanMonthTitle}_${timestamp}.xlsx`;
      
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      console.log('📋 Proof of calculation downloaded successfully');
      
      // Show success message
      setError(null);
      
    } catch (error) {
      console.error('Error downloading proof of calculation:', error);
      if (error.response?.data?.error) {
        setError(`Failed to generate proof of calculation: ${error.response.data.error}`);
      } else if (error.response?.status === 500) {
        setError('Server error while generating proof of calculation. Please check your data and try again.');
      } else {
        setError('Failed to download proof of calculation. Please try again.');
      }
    } finally {
      setDownloadingProof(false);
    }
  };

  return (
    <div className="p-6">
      <h2 className="text-2xl font-bold text-blue-800 mb-6">Executive Target vs Billed Analysis</h2>

      {error && (
        <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
          {error}
        </div>
      )}

      {/* Sheet Selection */}
      <div className="bg-white p-4 rounded-lg shadow mb-6">
        <h3 className="text-lg font-semibold text-blue-700 mb-4">Sheet Configuration</h3>
        <div className="grid grid-cols-2 gap-6">
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

          <div>
            <label className="block font-semibold mb-2">Target Sheet</label>
            <select 
              className="w-full p-2 border border-gray-300 rounded" 
              value={targetSheet} 
              onChange={e => setTargetSheet(e.target.value)}
            >
              <option value="">Select Sheet</option>
              {targetSheets.map(sheet => <option key={sheet} value={sheet}>{sheet}</option>)}
            </select>

            <label className="block mt-4 font-semibold mb-2">Target Header Row</label>
            <input 
              type="number" 
              className="w-full p-2 border border-gray-300 rounded" 
              min={1} 
              value={targetHeader}
              onChange={e => setTargetHeader(Number(e.target.value))} 
            />
          </div>
        </div>

        <button
          className="bg-blue-600 text-white px-6 py-2 rounded hover:bg-blue-700 disabled:bg-gray-400"
          onClick={fetchColumns}
          disabled={!salesSheet || !targetSheet || loading}
        >
          {loading ? 'Loading...' : 'Load Columns & Auto-Map'}
        </button>
      </div>

      {/* Column Mapping */}
      {salesColumns.length > 0 && targetColumns.length > 0 && (
        <div className="bg-white p-4 rounded-lg shadow mb-6">
          <h3 className="text-lg font-semibold text-blue-700 mb-4">Column Mapping</h3>
          
          {/* Sales Columns */}
          <div className="mb-6">
            <h4 className="text-md font-semibold text-gray-700 mb-3">Sales Columns</h4>
            <div className="grid grid-cols-3 gap-4">
              {[
                { key: 'date', label: 'Date *', required: true },
                { key: 'value', label: 'Value *', required: true },
                { key: 'qty', label: 'Quantity *', required: true },
                { key: 'exec', label: 'Executive *', required: true },
                { key: 'area', label: 'Area/Branch', required: false },
                { key: 'product_group', label: 'Product Group', required: false },
                { key: 'sl_code', label: 'SL Code', required: false }
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

          {/* Target Columns */}
          <div>
            <h4 className="text-md font-semibold text-gray-700 mb-3">Target Columns</h4>
            <div className="grid grid-cols-3 gap-4">
              {[
                { key: 'value', label: 'Value *', required: true },
                { key: 'qty', label: 'Quantity *', required: true },
                { key: 'exec', label: 'Executive *', required: true },
                { key: 'exec_code', label: 'Executive Code', required: false }, // NEW: Executive Code
                { key: 'area', label: 'Area/Branch', required: false },
                { key: 'product_group', label: 'Product Group', required: false },
                { key: 'sl_code', label: 'SL Code', required: false }
              ].map(({ key, label, required }) => (
                <div key={key}>
                  <label className="block font-medium mb-1">
                    {label} {required && <span className="text-red-500">*</span>}
                  </label>
                  <select
                    className={`w-full p-2 border rounded ${required && !columnSelections.budget[key] ? 'border-red-300' : 'border-gray-300'}`}
                    value={columnSelections.budget[key] || ''}
                    onChange={(e) =>
                      setColumnSelections(prev => ({
                        ...prev,
                        budget: { ...prev.budget, [key]: e.target.value }
                      }))
                    }
                  >
                    <option value="">Select Column</option>
                    {targetColumns.map(col => (
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
            {/* Executive Filter */}
            {executiveOptions.length > 0 && (
              <div>
                <label className="block font-semibold mb-2">
                  Executives ({filters.selectedExecutives.length} of {executiveOptions.length})
                </label>
                <div className="max-h-60 overflow-y-auto border border-gray-300 rounded p-3">
                  <label className="flex items-center mb-3">
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
                <div className="max-h-60 overflow-y-auto border border-gray-300 rounded p-3">
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

            {/* Month Filter */}
            {monthOptions.length > 0 && (
              <div>
                <label className="block font-semibold mb-3">
                  Months ({filters.selectedMonths.length} of {monthOptions.length})
                </label>
                <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-3">
                  <label className="flex items-center mb-3">
                    <input
                      type="checkbox"
                      checked={filters.selectedMonths.length === monthOptions.length}
                      onChange={(e) => {
                        if (e.target.checked) {
                          setFilters(prev => ({ ...prev, selectedMonths: monthOptions }));
                        } else {
                          setFilters(prev => ({ ...prev, selectedMonths: [] }));
                        }
                      }}
                      className="mr-3"
                    />
                    <span className="font-medium text-xm">Select All</span>
                  </label>
                  {monthOptions.map(month => (
                    <label key={month} className="flex items-center mb-3">
                      <input
                        type="checkbox"
                        checked={filters.selectedMonths.includes(month)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setFilters(prev => ({
                              ...prev,
                              selectedMonths: [...prev.selectedMonths, month]
                            }));
                          } else {
                            setFilters(prev => ({
                              ...prev,
                              selectedMonths: prev.selectedMonths.filter(m => m !== month)
                            }));
                          }
                        }}
                        className="mr-3"
                      />
                      <span className="text-xm">{month}</span>
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
          disabled={loading || !salesColumns.length || !targetColumns.length}
          className="bg-red-600 text-white px-6 py-2 rounded hover:bg-red-700 disabled:bg-gray-400"
        >
          {loading ? 'Calculating...' : 'Calculate Executive Target vs Billed'}
        </button>
      </div>

      {/* Results */}
      {results && (
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-xl font-bold text-blue-700 mb-4">Results</h3>

          {/* Success Message */}
          <div className="bg-green-100 border border-green-400 text-green-700 px-4 py-3 rounded mb-4">
            ✅ Results calculated and automatically added to consolidated reports!
          </div>

          {/* Download Buttons */}
          <div className="flex justify-center gap-4 mb-6">
            <button
              onClick={handleDownloadPpt}
              disabled={downloadingPpt || !results}
              className="bg-green-600 text-white px-6 py-2 rounded hover:bg-green-700"
            >
              {downloadingPpt ? (
                <>
                  <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white"></div>
                  Generating PPT...
                </>
              ) : (
                <>
                  📊 Download PowerPoint
                </>
              )}
            </button>

            <button
              onClick={handleDownloadProof}
              disabled={downloadingProof || !results}
              className="bg-yellow-600 text-white px-6 py-2 rounded hover:bg-yellow-700 disabled:bg-gray-400"
            >
              {downloadingProof ? (
                <>
                  <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white"></div>
                  Generating Excel...
                </>
              ) : (
                <>
                  📋 Download Proof of Calculation
                </>
              )}
            </button>
          </div>

          {/* Data Tables */}
          {['budget_vs_billed_qty', 'budget_vs_billed_value', 'overall_sales_qty', 'overall_sales_value'].map((type, index) => {
            const labelMap = {
              'budget_vs_billed_qty': 'Target vs Billed Quantity',
              'budget_vs_billed_value': 'Target vs Billed Value',
              'overall_sales_qty': 'Overall Sales Quantity',
              'overall_sales_value': 'Overall Sales Values'
            };
            
            const columnOrderMap = {
              'budget_vs_billed_qty': ['Executive', 'Target Qty', 'Billed Qty', '%'],
              'budget_vs_billed_value': ['Executive', 'Target Value', 'Billed Value', '%'],
              'overall_sales_qty': ['Executive', 'Target Qty', 'Billed Qty'],
              'overall_sales_value': ['Executive', 'Target Value', 'Billed Value']
            };
            
            const rows = results[type] || [];
            const orderedCols = columnOrderMap[type].filter(col => {
              // Map display column names to actual data column names
              const columnMap = {
                'Target Qty': 'Budget Qty',
                'Target Value': 'Budget Value'
              };
              const actualCol = columnMap[col] || col;
              return rows[0] && actualCol in rows[0];
            });
            
            if (!rows.length) return null;
            
            return (
              <div key={type} className="mb-8">
                <h4 className="text-lg font-semibold text-gray-800 mb-3">{labelMap[type]}</h4>
                <div className="overflow-x-auto">
                  <table className="min-w-full table-auto border-collapse border border-gray-300">
                    <thead>
                      <tr className="bg-blue-600 text-white">
                        {orderedCols.map(col => (
                          <th key={col} className="border border-gray-300 px-4 py-2 text-left font-semibold">
                            {col}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((row, i) => (
                        <tr 
                          key={i} 
                          className={`
                            ${row.Executive === 'TOTAL' 
                              ? 'bg-gray-200 font-bold border-t-2 border-gray-400' 
                              : i % 2 === 0 
                                ? 'bg-gray-50' 
                                : 'bg-white'
                            } hover:bg-blue-50
                          `}
                        >
                          {orderedCols.map(col => {
                            // Map display column names to actual data column names
                            const columnMap = {
                              'Target Qty': 'Budget Qty',
                              'Target Value': 'Budget Value'
                            };
                            const actualCol = columnMap[col] || col;
                            
                            return (
                              <td key={col} className="border border-gray-300 px-4 py-2">
                                {col === '%' 
                                  ? `${Number(row[actualCol]).toFixed(2)}%` 
                                  : typeof row[actualCol] === 'number' 
                                    ? Number(row[actualCol]).toLocaleString() 
                                    : row[actualCol]
                                }
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

export default ExecutiveBudgetVsBilled;
