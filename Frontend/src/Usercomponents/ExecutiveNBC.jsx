import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import { addReportToStorage } from '../utils/consolidatedStorage';

const CustomerODAnalysis = () => {
 const { selectedFiles } = useExcelData();
 
 const [activeTab, setActiveTab] = useState('customers');
 
 return (
   <div className="p-6">
     <h2 className="text-2xl font-bold text-blue-800 mb-6">Customer & OD Analysis</h2>
     
     {/* Tab Navigation */}
     <div className="flex mb-6">
       <button
         className={`px-6 py-3 rounded-t-lg font-medium ${
           activeTab === 'customers' 
             ? 'bg-blue-600 text-white border-b-2 border-blue-600' 
             : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
         }`}
         onClick={() => setActiveTab('customers')}
       >
         Number Of Billed Customers
       </button>
       <button
         className={`px-6 py-3 rounded-t-lg font-medium ml-2 ${
           activeTab === 'od_target' 
             ? 'bg-blue-600 text-white border-b-2 border-blue-600' 
             : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
         }`}
         onClick={() => setActiveTab('od_target')}
       >
         OD Target
       </button>
     </div>
     
     {/* Tab Content */}
     <div className="border-2 border-gray-200 rounded-lg p-6">
       {activeTab === 'customers' && <BilledCustomersTab />}
       {activeTab === 'od_target' && <ODTargetTab />}
     </div>
   </div>
 );
};

// Billed Customers Tab Component - WITH CONSOLIDATION
const BilledCustomersTab = () => {
 const { selectedFiles } = useExcelData();
 
 // Sheet configurations
 const [sheet, setSheet] = useState('Sheet1');
 const [sheets, setSheets] = useState([]);
 const [headerRow, setHeaderRow] = useState(1);
 
 // Column configurations
 const [columns, setColumns] = useState([]);
 const [columnSelections, setColumnSelections] = useState({
   date: '',
   branch: '',
   customer_id: '',
   executive: ''
 });
 
 // Options and filters
 const [availableMonths, setAvailableMonths] = useState([]);
 const [branchOptions, setBranchOptions] = useState([]);
 const [executiveOptions, setExecutiveOptions] = useState([]);
 const [filters, setFilters] = useState({
   selectedMonths: [],
   selectedBranches: [],
   selectedExecutives: []
 });
 
 // State management
 const [results, setResults] = useState(null);
 const [loading, setLoading] = useState(false);
 const [downloadingPpt, setDownloadingPpt] = useState(false);
 const [error, setError] = useState(null);
 
 // Fetch available sheet names
 const fetchSheets = async () => {
   if (!selectedFiles.salesFile) return;
   
   try {
     const res = await axios.post('http://localhost:5000/api/branch/sheets', { filename: selectedFiles.salesFile });
     if (res.data && res.data.sheets) {
       setSheets(res.data.sheets);
     } else {
       setSheets([]);
     }
   } catch (error) {
     console.error('Error fetching sheets:', error);
     setError(`Failed to load sheet names: ${error.response?.data?.error || error.message}`);
     setSheets([]);
   }
 };
 
 useEffect(() => {
   fetchSheets();
 }, [selectedFiles.salesFile]);
 
 // Fetch columns and auto-map
 const fetchColumns = async () => {
   if (!sheet) {
     setError('Please select a sheet first');
     return;
   }
   
   setLoading(true);
   setError(null);
   
   try {
     // Step 1: Get columns
     const colRes = await axios.post('http://localhost:5000/api/branch/get_columns', {
       filename: selectedFiles.salesFile,
       sheet_name: sheet,
       header: headerRow
     });
     
     if (!colRes.data || !colRes.data.columns) {
       setError('No columns found in the sheet');
       setColumns([]);
       return;
     }
     
     setColumns(colRes.data.columns);
     
     // Step 2: Auto-map columns
     try {
       const mapRes = await axios.post('http://localhost:5000/api/executive/customer_auto_map_columns', {
         sales_file_path: `uploads/${selectedFiles.salesFile}`
       });
       
       if (mapRes.data && mapRes.data.success && mapRes.data.mapping) {
         setColumnSelections(mapRes.data.mapping);
         
         // Step 3: Auto-load options - WAIT for column selections to be set
         setTimeout(async () => {
           await loadFilterOptions(mapRes.data.mapping);
         }, 200);
       } else {
         setError('Auto-mapping failed. Please map columns manually.');
       }
     } catch (mapError) {
       setError('Auto-mapping failed. Please map columns manually.');
     }
     
   } catch (error) {
     console.error('Error fetching columns:', error);
     setError(`Failed to load columns: ${error.response?.data?.error || error.message}`);
     setColumns([]);
   } finally {
     setLoading(false);
   }
 };
 
 // Load filter options function
 const loadFilterOptions = async (mappings = null) => {
   const mapping = mappings || columnSelections;
   
   // Validate mappings
   const requiredColumns = ['date', 'branch', 'customer_id', 'executive'];
   const missingColumns = requiredColumns.filter(col => !mapping[col]);
   
   if (missingColumns.length > 0) {
     console.warn('Missing required columns:', missingColumns);
     return;
   }
   
   try {
     const res = await axios.post('http://localhost:5000/api/executive/customer_get_options', {
       sales_file_path: `uploads/${selectedFiles.salesFile}`
     });
     
     if (res.data && res.data.success) {
       setAvailableMonths(res.data.available_months || []);
       setBranchOptions(res.data.branches || []);
       setExecutiveOptions(res.data.executives || []);
       
       // Set all options as selected by default
       setFilters({
         selectedMonths: res.data.available_months || [],
         selectedBranches: res.data.branches || [],
         selectedExecutives: res.data.executives || []
       });
       
       // Clear any mapping-related errors
       if (error && (error.includes('map columns') || error.includes('Auto-mapping'))) {
         setError(null);
       }
     }
   } catch (error) {
     console.error('Error loading filter options:', error);
   }
 };
 
 // Watch for column selection changes and auto-load options
 useEffect(() => {
   const hasAllColumns = columnSelections.date && columnSelections.branch && 
                        columnSelections.customer_id && columnSelections.executive;
   
   if (hasAllColumns && columns.length > 0) {
     const timer = setTimeout(() => {
       loadFilterOptions();
     }, 300);
     
     return () => clearTimeout(timer);
   }
 }, [columnSelections.date, columnSelections.branch, columnSelections.customer_id, columnSelections.executive]);
 
 // Function to add customer results to consolidated storage
 const addCustomerReportsToStorage = (resultsData) => {
   try {
     const customerReports = [];
     
     // Process each financial year result
     Object.entries(resultsData).forEach(([financialYear, data]) => {
       customerReports.push({
         df: data.data || [],
         title: `NUMBER OF BILLED CUSTOMERS - FY ${financialYear}`,
         percent_cols: [] // No percentage columns for customer reports
       });
     });

     if (customerReports.length > 0) {
       addReportToStorage(customerReports, 'customers_results');
       console.log(`✅ Added ${customerReports.length} customer reports to consolidated storage`);
     }
   } catch (error) {
     console.error('Error adding customer reports to consolidated storage:', error);
   }
 };

 // Handle generate report
 const handleGenerateReport = async () => {
   if (!selectedFiles.salesFile) {
     setError('Please upload a sales file');
     return;
   }
   
   if (!filters.selectedMonths.length) {
     setError('Please select at least one month');
     return;
   }
   
   setLoading(true);
   setError(null);
   
   try {
     const payload = {
       sales_file_path: `uploads/${selectedFiles.salesFile}`,
       date_column: columnSelections.date,
       branch_column: columnSelections.branch,
       customer_id_column: columnSelections.customer_id,
       executive_column: columnSelections.executive,
       selected_months: filters.selectedMonths,
       selected_branches: filters.selectedBranches,
       selected_executives: filters.selectedExecutives
     };
     
     const res = await axios.post('http://localhost:5000/api/executive/calculate_customer_analysis', payload);
     
     if (res.data && res.data.success) {
       setResults(res.data.results);
       setError(null);
       
       // 🎯 ADD TO CONSOLIDATED REPORTS - Just like Streamlit!
       addCustomerReportsToStorage(res.data.results);
       
     } else {
       setError(res.data?.error || 'Failed to generate report');
     }
   } catch (error) {
     console.error('Error generating report:', error);
     setError(`Error generating report: ${error.response?.data?.error || error.message}`);
   } finally {
     setLoading(false);
   }
 };
 
 // Handle download PPT
 const handleDownloadPpt = async (financialYear) => {
   if (!results || !results[financialYear]) {
     setError('No results available for PPT generation');
     return;
   }
   
   setDownloadingPpt(true);
   setError(null);
   
   try {
     const payload = {
       results_data: { results: results },
       title: `NUMBER OF BILLED CUSTOMERS - FY ${financialYear}`,
       logo_file: null
     };
     
     const response = await axios.post('http://localhost:5000/api/executive/generate_customer_ppt', payload, {
       responseType: 'blob',
       headers: {
         'Content-Type': 'application/json',
       },
     });
     
     const blob = new Blob([response.data], {
       type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
     });
     
     const url = window.URL.createObjectURL(blob);
     const link = document.createElement('a');
     link.href = url;
     link.download = `Billed_Customers_FY_${financialYear}.pptx`;
     document.body.appendChild(link);
     link.click();
     link.remove();
     window.URL.revokeObjectURL(url);
     
   } catch (error) {
     console.error('Error downloading PPT:', error);
     setError('Failed to download PowerPoint presentation');
   } finally {
     setDownloadingPpt(false);
   }
 };
 
 if (!selectedFiles.salesFile) {
   return (
     <div className="text-center py-8">
       <p className="text-gray-600">⚠️ Please upload a sales file to use this feature</p>
     </div>
   );
 }
 
 return (
   <div>
     {error && (
       <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
         {error}
       </div>
     )}
     
     {/* Sheet Selection */}
     <div className="bg-white p-4 rounded-lg shadow mb-6">
       <h3 className="text-lg font-semibold text-blue-700 mb-4">Sheet Configuration</h3>
       <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
         <div>
           <label className="block font-semibold mb-2">Select Sheet</label>
           <select 
             className="w-full p-2 border border-gray-300 rounded" 
             value={sheet} 
             onChange={e => setSheet(e.target.value)}
             disabled={loading}
           >
             <option value="">Select Sheet</option>
             {sheets.map(sheetName => (
               <option key={sheetName} value={sheetName}>{sheetName}</option>
             ))}
           </select>
         </div>
         <div>
           <label className="block font-semibold mb-2">Header Row (1-based)</label>
           <input 
             type="number" 
             className="w-full p-2 border border-gray-300 rounded" 
             min={1} 
             max={10}
             value={headerRow}
             onChange={e => setHeaderRow(Number(e.target.value))}
             disabled={loading}
           />
         </div>
       </div>
       <button
         className="mt-4 bg-blue-600 text-white px-6 py-2 rounded hover:bg-blue-700 disabled:bg-gray-400"
         onClick={fetchColumns}
         disabled={!sheet || loading}
       >
         {loading ? 'Loading...' : 'Load Columns & Auto-Map'}
       </button>
     </div>
     
     {/* Column Mapping */}
     {columns.length > 0 && (
       <div className="bg-white p-4 rounded-lg shadow mb-6">
         <h3 className="text-lg font-semibold text-blue-700 mb-4">Column Mapping</h3>
         <div className="grid grid-cols-2 gap-4">
           <div>
             <label className="block font-medium mb-1">Date Column *</label>
             <select
               className="w-full p-2 border border-gray-300 rounded"
               value={columnSelections.date || ''}
               onChange={(e) => setColumnSelections(prev => ({...prev, date: e.target.value}))}
             >
               <option value="">Select Column</option>
               {columns.map(col => (
                 <option key={col} value={col}>{col}</option>
               ))}
             </select>
           </div>
           <div>
             <label className="block font-medium mb-1">Branch Column *</label>
             <select
               className="w-full p-2 border border-gray-300 rounded"
               value={columnSelections.branch || ''}
               onChange={(e) => setColumnSelections(prev => ({...prev, branch: e.target.value}))}
             >
               <option value="">Select Column</option>
               {columns.map(col => (
                 <option key={col} value={col}>{col}</option>
               ))}
             </select>
           </div>
           <div>
             <label className="block font-medium mb-1">Customer ID Column *</label>
             <select
               className="w-full p-2 border border-gray-300 rounded"
               value={columnSelections.customer_id || ''}
               onChange={(e) => setColumnSelections(prev => ({...prev, customer_id: e.target.value}))}
             >
               <option value="">Select Column</option>
               {columns.map(col => (
                 <option key={col} value={col}>{col}</option>
               ))}
             </select>
           </div>
           <div>
             <label className="block font-medium mb-1">Executive Column *</label>
             <select
               className="w-full p-2 border border-gray-300 rounded"
               value={columnSelections.executive || ''}
               onChange={(e) => setColumnSelections(prev => ({...prev, executive: e.target.value}))}
             >
               <option value="">Select Column</option>
               {columns.map(col => (
                 <option key={col} value={col}>{col}</option>
               ))}
             </select>
           </div>
         </div>
       </div>
     )}
     
     {/* Filters */}
     {(availableMonths.length > 0 || branchOptions.length > 0 || executiveOptions.length > 0) && (
       <div className="bg-white p-4 rounded-lg shadow mb-6">
         <h3 className="text-lg font-semibold text-blue-700 mb-4">Filter Options</h3>
         
         {/* Months Filter */}
         {availableMonths.length > 0 && (
           <div className="mb-4">
             <label className="block font-semibold mb-2">
               Select Months ({filters.selectedMonths.length} of {availableMonths.length})
             </label>
             <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
               <label className="flex items-center mb-2">
                 <input
                   type="checkbox"
                   checked={filters.selectedMonths.length === availableMonths.length}
                   onChange={(e) => {
                     if (e.target.checked) {
                       setFilters(prev => ({...prev, selectedMonths: availableMonths}));
                     } else {
                       setFilters(prev => ({...prev, selectedMonths: []}));
                     }
                   }}
                   className="mr-2"
                 />
                 <span className="font-medium text-sm">Select All</span>
               </label>
               {availableMonths.map(month => (
                 <label key={month} className="flex items-center mb-1">
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
                     className="mr-2"
                   />
                   <span className="text-xs">{month}</span>
                 </label>
               ))}
             </div>
           </div>
         )}
         
         <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
           {/* Branches Filter */}
           {branchOptions.length > 0 && (
             <div>
               <label className="block font-semibold mb-2">
                 Branches ({filters.selectedBranches.length} of {branchOptions.length})
               </label>
               <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
                 <label className="flex items-center mb-2">
                   <input
                     type="checkbox"
                     checked={filters.selectedBranches.length === branchOptions.length}
                     onChange={(e) => {
                       if (e.target.checked) {
                         setFilters(prev => ({...prev, selectedBranches: branchOptions}));
                       } else {
                         setFilters(prev => ({...prev, selectedBranches: []}));
                       }
                     }}
                     className="mr-2"
                   />
                   <span className="font-medium text-sm">Select All</span>
                 </label>
                 {branchOptions.map(branch => (
                   <label key={branch} className="flex items-center mb-1">
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
                       className="mr-2"
                     />
                     <span className="text-xs">{branch}</span>
                   </label>
                 ))}
               </div>
             </div>
           )}
           
           {/* Executives Filter */}
           {executiveOptions.length > 0 && (
             <div>
               <label className="block font-semibold mb-2">
                 Executives ({filters.selectedExecutives.length} of {executiveOptions.length})
               </label>
               <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
                 <label className="flex items-center mb-2">
                   <input
                     type="checkbox"
                     checked={filters.selectedExecutives.length === executiveOptions.length}
                     onChange={(e) => {
                       if (e.target.checked) {
                         setFilters(prev => ({...prev, selectedExecutives: executiveOptions}));
                       } else {
                         setFilters(prev => ({...prev, selectedExecutives: []}));
                       }
                     }}
                     className="mr-2"
                   />
                   <span className="font-medium text-sm">Select All</span>
                 </label>
                 {executiveOptions.map(exec => (
                   <label key={exec} className="flex items-center mb-1">
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
                       className="mr-2"
                     />
                     <span className="text-xs">{exec}</span>
                   </label>
                 ))}
               </div>
             </div>
           )}
         </div>
       </div>
     )}
     
     {/* Generate Report Button */}
     <div className="text-center mb-6">
       <button
         onClick={handleGenerateReport}
         disabled={loading || !columns.length || !filters.selectedMonths.length}
         className="bg-green-600 text-white px-8 py-3 rounded-lg text-lg font-semibold hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
       >
         {loading ? 'Generating...' : 'Generate Report'}
       </button>
     </div>
     
     {/* Results */}
     {results && Object.keys(results).length > 0 && (
       <div className="bg-white p-6 rounded-lg shadow">
         <h3 className="text-xl font-bold text-blue-700 mb-4">Customer Analysis Results</h3>
         
         {/* Success Message */}
         <div className="bg-green-100 border border-green-400 text-green-700 px-4 py-3 rounded mb-4">
           ✅ Results calculated and automatically added to consolidated reports!
         </div>
         
         {Object.entries(results).map(([financialYear, data]) => (
           <div key={financialYear} className="mb-8">
             <div className="flex justify-between items-center mb-4">
               <h4 className="text-lg font-semibold text-gray-800">Financial Year: {financialYear}</h4>
               <button
                 onClick={() => handleDownloadPpt(financialYear)}
                 disabled={downloadingPpt}
                 className="bg-purple-600 text-white px-4 py-2 rounded hover:bg-purple-700 disabled:bg-gray-400"
               >
                 {downloadingPpt ? 'Generating PPT...' : 'Download PPT'}
               </button>
             </div>
             
             <div className="overflow-x-auto">
               <table className="min-w-full table-auto border-collapse border border-gray-300">
                 <thead>
                   <tr className="bg-blue-600 text-white">
                     {data.columns.filter(col => col.toLowerCase() !== 's.no' && col.toLowerCase() !== 'sno').map(col => (
                       <th key={col} className="border border-gray-300 px-4 py-2 text-left font-semibold">
                         {col}
                       </th>
                     ))}
                   </tr>
                 </thead>
                 <tbody>
                   {data.data.map((row, i) => (
                     <tr 
                       key={i} 
                       className={`
                         ${row['Executive Name'] === 'GRAND TOTAL' 
                           ? 'bg-gray-200 font-bold border-t-2 border-gray-400' 
                           : i % 2 === 0 
                             ? 'bg-gray-50' 
                             : 'bg-white'
                         } hover:bg-blue-50
                       `}
                     >
                       {data.columns.filter(col => col.toLowerCase() !== 's.no' && col.toLowerCase() !== 'sno').map((col, j) => (
                         <td key={j} className="border border-gray-300 px-4 py-2">
                           {row[col] || ''}
                         </td>
                       ))}
                     </tr>
                   ))}
                 </tbody>
               </table>
             </div>
           </div>
         ))}
       </div>
     )}
   </div>
 );
};

// OD Target Tab Component - WITH CONSOLIDATION
const ODTargetTab = () => {
 const { selectedFiles } = useExcelData();
 
 // File selection
 const [fileChoice, setFileChoice] = useState('os_feb'); // Default to current month
 const [currentFile, setCurrentFile] = useState(null);
 
 // Sheet configurations
 const [sheet, setSheet] = useState('Sheet1');
 const [sheets, setSheets] = useState([]);
 const [headerRow, setHeaderRow] = useState(1);
 
 // Column configurations
 const [columns, setColumns] = useState([]);
 const [columnSelections, setColumnSelections] = useState({
   area: '',
   net_value: '',
   due_date: '',
   executive: ''
 });
 
 // Options and filters
 const [yearOptions, setYearOptions] = useState([]);
 const [branchOptions, setBranchOptions] = useState([]);
 const [executiveOptions, setExecutiveOptions] = useState([]);
 const [filters, setFilters] = useState({
   selectedYears: [],
   selectedBranches: [],
   selectedExecutives: [],
   tillMonth: 'January'
 });
 
 // State management
 const [results, setResults] = useState(null);
 const [loading, setLoading] = useState(false);
 const [downloadingPpt, setDownloadingPpt] = useState(false);
 const [error, setError] = useState(null);
 
 // Update current file when choice or files change
 useEffect(() => {
   // Reset states when changing files
   setError(null);
   setSheet('');
   setSheets([]);
   setColumns([]);
   setColumnSelections({
     area: '',
     net_value: '',
     due_date: '',
     executive: ''
   });
   
   // Determine current file based on choice
   let newCurrentFile = null;
   if (fileChoice === 'os_jan' && selectedFiles.osPrevFile) {
     newCurrentFile = selectedFiles.osPrevFile;
   } else if (fileChoice === 'os_feb' && selectedFiles.osCurrFile) {
     newCurrentFile = selectedFiles.osCurrFile;
   }
   
   setCurrentFile(newCurrentFile);
 }, [fileChoice, selectedFiles.osPrevFile, selectedFiles.osCurrFile]);
 
 // Fetch sheets when file changes
 useEffect(() => {
   if (currentFile) {
     fetchSheets();
   } else {
     setSheets([]);
   }
 }, [currentFile]);
 
 const fetchSheets = async () => {
   if (!currentFile) return;
   
   setLoading(true);
   setError(null);
   
   try {
     const res = await axios.post('http://localhost:5000/api/branch/sheets', { 
       filename: currentFile 
     });
     
     if (res.data && res.data.sheets && Array.isArray(res.data.sheets)) {
       setSheets(res.data.sheets);
       if (res.data.sheets.includes('Sheet1')) {
         setSheet('Sheet1');
       } else if (res.data.sheets.length > 0) {
         setSheet(res.data.sheets[0]); // Set first sheet as default
       }  
     } else {
       setError('No sheets found in the file');
       setSheets([]);
     }
   } catch (error) {
     console.error('Error fetching sheets:', error);
     const errorMsg = error.response?.data?.error || error.message;
     setError(`Failed to load sheet names: ${errorMsg}`);
     setSheets([]);
   } finally {
     setLoading(false);
   }
 };
 
 // Fetch columns and auto-map
 const fetchColumns = async () => {
   if (!sheet || !currentFile) {
     setError('Please select a sheet first');
     return;
   }
   
   setLoading(true);
   setError(null);
   
   try {
     // Step 1: Get columns
     const colRes = await axios.post('http://localhost:5000/api/branch/get_columns', {
       filename: currentFile,
       sheet_name: sheet,
       header: headerRow
     });
     
     if (!colRes.data || !colRes.data.columns || !Array.isArray(colRes.data.columns)) {
       setError('No columns found in the sheet');
       setColumns([]);
       return;
     }
     
     setColumns(colRes.data.columns);
     
     // Step 2: Auto-map columns
     try {
       const mapRes = await axios.post('http://localhost:5000/api/executive/od_target_auto_map_columns', {
         os_file_path: `uploads/${currentFile}`
       });
       
       if (mapRes.data && mapRes.data.success && mapRes.data.mapping) {
         setColumnSelections(mapRes.data.mapping);
         
         // Step 3: Auto-load options with delay
         setTimeout(async () => {
           await loadFilterOptions(mapRes.data.mapping);
         }, 200);
       } else {
         setError('Auto-mapping failed. Please map columns manually.');
       }
     } catch (mapError) {
       setError('Auto-mapping failed. Please map columns manually.');
     }
   } catch (error) {
     console.error('Error fetching columns:', error);
     const errorMsg = error.response?.data?.error || error.message;
     setError(`Failed to load columns: ${errorMsg}`);
     setColumns([]);
   } finally {
     setLoading(false);
   }
 };
 
 // Load filter options function
 const loadFilterOptions = async (mappings = null) => {
   const mapping = mappings || columnSelections;
   
   // Validate mappings
   const requiredColumns = ['due_date', 'area', 'executive'];
   const missingColumns = requiredColumns.filter(col => !mapping[col]);
   
   if (missingColumns.length > 0 || !currentFile) {
     console.warn('Missing required columns or file:', missingColumns);
     return;
   }
   
   try {
     const res = await axios.post('http://localhost:5000/api/executive/od_target_get_options', {
       os_file_path: `uploads/${currentFile}`
     });
     
     if (res.data && res.data.success) {
       setYearOptions(res.data.years || []);
       setBranchOptions(res.data.branches || []);
       setExecutiveOptions(res.data.executives || []);
       
       // Set all options as selected by default
       setFilters(prev => ({
         ...prev,
         selectedYears: res.data.years || [],
         selectedBranches: res.data.branches || [],
         selectedExecutives: res.data.executives || []
       }));
       
       // Clear mapping-related errors
       if (error && (error.includes('map columns') || error.includes('Auto-mapping'))) {
         setError(null);
       }
     }
   } catch (error) {
     console.error('Error loading filter options:', error);
   }
 };
 
 // Watch for column selection changes and auto-load options
 useEffect(() => {
   const hasAllColumns = columnSelections.area && columnSelections.net_value && 
                        columnSelections.due_date && columnSelections.executive;
   
   if (hasAllColumns && columns.length > 0 && currentFile) {
     const timer = setTimeout(() => {
       loadFilterOptions();
     }, 300);
     
     return () => clearTimeout(timer);
   }
 }, [columnSelections.area, columnSelections.net_value, columnSelections.due_date, columnSelections.executive]);
 
 // Function to add OD target results to consolidated storage
 const addODTargetReportsToStorage = (resultsData) => {
   try {
     const odTargetReports = [{
       df: resultsData.data || [],
       title: `OD Target - ${resultsData.end_date || 'All Periods'}`,
       percent_cols: [] // No percentage columns for OD target reports
     }];

     if (odTargetReports.length > 0) {
       addReportToStorage(odTargetReports, 'od_results');
       console.log(`✅ Added ${odTargetReports.length} OD target reports to consolidated storage`);
     }
   } catch (error) {
     console.error('Error adding OD target reports to consolidated storage:', error);
   }
 };

 // Handle generate report
 const handleGenerateReport = async () => {
   if (!currentFile) {
     setError('Please select an OS file');
     return;
   }
   
   if (!filters.selectedYears.length || !filters.tillMonth) {
     setError('Please select at least one year and one month');
     return;
   }
   
   setLoading(true);
   setError(null);
   
   try {
     const payload = {
       os_file_path: `uploads/${currentFile}`,
       area_column: columnSelections.area,
       net_value_column: columnSelections.net_value,
       due_date_column: columnSelections.due_date,
       executive_column: columnSelections.executive,
       selected_branches: filters.selectedBranches,
       selected_years: filters.selectedYears,
       till_month: filters.tillMonth,
       selected_executives: filters.selectedExecutives
     };
     
     const res = await axios.post('http://localhost:5000/api/executive/calculate_od_target', payload);
     
     if (res.data && res.data.success) {
       setResults(res.data);
       setError(null);
       
       // 🎯 ADD TO CONSOLIDATED REPORTS - Just like Streamlit!
       addODTargetReportsToStorage(res.data);
       
     } else {
       setError(res.data?.error || 'Failed to generate report');
     }
   } catch (error) {
     console.error('Error generating report:', error);
     const errorMsg = error.response?.data?.error || error.message;
     setError(`Error generating report: ${errorMsg}`);
   } finally {
     setLoading(false);
   }
 };
 
 // Handle download PPT
 const handleDownloadPpt = async () => {
   if (!results) {
     setError('No results available for PPT generation');
     return;
   }
   
   setDownloadingPpt(true);
   setError(null);
   
   try {
     const title = `OD Target - ${results.end_date || 'All Periods'}`;
     
     const payload = {
       results_data: results,
       title: title,
       logo_file: null
     };
     
     const response = await axios.post('http://localhost:5000/api/executive/generate_od_target_ppt', payload, {
       responseType: 'blob',
       headers: {
         'Content-Type': 'application/json',
       },
     });
     
     const blob = new Blob([response.data], {
       type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
     });
     
     const url = window.URL.createObjectURL(blob);
     const link = document.createElement('a');
     link.href = url;
     link.download = `OD_Target_${results.end_date || 'Report'}.pptx`;
     document.body.appendChild(link);
     link.click();
     link.remove();
     window.URL.revokeObjectURL(url);
     
   } catch (error) {
     console.error('Error downloading PPT:', error);
     setError('Failed to download PowerPoint presentation');
   } finally {
     setDownloadingPpt(false);
   }
 };
 
 const monthOptions = [
   'January', 'February', 'March', 'April', 'May', 'June',
   'July', 'August', 'September', 'October', 'November', 'December'
 ];
 
 return (
   <div>
     {error && (
       <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
         {error}
       </div>
     )}
     
     {/* File Selection */}
     <div className="bg-white p-4 rounded-lg shadow mb-6">
       <h3 className="text-lg font-semibold text-blue-700 mb-4">Choose OS File</h3>
       <div className="flex gap-4 mb-4">
         <label className="flex items-center">
           <input
             type="radio"
             value="os_jan"
             checked={fileChoice === 'os_jan'}
             onChange={(e) => setFileChoice(e.target.value)}
             className="mr-2"
           />
           <span className="text-sm">
             OS-Previous Month 
             {selectedFiles.osPrevFile ? (
               <span className="text-green-600">(✓ Uploaded)</span>
             ) : (
               <span className="text-red-600">(✗ Not uploaded)</span>
             )}
           </span>
         </label>
         <label className="flex items-center">
           <input
             type="radio"
             value="os_feb"
             checked={fileChoice === 'os_feb'}
             onChange={(e) => setFileChoice(e.target.value)}
             className="mr-2"
           />
           <span className="text-sm">
             OS-Current Month 
             {selectedFiles.osCurrFile ? (
               <span className="text-green-600">(✓ Uploaded)</span>
             ) : (
               <span className="text-red-600">(✗ Not uploaded)</span>
             )}
           </span>
         </label>
       </div>
       
       <div className="text-sm text-gray-600">
         <strong>Selected file:</strong> {currentFile || 'None available'}
       </div>
     </div>
     
     {!currentFile ? (
       <div className="text-center py-8">
         <p className="text-gray-600">⚠️ No OS file selected or uploaded</p>
         <p className="text-sm text-gray-500 mt-2">
           Please upload the {fileChoice === 'os_jan' ? 'OS-Previous Month' : 'OS-Current Month'} file first.
         </p>
       </div>
     ) : (
       <>
         {/* Sheet Selection */}
         <div className="bg-white p-4 rounded-lg shadow mb-6">
           <h3 className="text-lg font-semibold text-blue-700 mb-4">Sheet Configuration</h3>
           <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
             <div>
               <label className="block font-semibold mb-2">Select Sheet</label>
               <select 
                 className="w-full p-2 border border-gray-300 rounded" 
                 value={sheet} 
                 onChange={e => setSheet(e.target.value)}
                 disabled={loading}
               >
                 <option value="">Select Sheet</option>
                 {sheets.map(sheetName => (
                   <option key={sheetName} value={sheetName}>{sheetName}</option>
                 ))}
               </select>
             </div>
             <div>
               <label className="block font-semibold mb-2">Header Row (1-based)</label>
               <input 
                 type="number" 
                 className="w-full p-2 border border-gray-300 rounded" 
                 min={1} 
                 max={10}
                 value={headerRow}
                 onChange={e => setHeaderRow(Number(e.target.value))}
                 disabled={loading}
               />
             </div>
           </div>
           <button
             className="mt-4 bg-blue-600 text-white px-6 py-2 rounded hover:bg-blue-700 disabled:bg-gray-400"
             onClick={fetchColumns}
             disabled={!sheet || loading || !currentFile}
           >
             {loading ? 'Loading...' : 'Load Columns & Auto-Map'}
           </button>
         </div>
         
         {/* Column Mapping */}
         {columns.length > 0 && (
           <div className="bg-white p-4 rounded-lg shadow mb-6">
             <h3 className="text-lg font-semibold text-blue-700 mb-4">Column Mapping</h3>
             <div className="grid grid-cols-2 gap-4">
               <div>
                 <label className="block font-medium mb-1">Area Column *</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={columnSelections.area || ''}
                   onChange={(e) => setColumnSelections(prev => ({...prev, area: e.target.value}))}
                 >
                   <option value="">Select Column</option>
                   {columns.map(col => (
                     <option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
               <div>
                 <label className="block font-medium mb-1">Net Value Column *</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={columnSelections.net_value || ''}
                   onChange={(e) => setColumnSelections(prev => ({...prev, net_value: e.target.value}))}
                 >
                   <option value="">Select Column</option>
                   {columns.map(col => (
                     <option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
               <div>
                 <label className="block font-medium mb-1">Due Date Column *</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={columnSelections.due_date || ''}
                   onChange={(e) => setColumnSelections(prev => ({...prev, due_date: e.target.value}))}
                 >
                   <option value="">Select Column</option>
                   {columns.map(col => (
                     <option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
               <div>
                 <label className="block font-medium mb-1">Executive Column *</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={columnSelections.executive || ''}
                   onChange={(e) => setColumnSelections(prev => ({...prev, executive: e.target.value}))}
                 >
                   <option value="">Select Column</option>
                   {columns.map(col => (
                     <option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
             </div>
           </div>
         )}
         
         {/* Date Filter */}
         {yearOptions.length > 0 && (
           <div className="bg-white p-4 rounded-lg shadow mb-6">
             <h3 className="text-lg font-semibold text-blue-700 mb-4">Due Date Filter</h3>
             <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
               <div>
                 <label className="block font-semibold mb-2">Select Years</label>
                 <div className="max-h-32 overflow-y-auto border border-gray-300 rounded p-2">
                   {yearOptions.map(year => (
                     <label key={year} className="flex items-center mb-1">
                       <input
                         type="checkbox"
                         checked={filters.selectedYears.includes(year)}
                         onChange={(e) => {
                           if (e.target.checked) {
                             setFilters(prev => ({
                               ...prev,
                               selectedYears: [...prev.selectedYears, year]
                             }));
                           } else {
                             setFilters(prev => ({
                               ...prev,
                               selectedYears: prev.selectedYears.filter(y => y !== year)
                             }));
                           }
                         }}
                         className="mr-2"
                       />
                       <span className="text-sm">{year}</span>
                     </label>
                   ))}
                 </div>
               </div>
               <div>
                 <label className="block font-semibold mb-2">Select Till Month</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={filters.tillMonth}
                   onChange={(e) => setFilters(prev => ({...prev, tillMonth: e.target.value}))}
                 >
                   {monthOptions.map(month => (
                     <option key={month} value={month}>{month}</option>
                   ))}
                 </select>
               </div>
             </div>
           </div>
         )}
         
         {/* Filters */}
         {(branchOptions.length > 0 || executiveOptions.length > 0) && (
           <div className="bg-white p-4 rounded-lg shadow mb-6">
             <h3 className="text-lg font-semibold text-blue-700 mb-4">Filter Options</h3>
             <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
               {/* Branches Filter */}
               {branchOptions.length > 0 && (
                 <div>
                   <label className="block font-semibold mb-2">
                     Branches ({filters.selectedBranches.length} of {branchOptions.length})
                   </label>
                   <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
                     <label className="flex items-center mb-2">
                       <input
                         type="checkbox"
                         checked={filters.selectedBranches.length === branchOptions.length}
                         onChange={(e) => {
                           if (e.target.checked) {
                             setFilters(prev => ({...prev, selectedBranches: branchOptions}));
                           } else {
                             setFilters(prev => ({...prev, selectedBranches: []}));
                           }
                         }}
                         className="mr-2"
                       />
                       <span className="font-medium text-sm">Select All</span>
                     </label>
                     {branchOptions.map(branch => (
                       <label key={branch} className="flex items-center mb-1">
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
                           className="mr-2"
                         />
                         <span className="text-xs">{branch}</span>
                       </label>
                     ))}
                   </div>
                 </div>
               )}
               
               {/* Executives Filter */}
               {executiveOptions.length > 0 && (
                 <div>
                   <label className="block font-semibold mb-2">
                     Executives ({filters.selectedExecutives.length} of {executiveOptions.length})
                   </label>
                   <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
                     <label className="flex items-center mb-2">
                       <input
                         type="checkbox"
                         checked={filters.selectedExecutives.length === executiveOptions.length}
                         onChange={(e) => {
                           if (e.target.checked) {
                             setFilters(prev => ({...prev, selectedExecutives: executiveOptions}));
                           } else {
                             setFilters(prev => ({...prev, selectedExecutives: []}));
                           }
                         }}
                         className="mr-2"
                       />
                       <span className="font-medium text-sm">Select All</span>
                     </label>
                     {executiveOptions.map(exec => (
                       <label key={exec} className="flex items-center mb-1">
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
                           className="mr-2"
                         />
                         <span className="text-xs">{exec}</span>
                       </label>
                     ))}
                   </div>
                 </div>
               )}
             </div>
           </div>
         )}
         
         {/* Generate Report Button */}
         <div className="text-center mb-6">
           <button
             onClick={handleGenerateReport}
             disabled={loading || !columns.length || !filters.selectedYears.length}
             className="bg-green-600 text-white px-8 py-3 rounded-lg text-lg font-semibold hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
           >
             {loading ? 'Generating...' : 'Generate Report'}
           </button>
         </div>
         
         {/* Results */}
         {results && (
           <div className="bg-white p-6 rounded-lg shadow">
             <div className="flex justify-between items-center mb-4">
               <h3 className="text-xl font-bold text-blue-700">
                 OD Target Results - {results.end_date || 'All Periods'}
               </h3>
               <button
                 onClick={handleDownloadPpt}
                 disabled={downloadingPpt}
                 className="bg-purple-600 text-white px-4 py-2 rounded hover:bg-purple-700 disabled:bg-gray-400"
               >
                 {downloadingPpt ? 'Generating PPT...' : 'Download PPT'}
               </button>
             </div>
             
             {/* Success Message */}
             <div className="bg-green-100 border border-green-400 text-green-700 px-4 py-3 rounded mb-4">
               ✅ Results calculated and automatically added to consolidated reports!
             </div>
             
             <div className="overflow-x-auto">
               <table className="min-w-full table-auto border-collapse border border-gray-300">
                 <thead>
                   <tr className="bg-blue-600 text-white">
                     {results.columns.filter(col => col.toLowerCase() !== 's.no' && col.toLowerCase() !== 'sno').map(col => (
                       <th key={col} className="border border-gray-300 px-4 py-2 text-left font-semibold">
                         {col}
                       </th>
                     ))}
                   </tr>
                 </thead>
                 <tbody>
                   {results.data.map((row, i) => (
                     <tr 
                       key={i} 
                       className={`
                         ${row['Executive'] === 'TOTAL' 
                           ? 'bg-gray-200 font-bold border-t-2 border-gray-400' 
                           : i % 2 === 0 
                             ? 'bg-gray-50' 
                             : 'bg-white'
                         } hover:bg-blue-50
                       `}
                     >
                       {results.columns.filter(col => col.toLowerCase() !== 's.no' && col.toLowerCase() !== 'sno').map((col, j) => (
                         <td key={j} className="border border-gray-300 px-4 py-2">
                           {col === 'TARGET' ? Number(row[col]).toFixed(2) : (row[col] || '')}
                         </td>
                       ))}
                     </tr>
                   ))}
                 </tbody>
               </table>
             </div>
           </div>
         )}
       </>
     )}
   </div>
 );
};

export default CustomerODAnalysis;
