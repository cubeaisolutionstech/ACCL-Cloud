import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import { addReportToStorage } from '../utils/consolidatedStorage';

const CustomerODAnalysis = () => {
 const { selectedFiles } = useExcelData();
 
 const [activeTab, setActiveTab] = useState('customers');
 
 // Persistent state for both tabs
 const [tabStates, setTabStates] = useState({
   customers: {
     // Sheet configurations
     sheet: 'Sheet1',
     sheets: [],
     headerRow: 1,
     
     // Column configurations
     columns: [],
     columnSelections: {
       date: '',
       branch: '',
       customer_id: '',
       executive: ''
     },
     
     // Options and filters
     availableMonths: [],
     branchOptions: [],
     executiveOptions: [],
     filters: {
       selectedMonths: [],
       selectedBranches: [],
       selectedExecutives: []
     },
     
     // State management
     results: null,
     loading: false,
     downloadingPpt: false,
     error: null
   },
   od_target: {
     // File selection
     fileChoice: 'os_feb',
     currentFile: null,
     
     // Sheet configurations
     sheet: 'Sheet1',
     sheets: [],
     headerRow: 1,
     
     // Column configurations
     columns: [],
     columnSelections: {
       area: '',
       net_value: '',
       due_date: '',
       executive: ''
     },
     
     // Options and filters
     yearOptions: [],
     branchOptions: [],
     executiveOptions: [],
     filters: {
       selectedYears: [],
       selectedBranches: [],
       selectedExecutives: [],
       tillMonth: 'January'
     },
     
     // State management
     results: null,
     loading: false,
     downloadingPpt: false,
     error: null
   }
 });

 // Helper function to update tab state
 const updateTabState = (tabName, updates) => {
   setTabStates(prev => ({
     ...prev,
     [tabName]: {
       ...prev[tabName],
       ...updates
     }
   }));
 };

 // Helper function to get current tab state
 const getCurrentTabState = () => tabStates[activeTab];
 
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
       {activeTab === 'customers' && (
         <BilledCustomersTab 
           tabState={tabStates.customers}
           updateTabState={(updates) => updateTabState('customers', updates)}
           selectedFiles={selectedFiles}
         />
       )}
       {activeTab === 'od_target' && (
         <ODTargetTab 
           tabState={tabStates.od_target}
           updateTabState={(updates) => updateTabState('od_target', updates)}
           selectedFiles={selectedFiles}
         />
       )}
     </div>
   </div>
 );
};

// Billed Customers Tab Component - WITH PERSISTENT STATE
const BilledCustomersTab = ({ tabState, updateTabState, selectedFiles }) => {
 
 // Fetch available sheet names
 const fetchSheets = async () => {
   if (!selectedFiles.salesFile) return;
   
   try {
     const res = await axios.post('http://localhost:5000/api/branch/sheets', { filename: selectedFiles.salesFile });
     if (res.data && res.data.sheets) {
       updateTabState({ sheets: res.data.sheets });
     } else {
       updateTabState({ sheets: [] });
     }
   } catch (error) {
     console.error('Error fetching sheets:', error);
     updateTabState({ 
       error: `Failed to load sheet names: ${error.response?.data?.error || error.message}`,
       sheets: []
     });
   }
 };
 
 useEffect(() => {
   fetchSheets();
 }, [selectedFiles.salesFile]);
 
 // Fetch columns and auto-map
 const fetchColumns = async () => {
   if (!tabState.sheet) {
     updateTabState({ error: 'Please select a sheet first' });
     return;
   }
   
   updateTabState({ loading: true, error: null });
   
   try {
     // Step 1: Get columns
     const colRes = await axios.post('http://localhost:5000/api/branch/get_columns', {
       filename: selectedFiles.salesFile,
       sheet_name: tabState.sheet,
       header: tabState.headerRow
     });
     
     if (!colRes.data || !colRes.data.columns) {
       updateTabState({ 
         error: 'No columns found in the sheet',
         columns: [],
         loading: false
       });
       return;
     }
     
     updateTabState({ columns: colRes.data.columns });
     
     // Step 2: Auto-map columns
     try {
       const mapRes = await axios.post('http://localhost:5000/api/executive/customer_auto_map_columns', {
         sales_file_path: `uploads/${selectedFiles.salesFile}`
       });
       
       if (mapRes.data && mapRes.data.success && mapRes.data.mapping) {
         updateTabState({ 
           columnSelections: mapRes.data.mapping,
           loading: false
         });
         
         // Step 3: Auto-load options - WAIT for column selections to be set
         setTimeout(async () => {
           await loadFilterOptions(mapRes.data.mapping);
         }, 200);
       } else {
         updateTabState({ 
           error: 'Auto-mapping failed. Please map columns manually.',
           loading: false
         });
       }
     } catch (mapError) {
       updateTabState({ 
         error: 'Auto-mapping failed. Please map columns manually.',
         loading: false
       });
     }
     
   } catch (error) {
     console.error('Error fetching columns:', error);
     updateTabState({ 
       error: `Failed to load columns: ${error.response?.data?.error || error.message}`,
       columns: [],
       loading: false
     });
   }
 };
 
 // Load filter options function
 const loadFilterOptions = async (mappings = null) => {
   const mapping = mappings || tabState.columnSelections;
   
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
       updateTabState({
         availableMonths: res.data.available_months || [],
         branchOptions: res.data.branches || [],
         executiveOptions: res.data.executives || [],
         filters: {
           selectedMonths: res.data.available_months || [],
           selectedBranches: res.data.branches || [],
           selectedExecutives: res.data.executives || []
         },
         error: tabState.error && (tabState.error.includes('map columns') || tabState.error.includes('Auto-mapping')) ? null : tabState.error
       });
     }
   } catch (error) {
     console.error('Error loading filter options:', error);
   }
 };
 
 // Watch for column selection changes and auto-load options
 useEffect(() => {
   const hasAllColumns = tabState.columnSelections.date && tabState.columnSelections.branch && 
                        tabState.columnSelections.customer_id && tabState.columnSelections.executive;
   
   if (hasAllColumns && tabState.columns.length > 0) {
     const timer = setTimeout(() => {
       loadFilterOptions();
     }, 300);
     
     return () => clearTimeout(timer);
   }
 }, [tabState.columnSelections.date, tabState.columnSelections.branch, tabState.columnSelections.customer_id, tabState.columnSelections.executive]);
 
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
     updateTabState({ error: 'Please upload a sales file' });
     return;
   }
   
   if (!tabState.filters.selectedMonths.length) {
     updateTabState({ error: 'Please select at least one month' });
     return;
   }
   
   updateTabState({ loading: true, error: null });
   
   try {
     const payload = {
       sales_file_path: `uploads/${selectedFiles.salesFile}`,
       date_column: tabState.columnSelections.date,
       branch_column: tabState.columnSelections.branch,
       customer_id_column: tabState.columnSelections.customer_id,
       executive_column: tabState.columnSelections.executive,
       selected_months: tabState.filters.selectedMonths,
       selected_branches: tabState.filters.selectedBranches,
       selected_executives: tabState.filters.selectedExecutives
     };
     
     const res = await axios.post('http://localhost:5000/api/executive/calculate_customer_analysis', payload);
     
     if (res.data && res.data.success) {
       updateTabState({ 
         results: res.data.results,
         error: null,
         loading: false
       });
       
       // 🎯 ADD TO CONSOLIDATED REPORTS
       addCustomerReportsToStorage(res.data.results);
       
     } else {
       updateTabState({ 
         error: res.data?.error || 'Failed to generate report',
         loading: false
       });
     }
   } catch (error) {
     console.error('Error generating report:', error);
     updateTabState({ 
       error: `Error generating report: ${error.response?.data?.error || error.message}`,
       loading: false
     });
   }
 };
 
 // Handle download PPT
 const handleDownloadPpt = async (financialYear) => {
   if (!tabState.results || !tabState.results[financialYear]) {
     updateTabState({ error: 'No results available for PPT generation' });
     return;
   }
   
   updateTabState({ downloadingPpt: true, error: null });
   
   try {
     const payload = {
       results_data: { results: tabState.results },
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
     updateTabState({ error: 'Failed to download PowerPoint presentation' });
   } finally {
     updateTabState({ downloadingPpt: false });
   }
 };

 // Handle column selection change
 const handleColumnChange = (columnType, value) => {
   updateTabState({
     columnSelections: {
       ...tabState.columnSelections,
       [columnType]: value
     }
   });
 };

 // Handle filter change
 const handleFilterChange = (filterType, value) => {
   updateTabState({
     filters: {
       ...tabState.filters,
       [filterType]: value
     }
   });
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
     {tabState.error && (
       <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
         {tabState.error}
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
             value={tabState.sheet} 
             onChange={e => updateTabState({ sheet: e.target.value })}
             disabled={tabState.loading}
           >
             <option value="">Select Sheet</option>
             {tabState.sheets.map(sheetName => (
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
             value={tabState.headerRow}
             onChange={e => updateTabState({ headerRow: Number(e.target.value) })}
             disabled={tabState.loading}
           />
         </div>
       </div>
       <button
         className="mt-4 bg-blue-600 text-white px-6 py-2 rounded hover:bg-blue-700 disabled:bg-gray-400"
         onClick={fetchColumns}
         disabled={!tabState.sheet || tabState.loading}
       >
         {tabState.loading ? 'Loading...' : 'Load Columns & Auto-Map'}
       </button>
     </div>
     
     {/* Column Mapping */}
     {tabState.columns.length > 0 && (
       <div className="bg-white p-4 rounded-lg shadow mb-6">
         <h3 className="text-lg font-semibold text-blue-700 mb-4">Column Mapping</h3>
         <div className="grid grid-cols-2 gap-4">
           <div>
             <label className="block font-medium mb-1">Date Column *</label>
             <select
               className="w-full p-2 border border-gray-300 rounded"
               value={tabState.columnSelections.date || ''}
               onChange={(e) => handleColumnChange('date', e.target.value)}
             >
               <option value="">Select Column</option>
               {tabState.columns.map(col => (
                 <option key={col} value={col}>{col}</option>
               ))}
             </select>
           </div>
           <div>
             <label className="block font-medium mb-1">Branch Column *</label>
             <select
               className="w-full p-2 border border-gray-300 rounded"
               value={tabState.columnSelections.branch || ''}
               onChange={(e) => handleColumnChange('branch', e.target.value)}
             >
               <option value="">Select Column</option>
               {tabState.columns.map(col => (
                 <option key={col} value={col}>{col}</option>
               ))}
             </select>
           </div>
           <div>
             <label className="block font-medium mb-1">Customer ID Column *</label>
             <select
               className="w-full p-2 border border-gray-300 rounded"
               value={tabState.columnSelections.customer_id || ''}
               onChange={(e) => handleColumnChange('customer_id', e.target.value)}
             >
               <option value="">Select Column</option>
               {tabState.columns.map(col => (
                 <option key={col} value={col}>{col}</option>
               ))}
             </select>
           </div>
           <div>
             <label className="block font-medium mb-1">Executive Column *</label>
             <select
               className="w-full p-2 border border-gray-300 rounded"
               value={tabState.columnSelections.executive || ''}
               onChange={(e) => handleColumnChange('executive', e.target.value)}
             >
               <option value="">Select Column</option>
               {tabState.columns.map(col => (
                 <option key={col} value={col}>{col}</option>
               ))}
             </select>
           </div>
         </div>
       </div>
     )}
     
     {/* Filters */}
     {(tabState.availableMonths.length > 0 || tabState.branchOptions.length > 0 || tabState.executiveOptions.length > 0) && (
       <div className="bg-white p-4 rounded-lg shadow mb-6">
         <h3 className="text-lg font-semibold text-blue-700 mb-4">Filter Options</h3>
         
         {/* Months Filter */}
         {tabState.availableMonths.length > 0 && (
           <div className="mb-4">
             <label className="block font-semibold mb-3">
               Select Months ({tabState.filters.selectedMonths.length} of {tabState.availableMonths.length})
             </label>
             <div className="max-h-60 overflow-y-auto border border-gray-300 rounded p-3">
               <label className="flex items-center mb-3">
                 <input
                   type="checkbox"
                   checked={tabState.filters.selectedMonths.length === tabState.availableMonths.length}
                   onChange={(e) => {
                     if (e.target.checked) {
                       handleFilterChange('selectedMonths', tabState.availableMonths);
                     } else {
                       handleFilterChange('selectedMonths', []);
                     }
                   }}
                   className="mr-3"
                 />
                 <span className="font-medium text-xm">Select All</span>
               </label>
               {tabState.availableMonths.map(month => (
                 <label key={month} className="flex items-center mb-2">
                   <input
                     type="checkbox"
                     checked={tabState.filters.selectedMonths.includes(month)}
                     onChange={(e) => {
                       if (e.target.checked) {
                         handleFilterChange('selectedMonths', [...tabState.filters.selectedMonths, month]);
                       } else {
                         handleFilterChange('selectedMonths', tabState.filters.selectedMonths.filter(m => m !== month));
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
         
         <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
           {/* Branches Filter */}
           {tabState.branchOptions.length > 0 && (
             <div>
               <label className="block font-semibold mb-3">
                 Branches ({tabState.filters.selectedBranches.length} of {tabState.branchOptions.length})
               </label>
               <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
                 <label className="flex items-center mb-3">
                   <input
                     type="checkbox"
                     checked={tabState.filters.selectedBranches.length === tabState.branchOptions.length}
                     onChange={(e) => {
                       if (e.target.checked) {
                         handleFilterChange('selectedBranches', tabState.branchOptions);
                       } else {
                         handleFilterChange('selectedBranches', []);
                       }
                     }}
                     className="mr-3"
                   />
                   <span className="font-medium text-xm">Select All</span>
                 </label>
                 {tabState.branchOptions.map(branch => (
                   <label key={branch} className="flex items-center mb-2">
                     <input
                       type="checkbox"
                       checked={tabState.filters.selectedBranches.includes(branch)}
                       onChange={(e) => {
                         if (e.target.checked) {
                           handleFilterChange('selectedBranches', [...tabState.filters.selectedBranches, branch]);
                         } else {
                           handleFilterChange('selectedBranches', tabState.filters.selectedBranches.filter(b => b !== branch));
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
           
           {/* Executives Filter */}
           {tabState.executiveOptions.length > 0 && (
             <div>
               <label className="block font-semibold mb-3">
                 Executives ({tabState.filters.selectedExecutives.length} of {tabState.executiveOptions.length})
               </label>
               <div className="max-h-60 overflow-y-auto border border-gray-300 rounded p-3">
                 <label className="flex items-center mb-3">
                   <input
                     type="checkbox"
                     checked={tabState.filters.selectedExecutives.length === tabState.executiveOptions.length}
                     onChange={(e) => {
                       if (e.target.checked) {
                         handleFilterChange('selectedExecutives', tabState.executiveOptions);
                       } else {
                         handleFilterChange('selectedExecutives', []);
                       }
                     }}
                     className="mr-3"
                   />
                   <span className="font-medium text-xm">Select All</span>
                 </label>
                 {tabState.executiveOptions.map(exec => (
                   <label key={exec} className="flex items-center mb-2">
                     <input
                       type="checkbox"
                       checked={tabState.filters.selectedExecutives.includes(exec)}
                       onChange={(e) => {
                         if (e.target.checked) {
                           handleFilterChange('selectedExecutives', [...tabState.filters.selectedExecutives, exec]);
                         } else {
                           handleFilterChange('selectedExecutives', tabState.filters.selectedExecutives.filter(e => e !== exec));
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
         </div>
       </div>
     )}
     
     {/* Generate Report Button */}
     <div className="text-center mb-6">
       <button
         onClick={handleGenerateReport}
         disabled={tabState.loading || !tabState.columns.length || !tabState.filters.selectedMonths.length}
         className="bg-red-600 text-white px-4 py-2 rounded disabled:bg-gray-400"
       >
         {tabState.loading ? 'Generating...' : 'Generate Report'}
       </button>
     </div>
     
     {/* Results */}
     {tabState.results && Object.keys(tabState.results).length > 0 && (
       <div className="bg-white p-6 rounded-lg shadow">
         <h3 className="text-xl font-bold text-blue-700 mb-4">Customer Analysis Results</h3>
         
         {/* Success Message */}
         <div className="bg-green-100 border border-green-400 text-green-700 px-4 py-3 rounded mb-4">
           ✅ Results calculated and automatically added to consolidated reports!
         </div>
         
         {Object.entries(tabState.results).map(([financialYear, data]) => (
           <div key={financialYear} className="mb-8">
             <div className="flex justify-between items-center mb-4">
               <h4 className="text-lg font-semibold text-gray-800">Financial Year: {financialYear}</h4>
               <button
                 onClick={() => handleDownloadPpt(financialYear)}
                 disabled={tabState.downloadingPpt}
                 className="bg-green-600 text-white px-4 py-2 rounded hover:bg-green-700 disabled:bg-gray-400"
               >
                 {tabState.downloadingPpt ? 'Generating PPT...' : 'Download PPT'}
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

// OD Target Tab Component - WITH PERSISTENT STATE
const ODTargetTab = ({ tabState, updateTabState, selectedFiles }) => {
 const monthOptions = [
   'January', 'February', 'March', 'April', 'May', 'June',
   'July', 'August', 'September', 'October', 'November', 'December'
 ];
 
 // Update current file when choice or files change
 useEffect(() => {
   // Reset some states when changing files
   updateTabState({
     error: null,
     sheet: '',
     sheets: [],
     columns: [],
     columnSelections: {
       area: '',
       net_value: '',
       due_date: '',
       executive: ''
     }
   });
   
   // Determine current file based on choice
   let newCurrentFile = null;
   if (tabState.fileChoice === 'os_jan' && selectedFiles.osPrevFile) {
     newCurrentFile = selectedFiles.osPrevFile;
   } else if (tabState.fileChoice === 'os_feb' && selectedFiles.osCurrFile) {
     newCurrentFile = selectedFiles.osCurrFile;
   }
   
   updateTabState({ currentFile: newCurrentFile });
 }, [tabState.fileChoice, selectedFiles.osPrevFile, selectedFiles.osCurrFile]);
 
 // Fetch sheets when file changes
 useEffect(() => {
   if (tabState.currentFile) {
     fetchSheets();
   } else {
     updateTabState({ sheets: [] });
   }
 }, [tabState.currentFile]);
 
 const fetchSheets = async () => {
   if (!tabState.currentFile) return;
   
   updateTabState({ loading: true, error: null });
   
   try {
     const res = await axios.post('http://localhost:5000/api/branch/sheets', { 
       filename: tabState.currentFile 
     });
     
     if (res.data && res.data.sheets && Array.isArray(res.data.sheets)) {
       const newSheet = res.data.sheets.includes('Sheet1') ? 'Sheet1' : res.data.sheets[0] || '';
       updateTabState({ 
         sheets: res.data.sheets,
         sheet: newSheet,
         loading: false
       });
     } else {
       updateTabState({ 
         error: 'No sheets found in the file',
         sheets: [],
         loading: false
       });
     }
   } catch (error) {
     console.error('Error fetching sheets:', error);
     const errorMsg = error.response?.data?.error || error.message;
     updateTabState({ 
       error: `Failed to load sheet names: ${errorMsg}`,
       sheets: [],
       loading: false
     });
   }
 };
 
 // Fetch columns and auto-map
 const fetchColumns = async () => {
   if (!tabState.sheet || !tabState.currentFile) {
     updateTabState({ error: 'Please select a sheet first' });
     return;
   }
   
   updateTabState({ loading: true, error: null });
   
   try {
     // Step 1: Get columns
     const colRes = await axios.post('http://localhost:5000/api/branch/get_columns', {
       filename: tabState.currentFile,
       sheet_name: tabState.sheet,
       header: tabState.headerRow
     });
     
     if (!colRes.data || !colRes.data.columns || !Array.isArray(colRes.data.columns)) {
       updateTabState({ 
         error: 'No columns found in the sheet',
         columns: [],
         loading: false
       });
       return;
     }
     
     updateTabState({ columns: colRes.data.columns });
     
     // Step 2: Auto-map columns
     try {
       const mapRes = await axios.post('http://localhost:5000/api/executive/od_target_auto_map_columns', {
         os_file_path: `uploads/${tabState.currentFile}`
       });
       
       if (mapRes.data && mapRes.data.success && mapRes.data.mapping) {
         updateTabState({ 
           columnSelections: mapRes.data.mapping,
           loading: false
         });
         
         // Step 3: Auto-load options with delay
         setTimeout(async () => {
           await loadFilterOptions(mapRes.data.mapping);
         }, 200);
       } else {
         updateTabState({ 
           error: 'Auto-mapping failed. Please map columns manually.',
           loading: false
         });
       }
     } catch (mapError) {
       updateTabState({ 
         error: 'Auto-mapping failed. Please map columns manually.',
         loading: false
       });
     }
   } catch (error) {
     console.error('Error fetching columns:', error);
     const errorMsg = error.response?.data?.error || error.message;
     updateTabState({ 
       error: `Failed to load columns: ${errorMsg}`,
       columns: [],
       loading: false
     });
   }
 };
 
 // Load filter options function
 const loadFilterOptions = async (mappings = null) => {
   const mapping = mappings || tabState.columnSelections;
   
   // Validate mappings
   const requiredColumns = ['due_date', 'area', 'executive'];
   const missingColumns = requiredColumns.filter(col => !mapping[col]);
   
   if (missingColumns.length > 0 || !tabState.currentFile) {
     console.warn('Missing required columns or file:', missingColumns);
     return;
   }
   
   try {
     const res = await axios.post('http://localhost:5000/api/executive/od_target_get_options', {
       os_file_path: `uploads/${tabState.currentFile}`
     });
     
     if (res.data && res.data.success) {
       updateTabState({
         yearOptions: res.data.years || [],
         branchOptions: res.data.branches || [],
         executiveOptions: res.data.executives || [],
         filters: {
           ...tabState.filters,
           selectedYears: res.data.years || [],
           selectedBranches: res.data.branches || [],
           selectedExecutives: res.data.executives || []
         },
         error: tabState.error && (tabState.error.includes('map columns') || tabState.error.includes('Auto-mapping')) ? null : tabState.error
       });
     }
   } catch (error) {
     console.error('Error loading filter options:', error);
   }
 };
 
 // Watch for column selection changes and auto-load options
 useEffect(() => {
   const hasAllColumns = tabState.columnSelections.area && tabState.columnSelections.net_value && 
                        tabState.columnSelections.due_date && tabState.columnSelections.executive;
   
   if (hasAllColumns && tabState.columns.length > 0 && tabState.currentFile) {
     const timer = setTimeout(() => {
       loadFilterOptions();
     }, 300);
     
     return () => clearTimeout(timer);
   }
 }, [tabState.columnSelections.area, tabState.columnSelections.net_value, tabState.columnSelections.due_date, tabState.columnSelections.executive]);
 
 // Function to add OD target results to consolidated storage
 const addODTargetReportsToStorage = (resultsData) => {
   try {
     const odTargetReports = [{
       df: resultsData.data || [],
       title: `OD Target - ${resultsData.end_date || 'All Periods'}`,
       percent_cols: []
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
   if (!tabState.currentFile) {
     updateTabState({ error: 'Please select an OS file' });
     return;
   }
   
   if (!tabState.filters.selectedYears.length || !tabState.filters.tillMonth) {
     updateTabState({ error: 'Please select at least one year and one month' });
     return;
   }
   
   updateTabState({ loading: true, error: null });
   
   try {
     const payload = {
       os_file_path: `uploads/${tabState.currentFile}`,
       area_column: tabState.columnSelections.area,
       net_value_column: tabState.columnSelections.net_value,
       due_date_column: tabState.columnSelections.due_date,
       executive_column: tabState.columnSelections.executive,
       selected_branches: tabState.filters.selectedBranches,
       selected_years: tabState.filters.selectedYears,
       till_month: tabState.filters.tillMonth,
       selected_executives: tabState.filters.selectedExecutives
     };
     
     const res = await axios.post('http://localhost:5000/api/executive/calculate_od_target', payload);
     
     if (res.data && res.data.success) {
       updateTabState({ 
         results: res.data,
         error: null,
         loading: false
       });
       
       // ADD TO CONSOLIDATED REPORTS
       addODTargetReportsToStorage(res.data);
       
     } else {
       updateTabState({ 
         error: res.data?.error || 'Failed to generate report',
         loading: false
       });
     }
   } catch (error) {
     console.error('Error generating report:', error);
     const errorMsg = error.response?.data?.error || error.message;
     updateTabState({ 
       error: `Error generating report: ${errorMsg}`,
       loading: false
     });
   }
 };
 
 // Handle download PPT
 const handleDownloadPpt = async () => {
   if (!tabState.results) {
     updateTabState({ error: 'No results available for PPT generation' });
     return;
   }
   
   updateTabState({ downloadingPpt: true, error: null });
   
   try {
     const title = `OD Target - ${tabState.results.end_date || 'All Periods'}`;
     
     const payload = {
       results_data: tabState.results,
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
     link.download = `OD_Target_${tabState.results.end_date || 'Report'}.pptx`;
     document.body.appendChild(link);
     link.click();
     link.remove();
     window.URL.revokeObjectURL(url);
     
   } catch (error) {
     console.error('Error downloading PPT:', error);
     updateTabState({ error: 'Failed to download PowerPoint presentation' });
   } finally {
     updateTabState({ downloadingPpt: false });
   }
 };

 // Handle column selection change
 const handleColumnChange = (columnType, value) => {
   updateTabState({
     columnSelections: {
       ...tabState.columnSelections,
       [columnType]: value
     }
   });
 };

 // Handle filter change
 const handleFilterChange = (filterType, value) => {
   updateTabState({
     filters: {
       ...tabState.filters,
       [filterType]: value
     }
   });
 };
 
 return (
   <div>
     {tabState.error && (
       <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
         {tabState.error}
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
             checked={tabState.fileChoice === 'os_jan'}
             onChange={(e) => updateTabState({ fileChoice: e.target.value })}
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
             checked={tabState.fileChoice === 'os_feb'}
             onChange={(e) => updateTabState({ fileChoice: e.target.value })}
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
         <strong>Selected file:</strong> {tabState.currentFile || 'None available'}
       </div>
     </div>
     
     {!tabState.currentFile ? (
       <div className="text-center py-8">
         <p className="text-gray-600">⚠️ No OS file selected or uploaded</p>
         <p className="text-sm text-gray-500 mt-2">
           Please upload the {tabState.fileChoice === 'os_jan' ? 'OS-Previous Month' : 'OS-Current Month'} file first.
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
                 value={tabState.sheet} 
                 onChange={e => updateTabState({ sheet: e.target.value })}
                 disabled={tabState.loading}
               >
                 <option value="">Select Sheet</option>
                 {tabState.sheets.map(sheetName => (
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
                 value={tabState.headerRow}
                 onChange={e => updateTabState({ headerRow: Number(e.target.value) })}
                 disabled={tabState.loading}
               />
             </div>
           </div>
           <button
             className="mt-4 bg-blue-600 text-white px-6 py-2 rounded hover:bg-blue-700 disabled:bg-gray-400"
             onClick={fetchColumns}
             disabled={!tabState.sheet || tabState.loading || !tabState.currentFile}
           >
             {tabState.loading ? 'Loading...' : 'Load Columns & Auto-Map'}
           </button>
         </div>
         
         {/* Column Mapping */}
         {tabState.columns.length > 0 && (
           <div className="bg-white p-4 rounded-lg shadow mb-6">
             <h3 className="text-lg font-semibold text-blue-700 mb-4">Column Mapping</h3>
             <div className="grid grid-cols-2 gap-4">
               <div>
                 <label className="block font-medium mb-1">Area Column *</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={tabState.columnSelections.area || ''}
                   onChange={(e) => handleColumnChange('area', e.target.value)}
                 >
                   <option value="">Select Column</option>
                   {tabState.columns.map(col => (
                     <option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
               <div>
                 <label className="block font-medium mb-1">Net Value Column *</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={tabState.columnSelections.net_value || ''}
                   onChange={(e) => handleColumnChange('net_value', e.target.value)}
                 >
                   <option value="">Select Column</option>
                   {tabState.columns.map(col => (
                     <option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
               <div>
                 <label className="block font-medium mb-1">Due Date Column *</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={tabState.columnSelections.due_date || ''}
                   onChange={(e) => handleColumnChange('due_date', e.target.value)}
                 >
                   <option value="">Select Column</option>
                   {tabState.columns.map(col => (
                     <option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
               <div>
                 <label className="block font-medium mb-1">Executive Column *</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={tabState.columnSelections.executive || ''}
                   onChange={(e) => handleColumnChange('executive', e.target.value)}
                 >
                   <option value="">Select Column</option>
                   {tabState.columns.map(col => (
                     <option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
             </div>
           </div>
         )}
         
         {/* Date Filter */}
         {tabState.yearOptions.length > 0 && (
           <div className="bg-white p-4 rounded-lg shadow mb-6">
             <h3 className="text-lg font-semibold text-blue-700 mb-4">Due Date Filter</h3>
             <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
               <div>
                 <label className="block font-semibold mb-2">Select Years</label>
                 <div className="max-h-32 overflow-y-auto border border-gray-300 rounded p-3">
                   {tabState.yearOptions.map(year => (
                     <label key={year} className="flex items-center mb-2">
                       <input
                         type="checkbox"
                         checked={tabState.filters.selectedYears.includes(year)}
                         onChange={(e) => {
                           if (e.target.checked) {
                             handleFilterChange('selectedYears', [...tabState.filters.selectedYears, year]);
                           } else {
                             handleFilterChange('selectedYears', tabState.filters.selectedYears.filter(y => y !== year));
                           }
                         }}
                         className="mr-3"
                       />
                       <span className="text-xm">{year}</span>
                     </label>
                   ))}
                 </div>
               </div>
               <div>
                 <label className="block font-semibold mb-3">Select Till Month</label>
                 <select
                   className="w-full p-2 border border-gray-300 rounded"
                   value={tabState.filters.tillMonth}
                   onChange={(e) => handleFilterChange('tillMonth', e.target.value)}
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
         {(tabState.branchOptions.length > 0 || tabState.executiveOptions.length > 0) && (
           <div className="bg-white p-4 rounded-lg shadow mb-6">
             <h3 className="text-lg font-semibold text-blue-700 mb-6">Filter Options</h3>
             <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
               {/* Branches Filter */}
               {tabState.branchOptions.length > 0 && (
                 <div>
                   <label className="block font-semibold mb-3">
                     Branches ({tabState.filters.selectedBranches.length} of {tabState.branchOptions.length})
                   </label>
                   <div className="max-h-60 overflow-y-auto border border-gray-300 rounded p-3">
                     <label className="flex items-center mb-3">
                       <input
                         type="checkbox"
                         checked={tabState.filters.selectedBranches.length === tabState.branchOptions.length}
                         onChange={(e) => {
                           if (e.target.checked) {
                             handleFilterChange('selectedBranches', tabState.branchOptions);
                           } else {
                             handleFilterChange('selectedBranches', []);
                           }
                         }}
                         className="mr-3"
                       />
                       <span className="font-medium text-xm">Select All</span>
                     </label>
                     {tabState.branchOptions.map(branch => (
                       <label key={branch} className="flex items-center mb-2">
                         <input
                           type="checkbox"
                           checked={tabState.filters.selectedBranches.includes(branch)}
                           onChange={(e) => {
                             if (e.target.checked) {
                               handleFilterChange('selectedBranches', [...tabState.filters.selectedBranches, branch]);
                             } else {
                               handleFilterChange('selectedBranches', tabState.filters.selectedBranches.filter(b => b !== branch));
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
               
               {/* Executives Filter */}
               {tabState.executiveOptions.length > 0 && (
                 <div>
                   <label className="block font-semibold mb-3">
                     Executives ({tabState.filters.selectedExecutives.length} of {tabState.executiveOptions.length})
                   </label>
                   <div className="max-h-60 overflow-y-auto border border-gray-300 rounded p-3">
                     <label className="flex items-center mb-3">
                       <input
                         type="checkbox"
                         checked={tabState.filters.selectedExecutives.length === tabState.executiveOptions.length}
                         onChange={(e) => {
                           if (e.target.checked) {
                             handleFilterChange('selectedExecutives', tabState.executiveOptions);
                           } else {
                             handleFilterChange('selectedExecutives', []);
                           }
                         }}
                         className="mr-3"
                       />
                       <span className="font-medium text-xm">Select All</span>
                     </label>
                     {tabState.executiveOptions.map(exec => (
                       <label key={exec} className="flex items-center mb-2">
                         <input
                           type="checkbox"
                           checked={tabState.filters.selectedExecutives.includes(exec)}
                           onChange={(e) => {
                             if (e.target.checked) {
                               handleFilterChange('selectedExecutives', [...tabState.filters.selectedExecutives, exec]);
                             } else {
                               handleFilterChange('selectedExecutives', tabState.filters.selectedExecutives.filter(e => e !== exec));
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
             </div>
           </div>
         )}
         
         {/* Generate Report Button */}
         <div className="text-center mb-6">
           <button
             onClick={handleGenerateReport}
             disabled={tabState.loading || !tabState.columns.length || !tabState.filters.selectedYears.length}
             className="bg-green-600 text-white px-8 py-3 rounded-lg text-lg font-semibold hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
           >
             {tabState.loading ? 'Generating...' : 'Generate Report'}
           </button>
         </div>
         
         {/* Results */}
         {tabState.results && (
           <div className="bg-white p-6 rounded-lg shadow">
             <div className="flex justify-between items-center mb-4">
               <h3 className="text-xl font-bold text-blue-700">
                 OD Target Results - {tabState.results.end_date || 'All Periods'}
               </h3>
               <button
                 onClick={handleDownloadPpt}
                 disabled={tabState.downloadingPpt}
                 className="bg-purple-600 text-white px-4 py-2 rounded hover:bg-purple-700 disabled:bg-gray-400"
               >
                 {tabState.downloadingPpt ? 'Generating PPT...' : 'Download PPT'}
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
                     {tabState.results.columns.filter(col => col.toLowerCase() !== 's.no' && col.toLowerCase() !== 'sno').map(col => (
                       <th key={col} className="border border-gray-300 px-4 py-2 text-left font-semibold">
                         {col}
                       </th>
                     ))}
                   </tr>
                 </thead>
                 <tbody>
                   {tabState.results.data.map((row, i) => (
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
                       {tabState.results.columns.filter(col => col.toLowerCase() !== 's.no' && col.toLowerCase() !== 'sno').map((col, j) => (
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
