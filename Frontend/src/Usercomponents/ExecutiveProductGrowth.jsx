import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import { addReportToStorage } from '../utils/consolidatedStorage';

const ExecutiveProductGrowth = () => {
 const { selectedFiles } = useExcelData();

 // Sheet configurations
 const [lySheet, setLySheet] = useState('Sheet1');
 const [cySheet, setCySheet] = useState('Sheet1');
 const [budgetSheet, setBudgetSheet] = useState('Sheet1');
 const [lySheets, setLySheets] = useState([]);
 const [cySheets, setCySheets] = useState([]);
 const [budgetSheets, setBudgetSheets] = useState([]);
 const [lyHeader, setLyHeader] = useState(1);
 const [cyHeader, setCyHeader] = useState(1);
 const [budgetHeader, setBudgetHeader] = useState(1);

 // Column configurations
 const [lyColumns, setLyColumns] = useState([]);
 const [cyColumns, setCyColumns] = useState([]);
 const [budgetColumns, setBudgetColumns] = useState([]);
 const [autoMap, setAutoMap] = useState({});

 // Filters and options
 const [lyMonthOptions, setLyMonthOptions] = useState([]);
 const [cyMonthOptions, setCyMonthOptions] = useState([]);
 const [executiveOptions, setExecutiveOptions] = useState([]);
 const [companyGroupOptions, setCompanyGroupOptions] = useState([]);
 const [productGroupOptions, setProductGroupOptions] = useState([]);
 const [filters, setFilters] = useState({
   lyMonth: '',
   cyMonth: '',
   selectedExecutives: [],
   selectedCompanyGroups: [],
   selectedProductGroups: []
 });

 // State management
 const [results, setResults] = useState(null);
 const [loading, setLoading] = useState(false);
 const [downloadingPpt, setDownloadingPpt] = useState(false);
 const [error, setError] = useState(null);

 // Column selections
 const [columnSelections, setColumnSelections] = useState({
   ly: {},
   cy: {},
   budget: {}
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
   if (selectedFiles.lastYearSalesFile) fetchSheets(selectedFiles.lastYearSalesFile, setLySheets);
   if (selectedFiles.salesFile) fetchSheets(selectedFiles.salesFile, setCySheets);
   if (selectedFiles.budgetFile) fetchSheets(selectedFiles.budgetFile, setBudgetSheets);
 }, [selectedFiles]);

 // Fetch columns after user selects sheet + header
 const fetchColumns = async () => {
   if (!lySheet || !cySheet || !budgetSheet) return;

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

     const [lyCols, cyCols, budgetCols] = await Promise.all([
       getCols(selectedFiles.lastYearSalesFile, lySheet, lyHeader),
       getCols(selectedFiles.salesFile, cySheet, cyHeader),
       getCols(selectedFiles.budgetFile, budgetSheet, budgetHeader)
     ]);

     setLyColumns(lyCols);
     setCyColumns(cyCols);
     setBudgetColumns(budgetCols);

     // Use Product Growth auto-mapping endpoint
     const res = await axios.post('http://localhost:5000/api/executive/product_auto_map_columns', {
       ly_file_path: `uploads/${selectedFiles.lastYearSalesFile}`,
       cy_file_path: `uploads/${selectedFiles.salesFile}`,
       budget_file_path: `uploads/${selectedFiles.budgetFile}`
     });

     const mapping = res.data;

     console.log('Product Growth auto-mapping response:', mapping);

     if (mapping?.ly_mapping && mapping?.cy_mapping && mapping?.budget_mapping) {
       setAutoMap(mapping);
       setColumnSelections({
         ly: { ...mapping.ly_mapping },
         cy: { ...mapping.cy_mapping },
         budget: { ...mapping.budget_mapping }
       });

       await fetchOptions();
     } else {
       console.warn("❌ Product Growth auto-mapping missing required fields:", mapping);
       setError("Auto-mapping failed: check if required columns were found.");
     }
   } catch (error) {
     console.error('Error in fetchColumns:', error);
     setError('Failed to load columns and auto-mapping');
   } finally {
     setLoading(false);
   }
 };

 const fetchOptions = async () => {
   try {
     const res = await axios.post('http://localhost:5000/api/executive/product_get_options', {
       ly_file_path: `uploads/${selectedFiles.lastYearSalesFile}`,
       cy_file_path: `uploads/${selectedFiles.salesFile}`,
       budget_file_path: `uploads/${selectedFiles.budgetFile}`
     });

     console.log('Product Growth options response:', res.data);

     // Handle executives
     const executives = res.data.executives || [];
     setExecutiveOptions(executives);
     
     // Handle company groups
     const companyGroups = res.data.company_groups || [];
     setCompanyGroupOptions(companyGroups);
     
     // Handle product groups
     const productGroups = res.data.product_groups || [];
     setProductGroupOptions(productGroups);
     
     // Handle months
     const lyMonths = res.data.ly_months || [];
     const cyMonths = res.data.cy_months || [];
     setLyMonthOptions(lyMonths);
     setCyMonthOptions(cyMonths);
     
     // Select all options by default
     setFilters(prev => ({
       ...prev,
       selectedExecutives: executives,
       selectedCompanyGroups: companyGroups,
       selectedProductGroups: productGroups,
       lyMonth: lyMonths.length > 0 ? lyMonths[lyMonths.length - 1] : '',
       cyMonth: cyMonths.length > 0 ? cyMonths[cyMonths.length - 1] : ''
     }));
     
     console.log('Processed Product Growth options:', {
       executives: executives.length,
       companyGroups: companyGroups.length,
       productGroups: productGroups.length,
       lyMonths: lyMonths.length,
       cyMonths: cyMonths.length
     });
   } catch (error) {
     console.error('Error fetching Product Growth options:', error);
     setError('Failed to load options');
   }
 };

 // Function to add Product Growth results to consolidated storage - UPDATED VERSION
 const addProductGrowthReportsToStorage = (resultsData) => {
   try {
     const productGrowthReports = [];
     
     // Add streamlit_result data to consolidated storage
     if (resultsData.streamlit_result) {
       Object.entries(resultsData.streamlit_result).forEach(([company, data]) => {
         // Add quantity table with transformed headers for storage
         if (data.qty_df && data.qty_df.length > 0) {
           // Transform the data to use display labels
           const transformedQtyData = data.qty_df.map(row => {
             const newRow = { ...row };
             // Rename keys for display in consolidated reports
             if ('LY_QTY' in newRow) {
               newRow['LAST YEAR QTY'] = newRow['LY_QTY'];
               delete newRow['LY_QTY'];
             }
             if ('CY_QTY' in newRow) {
               newRow['CURRENT YEAR QTY'] = newRow['CY_QTY'];
               delete newRow['CY_QTY'];
             }
             if ('ACHIEVEMENT %' in newRow) {
               newRow['GROWTH %'] = newRow['ACHIEVEMENT %'];
               delete newRow['ACHIEVEMENT %'];
             }
             return newRow;
           });

           productGrowthReports.push({
             df: transformedQtyData,
             title: `${company} - Quantity Growth - Last Year: ${filters.lyMonth} vs Current Year: ${filters.cyMonth}`,
             percent_cols: [4] // GROWTH % column index
           });
         }
         
         // Add value table with transformed headers
         if (data.value_df && data.value_df.length > 0) {
           // Transform the data to use display labels
           const transformedValueData = data.value_df.map(row => {
             const newRow = { ...row };
             // Rename keys for display in consolidated reports
             if ('LY_VALUE' in newRow) {
               newRow['LAST YEAR VALUE'] = newRow['LY_VALUE'];
               delete newRow['LY_VALUE'];
             }
             if ('CY_VALUE' in newRow) {
               newRow['CURRENT YEAR VALUE'] = newRow['CY_VALUE'];
               delete newRow['CY_VALUE'];
             }
             if ('ACHIEVEMENT %' in newRow) {
               newRow['GROWTH %'] = newRow['ACHIEVEMENT %'];
               delete newRow['ACHIEVEMENT %'];
             }
             return newRow;
           });

           productGrowthReports.push({
             df: transformedValueData,
             title: `${company} - Value Growth - Last Year: ${filters.lyMonth} vs Current Year: ${filters.cyMonth}`,
             percent_cols: [4] // GROWTH % column index
           });
         }
       });
     }

     // Add overall summary tables if available (with label transformation)
     if (resultsData.overall_growth_qty && resultsData.overall_growth_qty.length > 0) {
       productGrowthReports.push({
         df: resultsData.overall_growth_qty,
         title: `Overall Growth Quantity Summary - Last Year: ${filters.lyMonth} vs Current Year: ${filters.cyMonth}`,
         percent_cols: [] // Will be dynamically detected based on column names containing '%'
       });
     }

     if (resultsData.overall_growth_value && resultsData.overall_growth_value.length > 0) {
       productGrowthReports.push({
         df: resultsData.overall_growth_value,
         title: `Overall Growth Value Summary - Last Year: ${filters.lyMonth} vs Current Year: ${filters.cyMonth}`,
         percent_cols: [] // Will be dynamically detected based on column names containing '%'
       });
     }

     if (productGrowthReports.length > 0) {
       addReportToStorage(productGrowthReports, 'product_results');
       console.log(`✅ Added ${productGrowthReports.length} Product Growth reports to consolidated storage`);
     }
   } catch (error) {
     console.error('Error adding Product Growth reports to consolidated storage:', error);
   }
 };

 const handleCalculate = async () => {
   if (!selectedFiles.lastYearSalesFile || !selectedFiles.salesFile || !selectedFiles.budgetFile) {
     setError('Please upload all three files (Last Year Sales, Current Year Sales, and Target)');
     return;
   }

   // Validate required columns
   const requiredLyColumns = ['date', 'quantity', 'value', 'company_group', 'product_group', 'executive'];
   const requiredCyColumns = ['date', 'quantity', 'value', 'company_group', 'product_group', 'executive'];
   const requiredBudgetColumns = ['quantity', 'value', 'company_group', 'product_group', 'executive'];
   
   const missingLyColumns = requiredLyColumns.filter(col => 
     !columnSelections.ly[col] || columnSelections.ly[col] === ''
   );
   const missingCyColumns = requiredCyColumns.filter(col => 
     !columnSelections.cy[col] || columnSelections.cy[col] === ''
   );
   const missingBudgetColumns = requiredBudgetColumns.filter(col => 
     !columnSelections.budget[col] || columnSelections.budget[col] === ''
   );
   
   if (missingLyColumns.length > 0) {
     setError(`Missing required Last Year column mappings: ${missingLyColumns.join(', ')}`);
     return;
   }
   
   if (missingCyColumns.length > 0) {
     setError(`Missing required Current Year column mappings: ${missingCyColumns.join(', ')}`);
     return;
   }
   
   if (missingBudgetColumns.length > 0) {
     setError(`Missing required Target column mappings: ${missingBudgetColumns.join(', ')}`);
     return;
   }

   setLoading(true);
   setError(null);

   try {
     const payload = {
       ly_file_path: `uploads/${selectedFiles.lastYearSalesFile}`,
       cy_file_path: `uploads/${selectedFiles.salesFile}`,
       budget_file_path: `uploads/${selectedFiles.budgetFile}`,
       
       // LY column mappings
       ly_date: columnSelections.ly.date,
       ly_quantity: columnSelections.ly.quantity,
       ly_value: columnSelections.ly.value,
       ly_company_group: columnSelections.ly.company_group,
       ly_product_group: columnSelections.ly.product_group,
       ly_executive: columnSelections.ly.executive,
       ly_sl_code: columnSelections.ly.sl_code,
       
       // CY column mappings
       cy_date: columnSelections.cy.date,
       cy_quantity: columnSelections.cy.quantity,
       cy_value: columnSelections.cy.value,
       cy_company_group: columnSelections.cy.company_group,
       cy_product_group: columnSelections.cy.product_group,
       cy_executive: columnSelections.cy.executive,
       cy_sl_code: columnSelections.cy.sl_code,
       
       // Budget column mappings
       budget_quantity: columnSelections.budget.quantity,
       budget_value: columnSelections.budget.value,
       budget_company_group: columnSelections.budget.company_group,
       budget_product_group: columnSelections.budget.product_group,
       budget_executive: columnSelections.budget.executive,
       budget_sl_code: columnSelections.budget.sl_code,
       
       // Filters
       selected_executives: filters.selectedExecutives,
       selected_company_groups: filters.selectedCompanyGroups,
       selected_product_groups: filters.selectedProductGroups,
       ly_month: filters.lyMonth,
       cy_month: filters.cyMonth
     };

     console.log('Product Growth calculate payload:', payload);

     const res = await axios.post('http://localhost:5000/api/executive/calculate_product_growth', payload);
     
     console.log('Raw response from server:', res.data);
     
     if (res.data.success) {
       setResults(res.data);
       setError(null);
       console.log('Product Growth calculation successful!');
       console.log('product_growth_results:', res.data.product_growth_results);
       console.log('overall_growth_qty:', res.data.overall_growth_qty);
       console.log('overall_growth_value:', res.data.overall_growth_value);
       
       // 🎯 ADD TO CONSOLIDATED REPORTS - Just like ExecutiveODC!
       addProductGrowthReportsToStorage(res.data);
       
     } else {
       setError(res.data.error || 'Failed to calculate results');
     }
   } catch (error) {
     console.error('Error calculating Product Growth:', error);
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
     const monthTitle = `Last Year: ${filters.lyMonth} vs Current Year: ${filters.cyMonth}`;
     
     const payload = {
       results_data: {
         product_growth_results: results.product_growth_results,
         overall_growth_qty: results.overall_growth_qty,
         overall_growth_value: results.overall_growth_value,
         streamlit_result: results.streamlit_result // Include streamlit format for PPT
       },
       month_title: monthTitle,
       logo_file: null // Add logo support later if needed
     };

     console.log('PPT generation payload:', payload);

     const response = await axios.post('http://localhost:5000/api/executive/generate_product_growth_ppt', payload, {
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
     link.download = `Product_Growth_${monthTitle.replace(/[^a-zA-Z0-9]/g, '_')}.pptx`;
     document.body.appendChild(link);
     link.click();
     link.remove();
     window.URL.revokeObjectURL(url);

     console.log('PPT downloaded successfully');
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

 // Helper function to render Streamlit-style tables per company - UPDATED VERSION
 const renderStreamlitStyleTables = () => {
   if (!results?.streamlit_result) return null;

   // Define the correct column order for display with ORIGINAL DATA KEYS
   // We keep the original keys but display different labels
   const qtyColumnOrder = ['PRODUCT GROUP', 'LY_QTY', 'TARGET_QTY', 'CY_QTY', 'ACHIEVEMENT %'];
   const valueColumnOrder = ['PRODUCT GROUP', 'LY_VALUE', 'TARGET_VALUE', 'CY_VALUE', 'ACHIEVEMENT %'];

   // Create display label mapping
   const getDisplayLabel = (col) => {
     const labelMap = {
       'LY_QTY': 'LAST YEAR QTY',
       'LY_VALUE': 'LAST YEAR VALUE',
       'CY_QTY': 'CURRENT YEAR QTY',
       'CY_VALUE': 'CURRENT YEAR VALUE',
       'ACHIEVEMENT %': 'GROWTH %'
     };
     return labelMap[col] || col;
   };

   return Object.entries(results.streamlit_result).map(([company, data]) => (
     <div key={company} className="mb-8">
       <h4 className="text-xl font-bold text-blue-700 mb-4">{company}</h4>
       
       {/* Quantity Table */}
       <div className="mb-6">
         <h5 className="text-lg font-semibold text-gray-800 mb-3">{company} - Quantity Growth</h5>
         <div className="overflow-x-auto">
           <table className="min-w-full table-auto border-collapse border border-gray-300">
             <thead>
               <tr className="bg-blue-600 text-white">
                 {qtyColumnOrder.map(col => (
                   <th key={col} className="border border-gray-300 px-4 py-2 text-left font-semibold">
                     {getDisplayLabel(col)}
                   </th>
                 ))}
               </tr>
             </thead>
             <tbody>
               {data.qty_df && data.qty_df.map((row, i) => (
                 <tr 
                   key={i} 
                   className={`
                     ${row['PRODUCT GROUP'] === 'TOTAL' 
                       ? 'bg-gray-200 font-bold border-t-2 border-gray-400' 
                       : i % 2 === 0 
                         ? 'bg-gray-50' 
                         : 'bg-white'
                     } hover:bg-blue-50
                   `}
                 >
                   {qtyColumnOrder.map((col, j) => (
                     <td key={j} className="border border-gray-300 px-4 py-2">
                       {col.includes('ACHIEVEMENT') && col.includes('%') 
                         ? `${Number(row[col] || 0).toFixed(2)}%` 
                         : typeof row[col] === 'number' 
                           ? Number(row[col] || 0).toLocaleString() 
                           : row[col] || ''
                       }
                     </td>
                   ))}
                 </tr>
               ))}
             </tbody>
           </table>
         </div>
       </div>

       {/* Value Table */}
       <div className="mb-6">
         <h5 className="text-lg font-semibold text-gray-800 mb-3">{company} - Value Growth</h5>
         <div className="overflow-x-auto">
           <table className="min-w-full table-auto border-collapse border border-gray-300">
             <thead>
               <tr className="bg-blue-600 text-white">
                 {valueColumnOrder.map(col => (
                   <th key={col} className="border border-gray-300 px-4 py-2 text-left font-semibold">
                     {getDisplayLabel(col)}
                   </th>
                 ))}
               </tr>
             </thead>
             <tbody>
               {data.value_df && data.value_df.map((row, i) => (
                 <tr 
                   key={i} 
                   className={`
                     ${row['PRODUCT GROUP'] === 'TOTAL' 
                       ? 'bg-gray-200 font-bold border-t-2 border-gray-400' 
                       : i % 2 === 0 
                         ? 'bg-gray-50' 
                         : 'bg-white'
                     } hover:bg-blue-50
                   `}
                 >
                   {valueColumnOrder.map((col, j) => (
                     <td key={j} className="border border-gray-300 px-4 py-2">
                       {col.includes('ACHIEVEMENT') && col.includes('%') 
                         ? `${Number(row[col] || 0).toFixed(2)}%` 
                         : typeof row[col] === 'number' 
                           ? Number(row[col] || 0).toLocaleString() 
                           : row[col] || ''
                       }
                     </td>
                   ))}
                 </tr>
               ))}
             </tbody>
           </table>
         </div>
       </div>
     </div>
   ));
 };

 return (
   <div className="p-6">
     <h2 className="text-2xl font-bold text-blue-800 mb-6">Executive Product Growth Analysis</h2>

     {error && (
       <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
         {error}
       </div>
     )}

     {/* Sheet Selection */}
     <div className="bg-white p-4 rounded-lg shadow mb-6">
       <h3 className="text-lg font-semibold text-blue-700 mb-4">Sheet Configuration</h3>
       <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
         <div>
           <label className="block font-semibold mb-2">Last Year Sales Sheet</label>
           <select 
             className="w-full p-2 border border-gray-300 rounded" 
             value={lySheet} 
             onChange={e => setLySheet(e.target.value)}
           >
             <option value="">Select Sheet</option>
             {lySheets.map(sheet => <option key={sheet} value={sheet}>{sheet}</option>)}
           </select>
           <label className="block mt-4 font-semibold mb-2">Last Year Header Row</label>
           <input 
             type="number" 
             className="w-full p-2 border border-gray-300 rounded" 
             min={1} 
             value={lyHeader}
             onChange={e => setLyHeader(Number(e.target.value))} 
           />
         </div>
         <div>
           <label className="block font-semibold mb-2">Current Sales Sheet</label>
           <select 
             className="w-full p-2 border border-gray-300 rounded" 
             value={cySheet} 
             onChange={e => setCySheet(e.target.value)}
           >
             <option value="">Select Sheet</option>
             {cySheets.map(sheet => <option key={sheet} value={sheet}>{sheet}</option>)}
           </select>
           <label className="block mt-4 font-semibold mb-2">Current Sales Header Row</label>
           <input 
             type="number" 
             className="w-full p-2 border border-gray-300 rounded" 
             min={1} 
             value={cyHeader}
             onChange={e => setCyHeader(Number(e.target.value))} 
           />
         </div>
         <div>
           <label className="block font-semibold mb-2">Target Sheet</label>
           <select 
             className="w-full p-2 border border-gray-300 rounded" 
             value={budgetSheet} 
             onChange={e => setBudgetSheet(e.target.value)}
           >
             <option value="">Select Sheet</option>
             {budgetSheets.map(sheet => <option key={sheet} value={sheet}>{sheet}</option>)}
           </select>
           <label className="block mt-4 font-semibold mb-2">Target Header Row</label>
           <input 
             type="number" 
             className="w-full p-2 border border-gray-300 rounded" 
             min={1} 
             value={budgetHeader}
             onChange={e => setBudgetHeader(Number(e.target.value))} 
           />
         </div>
       </div>
       <button
         className="mt-4 bg-blue-600 text-white px-6 py-2 rounded hover:bg-blue-700 disabled:bg-gray-400"
         onClick={fetchColumns}
         disabled={!lySheet || !cySheet || !budgetSheet || loading}
       >
         {loading ? 'Loading...' : 'Load Columns & Auto-Map'}
       </button>
     </div>

     {/* Column Mapping */}
     {lyColumns.length > 0 && cyColumns.length > 0 && budgetColumns.length > 0 && (
       <div className="bg-white p-4 rounded-lg shadow mb-6">
         <h3 className="text-lg font-semibold text-blue-700 mb-4">Column Mapping</h3>
         
         {/* Last Year Sales Columns */}
         <div className="mb-6">
           <h4 className="text-md font-semibold text-gray-700 mb-3">Last Year Sales Columns</h4>
           <div className="grid grid-cols-3 gap-4">
             {[
               { key: 'date', label: 'Date *', required: true },
               { key: 'quantity', label: 'Quantity *', required: true },
               { key: 'value', label: 'Value *', required: true },
               { key: 'company_group', label: 'Company Group *', required: true },
               { key: 'product_group', label: 'Product Group *', required: true },
               { key: 'executive', label: 'Executive *', required: true },
               { key: 'sl_code', label: 'SL Code', required: false }
             ].map(({ key, label, required }) => (
               <div key={key}>
                 <label className="block font-medium mb-1">
                   {label} {required && <span className="text-red-500">*</span>}
                 </label>
                 <select
                   className={`w-full p-2 border rounded ${required && !columnSelections.ly[key] ? 'border-red-300' : 'border-gray-300'}`}
                   value={columnSelections.ly[key] || ''}
                   onChange={(e) =>
                     setColumnSelections(prev => ({
                       ...prev,
                       ly: { ...prev.ly, [key]: e.target.value }
                     }))
                   }
                 >
                   <option value="">Select Column</option>
                   {lyColumns.map(col => (
                     <option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
             ))}
           </div>
         </div>

         {/* Current Sales Columns */}
         <div className="mb-6">
           <h4 className="text-md font-semibold text-gray-700 mb-3">Current Sales Columns</h4>
           <div className="grid grid-cols-3 gap-4">
             {[
               { key: 'date', label: 'Date *', required: true },
               { key: 'quantity', label: 'Quantity *', required: true },
               { key: 'value', label: 'Value *', required: true },
               { key: 'company_group', label: 'Company Group *', required: true },
               { key: 'product_group', label: 'Product Group *', required: true },
               { key: 'executive', label: 'Executive *', required: true },
               { key: 'sl_code', label: 'SL Code', required: false }
             ].map(({ key, label, required }) => (
               <div key={key}>
                 <label className="block font-medium mb-1">
                   {label} {required && <span className="text-red-500">*</span>}
                 </label>
                 <select
                   className={`w-full p-2 border rounded ${required && !columnSelections.cy[key] ? 'border-red-300' : 'border-gray-300'}`}
                   value={columnSelections.cy[key] || ''}
                   onChange={(e) =>
                     setColumnSelections(prev => ({
                       ...prev,
                       cy: { ...prev.cy, [key]: e.target.value }
                     }))
                   }
                 >
                   <option value="">Select Column</option>
                   {cyColumns.map(col => (<option key={col} value={col}>{col}</option>
                   ))}
                 </select>
               </div>
             ))}
           </div>
         </div>

         {/* Budget Columns */}
         <div>
           <h4 className="text-md font-semibold text-gray-700 mb-3"> Target Columns</h4>
           <div className="grid grid-cols-3 gap-4">
             {[
               { key: 'quantity', label: 'Quantity *', required: true },
               { key: 'value', label: 'Value *', required: true },
               { key: 'company_group', label: 'Company Group *', required: true },
               { key: 'product_group', label: 'Product Group *', required: true },
               { key: 'executive', label: 'Executive *', required: true },
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
                   {budgetColumns.map(col => (
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
     {(lyMonthOptions.length > 0 || cyMonthOptions.length > 0 || executiveOptions.length > 0 || 
       companyGroupOptions.length > 0 || productGroupOptions.length > 0) && (
       <div className="bg-white p-4 rounded-lg shadow mb-6">
         <h3 className="text-lg font-semibold text-blue-700 mb-4">Filters</h3>
         <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
           {/* Last Year Month Filter */}
           {lyMonthOptions.length > 0 && (
             <div>
               <label className="block font-semibold mb-2">Last Year Month</label>
               <select
                 className="w-full p-2 border border-gray-300 rounded"
                 value={filters.lyMonth}
                 onChange={(e) => setFilters(prev => ({ ...prev, lyMonth: e.target.value }))}
               >
                 <option value="">Select Month</option>
                 {lyMonthOptions.map(month => (
                   <option key={month} value={month}>{month}</option>
                 ))}
               </select>
             </div>
           )}

           {/* Current Year Month Filter */}
           {cyMonthOptions.length > 0 && (
             <div>
               <label className="block font-semibold mb-2">Current Sales Month</label>
               <select
                 className="w-full p-2 border border-gray-300 rounded"
                 value={filters.cyMonth}
                 onChange={(e) => setFilters(prev => ({ ...prev, cyMonth: e.target.value }))}
               >
                 <option value="">Select Month</option>
                 {cyMonthOptions.map(month => (
                   <option key={month} value={month}>{month}</option>
                 ))}
               </select>
             </div>
           )}

           {/* Executive Filter */}
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
                         setFilters(prev => ({ ...prev, selectedExecutives: executiveOptions }));
                       } else {
                         setFilters(prev => ({ ...prev, selectedExecutives: [] }));
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

           {/* Company Group Filter */}
           {companyGroupOptions.length > 0 && (
             <div>
               <label className="block font-semibold mb-2">
                 Company Groups ({filters.selectedCompanyGroups.length} of {companyGroupOptions.length})
               </label>
               <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
                 <label className="flex items-center mb-2">
                   <input
                     type="checkbox"
                     checked={filters.selectedCompanyGroups.length === companyGroupOptions.length}
                     onChange={(e) => {
                       if (e.target.checked) {
                         setFilters(prev => ({ ...prev, selectedCompanyGroups: companyGroupOptions }));
                       } else {
                         setFilters(prev => ({ ...prev, selectedCompanyGroups: [] }));
                       }
                     }}
                     className="mr-2"
                   />
                   <span className="font-medium text-sm">Select All</span>
                 </label>
                 {companyGroupOptions.map(group => (
                   <label key={group} className="flex items-center mb-1">
                     <input
                       type="checkbox"
                       checked={filters.selectedCompanyGroups.includes(group)}
                       onChange={(e) => {
                         if (e.target.checked) {
                           setFilters(prev => ({
                             ...prev,
                             selectedCompanyGroups: [...prev.selectedCompanyGroups, group]
                           }));
                         } else {
                           setFilters(prev => ({
                             ...prev,
                             selectedCompanyGroups: prev.selectedCompanyGroups.filter(g => g !== group)
                           }));
                         }
                       }}
                       className="mr-2"
                     />
                     <span className="text-xs">{group}</span>
                   </label>
                 ))}
               </div>
             </div>
           )}

           {/* Product Group Filter */}
           {productGroupOptions.length > 0 && (
             <div>
               <label className="block font-semibold mb-2">
                 Product Groups ({filters.selectedProductGroups.length} of {productGroupOptions.length})
               </label>
               <div className="max-h-40 overflow-y-auto border border-gray-300 rounded p-2">
                 <label className="flex items-center mb-2">
                   <input
                     type="checkbox"
                     checked={filters.selectedProductGroups.length === productGroupOptions.length}
                     onChange={(e) => {
                       if (e.target.checked) {
                         setFilters(prev => ({ ...prev, selectedProductGroups: productGroupOptions }));
                       } else {
                         setFilters(prev => ({ ...prev, selectedProductGroups: [] }));
                       }
                     }}
                     className="mr-2"
                   />
                   <span className="font-medium text-sm">Select All</span>
                 </label>
                 {productGroupOptions.map(group => (
                   <label key={group} className="flex items-center mb-1">
                     <input
                       type="checkbox"
                       checked={filters.selectedProductGroups.includes(group)}
                       onChange={(e) => {
                         if (e.target.checked) {
                           setFilters(prev => ({
                             ...prev,
                             selectedProductGroups: [...prev.selectedProductGroups, group]
                           }));
                         } else {
                           setFilters(prev => ({
                             ...prev,
                             selectedProductGroups: prev.selectedProductGroups.filter(g => g !== group)
                           }));
                         }
                       }}
                       className="mr-2"
                     />
                     <span className="text-xs">{group}</span>
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
         disabled={loading || !lyColumns.length || !cyColumns.length || !budgetColumns.length}
         className="bg-green-600 text-white px-8 py-3 rounded-lg text-lg font-semibold hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
       >
         {loading ? 'Calculating...' : 'Calculate Product Growth'}
       </button>
     </div>

     {/* Results */}
     {results && (
       <div className="bg-white p-6 rounded-lg shadow">
         <h3 className="text-xl font-bold text-blue-700 mb-4">
           Product Growth Analysis Results - {results.ly_month} vs {results.cy_month}
         </h3>

         {/* Success Message */}
         <div className="bg-green-100 border border-green-400 text-green-700 px-4 py-3 rounded mb-4">
           ✅ Results calculated and automatically added to consolidated reports!
         </div>

         {/* Download PPT Button */}
         <div className="text-center mb-6">
           <button
             onClick={handleDownloadPpt}
             disabled={downloadingPpt || !results}
             className="bg-purple-600 text-white px-6 py-3 rounded-lg text-lg font-semibold hover:bg-purple-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
           >
             {downloadingPpt ? 'Generating PPT...' : 'Download PowerPoint'}
           </button>
         </div>

         {/* Display results in Streamlit format - separate tables per company */}
         {renderStreamlitStyleTables()}

         {/* Overall Summary Tables */}
         {results.overall_growth_qty && results.overall_growth_qty.length > 0 && (
           <div className="mb-8">
             <h4 className="text-lg font-semibold text-gray-800 mb-3">Overall Growth Quantity Summary</h4>
             <div className="overflow-x-auto">
               <table className="min-w-full table-auto border-collapse border border-gray-300">
                 <thead>
                   <tr className="bg-blue-600 text-white">
                     {Object.keys(results.overall_growth_qty[0]).map(col => (
                       <th key={col} className="border border-gray-300 px-4 py-2 text-left font-semibold">
                         {col}
                       </th>
                     ))}
                   </tr>
                 </thead>
                 <tbody>
                   {results.overall_growth_qty.map((row, i) => (
                     <tr 
                       key={i} 
                       className={`${i % 2 === 0 ? 'bg-gray-50' : 'bg-white'} hover:bg-blue-50`}
                     >
                       {Object.entries(row).map(([col, value], j) => (
                         <td key={j} className="border border-gray-300 px-4 py-2">
                           {col.includes('Growth') && col.includes('%') 
                             ? `${Number(value).toFixed(2)}%` 
                             : col.includes('Budget vs Actual') && col.includes('%')
                               ? `${Number(value).toFixed(2)}%`
                             : typeof value === 'number' 
                               ? Number(value).toLocaleString() 
                               : value
                           }
                         </td>
                       ))}
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

export default ExecutiveProductGrowth;
