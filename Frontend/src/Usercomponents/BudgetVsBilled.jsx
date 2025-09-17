// components/BudgetVsBilled.jsx
import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import ExecSelector from './ExecSelector';
import BranchSelector from './BranchSelector';
import SearchableSelect from './SearchableSelect'; // Import SearchableSelect
import { addReportToStorage } from '../utils/consolidatedStorage';

const BudgetVsBilled = () => {
  const { selectedFiles } = useExcelData();

  const [salesSheet, setSalesSheet] = useState('Sheet1');
  const [budgetSheet, setBudgetSheet] = useState('Sheet1');
  const [salesSheets, setSalesSheets] = useState([]);
  const [budgetSheets, setBudgetSheets] = useState([]);
  const [salesHeader, setSalesHeader] = useState(1);
  const [budgetHeader, setBudgetHeader] = useState(1);
  const [salesColumns, setSalesColumns] = useState([]);
  const [budgetColumns, setBudgetColumns] = useState([]);
  const [autoMap, setAutoMap] = useState({});
  const [monthOptions, setMonthOptions] = useState([]);
  const [filters, setFilters] = useState({
    selectedMonth: '',
    selectedSalesExecs: [],
    selectedBudgetExecs: [],
    selectedBranches: []
  });
  const [salesExecList, setSalesExecList] = useState([]);
  const [budgetExecList, setBudgetExecList] = useState([]);
  const [branchList, setBranchList] = useState([]);
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadingColumns, setLoadingColumns] = useState(false);
  const [loadingReport, setLoadingReport] = useState(false);
  
  // Only keep proof generation loading state
  const [loadingProofGeneration, setLoadingProofGeneration] = useState(false);

  // Fetch available sheet names from backend
  const fetchSheets = async (filename, setter) => {
    const res = await axios.post('http://localhost:5000/api/branch/sheets', { filename });
    setter(res.data.sheets);
  };

  useEffect(() => {
    if (selectedFiles.salesFile) fetchSheets(selectedFiles.salesFile, setSalesSheets);
    if (selectedFiles.budgetFile) fetchSheets(selectedFiles.budgetFile, setBudgetSheets);
  }, [selectedFiles]);

  const [columnSelections, setColumnSelections] = useState({
  sales: {},
  budget: {}
  });

  // Fetch columns after user selects sheet + header
  const fetchColumns = async () => {
    if (!salesSheet || !budgetSheet) return;
    setLoadingColumns(true);
    try{
    const getCols = async (filename, sheet_name, header) => {
      const res = await axios.post('http://localhost:5000/api/branch/get_columns', {
        filename,
        sheet_name,
        header
      });
      return res.data.columns || [];
    };

    const [salesCols, budgetCols] = await Promise.all([
      getCols(selectedFiles.salesFile, salesSheet, salesHeader),
      getCols(selectedFiles.budgetFile, budgetSheet, budgetHeader)
    ]);

    setSalesColumns(salesCols);
    setBudgetColumns(budgetCols);

    const res = await axios.post('http://localhost:5000/api/branch/auto_map_columns', {
  sales_columns: salesCols,
  budget_columns: budgetCols
});

const mapping = res.data;

if (
  mapping?.sales_mapping?.executive &&
  mapping?.budget_mapping?.executive &&
  mapping?.sales_mapping?.date
) {
  setAutoMap(mapping);
  setColumnSelections({
  sales: { ...mapping.sales_mapping },
  budget: { ...mapping.budget_mapping }
});

  await fetchExecAndBranches(mapping);
  await fetchMonths(mapping.sales_mapping.date);
} else {
  console.warn("❌ Auto-mapping missing required fields:", mapping);
  alert("Auto-mapping failed: check if 'date' or 'executive' columns were found.");
}
    }catch (err) {
      console.error('Error in fetchcolumn:', err);
    }finally{
      setLoadingColumns(false);
    }
  };

  const fetchExecAndBranches = async (autoMapData) => {
  const salesExec = autoMapData?.sales_mapping?.executive || '';
  const budgetExec = autoMapData?.budget_mapping?.executive || '';
  const salesArea = autoMapData?.sales_mapping?.area || '';
  const budgetArea = autoMapData?.budget_mapping?.area || '';

  const res = await axios.post('http://localhost:5000/api/branch/get_exec_branch_options', {
    sales_filename: selectedFiles.salesFile,
    budget_filename: selectedFiles.budgetFile,
    sales_sheet: salesSheet,
    budget_sheet: budgetSheet,
    sales_header: salesHeader,
    budget_header: budgetHeader,
    sales_exec_col: salesExec,
    budget_exec_col: budgetExec,
    sales_area_col: salesArea,
    budget_area_col: budgetArea
  });

  setSalesExecList(res.data.sales_executives);
  setBudgetExecList(res.data.budget_executives);
  setBranchList(res.data.branches);
};


const fetchMonths = async (dateCol) => {
  if (!dateCol) return;

  const res = await axios.post('http://localhost:5000/api/branch/extract_months', {
    sales_filename: selectedFiles.salesFile,
    sales_sheet: salesSheet,
    sales_header: salesHeader,
    sales_date_col: dateCol
  });

  const months = res.data.months || [];
  setMonthOptions(months);

  // Auto-select first month if available and not already selected
  if (months.length > 0 && !filters.selectedMonth) {
    setFilters((prev) => ({ ...prev, selectedMonth: months[0] }));
  }
};


const addToConsolidatedReports = (resultsData) => {
    try {
      const monthTitle = filters.selectedMonth || 'All Months';

      // Apply consistent column formatting with units - matching individual reports
      const formatDataWithUnits = (data, columns, reportType) => {
        if (!data || !Array.isArray(data)) return [];
        
        return data.map(row => {
          const formattedRow = { ...row };
          
          // Apply consistent column naming based on report type
          if (reportType === 'budget_qty') {
            // Rename columns to match individual report format
            if (formattedRow['Budget Qty'] !== undefined) {
              formattedRow['Budget Qty/Mt'] = formattedRow['Budget Qty'];
              delete formattedRow['Budget Qty'];
            }
            if (formattedRow['Billed Qty'] !== undefined) {
              formattedRow['Billed Qty/Mt'] = formattedRow['Billed Qty'];
              delete formattedRow['Billed Qty'];
            }
          } else if (reportType === 'budget_value') {
            if (formattedRow['Budget Value'] !== undefined) {
              formattedRow['Budget Value/L'] = formattedRow['Budget Value'];
              delete formattedRow['Budget Value'];
            }
            if (formattedRow['Billed Value'] !== undefined) {
              formattedRow['Billed Value/L'] = formattedRow['Billed Value'];
              delete formattedRow['Billed Value'];
            }
          } else if (reportType === 'overall_qty') {
            if (formattedRow['Budget Qty'] !== undefined) {
              formattedRow['Budget Qty/Mt'] = formattedRow['Budget Qty'];
              delete formattedRow['Budget Qty'];
            }
            if (formattedRow['Billed Qty'] !== undefined) {
              formattedRow['Billed Qty/Mt'] = formattedRow['Billed Qty'];
              delete formattedRow['Billed Qty'];
            }
          } else if (reportType === 'overall_value') {
            if (formattedRow['Budget Value'] !== undefined) {
              formattedRow['Budget Value/L'] = formattedRow['Budget Value'];
              delete formattedRow['Budget Value'];
            }
            if (formattedRow['Billed Value'] !== undefined) {
              formattedRow['Billed Value/L'] = formattedRow['Billed Value'];
              delete formattedRow['Billed Value'];
            }
          }
          
          return formattedRow;
        });
      };

      // Define proper column orders with units - matching individual reports
      const columnOrderMap = {
        budget_qty: ['Area', 'Budget Qty/Mt', 'Billed Qty/Mt', '%'],
        budget_value: ['Area', 'Budget Value/L', 'Billed Value/L', '%'],
        overall_qty: ['Area', 'Budget Qty/Mt', 'Billed Qty/Mt'],
        overall_value: ['Area', 'Budget Value/L', 'Billed Value/L']
      };

      const branchReports = [
        {
          df: formatDataWithUnits(resultsData.budget_vs_billed_qty.data || [], 
                                resultsData.budget_vs_billed_qty.columns || [], 
                                'budget_qty'),
          columns: columnOrderMap.budget_qty,
          title: `BUDGET VS BILLED - QUANTITY - ${monthTitle.toUpperCase()}`,
          percent_cols: [3]
        },
        {
          df: formatDataWithUnits(resultsData.budget_vs_billed_value.data || [], 
                                resultsData.budget_vs_billed_value.columns || [], 
                                'budget_value'),
          columns: columnOrderMap.budget_value,
          title: `BUDGET VS BILLED - VALUE - ${monthTitle.toUpperCase()}`,
          percent_cols: [3]
        },
        {
          df: formatDataWithUnits(resultsData.overall_sales_qty.data || [], 
                                resultsData.overall_sales_qty.columns || [], 
                                'overall_qty'),
          columns: columnOrderMap.overall_qty,
          title: `OVERALL SALES - QUANTITY - ${monthTitle.toUpperCase()}`,
          percent_cols: []
        },
        {
          df: formatDataWithUnits(resultsData.overall_sales_value.data || [], 
                                resultsData.overall_sales_value.columns || [], 
                                'overall_value'),
          columns: columnOrderMap.overall_value,
          title: `OVERALL SALES - VALUE - ${monthTitle.toUpperCase()}`,
          percent_cols: []
        }
      ];

      // Store with consistent formatting
      addReportToStorage(branchReports, 'branch_budget_results');

      console.log(`Added ${branchReports.length} branch reports to consolidated storage with proper column formatting and units`);
    } catch (error) {
      console.error('Error adding branch reports to consolidated storage:', error);
    }
  };


const handleCalculate = async () => {
  setLoadingReport(true);
    const payload = {
      sales_filename: selectedFiles.salesFile,
      budget_filename: selectedFiles.budgetFile,
      sales_sheet: salesSheet,
      budget_sheet: budgetSheet,
      sales_header: salesHeader,
      budget_header: budgetHeader,
      selected_month: filters.selectedMonth,
      selected_sales_execs: filters.selectedSalesExecs,
      selected_budget_execs: filters.selectedBudgetExecs,
      selected_branches: filters.selectedBranches,

      sales_date_col: columnSelections.sales.date,
      sales_value_col: columnSelections.sales.value,
      sales_qty_col: columnSelections.sales.quantity,
      sales_area_col: columnSelections.sales.area,
      sales_product_group_col: columnSelections.sales.product_group,
      sales_sl_code_col: columnSelections.sales.sl_code,
      sales_exec_col: columnSelections.sales.executive,

      budget_value_col: columnSelections.budget.value,
      budget_qty_col: columnSelections.budget.quantity,
      budget_area_col: columnSelections.budget.area,
      budget_product_group_col: columnSelections.budget.product_group,
      budget_sl_code_col: columnSelections.budget.sl_code,
      budget_exec_col: columnSelections.budget.executive
    };

    try{
    const res = await axios.post('http://localhost:5000/api/branch/calculate_budget_vs_billed', payload);
    if (res.data && res.data.budget_vs_billed_qty) {
      setResults(res.data);
      addToConsolidatedReports(res.data);
    } else {
      console.warn("No valid data returned from backend.");
    }
  } catch (err) {
    console.error("❌ Error calculating Branch Budget vs Billed:", err);
  } finally{
    setLoadingReport(false);
  }
  };

  // SIMPLIFIED PROOF OF CALCULATION - Only generation function

  const generateProofOfCalculation = async () => {
    if (!areColumnsValid()) {
      alert('Please ensure all required columns are selected before generating proof.');
      return;
    }

    setLoadingProofGeneration(true);
    try {
      const payload = createProofPayload();
      const res = await axios.post(
        'http://localhost:5000/api/branch/generate_proof_of_calculation',
        payload,
        { responseType: 'blob' }
      );

      const blob = new Blob([res.data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Proof_of_Calculation_${filters.selectedMonth.replace(' ', '_')}.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);

      console.log('✅ Proof of calculation downloaded successfully');
    } catch (err) {
      console.error('❌ Error generating proof of calculation:', err);
      alert('Failed to generate proof of calculation. Please check your inputs and try again.');
    } finally {
      setLoadingProofGeneration(false);
    }
  };

  const createProofPayload = () => {
    return {
      sales_filename: selectedFiles.salesFile,
      budget_filename: selectedFiles.budgetFile,
      sales_sheet: salesSheet,
      budget_sheet: budgetSheet,
      sales_header: salesHeader,
      budget_header: budgetHeader,
      selected_month: filters.selectedMonth,
      selected_executives: [...new Set([...filters.selectedSalesExecs, ...filters.selectedBudgetExecs])],
      selected_branches: filters.selectedBranches,

      sales_date_col: columnSelections.sales.date,
      sales_value_col: columnSelections.sales.value,
      sales_qty_col: columnSelections.sales.quantity,
      sales_area_col: columnSelections.sales.area,
      sales_product_group_col: columnSelections.sales.product_group,
      sales_sl_code_col: columnSelections.sales.sl_code,
      sales_exec_col: columnSelections.sales.executive,

      budget_value_col: columnSelections.budget.value,
      budget_qty_col: columnSelections.budget.quantity,
      budget_area_col: columnSelections.budget.area,
      budget_product_group_col: columnSelections.budget.product_group,
      budget_sl_code_col: columnSelections.budget.sl_code,
      budget_exec_col: columnSelections.budget.executive
    };
  };

  const areColumnsValid = () => {
    const requiredSalesColumns = ['date', 'value', 'quantity', 'area', 'product_group', 'sl_code', 'executive'];
    const requiredBudgetColumns = ['value', 'quantity', 'area', 'product_group', 'sl_code', 'executive'];

    const salesValid = requiredSalesColumns.every(col => columnSelections.sales[col]);
    const budgetValid = requiredBudgetColumns.every(col => columnSelections.budget[col]);

    return salesValid && budgetValid && filters.selectedMonth;
  };
  
  return (
  <div>
    <h2 className="text-xl font-bold text-blue-800 mb-4">Budget vs Billed Report</h2>

    {/* Sheet Selection - Keep as regular select */}
    <div className="grid grid-cols-2 gap-6 mb-6">
      <div>
        <label className="block font-semibold mb-1">Sales Sheet Name</label>
        <select className="w-full p-2 border" value={salesSheet} onChange={e => setSalesSheet(e.target.value)}>
          <option value="">Select</option>
          {salesSheets.map(sheet => <option key={sheet}>{sheet}</option>)}
        </select>

        <label className="block mt-4 font-semibold mb-1">Sales Header Row (1-based)</label>
        <input type="number" className="w-full p-2 border" min={1} value={salesHeader}
          onChange={e => setSalesHeader(Number(e.target.value))} />
      </div>

      <div>
        <label className="block font-semibold mb-1">Budget Sheet Name</label>
        <select className="w-full p-2 border" value={budgetSheet} onChange={e => setBudgetSheet(e.target.value)}>
          <option value="">Select</option>
          {budgetSheets.map(sheet => <option key={sheet}>{sheet}</option>)}
        </select>

        <label className="block mt-4 font-semibold mb-1">Budget Header Row (1-based)</label>
        <input type="number" className="w-full p-2 border" min={1} value={budgetHeader}
          onChange={e => setBudgetHeader(Number(e.target.value))} />
      </div>
    </div>

    {/* Load Columns */}
    {!(salesColumns.length > 0 && budgetColumns.length > 0) && (
      <div className="mb-4">
        <button
          className="bg-blue-600 text-white px-6 py-2 rounded hover:bg-blue-700 disabled:bg-gray-400"
          onClick={fetchColumns}
          disabled={!salesSheet || !budgetSheet || loadingColumns}
        >
          {loadingColumns ? 'Loading...' : 'Load Columns & Auto Map'}
        </button>
      </div>
    )}

    {salesColumns.length > 0 && budgetColumns.length > 0 && (
      <>
  <div className="mt-6">
    <h3 className="text-blue-700 font-bold text-lg mb-4">Sales Column Mapping</h3>
    <div className="grid grid-cols-3 gap-4">
      <div>
        <label className="block font-semibold mb-1">Sales Date</label>
        <SearchableSelect
          options={salesColumns}
          value={columnSelections.sales.date || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, date: value }
            }))
          }
          placeholder="Select Date Column"
          className="w-full p-2 border"
        />

        <label className="block font-semibold mt-2 mb-1">Sales Area</label>
        <SearchableSelect
          options={salesColumns}
          value={columnSelections.sales.area || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, area: value }
            }))
          }
          placeholder="Select Area Column"
          className="w-full p-2 border"
        />
      </div>

      <div>
        <label className="block font-semibold mb-1">Sales Value</label>
        <SearchableSelect
          options={salesColumns}
          value={columnSelections.sales.value || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, value: value }
            }))
          }
          placeholder="Select Value Column"
          className="w-full p-2 border"
        />

        <label className="block font-semibold mt-2 mb-1">Sales Quantity</label>
        <SearchableSelect
          options={salesColumns}
          value={columnSelections.sales.quantity || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, quantity: value }
            }))
          }
          placeholder="Select Quantity Column"
          className="w-full p-2 border"
        />
      </div>

      <div>
        <label className="block font-semibold mb-1">Sales Product Group</label>
        <SearchableSelect
          options={salesColumns}
          value={columnSelections.sales.product_group || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, product_group: value }
            }))
          }
          placeholder="Select Product Group Column"
          className="w-full p-2 border"
        />

        <label className="block font-semibold mt-2 mb-1">Sales SL Code</label>
        <SearchableSelect
          options={salesColumns}
          value={columnSelections.sales.sl_code || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, sl_code: value }
            }))
          }
          placeholder="Select SL Code Column"
          className="w-full p-2 border"
        />

        <label className="block font-semibold mt-2 mb-1">Sales Executive</label>
        <SearchableSelect
          options={salesColumns}
          value={columnSelections.sales.executive || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              sales: { ...prev.sales, executive: value }
            }))
          }
          placeholder="Select Executive Column"
          className="w-full p-2 border"
        />
      </div>
    </div>

    <h3 className="text-blue-700 font-bold text-lg mt-6 mb-4">Budget Column Mapping</h3>
    <div className="grid grid-cols-3 gap-4">
      <div>
        <label className="block font-semibold mb-1">Budget Area</label>
        <SearchableSelect
          options={budgetColumns}
          value={columnSelections.budget.area || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, area: value }
            }))
          }
          placeholder="Select Area Column"
          className="w-full p-2 border"
        />

        <label className="block font-semibold mt-2 mb-1">Budget Value</label>
        <SearchableSelect
          options={budgetColumns}
          value={columnSelections.budget.value || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, value: value }
            }))
          }
          placeholder="Select Value Column"
          className="w-full p-2 border"
        />
      </div>

      <div>
        <label className="block font-semibold mb-1">Budget Quantity</label>
        <SearchableSelect
          options={budgetColumns}
          value={columnSelections.budget.quantity || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, quantity: value }
            }))
          }
          placeholder="Select Quantity Column"
          className="w-full p-2 border"
        />

        <label className="block font-semibold mt-2 mb-1">Budget Product Group</label>
        <SearchableSelect
          options={budgetColumns}
          value={columnSelections.budget.product_group || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, product_group: value }
            }))
          }
          placeholder="Select Product Group Column"
          className="w-full p-2 border"
        />
      </div>

      <div>
        <label className="block font-semibold mb-1">Budget SL Code</label>
        <SearchableSelect
          options={budgetColumns}
          value={columnSelections.budget.sl_code || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, sl_code: value }
            }))
          }
          placeholder="Select SL Code Column"
          className="w-full p-2 border"
        />

        <label className="block font-semibold mt-2 mb-1">Budget Executive</label>
        <SearchableSelect
          options={budgetColumns}
          value={columnSelections.budget.executive || ''}
          onChange={(value) =>
            setColumnSelections(prev => ({
              ...prev,
              budget: { ...prev.budget, executive: value }
            }))
          }
          placeholder="Select Executive Column"
          className="w-full p-2 border"
        />
      </div>
    </div>
  </div>

<ExecSelector
  salesExecList={salesExecList}
  budgetExecList={budgetExecList}
  onChange={({ sales, budget }) => {
    setFilters(prev => ({
      ...prev,
      selectedSalesExecs: sales,
      selectedBudgetExecs: budget
    }));
  }}
/>
<BranchSelector
  branchList={branchList}
  onChange={(selected) =>
    setFilters(prev => ({ ...prev, selectedBranches: selected }))
  }
/>
{/* Keep month selection as regular select */}
{monthOptions.length > 0 && (
  <div className="mt-6">
    <label className="block font-semibold mb-1">Select Month</label>
    <select
      className="w-full p-2 border"
      value={filters.selectedMonth}
      onChange={(e) => setFilters((prev) => ({ ...prev, selectedMonth: e.target.value }))}
    >
      {monthOptions.map(month => (
        <option key={month} value={month}>{month}</option>
      ))}
    </select>
  </div>
)}

{/* Simplified Action Buttons - Only 2 buttons now */}
<div className="mt-6 space-y-4">
  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
    <button
      onClick={handleCalculate}
      className="bg-red-600 text-white px-6 py-2 rounded hover:bg-red-700 disabled:bg-gray-400"
      disabled={loadingReport || !areColumnsValid()}
    >
      {loadingReport ? 'Generating...' : 'Generate Budget vs Billed Report'}
    </button>

    <button
      onClick={generateProofOfCalculation}
      className="bg-yellow-600 text-white px-6 py-2 rounded hover:bg-yellow-700 disabled:bg-gray-400"
      disabled={loadingProofGeneration || !areColumnsValid()}
    >
      {loadingProofGeneration ? 'Generating...' : 'Download Proof of Calculation'}
    </button>
  </div>
</div>

{results && (
  <div className="mt-8">
    <h3 className="text-lg font-bold text-blue-700 mb-2">Results</h3>

    {['Qty', 'Value', 'OverallQty', 'OverallValue'].map((type, index) => {
      const labelMap = {
        Qty: 'Budget vs Billed Quantity',
        Value: 'Budget vs Billed Value',
        OverallQty: 'Overall Sales Quantity',
        OverallValue: 'Overall Sales Value'
      };
      const resultKeyMap = {
        Qty: 'budget_vs_billed_qty',
        Value: 'budget_vs_billed_value',
        OverallQty: 'overall_sales_qty',
        OverallValue: 'overall_sales_value'
      };
      const resultBlock = results[resultKeyMap[type]] || {};
      const rows = resultBlock.data || [];
      const orderedCols = resultBlock.columns || [];
      return (
        <div key={type} className="mt-6">
          <h4 className="font-semibold mb-1">{labelMap[type]}</h4>
          <div className="overflow-x-auto">
            <table className="table-auto w-full border text-sm">
              <thead>
                <tr>
                  {orderedCols.map(col => (
                    <th key={col} className="border px-2 py-1 bg-gray-200">{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, i) => (
                  <tr key={i}>
                    {orderedCols.map(col => (
                      <td key={col} className="border px-2 py-1">{row[col]}</td>
                    ))}
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

{results && (
  <div className="mt-6">
    <button
      className="bg-green-600 text-white px-6 py-2 rounded hover:bg-green-700"
      onClick={async () => {
        try {
          const res = await axios.post(
            'http://localhost:5000/api/branch/download_ppt',
            {
              month_title: filters.selectedMonth, // Changed from 'month' to 'month_title'
              budget_vs_billed_qty: results.budget_vs_billed_qty.data || [], // Extract data array
              budget_vs_billed_value: results.budget_vs_billed_value.data || [], // Extract data array
              overall_sales_qty: results.overall_sales_qty.data || [], // Extract data array
              overall_sales_value: results.overall_sales_value.data || [], // Extract data array
              logo_file: null // Add logo_file if needed
            },
            { responseType: 'blob' }
          );

          const blob = new Blob([res.data], {
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
          });
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `Budget_vs_Billed_${filters.selectedMonth.replace(' ', '_')}.pptx`;
          a.click();
          window.URL.revokeObjectURL(url);
        } catch (error) {
          console.error('Error downloading PPT:', error);
          alert('Failed to download PPT. Please try again.');
        }
      }}
    >
      Download Budget VS Billed PPT
    </button>
  </div>
)}
</>
)}
    </div>
  );
};

export default BudgetVsBilled;
