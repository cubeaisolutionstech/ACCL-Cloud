import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import SearchableSelect from './SearchableSelect'; // Import SearchableSelect
import { addReportToStorage } from '../utils/consolidatedStorage';

const ProductGrowth = () => {
  const { selectedFiles } = useExcelData();
  const [loadingFilters, setLoadingFilters] = useState(false);
  const [loadingReport, setLoadingReport] = useState(false);

  const [sheets, setSheets] = useState({ ly: [], cy: [], budget: [] });
  const [selectedSheet, setSelectedSheet] = useState({ ly: 'Sheet1', cy: 'Sheet1', budget: 'Sheet1' });
  const [headers, setHeaders] = useState({ ly: 1, cy: 1, budget: 1 });

  const [columns, setColumns] = useState({ ly: {}, cy: {}, budget: {} });
  const [allCols, setAllCols] = useState({ ly: [], cy: [], budget: [] });
  const [mappingsOverride, setMappingsOverride] = useState({ ly: {}, cy: {}, budget: {} });

  const [availableMonths, setAvailableMonths] = useState({ ly: [], cy: [] });
  const [selectedMonths, setSelectedMonths] = useState({ ly: [], cy: [] });

  const [executives, setExecutives] = useState([]);
  const [companyGroups, setCompanyGroups] = useState([]);
  const [selectedExecutives, setSelectedExecutives] = useState([]);
  const [selectedGroups, setSelectedGroups] = useState([]);

  const [groupResults, setGroupResults] = useState({});

  useEffect(() => {
    if (selectedFiles.lastYearSalesFile && selectedFiles.salesFile && selectedFiles.budgetFile) {
      fetchSheetNames();
    }
  }, [selectedFiles]);

  const fetchSheetNames = async () => {
    const load = async (filename) => {
      const res = await axios.post('http://localhost:5000/api/branch/sheets', { filename });
      return res.data.sheets || [];
    };
    const [lySheets, cySheets, budgetSheets] = await Promise.all([
      load(selectedFiles.lastYearSalesFile),
      load(selectedFiles.salesFile),
      load(selectedFiles.budgetFile)
    ]);
    setSheets({ ly: lySheets, cy: cySheets, budget: budgetSheets });
  };

  const handleAutoMap = async () => {
    setLoadingFilters(true);
    try {
      const fetchColumns = async (filename, sheet, header) => {
        const res = await axios.post('http://localhost:5000/api/branch/get_columns', {
          filename,
          sheet_name: sheet,
          header
        });
        return res.data.columns || [];
      };

      const [lyCols, cyCols, budgetCols] = await Promise.all([
        fetchColumns(selectedFiles.lastYearSalesFile, selectedSheet.ly, headers.ly),
        fetchColumns(selectedFiles.salesFile, selectedSheet.cy, headers.cy),
        fetchColumns(selectedFiles.budgetFile, selectedSheet.budget, headers.budget)
      ]);

      const res = await axios.post('http://localhost:5000/api/branch/auto_map_product_growth', {
        ly_columns: lyCols,
        cy_columns: cyCols,
        budget_columns: budgetCols
      });

      const newMappings = {
        ly: res.data.ly_mapping,
        cy: res.data.cy_mapping,
        budget: res.data.budget_mapping
      };

      setColumns(newMappings);
      setAllCols({ ly: lyCols, cy: cyCols, budget: budgetCols });
      setMappingsOverride(newMappings);
    } catch (error) {
      console.error('Error in handleAutoMap:', error);
      alert('Failed to load columns. Please check your file selections.');
    } finally {
      setLoadingFilters(false);
    }
  };

  useEffect(() => {
    const hasValidMappings = mappingsOverride.ly?.date && mappingsOverride.cy?.date && mappingsOverride.budget?.executive;

    if (
      selectedFiles.lastYearSalesFile &&
      selectedFiles.salesFile &&
      selectedFiles.budgetFile &&
      selectedSheet.ly &&
      selectedSheet.cy &&
      selectedSheet.budget &&
      hasValidMappings
    ) {
      fetchFilters();
    }
  }, [mappingsOverride, selectedSheet, selectedFiles]);

  const standardizeName = (name) => {
    if (!name) return "";
    let cleaned = name.trim().toLowerCase().replace(/[^a-zA-Z0-9\s]/g, '');
    cleaned = cleaned.replace(/\s+/g, ' ');
    const titleCase = cleaned.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');

    const generalVariants = ['general', 'gen', 'generals', 'general ', 'genral', 'generl'];
    if (generalVariants.includes(cleaned)) {
      return 'General';
    }
    return titleCase;
  };

  useEffect(() => {
    if (availableMonths.ly.length > 0) {
      const defaultLY = availableMonths.ly.includes("0") ? ["0"] : [];
      setSelectedMonths((prev) => ({ ...prev, ly: defaultLY }));
    }

    if (availableMonths.cy.length > 0) {
      setSelectedMonths((prev) => ({ ...prev, cy: [availableMonths.cy[0]] }));
    }
  }, [availableMonths]);

  const fetchFilters = async () => {
    setLoadingFilters(true);
    const payload = {
      ly_filename: selectedFiles.lastYearSalesFile,
      ly_sheet: selectedSheet.ly,
      ly_header: headers.ly,
      ly_date_col: mappingsOverride.ly?.date,
      ly_exec_col: mappingsOverride.ly?.executive,
      ly_group_col: mappingsOverride.ly?.company_group,

      cy_filename: selectedFiles.salesFile,
      cy_sheet: selectedSheet.cy,
      cy_header: headers.cy,
      cy_date_col: mappingsOverride.cy?.date,
      cy_exec_col: mappingsOverride.cy?.executive,
      cy_group_col: mappingsOverride.cy?.company_group,

      budget_filename: selectedFiles.budgetFile,
      budget_sheet: selectedSheet.budget,
      budget_header: headers.budget,
      budget_exec_col: mappingsOverride.budget?.executive,
      budget_group_col: mappingsOverride.budget?.company_group
    };

    console.log("fetchFilters payload:", payload);

    try {
      const res = await axios.post('http://localhost:5000/api/branch/get_product_growth_filters', payload);
      console.log("Filters response:", res.data);

      setAvailableMonths({ ly: res.data.ly_months, cy: res.data.cy_months });
      setExecutives(res.data.executives);
      setCompanyGroups(Array.from(new Set(res.data.company_groups.map(standardizeName))).sort());
      setSelectedExecutives(res.data.executives);
      setSelectedGroups(res.data.company_groups);
    } catch (error) {
      console.error("Failed to fetch filters:", error);
      alert("Error loading month and filter options. Check console.");
    } finally {
      setLoadingFilters(false);
    }
  };

  const addProductGrowthToConsolidatedReports = (resultsData) => {
    try {
      console.log("Adding Product Growth to consolidated storage:", resultsData);
      
      // Apply consistent column formatting with units for Product Growth reports
      const formatProductGrowthDataWithUnits = (data, reportType) => {
        if (!data || !Array.isArray(data)) {
          console.warn(`Invalid data provided for ${reportType}:`, data);
          return [];
        }
        
        console.log(`Formatting ${reportType} data with ${data.length} rows:`, data.slice(0, 2)); // Log first 2 rows
        
        return data.map(row => {
          const formattedRow = { ...row };
          
          // Ensure all numeric values are properly formatted and apply column mappings
          if (reportType === 'quantity') {
            // Define expected columns and verify they exist
            const expectedCols = ['PRODUCT NAME', 'LAST YEAR QTY/MT', 'BUDGET QTY/MT', 'CURRENT YEAR QTY/MT', 'GROWTH %'];
            
            // Column mapping to ensure consistent naming with units
            const columnMap = {
              'PRODUCT NAME': 'PRODUCT NAME',
              'LAST YEAR QTY/MT': 'LAST YEAR QTY/MT',
              'LAST_YEAR_QTY/MT': 'LAST YEAR QTY/MT',
              'BUDGET QTY/MT': 'BUDGET QTY/MT', 
              'BUDGET_QTY/MT': 'BUDGET QTY/MT',
              'CURRENT YEAR QTY/MT': 'CURRENT YEAR QTY/MT',
              'CURRENT_YEAR_QTY/MT': 'CURRENT YEAR QTY/MT',
              'GROWTH %': 'GROWTH %'
            };
            
            // Check for missing columns and warn
            expectedCols.forEach(col => {
              const hasColumn = Object.keys(formattedRow).some(key => 
                columnMap[key] === col || key === col
              );
              if (!hasColumn) {
                console.warn(`Missing column in quantity data: ${col}`);
              }
            });
            
            // Apply column mapping
            Object.keys(columnMap).forEach(oldKey => {
              if (formattedRow[oldKey] !== undefined) {
                formattedRow[columnMap[oldKey]] = formattedRow[oldKey];
                if (oldKey !== columnMap[oldKey]) {
                  delete formattedRow[oldKey];
                }
              }
            });
            
          } else if (reportType === 'value') {
            // Define expected columns and verify they exist
            const expectedCols = ['PRODUCT NAME', 'LAST YEAR VALUE/L', 'BUDGET VALUE/L', 'CURRENT YEAR VALUE/L', 'GROWTH %'];
            
            // Column mapping to ensure consistent naming with units
            const columnMap = {
              'PRODUCT NAME': 'PRODUCT NAME',
              'LAST YEAR VALUE/L': 'LAST YEAR VALUE/L',
              'LAST_YEAR_VALUE/L': 'LAST YEAR VALUE/L',
              'BUDGET VALUE/L': 'BUDGET VALUE/L',
              'BUDGET_VALUE/L': 'BUDGET VALUE/L',
              'CURRENT YEAR VALUE/L': 'CURRENT YEAR VALUE/L', 
              'CURRENT_YEAR_VALUE/L': 'CURRENT YEAR VALUE/L',
              'GROWTH %': 'GROWTH %'
            };
            
            // Check for missing columns and warn
            expectedCols.forEach(col => {
              const hasColumn = Object.keys(formattedRow).some(key => 
                columnMap[key] === col || key === col
              );
              if (!hasColumn) {
                console.warn(`Missing column in value data: ${col}`);
              }
            });
            
            // Apply column mapping
            Object.keys(columnMap).forEach(oldKey => {
              if (formattedRow[oldKey] !== undefined) {
                formattedRow[columnMap[oldKey]] = formattedRow[oldKey];
                if (oldKey !== columnMap[oldKey]) {
                  delete formattedRow[oldKey];
                }
              }
            });
          }
          
          return formattedRow;
        });
      };

      // Define proper column orders with units - matching individual reports
      const columnOrderMap = {
        quantity: ['PRODUCT NAME', 'LAST YEAR QTY/MT', 'BUDGET QTY/MT', 'CURRENT YEAR QTY/MT', 'GROWTH %'],
        value: ['PRODUCT NAME', 'LAST YEAR VALUE/L', 'BUDGET VALUE/L', 'CURRENT YEAR VALUE/L', 'GROWTH %']
      };

      const productGrowthReports = [];

      // Validate resultsData structure
      if (!resultsData || typeof resultsData !== 'object') {
        console.error("Invalid resultsData provided:", resultsData);
        return;
      }

      // Process each company group's results
      Object.keys(resultsData).forEach(groupName => {
        const groupData = resultsData[groupName];
        console.log(`Processing group ${groupName}:`, groupData);
        
        // Validate group data structure
        if (!groupData || typeof groupData !== 'object') {
          console.warn(`Invalid group data for ${groupName}:`, groupData);
          return;
        }
        
        // Process quantity data
        if (groupData.qty_df && Array.isArray(groupData.qty_df) && groupData.qty_df.length > 0) {
          console.log(`Adding quantity report for ${groupName}, rows: ${groupData.qty_df.length}`);
          const formattedQtyData = formatProductGrowthDataWithUnits(groupData.qty_df, 'quantity');
          
          if (formattedQtyData.length > 0) {
            productGrowthReports.push({
              df: formattedQtyData,
              columns: columnOrderMap.quantity,
              title: `${groupName.toUpperCase()} - QUANTITY GROWTH`,
              percent_cols: [4] // GROWTH % column index
            });
          }
        } else {
          console.log(`No valid quantity data for ${groupName}`);
        }
        
        // Process value data
        if (groupData.value_df && Array.isArray(groupData.value_df) && groupData.value_df.length > 0) {
          console.log(`Adding value report for ${groupName}, rows: ${groupData.value_df.length}`);
          const formattedValueData = formatProductGrowthDataWithUnits(groupData.value_df, 'value');
          
          if (formattedValueData.length > 0) {
            productGrowthReports.push({
              df: formattedValueData,
              columns: columnOrderMap.value,
              title: `${groupName.toUpperCase()} - VALUE GROWTH`,
              percent_cols: [4] // GROWTH % column index
            });
          }
        } else {
          console.log(`No valid value data for ${groupName}`);
        }
      });

      // Add reports to storage if any were created
      if (productGrowthReports.length > 0) {
        addReportToStorage(productGrowthReports, 'product_growth_results');
        console.log(`✅ Successfully added ${productGrowthReports.length} Product Growth reports to consolidated storage`);
      } else {
        console.warn("⚠️ No Product Growth reports to add to consolidated storage - no valid data found");
      }
      
    } catch (error) {
      console.error('❌ Error adding Product Growth reports to consolidated storage:', error);
      // Optionally re-throw or handle the error based on your application's error handling strategy
      throw new Error(`Failed to add Product Growth reports: ${error.message}`);
    }
  };

  const handleGenerateReport = async () => {
    setLoadingReport(true);
    try {
      const payload = {
        ly_filename: selectedFiles.lastYearSalesFile,
        ly_sheet: selectedSheet.ly,
        ly_header: headers.ly,
        cy_filename: selectedFiles.salesFile,
        cy_sheet: selectedSheet.cy,
        cy_header: headers.cy,
        budget_filename: selectedFiles.budgetFile,
        budget_sheet: selectedSheet.budget,
        budget_header: headers.budget,

        ly_months: selectedMonths.ly,
        cy_months: selectedMonths.cy,

        ly_date_col: mappingsOverride.ly.date,
        cy_date_col: mappingsOverride.cy.date,
        ly_qty_col: mappingsOverride.ly.quantity,
        cy_qty_col: mappingsOverride.cy.quantity,
        ly_value_col: mappingsOverride.ly.value,
        cy_value_col: mappingsOverride.cy.value,

        budget_qty_col: mappingsOverride.budget.quantity,
        budget_value_col: mappingsOverride.budget.value,

        ly_product_col: mappingsOverride.ly.product_group,
        cy_product_col: mappingsOverride.cy.product_group,
        budget_product_group_col: mappingsOverride.budget.product_group,

        ly_company_group_col: mappingsOverride.ly.company_group,
        cy_company_group_col: mappingsOverride.cy.company_group,
        budget_company_group_col: mappingsOverride.budget.company_group,

        ly_exec_col: mappingsOverride.ly.executive,
        cy_exec_col: mappingsOverride.cy.executive,
        budget_exec_col: mappingsOverride.budget.executive,

        selected_executives: selectedExecutives,
        selected_company_groups: selectedGroups
      };

      console.log("📤 Product Growth API Payload:", payload);

      const res = await axios.post("http://localhost:5000/api/branch/calculate_product_growth", payload);
      console.log("📥 Product Growth API Response:", res.data);
      
      if (res && res.data && res.data.status === 'success' && res.data.results) {
        setGroupResults({ results: res.data.results });
        
        // Add to consolidated storage
        addProductGrowthToConsolidatedReports(res.data.results);
        
        console.log("✅ Product Growth results processed successfully");
      } else {
        console.error("❌ Invalid response structure:", res.data);
        alert("No results received or invalid response format.");
      }
    } catch (error) {
      console.error("❌ Failed to calculate product growth:", error);
      const errorMsg = error?.response?.data?.error || error.message || "Unknown error";
      alert(`Error generating product growth report: ${errorMsg}`);
    } finally {
      setLoadingReport(false);
    }
  };

  const handleDownloadPPT = async () => {
    try {
      const filteredGroups = groupResults?.results
        ? Object.fromEntries(
            Object.entries(groupResults.results).filter(
              ([_, v]) => v && typeof v === "object" && v.qty_df && v.value_df
            )
          )
        : {};

      const payload = {
        group_results: filteredGroups,
        month_title: "Product Growth"  // Clean title without LY/CY month details
      };

      const response = await axios.post(
        "http://localhost:5000/api/branch/download_product_growth_ppt",
        payload,
        { responseType: "blob" }
      );

      const blob = new Blob([response.data], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      });

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Product_Growth_Report.pptx`;
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error("PPT download failed:", error);
      alert("PPT generation failed. Check console.");
    }
  };

  // Helper function to get the appropriate label for each file type
  const getFileTypeLabel = (key, context = 'sheet') => {
    const labels = {
      ly: {
        sheet: 'Last Year Sales',
        mapping: 'Last Year Sales',
        month: 'Last Year Months'
      },
      cy: {
        sheet: 'Current Month Sales',
        mapping: 'Current Month Sales', 
        month: 'Current Month'
      },
      budget: {
        sheet: 'Budget',
        mapping: 'Budget',
        month: 'Budget'
      }
    };
    return labels[key]?.[context] || key;
  };

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold text-blue-800 mb-4">Product Growth</h2>
      
      {/* Sheet Selection Grid - Keep as regular select */}
      <div className="grid grid-cols-3 gap-4 mb-6">
        {['ly', 'cy', 'budget'].map((key) => (
          <div key={key}>
            <label className="font-semibold block mb-1">
              {getFileTypeLabel(key, 'sheet')} Sheet
            </label>
            <select
              className="w-full p-2 border"
              value={selectedSheet[key]}
              onChange={(e) => setSelectedSheet(prev => ({ ...prev, [key]: e.target.value }))}
            >
              <option value="">Select</option>
              {sheets[key].map(s => <option key={s}>{s}</option>)}
            </select>

            <label className="block mt-2">Header Row</label>
            <input
              type="number"
              className="w-full p-2 border"
              value={headers[key]}
              min={1}
              onChange={(e) => setHeaders(prev => ({ ...prev, [key]: Number(e.target.value) }))}
            />
          </div>
        ))}
      </div>

      {/* Load Columns & Auto Map Button */}
      {Object.keys(mappingsOverride.ly || {}).length === 0 && (
        <div className="mb-4">
          <button
            onClick={handleAutoMap}
            className="bg-blue-600 text-white px-4 py-2 rounded flex items-center justify-center min-w-[220px] hover:bg-blue-700 disabled:bg-gray-400"
            disabled={loadingFilters}
          >
            {loadingFilters ? 'Loading...' : 'Load Columns & Auto Map'}
          </button>
        </div>
      )}

      {/* Column Mapping Preview - Updated with SearchableSelect */}
      {Object.keys(mappingsOverride.ly || {}).length > 0 && (
        <>
          <div className="space-y-10">
            {['ly', 'cy', 'budget'].map((key) => (
              <div key={key}>
                <h3 className="text-lg font-bold text-blue-700 mb-4">
                  {getFileTypeLabel(key, 'mapping')} Mapping
                </h3>
                <div className="grid grid-cols-3 gap-4">
                  {Object.entries(columns[key]).map(([colKey, label]) => (
                    <div key={colKey}>
                      <label className="block font-semibold mb-1 capitalize">
                        {colKey.replace(/_/g, ' ')}
                      </label>
                      <SearchableSelect
                        options={allCols[key] || []}
                        value={mappingsOverride[key]?.[colKey] || ''}
                        onChange={(value) =>
                          setMappingsOverride((prev) => ({
                            ...prev,
                            [key]: { ...prev[key], [colKey]: value }
                          }))
                        }
                        placeholder={`Select ${colKey.replace(/_/g, ' ')}`}
                        className="w-full p-2 border"
                      />
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>

          {/* Loading Indicator */}
          {loadingFilters && (
            <div className="flex justify-center items-center mt-6 mb-6">
              <div className="animate-spin rounded-full h-8 w-8 border-t-4 border-blue-500 border-opacity-75"></div>
              <span className="ml-3 text-blue-700 font-medium">Loading filters...</span>
            </div>
          )}
    
          {/* Month Range Selection - Keep as regular select for CY */}
          {availableMonths.ly.length > 0 && availableMonths.cy.length > 0 && (
            <div className="mt-10">
              <h3 className="text-blue-700 font-semibold text-lg mb-2">Select Month Range</h3>

              <div className="grid grid-cols-2 gap-6">
                
                {/* Last Year Months (with checkboxes) */}
                <div>
                  <label className="block mb-1 font-medium">
                    {getFileTypeLabel('ly', 'month')}
                  </label>

                  {/* Select All */}
                  <div className="flex items-center mb-2">
                    <input
                      type="checkbox"
                      className="mr-2"
                      checked={selectedMonths.ly.length === availableMonths.ly.length}
                      onChange={(e) => {
                        setSelectedMonths((prev) => ({
                          ...prev,
                          ly: e.target.checked ? [...availableMonths.ly] : [],
                        }));
                      }}
                    />
                    <span>Select All</span>
                  </div>

                  {/* Month Checkboxes */}
                  <div className="grid grid-cols-2 gap-2">
                    {availableMonths.ly.map((month) => (
                      <label key={month} className="flex items-center">
                        <input
                          type="checkbox"
                          className="mr-2"
                          value={month}
                          checked={selectedMonths.ly.includes(month)}
                          onChange={(e) => {
                            const checked = e.target.checked;
                            setSelectedMonths((prev) => {
                              const updated = checked
                                ? [...prev.ly, month]
                                : prev.ly.filter((m) => m !== month);
                              return { ...prev, ly: updated };
                            });
                          }}
                        />
                        {month}
                      </label>
                    ))}
                  </div>
                </div>

                {/* Current Month (single-select) - Keep as regular select */}
                <div>
                  <label className="block mb-1 font-medium">
                    {getFileTypeLabel('cy', 'month')}
                  </label>
                  <select
                    className="w-full border p-2"
                    value={selectedMonths.cy[0] || ''}
                    onChange={(e) =>
                      setSelectedMonths((prev) => ({
                        ...prev,
                        cy: [e.target.value],
                      }))
                    }
                  >
                    {availableMonths.cy.map((m) => (
                      <option key={m} value={m}>{m}</option>
                    ))}
                  </select>
                </div>
              </div>
            </div>
          )}

          {/* Filters Section */}
          {executives.length > 0 && (
            <div className="mt-10">
              <h3 className="text-blue-700 font-semibold text-lg mb-2">Filter Options</h3>

              <div className="grid grid-cols-2 gap-6">
                {/* Executives */}
                <div>
                  <label className="block mb-1 font-medium">Executives</label>
                  <div className="flex items-center mb-2">
                    <input
                      type="checkbox"
                      className="mr-2"
                      checked={selectedExecutives.length === executives.length}
                      onChange={(e) =>
                        setSelectedExecutives(e.target.checked ? [...executives] : [])
                      }
                    />
                    <span>Select All</span>
                  </div>
                  <div className="max-h-40 overflow-y-auto border p-2 rounded">
                    {executives.map((exec) => (
                      <label key={exec} className="flex items-center mb-1">
                        <input
                          type="checkbox"
                          className="mr-2"
                          value={exec}
                          checked={selectedExecutives.includes(exec)}
                          onChange={(e) => {
                            const checked = e.target.checked;
                            setSelectedExecutives((prev) => {
                              if (checked) return [...prev, exec];
                              return prev.filter((x) => x !== exec);
                            });
                          }}
                        />
                        {exec}
                      </label>
                    ))}
                  </div>
                </div>

                {/* Company Groups */}
                <div>
                  <label className="block mb-1 font-medium">Company Groups</label>
                  <div className="flex items-center mb-2">
                    <input
                      type="checkbox"
                      className="mr-2"
                      checked={selectedGroups.length === companyGroups.length}
                      onChange={(e) =>
                        setSelectedGroups(e.target.checked ? [...companyGroups] : [])
                      }
                    />
                    <span>Select All</span>
                  </div>
                  <div className="max-h-40 overflow-y-auto border p-2 rounded">
                    {companyGroups.map((grp) => (
                      <label key={grp} className="flex items-center mb-1">
                        <input
                          type="checkbox"
                          className="mr-2"
                          value={grp}
                          checked={selectedGroups.includes(grp)}
                          onChange={(e) => {
                            const checked = e.target.checked;
                            setSelectedGroups((prev) => {
                              if (checked) return [...prev, grp];
                              return prev.filter((x) => x !== grp);
                            });
                          }}
                        />
                        {grp}
                      </label>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Generate Button */}
          <div className="mt-8">
            <button 
              onClick={handleGenerateReport} 
              className="bg-red-600 text-white px-4 py-2 rounded disabled:bg-gray-400" 
              disabled={loadingReport}
            >
              {loadingReport ? 'Generating...' : 'Generate Product Growth Report'}
            </button>
          </div>

          {/* Results Table */}
          {groupResults?.results &&
            Object.entries(groupResults.results).map(([group, data]) => (
              <div key={group} className="mb-10 border rounded p-4 shadow">
                {/* Quantity Table */}
                <h4 className="text-lg font-semibold mb-2 text-blue-700">{group} - Quantity Growth</h4>
                <div className="overflow-auto">
                  {Array.isArray(data.qty_df) && data.qty_df.length > 0 ? (
                    <table className="table-auto w-full border text-sm mb-6">
                      <thead>
                        <tr>
                          {["PRODUCT NAME", "LAST YEAR QTY/MT", "BUDGET QTY/MT", "CURRENT YEAR QTY/MT", "GROWTH %"].map((col) => (
                            <th key={col} className="border px-2 py-1 bg-blue-100">{col}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {data.qty_df.map((row, idx) => (
                          <tr key={idx} className={idx === data.qty_df.length - 1 && row['PRODUCT NAME'] === 'TOTAL' ? 'bg-gray-200 font-bold' : ''}>
                            <td className="border px-2 py-1">{row['PRODUCT NAME']}</td>
                            <td className="border px-2 py-1">{typeof row['LAST YEAR QTY/MT'] === 'number' ? row['LAST YEAR QTY/MT'].toFixed(2) : row['LAST YEAR QTY/MT'] || '0.00'}</td>
                            <td className="border px-2 py-1">{typeof row['BUDGET QTY/MT'] === 'number' ? row['BUDGET QTY/MT'].toFixed(2) : row['BUDGET QTY/MT'] || '0.00'}</td>
                            <td className="border px-2 py-1">{typeof row['CURRENT YEAR QTY/MT'] === 'number' ? row['CURRENT YEAR QTY/MT'].toFixed(2) : row['CURRENT YEAR QTY/MT'] || '0.00'}</td>
                            <td className="border px-2 py-1">{typeof row['GROWTH %'] === 'number' ? row['GROWTH %'].toFixed(2) + '%' : (row['GROWTH %'] || '0.00') + '%'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <p className="text-sm text-gray-500">No quantity data available for {group}.</p>
                  )}
                </div>

                {/* Value Table */}
                <h4 className="text-lg font-semibold mb-2 text-green-700">{group} - Value Growth</h4>
                <div className="overflow-auto">
                  {Array.isArray(data.value_df) && data.value_df.length > 0 ? (
                    <table className="table-auto w-full border text-sm">
                      <thead>
                        <tr>
                          {["PRODUCT NAME", "LAST YEAR VALUE/L", "BUDGET VALUE/L", "CURRENT YEAR VALUE/L", "GROWTH %"].map((col) => (
                            <th key={col} className="border px-2 py-1 bg-green-100">{col}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {data.value_df.map((row, idx) => (
                          <tr key={idx} className={idx === data.value_df.length - 1 && row['PRODUCT NAME'] === 'TOTAL' ? 'bg-gray-200 font-bold' : ''}>
                            <td className="border px-2 py-1">{row['PRODUCT NAME']}</td>
                            <td className="border px-2 py-1">{typeof row['LAST YEAR VALUE/L'] === 'number' ? row['LAST YEAR VALUE/L'].toFixed(2) : row['LAST YEAR VALUE/L'] || '0.00'}</td>
                            <td className="border px-2 py-1">{typeof row['BUDGET VALUE/L'] === 'number' ? row['BUDGET VALUE/L'].toFixed(2) : row['BUDGET VALUE/L'] || '0.00'}</td>
                            <td className="border px-2 py-1">{typeof row['CURRENT YEAR VALUE/L'] === 'number' ? row['CURRENT YEAR VALUE/L'].toFixed(2) : row['CURRENT YEAR VALUE/L'] || '0.00'}</td>
                            <td className="border px-2 py-1">{typeof row['GROWTH %'] === 'number' ? row['GROWTH %'].toFixed(2) + '%' : (row['GROWTH %'] || '0.00') + '%'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <p className="text-sm text-gray-500">No value data available for {group}.</p>
                  )}
                </div>
              </div>
            ))}

          {/* Download PPT Button */}
          {groupResults?.results && Object.keys(groupResults.results).length > 0 && (
            <button
              onClick={handleDownloadPPT}
              className="mt-6 bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-4 rounded"
            >
              Download Product Growth PPT
            </button>
          )}
        </>
      )}
    </div>
  );
};

export default ProductGrowth;
