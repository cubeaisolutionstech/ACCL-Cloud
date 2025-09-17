import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';
import SearchableSelect from './SearchableSelect'; // Import SearchableSelect
import OdTargetSubTab from './OdTargetSubTab';
import { addReportToStorage } from '../utils/consolidatedStorage';

const NumberOfBilledCustomers = () => {
  const { selectedFiles } = useExcelData();
  const [activeTab, setActiveTab] = useState('nbc');

  const [salesSheets, setSalesSheets] = useState([]);
  const [salesSheet, setSalesSheet] = useState('');
  const [salesHeader, setSalesHeader] = useState(1);

  const [nbcColumns, setNbcColumns] = useState({});
  const [allSalesCols, setAllSalesCols] = useState([]);
  const [filters, setFilters] = useState({ branches: [], executives: [] });
  const [selectedBranches, setSelectedBranches] = useState([]);
  const [selectedExecutives, setSelectedExecutives] = useState([]);
  const [selectAllBranches, setSelectAllBranches] = useState(true);
  const [selectAllExecutives, setSelectAllExecutives] = useState(true);

  const [nbcResults, setNbcResults] = useState({});
  const [loading, setLoading] = useState(false);
  const [loadingAutoMap, setLoadingAutoMap] = useState(false);
  const [loadingReport, setLoadingReport] = useState(false);
  const [downloading, setDownloading] = useState(false); // Fixed: Moved to component level

  const fetchSheetNames = async () => {
    try {
      const res = await axios.post('http://localhost:5000/api/branch/sheets', {
        filename: selectedFiles.salesFile
      });
      setSalesSheets(res.data.sheets || []);
      if (res.data.sheets?.length > 0) {
        setSalesSheet(res.data.sheets[0]);
      }
    } catch (err) {
      console.error("Error fetching sheet names", err);
    }
  };

  const fetchAutoMap = async () => {
    setLoadingAutoMap(true);
    try {
      const res = await axios.post('http://localhost:5000/api/branch/get_columns', {
        filename: selectedFiles.salesFile,
        sheet_name: salesSheet,
        header: salesHeader
      });
      setAllSalesCols(res.data.columns || []);
      console.log("Fetched columns:", res.data.columns);

      const mapRes = await axios.post('http://localhost:5000/api/branch/get_nbc_columns', {
        filename: selectedFiles.salesFile,
        sheet_name: salesSheet,
        header: salesHeader
      });
      setNbcColumns(mapRes.data.mapping || {});
      console.log("Column mapping:", mapRes.data.mapping);

      // Auto-fetch filters after column mapping is complete
      if (mapRes.data.mapping) {
        const filterRes = await axios.post('http://localhost:5000/api/branch/get_nbc_filters', {
          filename: selectedFiles.salesFile,
          sheet_name: salesSheet,
          header: salesHeader,
          date_col: mapRes.data.mapping.date,
          branch_col: mapRes.data.mapping.branch,
          executive_col: mapRes.data.mapping.executive
        });
        const { branches, executives } = filterRes.data;
        setFilters({ branches, executives });
        setSelectedBranches(branches);
        setSelectedExecutives(executives);
      }
    } catch (err) {
      console.error("Error in auto mapping NBC columns", err);
      alert("Failed to load columns and mappings. Please check your file and try again.");
    } finally {
      setLoadingAutoMap(false);
    }
  };

  const addBranchNBCReportsToStorage = (resultsData) => {
    try {
      const customerReports = [];

      // Loop over each FY in results
      Object.entries(resultsData).forEach(([fyKey, fyData]) => {
        // Extract the most recent month for the title
        const mostRecentMonth = fyData.recent_month || 'MONTH';
        
        customerReports.push({
          df: fyData.data || [],
          title: `NUMBER OF BILLED CUSTOMERS - ${mostRecentMonth}`,
          percent_cols: [], // No percentage columns
          columns: ["Branch", ...(fyData.months || [])] // Ensure Branch is first column
        });
      });

      if (customerReports.length > 0) {
        addReportToStorage(customerReports, 'branch_nbc_results');
        console.log(`Stored ${customerReports.length} Branch NBC reports to consolidated storage`);
      }
    } catch (error) {
      console.error("Error storing Branch NBC reports:", error);
    }
  };

  const handleGenerateReport = async () => {
    setLoadingReport(true);
    try {
      // Validate required columns are selected
      const requiredColumns = ['date', 'branch', 'customer_id', 'executive'];
      const missingColumns = requiredColumns.filter(col => !nbcColumns[col]);
      
      if (missingColumns.length > 0) {
        alert(`Please select columns for: ${missingColumns.join(', ')}`);
        return;
      }

      const payload = {
        filename: selectedFiles.salesFile,
        sheet_name: salesSheet,
        header: salesHeader,
        date_col: nbcColumns.date,
        branch_col: nbcColumns.branch,
        customer_id_col: nbcColumns.customer_id,
        executive_col: nbcColumns.executive,
        selected_branches: selectedBranches,
        selected_executives: selectedExecutives
      };

      console.log("Generating NBC report with payload:", payload);

      const res = await axios.post('http://localhost:5000/api/branch/calculate_nbc_table', payload);
      
      if (res.data && res.data.results) {
        console.log("NBC Results received:", res.data.results);
        setNbcResults(res.data.results);
        addBranchNBCReportsToStorage(res.data.results);
      } else {
        console.warn("No NBC results received from server");
        alert("No data was generated. Please check your filters and try again.");
      }
    } catch (err) {
      console.error("Error generating NBC report:", err);
      alert("Failed to generate NBC report. Please check the console for details.");
    } finally {
      setLoadingReport(false);
    }
  };

  const handleDownloadPPT = async (fyKey) => {
    // Prevent multiple downloads
    if (downloading) {
      console.log("Download already in progress...");
      return;
    }

    try {
      setDownloading(true);
      
      const fyData = nbcResults[fyKey];
      if (!fyData || !fyData.data || fyData.data.length === 0) {
        alert("No data available for download");
        return;
      }

      const mostRecentMonth = fyData.recent_month || 'MONTH';
      
      const payload = {
        data: fyData.data,
        title: `NUMBER OF BILLED CUSTOMERS - ${mostRecentMonth}`,
        months: fyData.months || [],
        financial_year: fyKey,
        logo_file: selectedFiles.logoFile
      };

      console.log("Downloading PPT with payload:", {
        dataRows: payload.data.length,
        title: payload.title,
        months: payload.months.length,
        fy: payload.financial_year,
        hasLogo: !!payload.logo_file
      });

      // Use fetch instead of axios to avoid any axios CORS handling
      const response = await fetch("http://localhost:5000/api/branch/download_nbc_ppt", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      // Check if response is ok
      if (!response.ok) {
        const errorText = await response.text();
        let errorMessage;
        try {
          const errorJson = JSON.parse(errorText);
          errorMessage = errorJson.error || `Server error: ${response.status}`;
        } catch {
          errorMessage = `Server error: ${response.status}`;
        }
        throw new Error(errorMessage);
      }

      // Check content type
      const contentType = response.headers.get("content-type");
      if (!contentType || !contentType.includes("application/vnd.openxmlformats-officedocument.presentationml.presentation")) {
        const text = await response.text();
        console.error("Unexpected response:", text);
        throw new Error("Server returned invalid file format");
      }

      // Get the blob
      const blob = await response.blob();
      
      if (blob.size === 0) {
        throw new Error("Downloaded file is empty");
      }

      console.log(`PPT blob size: ${blob.size} bytes`);

      // Create download link
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `NBC_Report_${mostRecentMonth.replace(/\s+/g, '_')}.pptx`;
      
      // Add to DOM, click, and remove
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      
      // Clean up
      window.URL.revokeObjectURL(url);
      
      console.log("NBC PPT downloaded successfully");
      
    } catch (err) {
      console.error("Failed to download NBC PPT:", err);
      
      let errorMessage = "NBC PPT download failed. ";
      
      if (err.message.includes("Failed to fetch")) {
        errorMessage += "Cannot connect to server. Please ensure the server is running.";
      } else if (err.message.includes("empty")) {
        errorMessage += "The generated file is empty. Please check your data.";
      } else if (err.message.includes("CORS")) {
        errorMessage += "Connection issue with the server.";
      } else {
        errorMessage += err.message || "Please try again.";
      }
      
      alert(errorMessage);
      
    } finally {
      setDownloading(false);
    }
  };

  useEffect(() => {
    if (selectedFiles.salesFile) {
      fetchSheetNames();
    }
  }, [selectedFiles.salesFile]);

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold text-blue-900 mb-4">Number of Billed Customers & OD Target</h2>

      <div className="flex space-x-4 mb-4">
        <button
          onClick={() => setActiveTab('nbc')}
          className={`px-4 py-2 rounded ${activeTab === 'nbc' ? 'bg-blue-700 text-white' : 'bg-gray-200 text-black'}`}
        >
          Number of Billed Customers
        </button>
        <button
          onClick={() => setActiveTab('od')}
          className={`px-4 py-2 rounded ${activeTab === 'od' ? 'bg-blue-700 text-white' : 'bg-gray-200 text-black'}`}
        >
          OD Target
        </button>
      </div>

      {activeTab === 'nbc' && (
        <div>
          {/* Sheet & Header Selection */}
          <div className="grid grid-cols-2 gap-4 mb-4">
            <div>
              <label className="block font-medium">Sales Sheet</label>
              <select 
                className="w-full p-2 border rounded" 
                value={salesSheet} 
                onChange={(e) => setSalesSheet(e.target.value)}
              >
                <option value="">Select Sheet</option>
                {salesSheets.map(sheet => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="block font-medium">Header Row</label>
              <input
                type="number"
                min={1}
                value={salesHeader}
                onChange={(e) => setSalesHeader(Number(e.target.value))}
                className="w-full p-2 border rounded"
              />
            </div>
          </div>

          <button 
            onClick={fetchAutoMap} 
            className="bg-blue-600 text-white px-4 py-2 rounded disabled:bg-gray-400" 
            disabled={loadingAutoMap || !salesSheet}
          >
            {loadingAutoMap ? 'Loading...' : 'Load Columns & Auto Map'}
          </button>

          {Object.keys(nbcColumns).length > 0 && (
            <div className="mt-6">
              <h3 className="font-bold text-blue-800 mb-2">Column Mapping</h3>
              <div className="grid grid-cols-2 gap-4">
                {['date', 'branch', 'customer_id', 'executive'].map((key) => (
                  <div key={key}>
                    <label className="block font-semibold mb-1 capitalize">
                      {key.replace(/_/g, ' ')} <span className="text-red-500">*</span>
                    </label>
                    <SearchableSelect
                      options={allSalesCols}
                      value={nbcColumns[key] || ''}
                      onChange={(value) => setNbcColumns(prev => ({ ...prev, [key]: value }))}
                      placeholder={`Select ${key.replace(/_/g, ' ')}`}
                      className="w-full p-2 border rounded"
                    />
                  </div>
                ))}
              </div>
            </div>
          )}

          {Object.keys(nbcColumns).length > 0 && (
            <>
              <div className="grid grid-cols-2 gap-6 mt-4">
                {/* Branch Filter */}
                <div>
                  <label className="block font-medium mb-2">Branches</label>
                  <div className="mb-2">
                    <input
                      type="checkbox"
                      id="selectAllBranches"
                      checked={selectAllBranches}
                      onChange={() => {
                        const newVal = !selectAllBranches;
                        setSelectAllBranches(newVal);
                        setSelectedBranches(newVal ? filters.branches : []);
                      }}
                    />
                    <label htmlFor="selectAllBranches" className="ml-2">Select All</label>
                  </div>
                  <div className="max-h-40 overflow-y-auto border p-2 rounded bg-gray-50">
                    {filters.branches.map(branch => (
                      <label key={branch} className="flex items-center mb-1">
                        <input
                          type="checkbox"
                          className="mr-2"
                          checked={selectedBranches.includes(branch)}
                          onChange={(e) => {
                            const checked = e.target.checked;
                            setSelectedBranches((prev) => {
                              const updated = checked 
                                ? [...prev, branch]
                                : prev.filter((x) => x !== branch);
                              setSelectAllBranches(updated.length === filters.branches.length);
                              return updated;
                            });
                          }}
                        />
                        <span className="text-sm">{branch}</span>
                      </label>
                    ))}
                  </div>
                </div>

                {/* Executive Filter */}
                <div>
                  <label className="block font-medium mb-2">Executives</label>
                  <div className="mb-2">
                    <input
                      type="checkbox"
                      id="selectAllExecutives"
                      checked={selectAllExecutives}
                      onChange={() => {
                        const newVal = !selectAllExecutives;
                        setSelectAllExecutives(newVal);
                        setSelectedExecutives(newVal ? filters.executives : []);
                      }}
                    />
                    <label htmlFor="selectAllExecutives" className="ml-2">Select All</label>
                  </div>
                  <div className="max-h-40 overflow-y-auto border p-2 rounded bg-gray-50">
                    {filters.executives.map(exec => (
                      <label key={exec} className="flex items-center mb-1">
                        <input
                          type="checkbox"
                          className="mr-2"
                          checked={selectedExecutives.includes(exec)}
                          onChange={(e) => {
                            const checked = e.target.checked;
                            setSelectedExecutives((prev) => {
                              const updated = checked 
                                ? [...prev, exec]
                                : prev.filter((x) => x !== exec);
                              setSelectAllExecutives(updated.length === filters.executives.length);
                              return updated;
                            });
                          }}
                        />
                        <span className="text-sm">{exec}</span>
                      </label>
                    ))}
                  </div>
                </div>
              </div>

              <button 
                onClick={handleGenerateReport} 
                className="mt-6 bg-red-600 text-white px-6 py-2 rounded disabled:bg-gray-400" 
                disabled={loadingReport}
              >
                {loadingReport ? 'Generating...' : 'Generate Billed Customers Report'}
              </button>

              {Object.entries(nbcResults).map(([fy, result]) => {
                const mostRecentMonth = result.recent_month || 'MONTH';
                return (
                  <div key={fy} className="mb-8 mt-6">
                    <h3 className="text-lg font-bold text-blue-700 mb-2">
                      NUMBER OF BILLED CUSTOMERS - {mostRecentMonth}
                    </h3>
                    <div className="overflow-auto border rounded shadow-sm">
                      <table className="table-auto w-full border-collapse text-sm">
                        <thead>
                          <tr className="bg-blue-600 text-white">
                            <th className="border px-3 py-2 text-left">Branch</th>
                            {result.months.map((month) => (
                              <th key={month} className="border px-3 py-2 text-center">{month}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {result.data.map((row, idx) => {
                            const isTotalRow = row["Branch"] === "GRAND TOTAL";
                            return (
                              <tr 
                                key={idx} 
                                className={isTotalRow ? "bg-gray-300 font-bold" : idx % 2 === 0 ? "bg-gray-50" : "bg-white"}
                              >
                                <td className="border px-3 py-2 font-medium">
                                  {row["Branch"] || "N/A"}
                                </td>
                                {result.months.map((month) => (
                                  <td key={month} className="border px-3 py-2 text-right">
                                    {row[month] ?? 0}
                                  </td>
                                ))}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                    
                    <div className="mt-3">
                      <button
                        className={`px-4 py-2 rounded text-white ${downloading ? 'bg-gray-400 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700'}`}
                        onClick={() => handleDownloadPPT(fy)}
                        disabled={downloading}
                      >
                        {downloading ? 'Downloading...' : `Download NBC Report - ${mostRecentMonth}`}
                      </button>
                      <p className="text-sm text-gray-600 mt-1">
                        Total Records: {result.data.length - 1} branches + Grand Total
                      </p>
                    </div>
                  </div>
                );
              })}
            </>
          )}
        </div>
      )}

      {activeTab === 'od' && (
        <div className="mt-10">
          <OdTargetSubTab />
        </div>
      )}
    </div>
  );
};

export default NumberOfBilledCustomers;
