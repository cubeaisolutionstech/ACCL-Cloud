import React, { useState, useEffect } from "react";
import axios from "axios";
import { useExcelData } from "../context/ExcelDataContext";
import SearchableSelect from './SearchableSelect'; // Import SearchableSelect
import { addReportToStorage } from '../utils/consolidatedStorage';


const OdTargetSubTab = () => {
  const { selectedFiles, logoFile } = useExcelData();

  const [odActiveFile, setOdActiveFile] = useState("OS-Current Month");
  const [sheetOptions, setSheetOptions] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headerRow, setHeaderRow] = useState(1);

  const [columns, setColumns] = useState([]);
  const [mapping, setMapping] = useState({});
  const [override, setOverride] = useState({});

  const [years, setYears] = useState([]);
  const [months] = useState([
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ]);
  const [selectedYears, setSelectedYears] = useState([]);
  const [tillMonth, setTillMonth] = useState("March");

  const [branches, setBranches] = useState([]);
  const [executives, setExecutives] = useState([]);
  const [selectedBranches, setSelectedBranches] = useState([]);
  const [selectedExecutives, setSelectedExecutives] = useState([]);

  const [resultTable, setResultTable] = useState([]);
  const [periodTitle, setPeriodTitle] = useState("");
  const [loading, setLoading] = useState(false);
  const [loadingAutoMap, setLoadingAutoMap] = useState(false); // Loader for Load Columns & Auto Map
  const [loadingReport, setLoadingReport] = useState(false); // Loader for Generate OD Target Report

  const activeFilename =
    odActiveFile === "OS-Current Month"
      ? selectedFiles.osCurrFile
      : selectedFiles.osPrevFile;

  useEffect(() => {
    if (activeFilename) {
      fetchSheets();
    }
  }, [odActiveFile, selectedFiles]);

  const fetchSheets = async () => {
    const res = await axios.post("http://localhost:5000/api/branch/sheets", {
      filename: activeFilename
    });
    setSheetOptions(res.data.sheets || []);
    if (res.data.sheets.length > 0) {
      setSelectedSheet(res.data.sheets[0]);
    }
  };

  const handleAutoMap = async () => {
    setLoadingAutoMap(true);
    try {
      const colRes = await axios.post("http://localhost:5000/api/branch/get_columns", {
        filename: activeFilename,
        sheet_name: selectedSheet,
        header: headerRow
      });
      setColumns(colRes.data.columns || []);
      console.log("get column",colRes.data.col)

      const mapRes = await axios.post("http://localhost:5000/api/branch/get_od_target_columns", {
        columns: colRes.data.columns || []
      });

      const mapped = mapRes.data.mapping || {};
      setMapping(mapped);
      setOverride(mapped);

      // Fetch unique values for filters
      const filterRes = await axios.post("http://localhost:5000/api/branch/get_column_unique_values", {
        filename: activeFilename,          
        sheet_name: selectedSheet,
        header: headerRow,
        column_names: [mapped.due_date, mapped.area, mapped.executive]
      });

      setYears(filterRes.data[mapped.due_date]?.years || []);
      setSelectedYears(filterRes.data[mapped.due_date]?.years || []);

      setBranches(filterRes.data[mapped.area]?.values || []);
      setSelectedBranches(filterRes.data[mapped.area]?.values || []);

      setExecutives(filterRes.data[mapped.executive]?.values || []);
      setSelectedExecutives(filterRes.data[mapped.executive]?.values || []);
    } catch (err) {
      console.error("Auto map failed", err);
    } finally {
      setLoadingAutoMap(false);
    }
  };  

  const addBranchODTargetReportsToStorage = (resultsData, endDate, osLabel) => {
    try {
      const odTargetReports = [{
        df: resultsData || [],
        title: `OD Target- ${endDate || 'All Periods'}`,
        percent_cols: [] // no % columns
      }];

      if (odTargetReports.length > 0) {
        const storageKey = osLabel === ""
          ? 'branch_od_results_current'
          : 'branch_od_results_previous';

        addReportToStorage(odTargetReports, storageKey);
        console.log(`✅ Stored ${osLabel} OD Target to "${storageKey}"`);

        // Dispatch custom event to notify ConsolidatedReportPanel
        const event = new Event('reportAdded');
        window.dispatchEvent(event);
      }
    } catch (error) {
      console.error("❌ Error storing Branch OD Target report:", error);
    }
  };

  const handleGenerateReport = async () => {
    setLoadingReport(true);
    try {
      setResultTable([]);
      setPeriodTitle("");

      const osLabel = odActiveFile;
      const activeFilename =
        odActiveFile === "OS-Previous Month"
          ? selectedFiles.osPrevFile
          : selectedFiles.osCurrFile;

      const payload = {
        filename: activeFilename,
        sheet_name: selectedSheet,
        header: headerRow,
        area_col: override.area,
        due_date_col: override.due_date,
        qty_col: override.net_value,
        executive_col: override.executive,
        selected_branches: selectedBranches,
        selected_executives: selectedExecutives,
        selected_years: selectedYears,
        till_month: tillMonth
      };

      const res = await axios.post(
        "http://localhost:5000/api/branch/calculate_od_target_table",
        payload
      );

      const table = res.data.table || [];
      const end = res.data.end || "";

      setResultTable(table);
      // FIXED: Clean title without "Value in Lakhs" - matching Streamlit exactly
      setPeriodTitle(`OD Target - ${end}`);

      // ✅ Add to consolidated storage
      addBranchODTargetReportsToStorage(table, end, osLabel);
    } catch (err) {
      console.error("❌ Error generating OD Target report:", err);
      alert("OD Target generation failed. Check console.");
    } finally {
      setLoadingReport(false);
    }
  };

  const handleDownloadPPT = async () => {
    try {
      const payload = {
        result: resultTable,
        title: periodTitle,
        logo_file: logoFile
      };
      console.log("⬇️ Sending to PPT API:", payload);

      const res = await axios.post(
        "http://localhost:5000/api/branch/download_od_target_ppt",
        payload,
        { responseType: "blob" }
      );

      const blob = new Blob([res.data], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      });

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${periodTitle.replace(/\s/g, "_")}_${Date.now()}.pptx`;
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error("PPT download failed", err);
    }
  };

   return (
    <div className="p-4">
      <h2 className="text-lg font-bold text-red-700 mb-4">📋 OD Target Report</h2>

      <div className="mb-4">
        <label className="font-semibold mr-4">Choose File:</label>
        <label className="mr-4">
          <input
            type="radio"
            value="OS-Previous Month"
            checked={odActiveFile === "OS-Previous Month"}
            onChange={() => setOdActiveFile("OS-Previous Month")}
          />{" "}
          OS - Previous Month
        </label>
        <label>
          <input
            type="radio"
            value="OS-Current Month"
            checked={odActiveFile === "OS-Current Month"}
            onChange={() => setOdActiveFile("OS-Current Month")}
          />{" "}
          OS - Current Month
        </label>
      </div>

      <div className="grid grid-cols-2 gap-4 mb-6">
        <div>
          <label className="block font-medium">Sheet Name</label>
          <select
            value={selectedSheet}
            onChange={(e) => setSelectedSheet(e.target.value)}
            className="w-full border p-2"
          >
            <option value="">Select Sheet</option>
            {sheetOptions.map((sheet) => (
              <option key={sheet} value={sheet}>
                {sheet}
              </option>
            ))}
          </select>
        </div>
        <div>
          <label className="block font-medium">Header Row</label>
          <input
            type="number"
            value={headerRow}
            onChange={(e) => setHeaderRow(Number(e.target.value))}
            className="w-full border p-2"
            min="1"
          />
        </div>
      </div>

      {/* Auto Map Button */}
      <div className="flex justify-start mb-6">
        <button
          onClick={handleAutoMap}
          className="bg-blue-600 text-white px-4 py-2 rounded flex items-center justify-center min-w-[220px]"
          disabled={loadingAutoMap}
        >
          {loadingAutoMap ? (
            <>
              <svg className="animate-spin h-5 w-5 mr-2 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"></path>
              </svg>
              Loading...
            </>
          ) : (
            'Load Columns & Auto Map'
          )}
        </button>
      </div>

      {columns.length > 0 && (
        <>
        <div className="grid grid-cols-2 gap-6 mb-6">
          {Object.entries({
            area: "Area Column",
            due_date: "Due Date Column",
            net_value: "Net Value Column",
            executive: "Executive Column"
          }).map(([key, label]) => (
            <div key={key}>
              <label className="block font-medium mb-1">{label}</label>
              <SearchableSelect
                options={columns}
                value={override[key] || ''}
                onChange={(value) => setOverride(prev => ({ ...prev, [key]: value }))}
                placeholder={`Select ${label.toLowerCase()}`}
                className="w-full p-2 border"
              />
            </div>
          ))}
        </div>

      <div className="mt-8">
        <h3 className="text-blue-700 font-semibold text-md mb-2">📆 Date Filters</h3>
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label>Years</label>
            <div className="flex items-center mb-2">
                  <input
                    type="checkbox"
                    className="mr-2"
                    checked={selectedYears.length === years.length}
                    onChange={(e) => setSelectedYears(e.target.checked ? [...years] : [])}
                  />
                  <span>Select All</span>
                </div>
                <div className="max-h-40 overflow-y-auto border p-2 rounded">
                  {years.map((y) => (
                    <label key={y} className="flex items-center mb-1">
                      <input
                        type="checkbox"
                        className="mr-2"
                        value={y}
                        checked={selectedYears.includes(y)}
                        onChange={(e) => {
                          const checked = e.target.checked;
                          setSelectedYears((prev) => {
                            if (checked) return [...prev, y];
                            return prev.filter((x) => x !== y);
                          });
                        }}
                      />
                      {y}
                    </label>
                  ))}
                </div>
          </div>
          <div>
            <label>Until Month</label>
            <select value={tillMonth} onChange={(e) => setTillMonth(e.target.value)} className="w-full border p-2">
              {months.map(m => <option key={m}>{m}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-2 gap-6 mt-6">
        <div>
          <label>Branches</label>
          <div className="flex items-center mb-2">
                <input
                  type="checkbox"
                  className="mr-2"
                  checked={selectedBranches.length === branches.length}
                  onChange={(e) => setSelectedBranches(e.target.checked ? [...branches] : [])}
                />
                <span>Select All</span>
              </div>
              <div className="max-h-40 overflow-y-auto border p-2 rounded">
                {branches.map((b) => (
                  <label key={b} className="flex items-center mb-1">
                    <input
                      type="checkbox"
                      className="mr-2"
                      value={b}
                      checked={selectedBranches.includes(b)}
                      onChange={(e) => {
                        const checked = e.target.checked;
                        setSelectedBranches((prev) => {
                          if (checked) return [...prev, b];
                          return prev.filter((x) => x !== b);
                        });
                      }}
                    />
                    {b}
                  </label>
                ))}
              </div>
        </div>
        <div>
          <label>Executives</label>
          <div className="flex items-center mb-2">
                <input
                  type="checkbox"
                  className="mr-2"
                  checked={selectedExecutives.length === executives.length}
                  onChange={(e) => setSelectedExecutives(e.target.checked ? [...executives] : [])}
                />
                <span>Select All</span>
              </div>
              <div className="max-h-40 overflow-y-auto border p-2 rounded">
                {executives.map((e) => (
                  <label key={e} className="flex items-center mb-1">
                    <input
                      type="checkbox"
                      className="mr-2"
                      value={e}
                      checked={selectedExecutives.includes(e)}
                      onChange={(ev) => {
                        const checked = ev.target.checked;
                        setSelectedExecutives((prev) => {
                          if (checked) return [...prev, e];
                          return prev.filter((x) => x !== e);
                        });
                      }}
                    />
                    {e}
                  </label>
                ))}
              </div>
        </div>
      </div>

      {/* Generate Button - only show after columns are loaded */}
          <button onClick={handleGenerateReport} disabled={loadingReport} className="bg-red-700 text-white mt-6 px-4 py-2 rounded disabled:bg-gray-400">
            {loadingReport ? 'Loading...' : ' Generate OD Target Report'}
          </button>

      {resultTable.length > 0 && (
        <div className="mt-10">
          <h3 className="text-xl font-bold text-green-700 mb-2">{periodTitle}</h3>
          <div className="overflow-x-auto">
            <table className="table-auto w-full border text-sm">
              <thead>
                <tr>
                  {Object.keys(resultTable[0]).map((col) => (
                    <th key={col} className="border px-2 py-1 bg-gray-100">{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {resultTable.map((row, idx) => (
                  <tr key={idx}>
                    {Object.values(row).map((val, j) => (
                      <td key={j} className="border px-2 py-1">{val}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {resultTable.length > 0 && (
            <button
              onClick={handleDownloadPPT}
              className="mt-4 bg-green-700 text-white px-4 py-2 rounded"
            >
               Download OD Target PPT
            </button>
          )}
        </div>
      )}
      </>
      )}
    </div>
  );
};

export default OdTargetSubTab;
