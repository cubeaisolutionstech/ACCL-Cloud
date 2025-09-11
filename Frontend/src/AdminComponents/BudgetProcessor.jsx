import React, { useState } from "react";
import api from "../api/axios";

const BudgetProcessor = () => {
  const [file, setFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headerRow, setHeaderRow] = useState(1);
  const [downloadLink, setDownloadLink] = useState(null);
  const [processedExcelBase64, setProcessedExcelBase64] = useState(null);
  const [columns, setColumns] = useState([]);
  const [preview, setPreview] = useState([]);
  const [customFilename, setCustomFilename] = useState("");
  const [loading, setLoading] = useState(false); 
  
  const [colMap, setColMap] = useState({
    customer_col: "",
    exec_code_col: "",
    exec_name_col: "",
    branch_col: "",
    region_col: "",
    cust_name_col: "",
  });

  const [metrics, setMetrics] = useState(null);

  const handleFileUpload = async (e) => {
    const f = e.target.files[0];
    if (!f) return;

    setFile(f);
    const data = new FormData();
    data.append("file", f);

    const res = await api.post("/upload-tools/sheet-names", data);
    setSheetNames(res.data.sheet_names);

    const consolidateSheet = res.data.sheet_names.find(s => s.toLowerCase().trim() === "consolidate");
    setSelectedSheet(consolidateSheet || res.data.sheet_names[0]);
  };

  const handlePreview = async () => {
    setLoading(true);
    try {
      const data = new FormData();
      data.append("file", file);
      data.append("sheet_name", selectedSheet);
      data.append("header_row", headerRow);

      const res = await api.post("/upload-tools/preview", data);
      setColumns(res.data.columns);
      setPreview(res.data.preview);

      // Auto map
      const auto = (key) => res.data.columns.find(c => c.toLowerCase().includes(key)) || "";
      setColMap({
        customer_col: auto("sl code"),
        exec_code_col: auto("executive code") || auto("code"),
        exec_name_col: auto("executive name"),
        branch_col: auto("branch"),
        region_col: auto("region"),
        cust_name_col: auto("party name")
      });
      } catch (err) {
      console.error("Preview error", err);
      alert("Failed to load preview");
    } finally {
      setLoading(false); 
    }
  };

  const handleProcess = async () => {
    setLoading(true);
    try {
      const data = new FormData();
      data.append("file", file);
      data.append("sheet_name", selectedSheet);
      data.append("header_row", headerRow);

      Object.entries(colMap).forEach(([key, val]) => {
        if (val) data.append(key, val);
      });

      const res = await api.post("/upload-budget-file", data);
      setPreview(res.data.preview);
      setMetrics(res.data.counts);

      const byteCharacters = atob(res.data.file_data);
      const byteNumbers = new Array(byteCharacters.length).fill().map((_, i) => byteCharacters.charCodeAt(i));
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      setDownloadLink(url);
      setProcessedExcelBase64(res.data.file_data);
    }catch(err) {
      console.error("Process error", err);
      alert("Failed to process file");
    }
    finally {
      setLoading(false);
      }
  };

  const handleSave = async () => {
    if (!processedExcelBase64) return alert("No processed file to save");

    setLoading(true); // ✅ Start
    try {
      const res = await api.post("/save-budget-file", {
        file_data: processedExcelBase64,
        filename: customFilename?.trim() || "Processed_Budget.xlsx"
      });
      alert("✅ File saved to DB. File ID: " + res.data.id);
    } catch (err) {
      console.error("Save error", err);
      alert("Saving failed");
    } finally {
      setLoading(false); // ✅ End
    }
  };

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold mb-4">Budget Processing</h2>

      <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="mb-4" />

      {sheetNames.length > 0 && (
        <>
          <label className="block mb-2">Header Row</label>
          <input
            type="number"
            value={headerRow}
            onChange={(e) => setHeaderRow(parseInt(e.target.value))}
            className="border p-2 mb-4 w-32"
          />

          <button onClick={handlePreview} className="bg-blue-600 text-white px-4 py-2 rounded mb-4">
            Load Sheet
          </button>
        </>
      )}

      {loading && (
        <div className="text-blue-600 font-medium my-2 flex items-center gap-2">
          <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"></path>
          </svg>
          Loading...
        </div>
      )}

      {columns.length > 0 && (
        <div className="grid md:grid-cols-2 gap-4 mb-4">
          {[
            ["Customer Code Column", "customer_col"],
            ["Executive Code Column", "exec_code_col"],
            ["Executive Name Column", "exec_name_col"],
            ["Branch Column", "branch_col"],
            ["Region Column", "region_col"],
            ["Customer Name Column", "cust_name_col"],
          ].map(([label, key]) => (
            <div key={key}>
              <label>{label}</label>
              <select
                value={colMap[key]}
                onChange={(e) => setColMap((prev) => ({ ...prev, [key]: e.target.value }))}
                className="w-full border p-2"
              >
                <option value="">-- Select --</option>
                {columns.map((col, i) => (
                  <option key={i} value={col}>{col}</option>
                ))}
              </select>
            </div>
          ))}
        </div>
      )}

      {columns.length > 0 && (
        <button onClick={handleProcess} className="bg-green-600 text-white px-4 py-2 rounded">
          Process Budget File
        </button>
      )}

      {metrics && (
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 my-6">
          {Object.entries(metrics).map(([key, val]) => (
            <div key={key} className="p-3 bg-gray-100 rounded shadow">
              <p className="text-sm">{key.replace("_", " ")}</p>
              <p className="text-xl font-bold">{val}</p>
            </div>
          ))}
        </div>
      )}
      {downloadLink && (
  <div className="mt-4">
    <a
      href={downloadLink}
      download="processed_budget.xlsx"
      className="bg-blue-600 text-white px-4 py-2 rounded inline-block"
    >
      Download Processed Budget File
    </a>
  </div>
)}
<div className="mb-4">
  <label className="block text-sm font-semibold text-gray-700">Enter Filename to Save</label>
  <input
    type="text"
    value={customFilename}
    onChange={(e) => setCustomFilename(e.target.value)}
    className="mt-1 p-2 border border-gray-300 rounded w-full"
    placeholder="e.g., Apr-2025_Processed_Budget"
  />
</div>

{processedExcelBase64 && (
  <button
    className="bg-indigo-600 text-white px-4 py-2 rounded mt-4"
    onClick={handleSave}
  >
    Save Processed Excel to DB
  </button>
)}
    </div>
  );
};

export default BudgetProcessor;
