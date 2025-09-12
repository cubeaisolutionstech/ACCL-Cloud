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
  
  // Enhanced loading states
  const [loading, setLoading] = useState(false);
  const [loadingMessage, setLoadingMessage] = useState("");
  const [successMessage, setSuccessMessage] = useState("");
  const [errorMessage, setErrorMessage] = useState("");
  
  const [colMap, setColMap] = useState({
    customer_col: "",
    exec_code_col: "",
    exec_name_col: "",
    branch_col: "",
    region_col: "",
    cust_name_col: "",
  });

  const [metrics, setMetrics] = useState(null);

  // Clear messages helper
  const clearMessages = () => {
    setSuccessMessage("");
    setErrorMessage("");
  };

  const handleFileUpload = async (e) => {
    const f = e.target.files[0];
    if (!f) return;

    clearMessages();
    setLoading(true);
    setLoadingMessage("Uploading file and reading sheet names...");

    try {
      setFile(f);
      const data = new FormData();
      data.append("file", f);

      const res = await api.post("/upload-tools/sheet-names", data);
      setSheetNames(res.data.sheet_names);

      const consolidateSheet = res.data.sheet_names.find(s => s.toLowerCase().trim() === "consolidate");
      setSelectedSheet(consolidateSheet || res.data.sheet_names[0]);
      
      setSuccessMessage(`✅ File "${f.name}" uploaded successfully! Found ${res.data.sheet_names.length} sheet(s).`);
    } catch (err) {
      console.error("Upload error", err);
      setErrorMessage("❌ Failed to upload file. Please try again.");
    } finally {
      setLoading(false);
      setLoadingMessage("");
    }
  };

  const handlePreview = async () => {
    clearMessages();
    setLoading(true);
    setLoadingMessage("Loading sheet preview...");
    
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

      setSuccessMessage(`✅ Sheet "${selectedSheet}" loaded successfully! Found ${res.data.columns.length} columns.`);
    } catch (err) {
      console.error("Preview error", err);
      setErrorMessage("❌ Failed to load preview. Please check your file and try again.");
    } finally {
      setLoading(false);
      setLoadingMessage("");
    }
  };

  const handleProcess = async () => {
    clearMessages();
    setLoading(true);
    setLoadingMessage("Processing budget file... This may take a moment.");
    
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

      setSuccessMessage("🎉 Budget file processed successfully! Your file is ready for download.");
    } catch (err) {
      console.error("Process error", err);
      setErrorMessage("❌ Failed to process file. Please check your column mappings and try again.");
    } finally {
      setLoading(false);
      setLoadingMessage("");
    }
  };

  const handleSave = async () => {
    if (!processedExcelBase64) {
      setErrorMessage("❌ No processed file to save");
      return;
    }

    clearMessages();
    setLoading(true);
    setLoadingMessage("Saving file to database...");
    
    try {
      const res = await api.post("/save-budget-file", {
        file_data: processedExcelBase64,
        filename: customFilename?.trim() || "Processed_Budget.xlsx"
      });
      setSuccessMessage(`✅ File saved to database successfully! File ID: ${res.data.id}`);
    } catch (err) {
      console.error("Save error", err);
      setErrorMessage("❌ Failed to save file to database. Please try again.");
    } finally {
      setLoading(false);
      setLoadingMessage("");
    }
  };

  return (
    <div className="p-4 max-w-4xl mx-auto">
      <h2 className="text-2xl font-bold mb-6 text-gray-800">Budget Processing Tool</h2>

      {/* File Upload Section */}
      <div className="mb-6 p-4 border-2 border-dashed border-gray-300 rounded-lg">
        <label className="block text-sm font-medium text-gray-700 mb-2">
          Select Excel File (.xlsx, .xls)
        </label>
        <input 
          type="file" 
          accept=".xlsx,.xls" 
          onChange={handleFileUpload} 
          className="mb-2 block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100" 
          disabled={loading}
        />
      </div>

      {/* Loading Indicator */}
      {loading && (
        <div className="mb-4 p-4 bg-blue-50 border border-blue-200 rounded-lg">
          <div className="flex items-center gap-3">
            <svg className="animate-spin h-5 w-5 text-blue-600" viewBox="0 0 24 24">
              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"></path>
            </svg>
            <span className="text-blue-800 font-medium">{loadingMessage}</span>
          </div>
        </div>
      )}

      {/* Success Message */}
      {successMessage && (
        <div className="mb-4 p-4 bg-green-50 border border-green-200 rounded-lg">
          <p className="text-green-800">{successMessage}</p>
        </div>
      )}

      {/* Error Message */}
      {errorMessage && (
        <div className="mb-4 p-4 bg-red-50 border border-red-200 rounded-lg">
          <p className="text-red-800">{errorMessage}</p>
        </div>
      )}

      {/* Sheet Configuration */}
      {sheetNames.length > 0 && (
        <div className="mb-6 p-4 bg-gray-50 rounded-lg">
          <h3 className="text-lg font-semibold mb-3 text-gray-800">Sheet Configuration</h3>
          
          <div className="grid md:grid-cols-2 gap-4 mb-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Select Sheet</label>
              <select
                value={selectedSheet}
                onChange={(e) => setSelectedSheet(e.target.value)}
                className="w-full border border-gray-300 rounded-md p-2"
                disabled={loading}
              >
                {sheetNames.map((sheet, i) => (
                  <option key={i} value={sheet}>{sheet}</option>
                ))}
              </select>
            </div>
            
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Header Row</label>
              <input
                type="number"
                value={headerRow}
                onChange={(e) => setHeaderRow(parseInt(e.target.value))}
                className="w-full border border-gray-300 rounded-md p-2"
                disabled={loading}
                min="1"
              />
            </div>
          </div>

          <button 
            onClick={handlePreview} 
            className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-md font-medium transition-colors"
            disabled={loading}
          >
            Load Sheet Preview
          </button>
        </div>
      )}

      {/* Column Mapping */}
      {columns.length > 0 && (
        <div className="mb-6 p-4 bg-gray-50 rounded-lg">
          <h3 className="text-lg font-semibold mb-3 text-gray-800">Column Mapping</h3>
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
                <label className="block text-sm font-medium text-gray-700 mb-1">{label}</label>
                <select
                  value={colMap[key]}
                  onChange={(e) => setColMap((prev) => ({ ...prev, [key]: e.target.value }))}
                  className="w-full border border-gray-300 rounded-md p-2"
                  disabled={loading}
                >
                  <option value="">-- Select Column --</option>
                  {columns.map((col, i) => (
                    <option key={i} value={col}>{col}</option>
                  ))}
                </select>
              </div>
            ))}
          </div>

          <button 
            onClick={handleProcess} 
            className="bg-green-600 hover:bg-green-700 text-white px-6 py-2 rounded-md font-medium transition-colors"
            disabled={loading}
          >
            Process Budget File
          </button>
        </div>
      )}

      {/* Processing Results */}
      {metrics && (
        <div className="mb-6 p-4 bg-gray-50 rounded-lg">
          <h3 className="text-lg font-semibold mb-3 text-gray-800">Processing Results</h3>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            {Object.entries(metrics).map(([key, val]) => (
              <div key={key} className="p-3 bg-white rounded-md shadow-sm border">
                <p className="text-sm text-gray-600 capitalize">{key.replace(/_/g, " ")}</p>
                <p className="text-xl font-bold text-gray-900">{val}</p>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Download and Save Section */}
      {processedExcelBase64 && (
        <div className="mb-6 p-4 bg-gray-50 rounded-lg">
          <h3 className="text-lg font-semibold mb-3 text-gray-800">Download & Save</h3>
          
          {downloadLink && (
            <div className="mb-4">
              <a
                href={downloadLink}
                download="processed_budget.xlsx"
                className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-md inline-block font-medium transition-colors"
              >
                📥 Download Processed File
              </a>
            </div>
          )}

          <div className="mb-4">
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Filename for Database Storage
            </label>
            <input
              type="text"
              value={customFilename}
              onChange={(e) => setCustomFilename(e.target.value)}
              className="w-full border border-gray-300 rounded-md p-2"
              placeholder="e.g., Apr-2025_Processed_Budget"
              disabled={loading}
            />
          </div>

          <button
            className="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-2 rounded-md font-medium transition-colors"
            onClick={handleSave}
            disabled={loading}
          >
            💾 Save to Database
          </button>
        </div>
      )}
    </div>
  );
};

export default BudgetProcessor;
