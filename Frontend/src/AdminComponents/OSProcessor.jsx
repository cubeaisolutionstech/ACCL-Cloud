import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import api from "../api/axios";

const OSProcessor = () => {
  const [file, setFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headerRow, setHeaderRow] = useState(1);
  const [columns, setColumns] = useState([]);
  const [execCodeCol, setExecCodeCol] = useState("");
  const [preview, setPreview] = useState([]);
  const [processedFile, setProcessedFile] = useState(null);
  const [customFilename, setCustomFilename] = useState("");

  // Enhanced loading states
  const [loadingSheets, setLoadingSheets] = useState(false);
  const [loadingProcess, setLoadingProcess] = useState(false);
  const [loadingSave, setLoadingSave] = useState(false);

  // Success/Error states
  const [successMessage, setSuccessMessage] = useState("");
  const [errorMessage, setErrorMessage] = useState("");

  // Auto-hide messages after 5 seconds
  useEffect(() => {
    if (successMessage || errorMessage) {
      const timer = setTimeout(() => {
        setSuccessMessage("");
        setErrorMessage("");
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [successMessage, errorMessage]);

  const handleFileUpload = (e) => {
    const uploaded = e.target.files[0];
    if (!uploaded) return;

    setFile(uploaded);
    setLoadingSheets(true);
    setErrorMessage("");

    try {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const workbook = XLSX.read(evt.target.result, { type: "binary" });
        const sheets = workbook.SheetNames;
        setSheetNames(sheets);
        setSelectedSheet(sheets[0]);

        const sheet = workbook.Sheets[sheets[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { header: headerRow, raw: false });
        setPreview(data.slice(0, 10));

        const headers = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          range: headerRow,
        })[0];
        setColumns(headers || []);
        setExecCodeCol(headers?.find(h => h.toLowerCase().includes("executive code")) || "");

        setSuccessMessage(`Successfully loaded ${sheets.length} sheet(s) from file. Auto-detected columns.`);
        
        // Reset dependent states
        setProcessedFile(null);
        setLoadingSheets(false);
      };
      reader.onerror = () => {
        setErrorMessage("Failed to read file. Please check the file format.");
        setLoadingSheets(false);
      };
      reader.readAsBinaryString(uploaded);
    } catch (err) {
      console.error("Error loading file:", err);
      setErrorMessage("Failed to load file. Please check the file format.");
      setLoadingSheets(false);
    }
  };

  const handleSheetChange = (newSheet) => {
    setSelectedSheet(newSheet);
    if (!file) return;

    try {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const workbook = XLSX.read(evt.target.result, { type: "binary" });
        const sheet = workbook.Sheets[newSheet];
        const data = XLSX.utils.sheet_to_json(sheet, { header: headerRow, raw: false });
        setPreview(data.slice(0, 10));

        const headers = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          range: headerRow,
        })[0];
        setColumns(headers || []);
        setExecCodeCol(headers?.find(h => h.toLowerCase().includes("executive code")) || "");
      };
      reader.readAsBinaryString(file);
    } catch (err) {
      console.error("Error changing sheet:", err);
      setErrorMessage("Failed to load sheet data.");
    }
  };

  const handleProcess = async () => {
    if (!file || !selectedSheet || !execCodeCol) {
      setErrorMessage("Please select a file, sheet, and Executive Code column before processing.");
      return;
    }

    setLoadingProcess(true);
    setErrorMessage("");

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("sheet_name", selectedSheet);
      formData.append("header_row", headerRow);
      formData.append("exec_code_col", execCodeCol);

      const res = await api.post("/process-os-file", formData, {
        responseType: "blob",
      });
      
      const processedUrl = URL.createObjectURL(res.data);
      setProcessedFile(processedUrl);
      setSuccessMessage("OS file processed successfully! You can now download or save the processed file.");
    } catch (err) {
      console.error("Processing error:", err);
      setErrorMessage("Failed to process file. Please check your data and try again.");
    } finally {
      setLoadingProcess(false);
    }
  };

  const handleSaveToDb = async () => {
    if (!file || !customFilename.trim()) {
      setErrorMessage("Please select a file and enter a filename before saving.");
      return;
    }

    setLoadingSave(true);
    setErrorMessage("");

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("filename", customFilename.trim());
      
      const res = await api.post("/save-os-file", formData);
      setSuccessMessage(`OS file saved to database successfully! File ID: ${res.data?.id || 'N/A'}`);
      setCustomFilename(""); // Clear filename after successful save
    } catch (err) {
      console.error("Save error:", err);
      setErrorMessage("Failed to save file to database. Please try again.");
    } finally {
      setLoadingSave(false);
    }
  };

  // Success/Error Message Component
  const MessageDisplay = () => {
    if (!successMessage && !errorMessage) return null;
    
    return (
      <div className={`p-4 rounded-md mb-4 ${
        successMessage 
          ? 'bg-green-50 border border-green-200 text-green-800' 
          : 'bg-red-50 border border-red-200 text-red-800'
      }`}>
        <div className="flex items-center">
          <div className="flex-shrink-0">
            {successMessage ? (
              <svg className="h-5 w-5 text-green-400" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
            ) : (
              <svg className="h-5 w-5 text-red-400" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
              </svg>
            )}
          </div>
          <div className="ml-3">
            <p className="text-sm font-medium">
              {successMessage || errorMessage}
            </p>
          </div>
          <div className="ml-auto">
            <button
              className="inline-flex text-gray-400 hover:text-gray-600"
              onClick={() => {
                setSuccessMessage("");
                setErrorMessage("");
              }}
            >
              <svg className="h-4 w-4" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" />
              </svg>
            </button>
          </div>
        </div>
      </div>
    );
  };

  // Loading Spinner Component
  const LoadingSpinner = () => (
    <svg className="animate-spin -ml-1 mr-3 h-4 w-4 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
    </svg>
  );

  return (
    <div className="bg-white shadow p-6 rounded">
      <h2 className="text-xl font-bold mb-4">Process OS File</h2>

      <MessageDisplay />

      <div className="space-y-4">
        {/* File Upload Section */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Select Excel File
          </label>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileUpload}
            className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
          />
          {loadingSheets && (
            <div className="mt-2 flex items-center text-blue-600">
              <LoadingSpinner />
              <span className="text-sm">Loading sheets...</span>
            </div>
          )}
        </div>

        {/* Sheet Configuration */}
        {sheetNames.length > 0 && (
          <div className="bg-gray-50 p-4 rounded-md space-y-4">
            <div className="grid md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">
                  Sheet Name
                </label>
                <select
                  className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  value={selectedSheet}
                  onChange={(e) => handleSheetChange(e.target.value)}
                >
                  {sheetNames.map((name, idx) => (
                    <option key={idx} value={name}>{name}</option>
                  ))}
                </select>
              </div>
              
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">
                  Header Row (0-based)
                </label>
                <input
                  type="number"
                  min="0"
                  className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  value={headerRow}
                  onChange={(e) => setHeaderRow(Number(e.target.value))}
                />
              </div>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                Executive Code Column <span className="text-red-500">*</span>
              </label>
              <select
                className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                value={execCodeCol}
                onChange={(e) => setExecCodeCol(e.target.value)}
              >
                <option value="">-- Select Column --</option>
                {columns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>

            <button
              onClick={handleProcess}
              disabled={loadingProcess || !execCodeCol}
              className="inline-flex items-center px-6 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {loadingProcess && <LoadingSpinner />}
              {loadingProcess ? 'Processing File...' : 'Process OS File'}
            </button>

            {loadingProcess && (
              <div className="text-sm text-gray-600">
                Please wait while we process your OS file. This may take a few moments...
              </div>
            )}
          </div>
        )}

        {/* Download and Save Section */}
        {processedFile && (
          <div className="bg-gray-50 p-4 rounded-md space-y-4">
            <div className="flex flex-wrap gap-4">
              <a
                href={processedFile}
                download="processed_os.xlsx"
                className="inline-flex items-center px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700"
              >
                Download Processed File
              </a>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Enter Filename to Save to Database
              </label>
              <input
                type="text"
                placeholder="Enter file name"
                value={customFilename}
                onChange={(e) => setCustomFilename(e.target.value)}
                className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
              />
            </div>

            <button
              onClick={handleSaveToDb}
              disabled={loadingSave || !customFilename.trim()}
              className="inline-flex items-center px-4 py-2 bg-yellow-600 text-white rounded-md hover:bg-yellow-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {loadingSave && <LoadingSpinner />}
              {loadingSave ? 'Saving...' : 'Save to Database'}
            </button>
          </div>
        )}

        {/* Preview Section (Optional) */}
        {preview.length > 0 && (
          <div className="bg-white border rounded-md p-4">
            <h4 className="font-medium text-gray-900 mb-3">Data Preview (First 10 rows)</h4>
            <div className="overflow-x-auto">
              <table className="min-w-full table-auto border">
                <thead className="bg-gray-50">
                  <tr>
                    {Object.keys(preview[0]).slice(0, 6).map((col, i) => (
                      <th key={i} className="border px-2 py-1 text-sm font-medium text-gray-900">
                        {col}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.slice(0, 5).map((row, i) => (
                    <tr key={i} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                      {Object.values(row).slice(0, 6).map((val, j) => (
                        <td key={j} className="border px-2 py-1 text-xs text-gray-700">
                          {String(val).substring(0, 50)}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {Object.keys(preview[0]).length > 6 && (
                <p className="text-xs text-gray-500 mt-2">
                  Showing first 6 columns of {Object.keys(preview[0]).length} total columns
                </p>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default OSProcessor;
