import React, { useState, useEffect } from "react";
import api from "../api/axios";

const SalesProcessor = () => {
  const [file, setFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headerRow, setHeaderRow] = useState(1);
  const [columns, setColumns] = useState([]);
  const [downloadLink, setDownloadLink] = useState(null);
  const [processedExcelBase64, setProcessedExcelBase64] = useState(null);
  const [customFilename, setCustomFilename] = useState("");

  // Enhanced loading states
  const [loadingSheets, setLoadingSheets] = useState(false);
  const [loadingPreview, setLoadingPreview] = useState(false);
  const [loadingProcess, setLoadingProcess] = useState(false);
  const [loadingSave, setLoadingSave] = useState(false);

  // Success/Error states
  const [successMessage, setSuccessMessage] = useState("");
  const [errorMessage, setErrorMessage] = useState("");

  // Column mappings
  const [execCodeCol, setExecCodeCol] = useState("");
  const [execNameCol, setExecNameCol] = useState("");
  const [productCol, setProductCol] = useState("");
  const [unitCol, setUnitCol] = useState("");
  const [quantityCol, setQuantityCol] = useState("");
  const [valueCol, setValueCol] = useState("");

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

  const autoDetectColumn = (columns, target, fallback = "") => {
    const lowerTarget = target.toLowerCase();
    return (
      columns.find((col) => col.toLowerCase().includes(lowerTarget)) ||
      fallback
    );
  };

  const handleFileUpload = async (e) => {
    const selectedFile = e.target.files[0];
    if (!selectedFile) return;

    setFile(selectedFile);
    setLoadingSheets(true);
    setErrorMessage("");

    try {
      const formData = new FormData();
      formData.append("file", selectedFile);
      const res = await api.post("/upload-tools/sheet-names", formData);
      const names = res.data.sheet_names;
      setSheetNames(names);
      setSelectedSheet(names.find(n => n.toLowerCase().includes("sales")) || names[0]);
      
      setSuccessMessage(`Successfully loaded ${names.length} sheet(s) from file`);
      
      // Reset dependent states
      setColumns([]);
      setDownloadLink(null);
      setProcessedExcelBase64(null);
      
      // Reset column mappings
      setExecCodeCol("");
      setExecNameCol("");
      setProductCol("");
      setUnitCol("");
      setQuantityCol("");
      setValueCol("");
    } catch (err) {
      console.error("Error loading sheets:", err);
      setErrorMessage("Failed to load sheets from file. Please check the file format.");
    } finally {
      setLoadingSheets(false);
    }
  };

  const handlePreview = async () => {
    if (!file || !selectedSheet) {
      setErrorMessage("Please select a file and sheet first.");
      return;
    }

    setLoadingPreview(true);
    setErrorMessage("");

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("sheet_name", selectedSheet);
      formData.append("header_row", headerRow);

      const res = await api.post("/upload-tools/preview", formData);
      setColumns(res.data.columns);

      // Auto-detect columns
      setExecCodeCol(autoDetectColumn(res.data.columns, "executive code"));
      setExecNameCol(autoDetectColumn(res.data.columns, "executive name"));
      setProductCol(autoDetectColumn(res.data.columns, "type (make)"));
      setUnitCol(autoDetectColumn(res.data.columns, "uom"));
      setQuantityCol(autoDetectColumn(res.data.columns, "quantity"));
      setValueCol(autoDetectColumn(res.data.columns, "value"));

      setSuccessMessage(`Successfully previewed ${res.data.columns.length} columns. Auto-mapped available fields.`);
    } catch (err) {
      console.error("Preview error:", err);
      setErrorMessage("Failed to load preview. Please check your file and sheet selection.");
    } finally {
      setLoadingPreview(false);
    }
  };

  const handleProcess = async () => {
    if (!execCodeCol) {
      setErrorMessage("Please map at least the Executive Code column before processing.");
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
      formData.append("exec_name_col", execNameCol);
      formData.append("product_col", productCol);
      formData.append("unit_col", unitCol);
      formData.append("quantity_col", quantityCol);
      formData.append("value_col", valueCol);

      const res = await api.post("/upload-sales-file", formData);

      const byteCharacters = atob(res.data.file_data);
      const byteNumbers = new Array(byteCharacters.length)
        .fill()
        .map((_, i) => byteCharacters.charCodeAt(i));
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      setDownloadLink(url);
      setProcessedExcelBase64(res.data.file_data);

      setSuccessMessage("Sales file processed successfully! You can now download or save the processed file.");
    } catch (err) {
      console.error("Process error:", err);
      setErrorMessage("Failed to process file. Please check your data and try again.");
    } finally {
      setLoadingProcess(false);
    }
  };

  const handleSave = async () => {
    if (!processedExcelBase64) {
      setErrorMessage("No processed file to save. Please process a file first.");
      return;
    }

    setLoadingSave(true);
    setErrorMessage("");

    try {
      const res = await api.post("/save-sales-file", {
        file_data: processedExcelBase64,
        filename: customFilename?.trim() || "Processed_Sales.xlsx",
      });
      
      setSuccessMessage(`File saved to database successfully! File ID: ${res.data.id}`);
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
    <div className="bg-white p-6 rounded shadow">
      <h2 className="text-xl font-semibold mb-4">Sales File Processing</h2>

      <MessageDisplay />

      <div className="space-y-4">
        {/* File Upload Section */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Select Excel File
          </label>
          <input
            type="file"
            accept=".xls,.xlsx"
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
                  onChange={(e) => setSelectedSheet(e.target.value)}
                >
                  {sheetNames.map((sheet) => (
                    <option key={sheet} value={sheet}>{sheet}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">
                  Header Row
                </label>
                <input
                  type="number"
                  min={0}
                  className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  value={headerRow}
                  onChange={(e) => setHeaderRow(parseInt(e.target.value))}
                />
              </div>
            </div>

            <button
              onClick={handlePreview}
              disabled={loadingPreview || !selectedSheet}
              className="inline-flex items-center px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {loadingPreview && <LoadingSpinner />}
              {loadingPreview ? 'Loading Sheet...' : 'Load Sheet'}
            </button>
          </div>
        )}

        {/* Column Mapping */}
        {columns.length > 0 && (
          <div className="bg-gray-50 p-4 rounded-md space-y-4">
            <h4 className="font-medium text-gray-900">Column Mapping</h4>
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Executive Code Column <span className="text-red-500">*</span>
                  </label>
                  <select
                    value={execCodeCol}
                    onChange={(e) => setExecCodeCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- Select Column --</option>
                    {columns.map((col) => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Executive Name Column
                  </label>
                  <select
                    value={execNameCol}
                    onChange={(e) => setExecNameCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- None --</option>
                    {columns.map((col) => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Product Column
                  </label>
                  <select
                    value={productCol}
                    onChange={(e) => setProductCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- None --</option>
                    {columns.map((col) => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Unit Column
                  </label>
                  <select
                    value={unitCol}
                    onChange={(e) => setUnitCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- None --</option>
                    {columns.map((col) => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Quantity Column
                  </label>
                  <select
                    value={quantityCol}
                    onChange={(e) => setQuantityCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- None --</option>
                    {columns.map((col) => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Value Column
                  </label>
                  <select
                    value={valueCol}
                    onChange={(e) => setValueCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- None --</option>
                    {columns.map((col) => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            <button
              onClick={handleProcess}
              disabled={loadingProcess || !execCodeCol}
              className="inline-flex items-center px-6 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {loadingProcess && <LoadingSpinner />}
              {loadingProcess ? 'Processing File...' : 'Process Sales File'}
            </button>

            {loadingProcess && (
              <div className="text-sm text-gray-600">
                Please wait while we process your sales file. This may take a few moments...
              </div>
            )}
          </div>
        )}

        {/* Download and Save Section */}
        {downloadLink && (
          <div className="bg-gray-50 p-4 rounded-md space-y-4">
            <div className="flex flex-wrap gap-4">
              <a
                href={downloadLink}
                download="processed_sales.xlsx"
                className="inline-flex items-center px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700"
              >
                Download Processed Sales File
              </a>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Enter Filename to Save to Database
              </label>
              <input
                type="text"
                value={customFilename}
                onChange={(e) => setCustomFilename(e.target.value)}
                className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                placeholder="e.g., Apr-2025_Processed_Sales"
              />
            </div>

            {processedExcelBase64 && (
              <button
                onClick={handleSave}
                disabled={loadingSave}
                className="inline-flex items-center px-4 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {loadingSave && <LoadingSpinner />}
                {loadingSave ? 'Saving...' : 'Save Sales File to DB'}
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default SalesProcessor;
