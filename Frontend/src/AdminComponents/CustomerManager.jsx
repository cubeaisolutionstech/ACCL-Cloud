import React, { useEffect, useState } from "react";
import api from "../api/axios";
import * as XLSX from "xlsx";

const CustomerManager = ({ onDataUpdated }) => {
  const [execs, setExecs] = useState([]);
  const [selectedExec, setSelectedExec] = useState("");
  const [assignedCustomers, setAssignedCustomers] = useState([]);
  const [unmappedCustomers, setUnmappedCustomers] = useState([]);
  const [newCodes, setNewCodes] = useState("");

  const [sheets, setSheets] = useState([]);
  const [sheetData, setSheetData] = useState([]);
  const [execNameCol, setExecNameCol] = useState("");
  const [execCodeCol, setExecCodeCol] = useState("");
  const [custCodeCol, setCustCodeCol] = useState("");
  const [custNameCol, setCustNameCol] = useState("");
  const [selectedToRemove, setSelectedToRemove] = useState([]);
  const [selectedToAssign, setSelectedToAssign] = useState([]);

  // Enhanced loading and success states (matching BranchRegionManager)
  const [processing, setProcessing] = useState(false);
  const [loadingCustomers, setLoadingCustomers] = useState(false);
  const [loadingAssign, setLoadingAssign] = useState(false);
  const [loadingRemove, setLoadingRemove] = useState(false);
  const [successMessage, setSuccessMessage] = useState("");
  const [errorMessage, setErrorMessage] = useState("");

  // Safe render function for potentially nested objects
  const safeRender = (value) => {
    if (value === null || value === undefined) {
      return "";
    }
    if (typeof value === 'object') {
      // If it's an object, try to extract a meaningful value
      if (value.name) return String(value.name);
      if (value.code) return String(value.code);
      if (value.value) return String(value.value);
      // Fallback to stringified object
      return JSON.stringify(value);
    }
    return String(value);
  };

  // Auto-hide messages after 5 seconds (matching BranchRegionManager)
  useEffect(() => {
    if (successMessage || errorMessage) {
      const timer = setTimeout(() => {
        setSuccessMessage("");
        setErrorMessage("");
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [successMessage, errorMessage]);

  // Success/Error Message Component (matching BranchRegionManager)
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

  // Loading Spinner Component (matching BranchRegionManager)
  const LoadingSpinner = () => (
    <svg className="animate-spin -ml-1 mr-3 h-4 w-4 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
    </svg>
  );

  const guessColumn = (headers, type) => {
    const aliases = {
      executive_name: ["executive name", "empname", "executive"],
      executive_code: ["executive code", "empcode", "ecode"],
      customer_code: ["customer code", "slcode", "custcode"],
      customer_name: ["customer name", "slname", "custname"],
    };

    const candidates = aliases[type] || [type];
    const lowerHeaders = headers.map((h) => h.toLowerCase());

    for (let alias of candidates) {
      const match = lowerHeaders.find((h) => h === alias);
      if (match) return headers[lowerHeaders.indexOf(match)];
    }
    for (let alias of candidates) {
      const match = lowerHeaders.find((h) => h.includes(alias));
      if (match) return headers[lowerHeaders.indexOf(match)];
    }
    return "";
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setErrorMessage("");
    setSuccessMessage("");

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.SheetNames[0];
        setSheets(workbook.SheetNames);

        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {
          defval: "",
        });
        setSheetData(jsonData);

        if (jsonData.length > 0) {
          const headers = Object.keys(jsonData[0]);
          setExecNameCol(guessColumn(headers, "executive_name"));
          setExecCodeCol(guessColumn(headers, "executive_code"));
          setCustCodeCol(guessColumn(headers, "customer_code"));
          setCustNameCol(guessColumn(headers, "customer_name"));

          setSuccessMessage(`File uploaded successfully! Found ${jsonData.length} rows with auto-mapped columns. Ready to process.`);
        } else {
          setErrorMessage("File appears to be empty. Please check your Excel file.");
        }
      } catch (error) {
        setErrorMessage("Failed to read Excel file. Please ensure it's a valid .xlsx file.");
        console.error("Error reading Excel:", error);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleProcessFile = async () => {
    if (!sheetData.length) {
      setErrorMessage("Please upload a file first!");
      return;
    }

    if (!execNameCol || !custCodeCol) {
      setErrorMessage("Please select at least Executive Name and Customer Code columns.");
      return;
    }

    try {
      setProcessing(true);
      setErrorMessage("");
      setSuccessMessage("");

      const payload = {
        data: sheetData,
        execNameCol,
        execCodeCol,
        custCodeCol,
        custNameCol,
      };
      
      await api.post("/bulk-assign-customers", payload);

      await fetchCustomers();
      await fetchUnmapped();

      if (onDataUpdated) {
        onDataUpdated();
      }

      setSuccessMessage(`Bulk assignment complete! Processed ${sheetData.length} customer records successfully.`);
      
      // Reset form after successful processing
      setSheetData([]);
      setSheets([]);
    } catch (err) {
      setErrorMessage("Failed to process file. Please check your data format and try again.");
      console.error("Error processing file:", err);
    } finally {
      setProcessing(false);
    }
  };

  const fetchExecs = async () => {
    try {
      const res = await api.get("/executives");
      console.log("Executives response:", res.data); // Debug log
      setExecs(res.data);
      if (res.data.length > 0) {
        // Safe access to executive name
        const firstExecName = typeof res.data[0].name === 'object' 
          ? res.data[0].name.name || res.data[0].name.code || JSON.stringify(res.data[0].name)
          : res.data[0].name;
        setSelectedExec(firstExecName);
      }
    } catch (error) {
      setErrorMessage("Failed to load executives list.");
      console.error("Error fetching executives:", error);
    }
  };

  const fetchCustomers = async () => {
    if (!selectedExec) return;
    
    setLoadingCustomers(true);
    try {
      const res = await api.get(`/customers?executive=${selectedExec}`);
      setAssignedCustomers(res.data);
    } catch (error) {
      setErrorMessage("Failed to load customer assignments.");
      console.error("Error fetching customers:", error);
    } finally {
      setLoadingCustomers(false);
    }
  };

  const fetchUnmapped = async () => {
    try {
      const res = await api.get("/customers/unmapped");
      setUnmappedCustomers(res.data);
    } catch (error) {
      setErrorMessage("Failed to load unmapped customers.");
      console.error("Error fetching unmapped customers:", error);
    }
  };

  const handleRemove = async (codes) => {
    if (!codes.length) {
      setErrorMessage("Please select customers to remove.");
      return;
    }

    setLoadingRemove(true);
    setErrorMessage("");
    
    try {
      await api.post("/remove-customer", {
        executive: selectedExec,
        customers: codes,
      });
      
      await fetchCustomers();
      await fetchUnmapped();
      
      setSelectedToRemove([]);
      setSuccessMessage(`Successfully removed ${codes.length} customer(s) from ${selectedExec}.`);
      
      if (onDataUpdated) {
        onDataUpdated();
      }
    } catch (error) {
      setErrorMessage("Failed to remove customers. Please try again.");
      console.error("Error removing customers:", error);
    } finally {
      setLoadingRemove(false);
    }
  };

  const handleAssign = async (codes) => {
    if (!codes.length) {
      setErrorMessage("Please select customers to assign.");
      return;
    }

    setLoadingAssign(true);
    setErrorMessage("");
    
    try {
      await api.post("/assign-customer", {
        executive: selectedExec,
        customers: codes,
      });
      
      await fetchCustomers();
      await fetchUnmapped();
      
      setSelectedToAssign([]);
      setSuccessMessage(`Successfully assigned ${codes.length} customer(s) to ${selectedExec}.`);
      
      if (onDataUpdated) {
        onDataUpdated();
      }
    } catch (error) {
      setErrorMessage("Failed to assign customers. Please try again.");
      console.error("Error assigning customers:", error);
    } finally {
      setLoadingAssign(false);
    }
  };

  const handleAddNew = async () => {
    const codes = newCodes
      .split("\n")
      .map((code) => code.trim())
      .filter(Boolean);
    
    if (codes.length === 0) {
      setErrorMessage("Please enter at least one customer code.");
      return;
    }

    setLoadingAssign(true);
    setErrorMessage("");
    
    try {
      await api.post("/assign-customer", {
        executive: selectedExec,
        customers: codes,
      });
      
      setNewCodes("");
      await fetchCustomers();
      await fetchUnmapped();
      
      setSuccessMessage(`Successfully added ${codes.length} new customer(s) to ${selectedExec}.`);
      
      if (onDataUpdated) {
        onDataUpdated();
      }
    } catch (error) {
      setErrorMessage("Failed to add new customers. Please check the codes and try again.");
      console.error("Error adding customers:", error);
    } finally {
      setLoadingAssign(false);
    }
  };

  // Safe get executive identifier
  const getExecutiveId = (exec) => {
    if (typeof exec.name === 'object') {
      return exec.name?.name || exec.name?.code || JSON.stringify(exec.name);
    }
    return exec.name || exec.id || JSON.stringify(exec);
  };

  useEffect(() => {
    fetchExecs();
  }, []);

  useEffect(() => {
    if (selectedExec) {
      fetchCustomers();
      fetchUnmapped();
    }
  }, [selectedExec]);

  return (
    <div>
      <h2 className="text-xl font-bold mb-4">Customer Code Management</h2>

      <MessageDisplay />

      {/* ---------------- BULK UPLOAD ---------------- */}
      <div className="mb-6 p-4 bg-blue-50 rounded-md">
        <h3 className="font-semibold mb-3">Bulk Assignment via Excel</h3>
        
        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Select Excel File
            </label>
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={handleExcelUpload}
              className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            />
          </div>

          {sheets.length > 0 && sheetData.length > 0 && (
            <div className="bg-white p-4 rounded-md border">
              <h4 className="font-medium text-gray-900 mb-3">Column Mapping</h4>
              <div className="grid md:grid-cols-2 gap-4 mb-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Executive Name Column <span className="text-red-500">*</span>
                  </label>
                  <select
                    value={execNameCol}
                    onChange={(e) => setExecNameCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- Select Column --</option>
                    {Object.keys(sheetData[0]).map((col) => (
                      <option key={col} value={col}>
                        {col}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Executive Code Column
                  </label>
                  <select
                    value={execCodeCol}
                    onChange={(e) => setExecCodeCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- None --</option>
                    {Object.keys(sheetData[0]).map((col) => (
                      <option key={col} value={col}>
                        {col}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Customer Code Column <span className="text-red-500">*</span>
                  </label>
                  <select
                    value={custCodeCol}
                    onChange={(e) => setCustCodeCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- Select Column --</option>
                    {Object.keys(sheetData[0]).map((col) => (
                      <option key={col} value={col}>
                        {col}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Customer Name Column
                  </label>
                  <select
                    value={custNameCol}
                    onChange={(e) => setCustNameCol(e.target.value)}
                    className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="">-- None --</option>
                    {Object.keys(sheetData[0]).map((col) => (
                      <option key={col} value={col}>
                        {col}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <button
                onClick={handleProcessFile}
                disabled={processing || !execNameCol || !custCodeCol}
                className="inline-flex items-center px-6 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {processing && <LoadingSpinner />}
                {processing ? 'Processing File...' : 'Process File'}
              </button>

              {processing && (
                <div className="mt-3 text-sm text-gray-600">
                  Please wait while we process your file. This may take a few moments...
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* ---------------- MANUAL CUSTOMER MANAGEMENT ---------------- */}
      <div className="p-4 bg-gray-50 rounded-md">
        <h3 className="font-semibold mb-3">Manual Management</h3>

        <div className="mb-4">
          <label className="block text-sm font-medium text-gray-700 mb-2">Select Executive:</label>
          <select
            value={selectedExec}
            onChange={(e) => setSelectedExec(e.target.value)}
            className="w-full md:w-1/3 border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
          >
            {execs.map((exec, index) => (
              <option key={getExecutiveId(exec) || index} value={getExecutiveId(exec)}>
                {safeRender(exec.name)}
              </option>
            ))}
          </select>
        </div>

        <div className="grid md:grid-cols-2 gap-6">
          <div>
            <h4 className="font-medium mb-2">Assigned Customers</h4>
            {loadingCustomers ? (
              <div className="flex items-center justify-center py-8">
                <LoadingSpinner />
                <span className="ml-2 text-gray-600">Loading customers...</span>
              </div>
            ) : (
              <>
                <div className="border rounded-md p-3 mb-3 max-h-40 overflow-y-auto bg-white">
                  {assignedCustomers.length === 0 ? (
                    <p className="text-gray-500 text-sm">No customers assigned yet.</p>
                  ) : (
                    assignedCustomers.map((c, index) => (
                      <label key={c || index} className="flex items-center py-1">
                        <input
                          type="checkbox"
                          value={c}
                          checked={selectedToRemove.includes(c)}
                          onChange={(e) =>
                            setSelectedToRemove(
                              e.target.checked
                                ? [...selectedToRemove, c]
                                : selectedToRemove.filter((x) => x !== c)
                            )
                          }
                          className="mr-2 rounded"
                        />
                        <span className="text-sm">{safeRender(c)}</span>
                      </label>
                    ))
                  )}
                </div>
                <button
                  onClick={() => handleRemove(selectedToRemove)}
                  disabled={loadingRemove || selectedToRemove.length === 0}
                  className="inline-flex items-center px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {loadingRemove && <LoadingSpinner />}
                  {loadingRemove ? 'Removing...' : `Remove Selected (${selectedToRemove.length})`}
                </button>
              </>
            )}
          </div>

          <div>
            <h4 className="font-medium mb-2">Unmapped Customers</h4>
            <div className="border rounded-md p-3 mb-3 max-h-40 overflow-y-auto bg-white">
              {unmappedCustomers.length === 0 ? (
                <p className="text-gray-500 text-sm">All customers are assigned.</p>
              ) : (
                unmappedCustomers.map((c, index) => (
                  <label key={c || index} className="flex items-center py-1">
                    <input
                      type="checkbox"
                      value={c}
                      checked={selectedToAssign.includes(c)}
                      onChange={(e) =>
                        setSelectedToAssign(
                          e.target.checked
                            ? [...selectedToAssign, c]
                            : selectedToAssign.filter((x) => x !== c)
                        )
                      }
                      className="mr-2 rounded"
                    />
                    <span className="text-sm">{safeRender(c)}</span>
                  </label>
                ))
              )}
            </div>
            <button
              onClick={() => handleAssign(selectedToAssign)}
              disabled={loadingAssign || selectedToAssign.length === 0}
              className="inline-flex items-center px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {loadingAssign && <LoadingSpinner />}
              {loadingAssign ? 'Assigning...' : `Assign Selected (${selectedToAssign.length})`}
            </button>
          </div>
        </div>

        <hr className="my-6" />

        <div>
          <h4 className="font-medium mb-2">Add New Customers</h4>
          <textarea
            value={newCodes}
            onChange={(e) => setNewCodes(e.target.value)}
            placeholder="Enter customer codes (one per line)&#10;Example:&#10;CUST001&#10;CUST002&#10;CUST003"
            className="w-full border border-gray-300 rounded-md px-3 py-2 mb-3 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
            rows={4}
          />
          <button
            onClick={handleAddNew}
            disabled={loadingAssign || !newCodes.trim()}
            className="inline-flex items-center px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {loadingAssign && <LoadingSpinner />}
            {loadingAssign ? 'Adding...' : 'Add New Customers'}
          </button>
        </div>
      </div>
    </div>
  );
};

export default CustomerManager;
