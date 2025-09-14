import React, { useEffect, useState, forwardRef, useImperativeHandle } from "react";
import api from "../api/axios";

const BranchRegionManager = forwardRef((props, ref) => {
  const [activeTab, setActiveTab] = useState("manual");

  // shared states
  const [branches, setBranches] = useState([]);
  const [regions, setRegions] = useState([]);
  const [executives, setExecutives] = useState([]);
  const [mappings, setMappings] = useState([]);

  const [newBranch, setNewBranch] = useState("");
  const [newRegion, setNewRegion] = useState("");

  const [selectedBranch, setSelectedBranch] = useState("");
  const [branchExecs, setBranchExecs] = useState([]);

  const [selectedRegion, setSelectedRegion] = useState("");
  const [regionBranches, setRegionBranches] = useState([]);

  // Enhanced file upload states
  const [uploadFile, setUploadFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headerRow, setHeaderRow] = useState(0);
  const [columns, setColumns] = useState([]);
  const [mapping, setMapping] = useState({
    exec_name_col: "",
    exec_code_col: "",
    branch_col: "",
    region_col: "",
  });

  // Enhanced loading and success states
  const [loadingSheets, setLoadingSheets] = useState(false);
  const [loadingPreview, setLoadingPreview] = useState(false);
  const [loadingProcess, setLoadingProcess] = useState(false);
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

  const fetchAll = async () => {
    try {
      const [bRes, rRes, eRes, mRes] = await Promise.all([
        api.get("/branches"),
        api.get("/regions"),
        api.get("/executives"),
        api.get("/mappings"),
      ]);
      setBranches(bRes.data);
      setRegions(rRes.data);
      setExecutives(eRes.data);
      setMappings(mRes.data);
    } catch (error) {
      console.error("Error fetching data:", error);
      setErrorMessage("Failed to fetch data. Please refresh the page.");
    }
  };

  // Expose refresh function to parent component
  useImperativeHandle(ref, () => ({
    refreshData: () => {
      fetchAll();
      // Reset any local states that might be affected by reset
      setSelectedBranch("");
      setSelectedRegion("");
      setBranchExecs([]);
      setRegionBranches([]);
      setSuccessMessage("Data refreshed after mappings reset");
    }
  }));

  useEffect(() => {
    fetchAll();
  }, []);

  // Auto-Mapping Logic
  const autoMapColumns = (columnList) => {
    const normalize = (s) => s.toLowerCase().replace(/\s/g, "");
    const findMatch = (keywords) =>
      columnList.find((col) => keywords.some((key) => normalize(col).includes(key)));
    return {
      exec_name_col: findMatch(["executivename", "ename", "empname"]) || "",
      exec_code_col: findMatch(["executivecode", "code", "empcode"]) || "",
      branch_col: findMatch(["branch"]) || "",
      region_col: findMatch(["region"]) || "",
    };
  };

  const handleSheetLoad = async () => {
    setLoadingSheets(true);
    setErrorMessage("");
    try {
      const formData = new FormData();
      formData.append("file", uploadFile);
      const res = await api.post("/get-sheet-names", formData);
      const sheetList = res.data.sheets || [];
      setSheetNames(sheetList);
      
      if (sheetList.length > 0) {
        setSelectedSheet(sheetList[0]);
        setSuccessMessage(`Successfully loaded ${sheetList.length} sheet(s) from file`);
      }
    } catch (err) {
      console.error("Error loading sheets", err);
      setErrorMessage("Failed to load sheets from file. Please check the file format.");
    } finally {
      setLoadingSheets(false);
    }
  };

  const handlePreview = async () => {
    setLoadingPreview(true);
    setErrorMessage("");
    try {
      const formData = new FormData();
      formData.append("file", uploadFile);
      formData.append("sheet_name", selectedSheet);
      formData.append("header_row", headerRow);
      const res = await api.post("/preview-excel", formData);
      const cols = res.data.columns;
      setColumns(cols);
      setMapping(autoMapColumns(cols)); // auto-mapping here
      setSuccessMessage(`Successfully previewed ${cols.length} columns. Auto-mapped available fields.`);
    } catch (err) {
      console.error("Error previewing columns", err);
      setErrorMessage("Failed to preview columns. Please check your file and sheet selection.");
    } finally {
      setLoadingPreview(false);
    }
  };

  const handleUpload = async () => {
    if (!mapping.exec_name_col || !mapping.branch_col || !mapping.region_col) {
      setErrorMessage("Please map Executive Name, Branch, and Region columns before processing.");
      return;
    }

    setLoadingProcess(true);
    setErrorMessage("");
    try {
      const formData = new FormData();
      formData.append("file", uploadFile);
      formData.append("sheet_name", selectedSheet);
      formData.append("header_row", headerRow);
      formData.append("exec_name_col", mapping.exec_name_col);
      formData.append("exec_code_col", mapping.exec_code_col);
      formData.append("branch_col", mapping.branch_col);
      formData.append("region_col", mapping.region_col);

      await api.post("/upload-branch-region-file", formData);
      
      // Reset form after successful upload
      setUploadFile(null);
      setSheetNames([]);
      setSelectedSheet("");
      setColumns([]);
      setMapping({
        exec_name_col: "",
        exec_code_col: "",
        branch_col: "",
        region_col: "",
      });
      
      setSuccessMessage("File processed successfully! Branch-Region mappings have been updated.");
      fetchAll();
    } catch (err) {
      console.error("Error processing file", err);
      setErrorMessage("Failed to process file. Please check your data and try again.");
    } finally {
      setLoadingProcess(false);
    }
  };

  useEffect(() => {
    const fetchBranchExecs = async () => {
      if (selectedBranch) {
        try {
          const res = await api.get(`/branch/${selectedBranch}/executives`);
          setBranchExecs(res.data);  // already selected ones
        } catch (error) {
          console.error("Error fetching branch executives:", error);
        }
      }
    };
    fetchBranchExecs();
  }, [selectedBranch]);

  // when a region is selected — get current branches
  useEffect(() => {
    const fetchRegionBranches = async () => {
      if (selectedRegion) {
        try {
          const res = await api.get(`/region/${selectedRegion}/branches`);
          setRegionBranches(res.data); // already selected ones
        } catch (error) {
          console.error("Error fetching region branches:", error);
        }
      }
    };
    fetchRegionBranches();
  }, [selectedRegion]);

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
    <div className="p-6">
      <h2 className="text-xl font-bold mb-4">Branch & Region Mapping</h2>

      <MessageDisplay />

      {/* UPDATED GRAY TABS - NOW CONSISTENT WITH ExecutiveManagement */}
      <div className="flex gap-4 mb-6">
        <button
          className="px-4 py-2 rounded bg-gray-300 text-black hover:bg-gray-400 transition-colors"
          onClick={() => setActiveTab("manual")}
        >
          Manual Entry
        </button>
        <button
          className="px-4 py-2 rounded bg-gray-300 text-black hover:bg-gray-400 transition-colors"
          onClick={() => setActiveTab("upload")}
        >
          File Upload
        </button>
      </div>

      {/* Manual Entry Tab */}
      {activeTab === "manual" && (
        <div className="bg-white p-4 border rounded shadow-sm">
          <div className="grid md:grid-cols-2 gap-4">
            <div>
              <h3 className="font-semibold">Create Branch</h3>
              <input
                type="text"
                value={newBranch}
                onChange={(e) => setNewBranch(e.target.value)}
                className="border px-3 py-2 rounded w-full mt-1 mb-2"
              />
              <button onClick={async () => {
                await api.post("/branch", { name: newBranch });
                setNewBranch("");
                fetchAll();
              }} className="bg-green-600 text-white px-4 py-2 rounded">Create</button>
            </div>
            <div>
              <h3 className="font-semibold">Create Region</h3>
              <input
                type="text"
                value={newRegion}
                onChange={(e) => setNewRegion(e.target.value)}
                className="border px-3 py-2 rounded w-full mt-1 mb-2"
              />
              <button onClick={async () => {
                await api.post("/region", { name: newRegion });
                setNewRegion("");
                fetchAll();
              }} className="bg-green-600 text-white px-4 py-2 rounded">Create</button>
            </div>
          </div>

          <hr className="my-6" />

          {/* Current Branches Table */}
          <h3 className="font-semibold mt-4 mb-2">Current Branches</h3>
          <table className="w-full border text-sm mb-4">
            <thead className="bg-gray-100">
              <tr>
                <th className="border px-2 py-1">Branch</th>
                <th className="border px-2 py-1">Executives</th>
                <th className="border px-2 py-1">Regions</th>
              </tr>
            </thead>
            <tbody>
              {branches.map((b) => {
                const execCount = mappings.find((m) => m.branch === b.name)?.executives?.length || 0;
                const regionName = mappings.find((m) => m.branch === b.name)?.region || "Unmapped";
                return (
                  <tr key={b.id}>
                    <td className="border px-2 py-1">{b.name}</td>
                    <td className="border px-2 py-1">{execCount}</td>
                    <td className="border px-2 py-1">{regionName}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {/* Delete Branch */}
          <div className="mb-6">
            <label className="block font-medium mb-1">Remove Branch</label>
            <select
              className="border px-2 py-1 mr-2"
              onChange={(e) => setSelectedBranch(e.target.value)}
              value={selectedBranch}
            >
              <option value="">Select Branch</option>
              {branches.map((b) => (
                <option key={b.id} value={b.name}>{b.name}</option>
              ))}
            </select>
            <button
              className="bg-red-600 text-white px-3 py-1 rounded"
              onClick={async () => {
                if (selectedBranch && window.confirm(`Delete branch '${selectedBranch}'?`)) {
                  await api.delete(`/branch/${selectedBranch}`);
                  setSelectedBranch("");
                  fetchAll();
                }
              }}
              disabled={!selectedBranch}
            >
              Delete
            </button>
          </div>

          {/* Current Regions Table */}
          <h3 className="font-semibold mt-4 mb-2">Current Regions</h3>
          <table className="w-full border text-sm mb-4">
            <thead className="bg-gray-100">
              <tr>
                <th className="border px-2 py-1">Region</th>
                <th className="border px-2 py-1"># Branches</th>
              </tr>
            </thead>
            <tbody>
              {regions.map((r) => {
                const branchCount = mappings.filter((m) => m.region === r.name).length;
                return (
                  <tr key={r.id}>
                    <td className="border px-2 py-1">{r.name}</td>
                    <td className="border px-2 py-1">{branchCount}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {/* Delete Region */}
          <div className="mb-6">
            <label className="block font-medium mb-1">Remove Region</label>
            <select
              className="border px-2 py-1 mr-2"
              onChange={(e) => setSelectedRegion(e.target.value)}
              value={selectedRegion}
            >
              <option value="">Select Region</option>
              {regions.map((r) => (
                <option key={r.id} value={r.name}>{r.name}</option>
              ))}
            </select>
            <button
              className="bg-red-600 text-white px-3 py-1 rounded"
              onClick={async () => {
                if (selectedRegion && window.confirm(`Delete region '${selectedRegion}'?`)) {
                  await api.delete(`/region/${selectedRegion}`);
                  setSelectedRegion("");
                  fetchAll();
                }
              }}
              disabled={!selectedRegion}
            >
              Delete
            </button>
          </div>

          <div className="grid md:grid-cols-2 gap-6">
            <div>
              <h4 className="font-medium mb-1">Map Executives to Branch</h4>
              <select className="border px-2 py-1 mb-2 w-full" value={selectedBranch} onChange={(e) => setSelectedBranch(e.target.value)}>
                <option value="">Select Branch</option>
                {branches.map((b) => (
                  <option key={b.id} value={b.name}>{b.name}</option>
                ))}
                </select>
               <div className="w-full border p-2 h-40 mb-2 overflow-y-auto">
                {[
                  ...executives.filter((e) => branchExecs.includes(e.name)),
                  ...executives.filter((e) => !branchExecs.includes(e.name)),
                ].map((e) => (
                  <label key={e.id} className="block">
                    <input
                      type="checkbox"
                      value={e.name}
                      checked={branchExecs.includes(e.name)}
                      onChange={(ev) => {
                        if (ev.target.checked) {
                          setBranchExecs([...branchExecs, e.name]);
                        } else {
                          setBranchExecs(branchExecs.filter((name) => name !== e.name));
                        }
                      }}
                      className="mr-2"
                    />
                    {e.name}
                  </label>
                ))}
              </div>
              <button onClick={async () => {
                await api.post("/map-branch-executives", {
                  branch: selectedBranch,
                  executives: branchExecs
                });
                fetchAll();
              }} className="bg-blue-600 text-white px-4 py-2 rounded">
                Update Mapping
              </button>
            </div>

            <div>
              <h4 className="font-medium mb-1">Map Branches to Region</h4>
              <select className="border px-2 py-1 mb-2 w-full" value={selectedRegion} onChange={(e) => setSelectedRegion(e.target.value)}>
                <option value="">Select Region</option>
                {regions.map((r) => (
                  <option key={r.id} value={r.name}>{r.name}</option>
                ))}
              </select>
              <div className="w-full border p-2 h-40 mb-2 overflow-y-auto">
                {[
                  ...branches.filter((b) => regionBranches.includes(b.name)),
                  ...branches.filter((b) => !regionBranches.includes(b.name)),
                ].map((b) => (
                  <label key={b.id} className="block">
                    <input
                      type="checkbox"
                      value={b.name}
                      checked={regionBranches.includes(b.name)}
                      onChange={(ev) => {
                        if (ev.target.checked) {
                          setRegionBranches([...regionBranches, b.name]);
                        } else {
                          setRegionBranches(regionBranches.filter((name) => name !== b.name));
                        }
                      }}
                      className="mr-2"
                    />
                    {b.name}
                  </label>
                ))}
              </div>
              <button onClick={async () => {
                await api.post("/map-region-branches", {
                  region: selectedRegion,
                  branches: regionBranches
                });
                fetchAll();
              }} className="bg-blue-600 text-white px-4 py-2 rounded">
                Update Mapping
              </button>
            </div>
          </div>

          <hr className="my-6" />

          {/* Mapping Summary */}
          <h3 className="font-semibold mb-2">Current Mappings</h3>
          <table className="w-full border text-sm mb-6">
            <thead className="bg-gray-100">
              <tr>
                <th className="border px-2 py-1">Branch</th>
                <th className="border px-2 py-1">Region</th>
                <th className="border px-2 py-1">Executives</th>
                <th className="border px-2 py-1">Count</th>
              </tr>
            </thead>
            <tbody>
              {mappings.map((m, idx) => (
                <tr key={idx}>
                  <td className="border px-2 py-1">{m.branch}</td>
                  <td className="border px-2 py-1">{m.region}</td>
                  <td className="border px-2 py-1 max-w-xs overflow-hidden overflow-ellipsis relative" title={m.executives.join(", ")}>
                    <div className="whitespace-normal break-words">{m.executives.join(", ")}</div>
                  </td>
                  <td className="border px-2 py-1">{m.count}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* File Upload Tab */}
      {activeTab === "upload" && (
        <div className="bg-white p-4 border rounded shadow-sm">
          <h3 className="font-semibold mb-4">Upload Branch-Region Mapping File</h3>
          
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Select Excel File
              </label>
              <input
                type="file"
                accept=".xlsx,.xls"
                className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
                onChange={(e) => {
                  setUploadFile(e.target.files[0]);
                  // Reset dependent states when file changes
                  setSheetNames([]);
                  setSelectedSheet("");
                  setColumns([]);
                  setMapping({
                    exec_name_col: "",
                    exec_code_col: "",
                    branch_col: "",
                    region_col: "",
                  });
                  setSuccessMessage("");
                  setErrorMessage("");
                }}
              />
            </div>

            {uploadFile && (
              <>
                <div className="bg-gray-50 p-4 rounded-md">
                  <p className="text-sm text-gray-600 mb-3">
                    File selected: <span className="font-medium">{uploadFile.name}</span>
                  </p>
                  
                  <button
                    onClick={handleSheetLoad}
                    disabled={loadingSheets}
                    className="inline-flex items-center px-4 py-2 bg-blue-500 text-white rounded-md hover:bg-blue-600 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {loadingSheets && <LoadingSpinner />}
                    {loadingSheets ? 'Loading Sheets...' : 'Load Sheets'}
                  </button>
                </div>

                {sheetNames.length > 0 && (
                  <div className="grid md:grid-cols-2 gap-4 bg-gray-50 p-4 rounded-md">
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
                        Header Row (0-based)
                      </label>
                      <input
                        type="number"
                        min="0"
                        className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                        value={headerRow}
                        onChange={(e) => setHeaderRow(e.target.value)}
                      />
                    </div>
                    
                    <div className="md:col-span-2">
                      <button
                        onClick={handlePreview}
                        disabled={loadingPreview || !selectedSheet}
                        className="inline-flex items-center px-4 py-2 bg-gray-700 text-white rounded-md hover:bg-gray-800 disabled:opacity-50 disabled:cursor-not-allowed"
                      >
                        {loadingPreview && <LoadingSpinner />}
                        {loadingPreview ? 'Analyzing Columns...' : 'Preview Columns'}
                      </button>
                    </div>
                  </div>
                )}

                {columns.length > 0 && (
                  <div className="bg-gray-50 p-4 rounded-md">
                    <h4 className="font-medium text-gray-900 mb-3">Column Mapping</h4>
                    <div className="grid md:grid-cols-2 gap-4 mb-4">
                      {[
                        ["Executive Name Column", "exec_name_col", true],
                        ["Executive Code Column", "exec_code_col", false],
                        ["Branch Column", "branch_col", true],
                        ["Region Column", "region_col", true],
                      ].map(([label, key, required]) => (
                        <div key={key}>
                          <label className="block text-sm font-medium text-gray-700 mb-1">
                            {label} {required && <span className="text-red-500">*</span>}
                          </label>
                          <select
                            className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                            value={mapping[key]}
                            onChange={(e) =>
                              setMapping({ ...mapping, [key]: e.target.value })
                            }
                          >
                            <option value="">-- Select Column --</option>
                            {columns.map((col) => (
                              <option key={col} value={col}>{col}</option>
                            ))}
                          </select>
                        </div>
                      ))}
                    </div>
                    
                    <button
                      onClick={handleUpload}
                      disabled={loadingProcess || !mapping.exec_name_col || !mapping.branch_col || !mapping.region_col}
                      className="inline-flex items-center px-6 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      {loadingProcess && <LoadingSpinner />}
                      {loadingProcess ? 'Processing File...' : 'Process File'}
                    </button>
                    
                    {loadingProcess && (
                      <div className="mt-3 text-sm text-gray-600">
                        Please wait while we process your file. This may take a few moments...
                      </div>
                    )}
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
});

// Add display name for easier debugging
BranchRegionManager.displayName = "BranchRegionManager";

export default BranchRegionManager;
