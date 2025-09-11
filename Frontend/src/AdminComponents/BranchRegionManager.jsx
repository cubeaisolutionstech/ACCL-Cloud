// frontend/components/BranchRegionManager.jsx

import React, { useEffect, useState } from "react";
import api from "../api/axios";

const BranchRegionManager = () => {
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

  const fetchAll = async () => {
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
  };

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

  // Upload tab state
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

  const [loadingSheets, setLoadingSheets] = useState(false); // Loader for Load Sheets
    const [loadingPreview, setLoadingPreview] = useState(false); // Loader for Preview Columns
    const [loadingProcess, setLoadingProcess] = useState(false); // Loader for Process File

  const handleSheetLoad = async () => {
    setLoadingSheets(true);
    try{
  const formData = new FormData();
  formData.append("file", uploadFile);
  const res = await api.post("/get-sheet-names", formData);

  const sheetList = res.data.sheets || [];
  setSheetNames(sheetList);

  // ✅ Auto-select first sheet
  if (sheetList.length > 0) {
    setSelectedSheet(sheetList[0]);
  }
}catch (err){
  console.error("Error loading sheets", err);
}finally{
  setLoadingSheets(false);
}
};


   const handlePreview = async () => {
    setLoadingPreview(true);
    try {
    const formData = new FormData();
    formData.append("file", uploadFile);
    formData.append("sheet_name", selectedSheet);
    formData.append("header_row", headerRow);
    const res = await api.post("/preview-excel", formData);
    const cols = res.data.columns;
    setColumns(cols);
    setMapping(autoMapColumns(cols)); // 👈 auto-mapping here
    }catch (err){
      console.error("Error previewing columns", err);
    } finally {
      setLoadingPreview(false);
    }
  };

  const handleUpload = async () => {
    setLoadingProcess(true);
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
    alert("✅ File processed!");
    fetchAll();
    } catch (err) {
      console.error("Error processing file", err);
    } finally {
      setLoadingProcess(false);
    }
  };

  useEffect(() => {
  const fetchBranchExecs = async () => {
    if (selectedBranch) {
      const res = await api.get(`/branch/${selectedBranch}/executives`);
      setBranchExecs(res.data);  // already selected ones
    }
  };
  fetchBranchExecs();
}, [selectedBranch]);

// when a region is selected — get current branches
useEffect(() => {
  const fetchRegionBranches = async () => {
    if (selectedRegion) {
      const res = await api.get(`/region/${selectedRegion}/branches`);
      setRegionBranches(res.data); // already selected ones
    }
  };
  fetchRegionBranches();
}, [selectedRegion]);

  return (
    <div className="p-6">
      <h2 className="text-xl font-bold mb-4">Branch & Region Mapping</h2>

      {/* Tabs */}
      <div className="flex mb-6">
        <button
          className={`mr-4 px-4 py-2 rounded-t ${
            activeTab === "manual" ? "bg-blue-600 text-white" : "bg-gray-200"
          }`}
          onClick={() => setActiveTab("manual")}
        >
          Manual Entry
        </button>
        <button
          className={`px-4 py-2 rounded-t ${
            activeTab === "upload" ? "bg-blue-600 text-white" : "bg-gray-200"
          }`}
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
    🗑️ Delete
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
    🗑️ Delete
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
                  <td className="border px-2 py-1">{m.executives.join(", ")}</td>
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
          <h3 className="font-semibold mb-2">Upload Branch-Region Mapping File</h3>
          <input type="file" className="mb-2" onChange={(e) => setUploadFile(e.target.files[0])} />
          {uploadFile && (
            <>
              <button onClick={handleSheetLoad} className="bg-blue-500 text-white px-3 py-1 rounded mb-3">Load Sheets</button>
              <div className="grid md:grid-cols-2 gap-4">
                <div>
                  <label>Sheet Name</label>
                  <select className="w-full border px-2 py-1" value={selectedSheet} onChange={(e) => setSelectedSheet(e.target.value)}>
                    <option value="">-- Select --</option>
                    {sheetNames.map((sheet) => (
                      <option key={sheet} value={sheet}>{sheet}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label>Header Row</label>
                  <input type="number" className="w-full border px-2 py-1" value={headerRow} onChange={(e) => setHeaderRow(e.target.value)} />
                </div>
              </div>
              <button onClick={handlePreview} className="mt-4 bg-gray-700 text-white px-4 py-2 rounded">Preview Columns</button>
            </>
          )}
          {columns.length > 0 && (
            <div className="mt-6 grid md:grid-cols-2 gap-4">
              {[
                ["Executive Name Column", "exec_name_col"],
                ["Executive Code Column", "exec_code_col"],
                ["Branch Column", "branch_col"],
                ["Region Column", "region_col"],
              ].map(([label, key]) => (
                <div key={key}>
                  <label>{label}</label>
                  <select className="w-full border px-2 py-1"
                    value={mapping[key]}
                    onChange={(e) => setMapping({ ...mapping, [key]: e.target.value })}
                  >
                    <option value="">-- Select --</option>
                    {columns.map((col) => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>
              ))}
              <div className="col-span-2">
                <button onClick={handleUpload} className="mt-4 bg-green-600 text-white px-6 py-2 rounded">
                  Process File
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default BranchRegionManager;
