import React, { useEffect, useState } from "react";
import api from "../api/axios";
import * as XLSX from "xlsx";

const CustomerManager = () => {
  const [execs, setExecs] = useState([]);
  const [selectedExec, setSelectedExec] = useState("");
  const [assignedCustomers, setAssignedCustomers] = useState([]);
  const [unmappedCustomers, setUnmappedCustomers] = useState([]);
  const [newCodes, setNewCodes] = useState("");

  const [excelFile, setExcelFile] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [sheetData, setSheetData] = useState([]);
  const [execNameCol, setExecNameCol] = useState("");
  const [execCodeCol, setExecCodeCol] = useState("");
  const [custCodeCol, setCustCodeCol] = useState("");
  const [custNameCol, setCustNameCol] = useState("");
  const [selectedToRemove, setSelectedToRemove] = useState([]);
  const [selectedToAssign, setSelectedToAssign] = useState([]);

  const guessColumn = (headers, type) => {
    const aliases = {
      executive_name: ["executive name", "empname", "executive"],
      executive_code: ["executive code", "empcode", "ecode"],
      customer_code: ["customer code", "slcode", "custcode"],
      customer_name: ["customer name", "slname", "custname"],
    };

    const candidates = aliases[type] || [type];
    const lowerHeaders = headers.map(h => h.toLowerCase());

    for (let alias of candidates) {
      const match = lowerHeaders.find(h => h === alias);
      if (match) return headers[lowerHeaders.indexOf(match)];
    }
    for (let alias of candidates) {
      const match = lowerHeaders.find(h => h.includes(alias));
      if (match) return headers[lowerHeaders.indexOf(match)];
    }
    return "";
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setExcelFile(file);

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.SheetNames[0];
      setSheets(workbook.SheetNames);
      setSelectedSheet(sheet);

      const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], { defval: "" });
      setSheetData(jsonData);

      if (jsonData.length > 0) {
        const headers = Object.keys(jsonData[0]);
        setExecNameCol(guessColumn(headers, "executive_name"));
        setExecCodeCol(guessColumn(headers, "executive_code"));
        setCustCodeCol(guessColumn(headers, "customer_code"));
        setCustNameCol(guessColumn(headers, "customer_name"));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleProcessFile = async () => {
    const payload = {
      data: sheetData,
      execNameCol,
      execCodeCol,
      custCodeCol,
      custNameCol
    };
    await api.post("/bulk-assign-customers", payload);
    fetchCustomers();
    fetchUnmapped();
    alert("Bulk assignment complete!");
  };

  const fetchExecs = async () => {
    const res = await api.get("/executives");
    setExecs(res.data);
    if (res.data.length > 0) setSelectedExec(res.data[0].name);
  };

  const fetchCustomers = async () => {
    const res = await api.get(`/customers?executive=${selectedExec}`);
    setAssignedCustomers(res.data);
  };

  const fetchUnmapped = async () => {
    const res = await api.get("/customers/unmapped");
    setUnmappedCustomers(res.data);
  };

  const handleRemove = async (codes) => {
    await api.post("/remove-customer", { executive: selectedExec, customers: codes });
    fetchCustomers();
    fetchUnmapped();
  };

  const handleAssign = async (codes) => {
    await api.post("/assign-customer", { executive: selectedExec, customers: codes });
    fetchCustomers();
    fetchUnmapped();
  };

  const handleAddNew = async () => {
    const codes = newCodes.split("\\n").map(code => code.trim()).filter(Boolean);
    if (codes.length > 0) {
      await api.post("/assign-customer", { executive: selectedExec, customers: codes });
      setNewCodes("");
      fetchCustomers();
      fetchUnmapped();
    }
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

    {/* ---------------- BULK UPLOAD ---------------- */}
    <div className="mb-6 p-4 bg-blue-50 rounded">
      <h3 className="font-semibold mb-2">Bulk Assignment via Excel</h3>
      <input type="file" accept=".xlsx" onChange={handleExcelUpload} className="mb-2" />
      {sheets.length > 0 && sheetData.length > 0 && (
        <>
          <label className="block">Executive Name Column:</label>
          <select value={execNameCol} onChange={(e) => setExecNameCol(e.target.value)} className="border p-2 mb-2">
            {Object.keys(sheetData[0]).map(col => <option key={col} value={col}>{col}</option>)}
          </select>

          <label className="block">Executive Code Column:</label>
          <select value={execCodeCol} onChange={(e) => setExecCodeCol(e.target.value)} className="border p-2 mb-2">
            <option value="">None</option>
            {Object.keys(sheetData[0]).map(col => <option key={col} value={col}>{col}</option>)}
          </select>

          <label className="block">Customer Code Column:</label>
          <select value={custCodeCol} onChange={(e) => setCustCodeCol(e.target.value)} className="border p-2 mb-2">
            {Object.keys(sheetData[0]).map(col => <option key={col} value={col}>{col}</option>)}
          </select>

          <label className="block">Customer Name Column:</label>
          <select value={custNameCol} onChange={(e) => setCustNameCol(e.target.value)} className="border p-2 mb-2">
            <option value="">None</option>
            {Object.keys(sheetData[0]).map(col => <option key={col} value={col}>{col}</option>)}
          </select>

          <button onClick={handleProcessFile} className="bg-green-600 text-white px-4 py-2 rounded mt-2">
            Process File
          </button>

          <h4 className="mt-4 font-medium">Preview:</h4>
          <table className="w-full border mt-2">
            <thead>
              <tr>
                {sheetData[0] && Object.keys(sheetData[0]).map(col => (
                  <th key={col} className="border px-2">{col}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sheetData.slice(0, 5).map((row, idx) => (
                <tr key={idx}>
                  {Object.values(row).map((val, j) => (
                    <td key={j} className="border px-2">{val}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}
    </div>

    {/* ---------------- MANUAL CUSTOMER MANAGEMENT ---------------- */}
    {execs.length > 0 && (
      <div className="mt-8 p-4 bg-white rounded shadow">
        <h3 className="text-lg font-semibold mb-2">Manual Customer Management</h3>

        <label className="block mb-1">Select Executive:</label>
        <select
          value={selectedExec}
          onChange={(e) => setSelectedExec(e.target.value)}
          className="border px-3 py-2 mb-4 rounded"
        >
          {execs.map((e) => (
            <option key={e.id} value={e.name}>{e.name}</option>
          ))}
        </select>

        <div className="grid md:grid-cols-2 gap-6">
          {/* Assigned Customers */}
          <div>
            <h4 className="font-medium mb-2">Assigned Customers ({assignedCustomers.length})</h4>
            {assignedCustomers.length > 0 ? (
              <>
                <ul className="max-h-60 overflow-y-scroll bg-gray-100 p-2 rounded text-sm">
                  {assignedCustomers.slice(0, 20).map(c => (
                    <li key={c.code}>{c.code} - {c.name}</li>
                  ))}
                </ul>
                {assignedCustomers.length > 20 && (
                  <p className="text-xs mt-1">Showing first 20 of {assignedCustomers.length}</p>
                )}
              </>
            ) : (
              <p>No customers assigned</p>
            )}
          </div>

          {/* Actions */}
          <div>
            <h4 className="font-medium mb-2">Actions</h4>

            {/* Remove Selected */}
{assignedCustomers.length > 0 && (
  <div className="mb-4">
    <label className="block font-medium mb-1">Remove Customers:</label>
    <div className="max-h-32 overflow-y-auto border p-2 rounded flex flex-col gap-1">
      {assignedCustomers.map((c) => (
        <label key={c.code} className="flex items-center gap-2 text-sm">
          <input
            type="checkbox"
            value={c.code}
            checked={selectedToRemove.includes(c.code)}
            onChange={e => {
              const checked = e.target.checked;
              if (checked) {
                setSelectedToRemove([...selectedToRemove, c.code]);
              } else {
                setSelectedToRemove(selectedToRemove.filter(code => code !== c.code));
              }
            }}
          />
          <span>{c.code} - {c.name}</span>
        </label>
      ))}
    </div>
    <button
      onClick={() => handleRemove(selectedToRemove)}
      className="bg-red-600 text-white px-4 py-2 rounded mt-2"
      disabled={selectedToRemove.length === 0}
    >
      Remove Selected
    </button>
  </div>
)}


            {/* Assign Selected */}
{unmappedCustomers.length > 0 && (
  <div className="mb-4">
    <label className="block font-medium mb-1">Assign Customers:</label>
    <select
      multiple
      className="w-full border p-2 h-32"
      value={selectedToAssign}
      onChange={(e) =>
        setSelectedToAssign(
          Array.from(e.target.selectedOptions, (option) => option.value)
        )
      }
    >
      {unmappedCustomers.map((c) => (
        <option key={c.code} value={c.code}>
          {c.code} - {c.name}
        </option>
      ))}
    </select>
    <button
      onClick={() => handleAssign(selectedToAssign)}
      className="bg-blue-600 text-white px-4 py-2 rounded mt-2"
      disabled={selectedToAssign.length === 0}
    >
      Assign Selected
    </button>
  </div>
)}


            {/* Add New Codes */}
            <label className="block mb-1">Add New Customer Codes (one per line):</label>
            <textarea
              rows={4}
              className="w-full border rounded p-2 mb-2"
              value={newCodes}
              onChange={(e) => setNewCodes(e.target.value)}
            />
            <button
              onClick={handleAddNew}
              className="bg-green-600 text-white px-4 py-2 rounded"
            >
              Add New Codes
            </button>
          </div>
        </div>

        {/* Unmapped Display */}
        {unmappedCustomers.length > 0 && (
          <div className="mt-6">
            <h4 className="font-medium mb-2">Unmapped Customers (showing up to 50)</h4>
            <ul className="bg-yellow-50 p-2 rounded max-h-60 overflow-y-scroll text-sm">
              {unmappedCustomers.slice(0, 50).map(c => (
                <li key={c.code}>{c.code} - {c.name}</li>
              ))}
            </ul>
          </div>
        )}
      </div>
    )}
  </div>
);
}
export default CustomerManager;