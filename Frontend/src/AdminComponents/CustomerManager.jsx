import React, { useEffect, useState } from "react";
import api from "../api/axios";
import * as XLSX from "xlsx";

const CustomerManager = () => {
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

  // ✅ processing state
  const [processing, setProcessing] = useState(false);

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

    const reader = new FileReader();
    reader.onload = (evt) => {
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
      }

      alert("✅ File uploaded successfully, ready to process!");
    };
    reader.readAsArrayBuffer(file);
  };

  const handleProcessFile = async () => {
    if (!sheetData.length) {
      alert("⚠️ Please upload a file first!");
      return;
    }

    try {
      setProcessing(true);
      alert("⏳ Please wait, it may take some time...");

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

      alert("✅ Bulk assignment complete!");
    } catch (err) {
      alert("❌ Error while processing file!");
      console.error(err);
    } finally {
      setProcessing(false);
    }
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
    await api.post("/remove-customer", {
      executive: selectedExec,
      customers: codes,
    });
    fetchCustomers();
    fetchUnmapped();
  };

  const handleAssign = async (codes) => {
    await api.post("/assign-customer", {
      executive: selectedExec,
      customers: codes,
    });
    fetchCustomers();
    fetchUnmapped();
  };

  const handleAddNew = async () => {
    const codes = newCodes
      .split("\n")
      .map((code) => code.trim())
      .filter(Boolean);
    if (codes.length > 0) {
      await api.post("/assign-customer", {
        executive: selectedExec,
        customers: codes,
      });
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
        <input
          type="file"
          accept=".xlsx"
          onChange={handleExcelUpload}
          className="mb-2"
        />

        {sheets.length > 0 && sheetData.length > 0 && (
          <>
            <label className="block">Executive Name Column:</label>
            <select
              value={execNameCol}
              onChange={(e) => setExecNameCol(e.target.value)}
              className="border p-2 mb-2"
            >
              {Object.keys(sheetData[0]).map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>

            <label className="block">Executive Code Column:</label>
            <select
              value={execCodeCol}
              onChange={(e) => setExecCodeCol(e.target.value)}
              className="border p-2 mb-2"
            >
              <option value="">None</option>
              {Object.keys(sheetData[0]).map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>

            <label className="block">Customer Code Column:</label>
            <select
              value={custCodeCol}
              onChange={(e) => setCustCodeCol(e.target.value)}
              className="border p-2 mb-2"
            >
              {Object.keys(sheetData[0]).map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>

            <label className="block">Customer Name Column:</label>
            <select
              value={custNameCol}
              onChange={(e) => setCustNameCol(e.target.value)}
              className="border p-2 mb-2"
            >
              <option value="">None</option>
              {Object.keys(sheetData[0]).map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>

            <button
              onClick={handleProcessFile}
              className="bg-green-600 text-white px-4 py-2 rounded mt-2"
              disabled={processing}
            >
              {processing ? "⏳ Processing..." : "Process File"}
            </button>
          </>
        )}
      </div>

      {/* ---------------- MANUAL CUSTOMER MANAGEMENT ---------------- */}
      <div className="p-4 bg-gray-50 rounded">
        <h3 className="font-semibold mb-2">Manual Management</h3>

        <label className="block mb-2">Select Executive:</label>
        <select
          value={selectedExec}
          onChange={(e) => setSelectedExec(e.target.value)}
          className="border p-2 mb-4"
        >
          {execs.map((exec) => (
            <option key={exec.id} value={exec.name}>
              {exec.name}
            </option>
          ))}
        </select>

        <h4 className="font-medium">Assigned Customers</h4>
        <ul className="border p-2 mb-4 max-h-40 overflow-y-auto">
          {assignedCustomers.map((c) => (
            <li key={c} className="flex items-center">
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
                className="mr-2"
              />
              {c}
            </li>
          ))}
        </ul>
        <button
          onClick={() => handleRemove(selectedToRemove)}
          className="bg-red-600 text-white px-3 py-1 rounded mb-4"
        >
          Remove Selected
        </button>

        <h4 className="font-medium">Unmapped Customers</h4>
        <ul className="border p-2 mb-4 max-h-40 overflow-y-auto">
          {unmappedCustomers.map((c) => (
            <li key={c} className="flex items-center">
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
                className="mr-2"
              />
              {c}
            </li>
          ))}
        </ul>
        <button
          onClick={() => handleAssign(selectedToAssign)}
          className="bg-blue-600 text-white px-3 py-1 rounded mb-4"
        >
          Assign Selected
        </button>

        <h4 className="font-medium">Add New Customers</h4>
        <textarea
          value={newCodes}
          onChange={(e) => setNewCodes(e.target.value)}
          placeholder="Enter customer codes (one per line)"
          className="border p-2 w-full mb-2"
        />
        <button
          onClick={handleAddNew}
          className="bg-green-600 text-white px-3 py-1 rounded"
        >
          Add
        </button>
      </div>
    </div>
  );
};

export default CustomerManager;
