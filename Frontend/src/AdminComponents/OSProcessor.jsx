import React, { useState } from "react";
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
  const [loading, setLoading] = useState(false); 

  const handleFileUpload = (e) => {
    const uploaded = e.target.files[0];
    setFile(uploaded);

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
      setColumns(headers);
      setExecCodeCol(headers.find(h => h.toLowerCase().includes("executive code")) || "");
    };
    reader.readAsBinaryString(uploaded);
  };

  const handleProcess = async () => {
    setLoading(true);
    if (!file || !selectedSheet || !execCodeCol) return alert("Fill all fields");

    const formData = new FormData();
    formData.append("file", file);
    formData.append("sheet_name", selectedSheet);
    formData.append("header_row", headerRow);
    formData.append("exec_code_col", execCodeCol);

    try {
      const res = await api.post("/process-os-file", formData, {
        responseType: "blob",
      });
      setProcessedFile(URL.createObjectURL(res.data));
    } catch (err) {
      console.error(err);
      alert("Processing failed");
    }finally{
      setLoading(false);
    }
  };
  const handleSaveToDb = async () => {
    setLoading(true);
  if (!file || !customFilename) return alert("Please select a file and enter a name");
  const formData = new FormData();
  formData.append("file", file);
  formData.append("filename", customFilename);
  try {
    await api.post("/save-os-file", formData);
    alert("OS File saved to database.");
  } catch (err) {
    console.error(err);
    alert("Failed to save file.");
  }finally{
    setLoading(false);
  }
};


  return (
    <div className="bg-white shadow p-6 rounded">
      <h2 className="text-xl font-bold mb-4">Process OS File</h2>

      <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="mb-4" />

      {sheetNames.length > 0 && (
        <>
          <div className="mb-4">
            <label className="block font-medium">Sheet Name:</label>
            <select
              className="border p-2 rounded w-full"
              value={selectedSheet}
              onChange={(e) => setSelectedSheet(e.target.value)}
            >
              {sheetNames.map((name, idx) => (
                <option key={idx} value={name}>{name}</option>
              ))}
            </select>
          </div>

          <div className="mb-4">
            <label className="block font-medium">Header Row (0-based):</label>
            <input
              type="number"
              className="border p-2 rounded w-full"
              value={headerRow}
              onChange={(e) => setHeaderRow(Number(e.target.value))}
            />
          </div>

          <div className="mb-4">
            <label className="block font-medium">Executive Code Column:</label>
            <select
              className="border p-2 rounded w-full"
              value={execCodeCol}
              onChange={(e) => setExecCodeCol(e.target.value)}
            >
              <option value="">-- Select Column --</option>
              {columns.map((col, idx) => (
                <option key={idx} value={col}>{col}</option>
              ))}
            </select>
          </div>

          {loading && (
        <div className="text-blue-600 font-medium my-2 flex items-center gap-2">
          <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"></path>
          </svg>
          Loading...
        </div>
      )}

          <div className="mb-4">
            <button onClick={handleProcess} className="bg-blue-600 text-white px-4 py-2 rounded">
              Process OS File
            </button>
          </div>
        </>
      )}

      {/* {preview.length > 0 && (
        <div className="mb-4">
          <h3 className="font-semibold mb-2">Preview</h3>
          <table className="table-auto w-full border">
            <thead>
              <tr>
                {Object.keys(preview[0]).map((col, i) => (
                  <th key={i} className="border px-2 py-1 text-sm">{col}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {preview.map((row, i) => (
                <tr key={i}>
                  {Object.values(row).map((val, j) => (
                    <td key={j} className="border px-2 py-1 text-xs">{val}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )} */}

      {processedFile && (
        <div className="mt-4">
          <a
            href={processedFile}
            download="processed_os.xlsx"
            className="bg-green-600 text-white px-4 py-2 rounded"
          >
            Download Processed File
          </a>
          
          <input
  type="text"
  placeholder="Enter file name"
  value={customFilename}
  onChange={(e) => setCustomFilename(e.target.value)}
  className="border p-2 rounded w-full mb-4"
/>
          <button
      className="bg-yellow-600 text-white px-4 py-2 rounded"
      onClick={handleSaveToDb}
    >
      Save to Database
    </button>
        </div>
      )}
    </div>
    
  );
};

export default OSProcessor;
