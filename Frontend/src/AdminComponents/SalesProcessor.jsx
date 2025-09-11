import React, { useState } from "react";
import api from "../api/axios";

const SalesProcessor = () => {
  const [file, setFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headerRow, setHeaderRow] = useState(1);

  const [columns, setColumns] = useState([]);
  const [preview, setPreview] = useState([]);
//   const [processedExcel, setProcessedExcel] = useState(null);
  const [downloadLink, setDownloadLink] = useState(null);
  const [processedExcelBase64, setProcessedExcelBase64] = useState(null);
  const [customFilename, setCustomFilename] = useState("");
  const [loading, setLoading] = useState(false); 

  


  // Column mappings
  const [execCodeCol, setExecCodeCol] = useState("");
  const [execNameCol, setExecNameCol] = useState("");
  const [productCol, setProductCol] = useState("");
  const [unitCol, setUnitCol] = useState("");
  const [quantityCol, setQuantityCol] = useState("");
  const [valueCol, setValueCol] = useState("");

  const handleFileUpload = async (e) => {
    const selectedFile = e.target.files[0];
    setFile(selectedFile);

    const formData = new FormData();
    formData.append("file", selectedFile);
    const res = await api.post("/upload-tools/sheet-names", formData);
    const names = res.data.sheet_names;
    setSheetNames(names);
    setSelectedSheet(names.find(n => n.toLowerCase().includes("sales")) || names[0]);
  };

  const handlePreview = async () => {
    setLoading(true);
    try {
      if (!file || !selectedSheet) return;

      const formData = new FormData();
      formData.append("file", file);
      formData.append("sheet_name", selectedSheet);
      formData.append("header_row", headerRow);

      const res = await api.post("/upload-tools/preview", formData);
      setColumns(res.data.columns);
      setPreview(res.data.preview);
      setColumns(res.data.columns);

      setExecCodeCol(autoDetectColumn(res.data.columns, "executive code"));
      setExecNameCol(autoDetectColumn(res.data.columns, "executive name"));
      setProductCol(autoDetectColumn(res.data.columns, "type (make)"));
      setUnitCol(autoDetectColumn(res.data.columns, "uom"));
      setQuantityCol(autoDetectColumn(res.data.columns, "quantity"));
      setValueCol(autoDetectColumn(res.data.columns, "value"));
    }catch(err){
      console.log(err);
    }finally{
      setLoading(false);
    }

  };

  

  const handleProcess = async () => {
    setLoading(true);
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
    setPreview(res.data.preview);

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
  }catch(err){
    console.log(err);
    }finally{
      setLoading(false);
  }
  };
  function autoDetectColumn(columns, target, fallback = "") {
  const lowerTarget = target.toLowerCase();
  return (
    columns.find((col) => col.toLowerCase().includes(lowerTarget)) ||
    fallback
  );
}

  return (
    <div>
      <h2 className="text-xl font-semibold mb-4">Sales File Processing</h2>

      <input type="file" accept=".xls,.xlsx" onChange={handleFileUpload} />

      {sheetNames.length > 0 && (
        <div className="my-4">
          <label>Sheet:</label>
          <select
            className="border p-1 ml-2"
            value={selectedSheet}
            onChange={(e) => setSelectedSheet(e.target.value)}
          >
            {sheetNames.map((sheet) => (
              <option key={sheet} value={sheet}>
                {sheet}
              </option>
            ))}
          </select>

          <label className="ml-4">Header Row:</label>
          <input
            type="number"
            min={0}
            className="border p-1 w-16 ml-2"
            value={headerRow}
            onChange={(e) => setHeaderRow(parseInt(e.target.value))}
          />
          <button
            onClick={handlePreview}
            className="ml-4 bg-blue-600 text-white px-4 py-1 rounded"
          >
            Load Sheet
          </button>
        </div>
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
        <div className="grid grid-cols-2 gap-4 mb-4">
          <div>
            <label>Executive Code Column:</label>
            <select
              value={execCodeCol}
              onChange={(e) => setExecCodeCol(e.target.value)}
              className="block w-full border p-1"
            >
              {columns.map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>

            <label>Executive Name Column:</label>
            <select
              value={execNameCol}
              onChange={(e) => setExecNameCol(e.target.value)}
              className="block w-full border p-1"
            >
              <option value="">-- None --</option>
              {columns.map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>

            <label>Product Column:</label>
            <select
              value={productCol}
              onChange={(e) => setProductCol(e.target.value)}
              className="block w-full border p-1"
            >
              <option value="">-- None --</option>
              {columns.map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label>Unit Column:</label>
            <select
              value={unitCol}
              onChange={(e) => setUnitCol(e.target.value)}
              className="block w-full border p-1"
            >
              <option value="">-- None --</option>
              {columns.map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>

            <label>Quantity Column:</label>
            <select
              value={quantityCol}
              onChange={(e) => setQuantityCol(e.target.value)}
              className="block w-full border p-1"
            >
              <option value="">-- None --</option>
              {columns.map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>

            <label>Value Column:</label>
            <select
              value={valueCol}
              onChange={(e) => setValueCol(e.target.value)}
              className="block w-full border p-1"
            >
              <option value="">-- None --</option>
              {columns.map((col) => (
                <option key={col} value={col}>
                  {col}
                </option>
              ))}
            </select>
          </div>
        </div>
      )}

      {columns.length > 0 && (
        <button
          onClick={handleProcess}
          className="bg-green-600 text-white px-4 py-2 rounded"
        >
          Process Sales File
        </button>
      )}

      {/* {preview.length > 0 && (
        <div className="mt-6">
          <h3 className="text-lg font-semibold mb-2">Preview</h3>
          <table className="w-full table-auto border">
            <thead>
              <tr>
                {Object.keys(preview[0]).map((col) => (
                  <th className="border px-2 py-1" key={col}>
                    {col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {preview.map((row, i) => (
                <tr key={i}>
                  {Object.values(row).map((val, idx) => (
                    <td className="border px-2 py-1" key={idx}>
                      {val}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )} */}

      {downloadLink && (
        <div className="mt-4">
          <a
            href={downloadLink}
            download="processed_sales.xlsx"
            className="bg-blue-600 text-white px-4 py-2 rounded inline-block"
          >
            Download Processed Sales File
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
    placeholder="e.g., Apr-2025_Processed_Sales"
  />
</div>

      {processedExcelBase64 && (
  <button
    className="bg-indigo-600 text-white px-4 py-2 rounded mt-4 ml-4"
    onClick={async () => {
      const res = await api.post("/save-sales-file", {
        file_data: processedExcelBase64,
        filename: customFilename?.trim() ||"Processed_Sales.xlsx",
      });
      alert("✅ File saved to DB. File ID: " + res.data.id);
    }}
  >
    Save Sales File to DB
  </button>
)}

    </div>
  );
};

export default SalesProcessor;
