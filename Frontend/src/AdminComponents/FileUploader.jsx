import React, { useState } from "react";
import api from "../api/axios";

const FileUploader = ({ label, endpoint }) => {
  const [file, setFile] = useState(null);
  const [sheet, setSheet] = useState("");
  const [header, setHeader] = useState(0);
  const [status, setStatus] = useState("");

  const handleUpload = async () => {
    if (!file) {
      setStatus("Please select a file.");
      return;
    }

    const formData = new FormData();
    formData.append("file", file);
    formData.append("sheet", sheet);
    formData.append("header", header);

    try {
      const res = await api.post(endpoint, formData, {
        responseType: "blob",
      });

      const blob = new Blob([res.data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const downloadUrl = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = downloadUrl;
      a.download = `processed_${file.name}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setStatus("File processed and downloaded!");
    } catch (err) {
      console.error(err);
      setStatus("Upload failed.");
    }
  };

  return (
    <div className="bg-white shadow-md rounded p-6 my-6">
      <h2 className="text-xl font-semibold mb-4">{label}</h2>

      <input
        type="file"
        accept=".xlsx,.xls"
        onChange={(e) => setFile(e.target.files[0])}
        className="mb-4"
      />

      <input
        type="text"
        placeholder="Sheet Name (optional)"
        value={sheet}
        onChange={(e) => setSheet(e.target.value)}
        className="border rounded p-2 mb-4 w-full"
      />

      <input
        type="number"
        placeholder="Header Row Index (default 0)"
        value={header}
        onChange={(e) => setHeader(e.target.value)}
        className="border rounded p-2 mb-4 w-full"
      />

      <button
        className="bg-blue-700 text-white px-4 py-2 rounded hover:bg-blue-800"
        onClick={handleUpload}
      >
        Upload & Process
      </button>

      {status && (
        <p className="mt-4 text-sm text-green-600">{status}</p>
      )}
    </div>
  );
};

export default FileUploader;
