import React, { useEffect, useState } from "react";
import api from "../api/axios";

const SavedFiles = () => {
  const [budgetFiles, setBudgetFiles] = useState([]);
  const [salesFiles, setSalesFiles] = useState([]);
  const [osFiles, setOsFiles] = useState([]);

  useEffect(() => {
    fetchAll();
  }, []);

  const fetchAll = async () => {
    try {
      const [bRes, sRes, oRes] = await Promise.all([
        api.get("/budget-files"),
        api.get("/sales-files"),
        api.get("/os-files"),
      ]);
      setBudgetFiles(bRes.data);
      setSalesFiles(sRes.data);
      setOsFiles(oRes.data);
    } catch (err) {
      console.error("Fetch error", err);
    }
  };
  const handleDelete = async (prefix, id) => {
    const confirmDelete = window.confirm("Are you sure you want to delete this file?");
    if (!confirmDelete) return;

    try {
      await api.delete(`/${prefix}-files/${id}`);
      fetchAll(); // refresh list
    } catch (err) {
      alert("Failed to delete file");
      console.error(err);
    }
  };

  const renderTable = (files, label, prefix) => (
    <div className="mb-8">
      <h2 className="text-lg font-semibold mb-2">{label}</h2>
      {files.length === 0 ? (
        <p className="text-gray-500">No files found.</p>
      ) : (
        <table className="min-w-full border text-sm">
          <thead className="bg-gray-100">
            <tr>
              <th className="border px-4 py-2">Filename</th>
              <th className="border px-4 py-2">Uploaded At</th>
              <th className="border px-4 py-2">Action</th>
            </tr>
          </thead>
          <tbody>
            {files.map((file) => (
              <tr key={file.id}>
                <td className="border px-4 py-2">{file.filename}</td>
                <td className="border px-4 py-2">
                  {new Date(file.uploaded_at).toLocaleString()}
                </td>
                <td className="border px-4 py-2">
                  <a
                    href={`http://localhost:5000/api/${prefix}-files/${file.id}/download`}
                    className="text-blue-600 hover:underline"
                    download
                  >
                    Download
                  </a>
                  <button
                    onClick={() => handleDelete(prefix, file.id)}
                    className="ml-4 text-red-600 hover:underline"
                  >
                    Delete
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );

  return (
    <div className="p-6 bg-white rounded shadow">
      {renderTable(budgetFiles, "Saved Budget Files", "budget")}
      {renderTable(salesFiles, "Saved Sales Files", "sales")}
      {renderTable(osFiles, "Saved OS Files", "os")}
    </div>
  );
};

export default SavedFiles;
