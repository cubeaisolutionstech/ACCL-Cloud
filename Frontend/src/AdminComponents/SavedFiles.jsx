import React, { useEffect, useState } from "react";
import api from "../api/axios";

const SavedFiles = () => {
  const [budgetFiles, setBudgetFiles] = useState([]);
  const [salesFiles, setSalesFiles] = useState([]);
  const [osFiles, setOsFiles] = useState([]);
  
  // Search states
  const [budgetSearch, setBudgetSearch] = useState({ fromDate: '', toDate: '' });
  const [salesSearch, setSalesSearch] = useState({ fromDate: '', toDate: '' });
  const [osSearch, setOsSearch] = useState({ fromDate: '', toDate: '' });

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

  // Filter files based on date range
  const filterFilesByDate = (files, searchCriteria) => {
    if (!searchCriteria.fromDate && !searchCriteria.toDate) {
      return files;
    }

    return files.filter(file => {
      const fileDate = new Date(file.uploaded_at);
      const fromDate = searchCriteria.fromDate ? new Date(searchCriteria.fromDate) : null;
      const toDate = searchCriteria.toDate ? new Date(searchCriteria.toDate + 'T23:59:59') : null;

      if (fromDate && toDate) {
        return fileDate >= fromDate && fileDate <= toDate;
      } else if (fromDate) {
        return fileDate >= fromDate;
      } else if (toDate) {
        return fileDate <= toDate;
      }
      return true;
    });
  };

  const handleSearchChange = (type, field, value) => {
    switch (type) {
      case 'budget':
        setBudgetSearch(prev => ({ ...prev, [field]: value }));
        break;
      case 'sales':
        setSalesSearch(prev => ({ ...prev, [field]: value }));
        break;
      case 'os':
        setOsSearch(prev => ({ ...prev, [field]: value }));
        break;
    }
  };

  const clearSearch = (type) => {
    switch (type) {
      case 'budget':
        setBudgetSearch({ fromDate: '', toDate: '' });
        break;
      case 'sales':
        setSalesSearch({ fromDate: '', toDate: '' });
        break;
      case 'os':
        setOsSearch({ fromDate: '', toDate: '' });
        break;
    }
  };

  const renderSearchBar = (type, searchCriteria) => (
    <div className="mb-4 p-4 bg-gray-50 rounded-lg border">
      <div className="flex flex-wrap items-center gap-4">
        <div className="flex items-center gap-2">
          <label className="text-sm font-medium text-gray-700">From:</label>
          <input
            type="date"
            value={searchCriteria.fromDate}
            onChange={(e) => handleSearchChange(type, 'fromDate', e.target.value)}
            className="px-3 py-1 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
          />
        </div>
        <div className="flex items-center gap-2">
          <label className="text-sm font-medium text-gray-700">To:</label>
          <input
            type="date"
            value={searchCriteria.toDate}
            onChange={(e) => handleSearchChange(type, 'toDate', e.target.value)}
            className="px-3 py-1 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
          />
        </div>
        <button
          onClick={() => clearSearch(type)}
          className="px-3 py-1 bg-gray-500 text-white rounded-md hover:bg-gray-600 transition-colors text-sm"
        >
          Clear
        </button>
      </div>
    </div>
  );

  const renderTable = (files, label, prefix, searchCriteria, searchType) => {
    const filteredFiles = filterFilesByDate(files, searchCriteria);
    
    return (
      <div className="mb-8 bg-white rounded-lg shadow-md overflow-hidden">
        <div className="bg-gray-100 px-6 py-4 border-b">
          <h2 className="text-lg font-semibold text-gray-800">{label}</h2>
          <p className="text-sm text-gray-600 mt-1">
            Showing {filteredFiles.length} of {files.length} files
          </p>
        </div>
        
        <div className="p-6">
          {renderSearchBar(searchType, searchCriteria)}
          
          {filteredFiles.length === 0 ? (
            <div className="text-center py-8">
              <p className="text-gray-500">
                {files.length === 0 ? "No files found." : "No files match the selected date range."}
              </p>
            </div>
          ) : (
            <div className="overflow-hidden border rounded-lg">
              {/* Fixed header */}
              <div className="bg-gray-100 border-b">
                <div className="grid grid-cols-12 gap-4 px-4 py-3 font-semibold text-gray-700">
                  <div className="col-span-4">Filename</div>
                  <div className="col-span-4">Uploaded At</div>
                  <div className="col-span-4">Action</div>
                </div>
              </div>
              
              {/* Scrollable content */}
              <div className="max-h-80 overflow-y-auto">
                {filteredFiles.map((file, index) => (
                  <div
                    key={file.id}
                    className={`grid grid-cols-12 gap-4 px-4 py-3 border-b hover:bg-gray-50 transition-colors ${
                      index % 2 === 0 ? 'bg-white' : 'bg-gray-25'
                    }`}
                  >
                    <div className="col-span-4 flex items-center">
                      <span className="text-sm font-medium text-gray-900 truncate" title={file.filename}>
                        {file.filename}
                      </span>
                    </div>
                    <div className="col-span-4 flex items-center">
                      <span className="text-sm text-gray-600">
                        {new Date(file.uploaded_at).toLocaleString()}
                      </span>
                    </div>
                    <div className="col-span-4 flex items-center gap-3">
                      <a
                        href={`http://localhost:5000/api/${prefix}-files/${file.id}/download`}
                        className="inline-flex items-center px-3 py-1 bg-blue-600 text-white text-sm rounded-md hover:bg-blue-700 transition-colors"
                        download
                      >
                        <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                        </svg>
                        Download
                      </a>
                      <button
                        onClick={() => handleDelete(prefix, file.id)}
                        className="inline-flex items-center px-3 py-1 bg-red-600 text-white text-sm rounded-md hover:bg-red-700 transition-colors"
                      >
                        <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                        </svg>
                        Delete
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    );
  };

  return (
    <div className="p-6 bg-gray-100 min-h-screen">
      <div className="max-w-7xl mx-auto">
        <h1 className="text-2xl font-bold text-gray-900 mb-8">Saved Files</h1>
        
        {renderTable(budgetFiles, "Budget Files", "budget", budgetSearch, "budget")}
        {renderTable(salesFiles, "Sales Files", "sales", salesSearch, "sales")}
        {renderTable(osFiles, "OS Files", "os", osSearch, "os")}
      </div>
    </div>
  );
};

export default SavedFiles;
