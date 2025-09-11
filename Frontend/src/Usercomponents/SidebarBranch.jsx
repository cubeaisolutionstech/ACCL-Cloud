// components/SidebarBranch.jsx
import React from 'react';
import axios from 'axios';
import { useExcelData } from '../context/ExcelDataContext';

const SidebarBranch = ({onLogout}) => {
  const { selectedFiles, setSelectedFiles } = useExcelData();

  const uploadFile = async (file, type) => {
    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await axios.post('/api/branch/upload', formData);
      const filename = response.data.filename;

      const keyMap = {
        sales: 'salesFile',
        budget: 'budgetFile',
        osPrev: 'osPrevFile',
        osCurr: 'osCurrFile',
        lastYearSales: 'lastYearSalesFile' 
      };

      setSelectedFiles((prev) => ({
        ...prev,
        [keyMap[type]]: filename
      }));
    } catch (err) {
      console.error(`Error uploading ${type} file:`, err);
      alert(`Failed to upload ${type} file.`);
    }
  };

  const handleFileChange = (e, type) => {
    const file = e.target.files[0];
    if (file) {
      uploadFile(file, type);
    }
  };

  return (
    <div className="w-72 bg-blue-900 text-white p-4 space-y-4 shadow-lg">
      <h2 className="text-xl font-bold mb-4">Upload Files</h2>

      {[
        { label: 'Current Month Sales Excel File', type: 'sales' },
        { label: 'Current Year Budget/Executive Target File', type: 'budget' },
        { label: 'OS Previous File', type: 'osPrev' },
        { label: 'OS Current File', type: 'osCurr' },
        { label: 'Last Year Sales File', type: 'lastYearSales' } 
      ].map(({ label, type }) => (
        <div key={type}>
          <label className="block mb-1 font-semibold">{label}</label>
          <input type="file" accept=".xlsx,.xls" onChange={(e) => handleFileChange(e, type)} className="text-black" />
          {selectedFiles[`${type}File`] && (
            <p className="text-sm mt-1 text-green-300">Uploaded: {selectedFiles[`${type}File`]}</p>
          )}
        </div>
      ))}
      {/* 🔒 Logout Button */}
      <div className="px-6 py-4 border-t">
        <button
          onClick={onLogout}
          className="w-full bg-red-600 text-white py-2 rounded hover:bg-red-700 text-sm font-semibold"
        >
          Logout
        </button>
      </div>
    </div>
  );
};

export default SidebarBranch;
