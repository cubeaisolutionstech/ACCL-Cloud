import React, { useEffect, useState } from "react";
import axios from 'axios';
import api from "../api/axios";
import CustomerManager from "./CustomerManager";

const ExecutiveManagement = () => {
  const [activeTab, setActiveTab] = useState("creation");
  const [executives, setExecutives] = useState([]);
  const [name, setName] = useState("");
  const [code, setCode] = useState("");
  const [selectedRemove, setSelectedRemove] = useState("");

  const fetchExecutives = async () => {
    try {
      const res = await axios.get("http://localhost:5000/api/executives-with-counts");
      console.log("API Response:", res.data); // Debug log
      setExecutives(res.data);
    } catch (error) {
      console.error("Error fetching executives:", error);
    }
  };

  const handleAdd = async () => {
    if (!name) return;
    try {
      await api.post("/executive", { name, code });
      setName("");
      setCode("");
      fetchExecutives();
    } catch (error) {
      console.error("Error adding executive:", error);
    }
  };

  const handleRemove = async () => {
    if (selectedRemove) {
      try {
        await api.delete(`/executive/${selectedRemove}`);
        setSelectedRemove("");
        fetchExecutives();
      } catch (error) {
        console.error("Error removing executive:", error);
      }
    }
  };

  const handleTabChange = (newTab) => {
    setActiveTab(newTab);
    if (newTab === "creation") {
      fetchExecutives();
    }
  };

  const handleDataUpdated = () => {
    fetchExecutives();
  };

  // Safe render function for potentially nested objects
  const safeRender = (value) => {
    if (value === null || value === undefined) {
      return "";
    }
    if (typeof value === 'object') {
      // If it's an object, try to extract a meaningful value
      if (value.name) return String(value.name);
      if (value.code) return String(value.code);
      if (value.value) return String(value.value);
      // Fallback to stringified object
      return JSON.stringify(value);
    }
    return String(value);
  };

  // Safe get executive identifier (for keys and values)
  const getExecutiveId = (executive) => {
    if (typeof executive.name === 'object') {
      return executive.name?.name || executive.name?.code || JSON.stringify(executive.name);
    }
    return executive.name || executive.id || JSON.stringify(executive);
  };

  useEffect(() => {
    fetchExecutives();
  }, []);

  return (
    <div className="bg-white p-6 rounded shadow">
      <div className="flex gap-4 mb-4">
        <button
          className="px-4 py-2 rounded bg-blue-600 text-white hover:bg-blue-700 transition-colors"
          onClick={() => handleTabChange("creation")}
        >
          Manual entry
        </button>
        <button
          className="px-4 py-2 rounded bg-green-600 text-white hover:bg-green-700 transition-colors"
          onClick={() => handleTabChange("customers")}
        >
          File upload
        </button>
      </div>

      {activeTab === "creation" && (
        <div>
          <h3 className="text-xl font-semibold mb-4">Add New Executive</h3>
          <div className="flex gap-4 mb-4">
            <input
              className="border px-4 py-2 w-1/3"
              placeholder="Executive Name"
              value={name}
              onChange={(e) => setName(e.target.value)}
            />
            <input
              className="border px-4 py-2 w-1/3"
              placeholder="Executive Code (optional)"
              value={code}
              onChange={(e) => setCode(e.target.value)}
            />
            <button className="bg-green-600 text-white px-4 py-2 rounded" onClick={handleAdd}>
              Add Executive
            </button>
          </div>

          <h3 className="text-lg font-semibold mb-2">Current Executives</h3>
          <table className="w-full border">
            <thead className="bg-gray-100">
              <tr>
                <th className="p-2 border">Name</th>
                <th className="p-2 border">Code</th>
                <th className="p-2 border">Customers</th>
                <th className="p-2 border">Branches</th>
              </tr>
            </thead>
            <tbody>
              {executives.map((e, index) => (
                <tr key={getExecutiveId(e) || index}>
                  <td className="border p-2">{safeRender(e.name)}</td>
                  <td className="border p-2">{safeRender(e.code)}</td>
                  <td className="border p-2">{safeRender(e.customers)}</td>
                  <td className="border p-2">{safeRender(e.branches)}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className="mt-6">
            <h4 className="font-medium mb-2">Remove Executive</h4>
            <select
              value={selectedRemove}
              onChange={(e) => setSelectedRemove(e.target.value)}
              className="border px-4 py-2 mr-4"
            >
              <option value="">Select Executive</option>
              {executives.map((e, index) => (
                <option key={getExecutiveId(e) || index} value={getExecutiveId(e)}>
                  {safeRender(e.name)}
                </option>
              ))}
            </select>
            <button onClick={handleRemove} className="bg-red-600 text-white px-4 py-2 rounded">
              Remove
            </button>
          </div>
        </div>
      )}

      {activeTab === "customers" && (
        <CustomerManager onDataUpdated={handleDataUpdated} />
      )}
    </div>
  );
};

export default ExecutiveManagement;
