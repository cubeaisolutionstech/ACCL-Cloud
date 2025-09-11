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
    const res = await axios.get("http://localhost:5000/api/executives-with-counts") // custom endpoint with customer + branch count
    setExecutives(res.data);
  };

  const handleAdd = async () => {
    if (!name) return;
    await api.post("/executive", { name, code });
    setName("");
    setCode("");
    fetchExecutives();
  };

  const handleRemove = async () => {
    if (selectedRemove) {
      await api.delete(`/executive/${selectedRemove}`);
      setSelectedRemove("");
      fetchExecutives();
    }
  };

  useEffect(() => {
    fetchExecutives();
  }, []);

  return (
    <div className="bg-white p-6 rounded shadow">
      <div className="flex gap-4 mb-4">
        <button
          className={`px-4 py-2 rounded ${activeTab === "creation" ? "bg-blue-600 text-white" : "bg-gray-200"}`}
          onClick={() => setActiveTab("creation")}
        >
          Executive Creation
        </button>
        <button
          className={`px-4 py-2 rounded ${activeTab === "customers" ? "bg-blue-600 text-white" : "bg-gray-200"}`}
          onClick={() => setActiveTab("customers")}
        >
          Customer Code Management
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
              {executives.map((e) => (
                <tr key={e.name}>
                  <td className="border p-2">{e.name}</td>
                  <td className="border p-2">{e.code}</td>
                  <td className="border p-2">{e.customers}</td>
                  <td className="border p-2">{e.branches}</td>
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
              {executives.map((e) => (
                <option key={e.name} value={e.name}>
                  {e.name}
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
  <CustomerManager />
)}
    </div>
  );
};

export default ExecutiveManagement;
