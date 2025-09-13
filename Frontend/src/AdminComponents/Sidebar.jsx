import React from "react";
import api from "../api/axios";

const Sidebar = ({ activeTab, setActiveTab, onLogout, onReset }) => {
  const tabs = [
    "Executives",
    "Branch & Region",
    "Company & Product",
    "Uploads",
    "Saved Files",
  ];

  const handleResetMappings = async () => {
    if (!window.confirm("⚠️ Are you sure you want to reset all mappings?")) return;

    try {
      const { data } = await api.post("/reset_all_mappings");
      if (data.success) {
        alert(✅ ${data.message});
        if (onReset) onReset();
      } else {
        alert(❌ Failed: ${data.message});
      }
    } catch (error) {
      console.error("Error resetting mappings:", error);
      alert("❌ Could not reset mappings. Check console for details.");
    }
  };
  
  return (
    <div className="fixed top-0 left-0 h-screen w-64 bg-gray-800 text-white p-6 flex flex-col justify-between shadow-lg">
      <div>
        <h2 className="text-2xl font-bold mb-6">Admin Portal</h2>
        <ul className="space-y-3">
          {tabs.map((tab) => (
            <li
              key={tab}
              className={`cursor-pointer p-2 rounded hover:bg-gray-700 ${
                activeTab === tab ? "bg-gray-700" : ""
              }`}
              onClick={() => setActiveTab(tab)}
            >
              {tab}
            </li>
          ))}
        </ul>
      </div>

      <div className="space-y-3">
        <button
          onClick={handleResetMappings}
          className="w-full bg-blue-600 text-white py-2 rounded-lg hover:bg-blue-700 text-sm font-semibold shadow transition-colors duration-200"
        >
          Reset Mappings
        </button>

        <button
          onClick={onLogout}
          className="w-full bg-red-400 text-white py-2 rounded-lg hover:bg-red-500 text-sm font-semibold shadow transition-colors duration-200"
        >
          Logout
        </button>
      </div>
    </div>
  );
};

export default Sidebar;
