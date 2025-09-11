import React from "react";

const Sidebar = ({ activeTab, setActiveTab, onLogout}) => {
  const tabs = [
    "Executives",
    "Branch & Region",
    "Company & Product",
    "Uploads",
    "Saved Files"
  ];

  return (
    <div className="w-64 bg-gray-800 text-white min-h-screen p-6">
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

export default Sidebar;
