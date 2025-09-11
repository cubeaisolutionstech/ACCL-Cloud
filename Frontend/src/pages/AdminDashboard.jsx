import React, { useState } from "react";
import Sidebar from "../AdminComponents/Sidebar.jsx";
import ExecutiveManager from "../AdminComponents/ExecutiveManager.jsx";
import CompanyProductManager from "../AdminComponents/CompanyProductManager.jsx";
import BranchRegionManager from "../AdminComponents/BranchRegionManager.jsx";
import BudgetProcessor from "../AdminComponents/BudgetProcessor.jsx";
import SalesProcessor from "../AdminComponents/SalesProcessor.jsx";
import OSProcessor from "../AdminComponents/OSProcessor.jsx";
import SavedFiles from "../AdminComponents/SavedFiles.jsx";
import { useNavigate } from "react-router-dom";

// --- Initial State Helpers ---
const getInitialBudgetState = () => ({
  file: null,
  sheetNames: [],
  selectedSheet: "",
  headerRow: 0,
  downloadLink: null,
  processedExcelBase64: null,
  columns: [],
  preview: [],
  customFilename: "",
  loading: false,
  colMap: {
    customer_col: "",
    exec_code_col: "",
    exec_name_col: "",
    branch_col: "",
    region_col: "",
    cust_name_col: "",
  },
  metrics: null,
});

const getInitialSalesState = () => ({
  file: null,
  sheetNames: [],
  selectedSheet: "",
  headerRow: 0,
  columns: [],
  preview: [],
  downloadLink: null,
  processedExcelBase64: null,
  customFilename: "",
  loading: false,
  execCodeCol: "",
  execNameCol: "",
  productCol: "",
  unitCol: "",
  quantityCol: "",
  valueCol: "",
});

const getInitialOSState = () => ({
  file: null,
  sheetNames: [],
  selectedSheet: "",
  headerRow: 1,
  columns: [],
  execCodeCol: "",
  preview: [],
  processedFile: null,
  customFilename: "",
  loading: false,
});

const AdminDashboard = ({ onLogout }) => {
  const [activeTab, setActiveTab] = useState("Executives");
  const [uploadSubTab, setUploadSubTab] = useState("Budget");
  const [sidebarVisible, setSidebarVisible] = useState(true);
  const navigate = useNavigate();

  // --- State for each processor ---
  const [budgetState, setBudgetState] = useState(getInitialBudgetState());
  const [salesState, setSalesState] = useState(getInitialSalesState());
  const [osState, setOSState] = useState(getInitialOSState());

  const handleLogoutClick = () => {
    localStorage.removeItem("auth_token");
    localStorage.removeItem("user_data");
    if (onLogout) onLogout();
    navigate("/");
  };

  const toggleSidebar = () => {
    setSidebarVisible(!sidebarVisible);
  };

  // --- Reset all states when mappings are reset ---
  const handleReset = () => {
    setBudgetState(getInitialBudgetState());
    setSalesState(getInitialSalesState());
    setOSState(getInitialOSState());

    // Easiest: reload whole app so Executives / Customers also refresh
    window.location.reload();
  };

  return (
    <div className="flex">
      {/* Sidebar */}
      {sidebarVisible && (
        <Sidebar
          activeTab={activeTab}
          setActiveTab={setActiveTab}
          onLogout={handleLogoutClick}
          onReset={handleReset} // ✅ reset callback
        />
      )}

      {/* Main Content */}
      <div
        className={`p-6 bg-gray-100 min-h-screen transition-all duration-300 ${
          sidebarVisible ? "ml-64 w-[calc(100%-16rem)]" : "w-full ml-0"
        }`}
      >
        {/* Top Section: Toggle Button + Logo */}
        <div className="flex justify-between items-start mb-4">
          {/* Arrow Button to toggle sidebar */}
          <button
            onClick={toggleSidebar}
            className="p-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition-colors duration-200 shadow-md"
            title={sidebarVisible ? "Hide Sidebar" : "Show Sidebar"}
          >
            {sidebarVisible ? (
              <svg
                className="w-5 h-5"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M15 19l-7-7 7-7"
                />
              </svg>
            ) : (
              <svg
                className="w-5 h-5"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M9 5l7 7-7 7"
                />
              </svg>
            )}
          </button>

          {/* Logo */}
          <div className="text-center">
            <img
              src="/acl_logo.jpg"
              alt="Company Logo"
              className="h-12 w-auto opacity-90 hover:opacity-100 transition-opacity duration-200"
            />

            {/* Company Name */}
            <div className="mt-2 text-center">
              <h1 className="text-2xl font-bold bg-gradient-to-r from-blue-600 via-purple-600 to-green-600 bg-clip-text text-transparent drop-shadow-lg">
                ACCL
              </h1>
            </div>
          </div>
        </div>

        {/* Tab-Specific Content */}
        {activeTab === "Executives" && <ExecutiveManager />}
        {activeTab === "Branch & Region" && <BranchRegionManager />}
        {activeTab === "Company & Product" && <CompanyProductManager />}
        {activeTab === "Uploads" && (
          <div>
            <div className="flex space-x-4 mb-4">
              {["Budget", "Sales", "OS"].map((tab) => (
                <button
                  key={tab}
                  onClick={() => setUploadSubTab(tab)}
                  className={`px-4 py-2 rounded ${
                    uploadSubTab === tab
                      ? "bg-blue-600 text-white"
                      : "bg-gray-200"
                  }`}
                >
                  {tab} Processing
                </button>
              ))}
            </div>

            {/* Pass state and setters as props */}
            {uploadSubTab === "Budget" && (
              <BudgetProcessor state={budgetState} setState={setBudgetState} />
            )}
            {uploadSubTab === "Sales" && (
              <SalesProcessor state={salesState} setState={setSalesState} />
            )}
            {uploadSubTab === "OS" && (
              <OSProcessor state={osState} setState={setOSState} />
            )}
          </div>
        )}
        {activeTab === "Saved Files" && <SavedFiles />}
      </div>
    </div>
  );
};

export default AdminDashboard;
