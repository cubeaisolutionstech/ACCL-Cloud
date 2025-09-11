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

// --- Add state for each processor at the parent level ---
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

  return (
    <div className="flex">
      {sidebarVisible && (
        <Sidebar activeTab={activeTab} setActiveTab={setActiveTab} onLogout={handleLogoutClick} />
      )}
      <div className={`flex-1 p-6 bg-gray-100 min-h-screen ${!sidebarVisible ? 'w-full' : ''}`}>
        {/* Logo and arrow button in same row */}
        <div className="flex justify-between items-start mb-4">
          {/* Arrow button to toggle sidebar */}
          <button
            onClick={toggleSidebar}
            className="p-2 bg-black-600 text-blue rounded-lg hover:bg-blue-700 transition-colors duration-200 shadow-md"
            title={sidebarVisible ? "Hide Sidebar" : "Show Sidebar"}
          >
            {sidebarVisible ? (
              <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
              </svg>
            ) : (
              <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
              </svg>
            )}
          </button>

          {/* Logo in top right */}
          <div className="text-center">
            <img 
              src="/acl_logo.jpg" 
              alt="Company Logo" 
              className="h-12 w-auto opacity-90 hover:opacity-100 transition-opacity duration-200"
            />
            
            {/* ACL Title with unique design */}
            <div className="mt-2 text-center">
              <h1 className="text-2xl font-bold bg-gradient-to-r from-blue-600 via-purple-600 to-green-600 bg-clip-text text-transparent drop-shadow-lg">
                ACCL
              </h1>
              {/* <div className="w-8 h-0.5 bg-gradient-to-r from-blue-500 to-green-500 mx-auto mt-1 rounded-full"></div> */}
            </div>
          </div>
        </div>

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

            {/* Pass state and setters as props to each processor */}
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
