// pages/UserDashboard.jsx
import React, { useState } from 'react';
import SidebarUploadPanel from '../Usercomponents/SidebarUploadPanel';

import BudgetVsBilled from '../Usercomponents/BudgetVsBilled';
import OdTargetVsCollection from '../Usercomponents/OdTargetVsCollection';
import ProductGrowth from '../Usercomponents/ProductGrowth';
import NumberOfBilledCustomers from '../Usercomponents/NumberOfBilledCustomers';

import ExecutiveBudgetVsBilled from '../Usercomponents/ExecutiveBudgetVsBilled';
import ExecutiveODC from '../Usercomponents/ExecutiveODC';
import ExecutiveProductGrowth from '../Usercomponents/ExecutiveProductGrowth';
import CustomerODAnalysis from '../Usercomponents/ExecutiveNBC';

import ConsolidatedReportPanel from '../Usercomponents/ConsolidatedReportPanel';
import Chatbot from '../Usercomponents/Chatbot';
import Cumulative from '../VisualizationComponents/Cumulative';
import Dashboard from '../VisualizationComponents/Dashboard';
import Auditor from '../VisualizationComponents/Auditor';
import { useNavigate } from "react-router-dom";


const UserDashboard = ({ onLogout }) => {
  const [branchTab, setBranchTab] = useState('budget');
  const [execTab, setExecTab] = useState('exec_budget');
  const [sidebarVisible, setSidebarVisible] = useState(true);
  const [mainTab, setMainTab] = useState('ppt_generation');
  const [visualizationSubTab, setVisualizationSubTab] = useState('cumulative');

  const [activeModule, setActiveModule] = useState('branch');
  const navigate = useNavigate();

  const handleLogoutClick = () => {
    localStorage.removeItem("auth_token");
    localStorage.removeItem("user_data");
    if (onLogout) onLogout();
    navigate("/");
  }

  const toggleSidebar = () => {
    setSidebarVisible(!sidebarVisible);
  };

  return (
    <div className="flex h-screen overflow-hidden">
      {/* 🧭 Sidebar with logout handler */}
      {sidebarVisible && mainTab === 'ppt_generation' && (
        <div className="w-72 bg-blue-900 text-white shadow-lg flex-shrink-0 overflow-y-auto scrollbar-hide">
        <SidebarUploadPanel mode={activeModule} onLogout={handleLogoutClick} />
        </div>
      )}

      {/* Main content area - Fixed position */}
      <div className={`flex-1 bg-gray-50 overflow-y-auto scrollbar-hide ${(!sidebarVisible || mainTab === 'visualization') ? 'w-full' : ''}`}>
        <div className="p-6">
          {/* Logo and arrow button in same row */}
          <div className="flex justify-between items-start mb-2">
            {/* Arrow button to toggle sidebar - only show for PPT Generation tab */}
            <div className="flex-1">
              {mainTab === 'ppt_generation' && (
                <button
                  onClick={toggleSidebar}
                  className="p-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors duration-200 shadow-md"
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
              )}
            </div>

            {/* Logo in top right */}
            <div className="text-center flex-shrink-0">
              <img 
                src="/acl_logo.jpg" 
                alt="Company Logo" 
                className="h-16 w-auto opacity-90 hover:opacity-100 transition-opacity duration-200"
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

          {/* Main Tab Navigation */}
          <div className="mb-6 flex space-x-4">
            <button
              className={`px-6 py-3 rounded-lg font-semibold transition-colors duration-200 ${
                mainTab === 'ppt_generation' 
                  ? 'bg-blue-600 text-white shadow-lg' 
                  : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
              }`}
              onClick={() => setMainTab('ppt_generation')}
            >
              PPT Generation
            </button>
            <button
              className={`px-6 py-3 rounded-lg font-semibold transition-colors duration-200 ${
                mainTab === 'visualization' 
                  ? 'bg-green-600 text-white shadow-lg' 
                  : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
              }`}
              onClick={() => setMainTab('visualization')}
            >
              Visualization
            </button>
            <button
            className={`px-6 py-3 rounded-lg font-semibold transition-colors duration-200 ${
              mainTab === 'excel_chat' 
                ? 'bg-yellow-600 text-white shadow-lg' 
                : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
            }`}
            onClick={() => setMainTab('excel_chat')}
          >
            🤖 Excel Chat
          </button>
          </div>

          {/* Main Tab Content */}
          {mainTab === 'ppt_generation' && (
            <>

          <ConsolidatedReportPanel />

          {/* Module Toggle */}
          <div className="mb-4 flex space-x-4">
            <button
              className={`px-4 py-2 rounded ${activeModule === 'branch' ? 'bg-green-700 text-white' : 'bg-green-200 text-black'}`}
              onClick={() => {
                setActiveModule('branch');
              }}
            >
              Branch
            </button>
            <button
              className={`px-4 py-2 rounded ${activeModule === 'executive' ? 'bg-purple-700 text-white' : 'bg-purple-200 text-black'}`}
              onClick={() => {
                setActiveModule('executive');
              }}
            >
              Executive
            </button>
          </div>

          {/* Tab Switcher */}
          {activeModule === 'branch' && (
            <div className="mb-4 flex space-x-2">
              <button
                className={`px-4 py-2 rounded-l-lg ${branchTab  === 'budget' ? 'bg-blue-700 text-white' : 'bg-blue-200 text-black'}`}
                onClick={() => setBranchTab('budget')}
              >
                Budget vs Billed
              </button>
              <button
                className={`px-4 py-2 ${branchTab  === 'od' ? 'bg-blue-700 text-white' : 'bg-blue-200 text-black'}`}
                onClick={() => setBranchTab('od')}
              >
                OD Target vs Collection
              </button>
              <button
                className={`px-4 py-2 ${branchTab  === 'product' ? 'bg-blue-700 text-white' : 'bg-blue-200 text-black'}`}
                onClick={() => setBranchTab('product')}
              >
                Product Growth
              </button>
              <button
                className={`px-4 py-2 rounded-r-lg ${branchTab  === 'nbc' ? 'bg-blue-700 text-white' : 'bg-blue-200 text-black'}`}
                onClick={() => setBranchTab('nbc')}
              >
                Number of Billed Customers
              </button>
            </div>
          )}

          {activeModule === 'executive' && (
            <div className="mb-4 flex space-x-2">
              <button
                className={`px-4 py-2 rounded-l-lg ${execTab  === 'exec_budget' ? 'bg-purple-700 text-white' : 'bg-purple-200 text-black'}`}
                onClick={() => setExecTab('exec_budget')}
              >
                Budget vs Billed
              </button>
              <button
                className={`px-4 py-2 ${execTab  === 'exec_od' ? 'bg-purple-700 text-white' : 'bg-purple-200 text-black'}`}
                onClick={() => setExecTab('exec_od')}
              >
                OD Target vs Collection
              </button>
              <button
                className={`px-4 py-2 ${execTab  === 'exec_product' ? 'bg-purple-700 text-white' : 'bg-purple-200 text-black'}`}
                onClick={() => setExecTab('exec_product')}
              >
                Product Growth
              </button>
              <button
                className={`px-4 py-2 rounded-r-lg ${execTab  === 'exec_nbc' ? 'bg-purple-700 text-white' : 'bg-purple-200 text-black'}`}
                onClick={() => setExecTab('exec_nbc')}
              >
                Customer & OD Analysis
              </button>
            </div>
          )}

          {/* Tab Content */}
          <div className="bg-white p-4 rounded shadow">
            {activeModule === 'branch' && (
              <>
                <div style={{ display: branchTab  === 'budget' ? 'block' : 'none' }}>
                  <BudgetVsBilled />
                </div>
                <div style={{ display: branchTab  === 'od' ? 'block' : 'none' }}>
                  <OdTargetVsCollection />
                </div>
                <div style={{ display: branchTab  === 'product' ? 'block' : 'none' }}>
                  <ProductGrowth />
                </div>
                <div style={{ display: branchTab  === 'nbc' ? 'block' : 'none' }}>
                  <NumberOfBilledCustomers />
                </div>
              </>
            )}

            {activeModule === 'executive' && (
              <>
                <div style={{ display: execTab  === 'exec_budget' ? 'block' : 'none' }}>
                  <ExecutiveBudgetVsBilled />
                </div>
                <div style={{ display: execTab  === 'exec_od' ? 'block' : 'none' }}>
                  <ExecutiveODC />
                </div>
                <div style={{ display: execTab  === 'exec_product' ? 'block' : 'none' }}>
                  <ExecutiveProductGrowth />
                </div>
                <div style={{ display: execTab  === 'exec_nbc' ? 'block' : 'none' }}>
                  <CustomerODAnalysis />
                </div>
              </>
            )}
          </div>
          </>
          )}

          {mainTab === 'visualization' && (
            <>
              {/* Visualization Sub-tabs */}
              <div className="mb-6 flex space-x-4">
                
                <button
                  className={`px-6 py-3 rounded-lg font-semibold transition-colors duration-200 ${
                    visualizationSubTab === 'cumulative' 
                      ? 'bg-green-600 text-white shadow-lg' 
                      : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
                  }`}
                  onClick={() => setVisualizationSubTab('cumulative')}
                >
                  📊 Cumulative
                </button>

                <button
                  className={`px-6 py-3 rounded-lg font-semibold transition-colors duration-200 ${
                    visualizationSubTab === 'auditor_format' 
                      ? 'bg-blue-600 text-white shadow-lg' 
                      : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
                  }`}
                  onClick={() => setVisualizationSubTab('auditor_format')}
                >
                  🔍 Auditor Format
                </button>

                <button
                  className={`px-6 py-3 rounded-lg font-semibold transition-colors duration-200 ${
                    visualizationSubTab === 'dashboard' 
                      ? 'bg-blue-600 text-white shadow-lg' 
                      : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
                  }`}
                  onClick={() => setVisualizationSubTab('dashboard')}
                >
                  Dashboard
                </button>
              
              </div>
              

              {/* Visualization Sub-tab Content */}
              

              {visualizationSubTab === 'cumulative' && (
                <Cumulative />
              )}

              {visualizationSubTab === 'auditor_format' && (
                <Auditor />
              )}

              {visualizationSubTab === 'dashboard' && (
                <Dashboard />
              )}
            </>
          )}
          {mainTab === 'excel_chat' && (
          <div className="bg-white p-4 rounded shadow">
            <Chatbot />
          </div>
        )}
        </div>
      </div>
    </div>
  );
};

export default UserDashboard;
