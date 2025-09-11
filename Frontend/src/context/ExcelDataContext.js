import React, { createContext, useContext, useState } from 'react';

const ExcelDataContext = createContext();

export const ExcelDataProvider = ({ children }) => {
  // Admin Context
  const [excelData, setExcelData] = useState({});

  // User Context
  const [selectedFiles, setSelectedFiles] = useState({
    salesFile: null,
    budgetFile: null,
    osPreviousFile: null,
    osCurrentFile: null,
    lySalesFile: null
  });

  // Set parsed data for a file type
  const setParsedExcelData = (type, data) => {
    setExcelData(prev => ({ ...prev, [type]: data }));
  };

  

  return (
    <ExcelDataContext.Provider value={{ excelData, setParsedExcelData , selectedFiles, setSelectedFiles}}>
      {children}
    </ExcelDataContext.Provider>
  );
};

export const useExcelData = () => useContext(ExcelDataContext); 