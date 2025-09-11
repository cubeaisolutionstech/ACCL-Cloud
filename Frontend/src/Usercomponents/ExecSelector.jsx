// components/ExecSelector.jsx
import React, { useEffect, useState } from 'react';

const ExecSelector = ({ salesExecList, budgetExecList, onChange }) => {
  const [selectAllSales, setSelectAllSales] = useState(true);
  const [selectAllBudget, setSelectAllBudget] = useState(true);
  const [selectedSales, setSelectedSales] = useState([]);
  const [selectedBudget, setSelectedBudget] = useState([]);

  useEffect(() => {
    // Initialize with all selected
    setSelectedSales(salesExecList);
    setSelectedBudget(budgetExecList);
    onChange({ sales: salesExecList, budget: budgetExecList });
  }, [salesExecList, budgetExecList]);

  const toggleExec = (type, exec) => {
    const setter = type === 'sales' ? setSelectedSales : setSelectedBudget;
    const selected = type === 'sales' ? selectedSales : selectedBudget;

    const newList = selected.includes(exec)
      ? selected.filter(e => e !== exec)
      : [...selected, exec];

    setter(newList);
    if (type === 'sales') setSelectAllSales(newList.length === salesExecList.length);
    else setSelectAllBudget(newList.length === budgetExecList.length);
    onChange({ sales: type === 'sales' ? newList : selectedSales, budget: type === 'budget' ? newList : selectedBudget });
  };

  const toggleAll = (type) => {
    const isSelectAll = type === 'sales' ? selectAllSales : selectAllBudget;
    const all = type === 'sales' ? salesExecList : budgetExecList;
    const setter = type === 'sales' ? setSelectedSales : setSelectedBudget;
    const setToggle = type === 'sales' ? setSelectAllSales : setSelectAllBudget;

    if (isSelectAll) {
      setter([]);
      setToggle(false);
    } else {
      setter(all);
      setToggle(true);
    }
    onChange({
      sales: type === 'sales' ? (isSelectAll ? [] : all) : selectedSales,
      budget: type === 'budget' ? (isSelectAll ? [] : all) : selectedBudget
    });
  };

  return (
    <div className="grid grid-cols-2 gap-6 mt-6">
      {/* Sales Executives */}
      <div>
        <h3 className="font-bold text-blue-800 mb-2">Sales Executives</h3>
        <label className="block mb-2">
          <input type="checkbox" checked={selectAllSales} onChange={() => toggleAll('sales')} className="mr-2" />
          Select All Sales Executives
        </label>
        <div className="space-y-1 max-h-48 overflow-y-auto">
          {salesExecList.map(exec => (
            <label key={exec} className="block text-sm">
              <input
                type="checkbox"
                checked={selectedSales.includes(exec)}
                onChange={() => toggleExec('sales', exec)}
                className="mr-2"
              />
              {exec}
            </label>
          ))}
        </div>
      </div>

      {/* Budget Executives */}
      <div>
        <h3 className="font-bold text-blue-800 mb-2">Budget Executives</h3>
        <label className="block mb-2">
          <input type="checkbox" checked={selectAllBudget} onChange={() => toggleAll('budget')} className="mr-2" />
          Select All Budget Executives
        </label>
        <div className="space-y-1 max-h-48 overflow-y-auto">
          {budgetExecList.map(exec => (
            <label key={exec} className="block text-sm">
              <input
                type="checkbox"
                checked={selectedBudget.includes(exec)}
                onChange={() => toggleExec('budget', exec)}
                className="mr-2"
              />
              {exec}
            </label>
          ))}
        </div>
      </div>
    </div>
  );
};

export default ExecSelector;
