// components/BranchSelector.jsx
import React, { useEffect, useState } from 'react';

const BranchSelector = ({ branchList = [], onChange }) => {
  const [selectedBranches, setSelectedBranches] = useState(branchList);
  const [selectAll, setSelectAll] = useState(true);

  useEffect(() => {
    setSelectedBranches(branchList);
    setSelectAll(true);
    onChange(branchList);
  }, [branchList]);

  const toggleBranch = (branch) => {
    const updated = selectedBranches.includes(branch)
      ? selectedBranches.filter(b => b !== branch)
      : [...selectedBranches, branch];

    setSelectedBranches(updated);
    setSelectAll(updated.length === branchList.length);
    onChange(updated);
  };

  const toggleAll = () => {
    if (selectAll) {
      setSelectedBranches([]);
      setSelectAll(false);
      onChange([]);
    } else {
      setSelectedBranches(branchList);
      setSelectAll(true);
      onChange(branchList);
    }
  };

  // Split into 4 columns for layout
  const columns = 4;
  const rowsPerCol = Math.ceil(branchList.length / columns);
  const colChunks = Array.from({ length: columns }, (_, i) =>
    branchList.slice(i * rowsPerCol, (i + 1) * rowsPerCol)
  );

  return (
    <div className="mt-6">
      <h3 className="text-blue-800 font-bold text-lg mb-2">Branch Selection</h3>
      <label className="block mb-3">
        <input type="checkbox" checked={selectAll} onChange={toggleAll} className="mr-2" />
        Select All Branches
      </label>
      <div className="grid grid-cols-4 gap-4">
        {colChunks.map((col, colIndex) => (
          <div key={colIndex} className="space-y-1">
            {col.map((branch) => (
              <label key={branch} className="block text-sm">
                <input
                  type="checkbox"
                  checked={selectedBranches.includes(branch)}
                  onChange={() => toggleBranch(branch)}
                  className="mr-2"
                />
                {branch}
              </label>
            ))}
          </div>
        ))}
      </div>
    </div>
  );
};

export default BranchSelector;
