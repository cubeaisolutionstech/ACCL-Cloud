import React, { useEffect, useState } from "react";
import api from "../api/axios";

const CompanyProductManager = () => {
  const [activeTab, setActiveTab] = useState("manual");

  const [products, setProducts] = useState([]);
  const [companies, setCompanies] = useState([]);
  const [mappings, setMappings] = useState([]);

  const [newProduct, setNewProduct] = useState("");
  const [newCompany, setNewCompany] = useState("");

  const [productToRemove, setProductToRemove] = useState("");
  const [companyToRemove, setCompanyToRemove] = useState("");

  const [selectedCompany, setSelectedCompany] = useState("");
  const [companyProducts, setCompanyProducts] = useState([]);

  // Enhanced file upload states
  const [uploadFile, setUploadFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headerRow, setHeaderRow] = useState(0);
  const [columns, setColumns] = useState([]);
  const [mapping, setMapping] = useState({
    company_col: "",
    product_col: "",
  });

  // Enhanced loading and success states
  const [loadingSheets, setLoadingSheets] = useState(false);
  const [loadingPreview, setLoadingPreview] = useState(false);
  const [loadingProcess, setLoadingProcess] = useState(false);
  const [successMessage, setSuccessMessage] = useState("");
  const [errorMessage, setErrorMessage] = useState("");

  // Auto-hide messages after 5 seconds
  useEffect(() => {
    if (successMessage || errorMessage) {
      const timer = setTimeout(() => {
        setSuccessMessage("");
        setErrorMessage("");
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [successMessage, errorMessage]);

  const fetchAll = async () => {
    const [prodRes, compRes, mapRes] = await Promise.all([
      api.get("/products"),
      api.get("/companies"),
      api.get("/company-product-mappings"),
    ]);
    setProducts(prodRes.data);
    setCompanies(compRes.data);
    setMappings(mapRes.data);
  };

  useEffect(() => {
    fetchAll();
  }, []);

  useEffect(() => {
    const fetchCompanyProducts = async () => {
      if (selectedCompany) {
        const res = await api.get(`/company/${selectedCompany}/products`);
        setCompanyProducts(res.data);
      }
    };
    fetchCompanyProducts();
  }, [selectedCompany]);

  const getCompaniesUsingProduct = (productName) => {
    return mappings
      .filter((m) => m.products.includes(productName))
      .map((m) => m.company);
  };

  const getProductsInCompany = (companyName) => {
    const m = mappings.find((m) => m.company === companyName);
    return m ? m.products : [];
  };

  const deleteProduct = async () => {
    await api.delete(`/product/${productToRemove}`);
    setProductToRemove("");
    fetchAll();
  };

  const deleteCompany = async () => {
    await api.delete(`/company/${companyToRemove}`);
    setCompanyToRemove("");
    fetchAll();
  };

  const autoMapColumns = (cols) => {
    const normalize = (s) => s.toLowerCase().replace(/\s/g, "");
    const findMatch = (keys) => cols.find((col) => keys.some((k) => normalize(col).includes(k)));
    return {
      company_col: findMatch(["company", "firm", "group"]) || "",
      product_col: findMatch(["product", "item", "category"]) || "",
    };
  };

  const handleSheetLoad = async () => {
    setLoadingSheets(true);
    setErrorMessage("");
    try {
      const formData = new FormData();
      formData.append("file", uploadFile);
      const res = await api.post("/get-sheet-names", formData);
      const sheets = res.data.sheets || [];
      setSheetNames(sheets);
      if (sheets.length > 0) {
        setSelectedSheet(sheets[0]);
        setSuccessMessage(`Successfully loaded ${sheets.length} sheet(s) from file`);
      }
    } catch (err) {
      console.error("Error loading sheets", err);
      setErrorMessage("Failed to load sheets from file. Please check the file format.");
    } finally {
      setLoadingSheets(false);
    }
  };

  const handlePreview = async () => {
    setLoadingPreview(true);
    setErrorMessage("");
    try {
      const formData = new FormData();
      formData.append("file", uploadFile);
      formData.append("sheet_name", selectedSheet);
      formData.append("header_row", headerRow);
      const res = await api.post("/preview-excel", formData);
      const cols = res.data.columns;
      setColumns(cols);
      setMapping(autoMapColumns(cols));
      setSuccessMessage(`Successfully previewed ${cols.length} columns. Auto-mapped available fields.`);
    } catch (err) {
      console.error("Error previewing columns", err);
      setErrorMessage("Failed to preview columns. Please check your file and sheet selection.");
    } finally {
      setLoadingPreview(false);
    }
  };

  const handleUpload = async () => {
    if (!mapping.company_col || !mapping.product_col) {
      setErrorMessage("Please map both Company and Product columns before processing.");
      return;
    }

    setLoadingProcess(true);
    setErrorMessage("");
    try {
      const formData = new FormData();
      formData.append("file", uploadFile);
      formData.append("sheet_name", selectedSheet);
      formData.append("header_row", headerRow);
      formData.append("company_col", mapping.company_col);
      formData.append("product_col", mapping.product_col);
      
      await api.post("/upload-company-product-file", formData);
      
      // Reset form after successful upload
      setUploadFile(null);
      setSheetNames([]);
      setSelectedSheet("");
      setColumns([]);
      setMapping({ company_col: "", product_col: "" });
      
      setSuccessMessage("File processed successfully! Company-Product mappings have been updated.");
      fetchAll();
    } catch (err) {
      console.error("Error processing file", err);
      setErrorMessage("Failed to process file. Please check your data and try again.");
    } finally {
      setLoadingProcess(false);
    }
  };

  // Success/Error Message Component
  const MessageDisplay = () => {
    if (!successMessage && !errorMessage) return null;
    
    return (
      <div className={`p-4 rounded-md mb-4 ${
        successMessage 
          ? 'bg-green-50 border border-green-200 text-green-800' 
          : 'bg-red-50 border border-red-200 text-red-800'
      }`}>
        <div className="flex items-center">
          <div className="flex-shrink-0">
            {successMessage ? (
              <svg className="h-5 w-5 text-green-400" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
            ) : (
              <svg className="h-5 w-5 text-red-400" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
              </svg>
            )}
          </div>
          <div className="ml-3">
            <p className="text-sm font-medium">
              {successMessage || errorMessage}
            </p>
          </div>
          <div className="ml-auto">
            <button
              className="inline-flex text-gray-400 hover:text-gray-600"
              onClick={() => {
                setSuccessMessage("");
                setErrorMessage("");
              }}
            >
              <svg className="h-4 w-4" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" />
              </svg>
            </button>
          </div>
        </div>
      </div>
    );
  };

  // Loading Spinner Component
  const LoadingSpinner = () => (
    <svg className="animate-spin -ml-1 mr-3 h-4 w-4 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
    </svg>
  );

  return (
    <div className="p-6">
      <h2 className="text-xl font-bold mb-4">Company & Product Mapping</h2>

      <MessageDisplay />

      <div className="flex mb-6">
        <button
          className={`mr-4 px-4 py-2 rounded-t ${
            activeTab === "manual" ? "bg-blue-600 text-white" : "bg-gray-200"
          }`}
          onClick={() => setActiveTab("manual")}
        >
          Manual Entry
        </button>
        <button
          className={`px-4 py-2 rounded-t ${
            activeTab === "upload" ? "bg-blue-600 text-white" : "bg-gray-200"
          }`}
          onClick={() => setActiveTab("upload")}
        >
          File Upload
        </button>
      </div>

      {/* MANUAL ENTRY */}
      {activeTab === "manual" && (
        <div className="grid md:grid-cols-2 gap-6">
          {/* Column 1: Create */}
          <div>
            <h3 className="font-semibold mb-2">Create Product Group</h3>
            <input
              type="text"
              className="border w-full p-2 mb-2"
              value={newProduct}
              onChange={(e) => setNewProduct(e.target.value)}
            />
            <button
              className="bg-green-600 text-white px-4 py-2 rounded"
              onClick={async () => {
                await api.post("/product", { name: newProduct });
                setNewProduct("");
                fetchAll();
              }}
            >
              Create Product
            </button>

            <h3 className="font-semibold mt-6 mb-2">Create Company Group</h3>
            <input
              type="text"
              className="border w-full p-2 mb-2"
              value={newCompany}
              onChange={(e) => setNewCompany(e.target.value)}
            />
            <button
              className="bg-green-600 text-white px-4 py-2 rounded"
              onClick={async () => {
                await api.post("/company", { name: newCompany });
                setNewCompany("");
                fetchAll();
              }}
            >
              Create Company
            </button>
          </div>

          {/* Column 2: Current + Delete */}
          <div>
            <h4 className="font-semibold mb-2">Current Products</h4>
            <select
              className="w-full border p-2 mb-2"
              value={productToRemove}
              onChange={(e) => setProductToRemove(e.target.value)}
            >
              <option value="">-- Remove Product --</option>
              {products.map((p) => (
                <option key={p.id} value={p.name}>
                  {p.name}
                </option>
              ))}
            </select>

            {productToRemove && (
              <div className="text-sm mb-2">
                <p className="text-yellow-700 font-semibold">
                  Warning: Removing <strong>{productToRemove}</strong> will affect:
                </p>
                <ul className="list-disc list-inside">
                  {getCompaniesUsingProduct(productToRemove).map((c) => (
                    <li key={c}>{c}</li>
                  ))}
                </ul>
                <button
                  onClick={deleteProduct}
                  className="bg-red-600 text-white px-3 py-1 mt-2 rounded"
                >
                  Remove Product
                </button>
              </div>
            )}

            <h4 className="font-semibold mt-6 mb-2">Current Companies</h4>
            <select
              className="w-full border p-2 mb-2"
              value={companyToRemove}
              onChange={(e) => setCompanyToRemove(e.target.value)}
            >
              <option value="">-- Remove Company --</option>
              {companies.map((c) => (
                <option key={c.id} value={c.name}>
                  {c.name}
                </option>
              ))}
            </select>

            {companyToRemove && (
              <div className="text-sm mb-2">
                <p className="text-yellow-700 font-semibold">
                  Warning: Removing <strong>{companyToRemove}</strong> will affect:
                </p>
                <ul className="list-disc list-inside">
                  {getProductsInCompany(companyToRemove).map((p) => (
                    <li key={p}>{p}</li>
                  ))}
                </ul>
                <button
                  onClick={deleteCompany}
                  className="bg-red-600 text-white px-3 py-1 mt-2 rounded"
                >
                  Remove Company
                </button>
              </div>
            )}
          </div>

          {/* Mapping Section */}
          <div className="md:col-span-2 mt-6 border-t pt-4">
            <h3 className="font-semibold mb-2">Map Products to Companies</h3>

            {companies.length > 0 ? (
              <>
                <select
                  className="w-full border p-2 mb-2"
                  value={selectedCompany}
                  onChange={(e) => setSelectedCompany(e.target.value)}
                >
                  <option value="">Select Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.name}>
                      {c.name}
                    </option>
                  ))}
                </select>

                {selectedCompany && (
                  <>
                    {/* Checkbox list for products, checked at top */}
                    <div className="w-full border p-2 h-40 mb-2 overflow-y-auto flex flex-col gap-1">
                      {[
                        ...products.filter(p => companyProducts.includes(p.name)),
                        ...products.filter(p => !companyProducts.includes(p.name))
                      ].map((p) => (
                        <label key={p.id} className="flex items-center gap-2">
                          <input
                            type="checkbox"
                            value={p.name}
                            checked={companyProducts.includes(p.name)}
                            onChange={(ev) => {
                              const checked = ev.target.checked;
                              if (checked) {
                                setCompanyProducts([...new Set([...companyProducts, p.name])]);
                              } else {
                                setCompanyProducts(companyProducts.filter(name => name !== p.name));
                              }
                            }}
                          />
                          <span>{p.name}</span>
                        </label>
                      ))}
                    </div> 
                    <button
                      className="bg-blue-600 text-white px-4 py-2 rounded"
                      onClick={async () => {
                        await api.post("/map-company-products", {
                          company: selectedCompany,
                          products: companyProducts,
                        });
                        fetchAll();
                      }}
                    >
                      Update Company Mapping
                    </button>
                  </>
                )}
              </>
            ) : (
              <p className="text-gray-500">Create companies first.</p>
            )}
          </div>

          {/* Mappings Table */}
          <div className="md:col-span-2 mt-6">
            <h3 className="font-semibold mb-2">Current Mappings</h3>
            {mappings.length > 0 ? (
              <table className="w-full text-sm border">
                <thead className="bg-gray-100">
                  <tr>
                    <th className="border px-2 py-1">Company</th>
                    <th className="border px-2 py-1">Products</th>
                    <th className="border px-2 py-1">Count</th>
                  </tr>
                </thead>
                <tbody>
                  {mappings.map((m, idx) => (
                    <tr key={idx}>
                      <td className="border px-2 py-1">{m.company}</td>
                      <td className="border px-2 py-1 max-w-xs overflow-hidden overflow-ellipsis relative" title={m.products.join(", ")}>
                        <div className="whitespace-normal break-words">{m.products.join(", ")}</div>
                      </td>
                      <td className="border px-2 py-1">{m.count}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <p className="text-gray-500">No mappings created yet.</p>
            )}
          </div>
        </div>
      )}

      {/* FILE UPLOAD */}
      {activeTab === "upload" && (
        <div className="bg-white p-4 border rounded shadow-sm">
          <h3 className="font-semibold mb-4">Upload Company-Product Mapping File</h3>
          
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Select Excel File
              </label>
              <input
                type="file"
                accept=".xlsx,.xls"
                className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-medium file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
                onChange={(e) => {
                  setUploadFile(e.target.files[0]);
                  // Reset dependent states when file changes
                  setSheetNames([]);
                  setSelectedSheet("");
                  setColumns([]);
                  setMapping({ company_col: "", product_col: "" });
                  setSuccessMessage("");
                  setErrorMessage("");
                }}
              />
            </div>

            {uploadFile && (
              <>
                <div className="bg-gray-50 p-4 rounded-md">
                  <p className="text-sm text-gray-600 mb-3">
                    File selected: <span className="font-medium">{uploadFile.name}</span>
                  </p>
                  
                  <button
                    onClick={handleSheetLoad}
                    disabled={loadingSheets}
                    className="inline-flex items-center px-4 py-2 bg-blue-500 text-white rounded-md hover:bg-blue-600 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {loadingSheets && <LoadingSpinner />}
                    {loadingSheets ? 'Loading Sheets...' : 'Load Sheets'}
                  </button>
                </div>

                {sheetNames.length > 0 && (
                  <div className="grid md:grid-cols-2 gap-4 bg-gray-50 p-4 rounded-md">
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">
                        Sheet Name
                      </label>
                      <select
                        className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                        value={selectedSheet}
                        onChange={(e) => setSelectedSheet(e.target.value)}
                      >
                        {sheetNames.map((s) => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">
                        Header Row (0-based)
                      </label>
                      <input
                        type="number"
                        min="0"
                        className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                        value={headerRow}
                        onChange={(e) => setHeaderRow(e.target.value)}
                      />
                    </div>
                    
                    <div className="md:col-span-2">
                      <button
                        onClick={handlePreview}
                        disabled={loadingPreview || !selectedSheet}
                        className="inline-flex items-center px-4 py-2 bg-gray-700 text-white rounded-md hover:bg-gray-800 disabled:opacity-50 disabled:cursor-not-allowed"
                      >
                        {loadingPreview && <LoadingSpinner />}
                        {loadingPreview ? 'Analyzing Columns...' : 'Preview Columns'}
                      </button>
                    </div>
                  </div>
                )}

                {columns.length > 0 && (
                  <div className="bg-gray-50 p-4 rounded-md">
                    <h4 className="font-medium text-gray-900 mb-3">Column Mapping</h4>
                    <div className="grid md:grid-cols-2 gap-4 mb-4">
                      {[
                        ["Company Group Column", "company_col"],
                        ["Product Group Column", "product_col"],
                      ].map(([label, key]) => (
                        <div key={key}>
                          <label className="block text-sm font-medium text-gray-700 mb-1">
                            {label} <span className="text-red-500">*</span>
                          </label>
                          <select
                            className="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                            value={mapping[key]}
                            onChange={(e) =>
                              setMapping({ ...mapping, [key]: e.target.value })
                            }
                          >
                            <option value="">-- Select Column --</option>
                            {columns.map((col) => (
                              <option key={col} value={col}>{col}</option>
                            ))}
                          </select>
                        </div>
                      ))}
                    </div>
                    
                    <button
                      onClick={handleUpload}
                      disabled={loadingProcess || !mapping.company_col || !mapping.product_col}
                      className="inline-flex items-center px-6 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      {loadingProcess && <LoadingSpinner />}
                      {loadingProcess ? 'Processing File...' : 'Process File'}
                    </button>
                    
                    {loadingProcess && (
                      <div className="mt-3 text-sm text-gray-600">
                        Please wait while we process your file. This may take a few moments...
                      </div>
                    )}
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default CompanyProductManager;
