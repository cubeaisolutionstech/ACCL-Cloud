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

  // File Upload
  const [uploadFile, setUploadFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headerRow, setHeaderRow] = useState(0);
  const [columns, setColumns] = useState([]);
  const [mapping, setMapping] = useState({
    company_col: "",
    product_col: "",
  });

  const [loadingSheets, setLoadingSheets] = useState(false); // Loader for Load Sheets
  const [loadingPreview, setLoadingPreview] = useState(false); // Loader for Preview Columns
  const [loadingProcess, setLoadingProcess] = useState(false); // Loader for Process File

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
    try {
    const formData = new FormData();
    formData.append("file", uploadFile);
    const res = await api.post("/get-sheet-names", formData);
    const sheets = res.data.sheets || [];
    setSheetNames(sheets);
    if (sheets.length > 0) {
      setSelectedSheet(sheets[0]);
    }
    } catch (err) {
      console.error("Error loading sheets", err);
    } finally {
      setLoadingSheets(false);
    }
  };

  const handlePreview = async () => {
    setLoadingPreview(true);
    try {
    const formData = new FormData();
    formData.append("file", uploadFile);
    formData.append("sheet_name", selectedSheet);
    formData.append("header_row", headerRow);
    const res = await api.post("/preview-excel", formData);
    const cols = res.data.columns;
    setColumns(cols);
    setMapping(autoMapColumns(cols));
    } catch (err) {
      console.error("Error previewing columns", err);
    } finally {
      setLoadingPreview(false);
    }
  };

  const handleUpload = async () => {
    setLoadingProcess(true);
    try {
    const formData = new FormData();
    formData.append("file", uploadFile);
    formData.append("sheet_name", selectedSheet);
    formData.append("header_row", headerRow);
    formData.append("company_col", mapping.company_col);
    formData.append("product_col", mapping.product_col);
    await api.post("/upload-company-product-file", formData);
    alert("✅ File processed!");
    fetchAll();
    } catch (err) {
      console.error("Error processing file", err);
    } finally {
      setLoadingProcess(false);
    }
  };

  return (
    <div className="p-6">
      <h2 className="text-xl font-bold mb-4">Company & Product Mapping</h2>

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
                  ⚠️ Removing <strong>{productToRemove}</strong> will affect:
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
                  🗑️ Remove Product
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
                  ⚠️ Removing <strong>{companyToRemove}</strong> will affect:
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
                  🗑️ Remove Company
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
                      <td className="border px-2 py-1">
                        {m.products.length > 0 ? m.products.join(", ") : "None"}
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
          <h3 className="font-semibold mb-2">Upload Company-Product Mapping File</h3>
          <input
            type="file"
            className="mb-2"
            onChange={(e) => setUploadFile(e.target.files[0])}
          />
          {uploadFile && (
            <>
              <button
                onClick={handleSheetLoad}
                className="bg-blue-500 text-white px-3 py-1 rounded mb-3"
                 disabled={loadingSheets}
              >
               {loadingSheets ? 'Loading...' : 'Load Sheets'}
              </button>
              <div className="grid md:grid-cols-2 gap-4">
                <div>
                  <label>Sheet Name</label>
                  <select
                    className="w-full border px-2 py-1"
                    value={selectedSheet}
                    onChange={(e) => setSelectedSheet(e.target.value)}
                  >
                    {sheetNames.map((s) => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label>Header Row</label>
                  <input
                    type="number"
                    className="w-full border px-2 py-1"
                    value={headerRow}
                    onChange={(e) => setHeaderRow(e.target.value)}
                  />
                </div>
              </div>
              <button
                onClick={handlePreview}
                className="mt-4 bg-gray-700 text-white px-4 py-2 rounded"
                disabled={loadingPreview}
              >
                {loadingPreview ? 'Loading...' : 'Preview Columns'}
              </button>
            </>
          )}
          {columns.length > 0 && (
            <div className="mt-6 grid md:grid-cols-2 gap-4">
              {[
                ["Company Group Column", "company_col"],
                ["Product Group Column", "product_col"],
              ].map(([label, key]) => (
                <div key={key}>
                  <label>{label}</label>
                  <select
                    className="w-full border px-2 py-1"
                    value={mapping[key]}
                    onChange={(e) =>
                      setMapping({ ...mapping, [key]: e.target.value })
                    }
                  >
                    <option value="">-- Select --</option>
                    {columns.map((col) => (
                      <option key={col} value={col}>
                        {col}
                      </option>
                    ))}
                  </select>
                </div>
              ))}
              <div className="col-span-2">
                 <button
                  onClick={handleUpload}
                  className="mt-4 bg-green-600 text-white px-6 py-2 rounded disabled:bg-gray-400"
                  disabled={loadingProcess}
                >
                  {loadingProcess ? 'Processing...' : 'Process File'}
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default CompanyProductManager;
