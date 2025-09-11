import React, { useState } from "react";
import axios from "axios";

function ExcelChat() {
  const [file, setFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [dfPreview, setDfPreview] = useState([]);
  const [columns, setColumns] = useState([]);
  const [requestId, setRequestId] = useState("");
  const [question, setQuestion] = useState("");
  const [response, setResponse] = useState(null);

  const uploadFile = async () => {
    const formData = new FormData();
    formData.append("file", file);
    formData.append("sheet", selectedSheet);
    formData.append("header", 0);
    formData.append("removeUnnamed", "true");
    formData.append("removeLastRow", "true");

    const res = await axios.post("http://127.0.0.1:5000/api/upload_excel", formData);
    setDfPreview(res.data.preview);
    setColumns(res.data.columns);
    setSheetNames(res.data.sheetNames);
    setRequestId(res.data.requestId);
  };

  const askQuestion = async () => {
    const res = await axios.post("http://127.0.0.1:5000/api/ask_question", {
      requestId,
      question
    });
    setResponse(res.data);
  };

  return (
    <div>
      <h2>Chat with Excel</h2>
      <input type="file" onChange={(e) => setFile(e.target.files[0])} />
      <button onClick={uploadFile}>Upload & Preview</button>

      {dfPreview.length > 0 && (
        <>
          <h4>Data Preview</h4>
          <table>
            <thead>
              <tr>
                {columns.map((col) => (
                  <th key={col}>{col}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {dfPreview.map((row, idx) => (
                <tr key={idx}>
                  {columns.map((col) => (
                    <td key={col}>{row[col]}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}

      <input
        type="text"
        value={question}
        onChange={(e) => setQuestion(e.target.value)}
        placeholder="Ask a question"
      />
      <button onClick={askQuestion}>Ask</button>

      {response && (
        <div>
          <h4>Answer:</h4>
          {typeof response.result === "string" ? (
            <p>{response.result}</p>
          ) : (
            <pre>{JSON.stringify(response.result, null, 2)}</pre>
          )}
          <p><strong>Code:</strong></p>
          <code>{response.code}</code>
        </div>
      )}
    </div>
  );
}

export default ExcelChat;
