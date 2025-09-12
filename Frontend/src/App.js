// App.js
import React, { useEffect, useState } from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import { ExcelDataProvider } from './context/ExcelDataContext';
import LoginSystem from './LoginSystem';
import AdminDashboard from './pages/AdminDashboard'; // adjust path if needed
import UserDashboard from './pages/UserDashboard';   // adjust path if needed

function App() {
  const [user, setUser] = useState(null);

  // On load, get user from localStorage
  useEffect(() => {
    const storedUser = localStorage.getItem('user_data');
    if (storedUser) {
      setUser(JSON.parse(storedUser));
    }
  }, []);

  const handleLogin = (loggedInUser) => {
    setUser(loggedInUser);
  };

  const handleLogout = () => {
    localStorage.removeItem('auth_token');
    localStorage.removeItem('user_data');
    setUser(null);
  };

  return (
    <ExcelDataProvider>
    <Router>
      <Routes>
        <Route
          path="/"
          element={
            user ? (
              user.user_type === 'admin' ? (
                <Navigate to="/admin/dashboard" />
              ) : (
                <Navigate to="/user/dashboard" />
              )
            ) : (
              <LoginSystem onLogin={handleLogin} />
            )
          }
        />

        <Route
          path="/admin/dashboard"
          element={
            user?.user_type === 'admin' ? (
              <AdminDashboard onLogout={handleLogout} />
            ) : (
              <Navigate to="/" />
            )
          }
        />

        <Route
          path="/user/dashboard"
          element={
            user?.user_type === 'regular' ? (
              <UserDashboard onLogout={handleLogout} />
            ) : (
              <Navigate to="/" />
            )
          }
        />
      </Routes>
    </Router>
    </ExcelDataProvider>
  );
}

export default App;
