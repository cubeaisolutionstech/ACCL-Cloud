import React, { useState, useEffect } from 'react';
import { 
  User, Shield, Mail, Lock, Phone, 
  UserPlus, LogIn, RefreshCw, Eye, EyeOff 
} from 'lucide-react';
import './LoginSystem.css';

const API_BASE_URL = 'http://localhost:5001/api';

const LoginSystem = ({ onLogin }) => {
  const [activeTab, setActiveTab] = useState('admin');
  const [currentView, setCurrentView] = useState('login'); 
  // views: login, register, otp, forgotPassword, forgotOtp, resetPassword
  const [showPassword, setShowPassword] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [user, setUser] = useState(null);

  // Forms
  const [loginForm, setLoginForm] = useState({ email: '', password: '' });
  const [registerForm, setRegisterForm] = useState({
    email: '', password: '', confirmPassword: '', full_name: '', phone: ''
  });
  const [otpForm, setOtpForm] = useState({ otp_code: '', user_id: null, email: '', purpose: '' });
  const [resetForm, setResetForm] = useState({ newPassword: '', confirmPassword: '' });
  const [message, setMessage] = useState({ type: '', text: '' });

  // On mount, check localStorage
  useEffect(() => {
    const token = localStorage.getItem('auth_token');
    const userData = localStorage.getItem('user_data');
    if (token && userData) {
      try {
        setUser(JSON.parse(userData));
      } catch (e) {
        localStorage.removeItem('auth_token');
        localStorage.removeItem('user_data');
      }
    }
  }, []);

  // API helper
  const apiCall = async (endpoint, method = 'GET', data = null) => {
    try {
      const config = { 
        method, 
        headers: { 
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        }
      };
      
      if (data) config.body = JSON.stringify(data);
      
      const token = localStorage.getItem('auth_token');
      if (token) config.headers.Authorization = `Bearer ${token}`;
      
      console.log(`Making API call to ${API_BASE_URL}${endpoint}`, { method, data });
      
      const response = await fetch(`${API_BASE_URL}${endpoint}`, config);
      const result = await response.json();
      
      console.log('API Response:', response.status, result);
      
      if (!response.ok) {
        throw new Error(result.error || result.message || `HTTP ${response.status}`);
      }
      
      return result;
    } catch (error) {
      console.error('API Error:', error);
      if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
        throw new Error('Unable to connect to server. Please check if the backend is running.');
      }
      throw error;
    }
  };

  // Handle tab switch
  const handleTabSwitch = (tab) => {
    console.log(`Switching to ${tab} tab`);
    setActiveTab(tab);
    setCurrentView('login');
    setLoginForm({ email: '', password: '' });
    setRegisterForm({ email: '', password: '', confirmPassword: '', full_name: '', phone: '' });
    setOtpForm({ otp_code: '', user_id: null, email: '', purpose: '' });
    setResetForm({ newPassword: '', confirmPassword: '' });
    setMessage({ type: '', text: '' });
    setShowPassword(false);
  };

  // Login
  const handleLogin = async (e) => {
    e.preventDefault();
    setIsLoading(true);
    setMessage({ type: '', text: '' });
    
    try {
      if (!loginForm.email.trim() || !loginForm.password.trim()) {
        throw new Error('Email and password are required');
      }

      const result = await apiCall('/login', 'POST', {
        email: loginForm.email.trim(),
        password: loginForm.password,
        user_type: activeTab
      });

      if (result.requires_otp) {
        setOtpForm({ 
          otp_code: '', 
          user_id: result.user_id, 
          email: loginForm.email.trim(), 
          purpose: 'login' 
        });
        setCurrentView('otp');
        setMessage({ 
          type: 'success', 
          text: result.message + (result.otp ? ` (OTP: ${result.otp})` : '')
        });
      } else if (result.token && result.user) {
        localStorage.setItem('auth_token', result.token);
        localStorage.setItem('user_data', JSON.stringify(result.user));
        setUser(result.user);
        if (onLogin) onLogin(result.user);
      }
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }
    setIsLoading(false);
  };

  // Register
  const handleRegister = async (e) => {
    e.preventDefault();
    
    // Validation
    if (!registerForm.full_name.trim()) {
      setMessage({ type: 'error', text: 'Full name is required' });
      return;
    }
    if (!registerForm.email.trim()) {
      setMessage({ type: 'error', text: 'Email is required' });
      return;
    }
    if (!registerForm.phone.trim()) {
      setMessage({ type: 'error', text: 'Phone number is required' });
      return;
    }
    if (registerForm.password !== registerForm.confirmPassword) {
      setMessage({ type: 'error', text: 'Passwords do not match' });
      return;
    }
    if (registerForm.password.length < 6) {
      setMessage({ type: 'error', text: 'Password must be at least 6 characters' });
      return;
    }
    
    setIsLoading(true);
    setMessage({ type: '', text: '' });
    
    try {
      const result = await apiCall('/register', 'POST', {
        email: registerForm.email.trim(),
        password: registerForm.password,
        full_name: registerForm.full_name.trim(),
        phone: registerForm.phone.trim(),
        user_type: activeTab
      });
      
      setMessage({ type: 'success', text: result.message });
      setCurrentView('login');
      
      // Clear registration form
      setRegisterForm({ 
        email: '', password: '', confirmPassword: '', full_name: '', phone: '' 
      });
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }
    setIsLoading(false);
  };

  // Verify login OTP
  const handleOtpVerification = async (e) => {
    e.preventDefault();
    
    if (!otpForm.otp_code.trim()) {
      setMessage({ type: 'error', text: 'OTP code is required' });
      return;
    }
    
    setIsLoading(true);
    setMessage({ type: '', text: '' });
    
    try {
      const result = await apiCall('/verify-otp', 'POST', {
        user_id: otpForm.user_id,
        otp_code: otpForm.otp_code.trim(),
        email: otpForm.email,
        purpose: otpForm.purpose
      });
      
      localStorage.setItem('auth_token', result.token);
      localStorage.setItem('user_data', JSON.stringify(result.user));
      setUser(result.user);
      
      if (onLogin) onLogin(result.user);
      
      // Reset forms
      setCurrentView('login');
      setOtpForm({ otp_code: '', user_id: null, email: '', purpose: '' });
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }
    setIsLoading(false);
  };

  // Resend OTP
  const handleResendOtp = async () => {
    setIsLoading(true);
    try {
      const result = await apiCall('/resend-otp', 'POST', { 
        user_id: otpForm.user_id, 
        purpose: otpForm.purpose 
      });
      setMessage({ 
        type: 'success', 
        text: result.message + (result.otp ? ` (OTP: ${result.otp})` : '') 
      });
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }
    setIsLoading(false);
  };

  // Forgot Password (send OTP)
  const handleForgotPassword = async (e) => {
    e.preventDefault();
    
    if (!loginForm.email.trim()) {
      setMessage({ type: 'error', text: 'Email is required' });
      return;
    }
    
    setIsLoading(true);
    setMessage({ type: '', text: '' });
    
    try {
      const result = await apiCall('/forgot-password', 'POST', { 
        email: loginForm.email.trim(), 
        user_type: activeTab 
      });
      
      setMessage({ 
        type: 'success', 
        text: result.message + (result.otp ? ` (OTP: ${result.otp})` : '') 
      });
      setOtpForm({ 
        otp_code: '', 
        email: loginForm.email.trim(), 
        user_id: null, 
        purpose: 'reset' 
      });
      setCurrentView('forgotOtp');
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }
    setIsLoading(false);
  };

  // Verify forgot OTP
  const handleForgotOtp = async (e) => {
    e.preventDefault();
    
    if (!otpForm.otp_code.trim()) {
      setMessage({ type: 'error', text: 'OTP code is required' });
      return;
    }
    
    setIsLoading(true);
    setMessage({ type: '', text: '' });
    
    try {
      const result = await apiCall('/verify-forgot-otp', 'POST', { 
        email: otpForm.email, 
        otp_code: otpForm.otp_code.trim() 
      });
      setMessage({ type: 'success', text: result.message });
      setCurrentView('resetPassword');
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }
    setIsLoading(false);
  };

  // Reset password
  const handleResetPassword = async (e) => {
    e.preventDefault();
    
    if (!resetForm.newPassword || !resetForm.confirmPassword) {
      setMessage({ type: 'error', text: 'Both password fields are required' });
      return;
    }
    if (resetForm.newPassword !== resetForm.confirmPassword) {
      setMessage({ type: 'error', text: 'Passwords do not match' });
      return;
    }
    if (resetForm.newPassword.length < 6) {
      setMessage({ type: 'error', text: 'Password must be at least 6 characters' });
      return;
    }
    
    setIsLoading(true);
    setMessage({ type: '', text: '' });
    
    try {
      const result = await apiCall('/reset-password', 'POST', { 
        email: otpForm.email, 
        otp_code: otpForm.otp_code, 
        new_password: resetForm.newPassword 
      });
      
      setMessage({ type: 'success', text: result.message });
      setCurrentView('login');
      
      // Clear forms
      setResetForm({ newPassword: '', confirmPassword: '' });
      setOtpForm({ otp_code: '', user_id: null, email: '', purpose: '' });
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }
    setIsLoading(false);
  };

  // Logout
  const handleLogout = () => {
    localStorage.removeItem('auth_token');
    localStorage.removeItem('user_data');
    setUser(null);
    setMessage({ type: 'success', text: 'Logged out successfully' });
  };

  // Auto-clear messages
  useEffect(() => {
    if (message.text) {
      const timer = setTimeout(() => setMessage({ type: '', text: '' }), 5000);
      return () => clearTimeout(timer);
    }
  }, [message]);

  // Get current view title
  const getViewTitle = () => {
    const userTypeDisplay = activeTab === 'admin' ? 'Admin' : 'User';
    switch (currentView) {
      case 'login': return `${userTypeDisplay} Login`;
      case 'register': return `${userTypeDisplay} Registration`;
      case 'otp': return 'Enter OTP';
      case 'forgotPassword': return 'Forgot Password';
      case 'forgotOtp': return 'Verify Reset OTP';
      case 'resetPassword': return 'Reset Password';
      default: return `${userTypeDisplay} Login`;
    }
  };

  // Get current view subtitle
  const getViewSubtitle = () => {
    switch (currentView) {
      case 'login': return `Sign in to your ${activeTab} account`;
      case 'register': return `Create a new ${activeTab} account`;
      case 'otp': return 'Check your email for the OTP code';
      case 'forgotPassword': return 'Enter your email to reset password';
      case 'forgotOtp': return 'Enter the OTP sent to your email';
      case 'resetPassword': return 'Create your new password';
      default: return '';
    }
  };

  // ------------- UI -------------
  
  // Logged in user dashboard
  if (user) {
    return (
      <div className="min-h-screen flex items-center justify-center p-4 bg-gradient-to-br">
        <div className="bg-white rounded-2xl shadow-2xl p-8 w-full max-w-md">
          <div className="text-center mb-8">
            <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
              {user.user_type === 'admin' ? 
                <Shield className="w-8 h-8 text-green-600" /> : 
                <User className="w-8 h-8 text-green-600" />
              }
            </div>
            <h2 className="text-2xl font-bold dashboard-text">Welcome, {user.full_name}!</h2>
            <p className="dashboard-subtitle capitalize">{user.user_type} Dashboard</p>
            <p className="text-sm dashboard-subtitle">{user.email}</p>
          </div>
          <button 
            onClick={handleLogout} 
            className="w-full mt-6 bg-red-600 text-white py-3 px-4 rounded-lg hover:bg-red-700 transition-colors font-medium"
          >
            Logout
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen flex items-center justify-center p-4 bg-gradient-to-br login-container">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden">

        {/* Tabs - Only show on login and register views */}
        {(currentView === 'login' || currentView === 'register') && (
          <div className="flex bg-gray-50 border-b">
            <button 
              onClick={() => handleTabSwitch('admin')}
              className={`flex-1 py-4 px-6 text-sm font-semibold transition-all duration-200 login-system-tab ${
                activeTab === 'admin' ? 'active' : ''
              }`}
            >
              <Shield className="w-4 h-4 inline mr-2" /> 
              Admin
            </button>
            <button 
              onClick={() => handleTabSwitch('regular')}
              className={`flex-1 py-4 px-6 text-sm font-semibold transition-all duration-200 login-system-tab ${
                activeTab === 'regular' ? 'active' : ''
              }`}
            >
              <User className="w-4 h-4 inline mr-2" /> 
              User
            </button>
          </div>
        )}

        <div className="p-8 form-content">
          {/* Header */}
          <div className="text-center mb-6">
            <h2 className="text-2xl font-bold text-gray-800">{getViewTitle()}</h2>
            <p className="text-gray-600 text-sm mt-1">{getViewSubtitle()}</p>
          </div>

          {/* Messages */}
          {message.text && (
            <div className={`mb-4 p-3 rounded-lg text-sm font-medium ${
              message.type === 'error' 
                ? 'alert-error' 
                : 'alert-success'
            }`}>
              {message.text}
            </div>
          )}

          {/* OTP Verification */}
          {currentView === 'otp' && (
            <form onSubmit={handleOtpVerification} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Enter OTP Code
                </label>
                <input 
                  type="text" 
                  placeholder="000000" 
                  value={otpForm.otp_code}
                  onChange={(e) => setOtpForm({ ...otpForm, otp_code: e.target.value })} 
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg text-center text-lg font-mono tracking-widest focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                  maxLength="6"
                  autoComplete="one-time-code"
                  required 
                />
              </div>
              <button 
                type="submit" 
                disabled={isLoading}
                className="w-full bg-blue-600 text-white py-3 rounded-lg font-medium hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
              >
                {isLoading ? 'Verifying...' : 'Verify OTP'}
              </button>
              <button 
                type="button" 
                onClick={handleResendOtp} 
                disabled={isLoading}
                className="w-full text-blue-600 py-2 font-medium hover:text-blue-700 disabled:opacity-50 link-button"
              >
                <RefreshCw className="w-4 h-4 inline mr-2" /> 
                Resend OTP
              </button>
              <button 
                type="button" 
                onClick={() => setCurrentView('login')} 
                className="w-full text-gray-600 py-2 font-medium hover:text-gray-700 link-button"
              >
                Back to Login
              </button>
            </form>
          )}

          {/* Login */}
          {currentView === 'login' && (
            <form onSubmit={handleLogin} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  <Mail className="w-4 h-4 inline mr-1" />
                  Email
                </label>
                <input 
                  type="email" 
                  placeholder="Enter your email" 
                  value={loginForm.email}
                  onChange={(e) => setLoginForm({ ...loginForm, email: e.target.value })} 
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                  autoComplete="email"
                  required 
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  <Lock className="w-4 h-4 inline mr-1" />
                  Password
                </label>
                <div className="relative">
                  <input 
                    type={showPassword ? 'text' : 'password'} 
                    placeholder="Enter your password" 
                    value={loginForm.password}
                    onChange={(e) => setLoginForm({ ...loginForm, password: e.target.value })} 
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 pr-12" 
                    autoComplete="current-password"
                    required 
                  />
                                    <button
                    type="button"
                    onClick={(e) => {
                      e.preventDefault();
                      setShowPassword(!showPassword);
                    }}
                    className="absolute right-3 top-1/2 transform -translate-y-1/2 text-gray-500 hover:text-gray-700"
                    style={{ 
                      outline: 'none',
                      border: 'none',
                      background: 'transparent'
                    }}
                  >
                    {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
              </div>
              <button 
                type="submit" 
                disabled={isLoading} 
                className="w-full bg-blue-600 text-white py-3 rounded-lg font-medium hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
              >
                {isLoading ? 'Signing in...' : 'Sign In'}
              </button>
              <div className="text-center space-y-2">
                <button 
                  type="button"
                  onClick={() => setCurrentView('forgotPassword')} 
                  className="text-blue-600 hover:text-blue-700 font-medium link-button"
                >
                  Forgot Password?
                </button>
                <div className="text-gray-500">or</div>
                <button 
                  type="button"
                  onClick={() => setCurrentView('register')} 
                  className="text-blue-600 hover:text-blue-700 font-medium link-button"
                >
                  Create new account
                </button>
              </div>
            </form>
          )}

          {/* Register */}
          {currentView === 'register' && (
            <form onSubmit={handleRegister} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Full Name
                </label>
                <input 
                  type="text" 
                  placeholder="Enter your full name" 
                  value={registerForm.full_name}
                  onChange={(e) => setRegisterForm({ ...registerForm, full_name: e.target.value })} 
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                  autoComplete="name"
                  required 
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  <Mail className="w-4 h-4 inline mr-1" />
                  Email
                </label>
                <input 
                  type="email" 
                  placeholder="Enter your email" 
                  value={registerForm.email}
                  onChange={(e) => setRegisterForm({ ...registerForm, email: e.target.value })} 
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                  autoComplete="email"
                  required 
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  <Phone className="w-4 h-4 inline mr-1" />
                  Phone
                </label>
                <input 
                  type="tel" 
                  placeholder="Enter your phone number" 
                  value={registerForm.phone}
                  onChange={(e) => setRegisterForm({ ...registerForm, phone: e.target.value })} 
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                  autoComplete="tel"
                  required 
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  <Lock className="w-4 h-4 inline mr-1" />
                  Password
                </label>
                <div className="relative">
                  <input 
                    type={showPassword ? 'text' : 'password'} 
                    placeholder="Create a password" 
                    value={registerForm.password}
                    onChange={(e) => setRegisterForm({ ...registerForm, password: e.target.value })} 
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 pr-12" 
                    autoComplete="new-password"
                    required 
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword(!showPassword)}
                    className="absolute right-3 top-1/2 transform -translate-y-1/2 text-gray-500 hover:text-gray-700"
                  >
                    {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Confirm Password
                </label>
                <input 
                  type={showPassword ? 'text' : 'password'} 
                  placeholder="Confirm your password" 
                  value={registerForm.confirmPassword}
                  onChange={(e) => setRegisterForm({ ...registerForm, confirmPassword: e.target.value })} 
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                  autoComplete="new-password"
                  required 
                />
              </div>
              <button 
                type="submit" 
                disabled={isLoading}
                className="w-full bg-green-600 text-white py-3 rounded-lg font-medium hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
              >
                {isLoading ? 'Creating Account...' : 'Create Account'}
              </button>
              <div className="text-center">
                <button 
                  type="button"
                  onClick={() => setCurrentView('login')} 
                  className="text-blue-600 hover:text-blue-700 font-medium link-button"
                >
                  Already have an account? Sign in
                </button>
              </div>
            </form>
          )}

          {/* Forgot Password */}
          {currentView === 'forgotPassword' && (
            <form onSubmit={handleForgotPassword} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  <Mail className="w-4 h-4 inline mr-1" />
                  Email
                </label>
                <input 
                  type="email" 
                  placeholder="Enter your email" 
                  value={loginForm.email}
                  onChange={(e) => setLoginForm({ ...loginForm, email: e.target.value })} 
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                  autoComplete="email"
                  required 
                />
              </div>
              <button 
                type="submit" 
                disabled={isLoading}
                className="w-full bg-yellow-600 text-white py-3 rounded-lg font-medium hover:bg-yellow-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
              >
                {isLoading ? 'Sending OTP...' : 'Send Reset OTP'}
              </button>
              <button 
                type="button" 
                onClick={() => setCurrentView('login')} 
                className="w-full text-gray-600 py-2 font-medium hover:text-gray-700 link-button"
              >
                Back to Login
              </button>
            </form>
          )}

          {/* Forgot OTP */}
          {currentView === 'forgotOtp' && (
            <form onSubmit={handleForgotOtp} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Enter Reset OTP
                </label>
                <input 
                  type="text" 
                  placeholder="000000" 
                  value={otpForm.otp_code}
                  onChange={(e) => setOtpForm({ ...otpForm, otp_code: e.target.value })} 
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg text-center text-lg font-mono tracking-widest focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                  maxLength="6"
                  autoComplete="one-time-code"
                  required 
                />
              </div>
              <button 
                type="submit" 
                disabled={isLoading}
                className="w-full bg-blue-600 text-white py-3 rounded-lg font-medium hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
              >
                {isLoading ? 'Verifying...' : 'Verify Reset OTP'}
              </button>
              <button 
                type="button" 
                onClick={() => setCurrentView('forgotPassword')} 
                className="w-full text-gray-600 py-2 font-medium hover:text-gray-700 link-button"
              >
                Back
              </button>
            </form>
          )}

          {/* Reset Password */}
          {currentView === 'resetPassword' && (
            <form onSubmit={handleResetPassword} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  <Lock className="w-4 h-4 inline mr-1" />
                  New Password
                </label>
                <div className="relative">
                  <input 
                    type={showPassword ? 'text' : 'password'} 
                    placeholder="Enter new password" 
                    value={resetForm.newPassword}
                    onChange={(e) => setResetForm({ ...resetForm, newPassword: e.target.value })} 
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 pr-12" 
                    autoComplete="new-password"
                    required 
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword(!showPassword)}
                    className="absolute right-3 top-1/2 transform -translate-y-1/2 text-gray-500 hover:text-gray-700"
                  >
                    {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Confirm New Password
                </label>
                <input 
                  type={showPassword ? 'text' : 'password'} 
                  placeholder="Confirm new password" 
                  value={resetForm.confirmPassword}
                  onChange={(e) => setResetForm({ ...resetForm, confirmPassword: e.target.value })} 
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500" 
                  autoComplete="new-password"
                  required 
                />
              </div>
              <button 
                type="submit" 
                disabled={isLoading}
                className="w-full bg-green-600 text-white py-3 rounded-lg font-medium hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
              >
                {isLoading ? 'Resetting Password...' : 'Reset Password'}
              </button>
            </form>
          )}
        </div>
      </div>
    </div>
  );
};

export default LoginSystem;
