import React, { useState, useEffect } from 'react';
import { User, Shield, Mail, Lock, Phone, UserPlus, LogIn, RefreshCw, Eye, EyeOff } from 'lucide-react';
import './LoginSystem.css';

const API_BASE_URL = 'http://localhost:5001/api';

const LoginSystem = ({ onLogin }) => {
  // Main state management
  const [activeTab, setActiveTab] = useState('admin');
  const [currentView, setCurrentView] = useState('login'); // 'login', 'register', 'otp'
  const [showPassword, setShowPassword] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [user, setUser] = useState(null);

  // Form states
  const [loginForm, setLoginForm] = useState({
    email: '',
    password: ''
  });

  const [registerForm, setRegisterForm] = useState({
    email: '',
    password: '',
    confirmPassword: '',
    full_name: '',
    phone: ''
  });

  const [otpForm, setOtpForm] = useState({
    otp_code: '',
    user_id: null
  });

  const [message, setMessage] = useState({ type: '', text: '' });

  // Check for existing auth on mount
  useEffect(() => {
    const token = localStorage.getItem('auth_token');
    const userData = localStorage.getItem('user_data');
    if (token && userData) {
      setUser(JSON.parse(userData));
    }
  }, []);

  // API calls
  const apiCall = async (endpoint, method = 'GET', data = null) => {
    try {
      const config = {
        method,
        headers: {
          'Content-Type': 'application/json',
        }
      };

      if (data) {
        config.body = JSON.stringify(data);
      }

      const token = localStorage.getItem('auth_token');
      if (token) {
        config.headers.Authorization = `Bearer ${token}`;
      }

      const response = await fetch(`${API_BASE_URL}${endpoint}`, config);
      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || 'An error occurred');
      }

      return result;
    } catch (error) {
      throw error;
    }
  };

  // Handle login
  const handleLogin = async (e) => {
    e.preventDefault();
    setIsLoading(true);

    try {
      const result = await apiCall('/login', 'POST', {
        ...loginForm,
        user_type: activeTab
      });

      if (result.requires_otp) {
        setOtpForm({ ...otpForm, user_id: result.user_id });
        setCurrentView('otp');
        setMessage({ type: 'success', text: result.message });

        // Show OTP in demo (remove in production)
        if (result.otp) {
          setMessage({
            type: 'info',
            text: `OTP sent! Demo OTP: ${result.otp}`
          });
        }
      } else if (result.token && result.user) {
        // Store auth data
        localStorage.setItem('auth_token', result.token);
        localStorage.setItem('user_data', JSON.stringify(result.user));

        setUser(result.user);
        if (onLogin) onLogin(result.user);
        setMessage({ type: 'success', text: result.message || 'Login successful' });

        // Reset forms
        setLoginForm({ email: '', password: '' });
      }
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }

    setIsLoading(false);
  };

  // Handle registration
  const handleRegister = async (e) => {
    e.preventDefault();

    if (registerForm.password !== registerForm.confirmPassword) {
      setMessage({ type: 'error', text: 'Passwords do not match' });
      return;
    }

    if (registerForm.password.length < 6) {
      setMessage({ type: 'error', text: 'Password must be at least 6 characters' });
      return;
    }

    setIsLoading(true);

    try {
      const result = await apiCall('/register', 'POST', {
        ...registerForm,
        user_type: activeTab
      });

      setMessage({ type: 'success', text: result.message });
      setCurrentView('login');
      setRegisterForm({
        email: '',
        password: '',
        confirmPassword: '',
        full_name: '',
        phone: ''
      });
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }

    setIsLoading(false);
  };

  // Handle OTP verification
  const handleOtpVerification = async (e) => {
    e.preventDefault();
    setIsLoading(true);

    try {
      const result = await apiCall('/verify-otp', 'POST', otpForm);

      // Store auth data
      localStorage.setItem('auth_token', result.token);
      localStorage.setItem('user_data', JSON.stringify(result.user));

      setUser(result.user);
      if (onLogin) onLogin(result.user);
      setMessage({ type: 'success', text: result.message });

      // Reset forms
      setLoginForm({ email: '', password: '' });
      setOtpForm({ otp_code: '', user_id: null });
      setCurrentView('login');
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }

    setIsLoading(false);
  };

  // Handle OTP resend
  const handleResendOtp = async () => {
    setIsLoading(true);

    try {
      const result = await apiCall('/resend-otp', 'POST', {
        user_id: otpForm.user_id
      });

      setMessage({ type: 'success', text: result.message });

      // Show OTP in demo (remove in production)
      if (result.otp) {
        setMessage({
          type: 'info',
          text: `OTP resent! Demo OTP: ${result.otp}`
        });
      }
    } catch (error) {
      setMessage({ type: 'error', text: error.message });
    }

    setIsLoading(false);
  };

  // Handle logout
  const handleLogout = () => {
    localStorage.removeItem('auth_token');
    localStorage.removeItem('user_data');
    setUser(null);
    setMessage({ type: 'success', text: 'Logged out successfully' });
  };

  // Clear messages after 5 seconds
  useEffect(() => {
    if (message.text) {
      const timer = setTimeout(() => {
        setMessage({ type: '', text: '' });
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [message]);

  // If user is logged in, show dashboard
  if (user) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-blue-900 via-purple-900 to-indigo-900 flex items-center justify-center p-4">
        <div className="bg-white rounded-2xl shadow-2xl p-8 w-full max-w-md">
          <div className="text-center mb-8">
            <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
              {user.user_type === 'admin' ? (
                <Shield className="w-8 h-8 text-green-600" />
              ) : (
                <User className="w-8 h-8 text-green-600" />
              )}
            </div>
            <h2 className="text-2xl font-bold text-gray-800">Welcome!</h2>
            <p className="text-gray-600 capitalize">{user.user_type} Dashboard</p>
          </div>

          <div className="space-y-4">
            <div className="bg-gray-50 p-4 rounded-lg">
              <p className="text-sm text-gray-600">Name</p>
              <p className="font-semibold text-gray-800">{user.full_name}</p>
            </div>

            <div className="bg-gray-50 p-4 rounded-lg">
              <p className="text-sm text-gray-600">Email</p>
              <p className="font-semibold text-gray-800">{user.email}</p>
            </div>

            <div className="bg-gray-50 p-4 rounded-lg">
              <p className="text-sm text-gray-600">Account Type</p>
              <p className="font-semibold text-gray-800 capitalize">{user.user_type}</p>
            </div>
          </div>

          <button
            onClick={handleLogout}
            className="w-full mt-6 bg-red-600 text-white py-3 px-4 rounded-lg hover:bg-red-700 transition-colors font-semibold"
          >
            Logout
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-900 via-purple-900 to-indigo-900 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden">

        {/* Header Tabs */}
        <div className="flex bg-gray-100">
          <button
            onClick={() => setActiveTab('admin')}
            className={`flex-1 py-4 px-6 text-sm font-semibold transition-colors ${activeTab === 'admin'
                ? 'bg-white text-blue-600 border-b-2 border-blue-600'
                : 'text-gray-600 hover:text-gray-800'
              }`}
          >
            <Shield className="w-4 h-4 inline mr-2" />
            Admin Login
          </button>
          <button
            onClick={() => setActiveTab('regular')}
            className={`flex-1 py-4 px-6 text-sm font-semibold transition-colors ${activeTab === 'regular'
                ? 'bg-white text-blue-600 border-b-2 border-blue-600'
                : 'text-gray-600 hover:text-gray-800'
              }`}
          >
            <User className="w-4 h-4 inline mr-2" />
            User Login
          </button>
        </div>

        <div className="p-8">
          {/* Messages */}
          {message.text && (
            <div className={`mb-4 p-3 rounded-lg text-sm ${message.type === 'error' ? 'bg-red-100 text-red-700 border border-red-200' :
                message.type === 'success' ? 'bg-green-100 text-green-700 border border-green-200' :
                  'bg-blue-100 text-blue-700 border border-blue-200'
              }`}>
              {message.text}
            </div>
          )}

          {/* OTP Verification View */}
          {currentView === 'otp' && (
            <div>
              <div className="text-center mb-6">
                <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mx-auto mb-4">
                  <Mail className="w-8 h-8 text-blue-600" />
                </div>
                <h2 className="text-2xl font-bold text-gray-800">Verify OTP</h2>
                <p className="text-gray-600">Enter the 6-digit code sent to your email</p>
              </div>

              <form onSubmit={handleOtpVerification} className="space-y-4">
                <div>
                  <input
                    type="text"
                    placeholder="Enter 6-digit OTP"
                    value={otpForm.otp_code}
                    onChange={(e) => setOtpForm({ ...otpForm, otp_code: e.target.value })}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-center text-lg tracking-widest"
                    style={{
                      background: "#fff",
                      color: "#1e293b",
                      fontSize: "1rem",
                      fontWeight: "400",
                      boxShadow: "none"
                    }}
                    maxLength="6"
                    required
                  />
                </div>

                <button
                  type="submit"
                  disabled={isLoading}
                  className="w-full bg-blue-600 text-white py-3 px-4 rounded-lg hover:bg-blue-700 transition-colors font-semibold disabled:opacity-50"
                >
                  {isLoading ? 'Verifying...' : 'Verify OTP'}
                </button>

                <button
                  type="button"
                  onClick={handleResendOtp}
                  disabled={isLoading}
                  className="w-full text-blue-600 py-2 px-4 rounded-lg hover:bg-blue-50 transition-colors font-medium disabled:opacity-50"
                >
                  <RefreshCw className="w-4 h-4 inline mr-2" />
                  Resend OTP
                </button>

                <button
                  type="button"
                  onClick={() => setCurrentView('login')}
                  className="w-full text-gray-600 py-2 px-4 rounded-lg hover:bg-gray-50 transition-colors"
                >
                  Back to Login
                </button>
              </form>
            </div>
          )}

          {/* Login View */}
          {currentView === 'login' && (
            <div>
              <div className="text-center mb-6">
                <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mx-auto mb-4">
                  {activeTab === 'admin' ? (
                    <Shield className="w-8 h-8 text-blue-600" />
                  ) : (
                    <User className="w-8 h-8 text-blue-600" />
                  )}
                </div>
                <h2 className="text-2xl font-bold text-gray-800 capitalize">
                  {activeTab} Login
                </h2>
                <p className="text-gray-600">
                  Sign in to your {activeTab} account
                </p>
              </div>

              <form onSubmit={handleLogin} className="space-y-4">
                <div>
                  <div className="relative">
                    <Mail className="w-5 h-5 text-gray-400 absolute left-3 top-3.5" />
                    <input
                      type="email"
                      placeholder="Email address"
                      value={loginForm.email}
                      onChange={(e) => setLoginForm({ ...loginForm, email: e.target.value })}
                      className="w-full pl-12 pr-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                      required
                    />
                  </div>
                </div>

                <div className="relative">
  <Lock className="w-5 h-5 text-gray-400 absolute left-3 top-3.5" />
  <input
  type={showPassword ? "text" : "password"}
  placeholder="Password"
  value={loginForm.password}
  onChange={(e) => setLoginForm({ ...loginForm, password: e.target.value })}
  className="w-full pl-12 pr-4 py-3 border border-gray-300 rounded-lg bg-white text-gray-900 focus:ring-2 focus:ring-blue-500 focus:border-transparent"
  required
/>
  {/* <button
  type="button"
  onClick={() => setShowPassword(!showPassword)}
  className="absolute right-3 top-3.5 text-gray-400 hover:text-gray-600"
>
  {showPassword ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
</button> */}
</div>
                <button
                  type="submit"
                  disabled={isLoading}
                  className="w-full bg-blue-600 text-white py-3 px-4 rounded-lg hover:bg-blue-700 transition-colors font-semibold disabled:opacity-50"
                >
                  {isLoading ? 'Signing in...' : (
                    <>
                      <LogIn className="w-4 h-4 inline mr-2" />
                      Sign In
                    </>
                  )}
                </button>
              </form>

              <div className="mt-6 text-center">
                <p className="text-gray-600">
                  Don't have an account?{' '}
                  <button
                    onClick={() => setCurrentView('register')}
                    className="text-blue-600 hover:text-blue-800 font-semibold"
                  >
                    Register here
                  </button>
                </p>
              </div>
            </div>
          )}

          {/* Register View */}
          {currentView === 'register' && (
            <div>
              <div className="text-center mb-6">
                <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
                  <UserPlus className="w-8 h-8 text-green-600" />
                </div>
                <h2 className="text-2xl font-bold text-gray-800">
                  Create {activeTab} Account
                </h2>
                <p className="text-gray-600">
                  Join us today as a {activeTab} user
                </p>
              </div>

              <form onSubmit={handleRegister} className="space-y-4">
         
                <div>
                  <div className="relative">
                    <User className="w-5 h-5 text-gray-400 absolute left-3 top-3.5" />
                    <input
                      type="Full name"
                      placeholder="Full Name"
                      value={registerForm.full_name}
                      onChange={(e) => setRegisterForm({ ...registerForm, full_name: e.target.value })}
                      className="w-full pl-12 pr-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                      required
                    />
                  </div>
                </div>

                <div>
                  <div className="relative">
                    <Mail className="w-5 h-5 text-gray-400 absolute left-3 top-3.5" />
                    <input
                      type="email"
                      placeholder="Email address"
                      value={registerForm.email}
                      onChange={(e) => setRegisterForm({ ...registerForm, email: e.target.value })}
                      className="w-full pl-12 pr-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                      required
                    />
                  </div>
                </div>

                <div>
                  <div className="relative">
                    <Phone className="w-5 h-5 text-gray-400 absolute left-3 top-3.5" />
                    <input
                      type="tel"
                      placeholder="Phone Number"
                      value={registerForm.phone}
                      onChange={(e) => setRegisterForm({ ...registerForm, phone: e.target.value })}
                      className="w-full pl-12 pr-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                      required
                    />
                  </div>
                </div>

                <div>
                  <div className="relative">
                    <Lock className="w-5 h-5 text-gray-400 absolute left-3 top-3.5" />
                    <input
                      type={showPassword ? "text" : "password"}
                      placeholder="Password"
                      value={registerForm.password}
                      onChange={(e) => setRegisterForm({ ...registerForm, password: e.target.value })}
                      className="w-full pl-12 pr-12 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                      required
                    />
                 
                  </div>
                </div>

                <div>
                  <div className="relative">
                    <Lock className="w-5 h-5 text-gray-400 absolute left-3 top-3.5" />
                    <input
                      type="password"
                      placeholder="Confirm Password"
                      value={registerForm.confirmPassword}
                      onChange={(e) => setRegisterForm({ ...registerForm, confirmPassword: e.target.value })}
                      className="w-full pl-12 pr-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                      required
                    />
                  </div>
                </div>

                <button
                  type="submit"
                  disabled={isLoading}
                  className="w-full bg-green-600 text-white py-3 px-4 rounded-lg hover:bg-green-700 transition-colors font-semibold disabled:opacity-50"
                >
                  {isLoading ? 'Creating Account...' : (
                    <>
                      <UserPlus className="w-4 h-4 inline mr-2" />
                      Create Account
                    </>
                  )}
                </button>
              </form>

              <div className="mt-6 text-center">
                <p className="text-gray-600">
                  Already have an account?{' '}
                  <button
                    onClick={() => setCurrentView('login')}
                    className="text-blue-600 hover:text-blue-800 font-semibold"
                  >
                    Sign in here
                  </button>
                </p>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default LoginSystem;
