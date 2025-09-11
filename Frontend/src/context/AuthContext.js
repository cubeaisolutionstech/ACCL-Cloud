import React, { createContext, useState, useContext, useEffect } from 'react';
import api from '../services/api';

// Create auth context
const AuthContext = createContext(null);

// Auth provider component
export const AuthProvider = ({ children }) => {
  const [state, setState] = useState({
    user: null,
    token: null,
    isAuthenticated: false,
    loading: true,
    error: null
  });

  // Check if user is authenticated on mount
  useEffect(() => {
    const checkAuthStatus = async () => {
      const token = sessionStorage.getItem('auth_token');
      
      if (!token) {
        setState(prev => ({ ...prev, loading: false }));
        return;
      }
      
      try {
        // Set token in API headers
        api.setAuthToken(token);
        
        // Get current user data
        const response = await api.get('/auth/me');
        
        setState({
          user: response.data,
          token,
          isAuthenticated: true,
          loading: false,
          error: null
        });
      } catch (error) {
        console.error('Auth check failed:', error);
        
        // Clear invalid token
        sessionStorage.removeItem('auth_token');
        api.removeAuthToken();
        
        setState({
          user: null,
          token: null,
          isAuthenticated: false,
          loading: false,
          error: 'Session expired. Please login again.'
        });
      }
    };
    
    checkAuthStatus();
  }, []);

  // Request OTP step
  const requestOtp = async (email, password, userType = 'user') => {
    setState(prev => ({ ...prev, loading: true, error: null }));
    
    try {
      const response = await api.post('/auth/request-otp', {
        email,
        password,
        userType
      });
      
      setState(prev => ({ ...prev, loading: false }));
      
      return {
        success: true,
        userId: response.data.userId,
        message: response.data.message
      };
    } catch (error) {
      setState(prev => ({
        ...prev,
        loading: false,
        error: error.response?.data?.message || 'Failed to send OTP'
      }));
      
      return {
        success: false,
        message: error.response?.data?.message || 'Failed to send OTP'
      };
    }
  };

  // Verify OTP and complete login
  const verifyOtp = async (userId, otp, userType = 'user') => {
    setState(prev => ({ ...prev, loading: true, error: null }));
    
    try {
      const response = await api.post('/auth/verify-otp', {
        userId,
        otp,
        userType
      });
      
      const { token, user } = response.data;
      
      // Store token in sessionStorage (more secure than localStorage)
      sessionStorage.setItem('auth_token', token);
      
      // Set token in API headers
      api.setAuthToken(token);
      
      setState({
        user,
        token,
        isAuthenticated: true,
        loading: false,
        error: null
      });
      
      return { success: true };
    } catch (error) {
      setState(prev => ({
        ...prev,
        loading: false,
        error: error.response?.data?.message || 'Invalid OTP'
      }));
      
      return {
        success: false,
        message: error.response?.data?.message || 'Invalid OTP'
      };
    }
  };

  // Register new user
  const register = async (userData) => {
    setState(prev => ({ ...prev, loading: true, error: null }));
    
    try {
      await api.post('/auth/register', userData);
      
      setState(prev => ({ ...prev, loading: false }));
      
      return { success: true };
    } catch (error) {
      setState(prev => ({
        ...prev,
        loading: false,
        error: error.response?.data?.message || 'Registration failed'
      }));
      
      return {
        success: false,
        message: error.response?.data?.message || 'Registration failed'
      };
    }
  };

  // Logout user
  const logout = async () => {
    setState(prev => ({ ...prev, loading: true }));
    
    try {
      // Call logout endpoint to invalidate token on server
      if (state.token) {
        await api.post('/auth/logout');
      }
    } catch (error) {
      console.error('Logout failed:', error);
    } finally {
      // Clear token from storage
      sessionStorage.removeItem('auth_token');
      
      // Remove token from API headers
      api.removeAuthToken();
      
      // Reset state
      setState({
        user: null,
        token: null,
        isAuthenticated: false,
        loading: false,
        error: null
      });
    }
  };

  // Clear errors
  const clearErrors = () => {
    setState(prev => ({ ...prev, error: null }));
  };

  // Context value
  const value = {
    ...state,
    requestOtp,
    verifyOtp,
    register,
    logout,
    clearErrors
  };

  return (
    <AuthContext.Provider value={value}>
      {children}
    </AuthContext.Provider>
  );
};

// Custom hook to use the auth context
export const useAuth = () => {
  const context = useContext(AuthContext);
  
  if (!context) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  
  return context;
};