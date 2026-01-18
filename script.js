/**
 * Course Add/Drop Management System
 * Shaqra University - CCIT
 * Main JavaScript Utilities
 * Version: 2.0
 * Updated: December 2025
 */

// ===================================
// CONFIGURATION
// ===================================

const CONFIG = {
  APP_NAME: 'Course Add/Drop System',
  VERSION: '2.0.0',
  UNIVERSITY: 'Shaqra University',
  DEPARTMENT: 'College of Computing and Information Technology',
  // Replace with your actual Google Apps Script Web App URL
  BACKEND_URL: 'https://script.google.com/macros/s/AKfycbydJgjmOx3TtoJaM89tOwGYInHxcDXLBo20JH3qh9DeSmJYrtrxGGsu09zRRWAdjdmTCA/exec',
  SESSION_KEY: 'dac_user_session',
  SETTINGS_KEY: 'dac_user_settings',
  TIMEOUT: 30000, // 30 seconds
  MAX_RETRY: 3,
  TOAST_DURATION: 5000, // 5 seconds
  DEBUG: true
};

// ===================================
// UTILITY FUNCTIONS
// ===================================

/**
 * Show loading overlay
 */
function showLoading(message = 'Loading...') {
  const overlay = document.createElement('div');
  overlay.className = 'loading-overlay';
  overlay.id = 'loadingOverlay';
  overlay.innerHTML = `
    <div class="spinner"></div>
    <p>${message}</p>
  `;
  document.body.appendChild(overlay);
}

/**
 * Hide loading overlay
 */
function hideLoading() {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) {
    overlay.remove();
  }
}

/**
 * Show alert message
 */
function showAlert(message, type = 'info') {
  const alertDiv = document.createElement('div');
  alertDiv.className = `alert alert-${type}`;
  alertDiv.innerHTML = `
    <span>${type === 'success' ? '✓' : type === 'danger' ? '✕' : type === 'warning' ? '⚠' : 'ℹ'}</span>
    <span>${message}</span>
  `;
  
  const container = document.querySelector('.container') || document.body;
  container.insertBefore(alertDiv, container.firstChild);
  
  // Auto remove after 5 seconds
  setTimeout(() => {
    alertDiv.remove();
  }, 5000);
}

/**
 * Validate email format
 */
function validateEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(String(email).toLowerCase());
}

/**
 * Validate phone number (supports international formats)
 * Accepts: +966501234567, 0501234567, 501234567, (505) 123-4567, etc.
 */
function validatePhone(phone) {
  // Remove common formatting characters
  const cleaned = String(phone).replace(/[\s\-\(\)\+\.]/g, '');
  
  // Check if it's a valid phone number (8-15 digits after cleaning)
  const re = /^[0-9]{8,15}$/;
  return re.test(cleaned);
}

/**
 * Validate required fields
 */
function validateForm(formId) {
  const form = document.getElementById(formId);
  if (!form) return false;
  
  let isValid = true;
  const inputs = form.querySelectorAll('[required]');
  
  inputs.forEach(input => {
    const errorDiv = input.parentElement.querySelector('.form-error');
    if (errorDiv) errorDiv.remove();
    
    if (!input.value.trim()) {
      isValid = false;
      const error = document.createElement('div');
      error.className = 'form-error';
      error.textContent = 'This field is required';
      input.parentElement.appendChild(error);
      input.style.borderColor = 'var(--danger-color)';
    } else {
      input.style.borderColor = '';
    }
    
    // Specific validations
    if (input.type === 'email' && input.value && !validateEmail(input.value)) {
      isValid = false;
      const error = document.createElement('div');
      error.className = 'form-error';
      error.textContent = 'Invalid email format';
      input.parentElement.appendChild(error);
      input.style.borderColor = 'var(--danger-color)';
    }
    
    if (input.type === 'tel' && input.value && !validatePhone(input.value)) {
      isValid = false;
      const error = document.createElement('div');
      error.className = 'form-error';
      error.textContent = 'Invalid phone number format';
      input.parentElement.appendChild(error);
      input.style.borderColor = 'var(--danger-color)';
    }
  });
  
  return isValid;
}

/**
 * Format date to readable string
 */
function formatDate(dateString) {
  if (!dateString) return 'N/A';
  const date = new Date(dateString);
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

/**
 * Get status badge HTML
 */
function getStatusBadge(status) {
  const statusMap = {
    'pending': 'badge-pending',
    'submitted': 'badge-submitted',
    'approved': 'badge-approved',
    'rejected': 'badge-rejected',
    'awaiting_advisor': 'badge-pending',
    'awaiting_hod': 'badge-pending',
    'awaiting_registrar': 'badge-pending',
    'completed': 'badge-approved'
  };
  
  const badgeClass = statusMap[status.toLowerCase()] || 'badge-pending';
  return `<span class="badge ${badgeClass}">${status}</span>`;
}

// ===================================
// SESSION MANAGEMENT
// ===================================

/**
 * Save user session
 */
function saveSession(userData) {
  localStorage.setItem(CONFIG.SESSION_KEY, JSON.stringify(userData));
}

/**
 * Get user session
 */
function getSession() {
  const session = localStorage.getItem(CONFIG.SESSION_KEY);
  return session ? JSON.parse(session) : null;
}

/**
 * Clear user session
 */
function clearSession() {
  localStorage.removeItem(CONFIG.SESSION_KEY);
}

/**
 * Check if user is logged in
 */
function isLoggedIn() {
  return getSession() !== null;
}

/**
 * Require authentication
 */
function requireAuth() {
  if (!isLoggedIn()) {
    window.location.href = 'pages/login.html';
    return false;
  }
  return true;
}

/**
 * Check user role
 */
function hasRole(requiredRole) {
  const session = getSession();
  if (!session) return false;
  return session.role === requiredRole;
}

/**
 * Logout user
 */
function logout() {
  if (confirm('Are you sure you want to logout?')) {
    clearSession();
    window.location.href = '../index.html';
  }
}

// ===================================
// API COMMUNICATION
// ===================================

/**
 * Make API call to Google Apps Script backend
 */
async function apiCall(action, data = {}, retryCount = 0) {
  try {
    // Validate backend URL
    if (!CONFIG.BACKEND_URL || CONFIG.BACKEND_URL.includes('YOUR_GOOGLE_APPS_SCRIPT')) {
      throw new Error('Backend URL not configured. Please set CONFIG.BACKEND_URL in script.js.');
    }

    showLoading(`Processing ${action}...`);
    
    if (CONFIG.DEBUG) {
      console.log(`📤 API Call: ${action}`, data);
    }
    
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), CONFIG.TIMEOUT);
    
    const requestBody = JSON.stringify({
      action: action,
      data: data
    });
    
    if (CONFIG.DEBUG) {
      console.log('📦 Request body:', requestBody);
    }
    
    const response = await fetch(CONFIG.BACKEND_URL, {
      method: 'POST',
      headers: {
        // Use text/plain to avoid CORS preflight while still sending JSON body
        'Content-Type': 'text/plain;charset=utf-8',
      },
      body: requestBody,
      signal: controller.signal,
      redirect: 'follow'
    });
    
    clearTimeout(timeoutId);
    
    if (!response.ok) {
      throw new Error(`HTTP Error: ${response.status} ${response.statusText}`);
    }
    
    const result = await response.json();
    hideLoading();
    
    if (CONFIG.DEBUG) {
      console.log(`📥 API Response: ${action}`, result);
    }
    
    if (result.success) {
      return result.data;
    } else {
      throw new Error(result.message || `Operation failed: ${action}`);
    }
  } catch (error) {
    hideLoading();
    
    // Handle timeout
    if (error.name === 'AbortError') {
      console.error('Request timeout:', error);
      showAlert('Request timed out. Please check your connection and try again.', 'danger');
    } 
    // Handle fetch errors
    else if (error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
      console.error('Network error:', error);
      showAlert('Unable to connect to server. Please check your internet connection or backend URL.', 'danger');
    }
    else {
      console.error('API Error:', error);
      showAlert(error.message || 'An error occurred. Please try again.', 'danger');
    }
    
    // Retry logic for network errors
    if (retryCount < CONFIG.MAX_RETRY && (error.message.includes('Network') || error.name === 'AbortError')) {
      console.log(`Retrying... (${retryCount + 1}/${CONFIG.MAX_RETRY})`);
      return new Promise(resolve => {
        setTimeout(() => {
          resolve(apiCall(action, data, retryCount + 1));
        }, 1000 * (retryCount + 1));
      });
    }
    
    throw error;
  }
}

/**
 * Register new user
 */
async function registerUser(formData) {
  return await apiCall('register', formData);
}

/**
 * Login user
 */
async function loginUser(email, password, role) {
  const userData = await apiCall('login', { email, password, role });
  saveSession(userData);
  return userData;
}

/**
 * Get all courses
 */
async function getCourses() {
  return await apiCall('getCourses');
}

/**
 * Get all advisors
 */
async function getAdvisors() {
  return await apiCall('getAdvisors');
}

/**
 * Submit add/drop request
 */
async function submitRequest(requestData) {
  const session = getSession();
  requestData.studentId = session.id;
  requestData.studentEmail = session.email;
  requestData.studentName = session.fullName;
  return await apiCall('submitRequest', requestData);
}

/**
 * Get requests for user
 */
async function getRequests(userRole, userId) {
  return await apiCall('getRequests', { userRole, userId });
}

/**
 * Update request status
 */
async function updateRequestStatus(requestId, status, comments) {
  const session = getSession();
  return await apiCall('updateStatus', {
    requestId,
    status,
    comments,
    approverRole: session.role,
    approverId: session.id,
    approverName: session.fullName
  });
}

/**
 * Update user profile
 */
async function updateProfile(profileData) {
  const session = getSession();
  profileData.userId = session.id;
  profileData.role = session.role; // Include role to identify the correct sheet
  const result = await apiCall('updateProfile', profileData);
  
  // Update session with new data
  const updatedSession = { ...session, ...profileData };
  delete updatedSession.role; // Don't store role twice
  saveSession(updatedSession);
  
  return result;
}

/**
 * Reset password
 */
async function resetPassword(email) {
  return await apiCall('resetPassword', { email });
}

/**
 * Change password
 */
async function changePassword(oldPassword, newPassword) {
  const session = getSession();
  return await apiCall('changePassword', {
    email: session.email,
    oldPassword,
    newPassword
  });
}

// ===================================
// UI HELPERS
// ===================================

/**
 * Populate dropdown with options
 */
function populateDropdown(selectId, options, valueKey = 'id', textKey = 'name') {
  const select = document.getElementById(selectId);
  if (!select) return;
  
  // Clear existing options except the first one (placeholder)
  while (select.options.length > 1) {
    select.remove(1);
  }
  
  // Add new options
  options.forEach(option => {
    const opt = document.createElement('option');
    opt.value = typeof option === 'object' ? option[valueKey] : option;
    opt.textContent = typeof option === 'object' ? option[textKey] : option;
    select.appendChild(opt);
  });
}

/**
 * Show modal
 */
function showModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    // Ensure any inline display:none from previous hides is cleared
    modal.style.display = 'flex';
    modal.classList.add('active');
  }
}

/**
 * Hide modal
 */
function hideModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    modal.classList.add('closing');
    // Wait for animation to complete before removing active class
    setTimeout(() => {
      modal.classList.remove('active');
      modal.classList.remove('closing');
      // Forcefully hide to avoid stuck overlay/backdrop
      modal.style.display = 'none';
    }, 300);
  }
}

/**
 * Initialize tooltips
 */
function initTooltips() {
  const tooltips = document.querySelectorAll('[data-tooltip]');
  tooltips.forEach(element => {
    element.addEventListener('mouseenter', (e) => {
      const tooltip = document.createElement('div');
      tooltip.className = 'tooltip';
      tooltip.textContent = e.target.getAttribute('data-tooltip');
      document.body.appendChild(tooltip);
      
      const rect = e.target.getBoundingClientRect();
      tooltip.style.top = rect.top - tooltip.offsetHeight - 5 + 'px';
      tooltip.style.left = rect.left + (rect.width / 2) - (tooltip.offsetWidth / 2) + 'px';
    });
    
    element.addEventListener('mouseleave', () => {
      const tooltip = document.querySelector('.tooltip');
      if (tooltip) tooltip.remove();
    });
  });
}

/**
 * Render table from data
 */
function renderTable(tableId, data, columns) {
  const table = document.getElementById(tableId);
  if (!table) return;
  
  let html = '<thead><tr>';
  columns.forEach(col => {
    html += `<th>${col.label}</th>`;
  });
  html += '</tr></thead><tbody>';
  
  if (data.length === 0) {
    html += `<tr><td colspan="${columns.length}" class="text-center">No data available</td></tr>`;
  } else {
    data.forEach(row => {
      html += '<tr>';
      columns.forEach(col => {
        let value = row[col.field];
        if (col.render) {
          value = col.render(value, row);
        }
        html += `<td>${value}</td>`;
      });
      html += '</tr>';
    });
  }
  
  html += '</tbody>';
  table.innerHTML = html;
}

/**
 * Export table to CSV
 */
function exportToCSV(tableId, filename = 'export.csv') {
  const table = document.getElementById(tableId);
  if (!table) return;
  
  let csv = [];
  const rows = table.querySelectorAll('tr');
  
  rows.forEach(row => {
    const cols = row.querySelectorAll('td, th');
    const rowData = Array.from(cols).map(col => {
      return '"' + col.textContent.replace(/"/g, '""') + '"';
    });
    csv.push(rowData.join(','));
  });
  
  const csvContent = csv.join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  window.URL.revokeObjectURL(url);
}

// ===================================
// FORM HELPERS
// ===================================

/**
 * Get form data as object
 */
function getFormData(formId) {
  const form = document.getElementById(formId);
  if (!form) return {};
  
  const formData = new FormData(form);
  const data = {};
  
  for (let [key, value] of formData.entries()) {
    // Auto-convert numeric fields
    if (['completedHours', 'gpa', 'port'].includes(key)) {
      data[key] = value ? parseFloat(value) : '';
    } else if (['age'].includes(key)) {
      data[key] = value ? parseInt(value) : '';
    } else {
      data[key] = value;
    }
  }
  
  return data;
}

/**
 * Reset form
 */
function resetForm(formId) {
  const form = document.getElementById(formId);
  if (form) {
    form.reset();
    // Remove error messages
    form.querySelectorAll('.form-error').forEach(error => error.remove());
    // Reset border colors
    form.querySelectorAll('input, select, textarea').forEach(input => {
      input.style.borderColor = '';
    });
  }
}

/**
 * Set form data
 */
function setFormData(formId, data) {
  const form = document.getElementById(formId);
  if (!form) return;
  
  Object.keys(data).forEach(key => {
    const input = form.elements[key];
    if (input) {
      input.value = data[key];
    }
  });
}

// ===================================
// EVENT LISTENERS
// ===================================

// Close modals when clicking outside
document.addEventListener('click', (e) => {
  if (e.target.classList.contains('modal')) {
    e.target.classList.remove('active');
  }
});

// Close modals on ESC key
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    document.querySelectorAll('.modal.active').forEach(modal => {
      modal.classList.remove('active');
    });
  }
});

// Auto-hide alerts after 5 seconds
document.addEventListener('DOMContentLoaded', () => {
  const alerts = document.querySelectorAll('.alert');
  alerts.forEach(alert => {
    setTimeout(() => {
      alert.style.opacity = '0';
      setTimeout(() => alert.remove(), 300);
    }, 5000);
  });
});

// ===================================
// INITIALIZATION
// ===================================

/**
 * Initialize common elements
 */
function initializeCommonElements() {
  // Add active class to current nav link
  const currentPath = window.location.pathname;
  document.querySelectorAll('.nav-links a, .sidebar-menu a').forEach(link => {
    if (link.getAttribute('href') === currentPath || 
        currentPath.includes(link.getAttribute('href'))) {
      link.classList.add('active');
    }
  });
  
  // Update user info in sidebar
  const session = getSession();
  if (session) {
    const userNameEl = document.querySelector('.user-name');
    const userRoleEl = document.querySelector('.user-role');
    const userAvatarEl = document.querySelector('.user-avatar');
    
    if (userNameEl) userNameEl.textContent = session.fullName;
    if (userRoleEl) userRoleEl.textContent = session.role.replace('_', ' ').toUpperCase();
    if (userAvatarEl) {
      const initials = session.fullName.split(' ').map(n => n[0]).join('').substring(0, 2);
      userAvatarEl.textContent = initials;
    }
  }
  
  // Add logout button functionality
  const logoutBtn = document.getElementById('logoutBtn');
  if (logoutBtn) {
    logoutBtn.addEventListener('click', (e) => {
      e.preventDefault();
      logout();
    });
  }
  
  // Navbar hamburger menu toggle
  const navToggle = document.querySelector('.hamburger');
  const navLinks = document.querySelector('.nav-links');
  
  if (navToggle && navLinks) {
    navToggle.addEventListener('click', (e) => {
      e.stopPropagation();
      navLinks.classList.toggle('active');
    });
    
    // Close nav when clicking on a link
    document.querySelectorAll('.nav-links a').forEach(link => {
      link.addEventListener('click', () => {
        navLinks.classList.remove('active');
      });
    });
    
    // Close nav when clicking outside
    document.addEventListener('click', (e) => {
      if (!navToggle.contains(e.target) && !navLinks.contains(e.target)) {
        navLinks.classList.remove('active');
      }
    });
  }
  
  // Mobile sidebar toggle
  const toggleBtn = document.getElementById('toggleSidebar');
  const sidebar = document.querySelector('.sidebar');
  
  if (toggleBtn && sidebar) {
    toggleBtn.addEventListener('click', () => {
      sidebar.classList.toggle('open');
      toggleBtn.setAttribute('aria-expanded', 
        sidebar.classList.contains('open') ? 'true' : 'false');
    });
    
    // Close sidebar when a link is clicked
    document.querySelectorAll('.sidebar-menu a').forEach(link => {
      link.addEventListener('click', () => {
        sidebar.classList.remove('open');
        toggleBtn.setAttribute('aria-expanded', 'false');
      });
    });
    
    // Close sidebar when clicking outside
    document.addEventListener('click', (e) => {
      if (!sidebar.contains(e.target) && !toggleBtn.contains(e.target)) {
        sidebar.classList.remove('open');
        toggleBtn.setAttribute('aria-expanded', 'false');
      }
    });
  }
}

// Run initialization when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeCommonElements);
} else {
  initializeCommonElements();
}

// ===================================
// EXPORTS (for use in other scripts)
// ===================================

window.DACSystem = {
  // Configuration
  CONFIG,
  
  // Utility functions
  showLoading,
  hideLoading,
  showAlert,
  validateEmail,
  validatePhone,
  validateForm,
  formatDate,
  getStatusBadge,
  
  // Session management
  saveSession,
  getSession,
  clearSession,
  isLoggedIn,
  requireAuth,
  hasRole,
  logout,
  
  // API calls
  apiCall,
  registerUser,
  loginUser,
  getCourses,
  getAdvisors,
  submitRequest,
  getRequests,
  updateRequestStatus,
  updateProfile,
  resetPassword,
  changePassword,
  
  // UI helpers
  populateDropdown,
  showModal,
  hideModal,
  renderTable,
  exportToCSV,
  
  // Form helpers
  getFormData,
  resetForm,
  setFormData
};

// ===================================
// ENHANCED FEATURES
// ===================================

/**
 * Toast Notification System
 */
class ToastNotification {
  constructor() {
    this.container = this.createContainer();
  }

  createContainer() {
    let container = document.getElementById('toast-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'toast-container';
      container.className = 'toast-container';
      document.body.appendChild(container);
    }
    return container;
  }

  show(message, type = 'info', duration = CONFIG.TOAST_DURATION) {
    const icons = {
      success: '✓',
      error: '✕',
      warning: '⚠',
      info: 'ℹ'
    };

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
      <div class="toast-icon">${icons[type] || icons.info}</div>
      <div class="toast-content">
        <div class="toast-message">${message}</div>
      </div>
      <button class="toast-close" onclick="this.parentElement.remove()">×</button>
    `;

    this.container.appendChild(toast);

    // Auto remove
    setTimeout(() => {
      toast.style.animation = 'slideOutRight 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, duration);

    return toast;
  }

  success(message) {
    return this.show(message, 'success');
  }

  error(message) {
    return this.show(message, 'danger');
  }

  warning(message) {
    return this.show(message, 'warning');
  }

  info(message) {
    return this.show(message, 'info');
  }
}

// Initialize Toast System
const toast = new ToastNotification();

/**
 * Enhanced Modal System
 */
class ModalManager {
  constructor() {
    this.modals = new Map();
  }

  create(id, title, content, options = {}) {
    const modal = document.createElement('div');
    modal.id = id;
    modal.className = 'modal-backdrop';
    modal.innerHTML = `
      <div class="modal-content">
        <div class="modal-header">
          <h3 class="modal-title">${title}</h3>
          <button class="toast-close" onclick="modalManager.close('${id}')">×</button>
        </div>
        <div class="modal-body">
          ${content}
        </div>
        ${options.footer ? `<div class="modal-footer">${options.footer}</div>` : ''}
      </div>
    `;

    // Close on backdrop click
    modal.addEventListener('click', (e) => {
      if (e.target === modal && !options.preventBackdropClose) {
        this.close(id);
      }
    });

    document.body.appendChild(modal);
    this.modals.set(id, modal);
    
    // Prevent body scroll
    document.body.style.overflow = 'hidden';

    return modal;
  }

  close(id) {
    const modal = this.modals.get(id);
    if (modal) {
      modal.style.animation = 'fadeOut 0.3s ease';
      setTimeout(() => {
        modal.remove();
        this.modals.delete(id);
        if (this.modals.size === 0) {
          document.body.style.overflow = '';
        }
      }, 300);
    }
  }

  confirm(title, message, onConfirm) {
    const id = 'confirm-modal-' + Date.now();
    const footer = `
      <button class="btn btn-secondary" onclick="modalManager.close('${id}')">Cancel</button>
      <button class="btn btn-primary" onclick="modalManager.handleConfirm('${id}')">Confirm</button>
    `;

    this.create(id, title, `<p>${message}</p>`, { footer });
    
    // Store callback
    this.modals.get(id)._onConfirm = onConfirm;
  }

  handleConfirm(id) {
    const modal = this.modals.get(id);
    if (modal && modal._onConfirm) {
      modal._onConfirm();
    }
    this.close(id);
  }
}

// Initialize Modal Manager
const modalManager = new ModalManager();

/**
 * Local Storage Helper
 */
class StorageHelper {
  static set(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (error) {
      console.error('Storage error:', error);
      return false;
    }
  }

  static get(key, defaultValue = null) {
    try {
      const item = localStorage.getItem(key);
      return item ? JSON.parse(item) : defaultValue;
    } catch (error) {
      console.error('Storage error:', error);
      return defaultValue;
    }
  }

  static remove(key) {
    try {
      localStorage.removeItem(key);
      return true;
    } catch (error) {
      console.error('Storage error:', error);
      return false;
    }
  }

  static clear() {
    try {
      localStorage.clear();
      return true;
    } catch (error) {
      console.error('Storage error:', error);
      return false;
    }
  }
}

/**
 * Form Validation Enhancement
 */
class FormValidator {
  constructor(formId) {
    this.form = document.getElementById(formId);
    this.errors = new Map();
  }

  addRule(fieldName, rules) {
    const field = this.form.querySelector(`[name="${fieldName}"]`);
    if (!field) return;

    field.addEventListener('blur', () => this.validateField(fieldName, rules));
    field.addEventListener('input', () => this.clearFieldError(fieldName));
  }

  validateField(fieldName, rules) {
    const field = this.form.querySelector(`[name="${fieldName}"]`);
    const value = field.value.trim();

    // Required check
    if (rules.required && !value) {
      this.setFieldError(fieldName, rules.requiredMessage || 'This field is required');
      return false;
    }

    // Email validation
    if (rules.email && value && !validateEmail(value)) {
      this.setFieldError(fieldName, 'Please enter a valid email address');
      return false;
    }

    // Min length
    if (rules.minLength && value.length < rules.minLength) {
      this.setFieldError(fieldName, `Minimum ${rules.minLength} characters required`);
      return false;
    }

    // Max length
    if (rules.maxLength && value.length > rules.maxLength) {
      this.setFieldError(fieldName, `Maximum ${rules.maxLength} characters allowed`);
      return false;
    }

    // Pattern
    if (rules.pattern && value && !rules.pattern.test(value)) {
      this.setFieldError(fieldName, rules.patternMessage || 'Invalid format');
      return false;
    }

    // Custom validation
    if (rules.custom) {
      const customError = rules.custom(value);
      if (customError) {
        this.setFieldError(fieldName, customError);
        return false;
      }
    }

    this.clearFieldError(fieldName);
    return true;
  }

  setFieldError(fieldName, message) {
    const field = this.form.querySelector(`[name="${fieldName}"]`);
    this.errors.set(fieldName, message);
    
    field.classList.add('is-invalid');
    field.style.borderColor = 'var(--danger-color)';
    
    // Remove existing error
    const existingError = field.parentElement.querySelector('.form-error');
    if (existingError) existingError.remove();
    
    // Add new error
    const errorDiv = document.createElement('div');
    errorDiv.className = 'form-error';
    errorDiv.textContent = message;
    field.parentElement.appendChild(errorDiv);
  }

  clearFieldError(fieldName) {
    const field = this.form.querySelector(`[name="${fieldName}"]`);
    this.errors.delete(fieldName);
    
    field.classList.remove('is-invalid');
    field.style.borderColor = '';
    
    const errorDiv = field.parentElement.querySelector('.form-error');
    if (errorDiv) errorDiv.remove();
  }

  validateAll(rules) {
    let isValid = true;
    for (const [fieldName, fieldRules] of Object.entries(rules)) {
      if (!this.validateField(fieldName, fieldRules)) {
        isValid = false;
      }
    }
    return isValid;
  }
}

/**
 * Data Table with Sorting and Filtering
 */
class DataTable {
  constructor(tableId, data, columns) {
    this.table = document.getElementById(tableId);
    this.data = data;
    this.columns = columns;
    this.sortColumn = null;
    this.sortDirection = 'asc';
    this.filterText = '';
  }

  render() {
    if (!this.table) return;

    let filteredData = this.filterData();
    let sortedData = this.sortData(filteredData);

    const thead = this.table.querySelector('thead');
    const tbody = this.table.querySelector('tbody');

    // Render header
    thead.innerHTML = `
      <tr>
        ${this.columns.map(col => `
          <th onclick="dataTable.sort('${col.key}')" class="cursor-pointer">
            ${col.label}
            ${this.sortColumn === col.key ? (this.sortDirection === 'asc' ? '↑' : '↓') : ''}
          </th>
        `).join('')}
      </tr>
    `;

    // Render rows
    tbody.innerHTML = sortedData.map(row => `
      <tr>
        ${this.columns.map(col => `
          <td>${col.render ? col.render(row[col.key], row) : row[col.key]}</td>
        `).join('')}
      </tr>
    `).join('');
  }

  sort(columnKey) {
    if (this.sortColumn === columnKey) {
      this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
      this.sortColumn = columnKey;
      this.sortDirection = 'asc';
    }
    this.render();
  }

  sortData(data) {
    if (!this.sortColumn) return data;

    return [...data].sort((a, b) => {
      let aVal = a[this.sortColumn];
      let bVal = b[this.sortColumn];

      if (typeof aVal === 'string') {
        aVal = aVal.toLowerCase();
        bVal = bVal.toLowerCase();
      }

      if (this.sortDirection === 'asc') {
        return aVal > bVal ? 1 : -1;
      } else {
        return aVal < bVal ? 1 : -1;
      }
    });
  }

  filter(searchText) {
    this.filterText = searchText.toLowerCase();
    this.render();
  }

  filterData() {
    if (!this.filterText) return this.data;

    return this.data.filter(row => {
      return this.columns.some(col => {
        const value = String(row[col.key]).toLowerCase();
        return value.includes(this.filterText);
      });
    });
  }

  updateData(newData) {
    this.data = newData;
    this.render();
  }
}

/**
 * Progress Bar
 */
function showProgress(percentage, containerId = 'progress-container') {
  let container = document.getElementById(containerId);
  if (!container) {
    container = document.createElement('div');
    container.id = containerId;
    container.className = 'progress-container';
    container.innerHTML = '<div class="progress-bar" style="width: 0%"></div>';
    document.body.appendChild(container);
  }

  const bar = container.querySelector('.progress-bar');
  bar.style.width = percentage + '%';

  if (percentage >= 100) {
    setTimeout(() => container.remove(), 500);
  }
}

/**
 * Debounce Function
 */
function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

/**
 * Throttle Function
 */
function throttle(func, limit) {
  let inThrottle;
  return function(...args) {
    if (!inThrottle) {
      func.apply(this, args);
      inThrottle = true;
      setTimeout(() => inThrottle = false, limit);
    }
  };
}

/**
 * Copy to Clipboard
 */
async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    toast.success('Copied to clipboard!');
    return true;
  } catch (error) {
    console.error('Copy failed:', error);
    toast.error('Failed to copy');
    return false;
  }
}

/**
 * Download File
 */
function downloadFile(data, filename, type = 'text/plain') {
  const blob = new Blob([data], { type });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
  toast.success('File downloaded successfully!');
}

/**
 * Print Element
 */
function printElement(elementId) {
  const element = document.getElementById(elementId);
  if (!element) return;

  const printWindow = window.open('', '', 'height=600,width=800');
  printWindow.document.write('<html><head><title>Print</title>');
  printWindow.document.write('<link rel="stylesheet" href="../style.css">');
  printWindow.document.write('</head><body>');
  printWindow.document.write(element.innerHTML);
  printWindow.document.write('</body></html>');
  printWindow.document.close();
  printWindow.print();
}

/**
 * Auto-save Form Data
 */
function enableAutoSave(formId, storageKey) {
  const form = document.getElementById(formId);
  if (!form) return;

  // Load saved data
  const savedData = StorageHelper.get(storageKey);
  if (savedData) {
    Object.keys(savedData).forEach(key => {
      const field = form.querySelector(`[name="${key}"]`);
      if (field) field.value = savedData[key];
    });
  }

  // Save on change
  const saveData = debounce(() => {
    const formData = new FormData(form);
    const data = Object.fromEntries(formData);
    StorageHelper.set(storageKey, data);
  }, 1000);

  form.addEventListener('input', saveData);
  form.addEventListener('change', saveData);
}

/**
 * Initialize App
 */
function initializeApp() {
  console.log(`${CONFIG.APP_NAME} v${CONFIG.VERSION}`);
  console.log(`${CONFIG.UNIVERSITY} - ${CONFIG.DEPARTMENT}`);
  
  // Add loading class to body
  document.body.classList.add('loaded');
  
  // Initialize tooltips if any
  initializeTooltips();
  
  // Check session on protected pages
  if (window.location.pathname.includes('dashboard')) {
    requireAuth();
  }
}

/**
 * Initialize Tooltips
 */
function initializeTooltips() {
  const tooltips = document.querySelectorAll('[data-tooltip]');
  tooltips.forEach(element => {
    element.title = element.getAttribute('data-tooltip');
  });
}

// Auto-initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeApp);
} else {
  initializeApp();
}

// Export enhanced utilities
window.DACUtils = {
  ...window.DACUtils,
  toast,
  modalManager,
  StorageHelper,
  FormValidator,
  DataTable,
  showProgress,
  debounce,
  throttle,
  copyToClipboard,
  downloadFile,
  printElement,
  enableAutoSave
};
