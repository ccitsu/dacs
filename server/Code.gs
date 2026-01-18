/**
 * Course Add/Drop Management System
 * Google Apps Script - Main Server Code
 * 
 * Deploy this as a Web App:
 * 1. Go to Extensions > Apps Script
 * 2. Copy all server files to the script editor
 * 3. Deploy > New Deployment > Web App
 * 4. Execute as: Me
 * 5. Who has access: Anyone
 * 6. Copy the Web App URL to CONFIG.BACKEND_URL in script.js
 */

// ===================================
// CONFIGURATION
// ===================================

const SPREADSHEET_ID = '1On5wouw72tDj5c6zHA9Lq3nNU5s94Ea3frI8aw1jpow';
const DRIVE_FOLDER_ID = '1Ck7nxtFZhP-bLNaYgJSq9yFsCZsm0S5R';

// Sheet names
const SHEETS = {
  STUDENTS: 'Students',
  ADVISORS: 'Advisors',
  HODS: 'HODs',
  REGISTRARS: 'Registrars',
  COURSES: 'Courses',
  REQUESTS: 'Requests',
  APPROVALS: 'Approvals',
  VERIFICATION_CODES: 'VerificationCodes',
  RESET_CODES: 'PasswordResetCodes'
};

// ===================================
// MAIN HANDLER
// ===================================

function doPost(e) {
  try {
    Logger.log('====== NEW REQUEST ======');
    Logger.log('e parameter type: ' + typeof e);

    if (!e) {
      throw new Error('No request object received');
    }

    var requestData;
    var action;
    var data;

    // Try multiple ways to get the data
    if (e.postData && e.postData.contents) {
      Logger.log('✓ Using e.postData.contents');
      const requestBody = e.postData.contents;
      Logger.log('Request body: ' + requestBody);
      try {
        requestData = JSON.parse(requestBody);
      } catch (parseErr) {
        throw new Error('Invalid JSON payload received. Ensure the Web App URL is correct and the request body is JSON. Details: ' + parseErr);
      }
      action = requestData.action;
      data = requestData.data;
    } else if (e.parameter && e.parameter.action) {
      Logger.log('✓ Using e.parameter');
      action = e.parameter.action;
      data = e.parameter.data;
      // If data is a string, try to parse it
      if (typeof data === 'string') {
        try {
          data = JSON.parse(data);
        } catch (ex) {
          Logger.log('Could not parse data as JSON');
        }
      }
    } else {
      Logger.log('✗ No valid data source found');
      Logger.log('Has postData: ' + !!e.postData);
      Logger.log('Has parameter: ' + !!e.parameter);
      Logger.log('Has parameters: ' + !!e.parameters);
      if (e.postData) {
        Logger.log('postData keys: ' + Object.keys(e.postData).join(', '));
      }
      throw new Error('No data received. Invalid request format.');
    }

    // Validate action
    if (!action) {
      throw new Error('Action parameter is required');
    }

    Logger.log(`Action: ${action}`);
    Logger.log(`Data: ${JSON.stringify(data)}`);

    // Basic payload validation per action
    validateActionPayload(action, data);

    let result;

    switch(action) {
      case 'sendVerificationCode':
        result = sendVerificationCode(data);
        break;
      case 'verifyCode':
        result = verifyCodeAndRegister(data);
        break;
      case 'register':
        result = registerUser(data);
        break;
      case 'login':
        result = loginUser(data);
        break;
      case 'getCourses':
        result = getCourses();
        break;
      case 'getAdvisors':
        result = getAdvisors();
        break;
      case 'submitRequest':
        result = submitRequest(data);
        break;
      case 'getRequests':
        result = getRequests(data);
        break;
      case 'updateStatus':
        result = updateRequestStatus(data);
        break;
      case 'sendResetCode':
      case 'sendPasswordResetCode':
        result = sendPasswordResetCode(data);
        break;
      case 'verifyResetCode':
      case 'verifyPasswordResetCode':
        result = verifyPasswordResetCode(data);
        break;
      case 'updateProfile':
        result = updateProfile(data);
        break;
      case 'resetPassword':
        result = resetPassword(data);
        break;
      case 'changePassword':
        result = changePassword(data);
        break;
      case 'downloadPDF':
        result = getPDFUrl(data);
        break;
      default:
        throw new Error(`Invalid action: ${action}`);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: result
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('Error: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Validate required payload keys per action to avoid malformed requests
function validateActionPayload(action, data) {
  const requires = {
    // For multi-course add/drop we expect structured lists and a request type
    submitRequest: ['studentId', 'studentEmail', 'studentName', 'requestType', 'reason'],
    getRequests: ['userRole', 'userId'],
    updateStatus: ['requestId', 'status', 'approverRole', 'approverId'],
    sendPasswordResetCode: ['email'],
    verifyPasswordResetCode: ['email', 'code', 'newPassword'],
    login: ['email', 'password'],
    register: ['email', 'password', 'fullName'],
    updateProfile: ['userId'],
    changePassword: ['email', 'oldPassword', 'newPassword']
  };

  const required = requires[action];
  if (!required) return;

  if (!data || typeof data !== 'object') {
    throw new Error('Invalid payload: expected object');
  }

  const missing = required.filter((key) => !data[key]);
  if (missing.length) {
    throw new Error('Missing required fields: ' + missing.join(', '));
  }

  if (data.email && !isValidEmail(data.email)) {
    throw new Error('Invalid email format');
  }
}

function isValidEmail(email) {
  const normalized = (email || '').trim();
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(normalized);
}

function doGet(e) {
  Logger.log('GET request received');
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'Course Add/Drop System API is running',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

// ===================================
// USER MANAGEMENT
// ===================================

/**
 * Generate a 6-digit verification code
 */
function generateVerificationCode() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

/**
 * Find user by email across all roles using role-specific column mapping
 */
function findUserByEmail(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const emailLower = (email || '').toLowerCase().trim();
  const userTypes = [
    { sheet: SHEETS.STUDENTS, role: 'student', emailCol: 3, passwordCol: 12, nameCol: 1 },
    { sheet: SHEETS.ADVISORS, role: 'advisor', emailCol: 2, passwordCol: 7, nameCol: 1 },
    { sheet: SHEETS.HODS, role: 'hod', emailCol: 3, passwordCol: 6, nameCol: 1 },
    { sheet: SHEETS.REGISTRARS, role: 'registrar', emailCol: 2, passwordCol: 5, nameCol: 1 }
  ];

  for (const userType of userTypes) {
    const sheet = ss.getSheetByName(userType.sheet);
    if (!sheet) continue;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const rowEmail = (rows[i][userType.emailCol] || '').toLowerCase().trim();
      if (rowEmail === emailLower) {
        return {
          sheet,
          sheetName: userType.sheet,
          role: userType.role,
          rowIndex: i,
          userId: rows[i][0],
          fullName: rows[i][userType.nameCol] || '',
          passwordCol: userType.passwordCol
        };
      }
    }
  }
  return null;
}

/**
 * Send verification code to email
 */
function sendVerificationCode(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);
  
  // Check if email already exists
  const existingUsers = studentsSheet.getDataRange().getValues();
  for (let i = 1; i < existingUsers.length; i++) {
    if (existingUsers[i][3] === data.email) {
      throw new Error('Email already registered');
    }
  }
  
  // Generate verification code
  const verificationCode = generateVerificationCode();
  const timestamp = new Date();
  const expiryTime = new Date(timestamp.getTime() + 15 * 60000); // 15 minutes expiry
  
  // Store verification code
  let verificationSheet = ss.getSheetByName(SHEETS.VERIFICATION_CODES);
  // Gracefully handle alternate sheet naming with space
  if (!verificationSheet) {
    verificationSheet = ss.getSheetByName('Verification Codes');
  }
  if (!verificationSheet) {
    verificationSheet = ss.insertSheet(SHEETS.VERIFICATION_CODES);
    verificationSheet.appendRow(['Email', 'Code', 'Data', 'Created', 'Expires', 'Verified']);
  }
  
  // Clean up old codes for this email
  const codes = verificationSheet.getDataRange().getValues();
  for (let i = codes.length - 1; i >= 1; i--) {
    if (codes[i][0] === data.email) {
      verificationSheet.deleteRow(i + 1);
    }
  }
  
  // Store new code with registration data
  verificationSheet.appendRow([
    data.email,
    verificationCode,
    JSON.stringify(data),
    timestamp.toISOString(),
    expiryTime.toISOString(),
    'false'
  ]);
  Logger.log('Verification code stored for ' + data.email + ' code=' + verificationCode);
  
  // Send verification email
  sendVerificationEmail(data.email, data.fullName, verificationCode);
  
  return {
    message: 'Verification code sent to your email',
    email: data.email
  };
}

/**
 * Verify code and complete registration
 */
function verifyCodeAndRegister(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const verificationSheet = ss.getSheetByName(SHEETS.VERIFICATION_CODES);
  
  if (!verificationSheet) {
    throw new Error('Verification system not initialized');
  }
  
  // Normalize inputs
  const emailInput = (data.email || '').toLowerCase().trim();
  const codeInput = (data.code || '').toString().trim();
  
  Logger.log('=== VERIFICATION DEBUG ===');
  Logger.log('Input Email: ' + emailInput);
  Logger.log('Input Code: ' + codeInput);
  
  const codes = verificationSheet.getDataRange().getValues();
  const now = new Date();
  
  Logger.log('Total rows: ' + (codes.length - 1));
  
  // Find matching code
  for (let i = 1; i < codes.length; i++) {
    const row = codes[i];
    const sheetEmail = (row[0] || '').toLowerCase().trim();
    const sheetCode = (row[1] || '').toString().trim();
    const sheetVerified = (row[5] || '').toString().toLowerCase().trim();
    
    Logger.log('Row ' + i + ' - Email: "' + sheetEmail + '", Code: "' + sheetCode + '", Verified: "' + sheetVerified + '"');
    
    // Match email and code (case-insensitive, trimmed)
    if (sheetEmail === emailInput && sheetCode === codeInput) {
      Logger.log('✓ Match found!');
      
      const expiryTime = new Date(row[4]);
      
      Logger.log('Expiry: ' + expiryTime.toISOString());
      Logger.log('Now: ' + now.toISOString());
      
      // Check if already verified
      if (sheetVerified === 'true') {
        Logger.log('✕ Already verified');
        throw new Error('Verification code already used');
      }
      
      // Check expiry
      if (now > expiryTime) {
        Logger.log('✕ Code expired');
        throw new Error('Verification code expired. Please request a new code.');
      }
      
      // Mark as verified
      verificationSheet.getRange(i + 1, 6).setValue('true');
      Logger.log('✓ Marked as verified');
      
      // Retrieve original registration data
      const registrationData = JSON.parse(row[2]);
      Logger.log('✓ Retrieved registration data');
      
      // Complete registration
      const result = registerUser(registrationData);
      Logger.log('✓ User registered');
      
      // Clean up verification code
      verificationSheet.deleteRow(i + 1);
      Logger.log('✓ Cleanup complete');
      
      return {
        message: 'Email verified successfully. Registration complete!',
        userId: result.userId,
        verified: true
      };
    }
  }
  
  Logger.log('✕ No match found');
  throw new Error('Invalid verification code');
}

function registerUser(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);
  
  // Check if email already exists
  const existingUsers = studentsSheet.getDataRange().getValues();
  for (let i = 1; i < existingUsers.length; i++) {
    if (existingUsers[i][3] === data.email) { // Assuming email is in column 4
      throw new Error('Email already registered');
    }
  }
  
  // Generate unique ID
  const userId = 'STU' + Date.now();
  const timestamp = new Date().toISOString();
  
  // Hash password (in production, use proper encryption)
  const hashedPassword = Utilities.base64Encode(data.password);
  
  // Add user to sheet
  studentsSheet.appendRow([
    userId,
    data.fullName,
    data.universityId,
    data.email,
    data.phone,
    data.department,
    data.level,
    data.year,
    data.semester,
    data.completedHours,
    data.gpa,
    data.advisorName,
    hashedPassword,
    data.role,
    data.status,
    timestamp,
    data.advisorId || '', // Column 17
    data.advisorEmail || '' // Column 18
  ]);
  
  // Send welcome email
  sendWelcomeEmail(data.email, data.fullName);
  
  return {
    message: 'Registration successful',
    userId: userId
  };
}

function loginUser(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hashedPassword = Utilities.base64Encode(data.password);
  
  Logger.log('=== LOGIN ATTEMPT ===');
  Logger.log('Email: ' + data.email);
  Logger.log('Password (original): ' + data.password);
  Logger.log('Password (hashed): ' + hashedPassword);
  Logger.log('Selected Role: ' + data.role);
  
  // Define column mapping for each user type
  const userTypes = [
    { 
      sheet: SHEETS.STUDENTS, 
      role: 'student',
      emailCol: 3,    // Column D (index 3)
      passwordCol: 12, // Column M (index 12)
      nameCol: 1,      // Column B (index 1)
      deptCol: 5,      // Column F (index 5)
      levelCol: 6,     // Column G (index 6)
      gpaCol: 10       // Column K (index 10)
    },
    { 
      sheet: SHEETS.ADVISORS, 
      role: 'advisor',
      emailCol: 2,     // Column C (index 2)
      passwordCol: 7,  // Column H (index 7)
      nameCol: 1,      // Column B (index 1)
      deptCol: 3,      // Column D (index 3)
      levelCol: null,
      gpaCol: null
    },
    { 
      sheet: SHEETS.HODS, 
      role: 'hod',
      emailCol: 3,     // Column D (index 3) per sheet layout
      passwordCol: 6,  // Column G (index 6) Base64 encoded
      nameCol: 1,      // Column B (index 1)
      deptCol: 5,      // Column F (index 5)
      levelCol: null,
      gpaCol: null
    },
    { 
      sheet: SHEETS.REGISTRARS, 
      role: 'registrar',
      emailCol: 2,     // Column C (index 2)
      passwordCol: 5,  // Column F (index 5)
      nameCol: 1,      // Column B (index 1)
      deptCol: 3,      // Column D (index 3)
      levelCol: null,
      gpaCol: null
    }
  ];
  
  // Only check the sheet matching the selected role
  for (const userType of userTypes) {
    // Skip if role doesn't match (if role is specified)
    if (data.role && userType.role !== data.role) {
      Logger.log('Skipping ' + userType.role + ' sheet (not selected role)');
      continue;
    }
    
    const sheet = ss.getSheetByName(userType.sheet);
    const users = sheet.getDataRange().getValues();
    
    Logger.log('--- Checking ' + userType.role + ' sheet ---');
    Logger.log('Total rows: ' + (users.length - 1));
    
    for (let i = 1; i < users.length; i++) {
      const row = users[i];
      
      Logger.log('Row ' + i + ' - Email in sheet: "' + row[userType.emailCol] + '"');
      Logger.log('Row ' + i + ' - Password in sheet: "' + row[userType.passwordCol] + '"');
      
      // Check email and password using the correct columns for this user type
      if (row[userType.emailCol] === data.email && row[userType.passwordCol] === hashedPassword) {
        Logger.log('✓ LOGIN SUCCESS for ' + userType.role);
        const userData = {
          id: row[0],
          fullName: row[userType.nameCol],
          email: row[userType.emailCol],
          role: userType.role,
          department: userType.deptCol !== null ? (row[userType.deptCol] || '') : '',
          level: userType.levelCol !== null ? (row[userType.levelCol] || '') : '',
          gpa: userType.gpaCol !== null ? (row[userType.gpaCol] || '') : ''
        };
        
        // Check for additional roles (for staff who have multiple roles)
        // If user is found in one sheet, check other staff sheets too
        const additionalRoles = [];
        if (userType.role !== 'student') {
          const staffTypes = [
            { sheet: SHEETS.ADVISORS, role: 'advisor', emailCol: 2, passwordCol: 7 },
            { sheet: SHEETS.HODS, role: 'hod', emailCol: 2, passwordCol: 6 },
            { sheet: SHEETS.REGISTRARS, role: 'registrar', emailCol: 2, passwordCol: 5 }
          ];
          
          for (const staffType of staffTypes) {
            if (staffType.role !== userType.role) {
              const staffSheet = ss.getSheetByName(staffType.sheet);
              const staffData = staffSheet.getDataRange().getValues();
              
              for (let j = 1; j < staffData.length; j++) {
                if (staffData[j][staffType.emailCol] === data.email && staffData[j][staffType.passwordCol] === hashedPassword) {
                  additionalRoles.push(staffType.role);
                  break;
                }
              }
            }
          }
        }
        
        if (additionalRoles.length > 0) {
          userData.additionalRoles = additionalRoles;
          userData.allRoles = [userType.role, ...additionalRoles];
        } else {
          userData.allRoles = [userType.role];
        }
        
        // Add advisor information for students
        if (userType.role === 'student') {
          userData.advisorName = row[11] || '';
          userData.advisorId = row[16] || '';
          userData.advisorEmail = row[17] || '';
        }
        
        return userData;
      }
    }
  }
  
  throw new Error('Invalid email or password');
}

function updateProfile(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Validate and convert data types
  if (data.phone && typeof data.phone === 'string') {
    data.phone = data.phone.trim();
  }
  if (data.completedHours && typeof data.completedHours === 'string') {
    data.completedHours = parseFloat(data.completedHours) || 0;
  }
  if (data.gpa && typeof data.gpa === 'string') {
    data.gpa = parseFloat(data.gpa) || 0;
  }

  // Define role-specific column mappings (0-based index for getValues array)
  const roleConfigs = {
    'student': {
      sheet: SHEETS.STUDENTS,
      phoneCol: 4,              // Column E (index 4)
      completedHoursCol: 9,     // Column J (index 9)
      gpaCol: 10                // Column K (index 10)
    },
    'advisor': {
      sheet: SHEETS.ADVISORS,
      phoneCol: 3,              // Column D (index 3)
      nameCol: 1                // Column B (index 1)
    },
    'hod': {
      sheet: SHEETS.HODS,
      phoneCol: 4,              // Column E (index 4)
      nameCol: 1                // Column B (index 1)
    },
    'registrar': {
      sheet: SHEETS.REGISTRARS,
      phoneCol: 3,              // Column D (index 3)
      nameCol: 1                // Column B (index 1)
    }
  };

  if (!data.role || !roleConfigs[data.role]) {
    throw new Error('Invalid user role');
  }

  const config = roleConfigs[data.role];
  const sheet = ss.getSheetByName(config.sheet);
  if (!sheet) {
    throw new Error('User sheet not found for role: ' + data.role);
  }

  const rows = sheet.getDataRange().getValues();
  let userRowIndex = -1;

  // Find user by ID (0-based column 0)
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.userId) {
      userRowIndex = i;
      break;
    }
  }

  if (userRowIndex < 0) {
    throw new Error('User not found');
  }

  // Prepare updates: collect ranges and values
  const ranges = [];
  const values = [];

  // Always update phone if provided
  if (config.phoneCol !== undefined && data.hasOwnProperty('phone') && data.phone !== undefined) {
    ranges.push(sheet.getRange(userRowIndex + 1, config.phoneCol + 1, 1, 1)); // +1 for 1-based column
    values.push(data.phone || '');
  }

  // Staff name update (allow advisors/HODs/registrars to edit name)
  if (data.role !== 'student' && config.nameCol !== undefined && data.hasOwnProperty('fullName') && data.fullName !== undefined) {
    ranges.push(sheet.getRange(userRowIndex + 1, config.nameCol + 1, 1, 1));
    values.push(data.fullName || '');
  }

  // Student-specific fields
  if (data.role === 'student') {
    if (config.completedHoursCol !== undefined && data.hasOwnProperty('completedHours') && data.completedHours !== undefined && data.completedHours !== '') {
      ranges.push(sheet.getRange(userRowIndex + 1, config.completedHoursCol + 1, 1, 1));
      values.push(data.completedHours);
    }
    if (config.gpaCol !== undefined && data.hasOwnProperty('gpa') && data.gpa !== undefined && data.gpa !== '') {
      ranges.push(sheet.getRange(userRowIndex + 1, config.gpaCol + 1, 1, 1));
      values.push(data.gpa);
    }
  }

  // Apply all updates
  for (let j = 0; j < ranges.length; j++) {
    ranges[j].setValue(values[j]);
  }

  return { message: 'Profile updated successfully' };
}

/**
 * Send a 6-digit reset code to the user's email
 */
function sendPasswordResetCode(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const user = findUserByEmail(data.email);
  if (!user) {
    throw new Error('Email not found');
  }

  const resetSheetName = SHEETS.RESET_CODES;
  let resetSheet = ss.getSheetByName(resetSheetName) || ss.getSheetByName('Password Reset Codes');
  if (!resetSheet) {
    resetSheet = ss.insertSheet(resetSheetName);
    resetSheet.appendRow(['Email', 'Code', 'UserId', 'Role', 'Created', 'Expires', 'Used']);
  }

  // Clear old codes for this email
  const rows = resetSheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if ((rows[i][0] || '').toLowerCase().trim() === (data.email || '').toLowerCase().trim()) {
      resetSheet.deleteRow(i + 1);
    }
  }

  const code = generateVerificationCode();
  const now = new Date();
  const expires = new Date(now.getTime() + 15 * 60 * 1000);

  resetSheet.appendRow([
    data.email,
    code,
    user.userId,
    user.role,
    now.toISOString(),
    expires.toISOString(),
    'false'
  ]);

  sendPasswordResetCodeEmail(data.email, user.fullName || user.role, code);
  Logger.log('Password reset code stored for ' + data.email + ' code=' + code);

  return { message: 'Reset code sent to your email' };
}

/**
 * Verify reset code and set new password
 */
function verifyPasswordResetCode(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resetSheetName = SHEETS.RESET_CODES;
  const resetSheet = ss.getSheetByName(resetSheetName) || ss.getSheetByName('Password Reset Codes');
  if (!resetSheet) {
    throw new Error('No reset request found. Please request a new code.');
  }

  const emailLower = (data.email || '').toLowerCase().trim();
  const codeInput = (data.code || '').trim();
  const rows = resetSheet.getDataRange().getValues();
  let matchRow = -1;
  for (let i = 1; i < rows.length; i++) {
    const [email, code, userId, role, created, expires, used] = rows[i];
    if ((email || '').toLowerCase().trim() === emailLower && (code || '').toString().trim() === codeInput) {
      if ((used || '').toString().toLowerCase() === 'true') {
        throw new Error('This code has already been used.');
      }
      if (new Date(expires) < new Date()) {
        throw new Error('This code has expired. Please request a new one.');
      }
      matchRow = i + 1; // sheet row index (1-based)
      break;
    }
  }

  if (matchRow === -1) {
    throw new Error('Invalid code or email.');
  }

  const user = findUserByEmail(data.email);
  if (!user) {
    throw new Error('User not found');
  }

  if (!data.newPassword) {
    throw new Error('New password is required');
  }

  const hashedPassword = Utilities.base64Encode(data.newPassword);
  user.sheet.getRange(user.rowIndex + 1, user.passwordCol + 1).setValue(hashedPassword);

  // Mark code as used
  resetSheet.getRange(matchRow, 7).setValue('true');

  return { message: 'Password updated successfully. You can now log in.' };
}

function resetPassword(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Generate temporary password
  const tempPassword = generateRandomPassword();
  const hashedPassword = Utilities.base64Encode(tempPassword);
  
  // Update password in sheet (check all user types)
  const userTypes = [SHEETS.STUDENTS, SHEETS.ADVISORS, SHEETS.HODS, SHEETS.REGISTRARS];
  
  for (const sheetName of userTypes) {
    const sheet = ss.getSheetByName(sheetName);
    const users = sheet.getDataRange().getValues();
    
    for (let i = 1; i < users.length; i++) {
      if (users[i][3] === data.email) {
        sheet.getRange(i + 1, 13).setValue(hashedPassword);
        
        // Send reset email
        sendPasswordResetEmail(data.email, users[i][1], tempPassword);
        
        return { message: 'Password reset link sent to your email' };
      }
    }
  }
  
  throw new Error('Email not found');
}

function changePassword(data) {
  const user = findUserByEmail(data.email);
  if (!user) {
    throw new Error('User not found');
  }

  const oldHashedPassword = Utilities.base64Encode(data.oldPassword);
  const newHashedPassword = Utilities.base64Encode(data.newPassword);

  // Read current password from the correct column
  const currentPassword = user.sheet.getRange(user.rowIndex + 1, user.passwordCol + 1).getValue();
  if (currentPassword !== oldHashedPassword) {
    throw new Error('Invalid old password');
  }

  // Update password in place
  user.sheet.getRange(user.rowIndex + 1, user.passwordCol + 1).setValue(newHashedPassword);
  return { message: 'Password changed successfully' };
}

// ===================================
// COURSE MANAGEMENT
// ===================================

function getCourses() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const coursesSheet = ss.getSheetByName(SHEETS.COURSES);
  const data = coursesSheet.getDataRange().getValues();
  
  const courses = [];
  for (let i = 1; i < data.length; i++) {
    // Each row is a specific section of a course activity
    courses.push({
      code: data[i][0],
      name: data[i][1],
      activityType: data[i][2],
      creditHours: data[i][3],
      section: data[i][4],
      day: data[i][5],
      time: data[i][6]
    });
  }
  
  return courses;
}

function getAdvisors() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const advisorsSheet = ss.getSheetByName(SHEETS.ADVISORS);
  const data = advisorsSheet.getDataRange().getValues();
  
  const advisors = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // Check if there's an ID
      advisors.push({
        id: data[i][0],
        name: data[i][1],
        email: data[i][2],
        department: data[i][3],
        status: data[i][4]
      });
    }
  }
  
  return advisors;
}

// ===================================
// REQUEST MANAGEMENT
// ===================================

function submitRequest(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const requestsSheet = ss.getSheetByName(SHEETS.REQUESTS);
  
  // Generate request ID
  const requestId = 'REQ' + Date.now();
  const timestamp = new Date().toISOString();
  
  // Handle multiple courses
  const coursesToAdd = data.coursesToAdd || [];
  const coursesToDrop = data.coursesToDrop || [];
  
  // Create summary strings
  const addedCoursesStr = coursesToAdd.map(c => `${c.code} (${c.activityType} - Section ${c.section})`).join('; ') || '';
  const droppedCoursesStr = coursesToDrop.join('; ') || '';
  
  // Add request
  requestsSheet.appendRow([
    requestId,
    data.studentId,
    data.studentName,
    data.studentEmail,
    data.requestType,
    addedCoursesStr, // Multiple courses as comma-separated string
    droppedCoursesStr, // Multiple courses as comma-separated string
    '', // sectionToAdd (not used for multiple)
    '', // courseToDrop (not used for multiple)
    data.reason,
    'awaiting_advisor',
    timestamp,
    '', // advisor approval
    '', // hod approval
    '', // registrar approval
    '', // pdf url
    JSON.stringify({
      coursesToAdd: coursesToAdd,
      coursesToDrop: coursesToDrop
    }),
    data.advisorId || '', // Advisor ID - column 18
    data.advisorName || '', // Advisor Name - column 19
    data.advisorEmail || '' // Advisor Email - column 20
  ]);
  
  // Log approval start
  logApproval(requestId, 'submitted', 'student', data.studentName, 'Request submitted');
  
  // Send notifications (email functions are stubs - just logging for now)
  Logger.log('Sending notifications for request ' + requestId);
  
  return {
    message: 'Request submitted successfully',
    requestId: requestId
  };
}

function getRequests(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const requestsSheet = ss.getSheetByName(SHEETS.REQUESTS);
  const requestsData = requestsSheet.getDataRange().getValues();
  const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);
  const studentsData = studentsSheet.getDataRange().getValues();
  const approvalsSheet = ss.getSheetByName(SHEETS.APPROVALS);
  const approvalsData = approvalsSheet ? approvalsSheet.getDataRange().getValues() : [];
  
  // Create a map of approvals (comments) for quick lookup by requestId and role
  const approvalMap = {};
  for (let i = 1; i < approvalsData.length; i++) {
    const requestId = approvalsData[i][0];
    const role = approvalsData[i][1];
    const comments = approvalsData[i][4] || '';
    const key = `${requestId}_${role}`;
    approvalMap[key] = comments;
  }
  
  // Determine HOD department (for scoping approved/rejected to the HOD's department)
  let hodDepartment = '';
  if (data.userRole === 'hod') {
    const hodsSheet = ss.getSheetByName(SHEETS.HODS);
    if (hodsSheet) {
      const hodsData = hodsSheet.getDataRange().getValues();
      for (let i = 1; i < hodsData.length; i++) {
        // Match by HOD ID (column A)
        if (hodsData[i][0] === data.userId) {
          // Department column F (index 5)
          hodDepartment = hodsData[i][5] || '';
          break;
        }
      }
    }
  }

  // Determine Registrar department (to scope registrar-visible requests to their department)
  let registrarDepartment = '';
  if (data.userRole === 'registrar') {
    const registrarsSheet = ss.getSheetByName(SHEETS.REGISTRARS);
    if (registrarsSheet) {
      const registrarsData = registrarsSheet.getDataRange().getValues();
      for (let i = 1; i < registrarsData.length; i++) {
        // Match by Registrar ID (column A)
        if (registrarsData[i][0] === data.userId) {
          // Department column D (index 3)
          registrarDepartment = registrarsData[i][3] || '';
          break;
        }
      }
    }
  }
  
  // Create a map of student data for quick lookup
  const studentMap = {};
  for (let i = 1; i < studentsData.length; i++) {
    studentMap[studentsData[i][0]] = {
      level: studentsData[i][6] || '', // Column G (index 6)
      gpa: studentsData[i][10] || '', // Column K (index 10)
      department: studentsData[i][5] || '', // Column F (index 5)
      completedHours: studentsData[i][9] || '' // Column J (index 9)
    };
  }
  
  const requests = [];
  
  for (let i = 1; i < requestsData.length; i++) {
    const row = requestsData[i];
    
    // Filter based on user role
    let includeRequest = false;
    
    if (data.userRole === 'student' && row[1] === data.userId) {
      includeRequest = true;
    } else if (data.userRole === 'advisor' && row[17] === data.userId) {
      // Advisor sees all requests for their assigned students (column 18 is advisorId) regardless of status
      includeRequest = true;
    } else if (data.userRole === 'hod') {
      // HOD sees:
      // - awaiting_hod (pending reviews)
      // - hodApproval approved/rejected (their reviewed items), regardless of current status
      const studentId = row[1];
      const studentDept = studentMap[studentId] ? (studentMap[studentId].department || '') : '';
      const isAwaiting = row[10] === 'awaiting_hod';
      const isReviewedByHOD = (row[13] === 'approved' || row[13] === 'rejected');
      // If HOD department is known, scope to that department; otherwise include all
      const departmentMatches = !hodDepartment || (studentDept === hodDepartment);
      if ((isAwaiting || isReviewedByHOD) && departmentMatches) {
        includeRequest = true;
      }
    } else if (data.userRole === 'registrar' && (row[10] === 'awaiting_registrar' || row[10] === 'completed' || row[10] === 'rejected')) {
      // Registrar sees only requests for their department
      const studentId = row[1];
      const studentDept = studentMap[studentId] ? (studentMap[studentId].department || '') : '';
      const departmentMatches = !registrarDepartment || (studentDept === registrarDepartment);
      if (departmentMatches) {
        includeRequest = true;
      }
    }
    
    if (includeRequest) {
      // Get student details from studentMap
      const studentId = row[1];
      const studentDetails = studentMap[studentId] || {};
      const requestId = row[0];
      
      // Get advisor comments from approvals map
      const advisorCommentsKey = `${requestId}_advisor`;
      const advisorComments = approvalMap[advisorCommentsKey] || '';
      
      requests.push({
        id: row[0],
        studentId: row[1],
        studentName: row[2],
        studentEmail: row[3],
        type: row[4],
        courseCode: row[5],
        courseName: row[6],
        sectionToAdd: row[7],
        courseToDrop: row[8],
        reason: row[9],
        status: row[10],
        submittedDate: row[11],
        advisorApproval: row[12],
        hodApproval: row[13],
        registrarApproval: row[14],
        pdfUrl: row[15],
        courseDetails: row[16] ? JSON.parse(row[16]) : {},
        advisorId: row[17],
        advisorName: row[18],
        advisorEmail: row[19],
        advisorComments: advisorComments,
        level: studentDetails.level || 'N/A',
        gpa: studentDetails.gpa || 'N/A',
        department: studentDetails.department || 'N/A',
        completedHours: studentDetails.completedHours || 'N/A'
      });
    }
  }
  
  return requests;
}

function updateRequestStatus(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const requestsSheet = ss.getSheetByName(SHEETS.REQUESTS);
  const requests = requestsSheet.getDataRange().getValues();
  const approverRole = data.approverRole;
  const desiredStatus = data.status;

  // Enforce allowed status transitions per role
  const allowedStatusByRole = {
    advisor: ['awaiting_hod', 'rejected'],
    hod: ['awaiting_registrar', 'rejected'],
    registrar: ['completed', 'rejected']
  };

  if (!allowedStatusByRole[approverRole] || !allowedStatusByRole[approverRole].includes(desiredStatus)) {
    throw new Error('Unauthorized status transition for role ' + approverRole);
  }
  
  for (let i = 1; i < requests.length; i++) {
    if (requests[i][0] === data.requestId) {
      const rowNum = i + 1;
      const request = requests[i];
      const currentStatus = request[10]; // Column K (index 10)

      // Prevent duplicate/invalid transitions
      if (currentStatus === desiredStatus) {
        return { message: 'Status already set to ' + desiredStatus };
      }
      if (currentStatus === 'completed') {
        throw new Error('Request already completed');
      }
      
      // Determine student's department from Students sheet (Column F / index 5)
      const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);
      const studentsData = studentsSheet.getDataRange().getValues();
      const studentId = request[1];
      let studentDepartment = '';
      for (let s = 1; s < studentsData.length; s++) {
        if (studentsData[s][0] === studentId) {
          studentDepartment = studentsData[s][5] || '';
          break;
        }
      }
      
      // Enforce department scoping for HOD and Registrar approvals
      if (data.approverRole === 'hod') {
        const hodsSheet = ss.getSheetByName(SHEETS.HODS);
        const hodsData = hodsSheet.getDataRange().getValues();
        let hodDepartment = '';
        for (let h = 1; h < hodsData.length; h++) {
          if (hodsData[h][0] === data.approverId) {
            // Department column F (index 5)
            hodDepartment = hodsData[h][5] || '';
            break;
          }
        }
        if (studentDepartment && hodDepartment && studentDepartment !== hodDepartment) {
          throw new Error('Unauthorized: HOD department mismatch for this student');
        }
      } else if (data.approverRole === 'registrar') {
        const registrarsSheet = ss.getSheetByName(SHEETS.REGISTRARS);
        const registrarsData = registrarsSheet.getDataRange().getValues();
        let registrarDepartment = '';
        for (let r = 1; r < registrarsData.length; r++) {
          if (registrarsData[r][0] === data.approverId) {
            // Department column D (index 3)
            registrarDepartment = registrarsData[r][3] || '';
            break;
          }
        }
        if (studentDepartment && registrarDepartment && studentDepartment !== registrarDepartment) {
          throw new Error('Unauthorized: Registrar department mismatch for this student');
        }
      }
      
      // Update status
      requestsSheet.getRange(rowNum, 11).setValue(desiredStatus);
      
      // Update approval columns based on role
      if (approverRole === 'advisor') {
        requestsSheet.getRange(rowNum, 13).setValue(desiredStatus === 'rejected' ? 'rejected' : 'approved');
      } else if (approverRole === 'hod') {
        requestsSheet.getRange(rowNum, 14).setValue(desiredStatus === 'rejected' ? 'rejected' : 'approved');
      } else if (approverRole === 'registrar') {
        requestsSheet.getRange(rowNum, 15).setValue(desiredStatus === 'rejected' ? 'rejected' : 'approved');
        
        // Generate PDF if completed
        if (desiredStatus === 'completed') {
          let pdfUrl = '';
          try {
            pdfUrl = generateApprovalPDF(request, data);
            requestsSheet.getRange(rowNum, 16).setValue(pdfUrl);
          } catch (pdfErr) {
            logError('generateApprovalPDF request ' + data.requestId, pdfErr);
            return { message: 'Status updated; PDF generation failed. Registrar can retry.' };
          }

          // Send final email with PDF (non-blocking)
          try {
            sendCompletionEmail(request, pdfUrl);
          } catch (emailErr) {
            logError('sendCompletionEmail request ' + data.requestId, emailErr);
          }
        }
      }
      
      // Log approval
      logApproval(data.requestId, desiredStatus, approverRole, data.approverName, data.comments);
      
      // Send notification emails
      sendStatusUpdateEmails(request, data);
      
      return { message: 'Status updated successfully' };
    }
  }
  
  throw new Error('Request not found');
}

function logApproval(requestId, status, role, approverName, comments) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const approvalsSheet = ss.getSheetByName(SHEETS.APPROVALS);
  const timestamp = new Date().toISOString();
  
  approvalsSheet.appendRow([
    requestId,
    role,
    approverName,
    status,
    comments || '',
    timestamp
  ]);
}

function logError(context, error) {
  try {
    Logger.log('ERROR [' + context + ']: ' + error);
    Logger.log(error && error.stack ? error.stack : 'no stack');
  } catch (e) {
    Logger.log('Failed to log error: ' + e);
  }
}

// ===================================
// HELPER FUNCTIONS
// ===================================

function generateRandomPassword() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let password = '';
  for (let i = 0; i < 12; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return password;
}

function getPDFUrl(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const requestsSheet = ss.getSheetByName(SHEETS.REQUESTS);
  const requests = requestsSheet.getDataRange().getValues();
  
  for (let i = 1; i < requests.length; i++) {
    if (requests[i][0] === data.requestId) {
      return { pdfUrl: requests[i][15] };
    }
  }
  
  throw new Error('PDF not found');
}
