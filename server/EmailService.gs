/**
 * Course Add/Drop Management System
 * Email Notification Module
 */

// ===================================
// EMAIL TEMPLATES
// ===================================

const EMAIL_CONFIG = {
  FROM_NAME: 'Course Add/Drop System - CCIT',
  SIGNATURE: `
    <br><br>
    <p style="color: #7f8c8d; font-size: 0.9em;">
      This is an automated email from the Course Add/Drop Management System<br>
      College of Computing and Information Technology<br>
      Shaqra University, Saudi Arabia<br>
      <br>
      Please do not reply to this email.
    </p>
  `
};

// ===================================
// WELCOME & AUTHENTICATION EMAILS
// ===================================

/**
 * Send verification code email
 */
function sendVerificationEmail(email, fullName, verificationCode) {
  try {
    const subject = 'Verify Your Email - Course Add/Drop System';
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
          .container { max-width: 600px; margin: 0 auto; padding: 20px; }
          .header { background: linear-gradient(135deg, #014b3a 0%, #02685a 100%); color: white; padding: 30px; text-align: center; border-radius: 8px 8px 0 0; }
          .content { background: #f8f9fa; padding: 30px; border-radius: 0 0 8px 8px; }
          .code-box { background: white; border: 3px dashed #014b3a; padding: 20px; margin: 20px 0; text-align: center; border-radius: 8px; }
          .code { font-size: 32px; font-weight: bold; color: #014b3a; letter-spacing: 8px; font-family: 'Courier New', monospace; }
          .warning { background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; border-radius: 4px; }
          .footer { text-align: center; padding: 20px; color: #666; font-size: 12px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>📚 Email Verification</h1>
            <p>Shaqra University - CCIT</p>
          </div>
          <div class="content">
            <h2>Hello ${fullName},</h2>
            <p>Thank you for registering for the Course Add/Drop Management System!</p>
            <p>To complete your registration, please use the verification code below:</p>
            
            <div class="code-box">
              <div style="color: #666; font-size: 14px; margin-bottom: 10px;">Your Verification Code</div>
              <div class="code">${verificationCode}</div>
            </div>
            
            <div class="warning">
              <strong>⚠️ Important:</strong>
              <ul style="margin: 10px 0; padding-left: 20px;">
                <li>This code will expire in <strong>15 minutes</strong></li>
                <li>Do not share this code with anyone</li>
                <li>If you didn't request this code, please ignore this email</li>
              </ul>
            </div>
            
            <p>Enter this code on the registration page to verify your email address and activate your account.</p>
            
            <p style="color: #666; font-size: 14px; margin-top: 30px;">
              Having trouble? Contact us at support@ccit.shaqra.edu.sa
            </p>
          </div>
          <div class="footer">
            <p>Course Add/Drop Management System</p>
            <p>College of Computing and Information Technology</p>
            <p>Shaqra University</p>
            <p style="margin-top: 10px; color: #999;">This is an automated message, please do not reply to this email.</p>
          </div>
        </div>
      </body>
      </html>
    `;
    
    sendEmailSafe({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    }, 'sendVerificationCodeEmail');
    
    Logger.log('Verification code sent to: ' + email);
  } catch (error) {
    Logger.log('Error sending verification email: ' + error.toString());
    throw new Error('Failed to send verification email. Please try again.');
  }
}

function sendWelcomeEmail(email, name) {
  const subject = 'Welcome to Course Add/Drop System';
  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
          <h2 style="color: #3498db; border-bottom: 2px solid #3498db; padding-bottom: 10px;">
            Welcome to Course Add/Drop System
          </h2>
          
          <p>Dear ${name},</p>
          
          <p>Your account has been successfully created in the Course Add/Drop Management System.</p>
          
          <p><strong>What's Next?</strong></p>
          <ul>
            <li>Log in to your account using your registered email and password</li>
            <li>Complete your profile information</li>
            <li>Submit course add/drop requests when needed</li>
            <li>Track your request status in real-time</li>
          </ul>
          
          <div style="background: #ecf0f1; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p style="margin: 0;"><strong>📧 Your Email:</strong> ${email}</p>
            <p style="margin: 10px 0 0 0;"><strong>🔗 Login URL:</strong> 
              <a href="https://ccitsu.github.io/dacs/index.html" style="color: #3498db;">Click here to login</a>
            </p>
          </div>
          
          <p>If you have any questions or need assistance, please contact the CCIT support team.</p>
          
          ${EMAIL_CONFIG.SIGNATURE}
        </div>
      </body>
    </html>
  `;
  
  sendEmailSafe({
    to: email,
    subject: subject,
    htmlBody: body
  }, 'sendWelcomeEmail');
}

function sendPasswordResetEmail(email, name, tempPassword) {
  const subject = 'Password Reset - Course Add/Drop System';
  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
          <h2 style="color: #e74c3c; border-bottom: 2px solid #e74c3c; padding-bottom: 10px;">
            Password Reset Request
          </h2>
          
          <p>Dear ${name},</p>
          
          <p>We received a request to reset your password. Your temporary password is:</p>
          
          <div style="background: #fff3cd; padding: 20px; border-radius: 5px; margin: 20px 0; text-align: center;">
            <p style="margin: 0; font-size: 1.5em; font-weight: bold; color: #856404;">${tempPassword}</p>
          </div>
          
          <p><strong>⚠️ Important:</strong></p>
          <ul>
            <li>Please login and change your password immediately</li>
            <li>This temporary password will expire in 24 hours</li>
            <li>Do not share this password with anyone</li>
          </ul>
          
          <p>
            <a href="https://ccitsu.github.io/dacs/pages/login.html" 
               style="display: inline-block; background: #3498db; color: white; padding: 12px 30px; 
                      text-decoration: none; border-radius: 5px; margin-top: 10px;">
              Login Now
            </a>
          </p>
          
          <p style="margin-top: 20px; color: #7f8c8d; font-size: 0.9em;">
            If you didn't request this password reset, please ignore this email or contact support if you're concerned.
          </p>
          
          ${EMAIL_CONFIG.SIGNATURE}
        </div>
      </body>
    </html>
  `;
  
  sendEmailSafe({
    to: email,
    subject: subject,
    htmlBody: body
  }, 'sendPasswordResetEmail');
}

/**
 * Send code-based password reset email
 */
function sendPasswordResetCodeEmail(email, name, code) {
  const subject = 'Password Reset Code - Course Add/Drop System';
  const htmlBody = `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50; }
          .container { max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px; }
          .header { background: linear-gradient(135deg, #0b5e55 0%, #0f7f72 100%); color: white; padding: 24px; border-radius: 8px 8px 0 0; text-align: center; }
          .code-box { background: #f8f9fa; border: 3px dashed #0f7f72; padding: 20px; text-align: center; border-radius: 8px; margin: 24px 0; }
          .code { font-size: 32px; font-weight: bold; letter-spacing: 8px; font-family: 'Courier New', monospace; color: #0f7f72; }
          .warning { background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; border-radius: 4px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h2>Password Reset Code</h2>
            <p>Course Add/Drop System</p>
          </div>
          <div style="padding: 24px;">
            <p>Dear ${name || 'user'},</p>
            <p>We received a request to reset the password for your account. Use the code below to continue:</p>
            <div class="code-box">
              <div class="code">${code}</div>
            </div>
            <div class="warning">
              <strong>Important:</strong>
              <ul style="margin: 10px 0 0 20px;">
                <li>The code expires in 15 minutes</li>
                <li>Do not share this code with anyone</li>
                <li>If you did not request this, ignore this email</li>
              </ul>
            </div>
            <p style="margin-top: 20px;">Enter this code on the reset page to set a new password.</p>
            ${EMAIL_CONFIG.SIGNATURE}
          </div>
        </div>
      </body>
    </html>
  `;
  
  sendEmailSafe({
    to: email,
    subject: subject,
    htmlBody: htmlBody
  }, 'sendPasswordResetCodeEmail');
}

// ===================================
// REQUEST SUBMISSION EMAILS
// ===================================

function sendStudentSubmissionEmail(email, name, requestId) {
  const subject = `Request Submitted - #${requestId}`;
  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
          <h2 style="color: #27ae60; border-bottom: 2px solid #27ae60; padding-bottom: 10px;">
            ✓ Request Submitted Successfully
          </h2>
          
          <p>Dear ${name},</p>
          
          <p>Your course add/drop request has been submitted successfully and is now being processed.</p>
          
          <div style="background: #d4edda; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p style="margin: 0;"><strong>Request ID:</strong> #${requestId}</p>
            <p style="margin: 10px 0 0 0;"><strong>Status:</strong> Awaiting Advisor Review</p>
          </div>
          
          <p><strong>Next Steps:</strong></p>
          <ol>
            <li>Your academic advisor will review your request</li>
            <li>If approved, it will go to the Head of Department</li>
            <li>Final approval will be processed by the Registrar</li>
            <li>You'll receive email updates at each stage</li>
          </ol>
          
          <p>You can track your request status anytime by logging into your dashboard.</p>
          
          <p>
            <a href="https://ccitsu.github.io/dacs/pages/student-dashboard.html" 
               style="display: inline-block; background: #3498db; color: white; padding: 12px 30px; 
                      text-decoration: none; border-radius: 5px; margin-top: 10px;">
              View Dashboard
            </a>
          </p>
          
          ${EMAIL_CONFIG.SIGNATURE}
        </div>
      </body>
    </html>
  `;
  
  sendEmailSafe({
    to: email,
    subject: subject,
    htmlBody: body
  }, 'sendStudentConfirmationEmail');
}

function sendAdvisorNotificationEmail(advisorEmail, studentName, requestId) {
  if (!advisorEmail) return; // Skip if no advisor email
  
  const subject = `New Request for Review - ${studentName}`;
  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
          <h2 style="color: #f39c12; border-bottom: 2px solid #f39c12; padding-bottom: 10px;">
            ⏳ New Request Awaiting Your Review
          </h2>
          
          <p>Dear Advisor,</p>
          
          <p>A new course add/drop request has been submitted by your advisee and requires your review.</p>
          
          <div style="background: #fff3cd; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p style="margin: 0;"><strong>Student:</strong> ${studentName}</p>
            <p style="margin: 10px 0 0 0;"><strong>Request ID:</strong> #${requestId}</p>
            <p style="margin: 10px 0 0 0;"><strong>Status:</strong> Awaiting Your Review</p>
          </div>
          
          <p>Please log in to your dashboard to review and take action on this request.</p>
          
          <p>
            <a href="https://ccitsu.github.io/dacs/pages/advisor-dashboard.html" 
               style="display: inline-block; background: #f39c12; color: white; padding: 12px 30px; 
                      text-decoration: none; border-radius: 5px; margin-top: 10px;">
              Review Request
            </a>
          </p>
          
          ${EMAIL_CONFIG.SIGNATURE}
        </div>
      </body>
    </html>
  `;
  
  sendEmailSafe({
    to: advisorEmail,
    subject: subject,
    htmlBody: body
  }, 'sendAdvisorNotificationEmail');
}

// ===================================
// STATUS UPDATE EMAILS
// ===================================

function sendEmailSafe(payload, context) {
  try {
    MailApp.sendEmail(payload);
  } catch (err) {
    Logger.log('Email send failed [' + context + ']: ' + err);
  }
}

function sendStatusUpdateEmails(request, updateData) {
  const requestId = request[0];
  const studentName = request[2];
  const studentEmail = request[3];
  const status = updateData.status;
  const approverRole = updateData.approverRole;
  const comments = updateData.comments || 'No additional comments';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Resolve approver email safely to avoid invalid email errors
  let approverEmail = '';
  try {
    if (approverRole === 'advisor') {
      // Prefer advisor email stored on the request
      approverEmail = (request[19] || '').toString().trim(); // Column T (index 19)
      // Fallback: Advisors sheet lookup by ID (Column A = ID, Column C = Email)
      if (!approverEmail && updateData.approverId) {
        const advisorsSheet = ss.getSheetByName(SHEETS.ADVISORS);
        if (advisorsSheet) {
          const advisorsData = advisorsSheet.getDataRange().getValues();
          for (let i = 1; i < advisorsData.length; i++) {
            if (advisorsData[i][0] === updateData.approverId) {
              approverEmail = (advisorsData[i][2] || '').toString().trim();
              break;
            }
          }
        }
      }
    } else if (approverRole === 'hod') {
      // HOD email from HODs sheet by ID (Column A = ID, Column D = Email)
      if (updateData.approverId) {
        const hodsSheet = ss.getSheetByName(SHEETS.HODS);
        if (hodsSheet) {
          const hodsData = hodsSheet.getDataRange().getValues();
          for (let i = 1; i < hodsData.length; i++) {
            if (hodsData[i][0] === updateData.approverId) {
              approverEmail = (hodsData[i][3] || '').toString().trim();
              break;
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error resolving approver email: ' + e);
  }
  
  // Email to student
  sendStudentStatusEmail(studentEmail, studentName, requestId, status, approverRole, comments);
  
  // Email to next approver
  if (status === 'awaiting_hod') {
    sendHODNotificationEmail(requestId, studentName);
  } else if (status === 'awaiting_registrar') {
    sendRegistrarNotificationEmail(requestId, studentName);
  }
  
  // Email to previous approver confirming their action
  const isValidEmail = (email) => !!email && email.indexOf('@') > 0;
  if (approverRole === 'advisor' && status !== 'rejected') {
    if (isValidEmail(approverEmail)) {
      sendApproverConfirmationEmail(approverEmail, 'Advisor', requestId, 'approved');
    } else {
      Logger.log('Skipping advisor confirmation email due to invalid email: ' + approverEmail);
    }
  } else if (approverRole === 'hod' && status !== 'rejected') {
    if (isValidEmail(approverEmail)) {
      sendApproverConfirmationEmail(approverEmail, 'Head of Department', requestId, 'approved');
    } else {
      Logger.log('Skipping HOD confirmation email due to invalid email: ' + approverEmail);
    }
  }
}

function sendStudentStatusEmail(email, name, requestId, status, approverRole, comments) {
  let statusText = '';
  let statusColor = '#3498db';
  let nextStep = '';
  
  if (status === 'awaiting_hod') {
    statusText = 'Approved by Advisor';
    statusColor = '#27ae60';
    nextStep = 'Your request has been approved by your academic advisor and is now being reviewed by the Head of Department.';
  } else if (status === 'awaiting_registrar') {
    statusText = 'Approved by Head of Department';
    statusColor = '#27ae60';
    nextStep = 'Your request has been approved by the Head of Department and is now awaiting final processing by the Registrar.';
  } else if (status === 'completed') {
    statusText = 'Completed';
    statusColor = '#27ae60';
    nextStep = 'Your request has been completed successfully! The final approval document has been attached.';
  } else if (status === 'rejected') {
    statusText = 'Rejected';
    statusColor = '#e74c3c';
    nextStep = `Your request was not approved by the ${approverRole}. Please review the comments below and contact your advisor if you have questions.`;
  }
  
  const subject = `Request ${statusText} - #${requestId}`;
  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
          <h2 style="color: ${statusColor}; border-bottom: 2px solid ${statusColor}; padding-bottom: 10px;">
            Request Status Updated
          </h2>
          
          <p>Dear ${name},</p>
          
          <p>Your course add/drop request has been updated.</p>
          
          <div style="background: #ecf0f1; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p style="margin: 0;"><strong>Request ID:</strong> #${requestId}</p>
            <p style="margin: 10px 0 0 0;"><strong>New Status:</strong> ${statusText}</p>
            <p style="margin: 10px 0 0 0;"><strong>Updated By:</strong> ${approverRole.charAt(0).toUpperCase() + approverRole.slice(1)}</p>
          </div>
          
          <p>${nextStep}</p>
          
          ${comments !== 'No additional comments' ? `
            <div style="background: #fff3cd; padding: 15px; border-radius: 5px; margin: 20px 0;">
              <p style="margin: 0;"><strong>Comments:</strong></p>
              <p style="margin: 10px 0 0 0;">${comments}</p>
            </div>
          ` : ''}
          
          <p>
            <a href="https://ccitsu.github.io/dacs/pages/student-dashboard.html" 
               style="display: inline-block; background: #3498db; color: white; padding: 12px 30px; 
                      text-decoration: none; border-radius: 5px; margin-top: 10px;">
              View Dashboard
            </a>
          </p>
          
          ${EMAIL_CONFIG.SIGNATURE}
        </div>
      </body>
    </html>
  `;
  
  sendEmailSafe({
    to: email,
    subject: subject,
    htmlBody: body
  }, 'sendStudentStatusEmail');
}

function sendHODNotificationEmail(requestId, studentName) {
  // Resolve HOD email by student's department via requestId → studentId → Students → HODs
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let hodEmail = '';
  try {
    const requestsSheet = ss.getSheetByName(SHEETS.REQUESTS);
    const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);
    const hodsSheet = ss.getSheetByName(SHEETS.HODS);
    if (requestsSheet && studentsSheet && hodsSheet) {
      const reqData = requestsSheet.getDataRange().getValues();
      let studentId = '';
      for (let i = 1; i < reqData.length; i++) {
        if (reqData[i][0] === requestId) { // Column A (index 0) = Request ID
          studentId = reqData[i][1]; // Column B (index 1) = Student ID
          break;
        }
      }
      if (studentId) {
        const studentsData = studentsSheet.getDataRange().getValues();
        let department = '';
        for (let s = 1; s < studentsData.length; s++) {
          if (studentsData[s][0] === studentId) { // Column A (index 0) = Student ID
            department = studentsData[s][5] || ''; // Column F (index 5) = Department
            break;
          }
        }
        if (department) {
          const hodsData = hodsSheet.getDataRange().getValues();
          for (let h = 1; h < hodsData.length; h++) {
            if ((hodsData[h][5] || '') === department) { // Column F (index 5) = Department
              hodEmail = (hodsData[h][3] || '').toString().trim(); // Column D (index 3) = HOD email
              break;
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('sendHODNotificationEmail resolution error: ' + e);
  }
  if (!hodEmail || hodEmail.indexOf('@') < 1) {
    Logger.log('sendHODNotificationEmail skipped: no valid HOD email resolved for request ' + requestId);
    return;
  }
  
  const subject = `Request Awaiting Review - ${studentName}`;
  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
          <h2 style="color: #f39c12; border-bottom: 2px solid #f39c12; padding-bottom: 10px;">
            ⏳ Request Awaiting Your Review
          </h2>
          
          <p>Dear Head of Department,</p>
          
          <p>A course add/drop request has been approved by the academic advisor and now requires your review.</p>
          
          <div style="background: #fff3cd; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p style="margin: 0;"><strong>Student:</strong> ${studentName}</p>
            <p style="margin: 10px 0 0 0;"><strong>Request ID:</strong> #${requestId}</p>
            <p style="margin: 10px 0 0 0;"><strong>Status:</strong> Awaiting Your Approval</p>
          </div>
          
          <p>Please log in to your dashboard to review this request.</p>
          
          <p>
            <a href="https://ccitsu.github.io/dacs/pages/hod-dashboard.html" 
               style="display: inline-block; background: #f39c12; color: white; padding: 12px 30px; 
                      text-decoration: none; border-radius: 5px; margin-top: 10px;">
              Review Request
            </a>
          </p>
          
          ${EMAIL_CONFIG.SIGNATURE}
        </div>
      </body>
    </html>
  `;
  
  sendEmailSafe({
    to: hodEmail,
    subject: subject,
    htmlBody: body
  }, 'sendHODNotificationEmail');
}

function sendRegistrarNotificationEmail(requestId, studentName) {
  // Resolve Registrar email by student's department via requestId → studentId → Students → Registrars
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let registrarEmail = '';
  try {
    const requestsSheet = ss.getSheetByName(SHEETS.REQUESTS);
    const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);
    const registrarsSheet = ss.getSheetByName(SHEETS.REGISTRARS);
    if (requestsSheet && studentsSheet && registrarsSheet) {
      const reqData = requestsSheet.getDataRange().getValues();
      let studentId = '';
      for (let i = 1; i < reqData.length; i++) {
        if (reqData[i][0] === requestId) { // Column A (index 0) = Request ID
          studentId = reqData[i][1]; // Column B (index 1) = Student ID
          break;
        }
      }
      if (studentId) {
        const studentsData = studentsSheet.getDataRange().getValues();
        let department = '';
        for (let s = 1; s < studentsData.length; s++) {
          if (studentsData[s][0] === studentId) { // Column A (index 0) = Student ID
            department = studentsData[s][5] || ''; // Column F (index 5) = Department
            break;
          }
        }
        if (department) {
          const registrarsData = registrarsSheet.getDataRange().getValues();
          for (let r = 1; r < registrarsData.length; r++) {
            if ((registrarsData[r][3] || '') === department) { // Column D (index 3) = Department
              registrarEmail = (registrarsData[r][2] || '').toString().trim(); // Column C (index 2) = Registrar email
              break;
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('sendRegistrarNotificationEmail resolution error: ' + e);
  }
  if (!registrarEmail || registrarEmail.indexOf('@') < 1) {
    Logger.log('sendRegistrarNotificationEmail skipped: no valid registrar email resolved for request ' + requestId);
    return;
  }
  
  const subject = `Request Ready for Processing - ${studentName}`;
  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
          <h2 style="color: #9b59b6; border-bottom: 2px solid #9b59b6; padding-bottom: 10px;">
            ⏳ Request Ready for Final Processing
          </h2>
          
          <p>Dear Registrar,</p>
          
          <p>A course add/drop request has been fully approved and is ready for final registration processing.</p>
          
          <div style="background: #e8daef; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p style="margin: 0;"><strong>Student:</strong> ${studentName}</p>
            <p style="margin: 10px 0 0 0;"><strong>Request ID:</strong> #${requestId}</p>
            <p style="margin: 10px 0 0 0;"><strong>Approvals:</strong> ✓ Advisor, ✓ HOD</p>
          </div>
          
          <p>Please process this request and generate the final approval document.</p>
          
          <p>
            <a href="https://ccitsu.github.io/dacs/pages/registrar-dashboard.html" 
               style="display: inline-block; background: #9b59b6; color: white; padding: 12px 30px; 
                      text-decoration: none; border-radius: 5px; margin-top: 10px;">
              Process Request
            </a>
          </p>
          
          ${EMAIL_CONFIG.SIGNATURE}
        </div>
      </body>
    </html>
  `;
  
  sendEmailSafe({
    to: registrarEmail,
    subject: subject,
    htmlBody: body
  }, 'sendRegistrarNotificationEmail');
}

function sendApproverConfirmationEmail(approverEmail, approverRole, requestId, action) {
  // Guard against invalid email inputs (e.g., IDs like "H123")
  if (!approverEmail || approverEmail.indexOf('@') < 1) {
    Logger.log('sendApproverConfirmationEmail skipped: invalid email "' + approverEmail + '" for ' + approverRole + ' on request ' + requestId);
    return;
  }
  const subject = `Confirmation: Request ${action.charAt(0).toUpperCase() + action.slice(1)}`;
  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
          <h2 style="color: #27ae60; border-bottom: 2px solid #27ae60; padding-bottom: 10px;">
            ✓ Action Confirmed
          </h2>
          
          <p>Dear ${approverRole},</p>
          
          <p>This confirms that you have ${action} the following request:</p>
          
          <div style="background: #d4edda; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p style="margin: 0;"><strong>Request ID:</strong> #${requestId}</p>
            <p style="margin: 10px 0 0 0;"><strong>Your Action:</strong> ${action.toUpperCase()}</p>
          </div>
          
          <p>The request has been moved to the next stage in the approval workflow.</p>
          
          ${EMAIL_CONFIG.SIGNATURE}
        </div>
      </body>
    </html>
  `;
  
  sendEmailSafe({
    to: approverEmail,
    subject: subject,
    htmlBody: body
  }, 'sendApproverConfirmationEmail');
}

// ===================================
// COMPLETION EMAIL WITH PDF
// ===================================

function sendCompletionEmail(request, pdfUrl) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const requestId = request[0];
  const studentId = request[1];
  const studentName = request[2];
  const studentEmail = request[3];
  const courseCode = request[5];
  const courseName = request[6];
  const advisorId = request[17];
  const advisorName = request[18];
  const advisorEmailFromRequest = request[19];

  // Resolve advisor email: prefer stored on request, fallback to Advisors sheet by ID
  let advisorEmail = (advisorEmailFromRequest || '').trim();
  if (!advisorEmail && advisorId) {
    const advisorsSheet = ss.getSheetByName(SHEETS.ADVISORS);
    if (advisorsSheet) {
      const advisorsData = advisorsSheet.getDataRange().getValues();
      for (let i = 1; i < advisorsData.length; i++) {
        if (advisorsData[i][0] === advisorId) {
          advisorEmail = advisorsData[i][2] || ''; // Column C (index 2) holds advisor email
          break;
        }
      }
    }
  }

  // Resolve student department to identify the HOD for this record
  let studentDepartment = '';
  const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);
  if (studentsSheet) {
    const studentsData = studentsSheet.getDataRange().getValues();
    for (let i = 1; i < studentsData.length; i++) {
      if (studentsData[i][0] === studentId) {
        studentDepartment = studentsData[i][5] || ''; // Column F (index 5) = Department
        break;
      }
    }
  }

  // Resolve HOD email by department match
  let hodEmail = '';
  if (studentDepartment) {
    const hodsSheet = ss.getSheetByName(SHEETS.HODS);
    if (hodsSheet) {
      const hodsData = hodsSheet.getDataRange().getValues();
      for (let i = 1; i < hodsData.length; i++) {
        if ((hodsData[i][5] || '') === studentDepartment) { // Column F (index 5) = Department
          hodEmail = hodsData[i][3] || ''; // Column D (index 3) = HOD email
          break;
        }
      }
    }
  }

  // Resolve Registrar email by department match
  let registrarEmail = '';
  if (studentDepartment) {
    const registrarsSheet = ss.getSheetByName(SHEETS.REGISTRARS);
    if (registrarsSheet) {
      const registrarsData = registrarsSheet.getDataRange().getValues();
      for (let i = 1; i < registrarsData.length; i++) {
        if ((registrarsData[i][3] || '') === studentDepartment) { // Column D (index 3) = Department
          registrarEmail = registrarsData[i][2] || ''; // Column C (index 2) = Registrar email
          break;
        }
      }
    }
  }

  // Get PDF blob
  const pdfFile = DriveApp.getFileById(pdfUrl.split('/d/')[1].split('/')[0]);
  const pdfBlob = pdfFile.getBlob();

  const subject = `Course Add/Drop Request Completed - #${requestId}`;
  const body = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #2c3e50;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
          <h2 style="color: #27ae60; border-bottom: 2px solid #27ae60; padding-bottom: 10px;">
            🎉 Request Completed Successfully!
          </h2>
          
          <p>Dear ${studentName},</p>
          
          <p>Your course add/drop request has been processed and completed successfully.</p>
          
          <div style="background: #d4edda; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p style="margin: 0;"><strong>Request ID:</strong> #${requestId}</p>
            <p style="margin: 10px 0 0 0;"><strong>Course:</strong> ${courseCode} - ${courseName}</p>
            <p style="margin: 10px 0 0 0;"><strong>Status:</strong> ✓ Completed</p>
          </div>
          
          <p><strong>Approvals Received From:</strong></p>
          <ul>
            <li>✓ Academic Advisor${advisorName ? ' (' + advisorName + ')' : ''}</li>
            <li>✓ Head of Department</li>
            <li>✓ Registrar</li>
          </ul>
          
          <p>The official approval document is attached to this email. Please keep it for your records.</p>
          
          <p>Your course registration has been updated accordingly.</p>
          
          ${EMAIL_CONFIG.SIGNATURE}
        </div>
      </body>
    </html>
  `;

  const recipients = [studentEmail, advisorEmail, hodEmail, registrarEmail].filter(Boolean);
  const uniqueRecipients = Array.from(new Set(recipients));

  if (uniqueRecipients.length === 0) {
    Logger.log('No recipients resolved for completion email for request ' + requestId);
    return;
  }

  sendEmailSafe({
    to: uniqueRecipients.join(','),
    subject: subject,
    htmlBody: body,
    attachments: [pdfBlob]
  }, 'sendCompletionEmail');
}
