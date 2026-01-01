/**
 * Course Add/Drop Management System
 * PDF Generation Module
 */

// ===================================
// PDF GENERATION
// ===================================

// Optional: set one of these for logo inclusion
// If DRIVE_LOGO_FILE_ID is set to a Drive file ID, the image will be embedded inline (recommended)
// Otherwise, if LOGO_URL is set to a public image URL, it will be referenced directly
const DRIVE_LOGO_FILE_ID = '15vly1qIRbr4XukziaVXBmeAKD4UUT8f9';
const LOGO_URL = '';

function getLogoDataUrl() {
  try {
    if (DRIVE_LOGO_FILE_ID) {
      const blob = DriveApp.getFileById(DRIVE_LOGO_FILE_ID).getBlob();
      const contentType = blob.getContentType() || 'image/png';
      const base64 = Utilities.base64Encode(blob.getBytes());
      return `data:${contentType};base64,${base64}`;
    }
    if (LOGO_URL) {
      return LOGO_URL;
    }
  } catch (e) {
    Logger.log('Logo fetch failed: ' + e);
  }
  return '';
}

function generateApprovalPDF(request, updateData) {
  // Normalize request input (can be a row array or an object)
  if (!request) {
    throw new Error('PDF generation failed: request data is missing');
  }
  var req = {};
  if (Array.isArray(request)) {
    req.requestId = request[0];
    req.studentId = request[1];
    req.studentName = request[2];
    req.studentEmail = request[3];
    req.requestType = request[4];
    req.courseCode = request[5];
    req.courseName = request[6];
    req.sectionNumber = request[7];
    req.reason = request[9];
    req.submittedDate = request[11];
  } else if (typeof request === 'object') {
    req.requestId = request.id;
    req.studentId = request.studentId;
    req.studentName = request.studentName;
    req.studentEmail = request.studentEmail;
    req.requestType = request.type;
    req.courseCode = request.courseCode;
    req.courseName = request.courseName;
    req.sectionNumber = request.sectionToAdd || '';
    req.reason = request.reason;
    req.submittedDate = request.submittedDate;
    req.courseDetails = request.courseDetails;
  } else {
    throw new Error('PDF generation failed: invalid request data format');
  }

  const requestId = req.requestId || '';
  const studentId = req.studentId || '';
  const studentName = req.studentName || '';
  const studentEmail = req.studentEmail || '';
  const requestType = (req.requestType || '').toString();
  const courseCode = req.courseCode || '';
  const courseName = req.courseName || '';
  const sectionNumber = req.sectionNumber || '';
  const reason = req.reason || '';
  const submittedDate = req.submittedDate || new Date().toISOString();

  // Structured course details (supports add/drop lists)
  let courseDetails = {};
  if (req.courseDetails) {
    courseDetails = req.courseDetails;
  } else if (Array.isArray(request) && request[16]) {
    try {
      courseDetails = JSON.parse(request[16]);
    } catch (e) {
      Logger.log('Failed to parse courseDetails JSON: ' + e);
    }
  }

  const coursesToAdd = Array.isArray(courseDetails.coursesToAdd) ? courseDetails.coursesToAdd : [];
  const coursesToDrop = Array.isArray(courseDetails.coursesToDrop) ? courseDetails.coursesToDrop : [];

  // Fallbacks for legacy single-course records
  const legacyAdd = (!coursesToAdd.length && courseCode) ? [courseCode + (sectionNumber ? ` (Section ${sectionNumber})` : '') + (courseName ? ` - ${courseName}` : '')] : [];
  const legacyDrop = (!coursesToDrop.length && courseName) ? [courseName] : [];

  const addList = coursesToAdd.length ? coursesToAdd : legacyAdd;
  const dropList = coursesToDrop.length ? coursesToDrop : legacyDrop;

  function formatCourseItem(course) {
    if (!course) return '';
    if (typeof course === 'string') return course;

    const code = course.code || '';
    const name = course.name || '';
    const main = [code, name].filter(Boolean).join(' - ');
    const metaParts = [];
    if (course.activityType) metaParts.push(course.activityType);
    if (course.section) metaParts.push('Section ' + course.section);
    const schedule = [course.day, course.time].filter(Boolean).join(' ');
    if (schedule) metaParts.push(schedule);
    if (course.creditHours) metaParts.push(course.creditHours + ' CH');
    return main + (metaParts.length ? ' (' + metaParts.join(' | ') + ')' : '');
  }

  function renderCourseList(items, emptyText) {
    if (!items || !items.length) {
      return `<div class="course-item" style="border-left-color:#bdc3c7;">${emptyText}</div>`;
    }
    return items.map(c => `<div class="course-item"><strong>${formatCourseItem(c)}</strong></div>`).join('');
  }
  
  // Get approval details from Approvals sheet
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const approvalsSheet = ss.getSheetByName(SHEETS.APPROVALS);
  const approvals = approvalsSheet.getDataRange().getValues();
  
  let advisorApproval = { name: '', date: '', comments: '' };
  let hodApproval = { name: '', date: '', comments: '' };
  let registrarApproval = { name: '', date: '', comments: '' };
  
  // Extract approval information
  for (let i = 1; i < approvals.length; i++) {
    if (approvals[i][0] === requestId) {
      const role = approvals[i][1];
      const name = approvals[i][2];
      const status = approvals[i][3];
      const comments = approvals[i][4];
      const date = approvals[i][5];
      
      if (role === 'advisor' && status === 'approved') {
        advisorApproval = { name, date, comments };
      } else if (role === 'hod' && status === 'approved') {
        hodApproval = { name, date, comments };
      } else if (role === 'registrar' && status === 'completed') {
        registrarApproval = { name, date, comments };
      }
    }
  }
  
  // Create HTML content for PDF - optimized for single page
  const logoSrc = getLogoDataUrl();
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        * {
          margin: 0;
          padding: 0;
          box-sizing: border-box;
        }
        body {
          font-family: 'Arial', sans-serif;
          line-height: 1.4;
          color: #2c3e50;
          padding: 15px;
          background: white;
        }
        @page {
          size: A4;
          margin: 10mm;
        }
        .header {
          text-align: center;
          border-bottom: 3px solid #0c542c;
          padding-bottom: 10px;
          margin-bottom: 12px;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 15px;
        }
        .logo-img {
          width: 50px;
          height: 50px;
          border-radius: 4px;
        }
        .header-text {
          text-align: center;
        }
        .university-name {
          font-size: 14px;
          font-weight: bold;
          color: #0c542c;
        }
        .department-name {
          font-size: 11px;
          color: #bc9461;
          font-weight: 600;
        }
        .document-title {
          font-size: 13px;
          font-weight: bold;
          color: #0c542c;
          text-align: center;
          margin: 8px 0;
          text-decoration: underline;
        }
        .section {
          margin-bottom: 10px;
          background: #f0f7f4;
          padding: 8px;
          border-radius: 4px;
          border-left: 3px solid #bc9461;
          page-break-inside: avoid;
        }
        .section-title {
          font-size: 11px;
          font-weight: bold;
          color: #0c542c;
          margin-bottom: 5px;
        }
        .field {
          display: flex;
          justify-content: space-between;
          margin-bottom: 3px;
          font-size: 10px;
        }
        .field-label {
          font-weight: bold;
          color: #0c542c;
          min-width: 120px;
        }
        .field-value {
          color: #2c3e50;
          flex: 1;
          text-align: right;
        }
        .course-item {
          background: white;
          padding: 6px;
          margin: 4px 0;
          border-left: 3px solid #27ae60;
          font-size: 9px;
          line-height: 1.3;
        }
        .course-item strong {
          color: #0c542c;
          display: block;
        }
        .course-detail {
          color: #555;
          margin-left: 5px;
        }
        .approval-chain {
          display: flex;
          justify-content: space-around;
          margin: 8px 0;
          align-items: center;
          font-size: 9px;
        }
        .approval-item {
          text-align: center;
          min-width: 80px;
        }
        .approval-check {
          font-size: 24px;
          color: #27ae60;
          font-weight: bold;
          display: block;
        }
        .approval-label {
          font-size: 8px;
          color: #0c542c;
          font-weight: bold;
        }
        .footer {
          border-top: 2px solid #0c542c;
          padding-top: 8px;
          margin-top: 12px;
          text-align: center;
          font-size: 8px;
          color: #666;
          page-break-inside: avoid;
        }
        .footer-logo {
          width: 40px;
          height: 40px;
          margin: 0 auto 5px;
          border-radius: 4px;
        }
        .reference-number {
          text-align: center;
          color: #7f8c8d;
          font-size: 9px;
          margin-bottom: 8px;
        }
      </style>
    </head>
    <body>
      <!-- Header with Logo -->
      <div class="header">
        ${logoSrc ? `<img src="${logoSrc}" class="logo-img"/>` : ''}
        <div>
          <div class="university-name">Shaqra University</div>
          <div class="department-name">College of Computing and Information Technology</div>
        </div>
      </div>
      
      <!-- Reference Number -->
      <div class="reference-number">
        Ref: CCIT/${requestId}/${new Date().getFullYear()} | ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })}
      </div>
      
      <!-- Document Title -->
      <div class="document-title">COURSE ADD/DROP APPROVAL FORM</div>
      
      <!-- Student Info -->
      <div class="section">
        <div class="section-title">STUDENT INFORMATION</div>
        <div class="field">
          <span class="field-label">Name:</span>
          <span class="field-value">${studentName}</span>
        </div>
        <div class="field">
          <span class="field-label">ID:</span>
          <span class="field-value">${studentId}</span>
        </div>
        <div class="field">
          <span class="field-label">Email:</span>
          <span class="field-value">${studentEmail}</span>
        </div>
        <div class="field">
          <span class="field-label">Request Type:</span>
          <span class="field-value"><strong>${requestType.toUpperCase()}</strong></span>
        </div>
      </div>
      
      <!-- Course Details Section -->
      <div class="section">
        <div class="section-title">COURSES TO ADD</div>
        ${renderCourseList(addList, 'No courses to add')}
      </div>

      <div class="section">
        <div class="section-title">COURSES TO DROP</div>
        ${renderCourseList(dropList, 'No courses to drop')}
      </div>

      <div class="section">
        <div class="section-title">REQUEST DETAILS</div>
        <div class="field">
          <span class="field-label">Request Type:</span>
          <span class="field-value"><strong>${requestType.toUpperCase()}</strong></span>
        </div>
        <div class="field">
          <span class="field-label">Reason:</span>
          <span class="field-value">${reason || 'N/A'}</span>
        </div>
        <div class="field">
          <span class="field-label">Submitted:</span>
          <span class="field-value">${new Date(submittedDate).toLocaleString()}</span>
        </div>
      </div>
      
      <!-- Approval Chain - Simplified -->
      <div class="section">
        <div class="section-title">APPROVAL STATUS</div>
        <div class="approval-chain">
          <div class="approval-item">
            <div class="approval-check">✓</div>
            <div class="approval-label">Academic Advisor</div>
          </div>
          <div class="approval-item">
            <div class="approval-check">✓</div>
            <div class="approval-label">Head of Department</div>
          </div>
          <div class="approval-item">
            <div class="approval-check">✓</div>
            <div class="approval-label">Registrar</div>
          </div>
        </div>
      </div>
      
      <!-- Footer -->
      <div class="footer">
        <p><strong>🏛️ Shaqra University - CCIT</strong></p>
        <p>Course Add/Drop Management System v1.0</p>
        <p>Generated: ${new Date().toLocaleString('en-US')}</p>
      </div>
    </body>
    </html>
  `;
  
  // Convert HTML to a Google Doc via Drive API and export to PDF
  const filename = `AddDrop_Approval_${requestId}_${studentName}`;
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const htmlBlob = Utilities.newBlob(htmlContent, 'text/html', filename + '.html');

  try {
    // Use Advanced Drive service to convert HTML -> Google Doc
    // Ensure Advanced Drive service is enabled in Apps Script project.
    const docFile = Drive.Files.insert({
      title: filename,
      mimeType: MimeType.GOOGLE_DOCS,
      parents: [{ id: DRIVE_FOLDER_ID }]
    }, htmlBlob);

    // Export Doc to PDF
    const pdfBlob = Drive.Files.export(docFile.id, 'application/pdf');
    pdfBlob.setName(filename + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);

    // Cleanup temp doc
    DriveApp.getFileById(docFile.id).setTrashed(true);

    // Share PDF
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Return direct download URL instead of preview URL
    return 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId();
  } catch (e) {
    // Fallback: create HTML file and convert its blob to PDF
    const tempHtml = folder.createFile(filename + '.html', htmlContent, MimeType.HTML);
    const pdfBlob = tempHtml.getBlob().getAs('application/pdf');
    pdfBlob.setName(filename + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    tempHtml.setTrashed(true);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Return direct download URL instead of preview URL
    return 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId();
  }
}

// ===================================
// ALTERNATIVE: HTML TO PDF CONVERTER
// ===================================

/**
 * Alternative method using Google Docs HTML to PDF conversion
 */
function generateApprovalPDFFromHTML(request, updateData) {
  const requestId = request[0];
  const studentName = request[2];
  
  // Generate HTML content (same as above)
  const htmlContent = generatePDFHTMLContent(request, updateData);
  
  // Create a temporary HTML file
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const htmlFile = folder.createFile(`temp_${requestId}.html`, htmlContent, MimeType.HTML);
  
  // Get the file as PDF
  const htmlBlob = htmlFile.getBlob();
  const pdfBlob = htmlBlob.getAs('application/pdf');
  pdfBlob.setName(`AddDrop_Approval_${requestId}_${studentName}.pdf`);
  
  // Create PDF file
  const pdfFile = folder.createFile(pdfBlob);
  
  // Delete temporary HTML file
  htmlFile.setTrashed(true);
  
  // Make PDF shareable
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // Return direct download URL instead of preview URL
  return 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId();
}

function generatePDFHTMLContent(request, updateData) {
  // This function extracts the HTML generation logic
  // to be reused across different PDF generation methods
  // (Same HTML content as in generateApprovalPDF)
  
  const requestId = request[0];
  const studentId = request[1];
  const studentName = request[2];
  const studentEmail = request[3];
  const requestType = request[4];
  const courseCode = request[5];
  const courseName = request[6];
  const sectionNumber = request[7];
  const reason = request[9];
  const submittedDate = request[11];
  
  // Get approvals
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const approvalsSheet = ss.getSheetByName(SHEETS.APPROVALS);
  const approvals = approvalsSheet.getDataRange().getValues();
  
  let advisorApproval = { name: 'N/A', date: new Date(), comments: '' };
  let hodApproval = { name: 'N/A', date: new Date(), comments: '' };
  let registrarApproval = { name: 'N/A', date: new Date(), comments: '' };
  
  for (let i = 1; i < approvals.length; i++) {
    if (approvals[i][0] === requestId) {
      const role = approvals[i][1];
      if (role === 'advisor') {
        advisorApproval = { name: approvals[i][2], date: approvals[i][5], comments: approvals[i][4] };
      } else if (role === 'hod') {
        hodApproval = { name: approvals[i][2], date: approvals[i][5], comments: approvals[i][4] };
      } else if (role === 'registrar') {
        registrarApproval = { name: approvals[i][2], date: approvals[i][5], comments: approvals[i][4] };
      }
    }
  }
  
  return `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; }
          .header { text-align: center; margin-bottom: 30px; }
          .section { margin: 20px 0; padding: 15px; background: #f5f5f5; }
          .field { margin: 5px 0; }
          .label { font-weight: bold; display: inline-block; width: 150px; }
          .approval { border: 2px solid green; padding: 10px; margin: 10px 0; }
        </style>
      </head>
      <body>
        <div class="header">
          ${getLogoDataUrl() ? `<img src="${getLogoDataUrl()}" style="height:40px;margin-right:10px;"/>` : ''}
          <h1>Shaqra University - CCIT</h1>
          <h2>Course Add/Drop Approval</h2>
          <p>Request ID: ${requestId}</p>
        </div>
        <div class="section">
          <h3>Student Information</h3>
          <div class="field"><span class="label">Name:</span> ${studentName}</div>
          <div class="field"><span class="label">ID:</span> ${studentId}</div>
          <div class="field"><span class="label">Email:</span> ${studentEmail}</div>
        </div>
        <div class="section">
          <h3>Request Details</h3>
          <div class="field"><span class="label">Type:</span> ${requestType}</div>
          <div class="field"><span class="label">Course:</span> ${courseCode} - ${courseName}</div>
          <div class="field"><span class="label">Reason:</span> ${reason}</div>
        </div>
        <div class="section">
          <h3>Approvals</h3>
          <div class="approval">
            <strong>Advisor:</strong> ${advisorApproval.name}<br>
            <strong>Date:</strong> ${new Date(advisorApproval.date).toLocaleDateString()}
          </div>
          <div class="approval">
            <strong>HOD:</strong> ${hodApproval.name}<br>
            <strong>Date:</strong> ${new Date(hodApproval.date).toLocaleDateString()}
          </div>
          <div class="approval">
            <strong>Registrar:</strong> ${registrarApproval.name}<br>
            <strong>Date:</strong> ${new Date(registrarApproval.date).toLocaleDateString()}
          </div>
        </div>
      </body>
    </html>
  `;
}
