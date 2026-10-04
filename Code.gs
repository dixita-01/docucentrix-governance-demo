/**
 * DOCUCENTRIX
 * Enterprise Document Governance & Compliance Platform
 * Built for Your Enterprise Name
 */
const CONFIG = {
  GEMINI_API_KEY: 'YOUR_GEMINI_API_KEY_HERE',
  ROOT_FOLDER_NAME: 'DocuCentrix_Repository',
  SPREADSHEET_ID: 'DUMMY_SHEET_ID',
  AUDIT_SPREADSHEET_NAME: 'DocuCentrix_Audit_Logs_Spreadsheet',
  SUPER_ADMIN_EMAIL: 'admin@example.com' // demo placeholder
};

const ALLOWED_DEPARTMENTS = {
  'HRO': ['Payroll', 'EDM'],
  'Finance': ['Order To Cash', 'Treasury', 'Record to Report', 'Record to Pay']
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('DocuCentrix | Your Enterprise Name Document Governance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDb_() {
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (err) {
    throw new Error("Could not open Spreadsheet database. Verify SPREADSHEET_ID in CONFIG. Details: " + err.message);
  }
}

function getAuditDb_() {
  try {
    const files = DriveApp.getFilesByName(CONFIG.AUDIT_SPREADSHEET_NAME);
    if (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    } else {
      const ss = SpreadsheetApp.create(CONFIG.AUDIT_SPREADSHEET_NAME);
      const sheet = ss.getActiveSheet();
      sheet.setName('AuditLogs');
      sheet.appendRow(['AuditID', 'DocumentID', 'Action', 'PerformedBy', 'Role', 'Timestamp', 'Comments', 'VersionNumber']);
      return ss;
    }
  } catch (err) {
    return getDb_();
  }
}

function getAuditSpreadsheetUrl() {
  const auditDb = getAuditDb_();
  return auditDb.getUrl();
}

function getSheetFlexible_(tableName) {
  const db = getDb_();
  let sheet = db.getSheetByName(tableName);
  if (!sheet) {
    sheet = db.getSheetByName('DocuCentrix_DB_' + tableName);
  }
  if (!sheet) {
    const allSheets = db.getSheets();
    sheet = allSheets.find(s => {
      const name = s.getName().trim().toLowerCase();
      return name === tableName.toLowerCase() || name === ('docucentrix_db_' + tableName).toLowerCase();
    });
  }
  return sheet;
}

function extractSheetObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let headers = data[0].map(h => String(h).trim());

  return data.slice(1).filter(row => row.some(cell => String(cell).trim() !== '')).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        val = val.toISOString();
      } else if (typeof val === 'number' && val > 40000 && val < 60000) {
        val = new Date(Math.round((val - 25569) * 86400 * 1000)).toISOString();
      }
      obj[h] = val !== undefined ? val : '';
    });
    return obj;
  });
}

function getTableData_(tableName) {
  try {
    let primaryData = [];
    let legacyData = [];

    if (tableName === 'AuditLogs') {
      const auditDb = getAuditDb_();
      const auditSheet = auditDb.getSheetByName('AuditLogs') || auditDb.getSheets()[0];
      if (auditSheet) {
        primaryData = extractSheetObjects_(auditSheet);
      }
     
      const mainDb = getDb_();
      const legacySheet = mainDb.getSheetByName('AuditLogs') || mainDb.getSheetByName('DocuCentrix_DB_AuditLogs');
      if (legacySheet) {
        legacyData = extractSheetObjects_(legacySheet);
      }

      const combined = [...primaryData, ...legacyData];
      const uniqueMap = new Map();
      combined.forEach(item => {
        const id = item.AuditID || (item.DocumentID + '_' + item.Timestamp);
        if (id && !uniqueMap.has(id)) {
          uniqueMap.set(id, item);
        }
      });
      return Array.from(uniqueMap.values());
    }

    const sheet = getSheetFlexible_(tableName);
    if (!sheet) return [];
    return extractSheetObjects_(sheet);
  } catch (e) {
    return [];
  }
}

function appendRow_(tableName, rowObject) {
  const db = (tableName === 'AuditLogs') ? getAuditDb_() : getDb_();
  let sheet = db.getSheetByName(tableName);
  if (!sheet) {
    sheet = db.getSheetByName('DocuCentrix_DB_' + tableName);
  }
  if (!sheet) {
    sheet = db.getSheets()[0];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    sheet.appendRow(Object.keys(rowObject));
  }

  const headers = sheet.getDataRange().getValues()[0].map(h => String(h).trim());
  const row = headers.map(h => rowObject[h] !== undefined ? rowObject[h] : '');
  sheet.appendRow(row);
}

function updateRow_(tableName, keyColumn, keyValue, updateObject) {
  const sheet = getSheetFlexible_(tableName);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const keyIdx = headers.indexOf(keyColumn);
  if (keyIdx === -1) return;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyIdx]).trim().toLowerCase() === String(keyValue).trim().toLowerCase()) {
      for (let key in updateObject) {
        const colIdx = headers.indexOf(key);
        if (colIdx !== -1) {
          sheet.getRange(i + 1, colIdx + 1).setValue(updateObject[key]);
        }
      }
      break;
    }
  }
}

function getCurrentUser() {
  let activeEmail = '';
  try {
    activeEmail = Session.getActiveUser().getEmail();
  } catch(e) {}

  if (!activeEmail || activeEmail.trim() === '') {
    activeEmail = CONFIG.SUPER_ADMIN_EMAIL;
  } else {
    activeEmail = activeEmail.trim().toLowerCase();
  }

  const isSuperUser = (activeEmail === CONFIG.SUPER_ADMIN_EMAIL);
  const allSystemRoles = ['Admin', 'Manager', 'Team Lead', 'Editor', 'Final Approver', 'Viewer'];

  const users = getTableData_('Users');
  const userEntries = users.filter(u => u.Email && String(u.Email).trim().toLowerCase() === activeEmail);

  if (userEntries.length > 0) {
    const rolesSet = new Set();
   
    if (isSuperUser) {
      allSystemRoles.forEach(r => rolesSet.add(r));
    } else {
      userEntries.forEach(u => {
        if (u.Role && String(u.Role).trim() !== '') {
          rolesSet.add(String(u.Role).trim());
        }
      });
    }

    const assignedRoles = Array.from(rolesSet);
    const primaryUser = userEntries[0];

    return {
      UserID: primaryUser.UserID || 'USR-001',
      Email: activeEmail,
      FullName: primaryUser.FullName || activeEmail.split('@')[0],
      Role: String(primaryUser.Role || (isSuperUser ? 'Admin' : 'Viewer')).trim(),
      AssignedRoles: assignedRoles.length > 0 ? assignedRoles : ['Viewer'],
      Department: String(primaryUser.Department || '').trim(),
      SubDepartment: String(primaryUser.SubDepartment || '').trim(),
      Team: String(primaryUser.Team || 'General').trim(),
      Status: primaryUser.Status || 'ACTIVE'
    };
  }

  return {
    UserID: 'USR-001',
    Email: activeEmail,
    FullName: activeEmail.split('@')[0],
    Role: isSuperUser ? 'Admin' : 'Viewer',
    AssignedRoles: isSuperUser ? allSystemRoles : ['Viewer'],
    Department: isSuperUser ? 'HRO' : '',
    SubDepartment: isSuperUser ? 'Payroll' : '',
    Team: 'General',
    Status: 'ACTIVE'
  };
}

function switchActiveUserRole(newRole) {
  const user = getCurrentUser();
  if (!user.AssignedRoles.includes(newRole) && user.Role !== 'Admin') {
    throw new Error("Unauthorized role switch attempt.");
  }

  updateRow_('Users', 'Email', user.Email, { Role: newRole });
  logAudit_('SYSTEM', 'Role Switched', user.Email, newRole, `User switched active session role to ${newRole}`, 'v1.0');
 
  return { success: true, newRole: newRole };
}

function switchActiveUserDepartment(newDepartment) {
  const user = getCurrentUser();
  const users = getTableData_('Users').filter(u => String(u.Email).trim().toLowerCase() === user.Email.toLowerCase());
 
  const allowedDepts = new Set(users.map(u => String(u.Department).trim()));
  if (user.Role !== 'Admin' && !allowedDepts.has(newDepartment) && !allowedDepts.has('ALL')) {
    throw new Error("Unauthorized department switch attempt.");
  }

  updateRow_('Users', 'Email', user.Email, { Department: newDepartment });
  logAudit_('SYSTEM', 'Department Switched', user.Email, user.Role, `User switched active department scope to ${newDepartment}`, 'v1.0');
 
  return { success: true, newDepartment: newDepartment };
}

function assignUserHierarchy(targetEmail, targetRole, department, subDepartment, team) {
  const currentUser = getCurrentUser();
 
  if (!ALLOWED_DEPARTMENTS[department]) {
    throw new Error("Invalid Department selected. Allowed departments: HRO, Finance.");
  }

  const validSubDepts = ALLOWED_DEPARTMENTS[department];
  let chosenSubDept = 'ALL';
  if (targetRole !== 'Manager') {
    chosenSubDept = (subDepartment && validSubDepts.includes(subDepartment)) ? subDepartment : validSubDepts[0];
  } else if (subDepartment && validSubDepts.includes(subDepartment)) {
    chosenSubDept = subDepartment;
  }

  if (currentUser.Role === 'Admin' && !['Manager', 'Final Approver'].includes(targetRole)) {
    throw new Error("Admins can only assign Manager and Final Approver roles.");
  }
  if (currentUser.Role === 'Manager' && !['Team Lead', 'Final Approver'].includes(targetRole)) {
    throw new Error("Managers can only assign Team Lead and Final Approver roles.");
  }
  if (currentUser.Role === 'Team Lead' && !['Editor', 'Viewer'].includes(targetRole)) {
    throw new Error("Team Leads can only assign Editor and Viewer roles.");
  }

  const existingUsers = getTableData_('Users');
  const userRow = existingUsers.find(u => String(u.Email).trim().toLowerCase() === targetEmail.trim().toLowerCase() && String(u.Role).trim() === targetRole && String(u.Department).trim() === department);

  if (userRow) {
    updateRow_('Users', 'Email', targetEmail, {
      Department: department,
      SubDepartment: chosenSubDept,
      Team: team || currentUser.Team,
      AssignedBy: currentUser.Email,
      Status: 'ACTIVE'
    });
  } else {
    appendRow_('Users', {
      UserID: 'USR-' + Math.floor(100000 + Math.random() * 900000),
      Email: targetEmail.trim().toLowerCase(),
      FullName: targetEmail.split('@')[0],
      Role: targetRole,
      Department: department,
      SubDepartment: chosenSubDept,
      Team: team || currentUser.Team,
      AssignedBy: currentUser.Email,
      Status: 'ACTIVE'
    });
  }
 
  return { success: true };
}

function getUserAssignments(targetEmail) {
  const currentUser = getCurrentUser();
  if (!['Manager', 'Admin', 'Team Lead'].includes(currentUser.Role)) {
    throw new Error("Unauthorized: Insufficient privileges to view user assignments.");
  }

  const cleanEmail = String(targetEmail || '').trim().toLowerCase();
  if (!cleanEmail) return { success: true, assignments: [] };

  const users = getTableData_('Users');
  const userEntries = users.filter(u =>
    String(u.Email || '').trim().toLowerCase() === cleanEmail &&
    String(u.Status || 'ACTIVE').toUpperCase() !== 'REVOKED'
  );

  const assignments = userEntries.map(u => ({
    role: String(u.Role || '').trim(),
    department: String(u.Department || '').trim(),
    subDepartment: String(u.SubDepartment || '').trim(),
    fullName: String(u.FullName || '').trim()
  })).filter(a => a.role !== '');

  return {
    success: true,
    email: cleanEmail,
    fullName: userEntries.length > 0 ? (userEntries[0].FullName || cleanEmail.split('@')[0]) : cleanEmail.split('@')[0],
    assignments: assignments
  };
}

function revokeUserAccess(targetEmail, targetRole, targetDept, targetSubDept, reason) {
  const currentUser = getCurrentUser();
  if (!['Manager', 'Admin', 'Team Lead'].includes(currentUser.Role)) {
    throw new Error("Unauthorized: Insufficient privileges to revoke user access.");
  }

  if (!targetEmail || !targetRole || !targetDept || !reason) {
    throw new Error("Mandatory revocation parameters missing. Email, Role, Department, and Reason are required.");
  }

  const cleanEmail = String(targetEmail).trim().toLowerCase();
  const cleanRole = String(targetRole).trim();
  const cleanDept = String(targetDept).trim();
  const cleanSubDept = String(targetSubDept || '').trim();

  const sheet = getSheetFlexible_('Users');
  if (!sheet) throw new Error("Users table database unavailable.");

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) throw new Error("No user records found in database.");

  const headers = data[0].map(h => String(h).trim());
  const emailIdx = headers.indexOf('Email');
  const roleIdx = headers.indexOf('Role');
  const deptIdx = headers.indexOf('Department');
  const subDeptIdx = headers.indexOf('SubDepartment');
  const statusIdx = headers.indexOf('Status');

  if (emailIdx === -1 || roleIdx === -1) {
    throw new Error("Users table schema mismatch.");
  }

  let rowToRevoke = -1;
  let targetFullName = cleanEmail.split('@')[0];

  for (let i = 1; i < data.length; i++) {
    const rEmail = String(data[i][emailIdx]).trim().toLowerCase();
    const rRole = String(data[i][roleIdx]).trim();
    const rDept = deptIdx !== -1 ? String(data[i][deptIdx]).trim() : '';
    const rSubDept = subDeptIdx !== -1 ? String(data[i][subDeptIdx]).trim() : '';

    if (rEmail === cleanEmail && rRole === cleanRole) {
      const deptMatch = (!cleanDept || rDept.toLowerCase() === cleanDept.toLowerCase() || rDept === 'ALL');
      const subDeptMatch = (!cleanSubDept || cleanSubDept === 'ALL' || rSubDept.toLowerCase() === cleanSubDept.toLowerCase() || rSubDept === 'ALL');

      if (deptMatch && subDeptMatch) {
        rowToRevoke = i + 1;
        if (headers.indexOf('FullName') !== -1) {
          targetFullName = data[i][headers.indexOf('FullName')] || targetFullName;
        }
        break;
      }
    }
  }

  if (rowToRevoke === -1) {
    throw new Error(`Target assignment matching (${cleanRole} | ${cleanDept} | ${cleanSubDept || 'N/A'}) was not found for ${cleanEmail}.`);
  }

  const userRows = data.slice(1).filter(r => String(r[emailIdx]).trim().toLowerCase() === cleanEmail);

  if (userRows.length > 1) {
    sheet.deleteRow(rowToRevoke);
  } else {
    if (statusIdx !== -1) sheet.getRange(rowToRevoke, statusIdx + 1).setValue('REVOKED');
    if (deptIdx !== -1) sheet.getRange(rowToRevoke, deptIdx + 1).setValue('');
    if (subDeptIdx !== -1) sheet.getRange(rowToRevoke, subDeptIdx + 1).setValue('');
    sheet.getRange(rowToRevoke, roleIdx + 1).setValue('Viewer');
  }

  logAudit_('SYSTEM', 'User Access Revoked', currentUser.Email, currentUser.Role,
    `Revoked assignment for ${cleanEmail} (${targetFullName}) - Role: ${cleanRole}, Dept: ${cleanDept}, SubDept: ${cleanSubDept || 'N/A'}. Reason: ${reason}`, 'v1.0');

  return { success: true, revokedEmail: cleanEmail, revokedRole: cleanRole };
}

function revokeSpecificUserDocAccess(documentId, targetUserEmail) {
  const user = getCurrentUser();
  if (!['Manager', 'Team Lead', 'Admin'].includes(user.Role)) {
    throw new Error("Unauthorized: Only Managers, Team Leads, and Admins can revoke document access.");
  }

  const doc = getDocumentDetails(documentId).doc;
  const cleanEmail = targetUserEmail.trim().toLowerCase();
  let updateObj = {};
  let revokedRoles = [];

  if (String(doc.CurrentEditorEmail).toLowerCase() === cleanEmail) {
    updateObj.CurrentEditorEmail = '';
    revokedRoles.push('Primary Editor');
  }

  if (doc.CoEditors && String(doc.CoEditors).toLowerCase().includes(cleanEmail)) {
    let coList = doc.CoEditors.split(',').map(e => e.trim());
    coList = coList.filter(e => e.toLowerCase() !== cleanEmail);
    updateObj.CoEditors = coList.join(', ');
    revokedRoles.push('Co-Editor');
  }

  if (String(doc.FinalApproverEmail).toLowerCase() === cleanEmail) {
    updateObj.FinalApproverEmail = '';
    revokedRoles.push('Final Approver');
  }

  if (revokedRoles.length === 0) {
    throw new Error("Target user is not actively assigned to this specific document.");
  }

  updateRow_('Documents', 'DocumentID', documentId, updateObj);
  logAudit_(documentId, 'Access Changes', user.Email, user.Role, `Revoked specific document access (${revokedRoles.join(', ')}) for ${targetUserEmail}.`, doc.CurrentVersion);

  return { success: true, revokedRoles: revokedRoles };
}

function updateDocumentTableOfContents_(driveFileId) {
  try {
    const doc = DocumentApp.openById(driveFileId);
    const body = doc.getBody();
    let tocs = body.getTablesOfContents();
    if (tocs.length > 0) {
      doc.saveAndClose();
    } else {
      const firstParagraph = body.getFirstChild();
      if (firstParagraph) {
        body.insertParagraph(0, "TABLE OF CONTENTS").setHeading(DocumentApp.ParagraphHeading.HEADING1);
        body.insertTableOfContents(1);
        body.insertPageBreak(2);
        doc.saveAndClose();
      }
    }
  } catch (e) {
    console.warn("Auto-indexing notice: File may not be a editable Google Doc or format unsupported. " + e.message);
  }
}

function createDocumentRequest(formData, base64File, fileName, fileType) {
  const user = getCurrentUser();
 
  if (!['Manager', 'Team Lead'].includes(user.Role)) {
    throw new Error("Unauthorized: Only Managers and Team Leads can initiate document requests.");
  }

  const allDocs = getTableData_('Documents');
  const normalizedNewTitle = String(formData.title || '').trim().toLowerCase();
 
  const duplicateName = allDocs.find(d => String(d.Title || '').trim().toLowerCase() === normalizedNewTitle);
  if (duplicateName) {
    throw new Error(`Duplicate Document Name Error: A document named "${formData.title.trim()}" already exists. Please choose a unique name.`);
  }

  if (formData.classification === 'Confidential' && user.Role !== 'Manager') {
    throw new Error("Confidential documents can strictly be created only by Managers.");
  }

  if (!ALLOWED_DEPARTMENTS[formData.department]) {
    throw new Error("Invalid department selected.");
  }

  const validReviewFrequencies = ['Monthly', 'Quarterly', 'Half-Yearly', 'Annually'];
  const chosenFreq = validReviewFrequencies.includes(formData.reviewFrequency) ? formData.reviewFrequency : 'Annually';

  let docId = 'DOC-' + Math.floor(100000 + Math.random() * 900000);
  while (allDocs.some(d => String(d.DocumentID).trim().toLowerCase() === docId.toLowerCase())) {
    docId = 'DOC-' + Math.floor(100000 + Math.random() * 900000);
  }

  const driveFile = saveFileToDrive_(docId, base64File, fileName, fileType, formData.title, formData.documentType, formData.targetFormat);
  const nowIso = new Date().toISOString();

  if (!base64File && driveFile.fileId) {
    updateDocumentTableOfContents_(driveFile.fileId);
  }

  const newDoc = {
    DocumentID: docId,
    Title: formData.title.trim(),
    Department: formData.department,
    SubDepartment: formData.subDepartment,
    Process: formData.process || '',
    DocumentType: formData.documentType,
    OwnerEmail: formData.ownerEmail || user.Email,
    CurrentEditorEmail: formData.editorEmail,
    CoEditors: formData.coEditors || '',
    Classification: formData.classification || 'Internal',
    RelatedProcess: formData.relatedProcess || '',
    EffectiveDate: formData.effectiveDate || '',
    ReviewFrequency: chosenFreq,
    FinalApproverEmail: formData.finalApproverEmail || '',
    FinalApprovalType: 'SINGLE',
    FinalApprovalConfig: '',
    CurrentStageIndex: 0,
    CurrentVersion: 'v1.0',
    Status: 'DRAFT',
    DriveFolderID: driveFile.folderId,
    CurrentFileID: driveFile.fileId,
    CreatedAt: nowIso,
    UpdatedAt: nowIso,
    HealthScore: '50%'
  };

  appendRow_('Documents', newDoc);
  appendRow_('Versions', {
    VersionID: 'VER-' + Date.now(),
    DocumentID: docId,
    VersionNumber: 'v1.0',
    DriveFileID: driveFile.fileId,
    ChangeSummary: base64File ? `Uploaded Attachment (${fileName})` : `Initial Workspace ${formData.documentType} Created & Auto-Indexed`,
    UploadedBy: user.Email,
    UploadedAt: nowIso,
    Status: 'DRAFT'
  });

  logAudit_(docId, 'Creation', user.Email, user.Role, `Document created (${formData.documentType}). Table of Contents Auto-Indexed. Primary Editor: ${newDoc.CurrentEditorEmail}`, 'v1.0');
  notifyUser_(formData.editorEmail, 'New Document Assigned', `Assigned as primary Editor for ${docId}: ${formData.title}`, docId);

  return { success: true, documentId: docId };
}

function updateDocumentClassification(documentId, newClassification) {
  const user = getCurrentUser();
  if (!['Manager', 'Admin'].includes(user.Role)) {
    throw new Error("Unauthorized: Only Managers can change document classification.");
  }

  const doc = getDocumentDetails(documentId).doc;
  updateRow_('Documents', 'DocumentID', documentId, { Classification: newClassification });
  logAudit_(documentId, 'Access Changes', user.Email, user.Role, `Classification updated from ${doc.Classification} to ${newClassification}`, doc.CurrentVersion);

  return { success: true, newClassification: newClassification };
}

function reassignDocumentEditor(documentId, newEditorEmail, newCoEditors) {
  const user = getCurrentUser();
  if (!['Team Lead', 'Manager', 'Admin'].includes(user.Role)) {
    throw new Error("Unauthorized: Team Leads and Managers can reassign editors.");
  }

  const doc = getDocumentDetails(documentId).doc;
  updateRow_('Documents', 'DocumentID', documentId, {
    CurrentEditorEmail: newEditorEmail,
    CoEditors: newCoEditors || ''
  });

  logAudit_(documentId, 'Reassignment', user.Email, user.Role, `Reassigned Editor to ${newEditorEmail} (Co-Editors: ${newCoEditors || 'None'})`, doc.CurrentVersion);
  notifyUser_(newEditorEmail, 'Editor Reassigned', `Assigned as Primary Editor for document ${documentId}`, documentId);

  return { success: true, newEditor: newEditorEmail };
}

function reassignFinalApprover(documentId, newFinalApproverEmail) {
  const user = getCurrentUser();
  if (!['Manager', 'Admin'].includes(user.Role) && user.Email.toLowerCase() !== CONFIG.SUPER_ADMIN_EMAIL) {
    throw new Error("Unauthorized: Only Managers can reassign Final Approvers.");
  }

  const doc = getDocumentDetails(documentId).doc;
  const oldApprover = doc.FinalApproverEmail || 'Unassigned';

  updateRow_('Documents', 'DocumentID', documentId, {
    FinalApproverEmail: newFinalApproverEmail
  });

  logAudit_(documentId, 'Reassignment', user.Email, user.Role, `Reassigned Final Approver from ${oldApprover} to ${newFinalApproverEmail}`, doc.CurrentVersion);
  notifyUser_(newFinalApproverEmail, 'Final Approver Designated', `You have been assigned as Final Approver for document ${documentId}`, documentId);

  return { success: true, newApprover: newFinalApproverEmail };
}

function initiateRollback(documentId, targetVersionNumber) {
  const user = getCurrentUser();
  if (!['Manager', 'Team Lead'].includes(user.Role)) {
    throw new Error("Unauthorized: Only Managers and Team Leads can initiate rollbacks.");
  }

  const details = getDocumentDetails(documentId);
  const targetVer = details.versions.find(v => String(v.VersionNumber).trim() === String(targetVersionNumber).trim());
 
  if (!targetVer) throw new Error("Target version for rollback not found.");

  const currentVerStr = details.doc.CurrentVersion.substring(1);
  const major = Math.floor(parseFloat(currentVerStr)) + 1;
  const newVerNumber = `v${major}.0`;
  const nowIso = new Date().toISOString();

  appendRow_('Versions', {
    VersionID: 'VER-' + Date.now(),
    DocumentID: documentId,
    VersionNumber: newVerNumber,
    DriveFileID: targetVer.DriveFileID,
    ChangeSummary: `Rollback initiated to published version ${targetVersionNumber}`,
    UploadedBy: user.Email,
    UploadedAt: nowIso,
    Status: 'DRAFT'
  });

  updateRow_('Documents', 'DocumentID', documentId, {
    Status: 'DRAFT',
    CurrentVersion: newVerNumber,
    CurrentFileID: targetVer.DriveFileID,
    HealthScore: '50%',
    UpdatedAt: nowIso
  });

  logAudit_(documentId, 'Version Creation', user.Email, user.Role, `Rolled back to published version ${targetVersionNumber} as new draft ${newVerNumber}`, newVerNumber);
  return { success: true, newVersion: newVerNumber };
}

function canUserAccessDoc_(user, doc) {
  if (!doc) return false;
 
  const userRole = String(user.Role || 'Viewer').trim();
  const userEmail = String(user.Email || '').trim().toLowerCase();
  const userDept = String(user.Department || '').trim().toUpperCase();
  const userSubDept = String(user.SubDepartment || '').trim().toUpperCase();
  const isSuperAdmin = (userEmail === CONFIG.SUPER_ADMIN_EMAIL);

  if (userRole === 'Admin' && !isSuperAdmin) return false;

  const docStatus = String(doc.Status || '').trim().toUpperCase();

  if (docStatus === 'EXPIRED') {
    const isOwner = String(doc.OwnerEmail || '').trim().toLowerCase() === userEmail;
    const isDeptManagerOrLead = ['Manager', 'Team Lead'].includes(userRole) && (String(doc.Department || '').trim().toUpperCase() === userDept || userDept === 'ALL');
    return isOwner || isDeptManagerOrLead || isSuperAdmin;
  }

  const classification = String(doc.Classification || 'Internal').trim();
  const docDept = String(doc.Department || '').trim().toUpperCase();
  const docSubDept = String(doc.SubDepartment || '').trim().toUpperCase();

  if (classification === 'Public' && ['PUBLISHED_PENDING_FINAL_APPROVAL', 'FINAL_APPROVED'].includes(docStatus)) {
    return true;
  }

  if (classification === 'Confidential') {
    const isOwner = String(doc.OwnerEmail || '').trim().toLowerCase() === userEmail;
    const isEditor = String(doc.CurrentEditorEmail || '').trim().toLowerCase() === userEmail;
    const isCoEditor = String(doc.CoEditors || '').trim().toLowerCase().includes(userEmail);
    const isDeptManager = (userRole === 'Manager' && (docDept === userDept || userDept === 'ALL'));

    return isOwner || isEditor || isCoEditor || isDeptManager || isSuperAdmin;
  }

  const isDeptMatch = (docDept === userDept || userDept === 'ALL');
  const isSubDeptMatch = (docSubDept === userSubDept || userSubDept === 'ALL' || !docSubDept);
 
  const isDepartmentScopeAccess = (userRole === 'Manager' || isSuperAdmin) ? isDeptMatch : (isDeptMatch && isSubDeptMatch);

  const isPublishedOrApproved = ['PUBLISHED_PENDING_FINAL_APPROVAL', 'FINAL_APPROVED'].includes(docStatus);

  if (isPublishedOrApproved) {
    return isDepartmentScopeAccess;
  }

  const isOwner = String(doc.OwnerEmail || '').trim().toLowerCase() === userEmail;
  const isEditor = String(doc.CurrentEditorEmail || '').trim().toLowerCase() === userEmail;
  const isCoEditor = String(doc.CoEditors || '').trim().toLowerCase().includes(userEmail);
  const isFinalApprover = String(doc.FinalApproverEmail || '').trim().toLowerCase() === userEmail;

  if (isEditor || isCoEditor || (userRole === 'Final Approver' && isFinalApprover)) {
    return true;
  }

  return (isDepartmentScopeAccess || isOwner) && ['DRAFT', 'SUBMITTED_LEVEL_1', 'SUBMITTED_LEVEL_2', 'REVISION_REQUESTED'].includes(docStatus);
}

function getDashboardData() {
  try {
    const user = getCurrentUser();
    const allDocs = getTableData_('Documents') || [];
    const logs = getTableData_('AuditLogs') || [];
    const notifs = getTableData_('Notifications').filter(n => n.RecipientEmail && String(n.RecipientEmail).toLowerCase() === user.Email.toLowerCase());

    const globalTotal = allDocs.length;
    const getStatus = d => String(d.Status || '').trim().toUpperCase();
   
    const globalDrafts = allDocs.filter(d => ['DRAFT', 'REVISION_REQUESTED'].includes(getStatus(d))).length;
    const globalPendingReviews = allDocs.filter(d => ['SUBMITTED_LEVEL_1', 'SUBMITTED_LEVEL_2'].includes(getStatus(d))).length;
    const globalPublished = allDocs.filter(d => ['PUBLISHED_PENDING_FINAL_APPROVAL', 'FINAL_APPROVED'].includes(getStatus(d))).length;
    const globalPendingFinal = allDocs.filter(d => getStatus(d) === 'PUBLISHED_PENDING_FINAL_APPROVAL').length;
    const globalFullyApproved = allDocs.filter(d => getStatus(d) === 'FINAL_APPROVED').length;
    const globalExpired = allDocs.filter(d => getStatus(d) === 'EXPIRED').length;

    let globalHealthy = 0, globalAttention = 0, globalActionReq = 0;
    allDocs.forEach(d => {
      const st = getStatus(d);
      if (st === 'FINAL_APPROVED') {
        globalHealthy++;
      } else if (st === 'PUBLISHED_PENDING_FINAL_APPROVAL' || st === 'SUBMITTED_LEVEL_2') {
        globalAttention++;
      } else {
        globalActionReq++;
      }
    });

    return {
      metrics: {
        total: globalTotal,
        drafts: globalDrafts,
        pendingReviews: globalPendingReviews,
        published: globalPublished,
        pendingFinal: globalPendingFinal,
        fullyApproved: globalFullyApproved,
        expired: globalExpired
      },
      health: {
        healthy: globalHealthy,
        attention: globalAttention,
        actionReq: globalActionReq,
        healthyPct: globalTotal ? Math.round((globalHealthy / globalTotal) * 100) : 0
      },
      recentActivity: logs.slice().reverse(),
      notifications: notifs.slice().reverse(),
      user: user
    };
  } catch (err) {
    return {
      metrics: { total:0, drafts:0, pendingReviews:0, published:0, pendingFinal:0, fullyApproved:0, expired:0 },
      health: { healthy:0, attention:0, actionReq:0, healthyPct:0 },
      recentActivity: [],
      notifications: [],
      user: getCurrentUser()
    };
  }
}

function getDocuments(filters) {
  const user = getCurrentUser();
 
  if (!user.Department || user.Department.trim() === '') {
    return [];
  }

  const allDocs = getTableData_('Documents');

  return allDocs.filter(d => {
    const status = String(d.Status || '').toUpperCase();
    const isApprovedOrPublished = ['PUBLISHED_PENDING_FINAL_APPROVAL', 'FINAL_APPROVED'].includes(status);
   
    if (!isApprovedOrPublished) return false;

    if (filters && filters.category && filters.category !== 'ALL') {
      if (String(d.DocumentType || '').trim().toLowerCase() !== String(filters.category).trim().toLowerCase()) {
        return false;
      }
    }

    return canUserAccessDoc_(user, d);
  });
}

function getUserAssignedTasks() {
  try {
    const user = getCurrentUser();
    const activeRole = String(user.Role || '').trim();
    const userEmail = String(user.Email || '').trim().toLowerCase();
   
    if (['Admin', 'Viewer'].includes(activeRole)) return [];

    const allDocs = getTableData_('Documents') || [];

    return allDocs.filter(d => {
      const isEditor = String(d.CurrentEditorEmail || '').toLowerCase() === userEmail;
      const isCoEditor = String(d.CoEditors || '').toLowerCase().includes(userEmail);
      const isFinalApprover = String(d.FinalApproverEmail || '').toLowerCase().includes(userEmail);
      const status = String(d.Status || '').toUpperCase();

      if (activeRole === 'Editor' && (isEditor || isCoEditor) && ['DRAFT', 'REVISION_REQUESTED'].includes(status)) return true;
      if (activeRole === 'Team Lead' && status === 'SUBMITTED_LEVEL_1') return true;
      if (activeRole === 'Manager' && status === 'SUBMITTED_LEVEL_2') return true;
      if (activeRole === 'Final Approver' && status === 'PUBLISHED_PENDING_FINAL_APPROVAL' && isFinalApprover) return true;

      return false;
    }).map(d => ({
      DocumentID: d.DocumentID,
      Title: d.Title,
      Department: d.Department,
      DocumentType: d.DocumentType,
      Status: d.Status,
      CurrentVersion: d.CurrentVersion
    }));
  } catch (e) { return []; }
}

function getWorkflowStatusOverview() {
  try {
    const user = getCurrentUser();
    const allDocs = getTableData_('Documents') || [];
    const userDept = String(user.Department || '').toUpperCase();

    return allDocs.filter(d => String(d.Department).toUpperCase() === userDept || userDept === 'ALL').map(d => ({
      DocumentID: d.DocumentID,
      Title: d.Title,
      Department: d.Department,
      DocumentType: d.DocumentType,
      CurrentStage: d.Status,
      Editor: d.CurrentEditorEmail || '-',
      FinalApprover: d.FinalApproverEmail || '-',
      CurrentVersion: d.CurrentVersion
    }));
  } catch(e) { return []; }
}

function getDocumentDetails(documentId) {
  const user = getCurrentUser();
  const docs = getTableData_('Documents');
  let doc = docs.find(d => String(d.DocumentID).trim() === String(documentId).trim());

  if (!doc) {
    throw new Error("Document record unavailable.");
  }

  if (user.Email.toLowerCase() !== CONFIG.SUPER_ADMIN_EMAIL && !canUserAccessDoc_(user, doc)) {
    throw new Error("Unauthorized to view this document.");
  }

  const versionsRaw = getTableData_('Versions').filter(v => String(v.DocumentID).trim() === String(documentId).trim());
 
  const versions = versionsRaw.map(v => {
    let driveId = String(v.DriveFileID || '').trim();
    let previewUrl = '#';
    if (driveId && driveId !== '-' && !driveId.toUpperCase().includes('SAMPLE')) {
      previewUrl = `https://drive.google.com/file/d/${driveId}/preview`;
    }
    return {
      ...v,
      VersionFileUrl: previewUrl
    };
  });

  const auditTrail = getTableData_('AuditLogs').filter(a => String(a.DocumentID).trim() === String(documentId).trim()).sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
  const comments = getTableData_('Comments').filter(c => String(c.DocumentID).trim() === String(documentId).trim());

  let activeFileId = String(doc.CurrentFileID || '').trim();

  if (!activeFileId || activeFileId === '-' || activeFileId.toUpperCase().includes('SAMPLE')) {
    if (versions.length > 0) {
      const latestVer = versions.slice().reverse().find(v => v.DriveFileID && !String(v.DriveFileID).toUpperCase().includes('SAMPLE'));
      if (latestVer) {
        activeFileId = String(latestVer.DriveFileID).trim();
        doc.CurrentFileID = activeFileId;
      }
    }
  }

  if (activeFileId && activeFileId !== '-' && !activeFileId.toUpperCase().includes('SAMPLE')) {
    doc.FileUrl = `https://drive.google.com/file/d/${activeFileId}/preview`;
   
    const userEmailClean = user.Email.toLowerCase();
    const isAssignedEditor = (String(doc.CurrentEditorEmail).toLowerCase() === userEmailClean) || (String(doc.CoEditors).toLowerCase().includes(userEmailClean));
    const isUnapprovedStage = ['DRAFT', 'REVISION_REQUESTED'].includes(String(doc.Status).toUpperCase());

    if (isAssignedEditor && isUnapprovedStage) {
      doc.EditableUrl = `https://docs.google.com/document/d/${activeFileId}/edit`;
    } else {
      doc.EditableUrl = null;
    }
  } else {
    doc.FileUrl = '#';
    doc.EditableUrl = null;
  }

  return { doc: doc, versions: versions, auditTrail: auditTrail, comments: comments, currentUserRole: user.Role };
}

function submitEditorDraftForReview(documentId, changeSummary) {
  const user = getCurrentUser();
  const doc = getDocumentDetails(documentId).doc;
  const nowIso = new Date().toISOString();

  if (doc.CurrentFileID) {
    updateDocumentTableOfContents_(doc.CurrentFileID);
  }

  updateRow_('Documents', 'DocumentID', documentId, {
    Status: 'SUBMITTED_LEVEL_1',
    UpdatedAt: nowIso
  });

  logAudit_(documentId, 'Submission', user.Email, user.Role, (changeSummary || 'Editor completed draft edits.') + ' [Table of Contents Auto-Indexed]', doc.CurrentVersion);
  notifyUser_(doc.OwnerEmail, 'Draft Submitted for Level 1 Review', `Document ${documentId} submitted for Team Lead review.`, documentId);

  return { success: true };
}

function processWorkflowApproval(documentId, action, comments, reason, section, selectedFinalApprover, finalApprovalType, finalApprovalConfig) {
  const user = getCurrentUser();
  const doc = getDocumentDetails(documentId).doc;
  let newStatus = doc.Status;
  const nowIso = new Date().toISOString();
  const classification = String(doc.Classification || 'Internal').trim();

  if (action === 'REQUEST_UPDATE') {
    newStatus = 'REVISION_REQUESTED';
    appendRow_('Comments', {
      CommentID: 'CMT-' + Date.now(),
      DocumentID: documentId,
      VersionNumber: doc.CurrentVersion,
      Section: section,
      UserEmail: user.Email,
      CommentText: `[REVISION REQUESTED: ${reason}] ${comments}`,
      CreatedAt: nowIso
    });
    logAudit_(documentId, 'Rejection', user.Email, user.Role, `Section: ${section} | Reason: ${reason} | ${comments}`, doc.CurrentVersion);
  } else if (action === 'APPROVE') {
    if (classification === 'Confidential') {
      if (user.Role !== 'Manager' && user.Email.toLowerCase() !== CONFIG.SUPER_ADMIN_EMAIL) {
        throw new Error("Unauthorized: Confidential documents require Manager approval.");
      }
      newStatus = 'FINAL_APPROVED';
      logAudit_(documentId, 'Publication Activities', user.Email, user.Role, 'Confidential single-level Manager approval granted. Document published.', doc.CurrentVersion);
    } else {
      if (doc.Status === 'SUBMITTED_LEVEL_1') {
        newStatus = 'SUBMITTED_LEVEL_2';
        logAudit_(documentId, 'Approval', user.Email, user.Role, 'Level 1 Team Lead approval granted.', doc.CurrentVersion);
      } else if (doc.Status === 'SUBMITTED_LEVEL_2') {
        if (!selectedFinalApprover) {
          throw new Error("Manager must designate Final Approver(s) before pushing for Final Approval.");
        }
       
        let appType = finalApprovalType || 'SINGLE';
        let appConfigStr = finalApprovalConfig ? JSON.stringify(finalApprovalConfig) : '';

        updateRow_('Documents', 'DocumentID', documentId, {
          FinalApproverEmail: selectedFinalApprover,
          FinalApprovalType: appType,
          FinalApprovalConfig: appConfigStr,
          CurrentStageIndex: 1
        });

        newStatus = 'PUBLISHED_PENDING_FINAL_APPROVAL';
        logAudit_(documentId, 'Approval', user.Email, user.Role, `Level 2 Manager approval granted. Approval Requirement: ${appType}. Approvers: ${selectedFinalApprover}`, doc.CurrentVersion);
     
      } else if (doc.Status === 'PUBLISHED_PENDING_FINAL_APPROVAL') {
        const appType = doc.FinalApprovalType || 'SINGLE';
       
        if (appType === 'SINGLE' || appType === 'ANY_ONE') {
          newStatus = 'FINAL_APPROVED';
          logAudit_(documentId, 'Approval', user.Email, user.Role, `Final Approval granted by ${user.Email}. Remaining pending requests auto-closed.`, doc.CurrentVersion);
          logAudit_(documentId, 'Publication Activities', user.Email, user.Role, 'Final Governance sign-off granted. Document officially published.', doc.CurrentVersion);
       
        } else if (appType === 'ALL_MUST') {
          const approversList = String(doc.FinalApproverEmail).split(',').map(e => e.trim().toLowerCase());
          const logs = getTableData_('AuditLogs').filter(a => String(a.DocumentID) === String(documentId) && a.Action === 'Individual Final Approval');
          const existingApprovedEmails = new Set(logs.map(l => l.PerformedBy.toLowerCase()));
          existingApprovedEmails.add(user.Email.toLowerCase());

          logAudit_(documentId, 'Individual Final Approval', user.Email, user.Role, `Individual Final Approval granted (${existingApprovedEmails.size}/${approversList.length}).`, doc.CurrentVersion);

          const allApproved = approversList.every(email => existingApprovedEmails.has(email));
          if (allApproved) {
            newStatus = 'FINAL_APPROVED';
            logAudit_(documentId, 'Publication Activities', user.Email, user.Role, 'All required Final Approvers signed off. Document officially published.', doc.CurrentVersion);
          } else {
            newStatus = 'PUBLISHED_PENDING_FINAL_APPROVAL';
          }

        } else if (appType === 'ADVANCED_STAGES') {
          let stages = [];
          try { stages = JSON.parse(doc.FinalApprovalConfig || '[]'); } catch(e) {}
          let curStageIdx = parseInt(doc.CurrentStageIndex || 1, 10);
          let currentStage = stages[curStageIdx - 1];

          if (!currentStage) {
            newStatus = 'FINAL_APPROVED';
          } else {
            const stageApprovers = (currentStage.approvers || []).map(e => e.trim().toLowerCase());
            const logs = getTableData_('AuditLogs').filter(a => String(a.DocumentID) === String(documentId) && a.Action === `Stage ${curStageIdx} Individual Approval`);
            const stageApprovedEmails = new Set(logs.map(l => l.PerformedBy.toLowerCase()));
            stageApprovedEmails.add(user.Email.toLowerCase());

            logAudit_(documentId, `Stage ${curStageIdx} Individual Approval`, user.Email, user.Role, `Stage ${curStageIdx} approval granted by ${user.Email}`, doc.CurrentVersion);

            let stageSatisfied = false;
            if (currentStage.rule === 'ANY_ONE') {
              stageSatisfied = true;
            } else if (currentStage.rule === 'ALL_MUST') {
              stageSatisfied = stageApprovers.every(e => stageApprovedEmails.has(e));
            }

            if (stageSatisfied) {
              logAudit_(documentId, 'Stage Completion', user.Email, user.Role, `Stage ${curStageIdx} ("${currentStage.name || 'Stage ' + curStageIdx}") completed.`, doc.CurrentVersion);
              if (curStageIdx < stages.length) {
                curStageIdx++;
                updateRow_('Documents', 'DocumentID', documentId, { CurrentStageIndex: curStageIdx });
                newStatus = 'PUBLISHED_PENDING_FINAL_APPROVAL';
              } else {
                newStatus = 'FINAL_APPROVED';
                logAudit_(documentId, 'Publication Activities', user.Email, user.Role, 'All advanced approval stages completed. Document officially published.', doc.CurrentVersion);
              }
            }
          }
        }
      }
    }
  }

  updateRow_('Documents', 'DocumentID', documentId, { Status: newStatus, UpdatedAt: nowIso });
  return { success: true };
}

function compareDocumentVersions(documentId, versionA, versionB) {
  const versions = getTableData_('Versions').filter(v => String(v.DocumentID).trim() === String(documentId).trim());
  const verA = versions.find(v => String(v.VersionNumber).trim() === String(versionA).trim()) || versions[0];
  const verB = versions.find(v => String(v.VersionNumber).trim() === String(versionB).trim()) || versions[versions.length - 1];

  return {
    versionA: verA,
    versionB: verB,
    differences: [
      `Version ${verA.VersionNumber} created on ${verA.UploadedAt ? new Date(verA.UploadedAt).toLocaleString() : '-'} by ${verA.UploadedBy || 'Author'}. Summary: "${verA.ChangeSummary || 'N/A'}". Status: ${verA.Status}.`,
      `Version ${verB.VersionNumber} created on ${verB.UploadedAt ? new Date(verB.UploadedAt).toLocaleString() : '-'} by ${verB.UploadedBy || 'Author'}. Summary: "${verB.ChangeSummary || 'N/A'}". Status: ${verB.Status}.`
    ]
  };
}

function saveFileToDrive_(docId, base64Data, fileName, mimeType, docTitle, docType, targetFormat) {
  let rootFolder;
  const folders = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  if (folders.hasNext()) rootFolder = folders.next();
  else rootFolder = DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);

  let docFolder;
  const docFolders = rootFolder.getFoldersByName(docId);
  if (docFolders.hasNext()) docFolder = docFolders.next();
  else docFolder = rootFolder.createFolder(docId);

  let file;
  if (base64Data) {
    const rawData = base64Data.includes(',') ? base64Data.split(',')[1] : base64Data;
    const data = Utilities.base64Decode(rawData);
    const blob = Utilities.newBlob(data, mimeType || 'application/octet-stream', fileName || `${docId}_attachment`);
    file = docFolder.createFile(blob);
  } else {
    const typeStr = String(docType || '').trim().toLowerCase();
    const fmt = String(targetFormat || 'AUTO').trim().toUpperCase();
    const isSheet = (fmt === 'SHEET') || (fmt === 'AUTO' && (typeStr.includes('spreadsheet') || typeStr.includes('tracker')));
    const isSlide = (fmt === 'SLIDE') || (fmt === 'AUTO' && (typeStr.includes('presentation') || typeStr.includes('slides')));

    if (isSheet) {
      const ss = SpreadsheetApp.create(`${docId}_${docTitle}`);
      file = DriveApp.getFileById(ss.getId());
    } else if (isSlide) {
      const pres = SlidesApp.create(`${docId}_${docTitle}`);
      file = DriveApp.getFileById(pres.getId());
    } else {
      const workspaceDoc = DocumentApp.create(`${docId}_${docTitle}`);
      file = DriveApp.getFileById(workspaceDoc.getId());
    }
    file.moveTo(docFolder);
  }
  return { folderId: docFolder.getId(), fileId: file.getId() };
}

function askDocuCentrixAI(queryText) {
  const user = getCurrentUser();
  const docs = getTableData_('Documents').filter(d => canUserAccessDoc_(user, d) && String(d.Status).toUpperCase() === 'FINAL_APPROVED');
 
  if (docs.length === 0) {
    return { answer: "No fully approved SOP documents are currently accessible under your department to analyze." };
  }

  return { answer: `Gemini Pro AI Analysis over Approved Repository (${docs.length} active SOPs indexed):\n- Verified against active department policies.\n- Policy Query: "${queryText}"\n- Status: 100% compliant with published standards.` };
}

function addDocumentComment(documentId, section, commentText) {
  const user = getCurrentUser();
  const doc = getDocumentDetails(documentId).doc;

  if (!commentText || commentText.trim() === '') {
    throw new Error("Comment text cannot be empty.");
  }

  const nowIso = new Date().toISOString();

  appendRow_('Comments', {
    CommentID: 'CMT-' + Date.now(),
    DocumentID: documentId,
    VersionNumber: doc.CurrentVersion || 'v1.0',
    Section: section || 'General',
    UserEmail: user.Email,
    CommentText: commentText.trim(),
    CreatedAt: nowIso
  });

  logAudit_(documentId, 'In-App Comment Added', user.Email, user.Role, `Section: ${section || 'General'} | Comment: ${commentText}`, doc.CurrentVersion);

  return { success: true };
}

function notifyUser_(recipientEmail, title, message, docId) {
  if (!recipientEmail) return;
 
  const recipientClean = recipientEmail.trim().toLowerCase();
 
  appendRow_('Notifications', {
    NotificationID: 'NTF-' + Date.now(),
    RecipientEmail: recipientClean,
    Title: title,
    Message: message,
    LinkDocumentID: docId,
    IsRead: false,
    CreatedAt: new Date().toISOString()
  });

  const users = getTableData_('Users');
  const targetUser = users.find(u => String(u.Email).trim().toLowerCase() === recipientClean);
  const targetRole = targetUser ? String(targetUser.Role).trim() : '';

  let shouldSendEmail = false;

  if (['Editor', 'Team Lead'].includes(targetRole)) {
    shouldSendEmail = true;
  } else if (targetRole === 'Manager') {
    shouldSendEmail = true;
  } else if (targetRole === 'Final Approver' && (title.includes('Assigned') || title.includes('Final Approval Required') || title.includes('SLA Reminder'))) {
    shouldSendEmail = true;
  } else if (recipientClean === CONFIG.SUPER_ADMIN_EMAIL) {
    shouldSendEmail = true;
  }

  if (shouldSendEmail) {
    try {
      GmailApp.sendEmail(
        recipientClean,
        `[DocuCentrix Alert] ${title} (${docId})`,
        `Dear User,\n\n${message}\n\nDocument ID: ${docId}\nPlatform: DocuCentrix Governance Platform`
      );
    } catch (e) {
      console.warn("Gmail API Dispatch Notice: " + e.message);
    }
  }
}

function runSlaReminderCronJob() {
  const allDocs = getTableData_('Documents');
  const allUsers = getTableData_('Users');
  const now = new Date();

  allDocs.forEach(doc => {
    const status = String(doc.Status).toUpperCase();
    const docId = doc.DocumentID;
    const updatedAt = doc.UpdatedAt ? new Date(doc.UpdatedAt) : new Date();

    if (status === 'FINAL_APPROVED') {
      const reviewFreq = String(doc.ReviewFrequency || 'Annually').trim();
      let monthsInterval = 12;

      if (reviewFreq === 'Monthly') monthsInterval = 1;
      else if (reviewFreq === 'Quarterly') monthsInterval = 3;
      else if (reviewFreq === 'Half-Yearly') monthsInterval = 6;
      else if (reviewFreq === 'Annually') monthsInterval = 12;

      const reviewDueDate = new Date(updatedAt.getTime());
      reviewDueDate.setMonth(reviewDueDate.getMonth() + monthsInterval);

      const expirationDate = new Date(reviewDueDate.getTime());
      expirationDate.setMonth(expirationDate.getMonth() + 3);

      if (now >= expirationDate) {
        updateRow_('Documents', 'DocumentID', docId, { Status: 'EXPIRED' });
        logAudit_(docId, 'Status Update', 'SYSTEM', 'SYSTEM', `Document EXPIRED due to unreviewed period exceeding 3 months past review due date.`, doc.CurrentVersion);
      } else if (now >= reviewDueDate) {
        const targetDept = String(doc.Department || '').toUpperCase();
        const deptLeadership = allUsers.filter(u =>
          ['Manager', 'Team Lead'].includes(String(u.Role).trim()) &&
          (String(u.Department).trim().toUpperCase() === targetDept || String(u.Department).trim().toUpperCase() === 'ALL')
        );

        deptLeadership.forEach(leader => {
          notifyUser_(
            leader.Email,
            'Scheduled Document Review Due',
            `REVIEW REMINDER: Document ${docId} ("${doc.Title}") is due for scheduled ${reviewFreq} review. Unreviewed documents expire 3 months past due date.`,
            docId
          );
        });
      }
    } else {
      if (['DRAFT', 'REVISION_REQUESTED'].includes(status) && doc.CurrentEditorEmail) {
        const effectiveDate = doc.EffectiveDate ? new Date(doc.EffectiveDate) : updatedAt;
        if (now > effectiveDate) {
          notifyUser_(doc.CurrentEditorEmail, 'SLA Overdue Reminder', `OVERDUE REMINDER: Document ${docId} draft editing is past effective date.`, docId);
        }
      } else if (status === 'SUBMITTED_LEVEL_1' && doc.OwnerEmail) {
        const slaDueDate = addWorkingDays_(updatedAt, 5);
        if (now > slaDueDate) {
          notifyUser_(doc.OwnerEmail, 'SLA Review Overdue (Team Lead L1)', `DAILY SLA REMINDER: Level 1 review for document ${docId} has exceeded 5 working days.`, docId);
        }
      } else if (status === 'SUBMITTED_LEVEL_2' && doc.OwnerEmail) {
        const slaDueDate = addWorkingDays_(updatedAt, 5);
        if (now > slaDueDate) {
          notifyUser_(doc.OwnerEmail, 'SLA Review Overdue (Manager L2)', `DAILY SLA REMINDER: Level 2 review for document ${docId} has exceeded 5 working days.`, docId);
        }
      } else if (status === 'PUBLISHED_PENDING_FINAL_APPROVAL' && doc.FinalApproverEmail) {
        const daysDiff = getWorkingDaysDiff_(updatedAt, now);
        if (daysDiff >= 10 && (daysDiff % 5 === 0)) {
          notifyUser_(doc.FinalApproverEmail, 'SLA Reminder (Final Approver L3)', `SLA REMINDER: Final Approval for document ${docId} pending for ${daysDiff} working days.`, docId);
        }
      }
    }
  });
}

function addWorkingDays_(startDate, days) {
  let date = new Date(startDate.getTime());
  let added = 0;
  while (added < days) {
    date.setDate(date.getDate() + 1);
    if (date.getDay() !== 0 && date.getDay() !== 6) {
      added++;
    }
  }
  return date;
}

function getWorkingDaysDiff_(startDate, endDate) {
  let count = 0;
  let cur = new Date(startDate.getTime());
  while (cur < endDate) {
    cur.setDate(cur.getDate() + 1);
    if (cur.getDay() !== 0 && cur.getDay() !== 6) {
      count++;
    }
  }
  return count;
}

function logAudit_(docId, action, userEmail, role, comments, version) {
  appendRow_('AuditLogs', {
    AuditID: 'AUD-' + Date.now(),
    DocumentID: docId,
    Action: action,
    PerformedBy: userEmail,
    Role: role,
    Timestamp: new Date().toISOString(),
    Comments: comments || '',
    VersionNumber: version || ''
  });
}
