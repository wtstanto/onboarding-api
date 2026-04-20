// Auntie Anne's Onboarding — Google Apps Script Webhook
// -------------------------------------------------------
// Deploy as Web App: Execute as Me, Who has access: Anyone
// After updating this code, go to Deploy → Manage deployments
// → edit → change to "New version" → Deploy
//
// Set GAS_SECRET below to match the GAS_SECRET env var in Railway.

const GAS_SECRET = 'don';   // ← change this if you change it in Railway
const SHEET_NAME = 'Sheet1';

// Sheet column layout (1-indexed):
// A(1)  submittedAt       B(2)  firstName       C(3)  lastName
// D(4)  email             E(5)  phone            F(6)  ssn (masked)
// G(7)  dob               H(8)  address1         I(9)  city
// J(10) state             K(11) zip              L(12) i9Status
// M(13) zipDriveUrl       N(14) i9Doc            O(15) i9DocNumber
// P(16) i9Issuer          Q(17) i9ExpDate        R(18) i9VerifiedDate
// S(19) i9VerifiedBy      T(20) i9FileId         U(21) startDate
// V(22) ecName            W(23) ecRelationship   X(24) ecPhone
// Y(25) overallStatus

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.secret !== GAS_SECRET) return json({ error: 'Unauthorized' });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    // ── Log new employee submission ───────────────────────────────────────
    if (data.action === 'log') {
      sheet.appendRow([
        data.submittedAt    || new Date().toISOString(), // A
        data.firstName      || '',  // B
        data.lastName       || '',  // C
        data.email          || '',  // D
        data.phone          || '',  // E
        data.ssn            || '',  // F
        data.dob            || '',  // G
        data.address1       || '',  // H
        data.city           || '',  // I
        data.state          || '',  // J
        data.zip            || '',  // K
        'pending',                  // L  i9Status
        data.zipDriveUrl    || '',  // M  zip Drive URL
        '', '', '', '', '', '',     // N-S  i9 fields (filled on completion)
        data.i9FileId       || '',  // T  individual I-9 PDF Drive file ID
        data.startDate      || '',  // U  employee start date
        data.ecName         || '',  // V  emergency contact name
        data.ecRelationship || '',  // W  emergency contact relationship
        data.ecPhone        || '',  // X  emergency contact phone
        'new',                      // Y  overallStatus
      ]);
      return json({ status: 'ok', rowId: sheet.getLastRow() });
    }

    // ── Return all employees ──────────────────────────────────────────────
    if (data.action === 'getAll') {
      const rows = sheet.getDataRange().getValues();
      const safeDate = (v) => { try { const d = new Date(v); return isNaN(d.getTime()) ? '' : d.toISOString(); } catch(e) { return ''; } };
      // Use reduce to preserve the original row index (= actual 1-based sheet row number)
      // so that rowId used in completeI9 / updateStatus always targets the right sheet row.
      const employees = rows.reduce((acc, row, i) => {
        // Skip header row (first cell is a non-date string like "submittedAt")
        if (i === 0 && row[0] && isNaN(new Date(row[0]).getTime())) return acc;
        while (row.length < 25) row.push('');
        const i9Complete = (row[11] || 'pending') === 'complete';
        acc.push({
          id:             i + 1,          // actual 1-based sheet row number
          submittedAt:    row[0] ? safeDate(row[0]) : '',
          firstName:      row[1]  || '',
          lastName:       row[2]  || '',
          email:          row[3]  || '',
          phone:          row[4]  || '',
          ssn:            row[5]  || '',
          dob:            row[6]  ? formatDate(row[6]) : '',
          address1:       row[7]  || '',
          city:           row[8]  || '',
          state:          row[9]  || '',
          zip:            String(row[10] || ''),
          i9Status:       row[11] || 'pending',
          driveUrl:       row[12] || '',
          i9FileId:       row[19] || '',
          startDate:      row[20] ? formatDate(row[20]) : '',
          ecName:         row[21] || '',
          ecRelationship: row[22] || '',
          ecPhone:        row[23] || '',
          overallStatus:  row[24] || 'new',
          i9s2: i9Complete ? {
            docTitle:     row[13] || '',
            docNumber:    row[14] || '',
            issuer:       row[15] || '',
            expDate:      row[16] || '',
            verifiedDate: row[17] ? safeDate(row[17]) : '',
            empName:      row[18] || '',
          } : null,
        });
        return acc;
      }, []);
      return json({ employees });
    }

    // ── Return a single row (for I-9 Section 2 workflow) ─────────────────
    if (data.action === 'getRow') {
      const rowId = parseInt(data.rowId);
      const lastCol = Math.max(sheet.getLastColumn(), 25);
      const row = sheet.getRange(rowId, 1, 1, lastCol).getValues()[0];
      return json({ row });
    }

    // ── Update overall status ─────────────────────────────────────────────
    if (data.action === 'updateStatus') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      sheet.getRange(rowId, 25).setValue(data.overallStatus || 'new');  // Y
      return json({ status: 'ok' });
    }

    // ── Mark I-9 complete ─────────────────────────────────────────────────
    if (data.action === 'completeI9') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      const lastCol = Math.max(sheet.getLastColumn(), 21);
      const existing = sheet.getRange(rowId, 1, 1, lastCol).getValues()[0];
      // Update columns L–U (12–21)
      sheet.getRange(rowId, 12, 1, 10).setValues([[
        'complete',                                 // L  i9Status
        existing[12] || '',                         // M  preserve zip URL
        data.docTitle  || '',                       // N
        data.docNumber || '',                       // O
        data.issuer    || '',                       // P
        data.expDate   || '',                       // Q
        new Date().toISOString(),                   // R  verifiedDate
        data.empName   || '',                       // S
        data.newI9FileId || existing[19] || '',     // T  updated I-9 file ID
        existing[20] || '',                         // U  preserve startDate
      ]]);
      return json({ status: 'ok' });
    }

    // ── Create a subfolder in Drive ──────────────────────────────────────
    if (data.action === 'createFolder') {
      const parent = DriveApp.getFolderById(data.parentFolderId);
      const folder = parent.createFolder(data.folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return json({ folderId: folder.getId(), url: folder.getUrl() });
    }

    // ── Upload a file to Drive ────────────────────────────────────────────
    if (data.action === 'uploadFile') {
      const folder = DriveApp.getFolderById(data.folderId);
      const bytes  = Utilities.base64Decode(data.fileData);
      const blob   = Utilities.newBlob(bytes, data.mimeType || 'application/octet-stream', data.filename);
      const file   = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return json({ fileId: file.getId(), url: file.getUrl() });
    }

    // ── Download a file from Drive as base64 ─────────────────────────────
    if (data.action === 'getFile') {
      const file  = DriveApp.getFileById(data.fileId);
      const bytes = file.getBlob().getBytes();
      return json({ fileData: Utilities.base64Encode(bytes) });
    }

    // ── Combined: get row data + download its I-9 file in one round trip ────
    // Replaces separate getRow + getFile calls to cut I-9 completion time ~50%.
    if (data.action === 'getRowAndFile') {
      const rowId   = parseInt(data.rowId);
      const lastCol = Math.max(sheet.getLastColumn(), 21);
      let actualRowId = rowId;
      let row = sheet.getRange(rowId, 1, 1, lastCol).getValues()[0];
      // Header-row offset detection (same logic as Flask side)
      if (row[0] && isNaN(new Date(String(row[0])).getTime())) {
        actualRowId = rowId + 1;
        row = sheet.getRange(actualRowId, 1, 1, lastCol).getValues()[0];
      }
      const i9FileId = (row[19] || '').toString().trim();
      let fileData = null;
      if (i9FileId) {
        const file  = DriveApp.getFileById(i9FileId);
        const bytes = file.getBlob().getBytes();
        fileData = Utilities.base64Encode(bytes);
      }
      return json({ row: row, actualRowId: actualRowId, fileId: i9FileId, fileData: fileData });
    }

    // ── Combined: replace I-9 file + update sheet in one round trip ──────────
    if (data.action === 'replaceAndCompleteI9') {
      const oldFile = DriveApp.getFileById(data.fileId);
      const name    = data.filename || oldFile.getName();
      const parents = oldFile.getParents();
      const parent  = parents.next();
      oldFile.setTrashed(true);
      const bytes   = Utilities.base64Decode(data.fileData);
      const blob    = Utilities.newBlob(bytes, 'application/pdf', name);
      const newFile = parent.createFile(blob);
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const newFileId = newFile.getId();
      // Update sheet row
      const rowId   = parseInt(data.rowId);
      const lastCol = Math.max(sheet.getLastColumn(), 21);
      const existing = sheet.getRange(rowId, 1, 1, lastCol).getValues()[0];
      sheet.getRange(rowId, 12, 1, 10).setValues([[
        'complete',
        existing[12] || '',
        data.docTitle  || '',
        data.docNumber || '',
        data.issuer    || '',
        data.expDate   || '',
        new Date().toISOString(),
        data.empName   || '',
        newFileId,
        existing[20] || '',
      ]]);
      return json({ fileId: newFileId, url: newFile.getUrl(), status: 'ok' });
    }

    // ── Replace a Drive file with updated content ─────────────────────────
    if (data.action === 'replaceFile') {
      const oldFile = DriveApp.getFileById(data.fileId);
      const name    = data.filename || oldFile.getName();
      const parents = oldFile.getParents();
      const parent  = parents.next();
      oldFile.setTrashed(true);
      const bytes   = Utilities.base64Decode(data.fileData);
      const blob    = Utilities.newBlob(bytes, 'application/pdf', name);
      const newFile = parent.createFile(blob);
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return json({ fileId: newFile.getId(), url: newFile.getUrl() });
    }

    // ── Insert header row ─────────────────────────────────────────────────
    if (data.action === 'setupHeaders') {
      const headers = [
        'Submitted At', 'First Name', 'Last Name', 'Email', 'Phone', 'SSN',
        'Date of Birth', 'Address', 'City', 'State', 'ZIP',
        'I-9 Status', 'Drive Folder URL',
        'I-9 Doc Title', 'I-9 Doc Number', 'I-9 Issuer', 'I-9 Exp Date',
        'I-9 Verified Date', 'I-9 Verified By', 'I-9 File ID',
        'Start Date', 'EC Name', 'EC Relationship', 'EC Phone', 'Overall Status',
      ];
      // Only insert if row 1 is not already a header
      const first = sheet.getRange(1, 1).getValue();
      if (!first || !isNaN(new Date(first).getTime())) {
        sheet.insertRowBefore(1);
      }
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      return json({ status: 'ok' });
    }

    return json({ error: 'Unknown action' });

  } catch (err) {
    return json({ error: err.message });
  }
}

function doGet(e) {
  return json({ status: 'ok' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDate(val) {
  if (!val) return '';
  try {
    const d = new Date(val);
    return isNaN(d) ? String(val) : d.toISOString().split('T')[0];
  } catch (e) {
    return String(val);
  }
}
