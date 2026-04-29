// Auntie Anne's Onboarding — Google Apps Script Webhook
// -------------------------------------------------------
// Deploy as Web App: Execute as Me, Who has access: Anyone
// After updating this code, go to Deploy → Manage deployments
// → edit → change to "New version" → Deploy
//
// Set GAS_SECRET below to match the GAS_SECRET env var in Railway.

const GAS_SECRET = 'Lq0-v_bj2kdZyqRG23vhmMn3OVdbOgr9qszYgCUG8kM';
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
// -- /de112b additions --
// Z(26)  gender                   AA(27) payRate
// AB(28) position                 AC(29) location
// AD(30) deptCode                 AE(31) hireDate (confirmed)
// AF(32) humanityCsvAt            AG(33) quCsvAt
// AH(34) zignalCsvAt              AI(35) adpSheetAt
// AJ(36) adpCsvAt
// -- /de112c additions --
// AS(45) tshirtSize

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
        data.gender         || '',  // Z(26)  gender
        '', '', '', '', '',         // AA–AE(27–31)  employment (filled by admin)
        '', '', '', '', '',         // AF–AJ(32–36)  export timestamps
        '', '', '', '',             // AK–AN(37–40)  working papers + status
        '', '', '',                 // AO–AQ(41–43)  lifecycle timestamps + reason
        '',                         // AR(44)  driveFolderId
        data.tshirtSize     || '',  // AS(45)  t-shirt size
        data.i9s1docs       || '',  // AT(46)  I-9 Section 1 doc data (JSON)
        data.testEntry ? 'TRUE' : '',  // AU(47)  test/demo submission flag
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
        while (row.length < 47) row.push('');
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
          // /de112b additions
          gender:         row[25] || '',
          tshirtSize:     row[44] || '',
          i9s1docs:       row[45] ? (() => { try { return JSON.parse(row[45]); } catch(e) { return null; } })() : null,
          testEntry:      row[46] === 'TRUE' || row[46] === true,
          payRate:        row[26] || '',
          position:       row[27] || '',
          location:       row[28] || '',
          deptCode:       row[29] || '',
          hireDate:       row[30] ? formatDate(row[30]) : '',
          humanityCsvAt:  row[31] ? safeDate(row[31]) : '',
          quCsvAt:        row[32] ? safeDate(row[32]) : '',
          zignalCsvAt:    row[33] ? safeDate(row[33]) : '',
          adpSheetAt:     row[34] ? safeDate(row[34]) : '',
          adpCsvAt:       row[35] ? safeDate(row[35]) : '',
          // /de112b: working papers (under-18)
          workingPapersGivenAt:    row[36] ? safeDate(row[36]) : '',
          workingPapersReturnedAt: row[37] ? safeDate(row[37]) : '',
          workingPapersFileId:     row[38] || '',
          // /de112b: lifecycle status
          status:          row[39] || 'onboarding',
          activatedAt:     row[40] ? safeDate(row[40]) : '',
          inactiveAt:      row[41] ? safeDate(row[41]) : '',
          inactiveReason:  row[42] || '',
          // /de112b: explicit Drive folder ID (set during folder creation, or
          // populated for legacy/imported employees pointing at a shared folder)
          driveFolderId:   row[43] || '',
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

    // ── Send welcome email via MailApp ────────────────────────────────────
    if (data.action === 'sendEmail') {
      MailApp.sendEmail({
        to:      data.to,
        subject: data.subject,
        body:    data.body,
        name:    data.fromName || 'Auntie Anne\'s',
        replyTo: data.replyTo  || '',
      });
      return json({ status: 'ok' });
    }

    // ── Create a subfolder in Drive ──────────────────────────────────────
    if (data.action === 'createFolder') {
      const parent = DriveApp.getFolderById(data.parentFolderId);
      const folder = parent.createFolder(data.folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      // Inherit editors and viewers from parent folder
      parent.getEditors().forEach(u => folder.addEditor(u));
      parent.getViewers().forEach(u => folder.addViewer(u));
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
        // /de112b additions
        'Gender',                                // Z
        'Pay Rate', 'Position', 'Location',      // AA–AC
        'Dept Code', 'Hire Date (confirmed)',    // AD, AE
        'Humanity CSV At', 'Qu CSV At',          // AF, AG
        'Zygnal CSV At',                         // AH
        'ADP Cheat Sheet At', 'ADP CSV At',      // AI, AJ
        // /de112c additions (AK–AS are working papers, lifecycle, Drive folder, t-shirt)
        'WP Given At', 'WP Returned At', 'WP File ID',  // AK, AL, AM
        'Status', 'Activated At', 'Inactive At', 'Inactive Reason',  // AN, AO, AP, AQ
        'Drive Folder ID',                       // AR
        'T-Shirt Size',                          // AS
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

    // ── /de112b: update employment details (pay rate, position, etc.) ─────
    // Columns AA(27), AB(28), AC(29), AD(30), AE(31)
    if (data.action === 'updateEmployment') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      sheet.getRange(rowId, 27, 1, 5).setValues([[
        data.payRate  || '',  // AA
        data.position || '',  // AB
        data.location || '',  // AC
        data.deptCode || '',  // AD
        data.hireDate || '',  // AE
      ]]);
      return json({ status: 'ok' });
    }

    // ── /de112b: stamp an export-download timestamp ───────────────────────
    // Called by Flask right after serving an export. col is 1-based.
    if (data.action === 'stampExport') {
      const rowId = parseInt(data.rowId);
      const col   = parseInt(data.col);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      if (!col   || col   < 1) return json({ error: 'Invalid col' });
      sheet.getRange(rowId, col).setValue(data.timestamp || new Date().toISOString());
      return json({ status: 'ok' });
    }

    // ── /de112b: get the parent folder of a Drive file ────────────────────
    // Used to find the employee's individual subfolder so CSV exports can be
    // saved alongside their PDFs.
    if (data.action === 'getFileParent') {
      try {
        const file = DriveApp.getFileById(data.fileId);
        const parents = file.getParents();
        if (!parents.hasNext()) return json({ error: 'No parent folder' });
        const parent = parents.next();
        return json({ folderId: parent.getId(), url: parent.getUrl() });
      } catch (err) {
        return json({ error: 'File not found: ' + err.message });
      }
    }

    // ── /de112b: get an employee's Drive folder URL by row ID ─────────────
    // Tries the cached driveFolderId column (AR/44) first; falls back to
    // deriving from the I-9 file's parent (column T/20). Used by the admin
    // "Open Drive folder" button.
    if (data.action === 'getEmployeeFolderUrl') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      try {
        const lastCol = sheet.getLastColumn();
        const numCols = Math.max(44, lastCol);
        const row = sheet.getRange(rowId, 1, 1, numCols).getValues()[0];
        // Pad in case sheet has fewer than 44 columns
        while (row.length < 44) row.push('');
        // Try cached folder ID first (AR / index 43)
        const cachedFolderId = (row[43] || '').toString().trim();
        if (cachedFolderId) {
          try {
            const folder = DriveApp.getFolderById(cachedFolderId);
            return json({ folderId: cachedFolderId, url: folder.getUrl(), source: 'cached' });
          } catch (err) {
            // Fall through to I-9 parent derivation
          }
        }
        // Fallback: derive from I-9 file's parent (T / index 19)
        const i9FileId = (row[19] || '').toString().trim();
        if (!i9FileId) return json({ error: 'No Drive folder for this employee (no folder ID cached, no I-9 file on record)' });
        const file = DriveApp.getFileById(i9FileId);
        const parents = file.getParents();
        if (!parents.hasNext()) return json({ error: 'I-9 file has no parent folder' });
        const parent = parents.next();
        return json({ folderId: parent.getId(), url: parent.getUrl(), source: 'derived' });
      } catch (err) {
        return json({ error: 'Could not resolve folder: ' + err.message });
      }
    }

    // ── /de112b: set working papers given/returned timestamps ─────────────
    // Toggle a boolean field for a row. If `value` is true (or omitted), writes
    // current ISO timestamp; if false, clears the cell.
    //   field='given'    → column AK (37)
    //   field='returned' → column AL (38)
    if (data.action === 'setWorkingPapers') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      const colByField = { given: 37, returned: 38 };
      const col = colByField[data.field];
      if (!col) return json({ error: 'Invalid field — must be given or returned' });
      const stamp = (data.value === false) ? '' : (data.timestamp || new Date().toISOString());
      sheet.getRange(rowId, col).setValue(stamp);
      return json({ status: 'ok' });
    }

    // ── /de112b: upload a working papers photo to the employee's Drive folder ──
    // Takes base64 file data, finds the employee's Drive folder (parent of their
    // I-9 file), uploads the photo there, writes the new file's ID to column AM.
    // Also stamps the "returned" timestamp (column AL) since the photo only
    // arrives once the papers have come back.
    if (data.action === 'uploadWorkingPapers') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      // Find employee's folder via their I-9 file (column T = 20)
      const i9FileId = sheet.getRange(rowId, 20).getValue();
      if (!i9FileId) return json({ error: 'No I-9 file on record — cannot find employee folder' });
      let folder;
      try {
        const i9File = DriveApp.getFileById(i9FileId);
        const parents = i9File.getParents();
        if (!parents.hasNext()) return json({ error: 'I-9 file has no parent folder' });
        folder = parents.next();
      } catch (err) {
        return json({ error: 'Could not access I-9 folder: ' + err.message });
      }
      // Upload the photo
      const bytes = Utilities.base64Decode(data.fileData);
      const blob  = Utilities.newBlob(
        bytes,
        data.mimeType || 'image/jpeg',
        data.filename || ('working-papers-' + new Date().toISOString().slice(0,10) + '.jpg')
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      // Write file ID to column AM (39) and timestamp to AL (38)
      sheet.getRange(rowId, 39).setValue(file.getId());
      sheet.getRange(rowId, 38).setValue(new Date().toISOString());
      return json({ status: 'ok', fileId: file.getId(), url: file.getUrl() });
    }

    // ── /de112b: set lifecycle status ──────────────────────────────────
    // Columns AN(40)=status, AO(41)=activatedAt, AP(42)=inactiveAt, AQ(43)=inactiveReason
    if (data.action === 'setStatus') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      const status = data.status;
      if (!['onboarding','active','inactive'].includes(status)) {
        return json({ error: 'Invalid status' });
      }
      const now = new Date().toISOString();
      sheet.getRange(rowId, 40).setValue(status);
      if (status === 'active') {
        // Stamp activatedAt if not already set; clear inactive fields
        const existing = sheet.getRange(rowId, 41).getValue();
        if (!existing) sheet.getRange(rowId, 41).setValue(now);
        sheet.getRange(rowId, 42).setValue('');
        sheet.getRange(rowId, 43).setValue('');
      } else if (status === 'inactive') {
        sheet.getRange(rowId, 42).setValue(now);
        if (data.reason) sheet.getRange(rowId, 43).setValue(String(data.reason).substring(0, 500));
      } else if (status === 'onboarding') {
        // Reactivation case: clear activated/inactive stamps so they re-run the flow
        sheet.getRange(rowId, 41).setValue('');
        sheet.getRange(rowId, 42).setValue('');
        sheet.getRange(rowId, 43).setValue('');
      }
      return json({ status: 'ok' });
    }


    // ── /de112b: write Drive folder ID into AR (44) for a given row ───────
    // Used by the import script to point legacy/imported employees at a
    // shared "Pre-Middleman" folder, and by the backend after createFolder
    // to cache the new folder ID so we don't have to derive it from the I-9
    // file's parent on every Drive-folder click.
    if (data.action === 'setDriveFolderId') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      if (!data.folderId) return json({ error: 'folderId required' });
      sheet.getRange(rowId, 44).setValue(String(data.folderId));
      return json({ status: 'ok' });
    }


    // Run this once from the GAS editor (or webhook) to insert fake rows so
    // you can test the new Employment Details + Export flow in the admin.
    // Both employees have fully-populated I-9 Section 1 doc details and
    // submitted form data. Neither has employment details filled (that's
    // the point — you set them from the admin to test the flow).
    //
    // To run from the editor: select `doPost` is wrong — instead add a
    // temporary function:  function seed() { doPost({postData:{contents:
    //   JSON.stringify({secret:GAS_SECRET, action:'seedTestEmployees'})}}); }
    // then run seed() once. Remove when done.
    if (data.action === 'seedTestEmployees') {
      const now = new Date().toISOString();
      const rows = [
        [
          now, 'Marcus', 'Johnson', 'marcus.johnson+test@example.com',
          '(302) 555-0101', '***-**-1234', '2000-01-15',
          '123 Main Street', 'Wilmington', 'DE', '19801',
          'pending', '',  // i9Status, driveUrl (no real Drive folder)
          '', '', '', '', '', '',  // i9 section 2 fields
          '', '2025-01-20',  // i9FileId, startDate
          'Jane Johnson', 'Parent', '(302) 555-0102',  // emergency contact
          'new',  // overallStatus
          'Male',  // gender
          '', '', '', '', '',  // pay rate, position, location, dept, hire date (blank intentionally)
          '', '', '', '', '',  // export timestamps (blank)
        ],
        [
          now, 'Sofia', 'Martinez', 'sofia.martinez+test@example.com',
          '(302) 555-0201', '***-**-5678', '2003-07-22',
          '456 Elm Ave', 'Newark', 'DE', '19711',
          'pending', '',
          '', '', '', '', '', '',
          '', '2025-02-01',
          'Rosa Martinez', 'Parent', '(302) 555-0202',
          'new',
          'Female',
          '', '', '', '', '',
          '', '', '', '', '',
        ],
      ];
      rows.forEach(r => sheet.appendRow(r));
      return json({ status: 'ok', inserted: rows.length });
    }

    // ── /de112b: remove all test employees (rows where email ends with '+test@example.com') ──
    // Clean up after admin preview testing.
    if (data.action === 'removeTestEmployees') {
      const all = sheet.getDataRange().getValues();
      let removed = 0;
      // Iterate from bottom to top so row indexes don't shift as we delete
      for (let i = all.length - 1; i > 0; i--) {
        const email = String(all[i][3] || '');
        if (email.includes('+test@example.com')) {
          sheet.deleteRow(i + 1);
          removed++;
        }
      }
      return json({ status: 'ok', removed: removed });
    }

    // ── Delete a single test/demo row ────────────────────────────────────
    if (data.action === 'deleteRow') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 2) return json({ error: 'Invalid rowId' });
      // Safety check: only allow deletion of rows flagged as testEntry
      const lastCol = Math.max(sheet.getLastColumn(), 47);
      const row = sheet.getRange(rowId, 1, 1, lastCol).getValues()[0];
      if (row[46] !== 'TRUE') return json({ error: 'Row is not a test entry — cannot delete' });
      sheet.deleteRow(rowId);
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

// ─────────────────────────────────────────────────────────────────────────
// Editor-runnable helpers for /de112b admin testing.
// From the Apps Script editor, select the function name from the toolbar
// dropdown and click Run. Uses the GAS_SECRET to call doPost() directly —
// no webhook round-trip needed.
// ─────────────────────────────────────────────────────────────────────────

function runSeedTestEmployees() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        secret: GAS_SECRET,
        action: 'seedTestEmployees',
      }),
    },
  };
  const result = doPost(fakeEvent);
  Logger.log('seedTestEmployees result: ' + result.getContent());
}

function runRemoveTestEmployees() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        secret: GAS_SECRET,
        action: 'removeTestEmployees',
      }),
    },
  };
  const result = doPost(fakeEvent);
  Logger.log('removeTestEmployees result: ' + result.getContent());
}
