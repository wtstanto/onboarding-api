// Auntie Anne's Onboarding — Google Apps Script Webhook
// -------------------------------------------------------
// Deploy this as a Web App:
//   Deploy → New deployment → Type: Web app
//   Execute as: Me
//   Who has access: Anyone
// Copy the Web App URL and set it as GAS_WEBHOOK_URL in Railway.
//
// Set GAS_SECRET below to match the GAS_SECRET env var you set in Railway.

const GAS_SECRET = 'REPLACE_WITH_YOUR_SECRET';  // e.g. 'pretzel2024'
const SHEET_NAME = 'Sheet1';

// Column layout (1-indexed):
// A(1)=submittedAt  B(2)=firstName   C(3)=lastName   D(4)=email
// E(5)=phone        F(6)=ssn         G(7)=dob        H(8)=address1
// I(9)=city         J(10)=state      K(11)=zip        L(12)=i9Status
// M(13)=driveUrl    N(14)=i9Doc      O(15)=i9DocNum  P(16)=i9Issuer
// Q(17)=i9ExpDate   R(18)=i9VerDate  S(19)=i9VerBy

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Auth check
    if (data.secret !== GAS_SECRET) {
      return json({ error: 'Unauthorized' });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    if (data.action === 'log') {
      sheet.appendRow([
        data.submittedAt || new Date().toISOString(),
        data.firstName   || '',
        data.lastName    || '',
        data.email       || '',
        data.phone       || '',
        data.ssn         || '',
        data.dob         || '',
        data.address1    || '',
        data.city        || '',
        data.state       || '',
        data.zip         || '',
        'pending',            // i9Status
        data.driveUrl    || '',
        '', '', '', '', '', ''  // i9 fields (filled later)
      ]);
      return json({ status: 'ok', rowId: sheet.getLastRow() });
    }

    if (data.action === 'getAll') {
      const rows = sheet.getDataRange().getValues();
      const employees = rows.map((row, i) => {
        const i9Complete = (row[11] || 'pending') === 'complete';
        return {
          id:          i + 1,
          submittedAt: row[0]  ? new Date(row[0]).toISOString() : '',
          firstName:   row[1]  || '',
          lastName:    row[2]  || '',
          email:       row[3]  || '',
          phone:       row[4]  || '',
          ssn:         row[5]  || '',
          dob:         row[6]  ? formatDate(row[6]) : '',
          address1:    row[7]  || '',
          city:        row[8]  || '',
          state:       row[9]  || '',
          zip:         row[10] || '',
          i9Status:    row[11] || 'pending',
          driveUrl:    row[12] || '',
          i9s2: i9Complete ? {
            docTitle:     row[13] || '',
            docNumber:    row[14] || '',
            issuer:       row[15] || '',
            expDate:      row[16] || '',
            verifiedDate: row[17] ? new Date(row[17]).toISOString() : '',
            empName:      row[18] || '',
          } : null
        };
      });
      return json({ employees });
    }

    if (data.action === 'completeI9') {
      const rowId = parseInt(data.rowId);
      if (!rowId || rowId < 1) return json({ error: 'Invalid rowId' });
      // Update columns L–S (12–19)
      sheet.getRange(rowId, 12, 1, 8).setValues([[
        'complete',
        sheet.getRange(rowId, 13).getValue(), // preserve existing driveUrl
        data.docTitle  || '',
        data.docNumber || '',
        data.issuer    || '',
        data.expDate   || '',
        new Date().toISOString(),
        data.empName   || '',
      ]]);
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
  } catch(e) {
    return String(val);
  }
}
