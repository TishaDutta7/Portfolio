const SPREADSHEET_ID = '1UGTBTasvxtcDQ3B7ZChI_Mdjf7bhvuhc7GBiKLArE14';
const SHEET_NAME = 'Form Responses';

function doPost(e) {
  try {
    // Accept normal form-encoded requests from the GitHub Pages frontend,
    // and also accept JSON if a future client sends it.
    let p = (e && e.parameter) ? Object.assign({}, e.parameter) : {};
    if ((!p.name || !p.email) && e && e.postData && e.postData.contents) {
      try {
        const parsed = JSON.parse(e.postData.contents);
        p = Object.assign(p, parsed);
      } catch (_) {}
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Timestamp','Name','Email','Brand / Company','Instagram URL',
        'LinkedIn URL','Website','Service Required','Budget Range',
        'Project Description','Preferred Contact','Additional Information'
      ]);
      sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const timestamp = new Date();
    sheet.appendRow([
      timestamp,
      p.name || '',
      p.email || '',
      p.brand || '',
      p.instagram || '',
      p.linkedin || '',
      p.website || '',
      p.service || '',
      p.budget || '',
      p.description || '',
      p.contactMethod || '',
      p.additional || ''
    ]);

    // Send every enquiry notification to Tisha's email address.
    const notificationEmail = 'tishadutta7320@gmail.com';
    if (notificationEmail) {
      const subject = 'New Portfolio Enquiry — ' + (p.name || 'New Lead');
      const body = [
        'NEW PORTFOLIO ENQUIRY',
        '',
        'Name: ' + (p.name || '-'),
        'Email: ' + (p.email || '-'),
        'Brand / Company: ' + (p.brand || '-'),
        'Instagram: ' + (p.instagram || '-'),
        'LinkedIn: ' + (p.linkedin || '-'),
        'Website: ' + (p.website || '-'),
        'Service: ' + (p.service || '-'),
        'Budget: ' + (p.budget || '-'),
        'Preferred contact: ' + (p.contactMethod || '-'),
        '',
        'Project description:',
        p.description || '-',
        '',
        'Additional information:',
        p.additional || '-',
        '',
        'Open Google Sheet:',
        'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit'
      ].join('\n');

      MailApp.sendEmail({
        to: notificationEmail,
        subject: subject,
        body: body,
        replyTo: p.email || undefined,
        name: 'Tisha Dutta Portfolio'
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(err);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput('Tisha Dutta Portfolio — Form endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
