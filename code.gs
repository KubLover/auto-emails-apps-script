function sendCustomEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Header indexes
  const headers = data[0];
  const userEmailIndex = headers.indexOf("User Email");
  const sendToIndex = headers.indexOf("Send email to");
  const passwordIndex = headers.indexOf("Password");
  const subjectIndex = headers.indexOf("Subject");
  const bodyIndex = headers.indexOf("Body");
  const footerIndex = headers.indexOf("Footer");
  
  // Add Status column if not already added
  let statusIndex = headers.indexOf("Status");
  if (statusIndex === -1) {
    sheet.getRange(1, headers.length + 1).setValue("Status");
    statusIndex = headers.length;
  }

  // Process each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const login = row[userEmailIndex];
    const recipient = row[sendToIndex];
    const password = row[passwordIndex];
    const subject = row[subjectIndex];
    const body = row[bodyIndex];
    const footer = row[footerIndex];

    // Compose final HTML email body the SLACK LINK needs to be changed.
    const htmlBody = `
      <p>Hello,</p>
      <p>${body}</p>
      <ul>
        <li><strong>Login:</strong> ${login}</li>
        <li><strong>Password:</strong> ${password}</li>
      </ul>
      <p>
      Once the system is live, you'll be able to log in.
      </p>
      <br>
      <p style="font-size: 1em; color: #666;">
        This is an automatically generated message. Please do not reply to this email.<br>
        For inquiries, contact our support team at our
        <a href="SLACK LINK" target="_blank">Slack channel</a>.
      </p>
    `;

    try {
      // Send the email with custom name and HTML body
      GmailApp.sendEmail(recipient, subject, '', {
        name: "System Administrator",
        htmlBody: htmlBody
      });

      // Update status column with success and timestamp
      sheet.getRange(i + 1, statusIndex + 1).setValue("Sent: " + new Date().toLocaleString());
    } catch (e) {
      // Update status column with failure and error
      sheet.getRange(i + 1, statusIndex + 1).setValue("Failed: " + e.message);
    }
  }
}
