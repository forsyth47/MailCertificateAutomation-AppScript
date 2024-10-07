function generateAndSendCertificates() {
  // Variables to be changed
  // templateId is the google presentation unique ID in the url. You must also have access to view the presentation. https://docs.google.com/presentation/d/<<templateId>>
  const templateId = '1-sbg-ibfvuisfvisfvuhsvdhuksvuyfvskvf'; // Replace with your actual template ID
  const senderName = "XYZ Startup"

  // PrintOut remaining mail quota
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 2;  // Skip header
  const numRows = sheet.getLastRow() - 1;  // Total rows of data
  const dataRange = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());
  const data = dataRange.getValues();
  
  // Process each form response
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const recipientName = row[2];  // Assuming Name is in column C (index 2)
    const email = row[4];  // Assuming Email is in column E (index 4)
    const course = "Full Stack Development Workshop";  // Assuming Course is static
    const date = "10/09/2024"; // Assigned date

    // Check if the certificate has already been sent
    if (row[7] === "sent") { // Assuming "sent" is in column H (index 7)
      console.log(`${i+1}. (Done Already) Mail sent to ${email} and ${name}`);
    } else {
      // Create certificate and send email
      const cert = createCertificateAsImage(templateId, name, course, date);
      sendCertificateEmail(recipientName, email, cert, senderName);
      
      // Mark as sent in the sheet
      sheet.getRange(startRow + i, 8).setValue("sent"); // Update the corresponding row in the "sent" column
      
      console.log(`${i+1}. Mail sent to ${email} and ${name}`);
      var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
      Logger.log("Remaining email quota: " + emailQuotaRemaining);
    }
  }
}

function createCertificateAsImage(templateId, name, course, date) {
  // Create a copy of the Google Slides template
  const template = DriveApp.getFileById(templateId);
  const copy = template.makeCopy(`Certificate for ${name}`);
  const slides = SlidesApp.openById(copy.getId());
  const slide = slides.getSlides()[0];  // Assuming the certificate is on the first slide
  const content = `This is to certify that ${name} has participated in and successfully completed a project as part of the ${course} at Mohamed Sathak A.J. College Of Engineering conducted on ${date} by Nainu Startups.`;

  // Replace placeholders in the slide with actual values
  slide.replaceAllText('<<content>>', content);
  
  slides.saveAndClose();

  // Export the slide as a PDF
  const pdfBlob = copy.getAs(MimeType.PDF);
  const pdfFile = DriveApp.createFile(pdfBlob).setName(`Certificate_${name}.pdf`);

  // Cleanup the temporary copy
  copy.setTrashed(true); 

  return pdfFile;
}

function sendCertificateEmail(recipientName, email, cert, senderName) {
  const subject = `Certificate of Completion for ${name}`;
  const body = `Dear ${recipientName},\n\nCongratulations on completing the course!\n\nPlease find your certificate attached as a PDF.\nKindly post the certificate on your LinkedIn account with tagging @iamjoshua47 and @nainu-startups.\n\nBest regards,\nJoshua`;

  // Send email with the certificate PDF as an attachment
  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    attachments: [cert],
    name: senderName
  });
}
