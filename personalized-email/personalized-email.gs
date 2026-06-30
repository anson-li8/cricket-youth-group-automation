const CONFIG = {
  COL_STUDENT_NAME  : 2,
  COL_STUDENT_EMAIL : 3,
  COL_PARENT_EMAIL  : 8,
  SENDER_NAME       : "Cricket Youth Volunteer Group Presidency",
  REPLY_TO_EMAIL    : "ansonyuanruili@gmail.com",
  SIGNUP_URL    : "https://www.signupgenius.com/go/10C0845A4A62EA7FBC61-64584677-city"
}
const STUDENT_VERSION = {
  SUBJECT: "{FIRST_NAME}, we need your help for the July 4th Dragon Dance! 🐉",
  GREETING: "Hi {FIRST_NAME},",
  PARAGRAPHS: [
    "Eric and I were just officially elected as President and VP for this upcoming year, and we really look forward to working with everyone!",
    "We are reaching out to you individually because we urgently need volunteers for the Dragon Dance at the July 4th Parade this Saturday. Since you are a valued member of our group, we would really appreciate your help to make sure our team has enough people to run this event.",
    "It will be a great way to start off our new term together, catch up with friends, and get some easy volunteer hours. Please use the button below to reserve your spot on the signup sheet!"
  ],
  CTA_BUTTON_TEXT: "Join the Dragon Dance Team"
};

const PARENT_VERSION = {
  SUBJECT: "July 4th Parade Dragon Dance — Volunteer Invitation for {FIRST_NAME} 🐉",
  GREETING: "Dear Parents & Family of {FIRST_NAME},",
  PARAGRAPHS: [
    "My name is Anson Li, and along with Eric Zhang, we were just officially elected as the new President and VP of the Cricket Youth Volunteer Group for the upcoming year.",
    "We are reaching out to your family because our group urgently needs additional youth volunteers for our Dragon Dance team at the upcoming Fourth of July Parade this Saturday morning.",
    "As {FIRST_NAME} is a highly valued member of our group, we would be incredibly grateful if they could join us to participate. It is a wonderful community celebration and an excellent event for them to earn some volunteer hours. If they are available to help, please use the button below to access the signup window."
  ],
  CTA_BUTTON_TEXT: "Register Your Child to Volunteer"
};
function executeDrafter() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRangeValues = sheet.getDataRange().getValues();
  let draftsGenerated = 0;
  for (let rowIdx = 1; rowIdx < dataRangeValues.length; rowIdx++) {
    const activeRow = dataRangeValues[rowIdx];
    const targetFullName = String(activeRow[SYSTEM_CONFIG.INDEX_STUDENT_NAME]).trim();
    const studentTargetEmail = String(activeRow[SYSTEM_CONFIG.INDEX_STUDENT_EMAIL]).trim();
    const parentTargetEmail = String(activeRow[SYSTEM_CONFIG.INDEX_STUDENT_PARENT_EMAIL] || activeRow[SYSTEM_CONFIG.INDEX_PARENT_EMAIL]).trim();
    if (!studentTargetEmail || !studentTargetEmail.includes("@") || !targetFullName) continue;
    const baseFirstName = targetFullName.split(" ")[0];
    let emailSubjectLine = "";
    let compiledHtmlMessage = "";
    const operationalMailOptions = {
      name: SYSTEM_CONFIG.SENDER_DISPLAY_NAME,
      replyTo: SYSTEM_CONFIG.ACCOUNT_REPLY_TO
    };
    const isParentOwnedInbox = (studentTargetEmail.toLowerCase() === parentTargetEmail.toLowerCase());
    if (isParentOwnedInbox) {
      emailSubjectLine = PARENT_VERSION.SUBJECT.replace(/{FIRST_NAME}/g, baseFirstName);
      compiledHtmlMessage = renderHtmlFramework(
        PARENT_VERSION.GREETING.replace(/{FIRST_NAME}/g, baseFirstName),
        PARENT_VERSION.PARAGRAPHS.map(str => str.replace(/{FIRST_NAME}/g, baseFirstName)),
        SYSTEM_CONFIG.CTA_SIGNUP_URL,
        PARENT_VERSION.CTA_BUTTON_TEXT
      );
    } else {
      emailSubjectLine = STUDENT_VERSION.SUBJECT.replace(/{FIRST_NAME}/g, baseFirstName);
      compiledHtmlMessage = renderHtmlFramework(
        STUDENT_VERSION.GREETING.replace(/{FIRST_NAME}/g, baseFirstName),
        STUDENT_VERSION.PARAGRAPHS,
        SYSTEM_CONFIG.CTA_SIGNUP_URL,
        STUDENT_VERSION.CTA_BUTTON_TEXT
      );
      if (parentTargetEmail && parentTargetEmail.includes("@")) {
        operationalMailOptions.cc = parentTargetEmail;
      }
    }
    operationalMailOptions.htmlBody = compiledHtmlMessage;
    const plainTextFallback = "Please configure an HTML-compatible email client to review this dispatch.";
    GmailApp.createDraft(studentTargetEmail, emailSubjectLine, plainTextFallback, operationalMailOptions);
    draftsGenerated++;
  }
  SpreadsheetApp.getUi().alert(`Done! ${draftsGenerated} drafts generated and ready to send.`);
}
// function generatePersonalizedDrafts() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   const data = sheet.getDataRange().getValues();
//   let draftCount = 0;
//   for (let i = 1; i < data.length; i++) { //skip 0 cuz of header row
//     const row = data[i];
//     const fullName = String(row[CONFIG.COL_STUDENT_NAME]).trim();
//     const studentEmail = String(row[CONFIG.COL_STUDENT_EMAIL]).trim();
//     const parentEmail = String(row[CONFIG.COL_PARENT_EMAIL]).trim();
//     if (!studentEmail || !studentEmail.includes("@") || !fullName) continue;
//     const firstName = fullName.split(" ")[0];
//     const personalizedSubject = TEMPLATE.SUBJECT.replace(/{NAME}/g, firstName);
//     const personalizedBody = TEMPLATE.BODY.replace(/{NAME}/g, firstName);
//     const mailOptions = {
//       name: CONFIG.SENDER_NAME,
//       replyTo: CONFIG.REPLY_TO_EMAIL
//     };
//     if (parentEmail && parentEmail.includes("@") && parentEmail !== studentEmail) {
//       mailOptions.cc = parentEmail;
//     }
//     GmailApp.createDraft(studentEmail, personalizedSubject, personalizedBody, mailOptions);
//     draftCount++;
//   }
//   SpreadsheetApp.getUi().alert(`Nice! Generated ${draftCount} drafts.`);
// }
function renderHtmlFramework(greetingLine, textParagraphsArray, actionButtonUrl, actionButtonText) {
  const HTML_paragraphsMarkup = textParagraphsArray.map(txt => 
    `<p style="margin: 0 0 16px 0; font-size: 16px; line-height: 1.6; color: #4A5568;">${txt}</p>`
  ).join("");

  return `
    <div style="margin: 0; padding: 40px 10px; background-color: #F7FAFC; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; margin: 0 auto; background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
        
        <tr>
          <td style="background: linear-gradient(135deg, #1A365D 0%, #2A4365 100%); padding: 35px 20px; text-align: center;">
            <span style="display: block; font-size: 22px; font-weight: 700; color: #FFFFFF; letter-spacing: 0.5px; text-transform: uppercase;">Cricket Youth Volunteer Group</span>
          </td>
        </tr>
        
        <tr>
          <td style="padding: 35px 30px;">
            <h2 style="margin: 0 0 20px 0; font-size: 18px; font-weight: 700; color: #1A365D;">${greetingLine}</h2>
            
            ${HTML_paragraphsMarkup}
            
            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin: 35px 0;">
              <tr>
                <td align="center">
                  <a href="${actionButtonUrl}" target="_blank" style="display: inline-block; background-color: #DD6B20; color: #FFFFFF; font-size: 16px; font-weight: 700; text-decoration: none; padding: 14px 30px; border-radius: 6px; box-shadow: 0 4px 6px rgba(221, 107, 32, 0.25);">
                    ${actionButtonText}
                  </a>
                </td>
              </tr>
            </table>
            
            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-top: 35px; padding-top: 25px; border-top: 1px solid #E2E8F0;">
              <tr>
                <td>
                  <p style="margin: 0 0 2px 0; font-size: 16px; font-weight: 700; color: #1A365D;">Anson Li</p>
                  <p style="margin: 0; font-size: 14px; color: #718096; font-weight: 500;">President, Cricket Youth Volunteer Group</p>
                </td>
              </tr>
            </table>
            
          </td>
        </tr>
      </table>
    </div>
  `;
}