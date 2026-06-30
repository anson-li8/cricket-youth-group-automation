const CONFIG = {
  COL_STUDENT_NAME  : 2,
  COL_STUDENT_EMAIL : 3,
  COL_PARENT_EMAIL  : 8,
  SENDER_NAME       : "Cricket Youth Volunteer Group Presidency",
  REPLY_TO_EMAIL    : "ansonyuanruili@gmail.com",
}

const TEMPLATE = {
  SUBJECT: "{NAME}, we need your help for the July 4th Dragon Dance! 🐉",
  BODY: `Hi {NAME},

Eric and I were just elected as President and VP for this year, and we are really looking forward to working with everyone! If you are new, welcome to the group!

We are reaching out to you individually because we urgently need a few more volunteers for the Dragon Dance at the July 4th Parade this Saturday. Since you are a valued member of the group, we would really appreciate your help to ensure we have enough people for this event to happen. (Plus, it is a fun way to get some easy volunteer hours!).

Sign Up Here: https://www.signupgenius.com/go/10C0845A4A62EA7FBC61-64584677-city

Let me know if you can make it, or if you have any questions!

Thank you,
Anson Li
President, Cricket Youth Volunteer Group`
};

function generatePersonalizedDrafts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let draftCount = 0;
  for (let i = 1; i < data.length; i++) { //skip 0 cuz of header row
    const row = data[i];
    const fullName = String(row[CONFIG.COL_STUDENT_NAME]).trim();
    const studentEmail = String(row[CONFIG.COL_STUDENT_EMAIL]).trim();
    const parentEmail = String(row[CONFIG.COL_PARENT_EMAIL]).trim();
    if (!studentEmail || !studentEmail.includes("@") || !fullName) continue;
    const firstName = fullName.split(" ")[0];
    const personalizedSubject = TEMPLATE.SUBJECT.replace(/{NAME}/g, firstName);
    const personalizedBody = TEMPLATE.BODY.replace(/{NAME}/g, firstName);
    const mailOptions = {
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.REPLY_TO_EMAIL
    };
    if (parentEmail && parentEmail.includes("@") && parentEmail !== studentEmail) {
      mailOptions.cc = parentEmail;
    }
    GmailApp.createDraft(studentEmail, personalizedSubject, personalizedBody, mailOptions);
    draftCount++;
  }
  SpreadsheetApp.getUi().alert(`Nice! Generated ${draftCount} drafts.`);
}