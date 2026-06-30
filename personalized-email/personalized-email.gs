const CONFIG = {
  COL_STUDENT_NAME  : 2,
  COL_STUDENT_EMAIL : 3,
  COL_PARENT_EMAIL  : 8,
  SENDER_NAME       : "Cricket Youth Volunteer Group Presidency",
  REPLY_TO_EMAIL    : "ansonyuanruili@gmail.com",
  SIGNUP_URL        : "https://www.signupgenius.com/go/10C0845A4A62EA7FBC61-64584677-city"
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Cricket Management')
    .addItem('Send Personalized Drafts', 'executeDrafter')
    .addToUi();
}

const STUDENT_VERSION = {
  SUBJECT: "{FIRST_NAME}, we need your help for the July 4th Dragon Dance!",
  GREETING: "Hi {FIRST_NAME},",
  PARAGRAPHS: [
    "Eric and I will be the President and VP for this upcoming year, and we really look forward to working with everyone!",
    "We are reaching out to you individually because we urgently need volunteers for the Dragon Dance at the July 4th Parade this Saturday. Since you are a valued member of our group, we would really appreciate your help to make sure our team has enough people to run this event.",
    "It'll be a great start to our new term together, you'll get some easy volunteer hours (4.5 hours), and you might even make some new friends. Please use the button below to sign up and let me know if you have any questions!"
  ],
  CTA_BUTTON_TEXT: "Join the Dragon Dance Team"
};

const PARENT_VERSION = {
  SUBJECT: "July 4th Parade Dragon Dance — Volunteer Invitation for {FIRST_NAME}",
  GREETING: "Dear Parents & Family of {FIRST_NAME},",
  PARAGRAPHS: [
    "My name is Anson Li, and along with Eric Zhang, we will be the new President and VP of the Cricket Youth Volunteer Group for the upcoming year.",
    "We are reaching out to your family because our group urgently needs additional youth volunteers for our Dragon Dance team at the upcoming Fourth of July Parade this Saturday morning.",
    "Since we highly value {FIRST_NAME} in our group, we would be incredibly grateful if they could join us. It's a wonderful community celebration and an excellent event to earn some volunteer hours. If they are available to help, please use the button below to access the signup sheet."
  ],
  CTA_BUTTON_TEXT: "Register Your Child to Volunteer"
};

function executeDrafter() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  let draftCount = 0;

  const parentOnly = {};
  const studentDirect = [];

  // sort roster into direct students vs parent-only
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const fullName = String(row[CONFIG.COL_STUDENT_NAME]).trim();
    const studentEmail = String(row[CONFIG.COL_STUDENT_EMAIL]).trim();
    const parentEmail = String(row[CONFIG.COL_PARENT_EMAIL]).trim();

    if (!fullName) continue;

    const firstName = fullName.split(" ")[0];
    const sEmail = studentEmail.toLowerCase();
    const pEmail = parentEmail.toLowerCase();

    const validStudent = sEmail.includes("@");
    const validParent = pEmail.includes("@");

    // parent only: student email is blank or exact same as parent
    if (!validStudent || sEmail === pEmail) {
      if (validParent) {
        if (!parentOnly[pEmail]) {
          parentOnly[pEmail] = { to: parentEmail, names: [] };
        }
        parentOnly[pEmail].names.push(firstName);
      }
    } else {
      // student has distinct email
      studentDirect.push({
        to: studentEmail,
        cc: validParent ? parentEmail : null,
        name: firstName
      });
    }
  }

  const baseOptions = {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO_EMAIL
  };
  const fallback = "Please use an HTML-compatible client to view this message.";

  // draft parent-only emails grouped by siblings
  for (const key in parentOnly) {
    const p = parentOnly[key];
    let joinedNames = p.names[0];
    
    if (p.names.length === 2) {
      joinedNames = `${p.names[0]} and ${p.names[1]}`;
    } else if (p.names.length > 2) {
      joinedNames = `${p.names.slice(0, -1).join(", ")}, and ${p.names[p.names.length - 1]}`;
    }

    const subject = PARENT_VERSION.SUBJECT.replace(/{FIRST_NAME}/g, joinedNames);
    const htmlBody = renderHtml(
      PARENT_VERSION.GREETING.replace(/{FIRST_NAME}/g, joinedNames),
      PARENT_VERSION.PARAGRAPHS.map(str => str.replace(/{FIRST_NAME}/g, joinedNames)),
      CONFIG.SIGNUP_URL,
      PARENT_VERSION.CTA_BUTTON_TEXT
    );

    const options = { ...baseOptions, htmlBody: htmlBody };
    GmailApp.createDraft(p.to, subject, fallback, options);
    draftCount++;
  }

  // draft student emails with parent cc
  for (const s of studentDirect) {
    const subject = STUDENT_VERSION.SUBJECT.replace(/{FIRST_NAME}/g, s.name);
    const htmlBody = renderHtml(
      STUDENT_VERSION.GREETING.replace(/{FIRST_NAME}/g, s.name),
      STUDENT_VERSION.PARAGRAPHS,
      CONFIG.SIGNUP_URL,
      STUDENT_VERSION.CTA_BUTTON_TEXT
    );

    const options = { ...baseOptions, htmlBody: htmlBody };
    if (s.cc) {
      options.cc = s.cc;
    }

    GmailApp.createDraft(s.to, subject, fallback, options);
    draftCount++;
  }

  SpreadsheetApp.getUi().alert(`Done! ${draftCount} drafts generated and ready to send.`);
}

function renderHtml(greeting, paragraphs, buttonUrl, buttonText) {
  const paragraphsHtml = paragraphs.map(p => 
    `<p style="margin: 0 0 16px 0; font-size: 16px; line-height: 1.6; color: #4A5568;">${p}</p>`
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
            <h2 style="margin: 0 0 20px 0; font-size: 18px; font-weight: 700; color: #1A365D;">${greeting}</h2>
            ${paragraphsHtml}
            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin: 35px 0;">
              <tr>
                <td align="center">
                  <a href="${buttonUrl}" target="_blank" style="display: inline-block; background-color: #DD6B20; color: #FFFFFF; font-size: 16px; font-weight: 700; text-decoration: none; padding: 14px 30px; border-radius: 6px; box-shadow: 0 4px 6px rgba(221, 107, 32, 0.25);">
                    ${buttonText}
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