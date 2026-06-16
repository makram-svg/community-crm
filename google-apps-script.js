/**
 * Google Apps Script for "Droob Community" CRM & Event Automation
 * Copy-paste this script into your Google Sheet (Extensions -> Apps Script)
 * Publish it as a Web App (Deploy -> New Deployment -> Web App)
 * Select Execute as: "Me" (your email) and Who has access: "Anyone"
 */

const SENDER_EMAIL = "M.akram@doroobangels.com";
const SENDER_NAME = "Droob Community | مجتمع دروب";

function getEmailHeader() {
  return `<div dir="rtl" style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;background:#07070f;border-radius:12px;overflow:hidden;border:1px solid #1e1e3a;"><div style="background:#0e0e1c;padding:20px 28px;border-bottom:1px solid #1e1e3a;"><div style="float:right;width:36px;height:36px;background:#3b82f6;border-radius:8px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:16px;color:#fff;font-family:monospace;text-align:center;line-height:36px;">D</div><div style="float:right;margin-right:10px;"><div style="font-size:15px;font-weight:600;color:#f0f0f8;">Droob Community</div><div style="font-size:11px;color:#505078;font-family:monospace;">مجتمع دروب الاستثماري</div></div><div style="clear:both;"></div></div><div style="padding:28px;">`;
}

function getEmailFooter() {
  return `</div><div style="background:#0e0e1c;padding:18px 28px;border-top:1px solid #1e1e3a;text-align:center;"><div style="font-size:12px;color:#505078;margin-bottom:10px;">Droob Community | مجتمع دروب الاستثماري</div><div style="margin-bottom:10px;"><a href="https://www.linkedin.com/company/droobangels" style="color:#3b82f6;text-decoration:none;font-size:12px;margin:0 8px;">LinkedIn</a><a href="mailto:M.akram@doroobangels.com" style="color:#3b82f6;text-decoration:none;font-size:12px;margin:0 8px;">تواصل معنا</a></div><div style="font-size:11px;color:#2a2a50;">© 2026 Droob Community. جميع الحقوق محفوظة.</div></div></div>`;
}

function getInfoCard(content) {
  return `<div style="background:#0e0e1c;border:1px solid #1e1e3a;border-radius:8px;padding:18px;margin:16px 0;text-align:right;">${content}</div>`;
}

function getDivider() {
  return `<div style="height:1px;background:#1e1e3a;margin:20px 0;"></div>`;
}

function getBtnPrimary(url, label) {
  return `<a href="${url}" style="display:inline-block;background:#3b82f6;color:#fff;padding:12px 28px;text-decoration:none;border-radius:8px;font-weight:600;font-size:15px;margin:6px;">${label}</a>`;
}

function getBtnGhost(url, label) {
  return `<a href="${url}" style="display:inline-block;background:none;color:#9090b8;padding:12px 28px;text-decoration:none;border-radius:8px;font-weight:500;font-size:14px;border:1px solid #1e1e3a;margin:6px;">${label}</a>`;
}

// --- CONFIGURATION: MAP YOUR GOOGLE FORM RESPONSES SHEET COLUMNS ---
// Replace these numbers with the column indices in your Google Form responses sheet
// Note: 0 is Column A (usually Timestamp), 1 is Column B, 2 is Column C, etc.
const COLUMN_MAPPING = {
  timestamp: 0,
  name: 1,       // Column B: الاسم الكامل
  email: 2,      // Column C: البريد الإلكتروني
  phone: 3,      // Column D: رقم الجوال
  role: 4,       // Column E: الدور (مستثمر أم رائد أعمال)
  company: 5,    // Column F: اسم الشركة أو المشروع
  industry: 6,   // Column G: المجال أو القطاع
  ticketSize: 7, // Column H: حجم الاستثمار (للمستثمرين)
  stage: 8,      // Column I: مرحلة المشروع / الاستثمار
  linkedin: 9,   // Column J: رابط لينكد إن
  bio: 10        // Column K: نبذة مختصرة
};

// Automatically initialize sheets if they don't exist
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Members Sheet
  let membersSheet = ss.getSheetByName("Members");
  if (!membersSheet) {
    membersSheet = ss.insertSheet("Members");
    membersSheet.appendRow([
      "ID", "Name", "Email", "Phone", "Role", "Company",
      "Industry", "Ticket Size", "Stage", "LinkedIn", "Bio",
      "Engagement Score", "Member Type", "Participation Type", "Date Added"
    ]);
    // Freeze header row
    membersSheet.setFrozenRows(1);
  }
  
  // 2. Events Sheet
  let eventsSheet = ss.getSheetByName("Events");
  if (!eventsSheet) {
    eventsSheet = ss.insertSheet("Events");
    eventsSheet.appendRow([
      "ID", "Title", "Date", "Location", "Description", "Date Created"
    ]);
    eventsSheet.setFrozenRows(1);
  }
  
  // 3. Invitations Sheet
  let invSheet = ss.getSheetByName("Invitations");
  if (!invSheet) {
    invSheet = ss.insertSheet("Invitations");
    invSheet.appendRow([
      "ID", "Event ID", "Member ID", "RSVP Status", "Attended", "Token", "Sent At", "Updated At"
    ]);
    invSheet.setFrozenRows(1);
  }

  // 4. Newsletter Subscribers Sheet
  let newsletterSheet = ss.getSheetByName("Newsletter");
  if (!newsletterSheet) {
    newsletterSheet = ss.insertSheet("Newsletter");
    newsletterSheet.appendRow(["ID", "Name", "Email", "Subscribed At"]);
    newsletterSheet.setFrozenRows(1);
  }

  // 5. Contact Messages Sheet
  let contactSheet = ss.getSheetByName("ContactMessages");
  if (!contactSheet) {
    contactSheet = ss.insertSheet("ContactMessages");
    contactSheet.appendRow(["ID", "Name", "Email", "Phone", "Message", "Created At", "Is Read"]);
    contactSheet.setFrozenRows(1);
  }

  // 6. Interest Registrations Sheet
  let interestSheet = ss.getSheetByName("InterestRegistrations");
  if (!interestSheet) {
    interestSheet = ss.insertSheet("InterestRegistrations");
    interestSheet.appendRow(["ID", "Name", "Email", "Phone", "Company", "LinkedIn", "Interest Type", "Message", "Created At"]);
    interestSheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName("EventGuests")) {
    const guestsSheet = ss.insertSheet("EventGuests");
    guestsSheet.appendRow(["eventId","guestName","guestEmail","guestPhone","addedBy","addedAt","isExternal","rsvpStatus","checkedIn"]);
  }

  Logger.log("Setup completed successfully!");
}

// --- GOOGLE FORM SUBMIT TRIGGER ---
// Run this function automatically when a form is submitted
function onFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const membersSheet = ss.getSheetByName("Members");
    if (!membersSheet) {
      setup(); // Auto-setup if sheets don't exist
    }

    const values = e.values;
    if (!values || values.length === 0) {
      Logger.log("No values found in submit event");
      return;
    }

    // Extract values based on COLUMN_MAPPING
    const name = values[COLUMN_MAPPING.name] || "";
    const email = (values[COLUMN_MAPPING.email] || "").trim();
    const phone = values[COLUMN_MAPPING.phone] || "";
    const roleInput = values[COLUMN_MAPPING.role] || "Entrepreneur";
    const company = values[COLUMN_MAPPING.company] || "";
    const industry = values[COLUMN_MAPPING.industry] || "";
    const ticketSize = values[COLUMN_MAPPING.ticketSize] || "";
    const stage = values[COLUMN_MAPPING.stage] || "";
    const linkedin = values[COLUMN_MAPPING.linkedin] || "";
    const bio = values[COLUMN_MAPPING.bio] || "";

    if (!name || !email) {
      Logger.log("Missing Name or Email: " + JSON.stringify(values));
      return; // Mandatory columns missing
    }

    // Normalize role
    let role = "Entrepreneur";
    const roleLower = roleInput.toLowerCase();
    if (roleLower.includes("invest") || roleLower.includes("مستثمر")) {
      role = "Investor";
    }

    const members = getSheetData(membersSheet);
    const existing = members.find(m => m.email.toLowerCase() === email.toLowerCase());

    const now = new Date().toISOString();

    if (existing) {
      // Update existing record
      const dataRange = membersSheet.getDataRange().getValues();
      for (let i = 1; i < dataRange.length; i++) {
        if (dataRange[i][2].toLowerCase() === email.toLowerCase()) {
          const row = i + 1;
          membersSheet.getRange(row, 2, 1, 10).setValues([[
            name,
            phone,
            role,
            company,
            industry,
            ticketSize,
            stage,
            linkedin,
            bio,
            existing.engagementScore || 50
          ]]);
          Logger.log("Updated existing member: " + email);
          break;
        }
      }
    } else {
      // Insert new member row
      const id = "m_" + Math.random().toString(36).substr(2, 9);
      membersSheet.appendRow([
        id, name, email, phone, role, company, industry, ticketSize, stage, linkedin, bio, 50, now
      ]);
      Logger.log("Created new member: " + email);

      // Send Automated Welcome Email via Gmail
      sendWelcomeEmail(name, email);
    }
  } catch (err) {
    Logger.log("Error in onFormSubmit trigger: " + err.toString());
  }
}

// Handle GET requests (API endpoints, RSVP Landing Page, and CRM Dashboard)
function doGet(e) {
  try {
    const action = e.parameter.action;
    const token = e.parameter.token;
    
    // 1. If it's an RSVP action via invitation token
    if (token) {
      return renderRsvpPage(token, e.parameter.status);
    }
    
    // 2. REST API Actions
    if (action === 'getData') {
      return handleGetData();
    }
    
    // 3. DEFAULT: Serve the entire CRM Dashboard Webpage!
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle("مجتمع دروب | Droob Community CRM")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } catch (err) {
    return HtmlService.createHtmlOutput("<h2>حدث خطأ أثناء تحميل لوحة التحكم:</h2><p>" + err.toString() + "</p>");
  }
}

// Handle POST requests (API writes and syncs)
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    
    if (action === 'addMember') {
      return handleAddMember(postData.member);
    }
    if (action === 'addEvent') {
      return handleAddEvent(postData.event);
    }
    if (action === 'sendInvitations') {
      return handleSendInvitations(postData.eventId, postData.memberIds);
    }
    if (action === 'updateInvitation') {
      return handleUpdateInvitation(postData.invitationId, postData.updates);
    }
    if (action === 'syncForm') {
      return handleFormSubmitWebhook(postData);
    }
    if (action === 'addNewsletterSubscriber') {
      return handleAddNewsletterSubscriber(postData.subscriber);
    }
    if (action === 'addContactMessage') {
      return handleAddContactMessage(postData.message);
    }
    if (action === 'addInterestRegistration') {
      return handleAddInterestRegistration(postData.registration);
    }
    if (action === 'event_invite') {
      return handleEventInvite(postData);
    }
    if (action === 'event_invite_external') {
      return handleEventInviteExternal(postData);
    }
    if (action === 'addExternalGuest') {
      return handleAddExternalGuest(postData.eventId, postData.guest);
    }
    if (action === 'newsletter_send') {
      return handleNewsletterSend(postData);
    }

    return jsonResponse({ error: "Invalid action" });
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// --- API HANDLERS ---

function handleGetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const members = getSheetData(ss.getSheetByName("Members"));
  const events = getSheetData(ss.getSheetByName("Events"));
  const invitations = getSheetData(ss.getSheetByName("Invitations"));
  const newsletterSubscribers = getSheetData(ss.getSheetByName("Newsletter"));
  const contactMessages = getSheetData(ss.getSheetByName("ContactMessages"));
  const interestRegistrations = getSheetData(ss.getSheetByName("InterestRegistrations"));

  return jsonResponse({
    members,
    events,
    invitations,
    newsletterSubscribers,
    contactMessages,
    interestRegistrations
  });
}

function handleAddMember(member) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Members");
  const id = "m_" + Math.random().toString(36).substr(2, 9);
  const now = new Date().toISOString();
  
  sheet.appendRow([
    id,
    member.name || "",
    member.email || "",
    member.phone || "",
    member.role || "Entrepreneur",
    member.company || "",
    member.industry || "",
    member.ticketSize || "",
    member.stage || "",
    member.linkedin || "",
    member.bio || "",
    member.engagementScore || 50,
    member.memberType || "مستمع",
    member.participationType || "مستمع",
    now
  ]);
  
  member.id = id;
  member.createdAt = now;
  return jsonResponse({ success: true, member });
}

function handleAddEvent(event) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Events");
  const id = "e_" + Math.random().toString(36).substr(2, 9);
  const now = new Date().toISOString();
  
  sheet.appendRow([
    id,
    event.title || "",
    event.date || "",
    event.location || "",
    event.description || "",
    now
  ]);
  
  event.id = id;
  return jsonResponse({ success: true, event });
}

function handleSendInvitations(eventId, memberIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName("Members");
  const eventsSheet = ss.getSheetByName("Events");
  const invSheet = ss.getSheetByName("Invitations");
  
  const members = getSheetData(membersSheet);
  const events = getSheetData(eventsSheet);
  
  const event = events.find(e => e.id === eventId);
  if (!event) return jsonResponse({ error: "Event not found" });
  
  const webAppUrl = ScriptApp.getService().getUrl();
  const now = new Date().toISOString();
  
  const results = [];
  
  memberIds.forEach(memberId => {
    const member = members.find(m => m.id === memberId);
    if (member) {
      // Check existing invitation
      let invRowIdx = findRowIndex(invSheet, 1, eventId, 2, memberId);
      let token = "";
      let invId = "";
      
      if (invRowIdx === -1) {
        invId = "inv_" + Math.random().toString(36).substr(2, 9);
        token = Math.random().toString(36).substr(2, 15) + Math.random().toString(36).substr(2, 15);
        invSheet.appendRow([
          invId, eventId, memberId, "Pending", "Pending", token, now, now
        ]);
      } else {
        // Retrieve existing token and ID
        const rowValues = invSheet.getRange(invRowIdx, 1, 1, 8).getValues()[0];
        invId = rowValues[0];
        token = rowValues[5];
      }
      
      // Send Email via Gmail
      const rsvpUrl = `${webAppUrl}?token=${token}`;
      const emailSent = sendHtmlEmail(member.name, member.email, event.title, event.date, event.location, event.description, rsvpUrl);
      
      results.push({ email: member.email, sent: emailSent });
    }
  });
  
  return jsonResponse({ success: true, results });
}

function handleUpdateInvitation(invitationId, updates) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invitations");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === invitationId) {
      const rowNum = i + 1;
      
      if (updates.rsvpStatus !== undefined) {
        const colIdx = headers.indexOf("RSVP Status") + 1;
        sheet.getRange(rowNum, colIdx).setValue(updates.rsvpStatus);
      }
      if (updates.attended !== undefined) {
        const colIdx = headers.indexOf("Attended") + 1;
        sheet.getRange(rowNum, colIdx).setValue(updates.attended);
      }
      
      const updatedColIdx = headers.indexOf("Updated At") + 1;
      sheet.getRange(rowNum, updatedColIdx).setValue(new Date().toISOString());
      
      // Recalculate engagement score
      const memberId = data[i][2];
      recalculateMemberEngagementScore(memberId);
      
      return jsonResponse({ success: true });
    }
  }
  
  return jsonResponse({ error: "Invitation not found" });
}

// REST Webhook backup option
function handleFormSubmitWebhook(postData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Members");
  const members = getSheetData(sheet);
  
  const email = postData.email;
  if (!email) return jsonResponse({ error: "Email is required" });
  
  const existing = members.find(m => m.email.toLowerCase() === email.toLowerCase());
  
  const memberData = {
    name: postData.name || "",
    email: email,
    phone: postData.phone || "",
    role: postData.role || "Entrepreneur",
    company: postData.company || "",
    industry: postData.industry || "",
    ticketSize: postData.ticketSize || "",
    stage: postData.stage || "",
    linkedin: postData.linkedin || "",
    bio: postData.bio || "",
    engagementScore: 50
  };
  
  if (existing) {
    // Update existing member row
    const dataRange = sheet.getDataRange().getValues();
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][2].toLowerCase() === email.toLowerCase()) {
        const row = i + 1;
        sheet.getRange(row, 2, 1, 10).setValues([[
          memberData.name,
          memberData.phone,
          memberData.role,
          memberData.company,
          memberData.industry,
          memberData.ticketSize,
          memberData.stage,
          memberData.linkedin,
          memberData.bio,
          existing.engagementScore
        ]]);
        break;
      }
    }
    return jsonResponse({ success: true, action: "updated" });
  } else {
    // Add new member
    const id = "m_" + Math.random().toString(36).substr(2, 9);
    sheet.appendRow([
      id, memberData.name, memberData.email, memberData.phone, memberData.role,
      memberData.company, memberData.industry, memberData.ticketSize, memberData.stage,
      memberData.linkedin, memberData.bio, 50, new Date().toISOString()
    ]);
    
    // Send automated welcome email to new subscribers
    sendWelcomeEmail(memberData.name, memberData.email);
    
    return jsonResponse({ success: true, action: "created", memberId: id });
  }
}

// --- RSVP RENDERING PAGE ---

function renderRsvpPage(token, immediateStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("Invitations");
  const membersSheet = ss.getSheetByName("Members");
  const eventsSheet = ss.getSheetByName("Events");
  
  const invData = invSheet.getDataRange().getValues();
  const invHeaders = invData[0];
  let invRowIdx = -1;
  let invitation = null;
  
  for (let i = 1; i < invData.length; i++) {
    if (invData[i][5] === token) {
      invRowIdx = i + 1;
      invitation = {};
      invHeaders.forEach((h, index) => {
        invitation[toCamelCase(h)] = invData[i][index];
      });
      break;
    }
  }
  
  if (!invitation) {
    return HtmlService.createHtmlOutput("<h2 style='text-align:center;font-family:sans-serif;margin-top:100px;'>دعوة غير صالحة أو منتهية الصلاحية</h2>");
  }
  
  const members = getSheetData(membersSheet);
  const events = getSheetData(eventsSheet);
  
  const member = members.find(m => m.id === invitation.memberId);
  const event = events.find(e => e.id === invitation.eventId);
  
  // If immediate status is clicked from email buttons
  if (immediateStatus && (immediateStatus === 'Confirmed' || immediateStatus === 'Declined')) {
    const colIdx = invHeaders.indexOf("RSVP Status") + 1;
    invSheet.getRange(invRowIdx, colIdx).setValue(immediateStatus);
    invSheet.getRange(invRowIdx, invHeaders.indexOf("Updated At") + 1).setValue(new Date().toISOString());
    invitation.rsvpStatus = immediateStatus;
    recalculateMemberEngagementScore(invitation.memberId);
  }
  
  // Generate beautiful HTML Page
  const html = `
  <!DOCTYPE html>
  <html dir="rtl" lang="ar">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>تأكيد الحضور - ${event.title}</title>
    <style>
      body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background-color: #0d1117;
        color: #c9d1d9;
        margin: 0;
        padding: 20px;
        display: flex;
        justify-content: center;
        align-items: center;
        min-height: 100vh;
      }
      .card {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 16px;
        padding: 40px;
        max-width: 500px;
        width: 100%;
        box-shadow: 0 10px 30px rgba(0,0,0,0.5);
        text-align: center;
      }
      .logo {
        background: linear-gradient(135deg, #1f6feb 0%, #00d2ff 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 26px;
        font-weight: 800;
        margin-bottom: 24px;
      }
      h2 {
        color: #ffffff;
        margin-top: 0;
        font-size: 22px;
      }
      .welcome {
        color: #58a6ff;
        font-size: 16px;
        margin-bottom: 20px;
      }
      .event-box {
        background: #0d1117;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 20px;
        margin: 20px 0;
        text-align: right;
      }
      .event-box p {
        margin: 8px 0;
        font-size: 15px;
      }
      .label {
        font-weight: bold;
        color: #8b949e;
      }
      .btn-container {
        display: flex;
        gap: 16px;
        margin-top: 30px;
      }
      .btn {
        flex: 1;
        padding: 14px;
        border-radius: 8px;
        font-size: 16px;
        font-weight: bold;
        cursor: pointer;
        border: none;
        transition: all 0.2s;
        text-decoration: none;
      }
      .btn-confirm {
        background: #238636;
        color: white;
      }
      .btn-confirm:hover { background: #2ea043; }
      .btn-decline {
        background: #21262d;
        color: #c9d1d9;
        border: 1px solid #30363d;
      }
      .btn-decline:hover { background: #30363d; }
      .success-message {
        display: none;
        padding: 15px;
        background: rgba(35, 134, 54, 0.15);
        border: 1px solid #238636;
        color: #56d364;
        border-radius: 8px;
        margin-top: 20px;
        font-weight: bold;
      }
      .status-badge {
        display: inline-block;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 14px;
        font-weight: bold;
        margin-bottom: 20px;
      }
      .status-Pending { background: #c6902622; color: #e3b341; border: 1px solid #c6902655; }
      .status-Confirmed { background: #23863622; color: #56d364; border: 1px solid #23863655; }
      .status-Declined { background: #da363722; color: #f85149; border: 1px solid #da363755; }
    </style>
  </head>
  <body>
    <div class="card">
      <div class="logo">مجتمع دروب | Droob Community</div>
      <h2>تأكيد الحضور للفعالية</h2>
      <div class="welcome">مرحباً بك، <strong>${member.name}</strong> (${member.company})</div>
      
      <div class="event-box">
        <p><span class="label">الفعالية:</span> ${event.title}</p>
        <p><span class="label">التاريخ:</span> ${event.date}</p>
        <p><span class="label">الموقع:</span> ${event.location}</p>
      </div>

      <div class="status-badge status-${invitation.rsvpStatus}" id="badge">
        حالة دعوتك الحالية: ${translateStatus(invitation.rsvpStatus)}
      </div>

      <div id="success" class="success-message">
        تم تحديث حالة حضورك بنجاح! شكراً لك.
      </div>

      <div class="btn-container" id="buttons">
        <button class="btn btn-confirm" onclick="submitRsvp('Confirmed')">تأكيد الحضور</button>
        <button class="btn btn-decline" onclick="submitRsvp('Declined')">اعتذار</button>
      </div>
    </div>

    <script>
      function submitRsvp(status) {
        document.getElementById('buttons').style.opacity = '0.5';
        document.getElementById('buttons').style.pointerEvents = 'none';
        
        // Fetch to this exact App Script URL with status parameter
        const url = window.location.href + '&status=' + status;
        
        fetch(url, { method: 'GET' })
          .then(() => {
            const badge = document.getElementById('badge');
            badge.className = 'status-badge status-' + status;
            badge.innerText = 'حالة دعوتك الحالية: ' + (status === 'Confirmed' ? 'مؤكد' : 'معتذر');
            
            const success = document.getElementById('success');
            success.style.display = 'block';
            success.innerText = status === 'Confirmed' ? 'تم تأكيد حضورك بنجاح! نتطلع لرؤيتك.' : 'تم تسجيل اعتذارك. نتمنى رؤيتك في فعاليات قادمة.';
            
            document.getElementById('buttons').style.display = 'none';
          })
          .catch(err => {
            alert('حدث خطأ أثناء حفظ الرد. يرجى المحاولة لاحقاً.');
            document.getElementById('buttons').style.opacity = '1';
            document.getElementById('buttons').style.pointerEvents = 'all';
          });
      }
      
      // Auto-show success if clicked from email
      if ("${immediateStatus}" !== "") {
        document.getElementById('success').style.display = 'block';
        document.getElementById('buttons').style.display = 'none';
      }
    </script>
  </body>
  </html>
  `;
  
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- UTILITY FUNCTIONS ---

function getSheetData(sheet) {
  if (!sheet) return [];
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];
  
  const headers = values[0];
  const items = [];
  
  for (let i = 1; i < values.length; i++) {
    const item = {};
    headers.forEach((header, index) => {
      const key = toCamelCase(header);
      item[key] = values[i][index];
    });
    items.push(item);
  }
  
  return items;
}

function findRowIndex(sheet, colIndex1, val1, colIndex2, val2) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex1 - 1] === val1 && data[i][colIndex2 - 1] === val2) {
      return i + 1; // 1-indexed row number
    }
  }
  return -1;
}

function toCamelCase(str) {
  return str.toLowerCase().replace(/[^a-zA-Z0-9]+(.)/g, (m, chr) => chr.toUpperCase())
            .replace(/[^a-zA-Z0-9]/g, '');
}

function translateStatus(status) {
  if (status === 'Confirmed') return 'مؤكد';
  if (status === 'Declined') return 'معتذر';
  return 'معلق';
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Send standard HTML email via GmailApp
function sendHtmlEmail(name, email, title, date, location, description, rsvpUrl) {
  try {
    const subject = `دعوة حضور: ${title}`;
    const body = `مرحباً بك، ${name}. تم دعوتك لحضور فعالية ${title}. يرجى تأكيد حضورك عبر الرابط التالي: ${rsvpUrl}`;
    
    // Responsive HTML Email Body
    const htmlBody = `
    <div dir="rtl" style="font-family:sans-serif; background-color:#0d1117; color:#c9d1d9; padding:30px; border-radius:12px; max-width:600px; margin:0 auto; border:1px solid #30363d; text-align:right;">
      <h2 style="color:#ffffff; border-bottom:1px solid #30363d; padding-bottom:15px; text-align:center;">دعوة حضور فعالية</h2>
      <p style="font-size:16px;">أهلاً بك، <strong>سيد/ة ${name}</strong> 👋</p>
      <p>يسرنا دعوتك لحضور اللقاء القادم لـ <strong>مجتمع دروب</strong>:</p>
      
      <div style="background-color:#161b22; border:1px solid #30363d; border-radius:8px; padding:20px; margin:20px 0;">
        <p style="margin:5px 0;"><strong>الفعالية:</strong> ${title}</p>
        <p style="margin:5px 0;"><strong>التاريخ والوقت:</strong> ${date}</p>
        <p style="margin:5px 0;"><strong>المكان:</strong> ${location}</p>
        <p style="font-size:14px; color:#8b949e; border-top:1px dashed #30363d; padding-top:10px; margin-top:10px;">${description}</p>
      </div>
      
      <p style="text-align:center; font-weight:bold; margin-top:30px;">يرجى تأكيد حضورك بالضغط على أحد الروابط أدناه:</p>
      <div style="text-align:center; margin:25px 0;">
        <a href="${rsvpUrl}&status=Confirmed" style="display:inline-block; background-color:#238636; color:#ffffff; padding:12px 30px; text-decoration:none; border-radius:6px; font-weight:bold; font-size:16px; margin:0 10px;">تأكيد الحضور</a>
        <a href="${rsvpUrl}&status=Declined" style="display:inline-block; background-color:#21262d; color:#c9d1d9; padding:12px 30px; text-decoration:none; border-radius:6px; font-weight:bold; font-size:16px; border:1px solid #30363d; margin:0 10px;">الاعتذار</a>
      </div>
    </div>
    `;
    
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: htmlBody,
      name: "مجتمع دروب | Droob Community"
    });
    
    return true;
  } catch (err) {
    Logger.log("Email error: " + err.toString());
    return false;
  }
}

// Send standard Welcome email for new subscribers
function sendWelcomeEmail(name, email) {
  try {
    const subject = `مرحباً بك في مجتمع دروب 👋`;
    const plainBody = `أهلاً ${name}، يسعدنا انضمامك إلى مجتمع دروب.`;
    const htmlBody = getEmailHeader() +
      `<p style="font-size:16px;color:#f0f0f8;margin-bottom:4px;">أهلاً <strong>${name}</strong> 👋</p>
      <p style="color:#9090b8;font-size:14px;line-height:1.8;margin-bottom:20px;">يسعدنا انضمامك إلى مجتمع دروب — المجتمع الذي يجمع نخبة المستثمرين الملائكيين ورواد الأعمال في المملكة العربية السعودية.</p>` +
      getDivider() +
      getInfoCard(
        `<p style="color:#9090b8;font-size:13px;margin:4px 0;"><span style="color:#3b82f6;">●</span> وصول حصري للفعاليات المغلقة والديوانيات</p>
        <p style="color:#9090b8;font-size:13px;margin:4px 0;"><span style="color:#3b82f6;">●</span> توافق ذكي مع المستثمرين أو رواد الأعمال المناسبين</p>
        <p style="color:#9090b8;font-size:13px;margin:4px 0;"><span style="color:#3b82f6;">●</span> مسار نمو داخل المجتمع يكافئ مشاركتك الحقيقية</p>`
      ) +
      getDivider() +
      `<p style="color:#9090b8;font-size:13px;line-height:1.8;">سيتواصل معك الفريق قريباً لترتيب أول خطوة. في الأثناء، إذا كان لديك أي استفسار لا تتردد في التواصل معنا مباشرة.</p>
      <div style="text-align:center;margin-top:24px;">` +
      getBtnGhost('mailto:M.akram@doroobangels.com', 'تواصل مع الفريق') +
      `</div>` +
      getEmailFooter();
    GmailApp.sendEmail(email, subject, plainBody, { htmlBody, replyTo: SENDER_EMAIL, name: SENDER_NAME });
  } catch(err) {
    Logger.log('Welcome email error: ' + err.toString());
  }
}

// --- NEWSLETTER, CONTACT, INTEREST HANDLERS ---

function handleAddNewsletterSubscriber(subscriber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Newsletter");
  if (!sheet) { setup(); return handleAddNewsletterSubscriber(subscriber); }

  const data = sheet.getDataRange().getValues();
  const duplicate = data.slice(1).find(row => row[2].toString().toLowerCase() === (subscriber.email || "").toLowerCase());
  if (duplicate) return jsonResponse({ success: false, error: "already_subscribed" });

  const id = subscriber.id || ("ns_" + Math.random().toString(36).substr(2, 9));
  const now = subscriber.subscribedAt || new Date().toISOString();
  sheet.appendRow([id, subscriber.name || "", subscriber.email || "", now]);

  try {
    const htmlBody = `<div dir="rtl" style="font-family:sans-serif;background:#f7f4ef;color:#3d3d3d;padding:30px;max-width:600px;margin:0 auto;border:1px solid rgba(0,0,0,0.1);border-radius:12px;">
      <h2 style="color:#1b4332;text-align:center;">شكراً لاشتراكك في نشرة مجتمع دروب! 📬</h2>
      <p>أهلاً <strong>${subscriber.name}</strong>، سيصلك كل جديد من مجتمعنا مباشرة على بريدك.</p>
      <hr style="border:none;border-top:1px solid rgba(0,0,0,0.1);margin:20px 0;">
      <p style="font-size:13px;color:#718096;margin:0;">Mohamed Akram<br>Community Manager — مجتمع دروب<br>📞 +966 549311704</p>
    </div>`;
    GmailApp.sendEmail(subscriber.email, "تم اشتراكك في نشرة مجتمع دروب", "", { htmlBody, name: "مجتمع دروب | Droob Community" });
  } catch(e) { Logger.log("Newsletter email error: " + e); }

  return jsonResponse({ success: true, id });
}

function handleAddContactMessage(msg) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ContactMessages");
  if (!sheet) { setup(); return handleAddContactMessage(msg); }

  const id = msg.id || ("cm_" + Math.random().toString(36).substr(2, 9));
  const now = msg.createdAt || new Date().toISOString();
  sheet.appendRow([id, msg.name || "", msg.email || "", msg.phone || "", msg.message || "", now, "false"]);

  try {
    const notifyEmail = Session.getActiveUser().getEmail();
    if (notifyEmail) {
      GmailApp.sendEmail(notifyEmail,
        `رسالة جديدة من ${msg.name} - مجتمع دروب`,
        `الاسم: ${msg.name}\nالبريد: ${msg.email}\nالهاتف: ${msg.phone || '-'}\nالرسالة:\n${msg.message}`,
        { name: "مجتمع دروب | Droob Community" }
      );
    }
  } catch(e) { Logger.log("Contact notify error: " + e); }

  return jsonResponse({ success: true, id });
}

function handleAddInterestRegistration(reg) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("InterestRegistrations");
  if (!sheet) { setup(); return handleAddInterestRegistration(reg); }

  const id = reg.id || ("ir_" + Math.random().toString(36).substr(2, 9));
  const now = reg.createdAt || new Date().toISOString();
  sheet.appendRow([id, reg.name || "", reg.email || "", reg.phone || "", reg.company || "", reg.linkedin || "", reg.interestType || "", reg.message || "", now]);

  try {
    const htmlBody = `<div dir="rtl" style="font-family:sans-serif;background:#f7f4ef;color:#3d3d3d;padding:30px;max-width:600px;margin:0 auto;border:1px solid rgba(0,0,0,0.1);border-radius:12px;">
      <h2 style="color:#1b4332;text-align:center;">تم استلام طلب اهتمامك ✅</h2>
      <p>أهلاً <strong>${reg.name}</strong>، تم تسجيل اهتمامك بـ <strong>${reg.interestType}</strong> بنجاح.</p>
      <p>سيتواصل معك فريق مجتمع دروب في أقرب وقت.</p>
      ${reg.linkedin ? `<p style="font-size:13px;color:#718096;">لينكد إن: <a href="${reg.linkedin}" style="color:#40916c;">${reg.linkedin}</a></p>` : ''}
      <hr style="border:none;border-top:1px solid rgba(0,0,0,0.1);margin:20px 0;">
      <p style="font-size:13px;color:#718096;margin:0;">Mohamed Akram<br>Community Manager — مجتمع دروب<br>📞 +966 549311704</p>
    </div>`;
    GmailApp.sendEmail(reg.email, "تم استلام طلبك - مجتمع دروب", "", { htmlBody, name: "مجتمع دروب | Droob Community" });
  } catch(e) { Logger.log("Interest email error: " + e); }

  return jsonResponse({ success: true, id });
}

function handleEventInvite(data) {
  const subject = "دعوة: " + (data.eventName || "فعالية دروب");
  const body = `
    <div dir="rtl" style="font-family:Arial,sans-serif;background:#f7f4ef;padding:32px;max-width:600px;margin:0 auto;border-radius:12px;border:1px solid rgba(0,0,0,0.1);">
      <h2 style="color:#1b4332;border-bottom:2px solid #1b4332;padding-bottom:12px;">دعوة لحضور فعالية</h2>
      <p style="font-size:16px;">أهلاً بك، <strong>${data.memberName || ''}</strong></p>
      <p style="color:#3d3d3d;">يسرنا دعوتك لحضور فعالية مجتمع دروب القادمة:</p>
      <div style="background:#edeade;border-radius:8px;padding:20px;margin:20px 0;border:1px solid rgba(0,0,0,0.08);">
        <p style="margin:6px 0;"><strong>📅 الفعالية:</strong> ${data.eventName || ''}</p>
        <p style="margin:6px 0;"><strong>🗓️ التاريخ:</strong> ${data.eventDate || ''}</p>
        <p style="margin:6px 0;"><strong>📍 المكان:</strong> ${data.eventLocation || ''}</p>
        ${data.eventMode ? `<p style="margin:6px 0;"><strong>📡 الطريقة:</strong> ${data.eventMode}</p>` : ''}
      </div>
      <div style="text-align:center;margin:28px 0;">
        <a href="${data.rsvpLink}&status=Confirmed" style="display:inline-block;background:#1b4332;color:#fff;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px;margin-left:12px;">✅ تأكيد الحضور</a>
        <a href="${data.rsvpLink}&status=Declined" style="display:inline-block;background:#f0ece4;color:#3d3d3d;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px;border:1px solid rgba(0,0,0,0.15);">❌ الاعتذار</a>
      </div>
      <hr style="border:none;border-top:1px solid rgba(0,0,0,0.1);margin:24px 0;">
      <p style="font-size:12px;color:#718096;text-align:center;">مجتمع دروب — شبكة المستثمرين ورواد الأعمال</p>
      <p style="font-size:12px;color:#718096;text-align:center;">Mohamed Akram | Community Manager | 📞 +966 549311704</p>
    </div>`;
  try {
    GmailApp.sendEmail(data.to, subject, '', { htmlBody: body });
    return jsonResponse({ success: true });
  } catch(e) {
    return jsonResponse({ error: e.toString() });
  }
}

function handleEventInviteExternal(data) {
  const subject = "أنت مدعو لحضور " + (data.eventName || "فعالية دروب") + " — دروب";
  const body = `
    <div dir="rtl" style="font-family:Arial,sans-serif;background:#f7f4ef;padding:32px;max-width:600px;margin:0 auto;border-radius:12px;border:1px solid rgba(0,0,0,0.1);">
      <h2 style="color:#1b4332;border-bottom:2px solid #1b4332;padding-bottom:12px;">تشرفنا بدعوتك</h2>
      <p style="font-size:16px;">أهلاً بك، <strong>${data.guestName || ''}</strong></p>
      <p style="color:#3d3d3d;">يسعدنا دعوتك لحضور فعالية خاصة بمجتمع دروب:</p>
      <div style="background:#edeade;border-radius:8px;padding:20px;margin:20px 0;border:1px solid rgba(0,0,0,0.08);">
        <p style="margin:6px 0;"><strong>📅 الفعالية:</strong> ${data.eventName || ''}</p>
        <p style="margin:6px 0;"><strong>🗓️ التاريخ:</strong> ${data.eventDate || ''}</p>
        <p style="margin:6px 0;"><strong>📍 المكان:</strong> ${data.eventLocation || ''}</p>
      </div>
      <div style="text-align:center;margin:28px 0;">
        <a href="${data.registrationLink}" style="display:inline-block;background:#1b4332;color:#fff;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px;">سجّل وأكد حضورك</a>
      </div>
      <p style="font-size:12px;color:#718096;text-align:center;">سيتم إنشاء حسابك في مجتمع دروب تلقائياً عند التسجيل.</p>
      <hr style="border:none;border-top:1px solid rgba(0,0,0,0.1);margin:24px 0;">
      <p style="font-size:12px;color:#718096;text-align:center;">مجتمع دروب — شبكة المستثمرين ورواد الأعمال</p>
      <p style="font-size:12px;color:#718096;text-align:center;">Mohamed Akram | Community Manager | 📞 +966 549311704</p>
    </div>`;
  try {
    GmailApp.sendEmail(data.to, subject, '', { htmlBody: body });
    return jsonResponse({ success: true });
  } catch(e) {
    return jsonResponse({ error: e.toString() });
  }
}

function handleAddExternalGuest(eventId, guest) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("EventGuests");
  if (!sheet) {
    sheet = ss.insertSheet("EventGuests");
    sheet.appendRow(["eventId","guestName","guestEmail","guestPhone","addedBy","addedAt","isExternal","rsvpStatus","checkedIn"]);
  }
  const now = new Date().toISOString();
  sheet.appendRow([
    eventId,
    guest.name || "",
    guest.email || "",
    guest.phone || "",
    guest.addedBy || "",
    now,
    true,
    "Pending",
    false
  ]);
  return jsonResponse({ success: true });
}

// Recalculates engagement score inside sheet data
function recalculateMemberEngagementScore(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("Invitations");
  const membersSheet = ss.getSheetByName("Members");
  
  const invs = getSheetData(invSheet).filter(i => i.memberId === memberId);
  if (invs.length === 0) return;
  
  const total = invs.length;
  const confirmed = invs.filter(i => i.rsvpStatus === 'Confirmed').length;
  const declined = invs.filter(i => i.rsvpStatus === 'Declined').length;
  const attended = invs.filter(i => i.attended === 'Yes').length;
  const missed = invs.filter(i => i.rsvpStatus === 'Confirmed' && i.attended === 'No').length;
  
  let score = 50; // Base score
  
  const attendanceRate = attended / total;
  const rsvpRate = (confirmed + declined) / total;
  
  score += Math.round(attendanceRate * 40); 
  score += Math.round(rsvpRate * 15);       
  score -= missed * 15;                     
  
  score = Math.max(0, Math.min(100, score));
  
  // Update in sheet
  const membersData = membersSheet.getDataRange().getValues();
  for (let i = 1; i < membersData.length; i++) {
    if (membersData[i][0] === memberId) {
      const row = i + 1;
      const scoreColIdx = membersData[0].indexOf("Engagement Score") + 1;
      membersSheet.getRange(row, scoreColIdx).setValue(score);
      break;
    }
  }
}

function handleNewsletterSend(data) {
  const subject = data.subject || 'نشرة مجتمع دروب';
  const htmlContent = data.body || '';
  const recipients = data.recipients || [];

  const emailTemplate = `
    <div dir="rtl" style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;">
      <div style="background:#1b4332;padding:24px;text-align:center;">
        <div style="color:white;font-size:22px;font-weight:800;">مجتمع دروب | Droob</div>
      </div>
      <div style="padding:32px;background:#ffffff;">
        <h2 style="color:#1b4332;margin-bottom:20px;">${subject}</h2>
        <div style="line-height:1.8;color:#3d3d3d;font-size:14px;">${htmlContent}</div>
      </div>
      <div style="background:#f7f4ef;padding:20px;text-align:center;font-size:12px;color:#718096;border-top:1px solid rgba(0,0,0,0.08);">
        مجتمع دروب — شبكة المستثمرين ورواد الأعمال<br>
        Mohamed Akram | Community Manager | 📞 +966 549311704<br>
        <a href="#" style="color:#40916c;margin-top:8px;display:inline-block;">إلغاء الاشتراك</a>
      </div>
    </div>`;

  let sent = 0;
  let errors = 0;

  recipients.forEach(r => {
    try {
      GmailApp.sendEmail(r.email, subject, '', {
        htmlBody: emailTemplate,
        name: 'مجتمع دروب | Droob Community'
      });
      sent++;
      Utilities.sleep(100); // avoid rate limits
    } catch(e) {
      Logger.log('Newsletter send error for ' + r.email + ': ' + e);
      errors++;
    }
  });

  return jsonResponse({ success: true, sent, errors });
}
