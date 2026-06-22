/**
 * Google Apps Script for "Droob Community" CRM & Event Automation
 * Copy-paste this script into your Google Sheet (Extensions -> Apps Script)
 * Publish it as a Web App (Deploy -> New Deployment -> Web App)
 * Select Execute as: "Me" (your email) and Who has access: "Anyone"
 */

const SENDER_EMAIL = "M.akram@doroobangels.com";
const SENDER_NAME = "Droob Community | مجتمع دروب";
const SPREADSHEET_ID = "1qbgjnG78au6wjQ4UxPzvKnQWKdZwlAjRasI6EarIc5Q"; // غرفة بيانات المجتمع

// ── Firebase config (same as index.html) ──
const FIREBASE_API_KEY  = "AIzaSyAyn4dAvl721BouB-0Z2tOt3AC6n7eie-E";
const FIREBASE_PROJECT  = "crm-droob";
const FIRESTORE_BASE    = "https://firestore.googleapis.com/v1/projects/" + FIREBASE_PROJECT + "/databases/(default)/documents";

// ── Fetch all docs from a Firestore collection ──
function firestoreGetCollection(collection) {
  var url = FIRESTORE_BASE + "/" + collection + "?key=" + FIREBASE_API_KEY + "&pageSize=500";
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(resp.getContentText());
  if (!json.documents) return [];
  return json.documents.map(function(doc) {
    var obj = { id: doc.name.split('/').pop() };
    Object.keys(doc.fields || {}).forEach(function(k) {
      var v = doc.fields[k];
      obj[k] = v.stringValue !== undefined ? v.stringValue
              : v.integerValue !== undefined ? Number(v.integerValue)
              : v.doubleValue !== undefined ? Number(v.doubleValue)
              : v.booleanValue !== undefined ? v.booleanValue
              : v.arrayValue ? (v.arrayValue.values||[]).map(function(i){ return i.stringValue||''; }).join(', ')
              : '';
    });
    return obj;
  });
}

// ── Main sync function — run manually or via time trigger ──
function syncFromFirebase() {
  var ss = getSpreadsheet();
  Logger.log('Starting Firebase → Sheets sync...');

  function upsertSheet(name, headers) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clearContents();
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    try { sheet.getRange(1,1,1,headers.length).setBackground('#1b4332').setFontColor('#ffffff').setFontWeight('bold'); } catch(e) {}
    return sheet;
  }

  // Members
  var members = firestoreGetCollection('members');
  if (members.length > 0) {
    var sheet = upsertSheet('أعضاء CRM', ['الاسم','البريد','الجوال','الشركة','الوظيفة','الدور','نوع العضوية','نقاط النشاط','الاهتمامات','تاريخ الانضمام']);
    members.forEach(function(m) {
      sheet.appendRow([
        m.name||'', m.email||'', m.phone||'', m.company||'', m.position||'',
        m.role==='Investor'?'مستثمر':'رائد أعمال',
        m.memberType||'مستمع', m.engagementScore||50,
        m.interests||'', m.addedDate||m.createdAt||''
      ]);
    });
    Logger.log('Members synced: ' + members.length);
  }

  // Events
  var events = firestoreGetCollection('events');
  if (events.length > 0) {
    var sheet = upsertSheet('الفعاليات', ['العنوان','التاريخ','المكان','النوع','الحالة','الوصف']);
    events.forEach(function(ev) {
      sheet.appendRow([
        ev.title||'', ev.date||'', ev.location||'',
        ev.type||'ديوانية',
        ev.status==='completed'?'مكتملة':'مقررة',
        ev.description||''
      ]);
    });
    Logger.log('Events synced: ' + events.length);
  }

  // Invitations
  var invitations = firestoreGetCollection('invitations');
  if (invitations.length > 0) {
    var mMap = {}; members.forEach(function(m){ mMap[m.id]=m; });
    var eMap = {}; events.forEach(function(e){ eMap[e.id]=e; });
    var sheet = upsertSheet('الدعوات', ['الفعالية','اسم العضو','البريد','RSVP','حضر','تاريخ الدعوة']);
    invitations.forEach(function(inv) {
      var m = mMap[inv.memberId]||{}, e = eMap[inv.eventId]||{};
      var rsvp = inv.rsvpStatus==='Confirmed'?'مؤكد':inv.rsvpStatus==='Declined'?'معتذر':'معلق';
      sheet.appendRow([e.title||'', m.name||'', m.email||'', rsvp, inv.attended==='Yes'?'حضر':'', inv.sentAt||'']);
    });
    Logger.log('Invitations synced: ' + invitations.length);
  }

  // Interest registrations
  var interests = firestoreGetCollection('interests');
  if (interests.length > 0) {
    var sheet = upsertSheet('طلبات الانضمام', ['الاسم','البريد','الجوال','الدور','التاريخ','الحالة']);
    interests.forEach(function(r) {
      sheet.appendRow([r.name||'', r.email||'', r.phone||'', r.role||'', r.createdAt||'', r.status||'pending']);
    });
  }

  // Meta tab
  var meta = ss.getSheetByName('آخر مزامنة') || ss.insertSheet('آخر مزامنة');
  meta.clearContents();
  meta.appendRow(['آخر مزامنة', new Date().toLocaleString('ar-SA')]);
  meta.appendRow(['أعضاء', members.length]);
  meta.appendRow(['فعاليات', events.length]);
  meta.appendRow(['دعوات', invitations.length]);

  Logger.log('Sync complete ✓');
  return { members: members.length, events: events.length, invitations: invitations.length };
}

// ── Install hourly trigger (run once manually) ──
function installTrigger() {
  // Delete existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncFromFirebase') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFromFirebase').timeBased().everyHours(1).create();
  Logger.log('Hourly trigger installed ✓');
}

function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch(e) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
}

function getEmailHeader() {
  return `<div dir="rtl" style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e2e8f0;"><div style="background:#1b4332;padding:20px 28px;"><div style="float:right;width:36px;height:36px;background:#40916c;border-radius:8px;text-align:center;line-height:36px;font-weight:700;font-size:16px;color:#fff;font-family:monospace;">D</div><div style="float:right;margin-right:10px;"><div style="font-size:15px;font-weight:600;color:#ffffff;">Droob Community</div><div style="font-size:11px;color:#b7e4c7;">مجتمع دروب الاستثماري</div></div><div style="clear:both;"></div></div><div style="padding:28px;background:#ffffff;">`;
}

function getEmailFooter() {
  return `</div><div style="background:#f7f4ef;padding:18px 28px;border-top:1px solid #e2e8f0;text-align:center;"><div style="font-size:12px;color:#718096;margin-bottom:10px;">Droob Community | مجتمع دروب الاستثماري</div><div style="margin-bottom:10px;"><a href="https://www.linkedin.com/company/droobangels" style="color:#1b4332;text-decoration:none;font-size:12px;margin:0 8px;">LinkedIn</a><a href="mailto:M.akram@doroobangels.com" style="color:#1b4332;text-decoration:none;font-size:12px;margin:0 8px;">تواصل معنا</a></div><div style="font-size:11px;color:#a0aec0;">© 2026 Droob Community. جميع الحقوق محفوظة.</div></div></div>`;
}

function getInfoCard(content) {
  return `<div style="background:#f7f4ef;border:1px solid #e2e8f0;border-radius:8px;padding:18px;margin:16px 0;text-align:right;">${content}</div>`;
}

function getDivider() {
  return `<div style="height:1px;background:#e2e8f0;margin:20px 0;"></div>`;
}

function getBtnPrimary(url, label) {
  return `<a href="${url}" style="display:inline-block;background:#1b4332;color:#ffffff;padding:12px 28px;text-decoration:none;border-radius:8px;font-weight:600;font-size:15px;margin:6px;">${label}</a>`;
}

function getBtnGhost(url, label) {
  return `<a href="${url}" style="display:inline-block;background:#ffffff;color:#3d3d3d;padding:12px 28px;text-decoration:none;border-radius:8px;font-weight:500;font-size:14px;border:1px solid #e2e8f0;margin:6px;">${label}</a>`;
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
  const ss = getSpreadsheet();
  
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

  if (!ss.getSheetByName("PortalRegistrations")) {
    const pSheet = ss.insertSheet("PortalRegistrations");
    pSheet.appendRow(["timestamp","role","firstName","lastName","email","phone","linkedin","region","entity","dealCount","ticketSize","investorType","sectors","startupName","stage","askAmount","startupSector","needs","bio","referral","goals","participationType","showInDirectory"]);
  }
  if (!ss.getSheetByName("Communications")) {
    const cSheet = ss.insertSheet("Communications");
    cSheet.appendRow(["timestamp","type","guestId","guestName","eventId","eventName","status","channel"]);
  }
  if (!ss.getSheetByName("NewsletterLog")) {
    const nlSheet = ss.insertSheet("NewsletterLog");
    nlSheet.appendRow(["timestamp","subject","recipientCount","sentBy"]);
  }

  Logger.log("Setup completed successfully!");
}

// --- GOOGLE FORM SUBMIT TRIGGER ---
// Run this function automatically when a form is submitted
function onFormSubmit(e) {
  try {
    const ss = getSpreadsheet();
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
    
    // 2. REST API Actions — support JSONP for CORS-free browser fetch
    if (action === 'getData') {
      const callback = e.parameter.callback;
      const data = handleGetDataRaw();
      if (callback) return jsonpResponse(data, callback);
      return jsonResponse(data);
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

// ── CHUNKED SYNC HANDLERS (called from doPost) ──

function syncMembers(members) {
  if (!members || members.length === 0) return jsonResponse({ success: true, skipped: true });
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('أعضاء CRM');
  if (!sheet) { sheet = ss.insertSheet('أعضاء CRM'); }
  sheet.clearContents();
  const headers = ['الاسم','البريد','الجوال','الشركة','الوظيفة','الدور','نوع العضوية','نقاط النشاط','الاهتمامات','تاريخ الانضمام'];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  try { sheet.getRange(1,1,1,headers.length).setBackground('#1b4332').setFontColor('#ffffff').setFontWeight('bold'); } catch(e) {}
  members.forEach(function(m) {
    sheet.appendRow([
      m.name||'', m.email||'', m.phone||'', m.company||'', m.position||'',
      m.role==='Investor'?'مستثمر':'رائد أعمال',
      m.memberType||'مستمع', m.engagementScore||50,
      Array.isArray(m.interests)?m.interests.join(', '):(m.interests||''),
      m.addedDate||m.createdAt||''
    ]);
  });
  return jsonResponse({ success: true, count: members.length });
}

function syncEvents(events) {
  if (!events || events.length === 0) return jsonResponse({ success: true, skipped: true });
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('الفعاليات');
  if (!sheet) { sheet = ss.insertSheet('الفعاليات'); }
  sheet.clearContents();
  const headers = ['العنوان','التاريخ','المكان','النوع','الحالة','الوصف'];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  try { sheet.getRange(1,1,1,headers.length).setBackground('#1b4332').setFontColor('#ffffff').setFontWeight('bold'); } catch(e) {}
  events.forEach(function(ev) {
    sheet.appendRow([
      ev.title||'', ev.date||'', ev.location||'',
      ev.type||'ديوانية',
      ev.status==='completed'?'مكتملة':'مقررة',
      ev.description||''
    ]);
  });
  return jsonResponse({ success: true, count: events.length });
}

function syncExtras(data) {
  const ss = getSpreadsheet();

  if (data.invitations && data.invitations.length > 0) {
    let sheet = ss.getSheetByName('الدعوات');
    if (!sheet) { sheet = ss.insertSheet('الدعوات'); }
    sheet.clearContents();
    const headers = ['الفعالية','اسم العضو','البريد','الجوال','RSVP','حضر','التقييم','تاريخ الدعوة'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    try { sheet.getRange(1,1,1,headers.length).setBackground('#1b4332').setFontColor('#ffffff').setFontWeight('bold'); } catch(e) {}
    data.invitations.forEach(function(inv) {
      var rsvp = inv.rsvpStatus==='Confirmed'?'مؤكد':inv.rsvpStatus==='Declined'?'معتذر':'معلق';
      sheet.appendRow([
        inv.eventTitle||'', inv.memberName||'', inv.memberEmail||'', inv.memberPhone||'',
        rsvp, inv.attended==='Yes'?'حضر':'',
        inv.rating||'', inv.sentAt||''
      ]);
    });
  }

  if (data.interestRegistrations && data.interestRegistrations.length > 0) {
    let sheet = ss.getSheetByName('طلبات الانضمام');
    if (!sheet) { sheet = ss.insertSheet('طلبات الانضمام'); }
    sheet.clearContents();
    const headers = ['الاسم','البريد','الجوال','الشركة','الدور','الرسالة','التاريخ','الحالة'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    try { sheet.getRange(1,1,1,headers.length).setBackground('#1b4332').setFontColor('#ffffff').setFontWeight('bold'); } catch(e) {}
    data.interestRegistrations.forEach(function(r) {
      var name = r.fullName||r.name||((r.firstName||'')+' '+(r.lastName||'')).trim();
      sheet.appendRow([
        name, r.email||'', r.phone||'',
        r.entity||r.startupName||r.company||'',
        r.role||r.interestType||'',
        r.bio||r.message||'',
        r.submittedAt||r.date||'',
        r.status||'pending'
      ]);
    });
  }

  var meta = ss.getSheetByName('آخر مزامنة');
  if (!meta) { meta = ss.insertSheet('آخر مزامنة'); }
  meta.clearContents();
  meta.appendRow(['آخر مزامنة', new Date().toLocaleString('ar-SA')]);
  meta.appendRow(['دعوات', data.invitations ? data.invitations.length : 0]);
  meta.appendRow(['طلبات الانضمام', data.interestRegistrations ? data.interestRegistrations.length : 0]);

  return jsonResponse({ success: true });
}

// Handle POST requests (API writes and syncs)
// ── SYNC CRM DATA TO GOOGLE SHEET (DATA ROOM) ──
function syncToSheet(data) {
  const ss = getSpreadsheet();

  function upsertSheet(name, headers) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); }
    sheet.clearContents();
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    try {
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#1b4332').setFontColor('#ffffff').setFontWeight('bold');
    } catch(e) {}
    return sheet;
  }

  // ── Members ──
  if (data.members && data.members.length > 0) {
    const sheet = upsertSheet('أعضاء CRM', [
      'الاسم','البريد','الجوال','الشركة','الوظيفة','الدور',
      'نوع العضوية','نقاط النشاط','الاهتمامات','تاريخ الانضمام'
    ]);
    data.members.forEach(function(m) {
      sheet.appendRow([
        m.name||'', m.email||'', m.phone||'', m.company||'', m.position||'',
        m.role==='Investor'?'مستثمر':'رائد أعمال',
        m.memberType||'مستمع', m.engagementScore||50,
        Array.isArray(m.interests)?m.interests.join(', '):(m.interests||''),
        m.addedDate||m.createdAt||''
      ]);
    });
  }

  // ── Events ──
  if (data.events && data.events.length > 0) {
    const sheet = upsertSheet('الفعاليات', [
      'العنوان','التاريخ','المكان','النوع','الحالة','الوصف'
    ]);
    data.events.forEach(function(ev) {
      sheet.appendRow([
        ev.title||'', ev.date||'', ev.location||'',
        ev.type||'ديوانية',
        ev.status==='completed'?'مكتملة':'مقررة',
        ev.description||''
      ]);
    });
  }

  // ── Invitations ──
  if (data.invitations && data.invitations.length > 0) {
    const sheet = upsertSheet('الدعوات', [
      'الفعالية','اسم العضو','البريد','الجوال','RSVP','حضر','التقييم','تاريخ الدعوة'
    ]);
    data.invitations.forEach(function(inv) {
      var rsvp = inv.rsvpStatus==='Confirmed'?'مؤكد':inv.rsvpStatus==='Declined'?'معتذر':'معلق';
      sheet.appendRow([
        inv.eventTitle||'', inv.memberName||'', inv.memberEmail||'', inv.memberPhone||'',
        rsvp, inv.attended==='Yes'?'حضر':'',
        inv.rating||'', inv.sentAt||''
      ]);
    });
  }

  // ── Interest Registrations ──
  if (data.interestRegistrations && data.interestRegistrations.length > 0) {
    const sheet = upsertSheet('طلبات الانضمام', [
      'الاسم','البريد','الجوال','الشركة','الدور','الرسالة','التاريخ','الحالة'
    ]);
    data.interestRegistrations.forEach(function(r) {
      var name = r.fullName||r.name||((r.firstName||'')+' '+(r.lastName||'')).trim();
      sheet.appendRow([
        name, r.email||'', r.phone||'',
        r.entity||r.startupName||r.company||'',
        r.role||r.interestType||'',
        r.bio||r.message||'',
        r.submittedAt||r.date||'',
        r.status||'pending'
      ]);
    });
  }

  // ── Update last sync timestamp in a meta sheet ──
  var meta = ss.getSheetByName('آخر مزامنة');
  if (!meta) meta = ss.insertSheet('آخر مزامنة');
  meta.clearContents();
  meta.appendRow(['آخر مزامنة', new Date().toLocaleString('ar-SA')]);
  meta.appendRow(['أعضاء', data.members ? data.members.length : 0]);
  meta.appendRow(['فعاليات', data.events ? data.events.length : 0]);
  meta.appendRow(['دعوات', data.invitations ? data.invitations.length : 0]);
  meta.appendRow(['طلبات الانضمام', data.interestRegistrations ? data.interestRegistrations.length : 0]);

  return jsonResponse({ success: true, syncedAt: new Date().toISOString() });
}

function doPost(e) {
  try {
    // Support both JSON body and FormData (multipart)
    let postData;
    if (e.parameter && e.parameter.data) {
      postData = JSON.parse(e.parameter.data);
    } else if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else {
      return jsonResponse({ error: 'no data received' });
    }
    const action = postData.action;
    
    if (action === 'addMember') {
      return handleAddMember(postData.member);
    }
    if (action === 'addEvent') {
      return handleAddEvent(postData.event);
    }
    if (action === 'sendInvitations') {
      return handleSendInvitations(postData);
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
    if (action === 'syncToSheet') { return syncToSheet(postData); }
    if (action === 'syncMembers') { return syncMembers(postData.members); }
    if (action === 'syncEvents') { return syncEvents(postData.events); }
    if (action === 'syncExtras') { return syncExtras(postData); }
    if (action === 'newsletter_send') {
      return handleNewsletterSend(postData);
    }
    if (action === 'rsvp_confirm') {
      return handleRsvpUpdate(postData.guestId, postData.eventId, 'Confirmed', postData);
    }
    if (action === 'rsvp_decline') {
      return handleRsvpUpdate(postData.guestId, postData.eventId, 'Declined', postData);
    }
    if (action === 'portal_registration') {
      return handlePortalRegistration(postData);
    }
    if (action === 'sendHtmlEmail') {
      const sent = sendHtmlEmail(
        postData.name, postData.email, postData.title,
        postData.date, postData.location, postData.description,
        postData.rsvpUrl, postData.ticketId, postData.mapsUrl
      );
      return jsonResponse({ success: sent });
    }

    return jsonResponse({ error: "Invalid action" });
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// --- API HANDLERS ---

function handleGetDataRaw() {
  const ss = getSpreadsheet();
  return {
    members:               getSheetData(ss.getSheetByName("Members")),
    events:                getSheetData(ss.getSheetByName("Events")),
    invitations:           getSheetData(ss.getSheetByName("Invitations")),
    newsletterSubscribers: getSheetData(ss.getSheetByName("Newsletter")),
    contactMessages:       getSheetData(ss.getSheetByName("ContactMessages")),
    interestRegistrations: getSheetData(ss.getSheetByName("InterestRegistrations"))
  };
}

function handleGetData() {
  return jsonResponse(handleGetDataRaw());
}

function handleAddMember(member) {
  const sheet = getSpreadsheet().getSheetByName("Members");
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
  const sheet = getSpreadsheet().getSheetByName("Events");
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

function handleSendInvitations(postData) {
  // Supports two modes:
  // 1. New mode: postData.invitations = [{invId, token, rsvpLink, memberId, memberName, memberEmail}]
  //    with postData.eventTitle, eventDate, eventLocation, eventDescription embedded
  // 2. Legacy fallback: postData.eventId + postData.memberIds (looks up Sheets)

  const results = [];

  // ── New mode: all data embedded from CRM ──
  if (postData.invitations && Array.isArray(postData.invitations)) {
    const eventTitle       = postData.eventTitle || '';
    const eventDate        = postData.eventDate || '';
    const eventLocation    = postData.eventLocation || '';
    const eventDescription = postData.eventDescription || '';

    postData.invitations.forEach(function(inv) {
      if (!inv.memberEmail) return;
      const sent = sendHtmlEmail(
        inv.memberName, inv.memberEmail,
        eventTitle, eventDate, eventLocation, eventDescription,
        inv.rsvpLink
      );
      results.push({ email: inv.memberEmail, sent: sent });
    });

    return jsonResponse({ success: true, results: results });
  }

  // ── Legacy fallback: look up Sheets ──
  const eventId   = postData.eventId;
  const memberIds = postData.memberIds || [];
  const ss = getSpreadsheet();
  const members = getSheetData(ss.getSheetByName("Members"));
  const events   = getSheetData(ss.getSheetByName("Events"));
  const invSheet = ss.getSheetByName("Invitations");

  const event = events.find(function(e) { return e.id === eventId; });
  if (!event) return jsonResponse({ error: "Event not found in Sheets" });

  const webAppUrl = ScriptApp.getService().getUrl();
  const now = new Date().toISOString();

  memberIds.forEach(function(memberId) {
    const member = members.find(function(m) { return m.id === memberId; });
    if (!member) return;

    let invRowIdx = findRowIndex(invSheet, 1, eventId, 2, memberId);
    let token = "";
    let invId = "";

    if (invRowIdx === -1) {
      invId  = "inv_" + Math.random().toString(36).substr(2, 9);
      token  = Math.random().toString(36).substr(2, 15) + Math.random().toString(36).substr(2, 15);
      invSheet.appendRow([invId, eventId, memberId, "Pending", "Pending", token, now, now]);
    } else {
      const rowValues = invSheet.getRange(invRowIdx, 1, 1, 8).getValues()[0];
      invId = rowValues[0];
      token = rowValues[5];
    }

    const rsvpUrl = webAppUrl + "?token=" + token;
    const emailSent = sendHtmlEmail(member.name, member.email, event.title, event.date, event.location, event.description, rsvpUrl);
    results.push({ email: member.email, sent: emailSent });
  });

  return jsonResponse({ success: true, results: results });
}

function handleUpdateInvitation(invitationId, updates) {
  const sheet = getSpreadsheet().getSheetByName("Invitations");
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
  const ss = getSpreadsheet();
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
  const ss = getSpreadsheet();
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

// JSONP response for cross-origin GET requests from browser
function jsonpResponse(data, callback) {
  const cb = callback || 'callback';
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(data) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// Send standard HTML email via GmailApp
function sendHtmlEmail(name, email, title, date, location, description, rsvpUrl, ticketId, mapsUrl) {
  try {
    const subject = `دعوة | ${title} — مجتمع دروب`;
    const plainBody = `أهلاً ${name}، دعوة لحضور ${title} بتاريخ ${date} في ${location}. تأكيد: ${rsvpUrl}`;
    const ticketDisplay = ticketId ? ticketId : '';
    const mapsBtn = mapsUrl ? `<p style="margin:10px 0;"><a href="${mapsUrl}" style="color:#3b82f6;font-size:13px;">🗺️ عرض الموقع على الخريطة</a></p>` : '';
    const htmlBody = getEmailHeader() +
      `<p style="font-size:16px;color:#1a202c;margin-bottom:4px;">أهلاً <strong>${name}</strong>،</p>
      <p style="color:#4a5568;font-size:14px;line-height:1.8;">يسرنا دعوتك لحضور فعالية مجتمع دروب القادمة.</p>` +
      getInfoCard(
        `<p style="color:#1a202c;font-size:16px;font-weight:600;margin:0 0 14px;">${title}</p>
        <p style="color:#4a5568;font-size:13px;margin:6px 0;"><span style="margin-left:8px;">📅</span>${date}</p>
        <p style="color:#4a5568;font-size:13px;margin:6px 0;"><span style="margin-left:8px;">📍</span>${location}</p>` +
        mapsBtn +
        (ticketDisplay ? `<p style="color:#4a5568;font-size:13px;margin:10px 0 0;"><span style="margin-left:8px;">🎫</span>رقم تذكرتك: <strong style="color:#1b4332;font-family:monospace;font-size:15px;letter-spacing:1px;">${ticketDisplay}</strong></p>` : '') +
        getDivider() +
        `<p style="color:#4a5568;font-size:13px;line-height:1.7;">${description || ''}</p>`
      ) +
      `<p style="color:#4a5568;font-size:13px;text-align:center;margin:20px 0 10px;">يرجى تأكيد حضورك:</p>
      <div style="text-align:center;">` +
      getBtnPrimary(rsvpUrl, 'تأكيد الحضور') +
      getBtnGhost(rsvpUrl.replace('rsvp=confirm','rsvp=decline'), 'الاعتذار') +
      `</div>
      <p style="color:#a0aec0;font-size:11px;text-align:center;margin-top:16px;">يرجى الرد قبل يوم من الفعالية</p>` +
      getEmailFooter();
    GmailApp.sendEmail(email, subject, plainBody, { htmlBody, replyTo: SENDER_EMAIL, name: SENDER_NAME });
    return true;
  } catch(err) {
    Logger.log('Email error: ' + err.toString());
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
  const sheet = getSpreadsheet().getSheetByName("Newsletter");
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
    GmailApp.sendEmail(subscriber.email, "تم اشتراكك في نشرة مجتمع دروب", "", { htmlBody, replyTo: SENDER_EMAIL, name: "Droob Community | مجتمع دروب" });
  } catch(e) { Logger.log("Newsletter email error: " + e); }

  return jsonResponse({ success: true, id });
}

function handleAddContactMessage(msg) {
  const sheet = getSpreadsheet().getSheetByName("ContactMessages");
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
        { name: "مجتمع دروب | Droob Community", replyTo: SENDER_EMAIL }
      );
    }
  } catch(e) { Logger.log("Contact notify error: " + e); }

  return jsonResponse({ success: true, id });
}

function handleAddInterestRegistration(reg) {
  const sheet = getSpreadsheet().getSheetByName("InterestRegistrations");
  if (!sheet) { setup(); return handleAddInterestRegistration(reg); }

  const id = reg.id || ("ir_" + Math.random().toString(36).substr(2, 9));
  const now = reg.createdAt || new Date().toISOString();
  sheet.appendRow([id, reg.name || "", reg.email || "", reg.phone || "", reg.company || "", reg.linkedin || "", reg.interestType || "", reg.message || "", now]);

  try {
    sendInterestConfirmationEmail(reg.name, reg.email, reg.interestType || 'entrepreneur');
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
    GmailApp.sendEmail(data.to, subject, '', { htmlBody: body, replyTo: SENDER_EMAIL, name: SENDER_NAME });
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
    GmailApp.sendEmail(data.to, subject, '', { htmlBody: body, replyTo: SENDER_EMAIL, name: SENDER_NAME });
    return jsonResponse({ success: true });
  } catch(e) {
    return jsonResponse({ error: e.toString() });
  }
}

function handleAddExternalGuest(eventId, guest) {
  const ss = getSpreadsheet();
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
  const ss = getSpreadsheet();
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
      sendNewsletterEmail(r.email, subject, htmlContent);
      sent++;
      Utilities.sleep(100);
    } catch(e) {
      Logger.log('Newsletter send error for ' + r.email + ': ' + e);
      errors++;
    }
  });

  return jsonResponse({ success: true, sent, errors });
}

function sendExternalInviteEmail(guestName, guestEmail, eventTitle, eventDate, eventLocation, registrationLink) {
  try {
    const subject = `أنت مدعو | ${eventTitle} — مجتمع دروب`;
    const plainBody = `أهلاً ${guestName}، أنت مدعو لحضور ${eventTitle}. سجّل: ${registrationLink}`;
    const htmlBody = getEmailHeader() +
      `<p style="font-size:16px;color:#f0f0f8;margin-bottom:4px;">أهلاً <strong>${guestName}</strong> 🌟</p>
      <p style="color:#9090b8;font-size:14px;line-height:1.8;">يتشرف مجتمع <strong style="color:#f0f0f8;">دروب</strong> بدعوتك لحضور فعاليتنا القادمة.</p>` +
      getInfoCard(
        `<p style="color:#f0f0f8;font-size:16px;font-weight:600;margin:0 0 14px;">${eventTitle}</p>
        <p style="color:#9090b8;font-size:13px;margin:6px 0;"><span style="color:#3b82f6;margin-left:8px;">📅</span>${eventDate}</p>
        <p style="color:#9090b8;font-size:13px;margin:6px 0;"><span style="color:#3b82f6;margin-left:8px;">📍</span>${eventLocation}</p>`
      ) +
      getDivider() +
      `<p style="color:#9090b8;font-size:13px;line-height:1.8;text-align:right;">مجتمع دروب يجمع نخبة المستثمرين الملائكيين ورواد الأعمال في المملكة. سجّل مكانك الآن وسيتم إنشاء حسابك تلقائياً.</p>
      <div style="text-align:center;margin-top:24px;">` +
      getBtnPrimary(registrationLink, '🔗 سجّل مكانك الآن') +
      `</div>
      <p style="color:#2a2a50;font-size:11px;text-align:center;margin-top:12px;">سيتم إنشاء حسابك في مجتمع دروب تلقائياً عند التسجيل</p>` +
      getEmailFooter();
    GmailApp.sendEmail(guestEmail, subject, plainBody, { htmlBody, replyTo: SENDER_EMAIL, name: SENDER_NAME });
    return true;
  } catch(err) {
    Logger.log('External invite email error: ' + err.toString());
    return false;
  }
}

function sendRsvpConfirmationEmail(name, email, eventTitle, eventDate, eventLocation, ticketId, calendarLink, mapsLink) {
  try {
    const subject = `✅ تم تأكيد حضورك — ${eventTitle}`;
    const plainBody = `تم تأكيد حضورك يا ${name} في ${eventTitle} بتاريخ ${eventDate}. تذكرتك: #DRB-${ticketId}`;
    const htmlBody = getEmailHeader() +
      `<div style="text-align:center;margin-bottom:20px;"><div style="width:52px;height:52px;background:#063d2a;border:1px solid #10b981;border-radius:50%;display:inline-block;line-height:52px;font-size:22px;">✅</div></div>
      <p style="font-size:16px;color:#f0f0f8;text-align:center;font-weight:600;margin-bottom:4px;">تم تأكيد حضورك!</p>
      <p style="color:#9090b8;font-size:14px;text-align:center;margin-bottom:20px;">نتطلع لرؤيتك يا <strong style="color:#f0f0f8;">${name}</strong></p>` +
      getInfoCard(
        `<p style="color:#f0f0f8;font-size:15px;font-weight:600;margin:0 0 12px;">${eventTitle}</p>
        <p style="color:#9090b8;font-size:13px;margin:6px 0;"><span style="color:#10b981;margin-left:8px;">📅</span>${eventDate}</p>
        <p style="color:#9090b8;font-size:13px;margin:6px 0;"><span style="color:#10b981;margin-left:8px;">📍</span>${eventLocation}</p>
        <p style="color:#9090b8;font-size:13px;margin:6px 0;"><span style="color:#10b981;margin-left:8px;">🎫</span>تذكرتك: <strong style="color:#f0f0f8;font-family:monospace;">#DRB-${ticketId}</strong></p>`
      ) +
      `<div style="text-align:center;margin-top:20px;">` +
      (calendarLink ? getBtnPrimary(calendarLink, '📆 أضف للتقويم') : '') +
      (mapsLink ? getBtnGhost(mapsLink, '🗺 عرض الموقع') : '') +
      `</div>` +
      getEmailFooter();
    GmailApp.sendEmail(email, subject, plainBody, { htmlBody, replyTo: SENDER_EMAIL, name: SENDER_NAME });
    return true;
  } catch(err) {
    Logger.log('RSVP confirm email error: ' + err.toString());
    return false;
  }
}

function sendRsvpDeclineEmail(name, email, eventTitle) {
  try {
    const subject = `تم تسجيل اعتذارك — مجتمع دروب`;
    const plainBody = `شكراً ${name} لإخبارنا. تم تسجيل اعتذارك عن ${eventTitle}.`;
    const htmlBody = getEmailHeader() +
      `<p style="font-size:16px;color:#f0f0f8;margin-bottom:4px;">شكراً لإخبارنا <strong>${name}</strong> 🙏</p>
      <p style="color:#9090b8;font-size:14px;line-height:1.8;margin-bottom:20px;">تم تسجيل اعتذارك عن حضور <strong style="color:#f0f0f8;">${eventTitle}</strong>. نتمنى أن نراك في الفعالية القادمة.</p>` +
      getInfoCard(
        `<p style="color:#9090b8;font-size:13px;margin:0;line-height:1.7;">ستصلك دعوة للفعالية القادمة من مجتمع دروب. إذا كنت ترغب بتعديل تفضيلاتك أو لديك أي استفسار، لا تتردد في التواصل معنا.</p>`
      ) +
      `<div style="text-align:center;margin-top:20px;">` +
      getBtnGhost('mailto:M.akram@doroobangels.com', 'تواصل مع الفريق') +
      `</div>` +
      getEmailFooter();
    GmailApp.sendEmail(email, subject, plainBody, { htmlBody, replyTo: SENDER_EMAIL, name: SENDER_NAME });
    return true;
  } catch(err) {
    Logger.log('RSVP decline email error: ' + err.toString());
    return false;
  }
}

function sendInterestConfirmationEmail(name, email, role) {
  try {
    const subject = `وصل طلبك — مجتمع دروب 📩`;
    const roleLabel = (role === 'investor' || role === 'مستثمر') ? 'مستثمر ملائكي' : 'مؤسس شركة ناشئة';
    const plainBody = `أهلاً ${name}، وصلنا طلب انضمامك لمجتمع دروب بصفتك ${roleLabel}.`;
    const htmlBody = getEmailHeader() +
      `<p style="font-size:16px;color:#f0f0f8;margin-bottom:4px;">أهلاً <strong>${name}</strong> 👋</p>
      <p style="color:#9090b8;font-size:14px;line-height:1.8;margin-bottom:20px;">وصلنا طلب انضمامك لمجتمع دروب بصفتك <strong style="color:#f0f0f8;">${roleLabel}</strong>. شكراً على اهتمامك.</p>` +
      getDivider() +
      getInfoCard(
        `<p style="color:#f0f0f8;font-size:13px;font-weight:600;margin:0 0 10px;">ماذا يحدث الآن؟</p>
        <p style="color:#9090b8;font-size:13px;margin:6px 0;"><span style="color:#3b82f6;margin-left:8px;">١</span>سيراجع الفريق طلبك خلال ٤٨ ساعة</p>
        <p style="color:#9090b8;font-size:13px;margin:6px 0;"><span style="color:#3b82f6;margin-left:8px;">٢</span>سيتواصل معك أحد أعضاء الفريق مباشرة</p>
        <p style="color:#9090b8;font-size:13px;margin:6px 0;"><span style="color:#3b82f6;margin-left:8px;">٣</span>ستصلك دعوة لأول فعالية مناسبة لاهتماماتك</p>`
      ) +
      getDivider() +
      `<p style="color:#9090b8;font-size:13px;text-align:center;">أي استفسار؟ تواصل معنا مباشرة</p>
      <div style="text-align:center;margin-top:12px;">` +
      getBtnGhost('mailto:M.akram@doroobangels.com', 'تواصل معنا') +
      `</div>` +
      getEmailFooter();
    GmailApp.sendEmail(email, subject, plainBody, { htmlBody, replyTo: SENDER_EMAIL, name: SENDER_NAME });
    return true;
  } catch(err) {
    Logger.log('Interest email error: ' + err.toString());
    return false;
  }
}

function sendNewsletterEmail(subscriberEmail, subject, contentHtml) {
  try {
    const fullSubject = `${subject} — مجتمع دروب 📰`;
    const plainBody = subject;
    const dateStr = new Date().toLocaleDateString('ar-SA', {year:'numeric',month:'long',day:'numeric'});
    const htmlBody = getEmailHeader() +
      `<div style="display:inline-block;background:#1e1e3a;border:1px solid #2a2a50;border-radius:6px;padding:4px 12px;font-size:11px;color:#93c5fd;font-weight:600;letter-spacing:.05em;margin-bottom:16px;">النشرة البريدية</div>
      <h2 style="color:#f0f0f8;font-size:18px;font-weight:600;margin:0 0 8px;">${subject}</h2>
      <p style="color:#505078;font-size:12px;margin:0 0 20px;">${dateStr}</p>` +
      getDivider() +
      contentHtml +
      getDivider() +
      `<p style="color:#2a2a50;font-size:11px;text-align:center;margin-top:16px;">وصلتك هذه النشرة لأنك مشترك في قائمة مجتمع دروب البريدية.</p>` +
      getEmailFooter();
    GmailApp.sendEmail(subscriberEmail, fullSubject, plainBody, { htmlBody, replyTo: SENDER_EMAIL, name: SENDER_NAME });
    return true;
  } catch(err) {
    Logger.log('Newsletter email error for ' + subscriberEmail + ': ' + err.toString());
    return false;
  }
}

function handleRsvpUpdate(guestId, eventId, status, data) {
  const ss = getSpreadsheet();
  const invSheet = ss.getSheetByName("Invitations");
  const membersSheet = ss.getSheetByName("Members");
  const eventsSheet = ss.getSheetByName("Events");

  const members = getSheetData(membersSheet);
  const events = getSheetData(eventsSheet);
  const member = members.find(m => m.id === guestId);
  const event = events.find(e => e.id === eventId);

  if (!member || !event) return jsonResponse({ error: "Guest or event not found" });

  // Update invitation status
  if (invSheet) {
    const invData = invSheet.getDataRange().getValues();
    const headers = invData[0];
    for (let i = 1; i < invData.length; i++) {
      if (invData[i][headers.indexOf('Member ID')] === guestId && invData[i][headers.indexOf('Event ID')] === eventId) {
        const colIdx = headers.indexOf("RSVP Status") + 1;
        const updatedIdx = headers.indexOf("Updated At") + 1;
        invSheet.getRange(i + 1, colIdx).setValue(status);
        invSheet.getRange(i + 1, updatedIdx).setValue(new Date().toISOString());
        recalculateMemberEngagementScore(guestId);
        break;
      }
    }
  }

  // Log to Communications sheet
  let commSheet = ss.getSheetByName("Communications");
  if (!commSheet) {
    commSheet = ss.insertSheet("Communications");
    commSheet.appendRow(["timestamp","type","guestId","guestName","eventId","eventName","status","channel"]);
  }
  commSheet.appendRow([new Date().toISOString(), 'rsvp', guestId, member.name || '', eventId, event.title || '', status, 'email']);

  // Send confirmation email
  if (status === 'Confirmed') {
    const ticketId = new Date().getFullYear() + String(new Date().getMonth()+1).padStart(2,'0') + '-' + (guestId || '').slice(-4).toUpperCase();
    const mapsLink = event.location ? 'https://maps.google.com/?q=' + encodeURIComponent(event.location) : '';
    sendRsvpConfirmationEmail(member.name, member.email, event.title, event.date, event.location, ticketId, '', mapsLink);
  } else {
    sendRsvpDeclineEmail(member.name, member.email, event.title);
  }

  return jsonResponse({ success: true, status, guestName: member.name, eventName: event.title });
}

function handlePortalRegistration(data) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("PortalRegistrations");
  if (!sheet) {
    sheet = ss.insertSheet("PortalRegistrations");
    sheet.appendRow(["timestamp","role","firstName","lastName","email","phone","linkedin","region","entity","dealCount","ticketSize","investorType","sectors","startupName","stage","askAmount","startupSector","needs","bio","referral","goals","participationType","showInDirectory"]);
  }
  sheet.appendRow([
    new Date().toISOString(),
    data.role || '', data.firstName || '', data.lastName || '',
    data.email || '', data.phone || '', data.linkedin || '', data.region || '',
    data.entity || '', data.dealCount || '', data.ticketSize || '', data.investorType || '',
    Array.isArray(data.sectors) ? data.sectors.join(', ') : (data.sectors || ''),
    data.startupName || '', data.stage || '', data.askAmount || '', data.startupSector || '',
    Array.isArray(data.needs) ? data.needs.join(', ') : (data.needs || ''),
    data.bio || '', data.referral || '',
    Array.isArray(data.goals) ? data.goals.join(', ') : (data.goals || ''),
    data.participationType || '', data.showInDirectory !== false ? 'yes' : 'no'
  ]);

  const name = (data.firstName || '') + ' ' + (data.lastName || '');
  const role = data.role || 'entrepreneur';
  sendInterestConfirmationEmail(name.trim(), data.email, role);

  return jsonResponse({ success: true, message: "Registration received" });
}
