/**
 * Google Apps Script for "Droob Angels" CRM & Event Automation
 * UPDATED: Sender name = Droob Angels, supports FormData payload field
 */

const SENDER_EMAIL = "m.akram@doroobangels.com";
const SENDER_NAME = "Droob Angels";
const SPREADSHEET_ID = "1nWRu6CI9C89AS6HxOShjmOwe7fTCTRiJxb370DgquVY";

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

function syncFromFirebase() {
  var ss = getSpreadsheet();
  Logger.log('Starting Firebase → Sheets sync...');

  function upsertSheet(name, headers) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clearContents();
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    try { sheet.getRange(1,1,1,headers.length).setBackground('#0a0a0a').setFontColor('#D4AF37').setFontWeight('bold'); } catch(e) {}
    return sheet;
  }

  var members = firestoreGetCollection('members');
  if (members.length > 0) {
    var sheet = upsertSheet('أعضاء CRM', ['الاسم','البريد','الجوال','الشركة','الوظيفة','الدور','نوع العضوية','نقاط النشاط','الاهتمامات','تاريخ الانضمام']);
    members.forEach(function(m) {
      sheet.appendRow([
        m.name||'', m.email||'', m.phone||'', m.company||'', m.position||'',
        m.role==='Investor'?'مستثمر':'رائد أعمال',
        m.memberType||'مستمع', m.engagementScore||50,
        Array.isArray(m.interests)?m.interests.join(', '):(m.interests||''),
        m.addedDate||m.createdAt||''
      ]);
    });
  }

  var events = firestoreGetCollection('events');
  if (events.length > 0) {
    var sheet = upsertSheet('الفعاليات', ['العنوان','التاريخ','المكان','النوع','الحالة','الوصف']);
    events.forEach(function(ev) {
      sheet.appendRow([ev.title||'', ev.date||'', ev.location||'', ev.type||'ديوانية', ev.status==='completed'?'مكتملة':'مقررة', ev.description||'']);
    });
  }

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
  }

  var interests = firestoreGetCollection('interests');
  if (interests.length > 0) {
    var sheet = upsertSheet('طلبات الانضمام', ['الاسم','البريد','الجوال','الدور','التاريخ','الحالة']);
    interests.forEach(function(r) {
      sheet.appendRow([r.name||'', r.email||'', r.phone||'', r.role||'', r.createdAt||'', r.status||'pending']);
    });
  }

  // ── English sheets (Members, Events, Invitations) ──
  if (members.length > 0) {
    var ms = ss.getSheetByName('Members') || ss.insertSheet('Members');
    ms.clearContents();
    ms.appendRow(['ID','Name','Email','Phone','Role','Company','Industry','Ticket Size','Stage','LinkedIn','Bio','Engagement Score','Member Type','Participation Type','Date Added']);
    ms.setFrozenRows(1);
    members.forEach(function(m) {
      ms.appendRow([m.id||'',m.name||'',m.email||'',m.phone||'',m.role||'',m.company||'',m.industry||'',m.ticketSize||'',m.stage||'',m.linkedin||'',m.bio||'',m.engagementScore||50,m.memberType||'',m.participationType||'',m.addedDate||m.createdAt||'']);
    });
  }
  if (events.length > 0) {
    var es = ss.getSheetByName('Events') || ss.insertSheet('Events');
    es.clearContents();
    es.appendRow(['ID','Title','Date','Location','Type','Status','Description','Created At']);
    es.setFrozenRows(1);
    events.forEach(function(ev) {
      es.appendRow([ev.id||'',ev.title||'',ev.date||'',ev.location||'',ev.type||'',ev.status||'',ev.description||'',ev.createdAt||'']);
    });
  }
  if (invitations.length > 0) {
    var mMap2 = {}; members.forEach(function(m){ mMap2[m.id]=m; });
    var eMap2 = {}; events.forEach(function(e){ eMap2[e.id]=e; });
    var is = ss.getSheetByName('Invitations') || ss.insertSheet('Invitations');
    is.clearContents();
    is.appendRow(['ID','Event ID','Event Title','Member ID','Member Name','Member Email','Member Phone','RSVP Status','Attended','Token','Sent At','Updated At']);
    is.setFrozenRows(1);
    invitations.forEach(function(inv) {
      var m2 = mMap2[inv.memberId]||{}, e2 = eMap2[inv.eventId]||{};
      is.appendRow([inv.id||'',inv.eventId||'',e2.title||'',inv.memberId||'',m2.name||inv.guestName||'',m2.email||inv.guestEmail||'',m2.phone||inv.guestPhone||'',inv.rsvpStatus||'',inv.attended||'',inv.token||'',inv.sentAt||'',inv.updatedAt||'']);
    });
  }
  if (interests.length > 0) {
    var irs = ss.getSheetByName('InterestRegistrations') || ss.insertSheet('InterestRegistrations');
    irs.clearContents();
    irs.appendRow(['ID','Name','Email','Phone','Role','Company','LinkedIn','Message','Created At','Status']);
    irs.setFrozenRows(1);
    interests.forEach(function(r) {
      irs.appendRow([r.id||'',r.name||'',r.email||'',r.phone||'',r.role||'',r.company||'',r.linkedin||'',r.bio||r.message||'',r.createdAt||'',r.status||'pending']);
    });
  }

  var meta = ss.getSheetByName('آخر مزامنة') || ss.insertSheet('آخر مزامنة');
  meta.clearContents();
  meta.appendRow(['آخر مزامنة', new Date().toLocaleString('ar-SA')]);
  meta.appendRow(['أعضاء', members.length]);
  meta.appendRow(['فعاليات', events.length]);
  meta.appendRow(['دعوات', invitations.length]);

  Logger.log('Sync complete');
  return { members: members.length, events: events.length, invitations: invitations.length };
}

function installAllTriggers() {
  // Remove existing triggers for both functions
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'syncFromFirebase' || fn === 'processEmailJobs') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFromFirebase').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('processEmailJobs').timeBased().everyMinutes(1).create();
  Logger.log('Triggers installed: syncFromFirebase (hourly), processEmailJobs (every minute)');
}

// Keep legacy name as alias
function installTrigger() {
  installAllTriggers();
}

function getSpreadsheet() {
  try { return SpreadsheetApp.getActiveSpreadsheet(); }
  catch(e) { return SpreadsheetApp.openById(SPREADSHEET_ID); }
}

function getEmailHeader() {
  return '<div dir="rtl" style="font-family:\'Segoe UI\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#0a0a0a;border-radius:12px;overflow:hidden;">' +
    '<div style="background:#0a0a0a;padding:24px 32px;border-bottom:1px solid #222;">' +
    '<div style="font-size:20px;font-weight:700;color:#D4AF37;letter-spacing:1px;">Droob Angels</div>' +
    '<div style="font-size:11px;color:#666;margin-top:2px;">دروب للملائكة</div>' +
    '</div>' +
    '<div style="padding:32px;background:#111;color:#ddd;">';
}

function getEmailFooter() {
  return '</div>' +
    '<div style="background:#0a0a0a;padding:18px 32px;text-align:center;border-top:1px solid #222;">' +
    '<div style="margin-bottom:10px;">' +
    '<a href="https://doroobangels.com/diwaniya" style="color:#D4AF37;text-decoration:none;font-size:13px;font-weight:600;">🌐 doroobangels.com</a>' +
    '</div>' +
    '<div style="margin-bottom:8px;">' +
    '<a href="https://www.linkedin.com/company/doroob-angels-investors/" style="display:inline-block;background:#0A66C2;color:#fff;padding:5px 14px;border-radius:4px;text-decoration:none;font-size:12px;font-weight:600;margin:0 4px;">in LinkedIn</a>' +
    '<a href="mailto:' + SENDER_EMAIL + '" style="color:#D4AF37;text-decoration:none;font-size:12px;margin:0 8px;">تواصل معنا</a>' +
    '</div>' +
    '<div style="font-size:11px;color:#444;margin-top:6px;">Droob Angels — مجتمع الملائكة</div>' +
    '</div></div>';
}

function getInfoCard(content) {
  return '<div style="background:#1a1a1a;border:1px solid #2a2a2a;border-right:4px solid #D4AF37;border-radius:8px;padding:18px;margin:16px 0;">' + content + '</div>';
}

function getDivider() {
  return '<div style="height:1px;background:#222;margin:20px 0;"></div>';
}

function getBtnPrimary(url, label) {
  return '<a href="' + url + '" style="display:inline-block;background:#D4AF37;color:#0a0a0a;padding:12px 28px;text-decoration:none;border-radius:8px;font-weight:700;font-size:15px;margin:6px;">' + label + '</a>';
}

function getBtnGhost(url, label) {
  return '<a href="' + url + '" style="display:inline-block;background:transparent;color:#D4AF37;padding:12px 28px;text-decoration:none;border-radius:8px;font-weight:500;font-size:14px;border:1px solid #D4AF37;margin:6px;">' + label + '</a>';
}

const COLUMN_MAPPING = {
  timestamp: 0, name: 1, email: 2, phone: 3, role: 4,
  company: 5, industry: 6, ticketSize: 7, stage: 8, linkedin: 9, bio: 10
};

function setup() {
  const ss = getSpreadsheet();
  const sheets = {
    "Members": ["ID","Name","Email","Phone","Role","Company","Industry","Ticket Size","Stage","LinkedIn","Bio","Engagement Score","Member Type","Participation Type","Date Added"],
    "Events": ["ID","Title","Date","Location","Description","Date Created"],
    "Invitations": ["ID","Event ID","Member ID","RSVP Status","Attended","Token","Sent At","Updated At"],
    "Newsletter": ["ID","Name","Email","Subscribed At"],
    "ContactMessages": ["ID","Name","Email","Phone","Message","Created At","Is Read"],
    "InterestRegistrations": ["ID","Name","Email","Phone","Company","LinkedIn","Interest Type","Message","Created At"],
    "EventGuests": ["eventId","guestName","guestEmail","guestPhone","addedBy","addedAt","isExternal","rsvpStatus","checkedIn"],
    "PortalRegistrations": ["timestamp","role","firstName","lastName","email","phone","linkedin","region","entity","dealCount","ticketSize","investorType","sectors","startupName","stage","askAmount","startupSector","needs","bio","referral","goals","participationType","showInDirectory"],
    "Communications": ["timestamp","type","guestId","guestName","eventId","eventName","status","channel"],
    "NewsletterLog": ["timestamp","subject","recipientCount","sentBy"]
  };
  Object.entries(sheets).forEach(([name, headers]) => {
    if (!ss.getSheetByName(name)) {
      const s = ss.insertSheet(name);
      s.appendRow(headers);
      s.setFrozenRows(1);
    }
  });
  Logger.log("Setup completed");
}

function onFormSubmit(e) {
  try {
    const ss = getSpreadsheet();
    if (!ss.getSheetByName("Members")) setup();
    const values = e.values;
    if (!values || !values.length) return;
    const name = values[COLUMN_MAPPING.name]||'', email = (values[COLUMN_MAPPING.email]||'').trim();
    if (!name || !email) return;
    const roleLower = (values[COLUMN_MAPPING.role]||'').toLowerCase();
    const role = (roleLower.includes('invest')||roleLower.includes('مستثمر')) ? 'Investor' : 'Entrepreneur';
    const members = getSheetData(ss.getSheetByName("Members"));
    const existing = members.find(m => m.email.toLowerCase()===email.toLowerCase());
    if (!existing) {
      const id = 'm_'+Math.random().toString(36).substr(2,9);
      ss.getSheetByName("Members").appendRow([id,name,email,values[COLUMN_MAPPING.phone]||'',role,values[COLUMN_MAPPING.company]||'',values[COLUMN_MAPPING.industry]||'',values[COLUMN_MAPPING.ticketSize]||'',values[COLUMN_MAPPING.stage]||'',values[COLUMN_MAPPING.linkedin]||'',values[COLUMN_MAPPING.bio]||'',50,new Date().toISOString()]);
      sendWelcomeEmail(name, email);
    }
  } catch(err) { Logger.log("onFormSubmit error: "+err); }
}

function doGet(e) {
  try {
    if (!e) return HtmlService.createHtmlOutput('<p>No event</p>');
    const params = e.parameter || {};
    const action = params.action, token = params.token;

    // Handle payload sent via GET — returns JSONP if callback provided
    if (params.payload) {
      try {
        const postData = JSON.parse(params.payload);
        const act = postData.action;
        const cb = params.callback || '';
        function jsResp(val) {
          var s = cb ? cb + '(' + JSON.stringify(val) + ')' : JSON.stringify(val);
          var mime = cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
          return ContentService.createTextOutput(s).setMimeType(mime);
        }
        if (act === 'sendInvitations') { var r = handleSendInvitations(postData); return jsResp({success:true}); }
        if (act === 'sendHtmlEmail')   { sendHtmlEmail(postData.name, postData.email, postData.title, postData.date, postData.location, postData.description, postData.rsvpUrl); return jsResp({success:true}); }
        if (act === 'newsletter_send') { handleNewsletterSend(postData); return jsResp({success:true}); }
      } catch(pe) { Logger.log('GET payload parse error: ' + pe); }
    }

    if (action==='syncNow') {
      var r = syncFromFirebase();
      return HtmlService.createHtmlOutput('<html dir="rtl"><head><meta charset="UTF-8"><style>body{font-family:sans-serif;background:#0a0a0a;color:#D4AF37;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;}</style></head><body><div style="text-align:center;background:#111;padding:40px;border-radius:16px;border:1px solid #D4AF37;"><h2>تمت المزامنة</h2><p>أعضاء: '+r.members+' | فعاليات: '+r.events+' | دعوات: '+r.invitations+'</p></div></body></html>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (action==='installTrigger') { installAllTriggers(); return HtmlService.createHtmlOutput('<p>All triggers installed: syncFromFirebase (hourly) + processEmailJobs (every minute)</p>'); }
    if (action==='checkTriggers') { return renderTriggersPage(); }
    if (token) return renderRsvpPage(token, params.status);
    if (action==='getData') {
      const cb = params.callback, data = handleGetDataRaw();
      return cb ? jsonpResponse(data, cb) : jsonResponse(data);
    }
    return HtmlService.createHtmlOutput('<p>Droob Angels Script ✓</p>');
  } catch(err) { return HtmlService.createHtmlOutput('<p>Error: '+err+'</p>'); }
}

// ── doPost — supports FormData (payload), FormData (data), and raw JSON ──
function doPost(e) {
  try {
    if (!e) return jsonResponse({ error: 'no event' });
    let raw = '';
    const params = e.parameter || {};
    const postData = e.postData || {};

    if (params.payload) {
      raw = params.payload;
    } else if (params.data) {
      raw = params.data;
    } else if (postData.contents) {
      raw = postData.contents;
    } else {
      return jsonResponse({ error: 'no data received' });
    }

    const body = JSON.parse(raw);
    const action = body.action;

    if (action==='addMember')              return handleAddMember(body.member);
    if (action==='addEvent')               return handleAddEvent(body.event);
    if (action==='sendInvitations')        return handleSendInvitations(body);
    if (action==='updateInvitation')       return handleUpdateInvitation(body.invitationId, body.updates);
    if (action==='syncForm')               return handleFormSubmitWebhook(body);
    if (action==='addNewsletterSubscriber') return handleAddNewsletterSubscriber(body.subscriber);
    if (action==='addContactMessage')      return handleAddContactMessage(body.message);
    if (action==='addInterestRegistration') return handleAddInterestRegistration(body.registration);
    if (action==='event_invite')           return handleEventInvite(body);
    if (action==='event_invite_external')  return handleEventInviteExternal(body);
    if (action==='addExternalGuest')       return handleAddExternalGuest(body.eventId, body.guest);
    if (action==='syncToSheet')            return syncToSheet(body);
    if (action==='syncMembers')            return syncMembers(body.members);
    if (action==='syncEvents')             return syncEvents(body.events);
    if (action==='syncExtras')             return syncExtras(body);
    if (action==='newsletter_send')        return handleNewsletterSend(body);
    if (action==='rsvp_confirm')           return handleRsvpUpdate(body.guestId, body.eventId, 'Confirmed', body);
    if (action==='rsvp_decline')           return handleRsvpUpdate(body.guestId, body.eventId, 'Declined', body);
    if (action==='portal_registration')    return handlePortalRegistration(body);
    if (action==='upsert_event')           return jsonResponse({ success: true });
    if (action==='upsert_invitation')      return jsonResponse({ success: true });
    if (action==='sendHtmlEmail') {
      const sent = sendHtmlEmail(body.name, body.email, body.title, body.date, body.location, body.description, body.rsvpUrl, body.ticketId, body.mapsUrl);
      return jsonResponse({ success: sent });
    }
    return jsonResponse({ error: 'Unknown action: ' + action });
  } catch(err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

function syncMembers(members) {
  if (!members||!members.length) return jsonResponse({success:true,skipped:true});
  const ss=getSpreadsheet(); let sheet=ss.getSheetByName('أعضاء CRM')||ss.insertSheet('أعضاء CRM');
  sheet.clearContents();
  const h=['الاسم','البريد','الجوال','الشركة','الوظيفة','الدور','نوع العضوية','نقاط النشاط','الاهتمامات','تاريخ الانضمام'];
  sheet.appendRow(h); sheet.setFrozenRows(1);
  try{sheet.getRange(1,1,1,h.length).setBackground('#0a0a0a').setFontColor('#D4AF37').setFontWeight('bold');}catch(e){}
  members.forEach(function(m){sheet.appendRow([m.name||'',m.email||'',m.phone||'',m.company||'',m.position||'',m.role==='Investor'?'مستثمر':'رائد أعمال',m.memberType||'مستمع',m.engagementScore||50,Array.isArray(m.interests)?m.interests.join(', '):(m.interests||''),m.addedDate||m.createdAt||'']);});
  return jsonResponse({success:true,count:members.length});
}

function syncEvents(events) {
  if (!events||!events.length) return jsonResponse({success:true,skipped:true});
  const ss=getSpreadsheet(); let sheet=ss.getSheetByName('الفعاليات')||ss.insertSheet('الفعاليات');
  sheet.clearContents();
  const h=['العنوان','التاريخ','المكان','النوع','الحالة','الوصف'];
  sheet.appendRow(h); sheet.setFrozenRows(1);
  try{sheet.getRange(1,1,1,h.length).setBackground('#0a0a0a').setFontColor('#D4AF37').setFontWeight('bold');}catch(e){}
  events.forEach(function(ev){sheet.appendRow([ev.title||'',ev.date||'',ev.location||'',ev.type||'ديوانية',ev.status==='completed'?'مكتملة':'مقررة',ev.description||'']);});
  return jsonResponse({success:true,count:events.length});
}

function syncExtras(data) {
  const ss=getSpreadsheet();
  if (data.invitations&&data.invitations.length>0) {
    let sheet=ss.getSheetByName('الدعوات')||ss.insertSheet('الدعوات');
    sheet.clearContents();
    const h=['الفعالية','اسم العضو','البريد','الجوال','RSVP','حضر','التقييم','تاريخ الدعوة'];
    sheet.appendRow(h); sheet.setFrozenRows(1);
    data.invitations.forEach(function(inv){var rsvp=inv.rsvpStatus==='Confirmed'?'مؤكد':inv.rsvpStatus==='Declined'?'معتذر':'معلق';sheet.appendRow([inv.eventTitle||'',inv.memberName||'',inv.memberEmail||'',inv.memberPhone||'',rsvp,inv.attended==='Yes'?'حضر':'',inv.rating||'',inv.sentAt||'']);});
  }
  if (data.interestRegistrations&&data.interestRegistrations.length>0) {
    let sheet=ss.getSheetByName('طلبات الانضمام')||ss.insertSheet('طلبات الانضمام');
    sheet.clearContents();
    sheet.appendRow(['الاسم','البريد','الجوال','الشركة','الدور','الرسالة','التاريخ','الحالة']); sheet.setFrozenRows(1);
    data.interestRegistrations.forEach(function(r){var name=r.fullName||r.name||((r.firstName||'')+' '+(r.lastName||'')).trim();sheet.appendRow([name,r.email||'',r.phone||'',r.entity||r.startupName||r.company||'',r.role||r.interestType||'',r.bio||r.message||'',r.submittedAt||r.date||'',r.status||'pending']);});
  }
  var meta=ss.getSheetByName('آخر مزامنة')||ss.insertSheet('آخر مزامنة');
  meta.clearContents();
  meta.appendRow(['آخر مزامنة',new Date().toLocaleString('ar-SA')]);
  meta.appendRow(['دعوات',data.invitations?data.invitations.length:0]);
  return jsonResponse({success:true});
}

function syncToSheet(data) {
  const ss=getSpreadsheet();
  function upsertSheet(name,headers){let s=ss.getSheetByName(name)||ss.insertSheet(name);s.clearContents();s.appendRow(headers);s.setFrozenRows(1);try{s.getRange(1,1,1,headers.length).setBackground('#0a0a0a').setFontColor('#D4AF37').setFontWeight('bold');}catch(e){}return s;}
  if(data.members&&data.members.length>0){const s=upsertSheet('أعضاء CRM',['الاسم','البريد','الجوال','الشركة','الوظيفة','الدور','نوع العضوية','نقاط النشاط','الاهتمامات','تاريخ الانضمام']);data.members.forEach(function(m){s.appendRow([m.name||'',m.email||'',m.phone||'',m.company||'',m.position||'',m.role==='Investor'?'مستثمر':'رائد أعمال',m.memberType||'مستمع',m.engagementScore||50,Array.isArray(m.interests)?m.interests.join(', '):(m.interests||''),m.addedDate||m.createdAt||'']);});}
  if(data.events&&data.events.length>0){const s=upsertSheet('الفعاليات',['العنوان','التاريخ','المكان','النوع','الحالة','الوصف']);data.events.forEach(function(ev){s.appendRow([ev.title||'',ev.date||'',ev.location||'',ev.type||'ديوانية',ev.status==='completed'?'مكتملة':'مقررة',ev.description||'']);});}
  if(data.invitations&&data.invitations.length>0){const s=upsertSheet('الدعوات',['الفعالية','اسم العضو','البريد','الجوال','RSVP','حضر','التقييم','تاريخ الدعوة']);data.invitations.forEach(function(inv){var rsvp=inv.rsvpStatus==='Confirmed'?'مؤكد':inv.rsvpStatus==='Declined'?'معتذر':'معلق';s.appendRow([inv.eventTitle||'',inv.memberName||'',inv.memberEmail||'',inv.memberPhone||'',rsvp,inv.attended==='Yes'?'حضر':'',inv.rating||'',inv.sentAt||'']);});}
  var meta=ss.getSheetByName('آخر مزامنة')||ss.insertSheet('آخر مزامنة');
  meta.clearContents();meta.appendRow(['آخر مزامنة',new Date().toLocaleString('ar-SA')]);meta.appendRow(['أعضاء',data.members?data.members.length:0]);meta.appendRow(['فعاليات',data.events?data.events.length:0]);meta.appendRow(['دعوات',data.invitations?data.invitations.length:0]);
  return jsonResponse({success:true,syncedAt:new Date().toISOString()});
}

function handleGetDataRaw() {
  const ss=getSpreadsheet();
  return {
    members: getSheetData(ss.getSheetByName("Members")),
    events: getSheetData(ss.getSheetByName("Events")),
    invitations: getSheetData(ss.getSheetByName("Invitations")),
    newsletterSubscribers: getSheetData(ss.getSheetByName("Newsletter")),
    contactMessages: getSheetData(ss.getSheetByName("ContactMessages")),
    interestRegistrations: getSheetData(ss.getSheetByName("InterestRegistrations"))
  };
}

function handleAddMember(member) {
  const sheet=getSpreadsheet().getSheetByName("Members");
  const id='m_'+Math.random().toString(36).substr(2,9), now=new Date().toISOString();
  sheet.appendRow([id,member.name||'',member.email||'',member.phone||'',member.role||'Entrepreneur',member.company||'',member.industry||'',member.ticketSize||'',member.stage||'',member.linkedin||'',member.bio||'',member.engagementScore||50,member.memberType||'مستمع',member.participationType||'مستمع',now]);
  return jsonResponse({success:true,member:{...member,id,createdAt:now}});
}

function handleAddEvent(event) {
  const sheet=getSpreadsheet().getSheetByName("Events");
  const id='e_'+Math.random().toString(36).substr(2,9), now=new Date().toISOString();
  sheet.appendRow([id,event.title||'',event.date||'',event.location||'',event.description||'',now]);
  return jsonResponse({success:true,event:{...event,id}});
}

function handleSendInvitations(postData) {
  const results=[];
  if (postData.invitations&&Array.isArray(postData.invitations)) {
    const eventTitle=postData.eventTitle||'', eventDate=postData.eventDate||'', eventLocation=postData.eventLocation||'', eventDescription=postData.eventDescription||'';
    postData.invitations.forEach(function(inv){
      if(!inv.memberEmail) return;
      const sent=sendHtmlEmail(inv.memberName,inv.memberEmail,eventTitle,eventDate,eventLocation,eventDescription,inv.rsvpLink);
      results.push({email:inv.memberEmail,sent});
    });
    return jsonResponse({success:true,results});
  }
  // Legacy fallback
  const eventId=postData.eventId, memberIds=postData.memberIds||[];
  const ss=getSpreadsheet(), members=getSheetData(ss.getSheetByName("Members")), events=getSheetData(ss.getSheetByName("Events")), invSheet=ss.getSheetByName("Invitations");
  const event=events.find(function(e){return e.id===eventId;});
  if(!event) return jsonResponse({error:'Event not found'});
  const webAppUrl=ScriptApp.getService().getUrl(), now=new Date().toISOString();
  memberIds.forEach(function(memberId){
    const member=members.find(function(m){return m.id===memberId;});
    if(!member) return;
    let invRowIdx=findRowIndex(invSheet,1,eventId,2,memberId), token='', invId='';
    if(invRowIdx===-1){invId='inv_'+Math.random().toString(36).substr(2,9);token=Math.random().toString(36).substr(2,15)+Math.random().toString(36).substr(2,15);invSheet.appendRow([invId,eventId,memberId,'Pending','Pending',token,now,now]);}
    else{const rv=invSheet.getRange(invRowIdx,1,1,8).getValues()[0];invId=rv[0];token=rv[5];}
    const rsvpUrl=webAppUrl+'?token='+token;
    const sent=sendHtmlEmail(member.name,member.email,event.title,event.date,event.location,event.description,rsvpUrl);
    results.push({email:member.email,sent});
  });
  return jsonResponse({success:true,results});
}

function handleUpdateInvitation(invitationId, updates) {
  const sheet=getSpreadsheet().getSheetByName("Invitations"), data=sheet.getDataRange().getValues(), headers=data[0];
  for(let i=1;i<data.length;i++){
    if(data[i][0]===invitationId){
      const rowNum=i+1;
      if(updates.rsvpStatus!==undefined) sheet.getRange(rowNum,headers.indexOf("RSVP Status")+1).setValue(updates.rsvpStatus);
      if(updates.attended!==undefined) sheet.getRange(rowNum,headers.indexOf("Attended")+1).setValue(updates.attended);
      sheet.getRange(rowNum,headers.indexOf("Updated At")+1).setValue(new Date().toISOString());
      recalculateMemberEngagementScore(data[i][2]);
      return jsonResponse({success:true});
    }
  }
  return jsonResponse({error:'Invitation not found'});
}

function handleFormSubmitWebhook(postData) {
  const ss=getSpreadsheet(), sheet=ss.getSheetByName("Members"), members=getSheetData(sheet), email=postData.email;
  if(!email) return jsonResponse({error:'Email required'});
  const existing=members.find(m=>m.email.toLowerCase()===email.toLowerCase());
  if(!existing){
    const id='m_'+Math.random().toString(36).substr(2,9);
    sheet.appendRow([id,postData.name||'',email,postData.phone||'',postData.role||'Entrepreneur',postData.company||'',postData.industry||'',postData.ticketSize||'',postData.stage||'',postData.linkedin||'',postData.bio||'',50,new Date().toISOString()]);
    sendWelcomeEmail(postData.name,email);
    return jsonResponse({success:true,action:'created',memberId:id});
  }
  return jsonResponse({success:true,action:'existing'});
}

function renderRsvpPage(token, immediateStatus) {
  const ss=getSpreadsheet(), invSheet=ss.getSheetByName("Invitations");
  const invData=invSheet.getDataRange().getValues(), invHeaders=invData[0];
  let invRowIdx=-1, invitation=null;
  for(let i=1;i<invData.length;i++){if(invData[i][5]===token){invRowIdx=i+1;invitation={};invHeaders.forEach((h,idx)=>{invitation[toCamelCase(h)]=invData[i][idx];});break;}}
  if(!invitation) return HtmlService.createHtmlOutput('<h2 style="text-align:center;font-family:sans-serif;margin-top:100px;">دعوة غير صالحة</h2>');
  const member=getSheetData(ss.getSheetByName("Members")).find(m=>m.id===invitation.memberId);
  const event=getSheetData(ss.getSheetByName("Events")).find(e=>e.id===invitation.eventId);
  if(immediateStatus&&(immediateStatus==='Confirmed'||immediateStatus==='Declined')){
    invSheet.getRange(invRowIdx,invHeaders.indexOf("RSVP Status")+1).setValue(immediateStatus);
    invSheet.getRange(invRowIdx,invHeaders.indexOf("Updated At")+1).setValue(new Date().toISOString());
    invitation.rsvpStatus=immediateStatus;
    recalculateMemberEngagementScore(invitation.memberId);
  }
  const html=`<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>تأكيد الحضور</title>
  <style>body{font-family:'Segoe UI',sans-serif;background:#0a0a0a;color:#ddd;margin:0;display:flex;justify-content:center;align-items:center;min-height:100vh;padding:20px;box-sizing:border-box;}
  .card{background:#111;border:1px solid #D4AF37;border-radius:16px;padding:40px;max-width:480px;width:100%;text-align:center;}
  .logo{color:#D4AF37;font-size:22px;font-weight:700;margin-bottom:24px;}
  h2{color:#fff;margin-top:0;}
  .event-box{background:#0a0a0a;border:1px solid #222;border-right:3px solid #D4AF37;border-radius:8px;padding:16px 20px;margin:20px 0;text-align:right;}
  .event-box p{margin:6px 0;font-size:14px;}
  .label{color:#888;}
  .btns{display:flex;gap:14px;margin-top:28px;}
  .btn{flex:1;padding:14px;border-radius:8px;font-size:15px;font-weight:700;cursor:pointer;border:none;}
  .btn-c{background:#D4AF37;color:#0a0a0a;}
  .btn-d{background:#1a1a1a;color:#888;border:1px solid #333;}
  .success{display:none;padding:14px;background:rgba(212,175,55,0.1);border:1px solid #D4AF37;color:#D4AF37;border-radius:8px;margin-top:16px;}
  .badge{display:inline-block;padding:5px 14px;border-radius:20px;font-size:13px;font-weight:700;margin-bottom:16px;background:#1a1a1a;color:#D4AF37;border:1px solid #D4AF37;}</style>
  </head><body><div class="card">
  <div class="logo">Droob Angels</div>
  <h2>تأكيد الحضور</h2>
  <p style="color:#888;">مرحباً <strong style="color:#fff;">${member?member.name:''}</strong></p>
  <div class="event-box">
    <p><span class="label">الفعالية: </span>${event?event.title:''}</p>
    <p><span class="label">التاريخ: </span>${event?event.date:''}</p>
    <p><span class="label">المكان: </span>${event?event.location:''}</p>
  </div>
  <div class="badge" id="badge">الحالة: ${translateStatus(invitation.rsvpStatus)}</div>
  <div class="success" id="success"></div>
  <div class="btns" id="buttons">
    <button class="btn btn-c" onclick="submitRsvp('Confirmed')">تأكيد الحضور</button>
    <button class="btn btn-d" onclick="submitRsvp('Declined')">اعتذار</button>
  </div></div>
  <script>
  function submitRsvp(s){
    document.getElementById('buttons').style.opacity='0.5';document.getElementById('buttons').style.pointerEvents='none';
    fetch(window.location.href+'&status='+s,{method:'GET'}).then(()=>{
      document.getElementById('badge').textContent='الحالة: '+(s==='Confirmed'?'مؤكد':'معتذر');
      var el=document.getElementById('success');el.style.display='block';
      el.textContent=s==='Confirmed'?'تم تأكيد حضورك — نتطلع لرؤيتك!':'تم تسجيل اعتذارك — نتمنى رؤيتك قريباً.';
      document.getElementById('buttons').style.display='none';
    });
  }
  if("${immediateStatus||''}"!==''){document.getElementById('success').style.display='block';document.getElementById('buttons').style.display='none';}
  </script></body></html>`;
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleAddNewsletterSubscriber(subscriber) {
  const sheet=getSpreadsheet().getSheetByName("Newsletter");
  if(!sheet){setup();return handleAddNewsletterSubscriber(subscriber);}
  const data=sheet.getDataRange().getValues();
  if(data.slice(1).find(r=>r[2].toString().toLowerCase()===(subscriber.email||'').toLowerCase())) return jsonResponse({success:false,error:'already_subscribed'});
  const id=subscriber.id||('ns_'+Math.random().toString(36).substr(2,9)), now=subscriber.subscribedAt||new Date().toISOString();
  sheet.appendRow([id,subscriber.name||'',subscriber.email||'',now]);
  try{MailApp.sendEmail(subscriber.email,'تم اشتراكك في نشرة Droob Angels','',{htmlBody:getEmailHeader()+'<p style="color:#ddd;">شكراً <strong>'+subscriber.name+'</strong>، تم اشتراكك في نشرة Droob Angels.</p>'+getEmailFooter(),replyTo:SENDER_EMAIL,name:SENDER_NAME});}catch(e){}
  return jsonResponse({success:true,id});
}

function handleAddContactMessage(msg) {
  const sheet=getSpreadsheet().getSheetByName("ContactMessages");
  if(!sheet){setup();return handleAddContactMessage(msg);}
  const id=msg.id||('cm_'+Math.random().toString(36).substr(2,9)), now=msg.createdAt||new Date().toISOString();
  sheet.appendRow([id,msg.name||'',msg.email||'',msg.phone||'',msg.message||'',now,'false']);
  try{const n=Session.getActiveUser().getEmail();if(n)MailApp.sendEmail(n,'رسالة جديدة من '+msg.name,'الاسم: '+msg.name+'\nالبريد: '+msg.email+'\nالرسالة:\n'+msg.message,{name:SENDER_NAME,replyTo:SENDER_EMAIL});}catch(e){}
  return jsonResponse({success:true,id});
}

function handleAddInterestRegistration(reg) {
  const sheet=getSpreadsheet().getSheetByName("InterestRegistrations");
  if(!sheet){setup();return handleAddInterestRegistration(reg);}
  const id=reg.id||('ir_'+Math.random().toString(36).substr(2,9)), now=reg.createdAt||new Date().toISOString();
  sheet.appendRow([id,reg.name||'',reg.email||'',reg.phone||'',reg.company||'',reg.linkedin||'',reg.interestType||'',reg.message||'',now]);
  try{sendInterestConfirmationEmail(reg.name,reg.email,reg.interestType||'entrepreneur');}catch(e){}
  return jsonResponse({success:true,id});
}

function handleEventInvite(data) {
  const subject='دعوة: '+(data.eventName||'فعالية Droob Angels');
  const html=getEmailHeader()+'<p style="font-size:16px;color:#fff;">أهلاً <strong>'+data.memberName+'</strong>،</p><p style="color:#aaa;line-height:1.8;">يسرنا دعوتك لحضور فعالية Droob Angels القادمة.</p>'+getInfoCard('<p style="color:#fff;font-size:16px;font-weight:600;margin:0 0 12px;">'+data.eventName+'</p><p style="color:#aaa;font-size:13px;margin:6px 0;">التاريخ: '+data.eventDate+'</p><p style="color:#aaa;font-size:13px;margin:6px 0;">المكان: '+data.eventLocation+'</p>')+'<div style="text-align:center;margin-top:20px;">'+getBtnPrimary(data.rsvpLink+'&status=Confirmed','تأكيد الحضور')+getBtnGhost(data.rsvpLink+'&status=Declined','اعتذار')+'</div>'+getEmailFooter();
  try{MailApp.sendEmail(data.to,subject,'',{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});return jsonResponse({success:true});}
  catch(e){return jsonResponse({error:e.toString()});}
}

function handleEventInviteExternal(data) {
  const subject='أنت مدعو — '+(data.eventName||'فعالية Droob Angels');
  const html=getEmailHeader()+'<p style="font-size:16px;color:#fff;">أهلاً <strong>'+data.guestName+'</strong>،</p>'+getInfoCard('<p style="color:#fff;font-size:15px;font-weight:600;margin:0 0 12px;">'+data.eventName+'</p><p style="color:#aaa;font-size:13px;margin:6px 0;">التاريخ: '+data.eventDate+'</p><p style="color:#aaa;font-size:13px;margin:6px 0;">المكان: '+data.eventLocation+'</p>')+'<div style="text-align:center;margin:20px 0;">'+getBtnPrimary(data.registrationLink,'سجّل مكانك الآن')+'</div>'+getEmailFooter();
  try{MailApp.sendEmail(data.to,subject,'',{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});return jsonResponse({success:true});}
  catch(e){return jsonResponse({error:e.toString()});}
}

function handleAddExternalGuest(eventId,guest) {
  const ss=getSpreadsheet();let sheet=ss.getSheetByName("EventGuests");
  if(!sheet){sheet=ss.insertSheet("EventGuests");sheet.appendRow(["eventId","guestName","guestEmail","guestPhone","addedBy","addedAt","isExternal","rsvpStatus","checkedIn"]);}
  sheet.appendRow([eventId,guest.name||'',guest.email||'',guest.phone||'',guest.addedBy||'',new Date().toISOString(),true,'Pending',false]);
  return jsonResponse({success:true});
}

function recalculateMemberEngagementScore(memberId) {
  const ss=getSpreadsheet(), invSheet=ss.getSheetByName("Invitations"), membersSheet=ss.getSheetByName("Members");
  const invs=getSheetData(invSheet).filter(i=>i.memberId===memberId);
  if(!invs.length) return;
  const total=invs.length, attended=invs.filter(i=>i.attended==='Yes').length, confirmed=invs.filter(i=>i.rsvpStatus==='Confirmed').length, declined=invs.filter(i=>i.rsvpStatus==='Declined').length;
  const missed=invs.filter(i=>i.rsvpStatus==='Confirmed'&&i.attended==='No').length;
  let score=50+Math.round(attended/total*40)+Math.round((confirmed+declined)/total*15)-(missed*15);
  score=Math.max(0,Math.min(100,score));
  const mData=membersSheet.getDataRange().getValues();
  for(let i=1;i<mData.length;i++){if(mData[i][0]===memberId){membersSheet.getRange(i+1,mData[0].indexOf("Engagement Score")+1).setValue(score);break;}}
}

function handleNewsletterSend(data) {
  const subject=data.subject||'نشرة Droob Angels', htmlContent=data.body||'', recipients=data.recipients||[];
  let sent=0,errors=0;
  recipients.forEach(r=>{
    try{sendNewsletterEmail(r.email,subject,htmlContent);sent++;Utilities.sleep(100);}
    catch(e){Logger.log('Newsletter error for '+r.email+': '+e);errors++;}
  });
  return jsonResponse({success:true,sent,errors});
}

function handleRsvpUpdate(guestId,eventId,status,data) {
  const ss=getSpreadsheet(), invSheet=ss.getSheetByName("Invitations"), members=getSheetData(ss.getSheetByName("Members")), events=getSheetData(ss.getSheetByName("Events"));
  const member=members.find(m=>m.id===guestId), event=events.find(e=>e.id===eventId);
  if(!member||!event) return jsonResponse({error:'Not found'});
  if(invSheet){const invData=invSheet.getDataRange().getValues(),h=invData[0];for(let i=1;i<invData.length;i++){if(invData[i][h.indexOf('Member ID')]===guestId&&invData[i][h.indexOf('Event ID')]===eventId){invSheet.getRange(i+1,h.indexOf("RSVP Status")+1).setValue(status);invSheet.getRange(i+1,h.indexOf("Updated At")+1).setValue(new Date().toISOString());recalculateMemberEngagementScore(guestId);break;}}}
  if(status==='Confirmed'){const ticketId=new Date().getFullYear()+String(new Date().getMonth()+1).padStart(2,'0')+'-'+(guestId||'').slice(-4).toUpperCase();sendRsvpConfirmationEmail(member.name,member.email,event.title,event.date,event.location,ticketId,'',event.location?'https://maps.google.com/?q='+encodeURIComponent(event.location):'');}
  else sendRsvpDeclineEmail(member.name,member.email,event.title);
  return jsonResponse({success:true,status,guestName:member.name,eventName:event.title});
}

function handlePortalRegistration(data) {
  const ss=getSpreadsheet();let sheet=ss.getSheetByName("PortalRegistrations")||ss.insertSheet("PortalRegistrations");
  if(sheet.getLastRow()===0) sheet.appendRow(["timestamp","role","firstName","lastName","email","phone","linkedin","region","entity","dealCount","ticketSize","investorType","sectors","startupName","stage","askAmount","startupSector","needs","bio","referral","goals","participationType","showInDirectory"]);
  sheet.appendRow([new Date().toISOString(),data.role||'',data.firstName||'',data.lastName||'',data.email||'',data.phone||'',data.linkedin||'',data.region||'',data.entity||'',data.dealCount||'',data.ticketSize||'',data.investorType||'',Array.isArray(data.sectors)?data.sectors.join(', '):(data.sectors||''),data.startupName||'',data.stage||'',data.askAmount||'',data.startupSector||'',Array.isArray(data.needs)?data.needs.join(', '):(data.needs||''),data.bio||'',data.referral||'',Array.isArray(data.goals)?data.goals.join(', '):(data.goals||''),data.participationType||'',data.showInDirectory!==false?'yes':'no']);
  sendInterestConfirmationEmail(((data.firstName||'')+' '+(data.lastName||'')).trim(),data.email,data.role||'entrepreneur');
  return jsonResponse({success:true});
}

// ── Email Templates ──────────────────────────────────────────────

function sendHtmlEmail(name,email,title,date,location,description,rsvpUrl,ticketId,mapsUrl) {
  try {
    // Admin notification mode
    if (!date && !location && !rsvpUrl && description) {
      const html=getEmailHeader()+
        '<p style="font-size:16px;color:#fff;margin-bottom:16px;">'+description+'</p>'+
        getEmailFooter();
      MailApp.sendEmail(email,title+' — Droob Angels','',{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});
      return true;
    }
    const subject='دعوة | '+title+' — مجتمع دروب';
    const plain='مرحباً '+name+'،\n\nنذكّركم بـ'+title+(date?'\n\nالتاريخ: '+date:'')+(location?'\nالموقع: '+location:'')+(mapsUrl?'\nرابط الموقع: '+mapsUrl:'')+(rsvpUrl?'\n\nتأكيد الحضور: '+rsvpUrl:'')+'\n\nيسعدنا حضوركم\nمجتمع دروب';
    const confirmUrl = rsvpUrl ? rsvpUrl + (rsvpUrl.includes('?') ? '&' : '?') + 'status=Confirmed' : '';
    const declineUrl  = rsvpUrl ? rsvpUrl + (rsvpUrl.includes('?') ? '&' : '?') + 'status=Declined'  : '';
    const mapsRow = mapsUrl ? '<p style="color:#aaa;font-size:13px;margin:6px 0;">رابط الموقع: <a href="'+mapsUrl+'" style="color:#D4AF37;">'+mapsUrl+'</a></p>' : '';
    const html=getEmailHeader()+
      '<p style="font-size:16px;color:#fff;margin-bottom:20px;">مرحباً <strong>'+name+'</strong>،</p>'+
      '<p style="color:#aaa;font-size:14px;line-height:1.8;margin-bottom:20px;">نذكّركم بـ <strong style="color:#fff;">'+title+'</strong> اليوم بإذن الله.</p>'+
      getInfoCard(
        (date?'<p style="color:#aaa;font-size:13px;margin:6px 0;"><strong style="color:#D4AF37;">الوقت:</strong> '+date+'</p>':'')+
        (location?'<p style="color:#aaa;font-size:13px;margin:6px 0;"><strong style="color:#D4AF37;">الموقع:</strong> '+location+'</p>':'')+
        mapsRow+
        (description?getDivider()+'<p style="color:#aaa;font-size:13px;line-height:1.7;">'+description+'</p>':'')
      )+
      (rsvpUrl ?
        '<p style="color:#aaa;font-size:13px;text-align:center;margin:24px 0 12px;">يرجى تأكيد حضوركم:</p>'+
        '<div style="text-align:center;">'+
        getBtnPrimary(confirmUrl,'تأكيد الحضور')+
        getBtnGhost(declineUrl,'اعتذار')+
        '</div>' : '')+
      '<p style="color:#888;font-size:13px;text-align:center;margin-top:24px;">يسعدنا حضوركم، ونلتقيكم على خير.</p>'+
      getEmailFooter();
    MailApp.sendEmail(email,subject,plain,{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});
    return true;
  } catch(err){Logger.log('Email error: '+err);return false;}
}

function sendWelcomeEmail(name,email) {
  try {
    const html=getEmailHeader()+
      '<p style="font-size:16px;color:#fff;">مرحباً <strong>'+name+'</strong>،</p>'+
      '<p style="color:#aaa;font-size:14px;line-height:1.8;">يسعدنا انضمامك إلى مجتمع دروب — المجتمع الذي يجمع نخبة المستثمرين الملائكيين ورواد الأعمال في المملكة.</p>'+
      getDivider()+
      getInfoCard('<p style="color:#D4AF37;font-size:13px;font-weight:600;margin:0 0 10px;">ماذا يمكنك توقعه:</p>'+
        '<p style="color:#aaa;font-size:13px;margin:6px 0;">وصول حصري للفعاليات المغلقة والديوانيات</p>'+
        '<p style="color:#aaa;font-size:13px;margin:6px 0;">توافق ذكي مع المستثمرين أو رواد الأعمال المناسبين</p>'+
        '<p style="color:#aaa;font-size:13px;margin:6px 0;">مسار نمو داخل المجتمع يكافئ مشاركتك الحقيقية</p>')+
      '<div style="text-align:center;margin-top:20px;">'+getBtnGhost('mailto:'+SENDER_EMAIL,'تواصل مع الفريق')+'</div>'+
      getEmailFooter();
    MailApp.sendEmail(email,'مرحباً بك في مجتمع دروب','مرحباً '+name+'، يسعدنا انضمامك.',{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});
  } catch(err){Logger.log('Welcome email error: '+err);}
}

// ── رسالة تذكير بالفعالية ──
function sendReminderEmail(name,email,title,date,location,mapsUrl) {
  try {
    const subject='تذكير | '+title+' — مجتمع دروب';
    const plain='مرحباً '+name+'،\n\nنذكّركم بـ'+title+(date?'\n\nالوقت: '+date:'')+(location?'\nالموقع: '+location:'')+(mapsUrl?'\nرابط الموقع: '+mapsUrl:'')+'\n\nيسعدنا حضوركم\nمجتمع دروب';
    const mapsRow=mapsUrl?'<p style="color:#aaa;font-size:13px;margin:6px 0;"><strong style="color:#D4AF37;">رابط الموقع:</strong> <a href="'+mapsUrl+'" style="color:#D4AF37;">'+mapsUrl+'</a></p>':'';
    const html=getEmailHeader()+
      '<p style="font-size:16px;color:#fff;margin-bottom:20px;">مرحباً <strong>'+name+'</strong>،</p>'+
      '<p style="color:#aaa;font-size:14px;line-height:1.8;margin-bottom:20px;">نذكّركم بـ <strong style="color:#fff;">'+title+'</strong> اليوم بإذن الله.</p>'+
      getInfoCard(
        (date?'<p style="color:#aaa;font-size:13px;margin:6px 0;"><strong style="color:#D4AF37;">الوقت:</strong> '+date+'</p>':'')+
        (location?'<p style="color:#aaa;font-size:13px;margin:6px 0;"><strong style="color:#D4AF37;">الموقع:</strong> '+location+'</p>':'')+
        mapsRow)+
      '<p style="color:#888;font-size:13px;text-align:center;margin-top:24px;">يسعدنا حضوركم، ونلتقيكم على خير.</p>'+
      getEmailFooter();
    MailApp.sendEmail(email,subject,plain,{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});
    return true;
  } catch(err){Logger.log('Reminder email error: '+err);return false;}
}

// ── رسالة ما بعد الفعالية ──
function sendPostEventEmail(name,email,title) {
  try {
    const subject='شكراً لحضوركم | '+title+' — مجتمع دروب';
    const plain='مرحباً '+name+'،\n\nشكراً لحضوركم '+title+'. نتطلع لرؤيتكم في الفعالية القادمة.\n\nمجتمع دروب';
    const html=getEmailHeader()+
      '<p style="font-size:16px;color:#fff;margin-bottom:20px;">مرحباً <strong>'+name+'</strong>،</p>'+
      '<p style="color:#aaa;font-size:14px;line-height:1.8;">شكراً لحضوركم فعالية <strong style="color:#D4AF37;">'+title+'</strong>.</p>'+
      '<p style="color:#aaa;font-size:14px;line-height:1.8;margin-top:12px;">نتطلع لرؤيتكم في الفعالية القادمة.</p>'+
      getDivider()+
      '<div style="text-align:center;margin-top:16px;">'+getBtnGhost('https://doroobangels.com/diwaniya','تصفح الفعاليات القادمة')+'</div>'+
      getEmailFooter();
    MailApp.sendEmail(email,subject,plain,{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});
    return true;
  } catch(err){Logger.log('Post-event email error: '+err);return false;}
}

function sendRsvpConfirmationEmail(name,email,eventTitle,eventDate,eventLocation,ticketId,calendarLink,mapsLink) {
  try {
    const html=getEmailHeader()+
      '<div style="text-align:center;margin-bottom:20px;"><div style="width:52px;height:52px;background:#1a1a0a;border:2px solid #D4AF37;border-radius:50%;display:inline-block;line-height:52px;font-size:20px;color:#D4AF37;">✓</div></div>'+
      '<p style="font-size:16px;color:#fff;text-align:center;font-weight:600;">تم تأكيد حضورك!</p>'+
      '<p style="color:#aaa;font-size:14px;text-align:center;margin-bottom:20px;">نتطلع لرؤيتك يا <strong style="color:#D4AF37;">'+name+'</strong></p>'+
      getInfoCard('<p style="color:#fff;font-size:15px;font-weight:600;margin:0 0 12px;">'+eventTitle+'</p>'+
        '<p style="color:#aaa;font-size:13px;margin:6px 0;">التاريخ: '+eventDate+'</p>'+
        '<p style="color:#aaa;font-size:13px;margin:6px 0;">المكان: '+eventLocation+'</p>'+
        '<p style="color:#aaa;font-size:13px;margin:6px 0;">رقم التذكرة: <strong style="color:#D4AF37;font-family:monospace;">#DRB-'+ticketId+'</strong></p>')+
      '<div style="text-align:center;margin-top:20px;">'+(mapsLink?getBtnGhost(mapsLink,'عرض الموقع'):'')+'</div>'+
      getEmailFooter();
    MailApp.sendEmail(email,'تم تأكيد حضورك — '+eventTitle,'تم تأكيد حضورك يا '+name+'. تذكرتك: #DRB-'+ticketId,{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});
    return true;
  } catch(err){Logger.log('RSVP confirm error: '+err);return false;}
}

function sendRsvpDeclineEmail(name,email,eventTitle) {
  try {
    const html=getEmailHeader()+
      '<p style="font-size:16px;color:#fff;">شكراً لإخبارنا <strong>'+name+'</strong></p>'+
      '<p style="color:#aaa;font-size:14px;line-height:1.8;">تم تسجيل اعتذارك عن حضور <strong style="color:#D4AF37;">'+eventTitle+'</strong>. نتمنى أن نراك في الفعالية القادمة.</p>'+
      '<div style="text-align:center;margin-top:20px;">'+getBtnGhost('mailto:'+SENDER_EMAIL+'','تواصل مع الفريق')+'</div>'+
      getEmailFooter();
    MailApp.sendEmail(email,'تم تسجيل اعتذارك','شكراً '+name,{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});
    return true;
  } catch(err){Logger.log('RSVP decline error: '+err);return false;}
}

function sendInterestConfirmationEmail(name,email,role) {
  try {
    const roleLabels={investor:'مستثمر ملاك',entrepreneur:'رائد أعمال',ecosystem:'مهتم بالمنظومة',cofounder:'شريك مؤسس',guest:'ضيف'};
    const roleLabel=roleLabels[role]||roleLabels[role==='مستثمر'?'investor':'entrepreneur']||'عضو';
    const html=getEmailHeader()+
      '<p style="font-size:16px;color:#fff;">أهلاً <strong>'+name+'</strong>،</p>'+
      '<p style="color:#aaa;font-size:14px;line-height:1.8;">وصلنا طلب انضمامك لمجتمع Droob Angels بصفتك <strong style="color:#D4AF37;">'+roleLabel+'</strong>. شكراً على اهتمامك.</p>'+
      getDivider()+
      getInfoCard('<p style="color:#D4AF37;font-size:13px;font-weight:600;margin:0 0 10px;">ماذا يحدث الآن؟</p>'+
        '<p style="color:#aaa;font-size:13px;margin:6px 0;">١ — سيراجع الفريق طلبك خلال 48 ساعة</p>'+
        '<p style="color:#aaa;font-size:13px;margin:6px 0;">٢ — سيتواصل معك أحد أعضاء الفريق مباشرة</p>'+
        '<p style="color:#aaa;font-size:13px;margin:6px 0;">٣ — ستصلك دعوة لأول فعالية مناسبة</p>')+
      '<div style="text-align:center;margin-top:20px;">'+getBtnGhost('mailto:'+SENDER_EMAIL+'','تواصل معنا')+'</div>'+
      getEmailFooter();
    MailApp.sendEmail(email,'وصل طلبك — Droob Angels','أهلاً '+name+', وصل طلبك.',{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});
    return true;
  } catch(err){Logger.log('Interest email error: '+err);return false;}
}

function sendNewsletterEmail(subscriberEmail,subject,contentHtml) {
  try {
    const html=getEmailHeader()+
      '<div style="display:inline-block;background:#1a1600;border:1px solid #D4AF37;border-radius:6px;padding:4px 12px;font-size:11px;color:#D4AF37;font-weight:600;margin-bottom:16px;">النشرة البريدية</div>'+
      '<h2 style="color:#fff;font-size:18px;margin:0 0 20px;">'+subject+'</h2>'+
      getDivider()+contentHtml+getDivider()+
      '<p style="color:#555;font-size:11px;text-align:center;">وصلتك هذه النشرة لأنك مشترك في قائمة Droob Angels البريدية.</p>'+
      getEmailFooter();
    MailApp.sendEmail(subscriberEmail,subject+' — Droob Angels','',{htmlBody:html,replyTo:SENDER_EMAIL,name:SENDER_NAME});
    return true;
  } catch(err){Logger.log('Newsletter error: '+err);return false;}
}

// ── Firestore emailJobs queue processor ──────────────────────────

function firestoreGetPendingEmailJobs() {
  // Query emailJobs where status == "pending"
  var url = FIRESTORE_BASE + '/emailJobs?key=' + FIREBASE_API_KEY + '&pageSize=20';
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(resp.getContentText());
  if (!json.documents) return [];
  return json.documents
    .map(function(doc) {
      var obj = { _docName: doc.name };
      Object.keys(doc.fields || {}).forEach(function(k) {
        var v = doc.fields[k];
        if (v.stringValue !== undefined) { obj[k] = v.stringValue; return; }
        if (v.integerValue !== undefined) { obj[k] = Number(v.integerValue); return; }
        if (v.booleanValue !== undefined) { obj[k] = v.booleanValue; return; }
        if (v.arrayValue) {
          obj[k] = (v.arrayValue.values || []).map(function(item) {
            if (item.mapValue) {
              var m = {};
              Object.keys(item.mapValue.fields || {}).forEach(function(mk) {
                var mv = item.mapValue.fields[mk];
                m[mk] = mv.stringValue !== undefined ? mv.stringValue
                       : mv.booleanValue !== undefined ? mv.booleanValue
                       : '';
              });
              return m;
            }
            return item.stringValue || '';
          });
          return;
        }
        if (v.mapValue) { obj[k] = v.mapValue; return; }
        obj[k] = '';
      });
      return obj;
    })
    .filter(function(job) { return job.status === 'pending'; });
}

function firestoreUpdateJobStatus(docName, status) {
  // docName from REST API is like "projects/xxx/databases/(default)/documents/..."
  // Need to prepend the base URL
  var fullUrl = docName.startsWith('http')
    ? docName
    : 'https://firestore.googleapis.com/v1/' + docName;
  var url = fullUrl + '?key=' + FIREBASE_API_KEY + '&updateMask.fieldPaths=status';
  var payload = JSON.stringify({
    fields: { status: { stringValue: status } }
  });
  UrlFetchApp.fetch(url, {
    method: 'PATCH',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true
  });
}

function ensureEmailJobsTrigger() {
  var hasTrigger = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'processEmailJobs';
  });
  if (!hasTrigger) {
    ScriptApp.newTrigger('processEmailJobs').timeBased().everyMinutes(1).create();
    Logger.log('processEmailJobs trigger was missing — auto-installed');
  }
}

function processEmailJobs() {
  try {
    // Auto-heal: install trigger if missing
    ensureEmailJobsTrigger();

    var jobs = firestoreGetPendingEmailJobs();
    Logger.log('processEmailJobs: found ' + jobs.length + ' pending jobs');
    jobs.forEach(function(job) {
      try {
        if (job.action === 'sendInvitations') {
          var invitations = job.invitations || [];
          var sent = 0, errors = 0;
          invitations.forEach(function(inv) {
            if (!inv.memberEmail) return;
            Logger.log('processEmailJobs: sending to ' + inv.memberEmail + ' (' + (inv.memberName || '') + ') for job ' + (job.id || job._docName));
            var ok = sendHtmlEmail(
              inv.memberName || '',
              inv.memberEmail,
              job.eventTitle || '',
              job.eventDate || '',
              job.eventLocation || '',
              job.eventDescription || '',
              inv.rsvpLink || ''
            );
            Logger.log('processEmailJobs: ' + inv.memberEmail + ' => ' + (ok ? 'SENT' : 'FAILED'));
            if (ok) sent++; else errors++;
            Utilities.sleep(200);
          });
          Logger.log('Job ' + (job.id || job._docName) + ': sent=' + sent + ' errors=' + errors);
          firestoreUpdateJobStatus(job._docName, errors === 0 ? 'processed' : 'partial');
        } else if (job.action === 'sendHtmlEmail') {
          Logger.log('processEmailJobs: sending single email to ' + job.email);
          var ok = sendHtmlEmail(job.name, job.email, job.title, job.date, job.location, job.description, job.rsvpUrl);
          Logger.log('processEmailJobs: ' + job.email + ' => ' + (ok ? 'SENT' : 'FAILED'));
          firestoreUpdateJobStatus(job._docName, ok ? 'processed' : 'failed');
        } else {
          Logger.log('Unknown job action: ' + job.action);
          firestoreUpdateJobStatus(job._docName, 'unknown_action');
        }
      } catch(jobErr) {
        Logger.log('Job error (' + (job.id || job._docName) + '): ' + jobErr);
        firestoreUpdateJobStatus(job._docName, 'failed');
      }
    });
  } catch(err) {
    Logger.log('processEmailJobs error: ' + err);
  }
}

function installEmailJobsTrigger() {
  // Now delegates to installAllTriggers for consistency
  installAllTriggers();
}

function renderTriggersPage() {
  var triggers = ScriptApp.getProjectTriggers();
  var rows = '';
  if (triggers.length === 0) {
    rows = '<tr><td colspan="4" style="color:#f66;text-align:center;padding:16px;">No triggers installed!</td></tr>';
  } else {
    triggers.forEach(function(t) {
      var evType = t.getEventType();
      var src = t.getTriggerSource();
      rows += '<tr>' +
        '<td>' + t.getHandlerFunction() + '</td>' +
        '<td>' + src + '</td>' +
        '<td>' + evType + '</td>' +
        '<td style="color:#4f4;">✓ Active</td>' +
        '</tr>';
    });
  }
  var installUrl = ScriptApp.getService().getUrl() + '?action=installTrigger';
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>body{font-family:sans-serif;background:#0a0a0a;color:#ddd;padding:30px;}' +
    'h2{color:#D4AF37;}' +
    'table{width:100%;border-collapse:collapse;margin-top:16px;}' +
    'th{background:#1a1a1a;color:#D4AF37;padding:10px 14px;text-align:left;border:1px solid #333;}' +
    'td{padding:10px 14px;border:1px solid #222;}' +
    '.btn{display:inline-block;margin-top:24px;background:#D4AF37;color:#0a0a0a;padding:10px 22px;border-radius:8px;text-decoration:none;font-weight:700;}' +
    '</style></head><body>' +
    '<h2>Installed Triggers</h2>' +
    '<table><thead><tr><th>Function</th><th>Source</th><th>Event Type</th><th>Status</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table>' +
    '<a class="btn" href="' + installUrl + '">Install / Reinstall All Triggers</a>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Utilities ────────────────────────────────────────────────────

function getSheetData(sheet) {
  if(!sheet) return [];
  const values=sheet.getDataRange().getValues();
  if(values.length<=1) return [];
  const headers=values[0];
  return values.slice(1).map(row=>{const item={};headers.forEach((h,i)=>{item[toCamelCase(h)]=row[i];});return item;});
}

function findRowIndex(sheet,colIndex1,val1,colIndex2,val2) {
  const data=sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){if(data[i][colIndex1-1]===val1&&data[i][colIndex2-1]===val2)return i+1;}
  return -1;
}

function toCamelCase(str) {
  return str.toLowerCase().replace(/[^a-zA-Z0-9]+(.)/g,(m,c)=>c.toUpperCase()).replace(/[^a-zA-Z0-9]/g,'');
}

function translateStatus(status) {
  return status==='Confirmed'?'مؤكد':status==='Declined'?'معتذر':'معلق';
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function jsonpResponse(data,callback) {
  return ContentService.createTextOutput((callback||'callback')+'('+JSON.stringify(data)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
