const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');

const LOGS_FILE = path.join(__dirname, 'data', 'email_logs.json');

// Ensure directory exists
if (!fs.existsSync(path.dirname(LOGS_FILE))) {
  fs.mkdirSync(path.dirname(LOGS_FILE), { recursive: true });
}

// Simple helper to append to a JSON file log
function logEmailLocally(emailDetails) {
  try {
    let logs = [];
    if (fs.existsSync(LOGS_FILE)) {
      const data = fs.readFileSync(LOGS_FILE, 'utf8');
      logs = JSON.parse(data || '[]');
    }
    logs.unshift({
      id: 'log_' + Math.random().toString(36).substr(2, 9),
      timestamp: new Date().toISOString(),
      ...emailDetails
    });
    // Keep only last 100 emails logged
    if (logs.length > 100) logs = logs.slice(0, 100);
    fs.writeFileSync(LOGS_FILE, JSON.stringify(logs, null, 2));
    console.log(`[Email Simulator] Email logged to ${LOGS_FILE} for ${emailDetails.to}`);
  } catch (err) {
    console.error('Error logging email locally:', err);
  }
}

// Mailer configuration (loaded from environment variables if present)
function getTransporter() {
  if (process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS) {
    return nodemailer.createTransport({
      host: process.env.SMTP_HOST,
      port: parseInt(process.env.SMTP_PORT) || 587,
      secure: process.env.SMTP_SECURE === 'true', // true for 465, false for other ports
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS
      }
    });
  }
  return null;
}

// Generate beautiful HTML Email Template (Arabic)
function generateInvitationHtml(memberName, eventTitle, eventDate, eventLocation, eventDescription, rsvpLink) {
  return `
  <!DOCTYPE html>
  <html dir="rtl" lang="ar">
  <head>
    <meta charset="UTF-8">
    <title>دعوة حضور: ${eventTitle}</title>
    <style>
      body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background-color: #0d1117;
        color: #c9d1d9;
        margin: 0;
        padding: 0;
        text-align: right;
      }
      .container {
        max-width: 600px;
        margin: 40px auto;
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 8px 24px rgba(0,0,0,0.5);
      }
      .header {
        background: linear-gradient(135deg, #1f6feb 0%, #00d2ff 100%);
        padding: 30px;
        text-align: center;
        border-bottom: 1px solid #30363d;
      }
      .header h1 {
        margin: 0;
        color: #ffffff;
        font-size: 24px;
        font-weight: 700;
      }
      .content {
        padding: 30px;
        line-height: 1.8;
      }
      .welcome {
        font-size: 18px;
        color: #58a6ff;
        margin-bottom: 20px;
      }
      .event-details {
        background: #0d1117;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 20px;
        margin: 20px 0;
      }
      .detail-item {
        margin-bottom: 12px;
        font-size: 15px;
      }
      .detail-title {
        font-weight: bold;
        color: #58a6ff;
      }
      .description {
        color: #8b949e;
        font-size: 14px;
        border-top: 1px dashed #30363d;
        padding-top: 12px;
        margin-top: 12px;
      }
      .cta-section {
        text-align: center;
        margin: 30px 0 10px;
      }
      .btn {
        display: inline-block;
        padding: 12px 30px;
        border-radius: 8px;
        text-decoration: none;
        font-weight: bold;
        font-size: 16px;
        transition: all 0.3s ease;
        margin: 0 10px 10px;
      }
      .btn-confirm {
        background-color: #238636;
        color: #ffffff !important;
      }
      .btn-decline {
        background-color: #21262d;
        color: #c9d1d9 !important;
        border: 1px solid #30363d;
      }
      .footer {
        background: #0d1117;
        padding: 20px;
        text-align: center;
        font-size: 12px;
        color: #8b949e;
        border-top: 1px solid #30363d;
      }
      .footer a {
        color: #58a6ff;
        text-decoration: none;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">
        <h1>مجتمع المستثمرين ورواد الأعمال</h1>
      </div>
      <div class="content">
        <div class="welcome">مرحباً بك، سيد/ة ${memberName} 👋</div>
        <p>يسعدنا ويشرفنا دعوتك لحضور فعاليتنا القادمة والمصممة خصيصاً لتعزيز التواصل والشراكات بين المستثمرين ورواد الأعمال في مجتمعنا.</p>
        
        <div class="event-details">
          <div class="detail-item"><span class="detail-title">الفعالية:</span> ${eventTitle}</div>
          <div class="detail-item"><span class="detail-title">التاريخ:</span> ${eventDate}</div>
          <div class="detail-item"><span class="detail-title">الموقع:</span> ${eventLocation}</div>
          <div class="description">${eventDescription}</div>
        </div>

        <p style="text-align: center; font-weight: bold; margin-bottom: 20px;">يرجى تأكيد حضورك أو اعتذارك بالضغط على أحد الأزرار أدناه:</p>

        <div class="cta-section">
          <a href="${rsvpLink}?status=Confirmed" class="btn btn-confirm">تأكيد الحضور</a>
          <a href="${rsvpLink}?status=Declined" class="btn btn-decline">اعتذار عن الحضور</a>
        </div>
      </div>
      <div class="footer">
        هذه الرسالة مرسلة تلقائياً من نظام إدارة المجتمع.<br>
        إذا كان لديك أي استفسار، يرجى الرد على هذا البريد الإلكتروني.
      </div>
    </div>
  </body>
  </html>
  `;
}

const emailService = {
  // Send invitation email
  sendInvitation: async ({ memberName, email, eventTitle, eventDate, eventLocation, eventDescription, token }) => {
    // Generate dashboard local link for RSVP (can be configured in production with domain)
    const host = process.env.CLIENT_URL || 'http://localhost:5173';
    const rsvpLink = `${host}/rsvp/${token}`;
    const subject = `دعوة حضور: ${eventTitle}`;
    const htmlContent = generateInvitationHtml(memberName, eventTitle, eventDate, eventLocation, eventDescription, rsvpLink);

    // Logging locally for debugging
    logEmailLocally({
      to: email,
      memberName,
      subject,
      rsvpLink,
      eventTitle
    });

    const transporter = getTransporter();
    if (transporter) {
      try {
        await transporter.sendMail({
          from: `"${process.env.SENDER_NAME || 'مجتمع المستثمرين'}" <${process.env.SENDER_EMAIL || process.env.SMTP_USER}>`,
          to: email,
          subject,
          html: htmlContent
        });
        console.log(`[SMTP] Email sent successfully to ${email}`);
        return { success: true, method: 'SMTP' };
      } catch (err) {
        console.error(`[SMTP Error] Failed to send email to ${email}:`, err);
        return { success: false, method: 'SMTP', error: err.message };
      }
    } else {
      console.log(`[Mock Send] Email would be sent to ${email} (No SMTP config)`);
      return { success: true, method: 'Mock' };
    }
  },

  // Get simulated email logs
  getEmailLogs: () => {
    try {
      if (fs.existsSync(LOGS_FILE)) {
        const data = fs.readFileSync(LOGS_FILE, 'utf8');
        return JSON.parse(data || '[]');
      }
      return [];
    } catch (err) {
      return [];
    }
  },

  clearEmailLogs: () => {
    try {
      fs.writeFileSync(LOGS_FILE, JSON.stringify([], null, 2));
      return true;
    } catch (err) {
      return false;
    }
  }
};

module.exports = emailService;
