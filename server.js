require('dotenv').config();
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const db = require('./database');
const emailService = require('./emailService');

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(bodyParser.json({ limit: '10mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '10mb' }));

// --- API ROUTES ---

// 1. Members Endpoints
app.get('/api/members', (req, res) => {
  try {
    const { role, industry, stage, search } = req.query;
    let members = db.getMembers();

    if (role) {
      members = members.filter(m => m.role.toLowerCase() === role.toLowerCase());
    }
    if (industry) {
      members = members.filter(m => m.industry.toLowerCase().includes(industry.toLowerCase()));
    }
    if (stage) {
      members = members.filter(m => m.stage && m.stage.toLowerCase() === stage.toLowerCase());
    }
    if (search) {
      const q = search.toLowerCase();
      members = members.filter(m => 
        m.name.toLowerCase().includes(q) || 
        m.email.toLowerCase().includes(q) || 
        m.company.toLowerCase().includes(q) ||
        m.phone.includes(q)
      );
    }

    res.json(members);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/members', (req, res) => {
  try {
    const member = db.addMember(req.body);
    res.status(201).json(member);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/members/:id', (req, res) => {
  try {
    const member = db.updateMember(req.params.id, req.body);
    if (!member) return res.status(404).json({ error: 'Member not found' });
    res.json(member);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/members/:id', (req, res) => {
  try {
    db.deleteMember(req.params.id);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/members/import', (req, res) => {
  try {
    const { members } = req.body;
    if (!Array.isArray(members)) {
      return res.status(400).json({ error: 'Invalid members format' });
    }
    
    const imported = [];
    members.forEach(m => {
      // Basic duplicate check by email
      const existing = db.getMembers().find(ex => ex.email.toLowerCase() === m.email.toLowerCase());
      if (!existing && m.name && m.email) {
        const added = db.addMember(m);
        imported.push(added);
      }
    });
    
    res.json({ success: true, count: imported.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 2. Events Endpoints
app.get('/api/events', (req, res) => {
  try {
    const events = db.getEvents();
    res.json(events);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/events', (req, res) => {
  try {
    const event = db.addEvent(req.body);
    res.status(201).json(event);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/events/:id', (req, res) => {
  try {
    db.deleteEvent(req.params.id);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 3. Invitation & RSVP Workflow
app.get('/api/events/:eventId/invitations', (req, res) => {
  try {
    const { eventId } = req.params;
    const invitations = db.getInvitations().filter(i => i.eventId === eventId);
    const members = db.getMembers();
    
    // Join with member details
    const joined = invitations.map(inv => {
      const member = members.find(m => m.id === inv.memberId);
      return {
        ...inv,
        memberName: member ? member.name : 'Unknown',
        memberEmail: member ? member.email : 'Unknown',
        memberRole: member ? member.role : 'Unknown',
        memberCompany: member ? member.company : ''
      };
    });
    
    res.json(joined);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Send invitations to a list of members
app.post('/api/events/:eventId/invite', async (req, res) => {
  try {
    const { eventId } = req.params;
    const { memberIds } = req.body; // Array of member IDs
    
    if (!Array.isArray(memberIds) || memberIds.length === 0) {
      return res.status(400).json({ error: 'No members selected' });
    }

    const event = db.getEvents().find(e => e.id === eventId);
    if (!event) return res.status(404).json({ error: 'Event not found' });

    const results = [];
    for (const memberId of memberIds) {
      const member = db.getMembers().find(m => m.id === memberId);
      if (member) {
        const invitation = db.addInvitation(eventId, memberId);
        // Trigger email send
        const mailResult = await emailService.sendInvitation({
          memberName: member.name,
          email: member.email,
          eventTitle: event.title,
          eventDate: event.date,
          eventLocation: event.location,
          eventDescription: event.description,
          token: invitation.token
        });
        results.push({ email: member.email, success: mailResult.success });
      }
    }

    res.json({ success: true, invitesSent: results.length, details: results });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update RSVP or attendance manually
app.put('/api/invitations/:id', (req, res) => {
  try {
    const { rsvpStatus, attended } = req.body;
    const inv = db.updateInvitation(req.params.id, { rsvpStatus, attended });
    if (!inv) return res.status(404).json({ error: 'Invitation not found' });

    // Recalculate engagement score for the member
    recalculateEngagementScore(inv.memberId);

    res.json(inv);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Public RSVP details by Token
app.get('/api/rsvp/:token', (req, res) => {
  try {
    const invitations = db.getInvitations();
    const inv = invitations.find(i => i.token === req.params.token);
    if (!inv) return res.status(404).json({ error: 'دعوة غير صالحة أو منتهية الصلاحية' });

    const event = db.getEvents().find(e => e.id === inv.eventId);
    const member = db.getMembers().find(m => m.id === inv.memberId);

    res.json({
      invitationId: inv.id,
      rsvpStatus: inv.rsvpStatus,
      event: {
        title: event.title,
        date: event.date,
        location: event.location,
        description: event.description
      },
      member: {
        name: member.name,
        email: member.email,
        company: member.company
      }
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Public RSVP Submit
app.post('/api/rsvp/:token', (req, res) => {
  try {
    const { status } = req.body; // 'Confirmed' or 'Declined'
    if (!['Confirmed', 'Declined'].includes(status)) {
      return res.status(400).json({ error: 'Invalid status' });
    }

    const inv = db.updateInvitationByToken(req.params.token, { rsvpStatus: status });
    if (!inv) return res.status(404).json({ error: 'Invitation not found' });

    // Recalculate engagement score
    recalculateEngagementScore(inv.memberId);

    res.json({ success: true, status: inv.rsvpStatus });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 4. Webhook for Google Sheets Integration
app.post('/api/webhook/google-sheets', (req, res) => {
  try {
    // Expected payload layout:
    // {
    //   "name": "Full Name",
    //   "email": "email@domain.com",
    //   "phone": "+9665...",
    //   "role": "Investor / Entrepreneur",
    //   "company": "Company Name",
    //   "industry": "Fintech",
    //   "ticketSize": "$10k-$50k",
    //   "stage": "Seed",
    //   "linkedin": "...",
    //   "bio": "..."
    // }
    const { name, email, phone, role, company, industry, ticketSize, stage, linkedin, bio } = req.body;

    if (!name || !email) {
      return res.status(400).json({ error: 'Name and Email are required parameters' });
    }

    // Check duplicate
    const existing = db.getMembers().find(m => m.email.toLowerCase() === email.toLowerCase());
    if (existing) {
      // Update existing record
      const updated = db.updateMember(existing.id, {
        name, phone, role, company, industry, ticketSize, stage, linkedin, bio
      });
      return res.json({ success: true, action: 'updated', member: updated });
    }

    // Add new member
    const newMember = db.addMember({
      name, email, phone, role, company, industry, ticketSize, stage, linkedin, bio
    });

    res.status(201).json({ success: true, action: 'created', member: newMember });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 5. Simulated Email Logs
app.get('/api/email-logs', (req, res) => {
  res.json(emailService.getEmailLogs());
});

app.delete('/api/email-logs', (req, res) => {
  const success = emailService.clearEmailLogs();
  res.json({ success });
});

// --- HELPER FUNCTIONS ---

function recalculateEngagementScore(memberId) {
  const invitations = db.getInvitations().filter(i => i.memberId === memberId);
  if (invitations.length === 0) return;

  const totalInvited = invitations.length;
  
  // Confirmed RSVPs
  const confirmed = invitations.filter(i => i.rsvpStatus === 'Confirmed').length;
  // Declined RSVPs
  const declined = invitations.filter(i => i.rsvpStatus === 'Declined').length;
  // Actual attendance
  const attended = invitations.filter(i => i.attended === 'Yes').length;
  // Missed event (RSVP'd Confirmed but didn't show up)
  const missed = invitations.filter(i => i.rsvpStatus === 'Confirmed' && i.attended === 'No').length;

  let score = 50; // Base score

  // Engagement math
  const attendanceRate = attended / totalInvited;
  const rsvpRate = (confirmed + declined) / totalInvited;

  score += Math.round(attendanceRate * 40); // Max +40 points for high attendance
  score += Math.round(rsvpRate * 15);       // Max +15 points for responding
  score -= missed * 15;                     // Deduct 15 points for each "no-show" without declining

  // Limit score range between 0 and 100
  score = Math.max(0, Math.min(100, score));

  db.updateMember(memberId, { engagementScore: score });
}

// Serve static HTML files
const path = require('path');
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'portal.html')));
app.get('/portal', (req, res) => res.sendFile(path.join(__dirname, 'portal.html')));
app.get('/admin', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
