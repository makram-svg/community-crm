const fs = require('fs');
const path = require('path');

const DB_DIR = path.join(__dirname, 'data');
const MEMBERS_FILE = path.join(DB_DIR, 'members.json');
const EVENTS_FILE = path.join(DB_DIR, 'events.json');
const INVITATIONS_FILE = path.join(DB_DIR, 'invitations.json');

// Ensure database directory exists
if (!fs.existsSync(DB_DIR)) {
  fs.mkdirSync(DB_DIR, { recursive: true });
}

// Helper to read JSON file safely
function readJSON(file, defaultData = []) {
  try {
    if (!fs.existsSync(file)) {
      fs.writeFileSync(file, JSON.stringify(defaultData, null, 2));
      return defaultData;
    }
    const data = fs.readFileSync(file, 'utf8');
    return JSON.parse(data || '[]');
  } catch (err) {
    console.error(`Error reading ${file}:`, err);
    return defaultData;
  }
}

// Helper to write JSON file atomically
function writeJSON(file, data) {
  try {
    const tempFile = `${file}.tmp`;
    fs.writeFileSync(tempFile, JSON.stringify(data, null, 2));
    fs.renameSync(tempFile, file);
    return true;
  } catch (err) {
    console.error(`Error writing ${file}:`, err);
    return false;
  }
}

const db = {
  // --- MEMBERS ---
  getMembers: () => readJSON(MEMBERS_FILE),
  
  saveMembers: (members) => writeJSON(MEMBERS_FILE, members),
  
  addMember: (member) => {
    const members = db.getMembers();
    const newMember = {
      id: 'm_' + Math.random().toString(36).substr(2, 9),
      name: member.name || '',
      email: member.email || '',
      phone: member.phone || '',
      role: member.role || 'Entrepreneur', // 'Investor' or 'Entrepreneur'
      company: member.company || '',
      industry: member.industry || '',
      ticketSize: member.ticketSize || '', // For investors: e.g. "$10k - $50k"
      stage: member.stage || '', // e.g. "Pre-seed", "Seed", "Series A"
      linkedin: member.linkedin || '',
      bio: member.bio || '',
      engagementScore: parseInt(member.engagementScore) || 50,
      createdAt: member.createdAt || new Date().toISOString()
    };
    members.push(newMember);
    db.saveMembers(members);
    return newMember;
  },

  updateMember: (id, updates) => {
    const members = db.getMembers();
    const index = members.findIndex(m => m.id === id);
    if (index === -1) return null;
    members[index] = { ...members[index], ...updates };
    db.saveMembers(members);
    return members[index];
  },

  deleteMember: (id) => {
    const members = db.getMembers();
    const filtered = members.filter(m => m.id !== id);
    db.saveMembers(filtered);
    // Clean up invitations
    const invitations = db.getInvitations();
    const filteredInvs = invitations.filter(i => i.memberId !== id);
    db.saveInvitations(filteredInvs);
    return true;
  },

  // --- EVENTS ---
  getEvents: () => readJSON(EVENTS_FILE),
  
  saveEvents: (events) => writeJSON(EVENTS_FILE, events),
  
  addEvent: (event) => {
    const events = db.getEvents();
    const newEvent = {
      id: 'e_' + Math.random().toString(36).substr(2, 9),
      title: event.title || '',
      date: event.date || '',
      location: event.location || '',
      description: event.description || '',
      createdAt: new Date().toISOString()
    };
    events.push(newEvent);
    db.saveEvents(events);
    return newEvent;
  },

  updateEvent: (id, updates) => {
    const events = db.getEvents();
    const index = events.findIndex(e => e.id === id);
    if (index === -1) return null;
    events[index] = { ...events[index], ...updates };
    db.saveEvents(events);
    return events[index];
  },

  deleteEvent: (id) => {
    const events = db.getEvents();
    const filtered = events.filter(e => e.id !== id);
    db.saveEvents(filtered);
    // Clean up invitations
    const invitations = db.getInvitations();
    const filteredInvs = invitations.filter(i => i.eventId !== id);
    db.saveInvitations(filteredInvs);
    return true;
  },

  // --- INVITATIONS ---
  getInvitations: () => readJSON(INVITATIONS_FILE),
  
  saveInvitations: (invitations) => writeJSON(INVITATIONS_FILE, invitations),
  
  addInvitation: (eventId, memberId) => {
    const invitations = db.getInvitations();
    // Check if invitation already exists
    const existing = invitations.find(i => i.eventId === eventId && i.memberId === memberId);
    if (existing) return existing;

    const newInv = {
      id: 'inv_' + Math.random().toString(36).substr(2, 9),
      eventId,
      memberId,
      rsvpStatus: 'Pending', // 'Pending', 'Confirmed', 'Declined'
      attended: 'Pending', // 'Pending', 'Yes', 'No'
      token: Math.random().toString(36).substr(2, 15) + Math.random().toString(36).substr(2, 15),
      sentAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    invitations.push(newInv);
    db.saveInvitations(invitations);
    return newInv;
  },

  updateInvitation: (id, updates) => {
    const invitations = db.getInvitations();
    const index = invitations.findIndex(i => i.id === id);
    if (index === -1) return null;
    invitations[index] = { 
      ...invitations[index], 
      ...updates, 
      updatedAt: new Date().toISOString() 
    };
    db.saveInvitations(invitations);
    return invitations[index];
  },

  updateInvitationByToken: (token, updates) => {
    const invitations = db.getInvitations();
    const index = invitations.findIndex(i => i.token === token);
    if (index === -1) return null;
    invitations[index] = { 
      ...invitations[index], 
      ...updates, 
      updatedAt: new Date().toISOString() 
    };
    db.saveInvitations(invitations);
    return invitations[index];
  },

  // --- SEED SEED DATA ---
  seed: () => {
    const members = db.getMembers();
    if (members.length > 0) return; // Already seeded

    console.log('Seeding initial data...');
    
    const seedMembers = [
      {
        name: 'عبدالرحمن آل سعود',
        email: 'abdulrahman.invest@example.com',
        phone: '+966501234567',
        role: 'Investor',
        company: 'بدر للاستثمار الجريء',
        industry: 'Fintech, SaaS',
        ticketSize: '$50k - $200k',
        stage: 'Seed, Series A',
        linkedin: 'https://linkedin.com/in/abdulrahman-investor',
        bio: 'مستثمر جريء مهتم بدعم الشركات الناشئة التقنية في المملكة ومنطقة الخليج.',
        engagementScore: 85
      },
      {
        name: 'سارة التميمي',
        email: 'sara.t@fintechventures.com',
        phone: '+966502345678',
        role: 'Investor',
        company: 'سنابل الاستثمارية',
        industry: 'E-commerce, Logistics',
        ticketSize: '$100k - $500k',
        stage: 'Series A, Series B',
        linkedin: 'https://linkedin.com/in/sara-tamimi',
        bio: 'مديرة استثمار مهتمة بقطاع الخدمات اللوجستية والتجارة الإلكترونية.',
        engagementScore: 90
      },
      {
        name: 'أحمد الشمري',
        email: 'ahmad@flowtech.sa',
        phone: '+966503456789',
        role: 'Entrepreneur',
        company: 'فلوتك للحلول المالية (FlowTech)',
        industry: 'Fintech',
        ticketSize: '',
        stage: 'Seed',
        linkedin: 'https://linkedin.com/in/ahmad-flowtech',
        bio: 'مؤسس منصة فلوتك لحلول الدفع الإلكتروني الميسر للشركات الصغيرة.',
        engagementScore: 75
      },
      {
        name: 'نورة القحطاني',
        email: 'noura@eduhub.sa',
        phone: '+966504567890',
        role: 'Entrepreneur',
        company: 'إديو هب (EduHub)',
        industry: 'Edtech',
        ticketSize: '',
        stage: 'Pre-seed',
        linkedin: 'https://linkedin.com/in/noura-eduhub',
        bio: 'شريكة مؤسسة لمنصة تعليمية مبتكرة تقدم دورات تقنية وتفاعلية للطلاب.',
        engagementScore: 65
      },
      {
        name: 'خالد المطيري',
        email: 'khaled.m@proptech.sa',
        phone: '+966505678901',
        role: 'Entrepreneur',
        company: 'عقارك الذكي (SmartProp)',
        industry: 'Proptech, Real Estate',
        ticketSize: '',
        stage: 'Seed',
        linkedin: 'https://linkedin.com/in/khaled-smartprop',
        bio: 'مهندس برمجيات ريادي يطور حلولاً ذكية لإدارة العقارات والمجمعات السكنية.',
        engagementScore: 80
      }
    ];

    seedMembers.forEach(m => db.addMember(m));

    const seedEvents = [
      {
        title: 'اللقاء الدوري الأول لمجتمع الرياديين والمستثمرين',
        date: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0], // 7 days from now
        location: 'مساحة عمل كود، الرياض',
        description: 'لقاء مفتوح للتعارف ومناقشة التوجهات الاستثمارية الجديدة في قطاع التقنية المالية بالمملكة العربية السعودية.'
      },
      {
        title: 'ورشة عمل: كيف تجهز شركتك لجولة الاستثمار القادمة؟',
        date: new Date(Date.now() + 14 * 24 * 60 * 60 * 1000).toISOString().split('T')[0], // 14 days from now
        location: 'فندق الهيلتون، قاعة المؤتمرات، الرياض',
        description: 'ورشة عمل تفصيلية بالتعاون مع مستشارين استثماريين لشرح آليات التقييم وتجهيز العروض الاستثمارية.'
      }
    ];

    seedEvents.forEach(e => {
      const addedEvent = db.addEvent(e);
      // Automatically send invitations to all seeded members
      const members = db.getMembers();
      members.forEach(m => {
        const inv = db.addInvitation(addedEvent.id, m.id);
        // Randomly set RSVP status for the first event to make data look alive
        if (addedEvent.title.includes('اللقاء الدوري الأول')) {
          const rand = Math.random();
          if (rand > 0.6) {
            db.updateInvitation(inv.id, { rsvpStatus: 'Confirmed', attended: 'Yes' });
          } else if (rand > 0.3) {
            db.updateInvitation(inv.id, { rsvpStatus: 'Confirmed', attended: 'Pending' });
          } else {
            db.updateInvitation(inv.id, { rsvpStatus: 'Declined', attended: 'No' });
          }
        }
      });
    });

    console.log('Seeding completed successfully!');
  }
};

// Initialize seed data
db.seed();

module.exports = db;
