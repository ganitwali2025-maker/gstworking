import { getDb, saveDb } from '../data/store.js';

const STEP_DEFINITIONS = [
  {
    id: 0,
    titleEn: 'GSTR-1 – Sales Output',
    titleHi: 'बिक्री विवरण',
    descEn: 'Record and verify outward supplies and output GST.',
    descHi: 'बिक्री एवं आउटपुट GST का मिलान और सत्यापन करें।'
  },
  {
    id: 1,
    titleEn: 'Input Tax Credit (ITC)',
    titleHi: 'इनपुट टैक्स क्रेडिट',
    descEn: 'Calculate eligible ITC from purchase records.',
    descHi: 'खरीद रिकॉर्ड के आधार पर पात्र ITC की गणना करें।'
  },
  {
    id: 2,
    titleEn: 'Reverse Charge Mechanism (RCM)',
    titleHi: 'रिवर्स चार्ज',
    descEn: 'Calculate RCM liability and eligible ITC.',
    descHi: 'RCM टैक्स देयता एवं पात्र ITC की गणना करें।'
  },
  {
    id: 3,
    titleEn: 'GSTR-2B – ITC Statement',
    titleHi: 'ITC विवरण',
    descEn: 'Verify eligible ITC as per GSTR-2B.',
    descHi: 'GSTR-2B के अनुसार पात्र ITC का सत्यापन करें।'
  },
  {
    id: 4,
    titleEn: 'ITC Reconciliation',
    titleHi: 'ITC मिलान',
    descEn: 'Reconcile purchase books with GSTR-2B and identify differences.',
    descHi: 'खरीद बुक एवं GSTR-2B का मिलान कर अंतर पहचानें।'
  }
];

export const getSteps = (req, res) => {
  try {
    const db = getDb();
    const storedSteps = db.steps || [];

    const merged = STEP_DEFINITIONS.map(def => {
      const match = storedSteps.find(s => s.id === def.id) || { done: false, notes: [] };
      return {
        ...def,
        done: !!match.done,
        notes: match.notes || []
      };
    });

    res.json({ success: true, steps: merged });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const updateStepStatus = (req, res) => {
  try {
    const db = getDb();
    const stepId = parseInt(req.params.id, 10);
    const { done } = req.body;

    if (!db.steps) db.steps = [];
    let step = db.steps.find(s => s.id === stepId);
    if (!step) {
      step = { id: stepId, done: false, notes: [] };
      db.steps.push(step);
    }

    step.done = !!done;
    saveDb(db);

    res.json({ success: true, step });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const addStepNote = (req, res) => {
  try {
    const db = getDb();
    const stepId = parseInt(req.params.id, 10);
    const { note } = req.body;

    if (!note || !note.trim()) {
      return res.status(400).json({ success: false, message: 'Note text cannot be empty.' });
    }

    if (!db.steps) db.steps = [];
    let step = db.steps.find(s => s.id === stepId);
    if (!step) {
      step = { id: stepId, done: false, notes: [] };
      db.steps.push(step);
    }

    step.notes.push(note.trim());
    saveDb(db);

    res.json({ success: true, notes: step.notes });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const deleteStepNote = (req, res) => {
  try {
    const db = getDb();
    const stepId = parseInt(req.params.id, 10);
    const noteIndex = parseInt(req.params.noteIndex, 10);

    const step = (db.steps || []).find(s => s.id === stepId);
    if (!step || !step.notes || !step.notes[noteIndex]) {
      return res.status(404).json({ success: false, message: 'Note not found.' });
    }

    step.notes.splice(noteIndex, 1);
    saveDb(db);

    res.json({ success: true, notes: step.notes });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};
