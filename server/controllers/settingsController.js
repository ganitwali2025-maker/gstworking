import { getDb, saveDb } from '../data/store.js';

export const getSettings = (req, res) => {
  try {
    const db = getDb();
    res.json({ success: true, settings: db.settings });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const updateSettings = (req, res) => {
  try {
    const db = getDb();
    const { businessName, gstin, filingFrequency, state, dueDateReminder, mismatchAlert, monthlySummaryEmail } = req.body;

    db.settings = {
      ...db.settings,
      businessName: businessName !== undefined ? businessName : db.settings.businessName,
      gstin: gstin !== undefined ? gstin : db.settings.gstin,
      filingFrequency: filingFrequency !== undefined ? filingFrequency : db.settings.filingFrequency,
      state: state !== undefined ? state : db.settings.state,
      dueDateReminder: dueDateReminder !== undefined ? !!dueDateReminder : db.settings.dueDateReminder,
      mismatchAlert: mismatchAlert !== undefined ? !!mismatchAlert : db.settings.mismatchAlert,
      monthlySummaryEmail: monthlySummaryEmail !== undefined ? !!monthlySummaryEmail : db.settings.monthlySummaryEmail
    };

    saveDb(db);
    res.json({ success: true, settings: db.settings, message: 'Settings saved successfully' });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};
