import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DB_PATH = path.join(__dirname, 'db.json');

const DEFAULT_DATA = {
  locked: false,
  settings: {
    businessName: 'Sharma Traders',
    gstin: '22AAAAA0000A1Z5',
    filingFrequency: 'Monthly',
    state: 'Chhattisgarh',
    dueDateReminder: true,
    mismatchAlert: true,
    monthlySummaryEmail: false
  },
  monthlyData: {
    'Apr 2026': {
      liability: [182000, 96000, 96000],
      itc: [140000, 70000, 70000],
      rec: { matchRate: 91, booksItc: 404000, gstr2bItc: 384000, ineligible: 12000 }
    },
    'May 2026': {
      liability: [205000, 110000, 110000],
      itc: [168000, 82000, 82000],
      rec: { matchRate: 93, booksItc: 468000, gstr2bItc: 452000, ineligible: 9000 }
    },
    'Jun 2026': {
      liability: [198000, 104000, 104000],
      itc: [176000, 90000, 90000],
      rec: { matchRate: 89, booksItc: 478000, gstr2bItc: 452000, ineligible: 18000 }
    },
    'Jul 2026': {
      liability: [231000, 125000, 125000],
      itc: [190000, 98000, 98000],
      rec: { matchRate: 94, booksItc: 502000, gstr2bItc: 486000, ineligible: 7000 }
    },
    'Aug 2026': {
      liability: [244000, 131000, 131000],
      itc: [214000, 121000, 121000],
      rec: { matchRate: 92, booksItc: 572000, gstr2bItc: 548000, ineligible: 11000 }
    },
    'Sep 2026': {
      liability: [262000, 142000, 142000],
      itc: [226000, 118000, 118000],
      rec: { matchRate: 95, booksItc: 590000, gstr2bItc: 562000, ineligible: 6000 }
    },
    'Oct 2026': {
      liability: [279000, 150000, 150000],
      itc: [231000, 140000, 140000],
      rec: { matchRate: 90, booksItc: 578000, gstr2bItc: 534000, ineligible: 16000 }
    }
  },
  rcmDetails: [
    { head: 'Rent', amount: 420000, rate: 18 },
    { head: 'Transport (GTA)', amount: 280000, rate: 5 },
    { head: 'Security service', amount: 150000, rate: 18 }
  ],
  steps: [
    { id: 0, done: false, notes: [] },
    { id: 1, done: false, notes: [] },
    { id: 2, done: false, notes: [] },
    { id: 3, done: false, notes: [] },
    { id: 4, done: false, notes: [] }
  ],
  invoices: [
    {
      id: 'inv-108',
      fy: '2026-27',
      q: 'Q2 (Jul-Sep)',
      m: 'Aug',
      d: '2026-08-01',
      p: 'Metalx Corporation (Opc) Private Limited',
      g: '22AATCM1200E1Z1',
      n: '108',
      t: 9900000,
      i: 0,
      c: 891000,
      s: 891000
    },
    {
      id: 'inv-109',
      fy: '2026-27',
      q: 'Q2 (Jul-Sep)',
      m: 'Aug',
      d: '2026-08-01',
      p: 'Passary Minerals Madhya PVT. LTD.',
      g: '22AAHCP9274B1ZI',
      n: '109',
      t: 51060000,
      i: 0,
      c: 4595400,
      s: 4595400
    },
    {
      id: 'inv-110',
      fy: '2026-27',
      q: 'Q2 (Jul-Sep)',
      m: 'Aug',
      d: '2026-08-01',
      p: 'Passary Minerals Madhya PVT. LTD.',
      g: '22AAHCP9274B1ZI',
      n: '110',
      t: 79680000,
      i: 0,
      c: 7171200,
      s: 7171200
    }
  ]
};

export function getDb() {
  try {
    if (!fs.existsSync(DB_PATH)) {
      saveDb(DEFAULT_DATA);
      return JSON.parse(JSON.stringify(DEFAULT_DATA));
    }
    const raw = fs.readFileSync(DB_PATH, 'utf-8');
    return JSON.parse(raw);
  } catch (err) {
    console.error('Error reading db.json:', err);
    return JSON.parse(JSON.stringify(DEFAULT_DATA));
  }
}

export function saveDb(data) {
  try {
    fs.writeFileSync(DB_PATH, JSON.stringify(data, null, 2), 'utf-8');
    return true;
  } catch (err) {
    console.error('Error writing to db.json:', err);
    return false;
  }
}
