import { MonthlyStats, StepItem, BusinessSettings, RcmItem, InvoiceRecord } from '../types/gst';

export const SEED_MONTHLY_DATA: Record<string, MonthlyStats> = {
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
  },
  'Nov 2026': {
    liability: [285000, 154000, 154000],
    itc: [238000, 142000, 142000],
    rec: { matchRate: 92, booksItc: 586000, gstr2bItc: 548000, ineligible: 15000 }
  },
  'Dec 2026': {
    liability: [292000, 158000, 158000],
    itc: [245000, 146000, 146000],
    rec: { matchRate: 94, booksItc: 602000, gstr2bItc: 574000, ineligible: 14000 }
  },
  'Jan 2027': {
    liability: [275000, 148000, 148000],
    itc: [228000, 138000, 138000],
    rec: { matchRate: 91, booksItc: 564000, gstr2bItc: 532000, ineligible: 13000 }
  },
  'Feb 2027': {
    liability: [268000, 144000, 144000],
    itc: [222000, 134000, 134000],
    rec: { matchRate: 93, booksItc: 550000, gstr2bItc: 524000, ineligible: 12000 }
  },
  'Mar 2027': {
    liability: [310000, 168000, 168000],
    itc: [260000, 156000, 156000],
    rec: { matchRate: 96, booksItc: 642000, gstr2bItc: 618000, ineligible: 17000 }
  }
};

export const SEED_STEPS: StepItem[] = [
  {
    id: 0,
    titleEn: 'GSTR-1 – Sales Output',
    titleHi: 'बिक्री विवरण',
    descEn: 'Record and verify outward supplies and output GST.',
    descHi: 'बिक्री एवं आउटपुट GST का मिलान और सत्यापन करें।',
    done: false,
    notes: []
  },
  {
    id: 1,
    titleEn: 'Input Tax Credit (ITC)',
    titleHi: 'इनपुट टैक्स क्रेडिट',
    descEn: 'Calculate eligible ITC from purchase records.',
    descHi: 'खरीद रिकॉर्ड के आधार पर पात्र ITC की गणना करें।',
    done: false,
    notes: []
  },
  {
    id: 2,
    titleEn: 'Reverse Charge Mechanism (RCM)',
    titleHi: 'रिवर्स चार्ज',
    descEn: 'Calculate RCM liability and eligible ITC.',
    descHi: 'RCM टैक्स देयता एवं पात्र ITC की गणना करें।',
    done: false,
    notes: []
  },
  {
    id: 3,
    titleEn: 'GSTR-2B – ITC Statement',
    titleHi: 'ITC विवरण',
    descEn: 'Verify eligible ITC as per GSTR-2B.',
    descHi: 'GSTR-2B के अनुसार पात्र ITC का सत्यापन करें।',
    done: false,
    notes: []
  },
  {
    id: 4,
    titleEn: 'ITC Reconciliation',
    titleHi: 'ITC मिलान',
    descEn: 'Reconcile purchase books with GSTR-2B and identify differences.',
    descHi: 'खरीद बुक एवं GSTR-2B का मिलान कर अंतर पहचानें।',
    done: false,
    notes: []
  }
];

export const SEED_RCM: RcmItem[] = [
  { head: 'Rent', amount: 420000, rate: 18 },
  { head: 'Transport (GTA)', amount: 280000, rate: 5 },
  { head: 'Security service', amount: 150000, rate: 18 }
];

export const SEED_SETTINGS: BusinessSettings = {
  businessName: 'Sharma Traders',
  gstin: '22AAAAA0000A1Z5',
  filingFrequency: 'Monthly',
  state: 'Chhattisgarh',
  dueDateReminder: true,
  mismatchAlert: true,
  monthlySummaryEmail: false
};

export const SEED_INVOICES: InvoiceRecord[] = [
  {
    id: 'inv-108',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-01',
    p: 'Metalx Corporation (Opc) Private...',
    g: '22AATCM1200E1Z1',
    n: '108',
    t: 9900000,
    i: 0,
    c: 891000,
    s: 891000
  },
  {
    id: 'inv-109',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-01',
    p: 'Passary Minerals Madhya PVT. L..',
    g: '22AAHCP9274B1ZI',
    n: '109',
    t: 51060000,
    i: 0,
    c: 4595400,
    s: 4595400
  },
  {
    id: 'inv-110',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-01',
    p: 'Passary Minerals Madhya PVT. L..',
    g: '22AAHCP9274B1ZI',
    n: '110',
    t: 79680000,
    i: 0,
    c: 7171200,
    s: 7171200
  },
  {
    id: 'inv-111',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-01',
    p: 'Passary Minerals Madhya PVT. L..',
    g: '22AAHCP9274B1ZI',
    n: '111',
    t: 61020000,
    i: 0,
    c: 5491800,
    s: 5491800
  },
  {
    id: 'inv-112',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-01',
    p: 'Metalx Corporation (Opc) Private...',
    g: '22AATCM1200E1Z1',
    n: '112',
    t: 240000,
    i: 0,
    c: 21600,
    s: 21600
  },
  {
    id: 'inv-113',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-04',
    p: 'Passary Minerals Private Limited',
    g: '21AABCP0611Q1ZO',
    n: '113',
    t: 196000000,
    i: 35280000,
    c: 0,
    s: 0
  },
  {
    id: 'inv-114',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-04',
    p: 'Passary Minerals Private Limited',
    g: '21AABCP0611Q1ZO',
    n: '114',
    t: 34888000,
    i: 6279840,
    c: 0,
    s: 0
  },
  {
    id: 'inv-115',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-05',
    p: 'Passary Minerals Private Limited',
    g: '21AABCP0611Q1ZO',
    n: '115',
    t: 134356000,
    i: 24184100,
    c: 0,
    s: 0
  },
  {
    id: 'inv-116',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-05',
    p: 'Passary Minerals Private Limited',
    g: '21AABCP0611Q1ZO',
    n: '116',
    t: 40432000,
    i: 7277800,
    c: 0,
    s: 0
  },
  {
    id: 'inv-117',
    fy: 'FY 2026–27',
    q: 'Q2 (Jul-Sep)',
    m: 'Aug',
    d: '2026-08-06',
    p: 'Passary Minerals Private Limited',
    g: '21AABCP0611Q1ZO',
    n: '117',
    t: 174832000,
    i: 31469800,
    c: 0,
    s: 0
  }
];
