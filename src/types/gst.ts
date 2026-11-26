export type ViewType =
  | 'itc'
  | 'liab'
  | 'steps'
  | 'g1'
  | 'set'
  | 'rcm'
  | 'rec'
  | 'itc-reg'
  | 'g2b'
  | 'reco-matched'
  | 'reco-missing-books'
  | 'reco-missing-2b';

export interface InvoiceRecord {
  id: string;
  fy: string;
  q: string;
  m: string;
  d: string;
  p: string;
  g: string;
  n: string;
  t: number; // in paise or numeric
  i: number;
  c: number;
  s: number;
  sheetSynced?: boolean;
  createdAt?: string;
  updatedAt?: string;
}

export interface MonthlyStats {
  liability: [number, number, number];
  itc: [number, number, number];
  rec: {
    matchRate: number;
    booksItc: number;
    gstr2bItc: number;
    ineligible: number;
  };
}

export interface StepItem {
  id: number;
  titleEn: string;
  titleHi: string;
  descEn: string;
  descHi: string;
  done: boolean;
  notes: string[];
}

export interface BusinessSettings {
  businessName: string;
  gstin: string;
  filingFrequency: string;
  state: string;
  dueDateReminder: boolean;
  mismatchAlert: boolean;
  monthlySummaryEmail: boolean;
}

export interface RcmItem {
  head: string;
  amount: number;
  rate: number;
}
