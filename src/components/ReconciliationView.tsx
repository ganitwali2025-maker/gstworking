import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { getQueryParams, updateQueryParams } from '../utils/router';
import defaultRegisters from '../data/defaultRegisters.json';
import { loadRegisterRows, saveRegisterRow, REGISTER_TABS } from '../services/registerService';

// ==========================================
// Types & Interfaces
// ==========================================
export type RecoStatus =
  | 'Matched'
  | 'Not Match'
  | 'Not in GSTR-2B'
  | 'Not in Purchase Register'
  | 'Debit Note';

export type AuditStatus = 'Verified' | 'Query' | 'Pending';

export interface AuditEntry {
  status: AuditStatus;
  note: string;
  updatedAt: string;
}

export interface AuditFieldChangeLog {
  id: string;
  invNo: string;
  gstin: string;
  field: string;
  oldValue: string;
  newValue: string;
  user: string;
  timestamp: string;
}

export interface AuditEditState {
  gstin: string;
  invNo: string;
  supplierName: string;
  date: string;
  taxable: number;
  igst: number;
  cgst: number;
  sgst: number;
  note: string;
  status: AuditStatus;
}

export interface ReconciledInvoice {
  id: string;
  status: RecoStatus;
  month2b: string; // e.g. 'Apr-2026' or '-'
  hasDuplicateMonth: boolean;
  duplicateMonthsList?: string[];
  itcMatchText: string; // e.g. '6/6 Match', '4/6 Match', '-'
  itcMatchCount: number; // 0 to 6
  mismatchedFields: string[]; // e.g. ['Taxable Not Match', 'IGST Not Match']
  date: string;
  invNo: string;
  gstin: string;
  supplierName: string;

  // Books (ITC Purchase Register) Values in Rupees
  bTaxable: number;
  bIgst: number;
  bCgst: number;
  bSgst: number;
  bTotal: number;

  // Govt GSTR-2B Values in Rupees
  gTaxable: number;
  gIgst: number;
  gCgst: number;
  gSgst: number;
  gTotal: number;

  // Difference (Books - Govt 2B) Values in Rupees
  dTaxable: number;
  dIgst: number;
  dCgst: number;
  dSgst: number;
  dTotal: number;

  // Paired 2B raw fields for side-by-side View
  gDate?: string;
  gInvNo?: string;
  gGstin?: string;
  gTradeName?: string;

  // Individual field match flags
  isGstinMatch: boolean;
  isInvMatch: boolean;
  isTaxableMatch: boolean;
  isIgstMatch: boolean;
  isCgstMatch: boolean;
  isSgstMatch: boolean;

  // Audit Info
  auditStatus: AuditStatus;
  auditNote: string;
  auditUpdatedAt?: string;
}

interface RawBookRecord {
  id: string;
  date: string;
  invNo: string;
  gstin: string;
  supplierName: string;
  taxable: number;
  igst: number;
  cgst: number;
  sgst: number;
  cess?: number;
  month: string;
}

interface Raw2bRecord {
  id: string;
  date: string;
  invNo: string;
  gstin: string;
  supplierName: string;
  taxable: number;
  igst: number;
  cgst: number;
  sgst: number;
  cess?: number;
  month: string;
}

// ==========================================
// Default Seed Data (Demonstrates All Cases)
// ==========================================
export const DEFAULT_BOOKS: RawBookRecord[] = [
  {
    id: 'b-1',
    date: '02/08/2026',
    invNo: 'BS-4421',
    gstin: '22AABCB4567M1Z2',
    supplierName: 'Bharat Steel & Alloys Pvt Ltd',
    taxable: 420000,
    igst: 0,
    cgst: 37800,
    sgst: 37800,
    month: 'Aug-2026'
  },
  {
    id: 'b-2',
    date: '05/08/2026',
    invNo: 'SK-1092',
    gstin: '22AAFCK8812D1Z4',
    supplierName: 'Shree Krishna Logistics & Spares',
    taxable: 280000,
    igst: 0,
    cgst: 25200,
    sgst: 25200,
    month: 'Aug-2026'
  },
  {
    id: 'b-3',
    date: '08/08/2026',
    invNo: 'TS-9901',
    gstin: '21AABCT1332N1Z8',
    supplierName: 'Tata Steel Downstream Products Ltd',
    taxable: 850000,
    igst: 153000,
    cgst: 0,
    sgst: 0,
    month: 'Aug-2026'
  },
  {
    id: 'b-4',
    date: '15/07/2026',
    invNo: 'GI/26-27/318',
    gstin: '21AILPG1979R1ZO',
    supplierName: 'Glaxy International',
    taxable: 1800000,
    igst: 324000,
    cgst: 0,
    sgst: 0,
    month: 'Jul-2026'
  },
  {
    id: 'b-5',
    date: '15/08/2026',
    invNo: 'DN/2026/089',
    gstin: '22AABCB4567M1Z2',
    supplierName: 'Bharat Steel & Alloys Pvt Ltd',
    taxable: -50000,
    igst: 0,
    cgst: -4500,
    sgst: -4500,
    month: 'Aug-2026'
  },
  {
    id: 'b-6',
    date: '22/08/2026',
    invNo: 'DN-TK-441',
    gstin: '21AABCT1332N1Z8',
    supplierName: 'Tata Steel Downstream Products Ltd',
    taxable: -25000,
    igst: -4500,
    cgst: 0,
    sgst: 0,
    month: 'Aug-2026'
  },
  {
    id: 'b-7',
    date: '25/08/2026',
    invNo: '26-27/25',
    gstin: '22BVZPK4622B1ZO',
    supplierName: 'Kalam Khan Transport',
    taxable: 778635,
    igst: 0,
    cgst: 70077.15,
    sgst: 70077.15,
    month: 'Aug-2026'
  },
  {
    id: 'b-8',
    date: '28/08/2026',
    invNo: 'AP-8841',
    gstin: '22AANCA4419K1ZM',
    supplierName: 'Adani Power & Infra Ltd',
    taxable: 1250000,
    igst: 225000,
    cgst: 0,
    sgst: 0,
    month: 'Aug-2026'
  }
];

export const DEFAULT_GSTR2B: Raw2bRecord[] = [
  // Exact Match for b-1 (Found in Aug-2026)
  {
    id: 'g-1',
    date: '02/08/2026',
    invNo: 'BS-4421',
    gstin: '22AABCB4567M1Z2',
    supplierName: 'Bharat Steel & Alloys Pvt Ltd',
    taxable: 420000,
    igst: 0,
    cgst: 37800,
    sgst: 37800,
    month: 'Aug-2026'
  },
  // Mismatch for b-2 (Taxable & Tax difference)
  {
    id: 'g-2',
    date: '05/08/2026',
    invNo: 'SK-1092',
    gstin: '22AAFCK8812D1Z4',
    supplierName: 'Shree Krishna Logistics & Spares',
    taxable: 275000, // Rs 5,000 difference
    igst: 0,
    cgst: 24750,
    sgst: 24750,
    month: 'Aug-2026'
  },
  // Match for b-3 in Jul-2026 (Different month demonstration)
  {
    id: 'g-3',
    date: '08/08/2026',
    invNo: 'TS-9901',
    gstin: '21AABCT1332N1Z8',
    supplierName: 'Tata Steel Downstream Products Ltd',
    taxable: 850000,
    igst: 153000,
    cgst: 0,
    sgst: 0,
    month: 'Jul-2026'
  },
  // Invoice No mismatch for b-4 (GI/26-27/318 vs GI-318)
  {
    id: 'g-4',
    date: '15/07/2026',
    invNo: 'GI-318',
    gstin: '21AILPG1979R1ZO',
    supplierName: 'Glaxy International',
    taxable: 1800000,
    igst: 324000,
    cgst: 0,
    sgst: 0,
    month: 'Jul-2026'
  },
  // Duplicate month simulation: TS-9901 also uploaded in Aug-2026
  {
    id: 'g-3-dup',
    date: '08/08/2026',
    invNo: 'TS-9901',
    gstin: '21AABCT1332N1Z8',
    supplierName: 'Tata Steel Downstream Products Ltd',
    taxable: 850000,
    igst: 153000,
    cgst: 0,
    sgst: 0,
    month: 'Aug-2026'
  },
  // Not in Books (Govt GSTR-2B only)
  {
    id: 'g-5',
    date: '20/08/2026',
    invNo: 'INV-7712',
    gstin: '27AABCS9912K1Z9',
    supplierName: 'Siemens India Industrial Ltd',
    taxable: 650000,
    igst: 117000,
    cgst: 0,
    sgst: 0,
    month: 'Aug-2026'
  },
  {
    id: 'g-6',
    date: '24/08/2026',
    invNo: 'RJ-901',
    gstin: '22AAPCR1120F1ZX',
    supplierName: 'Rajasthan Mining Works',
    taxable: 340000,
    igst: 0,
    cgst: 30600,
    sgst: 30600,
    month: 'Aug-2026'
  }
];

// ==========================================
// Formatting Helpers
// ==========================================
export const round2 = (val: number): number => Math.round((val + Number.EPSILON) * 100) / 100;

export const cleanStr = (s?: string): string => (s || '').trim().toUpperCase().replace(/\s+/g, ' ');

export const cleanInv = (s?: string): string => (s || '').trim().toUpperCase().replace(/[\s\-_/]/g, '');

export const ALL_FY_MONTHS = [
  { value: 'Apr-2026', label: 'April-2026', short: 'Apr' },
  { value: 'May-2026', label: 'May-2026', short: 'May' },
  { value: 'Jun-2026', label: 'June-2026', short: 'Jun' },
  { value: 'Jul-2026', label: 'July-2026', short: 'Jul' },
  { value: 'Aug-2026', label: 'August-2026', short: 'Aug' },
  { value: 'Sep-2026', label: 'September-2026', short: 'Sep' },
  { value: 'Oct-2026', label: 'October-2026', short: 'Oct' },
  { value: 'Nov-2026', label: 'November-2026', short: 'Nov' },
  { value: 'Dec-2026', label: 'December-2026', short: 'Dec' },
  { value: 'Jan-2027', label: 'January-2027', short: 'Jan' },
  { value: 'Feb-2027', label: 'February-2027', short: 'Feb' },
  { value: 'Mar-2027', label: 'March-2027', short: 'Mar' }
];

// GSTR-2B Month Search Priority function:
// 1. Same working month = 0 (Highest)
// 2. Next/future months = 10 + distance
// 3. Previous/past months = 100 + distance
export function getMonthPriority(targetM: string, currentM: string): number {
  if (targetM === currentM) return 0;
  const currIdx = ALL_FY_MONTHS.findIndex((m) => m.value === currentM);
  const targetIdx = ALL_FY_MONTHS.findIndex((m) => m.value === targetM);
  if (currIdx === -1 || targetIdx === -1) return 999;
  if (targetIdx > currIdx) {
    return 10 + (targetIdx - currIdx);
  }
  return 100 + (currIdx - targetIdx);
}

export const parseRecordMonth = (item: any, defaultMonth: string = 'Oct'): string => {
  const mRaw = (item.m || '').trim();
  const dRaw = (item.d || item.date || '').trim();
  let parsedMonth = mRaw;
  let parsedYear = '2026';
  if (item.fy) {
    const fyMatch = String(item.fy).replace(/[^0-9]/g, '').slice(0, 4);
    if (fyMatch) parsedYear = fyMatch;
  } else if (dRaw) {
    const yrMatch = dRaw.match(/\b(20\d\d)\b/);
    if (yrMatch) parsedYear = yrMatch[1];
  }

  if (!parsedMonth && dRaw) {
    const parts = dRaw.split(/[-/]/);
    const MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        const mNum = parseInt(parts[1], 10);
        if (mNum >= 1 && mNum <= 12) parsedMonth = MONTH_NAMES[mNum];
      } else if (parts[2].length === 4) {
        const mNum = parseInt(parts[1], 10);
        if (mNum >= 1 && mNum <= 12) parsedMonth = MONTH_NAMES[mNum];
      }
    }
  }

  const MONTH_MAP: Record<string, string> = {
    jan: 'Jan', january: 'Jan',
    feb: 'Feb', february: 'Feb',
    mar: 'Mar', march: 'Mar',
    apr: 'Apr', april: 'Apr',
    may: 'May',
    jun: 'Jun', june: 'Jun',
    jul: 'Jul', july: 'Jul',
    aug: 'Aug', august: 'Aug',
    sep: 'Sep', sept: 'Sep', september: 'Sep',
    oct: 'Oct', october: 'Oct',
    nov: 'Nov', november: 'Nov',
    dec: 'Dec', december: 'Dec'
  };
  const mClean = parsedMonth.toLowerCase();
  const stdMonth = MONTH_MAP[mClean] || (mClean ? mClean.charAt(0).toUpperCase() + mClean.slice(1, 3) : defaultMonth);
  if (['Jan', 'Feb', 'Mar'].includes(stdMonth) && parsedYear === '2026') {
    parsedYear = '2027';
  }
  return `${stdMonth}-${parsedYear}`;
};

export const formatAmount = (val: number): string => {
  if (isNaN(val) || val === null || val === undefined) return '0.00';
  const isNeg = val < 0;
  const absVal = Math.abs(val);
  const formatted = absVal.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  return (isNeg ? '-' : '') + formatted;
};

export const formatCurrency = (val: number): string => '₹' + formatAmount(val);

// Helper to render GST amount cells on lavender (#F3EBFE) background
export const renderGstAmount = (
  val: number,
  otherVal: number,
  isBooks: boolean,
  isTotal: boolean = false
) => {
  const formatted = formatAmount(val);

  // Minus amount: bold red for clear readability on lavender
  if (val < 0) {
    return (
      <span style={{ color: '#dc2626', fontWeight: 800 }}>
        {formatted}
      </span>
    );
  }

  // 0.00: subtle neutral grey
  if (val === 0) {
    return (
      <span style={{ color: '#64748b', fontWeight: 600 }}>
        0.00
      </span>
    );
  }

  // Rs 1 tolerance comparison: val is current cell, otherVal is comparison side
  const diff = round2(val - otherVal);
  const hasDiff = Math.abs(diff) > 1.00;

  if (!hasDiff) {
    return (
      <span style={{ color: isTotal ? '#3C1A7A' : '#2A2161', fontWeight: isTotal ? 800 : 600 }}>
        {formatted}
      </span>
    );
  }

  // If other side is zero or negative (e.g., missing invoice or debit note)
  if (otherVal <= 0) {
    return (
      <span style={{ color: isTotal ? '#3C1A7A' : '#2A2161', fontWeight: isTotal ? 800 : 600 }}>
        {formatted}
      </span>
    );
  }

  const isHigher = diff > 1.00;
  const color = isHigher ? '#15803d' : '#dc2626';
  const symbol = isHigher ? '▲' : '▼';
  const sideName = isBooks ? 'Books' : 'Govt 2B';
  const comparisonText = isHigher ? 'zyada hai' : 'kam hai';
  const absDiff = formatAmount(Math.abs(round2(val - otherVal)));
  const tooltip = `${sideName} ₹${absDiff} ${comparisonText}`;

  return (
    <span
      title={tooltip}
      style={{
        color,
        fontWeight: 800,
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'flex-end',
        gap: '4px',
        cursor: 'help'
      }}
    >
      <span>{formatted}</span>
      <span style={{ fontSize: '11px', fontWeight: 900, lineHeight: 1 }}>
        {symbol}
      </span>
    </span>
  );
};


interface ReconciliationViewProps {
  showToast?: (msg: string) => void;
  fy?: string;
  month?: string;
  monthlyData?: any;
}

export default function ReconciliationView({
  showToast = () => { },
  fy = 'FY 2026–27',
  month,
  monthlyData
}: ReconciliationViewProps) {
  // Filters & Search
  const [statusFilter, setStatusFilter] = useState<string>('ALL');
  const [searchQuery, setSearchQuery] = useState<string>('');

  // Working Month derived directly from TopBar HeadingTemplate props
  const workingMonth = useMemo(() => {
    let year = '2026';
    if (fy) {
      const match = fy.replace(/[^0-9]/g, '').slice(0, 4);
      if (match) year = match;
    }

    let m = (month || '').trim();
    if (!m || m.toUpperCase() === 'ALL') {
      m = 'Oct';
    }

    const MONTH_MAP: Record<string, string> = {
      jan: 'Jan', january: 'Jan',
      feb: 'Feb', february: 'Feb',
      mar: 'Mar', march: 'Mar',
      apr: 'Apr', april: 'Apr',
      may: 'May',
      jun: 'Jun', june: 'Jun',
      jul: 'Jul', july: 'Jul',
      aug: 'Aug', august: 'Aug',
      sep: 'Sep', sept: 'Sep', september: 'Sep',
      oct: 'Oct', october: 'Oct',
      nov: 'Nov', november: 'Nov',
      dec: 'Dec', december: 'Dec'
    };

    const stdMonth = MONTH_MAP[m.toLowerCase()] || (m ? m.charAt(0).toUpperCase() + m.slice(1, 3) : 'Oct');
    if (['Jan', 'Feb', 'Mar'].includes(stdMonth) && year === '2026') {
      year = '2027';
    }

    return `${stdMonth}-${year}`;
  }, [month, fy]);

  // Refresh & Lock State
  const [refreshTrigger, setRefreshTrigger] = useState<number>(0);
  const [isRefreshing, setIsRefreshing] = useState<boolean>(false);

  // Month-wise Lock state (persisted in localStorage)
  const [locks, setLocks] = useState<Record<string, {
    isLocked: boolean;
    lockedAt: string;
    lockedBy: string;
    snapshot?: ReconciledInvoice[];
  }>>(() => {
    try {
      const raw = localStorage.getItem('gst_reco_locks_v1');
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });

  const currentLock = locks[workingMonth];
  const isLocked = Boolean(currentLock?.isLocked);

  // Modals state
  const [viewingRecord, setViewingRecord] = useState<ReconciledInvoice | null>(null);
  const [auditingRecord, setAuditingRecord] = useState<ReconciledInvoice | null>(null);
  const [auditFormStatus, setAuditFormStatus] = useState<AuditStatus>('Verified');
  const [auditFormNote, setAuditFormNote] = useState<string>('');
  const [auditEditState, setAuditEditState] = useState<AuditEditState>({
    gstin: '',
    invNo: '',
    supplierName: '',
    date: '',
    taxable: 0,
    igst: 0,
    cgst: 0,
    sgst: 0,
    note: '',
    status: 'Verified'
  });
  const [auditChangeLogs, setAuditChangeLogs] = useState<AuditFieldChangeLog[]>(() => {
    try {
      const raw = localStorage.getItem('gst_audit_change_logs_v1');
      return raw ? JSON.parse(raw) : [];
    } catch {
      return [];
    }
  });
  const [confirmDialog, setConfirmDialog] = useState<{
    title: string;
    message: string;
    confirmLabel: string;
    onConfirm: () => void;
  } | null>(null);

  // Audit notes store
  const [auditNotes, setAuditNotes] = useState<Record<string, AuditEntry>>(() => {
    try {
      const saved = localStorage.getItem('gst_reco_audit_v1');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  // Sync tab navigation if query params exist
  useEffect(() => {
    try {
      const q: { fy?: string; month?: string; tab?: string } = getQueryParams();
      if (q.tab === 'matched') setStatusFilter('Matched');
      else if (q.tab === 'not-match') setStatusFilter('Not Match');
      else if (q.tab === 'not-in-book') setStatusFilter('Not in Purchase Register');
      else if (q.tab === 'not-in-2b') setStatusFilter('Not in GSTR-2B');
      else if (q.tab === 'debit-note') setStatusFilter('Debit Note');
    } catch { }
  }, []);

  // -------------------------------------------------------------
  // Data Loading: Read live books & GSTR-2B or fall back to seeds
  // -------------------------------------------------------------
  const booksRecords = useMemo<RawBookRecord[]>(() => {
    try {
      let rawList: any[] = [];
      const raw = localStorage.getItem('itc_register_v1');
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed) && parsed.length > 0) {
          rawList = parsed;
        }
      }
      if (rawList.length === 0 && Array.isArray((defaultRegisters as any)?.itc)) {
        rawList = (defaultRegisters as any).itc;
      }

      if (rawList.length > 0) {
        return rawList.map((item: any, idx: number) => {
          let rawTaxable = typeof item.t === 'number' ? item.t : (typeof item.taxable === 'number' ? item.taxable : 0);
          let rawIgst = typeof item.i === 'number' ? item.i : (typeof item.igst === 'number' ? item.igst : 0);
          let rawCgst = typeof item.c === 'number' ? item.c : (typeof item.cgst === 'number' ? item.cgst : 0);
          let rawSgst = typeof item.s === 'number' ? item.s : (typeof item.sgst === 'number' ? item.sgst : 0);
          let rawCess = typeof item.cess === 'number' ? item.cess : 0;

          // Intelligent unshift: if rawTaxable is actually Total Invoice Value (equals sum of taxes + taxable)
          const sumShift = rawIgst + rawCgst + rawSgst + rawCess;
          if (sumShift !== 0 && Math.abs(rawTaxable - sumShift) < 10) {
            rawTaxable = rawIgst;
            rawIgst = rawCgst;
            rawCgst = rawSgst;
            rawSgst = rawCess;
            rawCess = 0;
          }

          const isPaise = Math.abs(rawTaxable) >= 100000;
          const monthFormatted = parseRecordMonth(item, 'Oct');

          return {
            id: item.id || `book-${idx}`,
            date: item.d || '',
            invNo: item.invNo || item.n || '',
            gstin: item.g || item.gstin || '',
            supplierName: item.p || item.tradeName || item.party || '',
            taxable: isPaise ? rawTaxable / 100 : rawTaxable,
            igst: isPaise ? rawIgst / 100 : rawIgst,
            cgst: isPaise ? rawCgst / 100 : rawCgst,
            sgst: isPaise ? rawSgst / 100 : rawSgst,
            cess: (rawCess || 0) / (isPaise ? 100 : 1),
            month: monthFormatted
          };
        });
      }
    } catch { }
    return DEFAULT_BOOKS;
  }, [refreshTrigger]);

  const gstr2bRecords = useMemo<Raw2bRecord[]>(() => {
    try {
      let rawList: any[] = [];
      const raw = localStorage.getItem('gstr2b_register_v1');
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed) && parsed.length > 0) {
          rawList = parsed;
        }
      }
      if (rawList.length === 0 && Array.isArray((defaultRegisters as any)?.g2b)) {
        rawList = (defaultRegisters as any).g2b;
      }

      if (rawList.length > 0) {
        return rawList.map((item: any, idx: number) => {
          const rawTaxable = typeof item.taxable === 'number' ? item.taxable : (typeof item.t === 'number' ? item.t : 0);
          const rawIgst = typeof item.igst === 'number' ? item.igst : (typeof item.i === 'number' ? item.i : 0);
          const rawCgst = typeof item.cgst === 'number' ? item.cgst : (typeof item.c === 'number' ? item.c : 0);
          const rawSgst = typeof item.sgst === 'number' ? item.sgst : (typeof item.s === 'number' ? item.s : 0);
          const isPaise = Math.abs(rawTaxable) >= 100000;
          const monthFormatted = parseRecordMonth(item, 'Oct');

          return {
            id: item.id || `g2b-${idx}`,
            date: item.d || '',
            invNo: item.invNo || item.n || '',
            gstin: item.gstin || item.g || '',
            supplierName: item.tradeName || item.p || item.supplierName || '',
            taxable: isPaise ? rawTaxable / 100 : rawTaxable,
            igst: isPaise ? rawIgst / 100 : rawIgst,
            cgst: isPaise ? rawCgst / 100 : rawCgst,
            sgst: isPaise ? rawSgst / 100 : rawSgst,
            cess: (item.cess || 0) / (isPaise ? 100 : 1),
            month: monthFormatted
          };
        });
      }
    } catch { }
    return DEFAULT_GSTR2B;
  }, [refreshTrigger]);

  // Initial mount: ensure fresh live data from server for Books & 2B
  useEffect(() => {
    let isMounted = true;
    async function syncRegisters() {
      try {
        await Promise.all([
          loadRegisterRows(REGISTER_TABS.ITC_PURCHASE),
          loadRegisterRows(REGISTER_TABS.GSTR_2B)
        ]);
        if (isMounted) {
          setRefreshTrigger((prev) => prev + 1);
        }
      } catch (err) {
        console.warn('Live register sync note:', err);
      }
    }
    syncRegisters();
    return () => {
      isMounted = false;
    };
  }, []);

  // -------------------------------------------------------------
  // Full Reconciliation Engine (Books -> 2B Directional Matching)
  // -------------------------------------------------------------
  const computedInvoices = useMemo<ReconciledInvoice[]>(() => {
    const results: ReconciledInvoice[] = [];
    const paired2bIds = new Set<string>();

    // Duplicate check map across all GSTR-2B months
    const gstr2bInvMonthMap = new Map<string, string[]>();
    gstr2bRecords.forEach((g) => {
      const key = `${cleanStr(g.gstin)}_${cleanInv(g.invNo)}`;
      const existing = gstr2bInvMonthMap.get(key) || [];
      if (!existing.includes(g.month)) {
        existing.push(g.month);
      }
      gstr2bInvMonthMap.set(key, existing);
    });

    // Step 0: Filter Books for ONLY the selected Working Month!
    const workingBooks = booksRecords.filter((b) => b.month === workingMonth);

    // Separate Debit Note rows from Books (negative amounts)
    const debitNoteBooks: RawBookRecord[] = [];
    const regularBooks: RawBookRecord[] = [];

    workingBooks.forEach((b) => {
      const isMinus = b.taxable < 0 || b.igst < 0 || b.cgst < 0 || b.sgst < 0;
      if (isMinus) {
        debitNoteBooks.push(b);
      } else {
        regularBooks.push(b);
      }
    });

    // Debit Note records -> Status "Debit Note", never paired, 2B Month "-"
    debitNoteBooks.forEach((b) => {
      const bTotal = round2(b.taxable + b.igst + b.cgst + b.sgst);
      const audit = auditNotes[b.id] || auditNotes[`${cleanStr(b.gstin)}_${cleanInv(b.invNo)}`];

      results.push({
        id: b.id,
        status: 'Debit Note',
        month2b: '-',
        hasDuplicateMonth: false,
        itcMatchText: '-',
        itcMatchCount: 0,
        mismatchedFields: [],
        date: b.date,
        invNo: b.invNo,
        gstin: b.gstin,
        supplierName: b.supplierName,
        bTaxable: b.taxable,
        bIgst: b.igst,
        bCgst: b.cgst,
        bSgst: b.sgst,
        bTotal,
        gTaxable: 0,
        gIgst: 0,
        gCgst: 0,
        gSgst: 0,
        gTotal: 0,
        dTaxable: b.taxable,
        dIgst: b.igst,
        dCgst: b.cgst,
        dSgst: b.sgst,
        dTotal: bTotal,
        isGstinMatch: false,
        isInvMatch: false,
        isTaxableMatch: false,
        isIgstMatch: false,
        isCgstMatch: false,
        isSgstMatch: false,
        auditStatus: audit?.status || 'Pending',
        auditNote: audit?.note || '',
        auditUpdatedAt: audit?.updatedAt
      });
    });

    // Sort candidate GSTR-2B records by month search order:
    // 1) Working month
    // 2) Next/future months
    // 3) Previous/past months
    const sorted2bRecords = [...gstr2bRecords].sort(
      (a, b) => getMonthPriority(a.month, workingMonth) - getMonthPriority(b.month, workingMonth)
    );

    // Step 1: Exact Match on ALL 6 fields (Matched)
    const unpairedBooksStage1: RawBookRecord[] = [];

    regularBooks.forEach((b) => {
      const bGstinClean = cleanStr(b.gstin);
      const bInvClean = cleanInv(b.invNo);

      let matched2b: Raw2bRecord | undefined;

      for (const g of sorted2bRecords) {
        if (paired2bIds.has(g.id)) continue;

        const gGstinClean = cleanStr(g.gstin);
        const gInvClean = cleanInv(g.invNo);

        const gstinMatch = bGstinClean === gGstinClean;
        const invMatch = bInvClean === gInvClean;
        const taxMatch = Math.abs(round2(b.taxable - g.taxable)) <= 1.00;
        const igstMatch = Math.abs(round2(b.igst - g.igst)) <= 1.00;
        const cgstMatch = Math.abs(round2(b.cgst - g.cgst)) <= 1.00;
        const sgstMatch = Math.abs(round2(b.sgst - g.sgst)) <= 1.00;

        if (gstinMatch && invMatch && taxMatch && igstMatch && cgstMatch && sgstMatch) {
          matched2b = g;
          break;
        }
      }

      if (matched2b) {
        paired2bIds.add(matched2b.id);
        const bTotal = round2(b.taxable + b.igst + b.cgst + b.sgst);
        const gTotal = round2(matched2b.taxable + matched2b.igst + matched2b.cgst + matched2b.sgst);

        const dupMonths = gstr2bInvMonthMap.get(`${bGstinClean}_${bInvClean}`) || [];
        const hasDup = dupMonths.length > 1;
        const audit = auditNotes[b.id] || auditNotes[`${bGstinClean}_${bInvClean}`];

        results.push({
          id: b.id,
          status: 'Matched',
          month2b: matched2b.month || workingMonth,
          hasDuplicateMonth: hasDup,
          duplicateMonthsList: dupMonths,
          itcMatchText: '6/6 Match',
          itcMatchCount: 6,
          mismatchedFields: [],
          date: b.date,
          invNo: b.invNo,
          gstin: b.gstin,
          supplierName: b.supplierName,
          bTaxable: b.taxable,
          bIgst: b.igst,
          bCgst: b.cgst,
          bSgst: b.sgst,
          bTotal,
          gTaxable: matched2b.taxable,
          gIgst: matched2b.igst,
          gCgst: matched2b.cgst,
          gSgst: matched2b.sgst,
          gTotal,
          dTaxable: round2(b.taxable - matched2b.taxable),
          dIgst: round2(b.igst - matched2b.igst),
          dCgst: round2(b.cgst - matched2b.cgst),
          dSgst: round2(b.sgst - matched2b.sgst),
          dTotal: round2(bTotal - gTotal),
          gDate: matched2b.date,
          gInvNo: matched2b.invNo,
          gGstin: matched2b.gstin,
          gTradeName: matched2b.supplierName,
          isGstinMatch: true,
          isInvMatch: true,
          isTaxableMatch: true,
          isIgstMatch: true,
          isCgstMatch: true,
          isSgstMatch: true,
          auditStatus: audit?.status || 'Pending',
          auditNote: audit?.note || '',
          auditUpdatedAt: audit?.updatedAt
        });
      } else {
        unpairedBooksStage1.push(b);
      }
    });

    // Step 2: Key Match (GST No + Invoice No)
    const unpairedBooksStage2: RawBookRecord[] = [];

    unpairedBooksStage1.forEach((b) => {
      const bGstinClean = cleanStr(b.gstin);
      const bInvClean = cleanInv(b.invNo);

      let paired2b: Raw2bRecord | undefined;

      for (const g of sorted2bRecords) {
        if (paired2bIds.has(g.id)) continue;
        if (bGstinClean === cleanStr(g.gstin) && bInvClean === cleanInv(g.invNo)) {
          paired2b = g;
          break;
        }
      }

      if (paired2b) {
        paired2bIds.add(paired2b.id);
        processMismatchedPair(b, paired2b);
      } else {
        unpairedBooksStage2.push(b);
      }
    });

    // Step 3: Fuzzy Match (GST No + Taxable within ₹1)
    const unpairedBooksStage3: RawBookRecord[] = [];

    unpairedBooksStage2.forEach((b) => {
      const bGstinClean = cleanStr(b.gstin);

      let paired2b: Raw2bRecord | undefined;

      for (const g of sorted2bRecords) {
        if (paired2bIds.has(g.id)) continue;
        if (bGstinClean === cleanStr(g.gstin) && Math.abs(round2(b.taxable - g.taxable)) <= 1.00) {
          paired2b = g;
          break;
        }
      }

      if (paired2b) {
        paired2bIds.add(paired2b.id);
        processMismatchedPair(b, paired2b);
      } else {
        unpairedBooksStage3.push(b);
      }
    });

    // Helper for pairing mismatched rows
    function processMismatchedPair(b: RawBookRecord, g: Raw2bRecord) {
      const bGstinClean = cleanStr(b.gstin);
      const bInvClean = cleanInv(b.invNo);
      const gGstinClean = cleanStr(g.gstin);
      const gInvClean = cleanInv(g.invNo);

      const isGstinMatch = bGstinClean === gGstinClean;
      const isInvMatch = bInvClean === gInvClean;
      const isTaxableMatch = Math.abs(round2(b.taxable - g.taxable)) <= 1.00;
      const isIgstMatch = Math.abs(round2(b.igst - g.igst)) <= 1.00;
      const isCgstMatch = Math.abs(round2(b.cgst - g.cgst)) <= 1.00;
      const isSgstMatch = Math.abs(round2(b.sgst - g.sgst)) <= 1.00;

      const mismatches: string[] = [];
      let matchCount = 0;

      if (isGstinMatch) matchCount++; else mismatches.push('GST No Not Match');
      if (isInvMatch) matchCount++; else mismatches.push('Invoice No Not Match');
      if (isTaxableMatch) matchCount++; else mismatches.push('Taxable Not Match');
      if (isIgstMatch) matchCount++; else mismatches.push('IGST Not Match');
      if (isCgstMatch) matchCount++; else mismatches.push('CGST Not Match');
      if (isSgstMatch) matchCount++; else mismatches.push('SGST Not Match');

      const bTotal = round2(b.taxable + b.igst + b.cgst + b.sgst);
      const gTotal = round2(g.taxable + g.igst + g.cgst + g.sgst);

      const dupMonths = gstr2bInvMonthMap.get(`${bGstinClean}_${bInvClean}`) || [];
      const hasDup = dupMonths.length > 1;
      const audit = auditNotes[b.id] || auditNotes[`${bGstinClean}_${bInvClean}`];

      results.push({
        id: b.id,
        status: 'Not Match',
        month2b: g.month || workingMonth,
        hasDuplicateMonth: hasDup,
        duplicateMonthsList: dupMonths,
        itcMatchText: `${matchCount}/6 Match`,
        itcMatchCount: matchCount,
        mismatchedFields: mismatches,
        date: b.date,
        invNo: b.invNo,
        gstin: b.gstin,
        supplierName: b.supplierName,
        bTaxable: b.taxable,
        bIgst: b.igst,
        bCgst: b.cgst,
        bSgst: b.sgst,
        bTotal,
        gTaxable: g.taxable,
        gIgst: g.igst,
        gCgst: g.cgst,
        gSgst: g.sgst,
        gTotal,
        dTaxable: round2(b.taxable - g.taxable),
        dIgst: round2(b.igst - g.igst),
        dCgst: round2(b.cgst - g.cgst),
        dSgst: round2(b.sgst - g.sgst),
        dTotal: round2(bTotal - gTotal),
        gDate: g.date,
        gInvNo: g.invNo,
        gGstin: g.gstin,
        gTradeName: g.supplierName,
        isGstinMatch,
        isInvMatch,
        isTaxableMatch,
        isIgstMatch,
        isCgstMatch,
        isSgstMatch,
        auditStatus: audit?.status || 'Pending',
        auditNote: audit?.note || '',
        auditUpdatedAt: audit?.updatedAt
      });
    }

    // Step 4: Unpaired Books -> "Not in GSTR-2B"
    unpairedBooksStage3.forEach((b) => {
      const bTotal = round2(b.taxable + b.igst + b.cgst + b.sgst);
      const audit = auditNotes[b.id] || auditNotes[`${cleanStr(b.gstin)}_${cleanInv(b.invNo)}`];

      results.push({
        id: b.id,
        status: 'Not in GSTR-2B',
        month2b: '-',
        hasDuplicateMonth: false,
        itcMatchText: '-',
        itcMatchCount: 0,
        mismatchedFields: ['Missing in GSTR-2B across all months'],
        date: b.date,
        invNo: b.invNo,
        gstin: b.gstin,
        supplierName: b.supplierName,
        bTaxable: b.taxable,
        bIgst: b.igst,
        bCgst: b.cgst,
        bSgst: b.sgst,
        bTotal,
        gTaxable: 0,
        gIgst: 0,
        gCgst: 0,
        gSgst: 0,
        gTotal: 0,
        dTaxable: b.taxable,
        dIgst: b.igst,
        dCgst: b.cgst,
        dSgst: b.sgst,
        dTotal: bTotal,
        isGstinMatch: false,
        isInvMatch: false,
        isTaxableMatch: false,
        isIgstMatch: false,
        isCgstMatch: false,
        isSgstMatch: false,
        auditStatus: audit?.status || 'Pending',
        auditNote: audit?.note || '',
        auditUpdatedAt: audit?.updatedAt
      });
    });

    // Step 5: Unpaired GSTR-2B Records for WORKING MONTH ONLY -> "Not in Purchase Register"
    gstr2bRecords.forEach((g) => {
      if (g.month === workingMonth && !paired2bIds.has(g.id)) {
        const gTotal = round2(g.taxable + g.igst + g.cgst + g.sgst);
        const audit = auditNotes[g.id] || auditNotes[`${cleanStr(g.gstin)}_${cleanInv(g.invNo)}`];

        results.push({
          id: g.id,
          status: 'Not in Purchase Register',
          month2b: g.month || workingMonth,
          hasDuplicateMonth: false,
          itcMatchText: '-',
          itcMatchCount: 0,
          mismatchedFields: ['Missing in Purchase Register'],
          date: g.date,
          invNo: g.invNo,
          gstin: g.gstin,
          supplierName: g.supplierName,
          bTaxable: 0,
          bIgst: 0,
          bCgst: 0,
          bSgst: 0,
          bTotal: 0,
          gTaxable: g.taxable,
          gIgst: g.igst,
          gCgst: g.cgst,
          gSgst: g.sgst,
          gTotal,
          dTaxable: round2(0 - g.taxable),
          dIgst: round2(0 - g.igst),
          dCgst: round2(0 - g.cgst),
          dSgst: round2(0 - g.sgst),
          dTotal: round2(0 - gTotal),
          gDate: g.date,
          gInvNo: g.invNo,
          gGstin: g.gstin,
          gTradeName: g.supplierName,
          isGstinMatch: false,
          isInvMatch: false,
          isTaxableMatch: false,
          isIgstMatch: false,
          isCgstMatch: false,
          isSgstMatch: false,
          auditStatus: audit?.status || 'Pending',
          auditNote: audit?.note || '',
          auditUpdatedAt: audit?.updatedAt
        });
      }
    });

    return results;
  }, [booksRecords, gstr2bRecords, workingMonth, auditNotes]);

  // If month is locked and snapshot exists, freeze matching result, status, amount:
  const reconciledInvoices = useMemo<ReconciledInvoice[]>(() => {
    if (isLocked && currentLock?.snapshot && currentLock.snapshot.length > 0) {
      return currentLock.snapshot;
    }
    return computedInvoices;
  }, [isLocked, currentLock, computedInvoices]);

  // Working Month Books Count (for verification check)
  const booksCount = useMemo(() => {
    if (isLocked && currentLock?.snapshot && currentLock.snapshot.length > 0) {
      return currentLock.snapshot.filter((r) => r.status !== 'Not in Purchase Register').length;
    }
    return booksRecords.filter((b) => b.month === workingMonth).length;
  }, [isLocked, currentLock, booksRecords, workingMonth]);

  // -------------------------------------------------------------
  // Filtered Display List
  // -------------------------------------------------------------
  const filteredRecords = useMemo(() => {
    return reconciledInvoices.filter((r) => {
      // If statusFilter is 'ALL', show all working month Books invoices (excluding 'Not in Purchase Register')
      if (statusFilter === 'ALL') {
        if (r.status === 'Not in Purchase Register') return false;
      } else if (r.status !== statusFilter) {
        return false;
      }

      // Search query filter
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const matchesInv = (r.invNo || '').toLowerCase().includes(q) || (r.gInvNo || '').toLowerCase().includes(q);
        const matchesGstin = (r.gstin || '').toLowerCase().includes(q) || (r.gGstin || '').toLowerCase().includes(q);
        const matchesSupplier = (r.supplierName || '').toLowerCase().includes(q) || (r.gTradeName || '').toLowerCase().includes(q);
        if (!matchesInv && !matchesGstin && !matchesSupplier) {
          return false;
        }
      }

      return true;
    });
  }, [reconciledInvoices, statusFilter, searchQuery]);

  // -------------------------------------------------------------
  // Summary KPIs (6 Cards)
  // -------------------------------------------------------------
  const summaryKpis = useMemo(() => {
    // Only Books invoices (Matched, Not Match, Not in GSTR-2B, Debit Note) make up Total Invoices
    const booksInvoices = reconciledInvoices.filter((r) => r.status !== 'Not in Purchase Register');
    const totalCount = booksInvoices.length;
    const totalItc = booksInvoices.reduce((acc, r) => acc + (r.bIgst + r.bCgst + r.bSgst), 0);

    const matchedList = reconciledInvoices.filter((r) => r.status === 'Matched');
    const matchedCount = matchedList.length;
    const matchedItc = matchedList.reduce((acc, r) => acc + (r.bIgst + r.bCgst + r.bSgst), 0);

    const notMatchList = reconciledInvoices.filter((r) => r.status === 'Not Match');
    const notMatchCount = notMatchList.length;
    const notMatchItc = notMatchList.reduce((acc, r) => acc + (r.bIgst + r.bCgst + r.bSgst), 0);

    const notIn2bList = reconciledInvoices.filter((r) => r.status === 'Not in GSTR-2B');
    const notIn2bCount = notIn2bList.length;
    const notIn2bItc = notIn2bList.reduce((acc, r) => acc + (r.bIgst + r.bCgst + r.bSgst), 0);

    // ONLY working month unpaired 2B records
    const notInBookList = reconciledInvoices.filter((r) => r.status === 'Not in Purchase Register');
    const notInBookCount = notInBookList.length;
    const notInBookItc = notInBookList.reduce((acc, r) => acc + (r.gIgst + r.gCgst + r.gSgst), 0);

    const debitNoteList = reconciledInvoices.filter((r) => r.status === 'Debit Note');
    const debitNoteCount = debitNoteList.length;
    const debitNoteItc = debitNoteList.reduce((acc, r) => acc + (r.bIgst + r.bCgst + r.bSgst), 0);

    return {
      totalCount,
      totalItc,
      matchedCount,
      matchedItc,
      notMatchCount,
      notMatchItc,
      notIn2bCount,
      notIn2bItc,
      notInBookCount,
      notInBookItc,
      debitNoteCount,
      debitNoteItc
    };
  }, [reconciledInvoices]);

  // Broadcast live reconciliation state for external cards & views (one-way sync)
  useEffect(() => {
    try {
      const payload = {
        workingMonth,
        updatedAt: Date.now(),
        summary: {
          totalCount: summaryKpis.totalCount,
          matchedCount: summaryKpis.matchedCount,
          notMatchCount: summaryKpis.notMatchCount,
          notIn2bCount: summaryKpis.notIn2bCount,
          notInBookCount: summaryKpis.notInBookCount,
          debitNoteCount: summaryKpis.debitNoteCount
        },
        invoices: reconciledInvoices.map((r) => ({
          id: r.id,
          status: r.status,
          month2b: r.month2b,
          date: r.date,
          invNo: r.invNo,
          gstin: r.gstin,
          supplierName: r.supplierName,
          taxable: r.status === 'Not in Purchase Register' ? r.gTaxable : r.bTaxable,
          igst: r.status === 'Not in Purchase Register' ? r.gIgst : r.bIgst,
          cgst: r.status === 'Not in Purchase Register' ? r.gCgst : r.bCgst,
          sgst: r.status === 'Not in Purchase Register' ? r.gSgst : r.bSgst,
          total: r.status === 'Not in Purchase Register' ? r.gTotal : r.bTotal
        }))
      };
      localStorage.setItem('gst_reco_active_data_v1', JSON.stringify(payload));
      window.dispatchEvent(new CustomEvent('reco-data-updated', { detail: payload }));
    } catch {}
  }, [reconciledInvoices, summaryKpis, workingMonth]);

  // Verification Check: Total Invoices = Matched + Not Match + Not in 2B + Debit Note = Books Count
  const isCheckPassed = useMemo(() => {
    const sum = summaryKpis.matchedCount + summaryKpis.notMatchCount + summaryKpis.notIn2bCount + summaryKpis.debitNoteCount;
    return sum === booksCount && summaryKpis.totalCount === booksCount;
  }, [summaryKpis, booksCount]);

  // -------------------------------------------------------------
  // Table Totals across filtered records
  // -------------------------------------------------------------
  const tableTotals = useMemo(() => {
    return filteredRecords.reduce(
      (acc, r) => ({
        bTaxable: acc.bTaxable + r.bTaxable,
        bIgst: acc.bIgst + r.bIgst,
        bCgst: acc.bCgst + r.bCgst,
        bSgst: acc.bSgst + r.bSgst,
        bTotal: acc.bTotal + r.bTotal,
        gTaxable: acc.gTaxable + r.gTaxable,
        gIgst: acc.gIgst + r.gIgst,
        gCgst: acc.gCgst + r.gCgst,
        gSgst: acc.gSgst + r.gSgst,
        gTotal: acc.gTotal + r.gTotal,
        dTaxable: acc.dTaxable + r.dTaxable,
        dIgst: acc.dIgst + r.dIgst,
        dCgst: acc.dCgst + r.dCgst,
        dSgst: acc.dSgst + r.dSgst,
        dTotal: acc.dTotal + r.dTotal
      }),
      {
        bTaxable: 0,
        bIgst: 0,
        bCgst: 0,
        bSgst: 0,
        bTotal: 0,
        gTaxable: 0,
        gIgst: 0,
        gCgst: 0,
        gSgst: 0,
        gTotal: 0,
        dTaxable: 0,
        dIgst: 0,
        dCgst: 0,
        dSgst: 0,
        dTotal: 0
      }
    );
  }, [filteredRecords]);

  // -------------------------------------------------------------
  // Audit Actions
  // -------------------------------------------------------------
  const handleOpenAudit = (rec: ReconciledInvoice) => {
    if (isLocked) {
      showToast('Reconciliation locked hai! Audit edit nahi ho sakta.');
      return;
    }
    setAuditingRecord(rec);
    setAuditEditState({
      gstin: rec.gstin || '',
      invNo: rec.invNo || '',
      supplierName: rec.supplierName || '',
      date: rec.date || '',
      taxable: rec.bTaxable || 0,
      igst: rec.bIgst || 0,
      cgst: rec.bCgst || 0,
      sgst: rec.bSgst || 0,
      note: rec.auditNote || '',
      status: rec.auditStatus || 'Verified'
    });
    setAuditFormStatus(rec.auditStatus || 'Verified');
    setAuditFormNote(rec.auditNote || '');
  };

  const handleSaveAudit = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (isLocked) {
      showToast('Reconciliation locked hai! Changes save nahi honge.');
      return;
    }
    if (!auditingRecord) return;

    const curTaxable = Number(auditEditState.taxable) || 0;
    const curIgst = Number(auditEditState.igst) || 0;
    const curCgst = Number(auditEditState.cgst) || 0;
    const curSgst = Number(auditEditState.sgst) || 0;
    const curTotal = round2(curTaxable + curIgst + curCgst + curSgst);

    // 1. Identify changes for audit change logs
    const changes: { field: string; oldVal: string; newVal: string }[] = [];
    if (cleanStr(auditingRecord.gstin) !== cleanStr(auditEditState.gstin)) {
      changes.push({ field: 'Supplier GSTIN', oldVal: auditingRecord.gstin || '-', newVal: auditEditState.gstin });
    }
    if (cleanInv(auditingRecord.invNo) !== cleanInv(auditEditState.invNo)) {
      changes.push({ field: 'Invoice No.', oldVal: auditingRecord.invNo || '-', newVal: auditEditState.invNo });
    }
    if (cleanStr(auditingRecord.supplierName) !== cleanStr(auditEditState.supplierName)) {
      changes.push({ field: 'Supplier Name', oldVal: auditingRecord.supplierName || '-', newVal: auditEditState.supplierName });
    }
    if (auditEditState.date && auditingRecord.date !== auditEditState.date) {
      changes.push({ field: 'Invoice Date', oldVal: auditingRecord.date || '-', newVal: auditEditState.date });
    }
    if (Math.abs(round2(auditingRecord.bTaxable - curTaxable)) > 0.01) {
      changes.push({ field: 'Taxable Value', oldVal: formatAmount(auditingRecord.bTaxable), newVal: formatAmount(curTaxable) });
    }
    if (Math.abs(round2(auditingRecord.bIgst - curIgst)) > 0.01) {
      changes.push({ field: 'IGST', oldVal: formatAmount(auditingRecord.bIgst), newVal: formatAmount(curIgst) });
    }
    if (Math.abs(round2(auditingRecord.bCgst - curCgst)) > 0.01) {
      changes.push({ field: 'CGST', oldVal: formatAmount(auditingRecord.bCgst), newVal: formatAmount(curCgst) });
    }
    if (Math.abs(round2(auditingRecord.bSgst - curSgst)) > 0.01) {
      changes.push({ field: 'SGST', oldVal: formatAmount(auditingRecord.bSgst), newVal: formatAmount(curSgst) });
    }

    const nowTime = new Date().toLocaleString('en-IN', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });

    if (changes.length > 0) {
      const newLogs: AuditFieldChangeLog[] = changes.map(c => ({
        id: `log-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        invNo: auditEditState.invNo || auditingRecord.invNo,
        gstin: auditEditState.gstin || auditingRecord.gstin,
        field: c.field,
        oldValue: c.oldVal,
        newValue: c.newVal,
        user: 'User',
        timestamp: nowTime
      }));
      const updatedLogs = [...newLogs, ...auditChangeLogs];
      setAuditChangeLogs(updatedLogs);
      try {
        localStorage.setItem('gst_audit_change_logs_v1', JSON.stringify(updatedLogs));
      } catch {}
    }

    // 2. Update Books Data in localStorage and database
    let rawList: any[] = [];
    try {
      const saved = localStorage.getItem('itc_register_v1');
      if (saved) {
        rawList = JSON.parse(saved);
      }
    } catch {}
    if (!Array.isArray(rawList) || rawList.length === 0) {
      if (Array.isArray((defaultRegisters as any)?.itc)) {
        rawList = [...(defaultRegisters as any).itc];
      }
    }

    const isPaise = rawList.some(
      (b) => (b.taxable && Math.abs(b.taxable) >= 100000) || (b.t && Math.abs(b.t) >= 100000)
    );

    const bGstinClean = cleanStr(auditingRecord.gstin);
    const bInvClean = cleanInv(auditingRecord.invNo);

    let foundIdx = rawList.findIndex(r => r.id === auditingRecord.id);
    if (foundIdx === -1) {
      foundIdx = rawList.findIndex(
        r => cleanStr(r.g || r.gstin) === bGstinClean && cleanInv(r.n || r.invNo) === bInvClean
      );
    }

    const tVal = isPaise ? Math.round(curTaxable * 100) : curTaxable;
    const iVal = isPaise ? Math.round(curIgst * 100) : curIgst;
    const cVal = isPaise ? Math.round(curCgst * 100) : curCgst;
    const sVal = isPaise ? Math.round(curSgst * 100) : curSgst;
    const totalVal = isPaise ? Math.round(curTotal * 100) : curTotal;

    let targetBookRow: any;
    if (foundIdx !== -1) {
      targetBookRow = {
        ...rawList[foundIdx],
        g: auditEditState.gstin,
        gstin: auditEditState.gstin,
        gstNo: auditEditState.gstin,
        n: auditEditState.invNo,
        invNo: auditEditState.invNo,
        p: auditEditState.supplierName,
        tradeName: auditEditState.supplierName,
        party: auditEditState.supplierName,
        partyName: auditEditState.supplierName,
        supplierName: auditEditState.supplierName,
        d: auditEditState.date,
        t: tVal,
        taxable: tVal,
        taxableValue: tVal,
        amount: tVal,
        i: iVal,
        igst: iVal,
        c: cVal,
        cgst: cVal,
        s: sVal,
        sgst: sVal,
        invVal: totalVal,
        totalInvoiceValue: totalVal
      };
      rawList[foundIdx] = targetBookRow;
    } else {
      targetBookRow = {
        id: `book-audit-${Date.now()}`,
        fy: 'FY 2026–27',
        q: 'Q3 (Oct-Dec)',
        m: 'Oct',
        d: auditEditState.date || auditingRecord.gDate || auditingRecord.date || '2026-10-15',
        p: auditEditState.supplierName,
        tradeName: auditEditState.supplierName,
        party: auditEditState.supplierName,
        partyName: auditEditState.supplierName,
        supplierName: auditEditState.supplierName,
        g: auditEditState.gstin,
        gstin: auditEditState.gstin,
        gstNo: auditEditState.gstin,
        n: auditEditState.invNo,
        invNo: auditEditState.invNo,
        t: tVal,
        taxable: tVal,
        taxableValue: tVal,
        amount: tVal,
        i: iVal,
        igst: iVal,
        c: cVal,
        cgst: cVal,
        s: sVal,
        sgst: sVal,
        invVal: totalVal,
        totalInvoiceValue: totalVal,
        month: workingMonth
      };
      rawList.push(targetBookRow);
    }

    try {
      localStorage.setItem('itc_register_v1', JSON.stringify(rawList));
    } catch {}

    try {
      await saveRegisterRow(REGISTER_TABS.ITC_PURCHASE, targetBookRow);
    } catch {}

    // 3. Check if all 6 fields match 2B
    const diffTaxable = round2(curTaxable - auditingRecord.gTaxable);
    const diffIgst = round2(curIgst - auditingRecord.gIgst);
    const diffCgst = round2(curCgst - auditingRecord.gCgst);
    const diffSgst = round2(curSgst - auditingRecord.gSgst);
    const liveGstinMatch = cleanStr(auditEditState.gstin) === cleanStr(auditingRecord.gGstin || auditingRecord.gstin);
    const liveInvMatch = cleanInv(auditEditState.invNo) === cleanInv(auditingRecord.gInvNo || auditingRecord.invNo);
    const liveTaxMatch = Math.abs(diffTaxable) <= 1.00;
    const liveIgstMatch = Math.abs(diffIgst) <= 1.00;
    const liveCgstMatch = Math.abs(diffCgst) <= 1.00;
    const liveSgstMatch = Math.abs(diffSgst) <= 1.00;

    const liveAllMatch = liveGstinMatch && liveInvMatch && liveTaxMatch && liveIgstMatch && liveCgstMatch && liveSgstMatch;

    const finalStatus: AuditStatus = liveAllMatch ? 'Verified' : auditEditState.status;
    const auditKey = auditingRecord.id;
    const updatedAudit: AuditEntry = {
      status: finalStatus,
      note: auditEditState.note.trim() || (changes.length > 0 ? `Audited and updated Books data (${changes.map(c => c.field).join(', ')})` : 'Audited and verified'),
      updatedAt: nowTime
    };

    const newNotes = {
      ...auditNotes,
      [auditKey]: updatedAudit,
      [`${cleanStr(auditEditState.gstin)}_${cleanInv(auditEditState.invNo)}`]: updatedAudit
    };
    setAuditNotes(newNotes);
    try {
      localStorage.setItem('gst_reco_audit_v1', JSON.stringify(newNotes));
    } catch {}

    // 4. Trigger UI refresh across Reconciliation and ITC Register
    setRefreshTrigger(prev => prev + 1);
    window.dispatchEvent(new CustomEvent('register-synced', { detail: { tab: REGISTER_TABS.ITC_PURCHASE } }));
    window.dispatchEvent(new Event('storage'));

    setAuditingRecord(null);
    showToast(liveAllMatch ? `Audit saved! ${auditEditState.invNo} is now Matched (ITC Verified).` : `Audit changes saved for ${auditEditState.invNo}.`);
  };

  // -------------------------------------------------------------
  // Refresh & Lock Actions
  // -------------------------------------------------------------
  const handleRefresh = async () => {
    if (isLocked) {
      showToast('Reconciliation locked hai! Is month me data refresh nahi ho sakta.');
      return;
    }
    setIsRefreshing(true);
    try {
      await Promise.all([
        loadRegisterRows(REGISTER_TABS.ITC_PURCHASE),
        loadRegisterRows(REGISTER_TABS.GSTR_2B)
      ]);
      setRefreshTrigger((prev) => prev + 1);
      showToast('Data refreshed');
    } catch {
      setRefreshTrigger((prev) => prev + 1);
      showToast('Data refreshed');
    } finally {
      setIsRefreshing(false);
    }
  };

  const handleToggleLock = () => {
    const userName =
      localStorage.getItem('username') ||
      localStorage.getItem('user_name') ||
      'Admin';
    const now = new Date();
    const dateFormatted = now.toLocaleString('en-IN', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });

    if (!isLocked) {
      setConfirmDialog({
        title: 'Lock Sheet Confirmation',
        message: `Kya aap ${workingMonth} ke records ko lock karna chahte hain? Lock hone par edit/delete nahi ho sakega.`,
        confirmLabel: 'Lock Sheet',
        onConfirm: () => {
          const newLocks = {
            ...locks,
            [workingMonth]: {
              isLocked: true,
              lockedAt: dateFormatted,
              lockedBy: userName,
              snapshot: computedInvoices
            }
          };
          setLocks(newLocks);
          try {
            localStorage.setItem('gst_reco_locks_v1', JSON.stringify(newLocks));

            const logsRaw = localStorage.getItem('gst_reco_lock_logs_v1');
            const logs = logsRaw ? JSON.parse(logsRaw) : [];
            logs.push({
              id: `log-${Date.now()}`,
              month: workingMonth,
              action: 'LOCK',
              timestamp: now.toISOString(),
              displayTime: dateFormatted,
              user: userName
            });
            localStorage.setItem('gst_reco_lock_logs_v1', JSON.stringify(logs));
          } catch { }

          showToast(`${workingMonth} reconciliation locked successfully.`);
        }
      });
    } else {
      setConfirmDialog({
        title: 'Unlock Sheet Confirmation',
        message: `Kya aap ${workingMonth} reconciliation sheet ko unlock karna chahte hain? Unlock ke baad changes kiye ja sakenge.`,
        confirmLabel: 'Unlock Sheet',
        onConfirm: () => {
          const newLocks = {
            ...locks,
            [workingMonth]: {
              isLocked: false,
              lockedAt: '',
              lockedBy: userName,
              snapshot: undefined
            }
          };
          setLocks(newLocks);
          try {
            localStorage.setItem('gst_reco_locks_v1', JSON.stringify(newLocks));

            const logsRaw = localStorage.getItem('gst_reco_lock_logs_v1');
            const logs = logsRaw ? JSON.parse(logsRaw) : [];
            logs.push({
              id: `log-${Date.now()}`,
              month: workingMonth,
              action: 'UNLOCK',
              timestamp: now.toISOString(),
              displayTime: dateFormatted,
              user: userName
            });
            localStorage.setItem('gst_reco_lock_logs_v1', JSON.stringify(logs));
          } catch { }

          showToast(`${workingMonth} reconciliation unlocked.`);
        }
      });
    }
  };

  // -------------------------------------------------------------
  // Export to Excel (All 23 Columns)
  // -------------------------------------------------------------
  const handleExportExcel = () => {
    try {
      const exportRows = filteredRecords.map((r) => ({
        'Status': r.status,
        '2B Month': r.month2b,
        'Date': r.date,
        'Invoice No.': r.invNo,
        'Supplier GSTIN': r.gstin,
        'Supplier Name': r.supplierName,
        'Books: Taxable': r.bTaxable,
        'Books: IGST': r.bIgst,
        'Books: CGST': r.bCgst,
        'Books: SGST': r.bSgst,
        'Govt 2B: Taxable': r.gTaxable,
        'Govt 2B: IGST': r.gIgst,
        'Govt 2B: CGST': r.gCgst,
        'Govt 2B: SGST': r.gSgst,
        'Audit Status': r.auditStatus || 'Pending',
        'Audit Note': r.auditNote || ''
      }));

      const ws = XLSX.utils.json_to_sheet(exportRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Books_Reco_2B');
      const filename = `Books_Reconciliation_with_GSTR-2B_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(wb, filename);
      showToast('Excel report downloaded successfully!');
    } catch (e) {
      console.error(e);
      showToast('Failed to export Excel.');
    }
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <section id="v-rec" className="reco-page-container">
      {/* ======================================================== */}
      {/* 1. Summary Cards (Count aur ITC Amount)                   */}
      {/* ======================================================== */}
          <div className="reco-kpis-grid">
            {/* Total Invoices */}
            <div className="reco-kpi-card">
              <div className="reco-kpi-title" style={{ color: '#350b5e' }}>Total Invoices</div>
              <div className="reco-kpi-count" style={{ color: '#350b5e' }}>{summaryKpis.totalCount}</div>
            </div>

            {/* Matched (Hara) */}
            <div className="reco-kpi-card">
              <div className="reco-kpi-title" style={{ color: '#15803d' }}>Matched Invoices</div>
              <div className="reco-kpi-count" style={{ color: '#15803d' }}>{summaryKpis.matchedCount}</div>
            </div>

            {/* Not Match (Laal) */}
            <div className="reco-kpi-card">
              <div className="reco-kpi-title" style={{ color: '#dc2626' }}>Not Match</div>
              <div className="reco-kpi-count" style={{ color: '#dc2626' }}>{summaryKpis.notMatchCount}</div>
            </div>

            {/* Not in GSTR-2B (Peela) */}
            <div className="reco-kpi-card">
              <div className="reco-kpi-title" style={{ color: '#b45309' }}>Not in GSTR-2B</div>
              <div className="reco-kpi-count" style={{ color: '#b45309' }}>{summaryKpis.notIn2bCount}</div>
            </div>

            {/* Not in Purchase Register (Neela) */}
            <div className="reco-kpi-card">
              <div className="reco-kpi-title" style={{ color: '#1d4ed8' }}>Not in Purchase Register</div>
              <div className="reco-kpi-count" style={{ color: '#1d4ed8' }}>{summaryKpis.notInBookCount}</div>
            </div>

            {/* Debit Note (Orange) */}
            <div className="reco-kpi-card">
              <div className="reco-kpi-title" style={{ color: '#c2410c' }}>Debit Note</div>
              <div className="reco-kpi-count" style={{ color: '#c2410c' }}>{summaryKpis.debitNoteCount}</div>
            </div>
          </div>

      {/* ======================================================== */}
      {/* 2. Unified Sheet Card (Header, Filters, Table)           */}
      {/* ======================================================== */}
      <div className="card reco-card">
        {/* Top Header Section */}
        <div className="reco-header-card">
          <div className="reco-header-left">
            <div className="reco-icon-badge">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
              </svg>
              <div className="reco-badge-check">✓</div>
            </div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                <h1 className="reco-title">
                  Books Reconciliation <span style={{ fontWeight: 400, color: '#64748b' }}>with GSTR-2B</span>
                </h1>
                {isLocked && (
                  <span className="reco-locked-badge" title={`Locked by ${currentLock?.lockedBy || 'Admin'} on ${currentLock?.lockedAt}`}>
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                      <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
                      <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
                    </svg>
                    Locked ({currentLock?.lockedAt || 'Locked'})
                  </span>
                )}
              </div>
              <p className="reco-subtitle">
                Systematic 6-field invoice verification between Purchase Books & GSTR-2B across all return periods ({workingMonth})
              </p>
            </div>
          </div>

          <div className="reco-header-right">
            {/* Search Input */}
            <div className="reco-search-box">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" className="reco-search-icon">
                <circle cx="11" cy="11" r="8"></circle>
                <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
              </svg>
              <input
                type="text"
                placeholder="Search Supplier, GSTIN, Inv No..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="reco-search-input"
              />
              {searchQuery && (
                <button
                  type="button"
                  className="reco-search-clear"
                  onClick={() => setSearchQuery('')}
                  title="Clear search"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Refresh Button */}
            <button
              type="button"
              className="btn-reco-refresh"
              onClick={handleRefresh}
              disabled={isRefreshing || isLocked}
              title={isLocked ? "Reconciliation locked hai! Data refresh nahi ho sakta." : "Reload latest Books & 2B data and re-run matching"}
            >
              <svg
                width="14"
                height="14"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.2"
                strokeLinecap="round"
                strokeLinejoin="round"
                className={isRefreshing ? 'spin-icon' : ''}
              >
                <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67" />
              </svg>
              <span>{isRefreshing ? 'Refreshing...' : 'Refresh'}</span>
            </button>

            {/* Lock / Unlock Button */}
            <button
              type="button"
              className={isLocked ? 'btn-reco-unlock' : 'btn-reco-lock'}
              onClick={handleToggleLock}
              title={isLocked ? "Click to unlock this month's reconciliation" : "Click to lock this month's reconciliation"}
            >
              {isLocked ? (
                <>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                    <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                    <path d="M7 11V7a5 5 0 0 1 9.9-1" />
                  </svg>
                  <span>Unlock</span>
                </>
              ) : (
                <>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                    <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
                    <path d="M7 11V7a5 5 0 0 1 10 0v4" />
                  </svg>
                  <span>Lock</span>
                </>
              )}
            </button>
          </div>
        </div>

        {/* ======================================================== */}
        {/* 3. Detailed Data Table (Exact 23 Columns)                 */}
        {/* ======================================================== */}
        <div className="reco-table-wrapper">
          <table className="reco-table" style={{ minWidth: "1895px", width: "1895px", borderCollapse: "separate", borderSpacing: 0 }}>
            <thead>
              {/* Group Header Tier */}
              <tr>
                <th rowSpan={2} style={{ width: '140px', minWidth: '140px', textAlign: 'center' }}>STATUS</th>
                <th rowSpan={2} style={{ width: '105px', minWidth: '105px', textAlign: 'center' }}>2B MONTH</th>
                <th rowSpan={2} style={{ width: '95px', minWidth: '95px', textAlign: 'center' }}>DATE</th>
                <th rowSpan={2} style={{ width: '140px', minWidth: '140px', textAlign: 'center' }}>INVOICE NO.</th>
                <th rowSpan={2} style={{ width: '155px', minWidth: '155px', textAlign: 'center' }}>SUPPLIER GSTIN</th>
                <th rowSpan={2} style={{ width: '220px', minWidth: '220px', textAlign: 'center' }}>SUPPLIER NAME</th>

                {/* Group Heading: Books (4 cols) */}
                <th colSpan={4} className="reco-group-th reco-group-books reco-books-divider" style={{ minWidth: '445px', background: '#350B5E' }}>
                  BOOKS
                </th>

                {/* Group Heading: Govt GSTR-2B (4 cols) */}
                <th colSpan={4} className="reco-group-th reco-group-g2b" style={{ minWidth: '445px', background: '#350B5E' }}>
                  GOVT GSTR-2B
                </th>



                <th rowSpan={2} style={{ width: '145px', minWidth: '145px', textAlign: 'center' }}>ACTION</th>
              </tr>

              {/* Subheader Tier */}
              <tr className="reco-subhead-tr">
                {/* Books Subheaders (4 cols) */}
                <th style={{ width: '115px', minWidth: '115px', background: '#350B5E' }}>TAXABLE</th>
                <th style={{ width: '110px', minWidth: '110px', background: '#350B5E' }}>IGST</th>
                <th style={{ width: '110px', minWidth: '110px', background: '#350B5E' }}>CGST</th>
                <th className="reco-books-divider" style={{ width: '110px', minWidth: '110px', background: '#350B5E' }}>SGST</th>

                {/* Govt GSTR-2B Subheaders (4 cols) */}
                <th style={{ width: '115px', minWidth: '115px', background: '#350B5E' }}>TAXABLE</th>
                <th style={{ width: '110px', minWidth: '110px', background: '#350B5E' }}>IGST</th>
                <th style={{ width: '110px', minWidth: '110px', background: '#350B5E' }}>CGST</th>
                <th style={{ width: '110px', minWidth: '110px', background: '#350B5E' }}>SGST</th>


              </tr>
            </thead>

            <tbody>
              {filteredRecords.length > 0 ? (
                filteredRecords.map((r) => {
                  const isMatched = r.status === 'Matched';
                  const isNotMatch = r.status === 'Not Match';
                  const isNotIn2b = r.status === 'Not in GSTR-2B';
                  const isNotInBook = r.status === 'Not in Purchase Register';
                  const isDebitNote = r.status === 'Debit Note';

                  const hasDiff = Math.abs(r.dTaxable) > 1 || Math.abs(r.dIgst) > 1 || Math.abs(r.dCgst) > 1 || Math.abs(r.dSgst) > 1;

                  return (
                    <tr key={r.id}>
                      {/* 1. Status Column */}
                      <td style={{ width: '140px', minWidth: '140px', textAlign: 'center' }}>
                        <span
                          className={`reco-status-pill ${isMatched
                            ? 'pill-matched'
                            : isNotMatch
                              ? 'pill-not-match'
                              : isNotIn2b
                                ? 'pill-not-in-2b'
                                : isNotInBook
                                  ? 'pill-not-in-book'
                                  : 'pill-debit-note'
                            }`}
                        >
                          <span className="reco-status-dot">●</span> {r.status}
                        </span>
                        {r.auditStatus === 'Verified' && (
                          <span className="audit-status-tag verified" title={r.auditNote || 'Audit Verified'}>
                            ✓ Verified
                          </span>
                        )}
                        {r.auditStatus === 'Query' && (
                          <span className="audit-status-tag query" title={r.auditNote || 'Query Pending'}>
                            ❓ Query
                          </span>
                        )}
                      </td>

                      {/* 2. 2B Month Column (with Duplicate warning if present) */}
                      <td style={{ width: '105px', minWidth: '105px', textAlign: 'center', fontWeight: 600, color: '#334155' }}>
                        {r.month2b}
                        {r.hasDuplicateMonth && (
                          <span
                            className="duplicate-warn-tag"
                            title={`⚠️ Duplicate Warning: Invoice also found in ${r.duplicateMonthsList?.join(', ')}`}
                          >
                            ⚠️
                          </span>
                        )}
                      </td>

                      {/* 3. Date */}
                      <td style={{ width: '95px', minWidth: '95px', textAlign: 'center', color: '#475569' }}>
                        {r.date || '-'}
                      </td>

                      {/* 5. Invoice No. */}
                      <td style={{ width: '140px', minWidth: '140px', textAlign: 'center' }}>
                        <div className="reco-inv-no" style={{ textAlign: 'center' }}>{r.invNo || '-'}</div>
                        {!r.isInvMatch && r.gInvNo && (
                          <div style={{ fontSize: '11px', color: '#dc2626', fontWeight: 600, textAlign: 'center' }}>
                            2B: {r.gInvNo}
                          </div>
                        )}
                      </td>

                      {/* 6. Supplier GSTIN */}
                      <td style={{ width: '155px', minWidth: '155px', textAlign: 'center' }}>
                        <div className="reco-gstin">{r.gstin || '-'}</div>
                        {!r.isGstinMatch && r.gGstin && (
                          <div style={{ fontSize: '11px', color: '#dc2626', fontWeight: 600, textAlign: 'center' }}>
                            2B: {r.gGstin}
                          </div>
                        )}
                      </td>

                      {/* 7. Supplier Name */}
                      <td style={{ width: '220px', minWidth: '220px', textAlign: 'center' }} title={r.supplierName}>
                        <div className="reco-party-name" style={{ textAlign: 'center', margin: '0 auto' }}>{r.supplierName || '-'}</div>
                      </td>

                      {/* ================= BOOKS GROUP (4 cols) ================= */}
                      <td className="num-val reco-gst-cell" style={{ width: '115px', minWidth: '115px' }}>
                        {renderGstAmount(r.bTaxable, r.gTaxable, true, false)}
                      </td>
                      <td className="num-val reco-gst-cell" style={{ width: '110px', minWidth: '110px' }}>
                        {renderGstAmount(r.bIgst, r.gIgst, true, false)}
                      </td>
                      <td className="num-val reco-gst-cell" style={{ width: '110px', minWidth: '110px' }}>
                        {renderGstAmount(r.bCgst, r.gCgst, true, false)}
                      </td>
                      <td className="num-val reco-gst-cell reco-books-divider" style={{ width: '110px', minWidth: '110px' }}>
                        {renderGstAmount(r.bSgst, r.gSgst, true, false)}
                      </td>

                      {/* ================= GOVT GSTR-2B GROUP (4 cols) ================= */}
                      <td className="num-val reco-gst-cell" style={{ width: '115px', minWidth: '115px' }}>
                        {renderGstAmount(r.gTaxable, r.bTaxable, false, false)}
                      </td>
                      <td className="num-val reco-gst-cell" style={{ width: '110px', minWidth: '110px' }}>
                        {renderGstAmount(r.gIgst, r.bIgst, false, false)}
                      </td>
                      <td className="num-val reco-gst-cell" style={{ width: '110px', minWidth: '110px' }}>
                        {renderGstAmount(r.gCgst, r.bCgst, false, false)}
                      </td>
                      <td className="num-val reco-gst-cell" style={{ width: '110px', minWidth: '110px' }}>
                        {renderGstAmount(r.gSgst, r.bSgst, false, false)}
                      </td>



                      {/* ================= ACTION BUTTONS ================= */}
                      <td className="reco-actions-col" style={{ width: '145px', minWidth: '145px', textAlign: 'center' }}>
                        <button
                          type="button"
                          className="btn-action-view"
                          onClick={() => setViewingRecord(r)}
                          title="View Side-by-Side Comparison"
                        >
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                            <circle cx="12" cy="12" r="3"></circle>
                          </svg>
                          <span>View</span>
                        </button>
                        <button
                          type="button"
                          className="btn-action-audit"
                          onClick={() => handleOpenAudit(r)}
                          disabled={isLocked}
                          style={isLocked ? { opacity: 0.5, cursor: 'not-allowed' } : {}}
                          title={isLocked ? "Reconciliation locked hai! Audit note edit nahi ho sakta." : "Add Audit Note & Mark Verified/Query"}
                        >
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M12 20h9"></path>
                            <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
                          </svg>
                          <span>Audit</span>
                        </button>
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={15} style={{ textAlign: 'center', padding: '42px', color: '#64748b' }}>
                    No invoices found matching selected filter criteria.
                  </td>
                </tr>
              )}
            </tbody>


          </table>
        </div>
      </div>

      {/* ======================================================== */}
      {/* 4. Side-by-Side View Modal                               */}
      {/* ======================================================== */}
      {viewingRecord && (
        <div className="reco-modal-backdrop" onClick={() => setViewingRecord(null)}>
          <div
            className="reco-modal-card"
            style={{ maxWidth: '820px', width: '95%' }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="reco-modal-header">
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <h3 style={{ margin: 0 }}>Invoice Reconciliation Comparison – {viewingRecord.invNo}</h3>
                <span className={`reco-status-pill ${viewingRecord.status === 'Matched' ? 'pill-matched' :
                  viewingRecord.status === 'Not Match' ? 'pill-not-match' :
                    viewingRecord.status === 'Not in GSTR-2B' ? 'pill-not-in-2b' :
                      viewingRecord.status === 'Not in Purchase Register' ? 'pill-not-in-book' : 'pill-debit-note'
                  }`}>
                  ● {viewingRecord.status}
                </span>
              </div>
              <button
                type="button"
                className="reco-modal-close"
                onClick={() => setViewingRecord(null)}
              >
                ✕
              </button>
            </div>

            <div className="reco-modal-body" style={{ maxHeight: '70vh', overflowY: 'auto' }}>
              {/* Match Summary Banner */}
              <div
                style={{
                  padding: '12px 16px',
                  marginBottom: '16px',
                  border: viewingRecord.itcMatchCount === 6 ? '1px solid #86efac' : '1px solid #fca5a5',
                  background: viewingRecord.itcMatchCount === 6 ? '#f0fdf4' : '#fef2f2',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between'
                }}
              >
                <div>
                  <div style={{ fontWeight: 700, color: viewingRecord.itcMatchCount === 6 ? '#15803d' : '#b91c1c' }}>
                    {viewingRecord.status === 'Debit Note'
                      ? 'Debit Note (Purchase Return) – ITC Reduction'
                      : viewingRecord.itcMatchCount === 6
                        ? '✓ 6/6 Fields Match – Verified & Eligible for ITC'
                        : `${viewingRecord.itcMatchText} – Differences detected`}
                  </div>
                  {viewingRecord.mismatchedFields.length > 0 && (
                    <div style={{ fontSize: '12px', color: '#dc2626', marginTop: '3px' }}>
                      Mismatches: {viewingRecord.mismatchedFields.join(' • ')}
                    </div>
                  )}
                </div>
                <div style={{ fontSize: '12.5px', fontWeight: 600, color: '#334155' }}>
                  2B Month: <strong>{viewingRecord.month2b}</strong>
                  {viewingRecord.hasDuplicateMonth && (
                    <span style={{ color: '#d97706', marginLeft: '6px' }}>
                      (⚠️ Also in {viewingRecord.duplicateMonthsList?.join(', ')})
                    </span>
                  )}
                </div>
              </div>

              {/* Side-by-Side Table */}
              <table className="reco-view-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                <thead>
                  <tr style={{ background: '#350b5e', color: '#ffffff' }}>
                    <th style={{ padding: '9px 12px', textAlign: 'left', width: '22%' }}>FIELD NAME</th>
                    <th style={{ padding: '9px 12px', textAlign: 'right', width: '28%' }}>BOOKS</th>
                    <th style={{ padding: '9px 12px', textAlign: 'right', width: '28%' }}>GOVT GSTR-2B</th>
                    <th style={{ padding: '9px 12px', textAlign: 'right', width: '22%' }}>DIFFERENCE</th>
                  </tr>
                </thead>
                <tbody>
                  {/* GSTIN */}
                  <tr style={{ borderBottom: '1px solid #edf2f7', background: !viewingRecord.isGstinMatch && viewingRecord.gGstin ? '#fee2e2' : 'transparent' }}>
                    <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>Supplier GSTIN</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 600 }}>{viewingRecord.gstin || '-'}</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 600 }}>{viewingRecord.gGstin || viewingRecord.gstin || '-'}</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right' }}>
                      <span className={`mini-match-pill ${viewingRecord.isGstinMatch ? 'match' : 'mismatch'}`}>
                        {viewingRecord.isGstinMatch ? 'Match' : 'Not Match'}
                      </span>
                    </td>
                  </tr>

                  {/* Invoice No */}
                  <tr style={{ borderBottom: '1px solid #edf2f7', background: !viewingRecord.isInvMatch && viewingRecord.gInvNo ? '#fee2e2' : 'transparent' }}>
                    <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>Invoice No.</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 600 }}>{viewingRecord.invNo || '-'}</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 600 }}>{viewingRecord.gInvNo || viewingRecord.invNo || '-'}</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right' }}>
                      <span className={`mini-match-pill ${viewingRecord.isInvMatch ? 'match' : 'mismatch'}`}>
                        {viewingRecord.isInvMatch ? 'Match' : 'Not Match'}
                      </span>
                    </td>
                  </tr>

                  {/* Supplier Name */}
                  <tr style={{ borderBottom: '1px solid #edf2f7' }}>
                    <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>Supplier Name</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right' }}>{viewingRecord.supplierName || '-'}</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right' }}>{viewingRecord.gTradeName || viewingRecord.supplierName || '-'}</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right' }}>-</td>
                  </tr>

                  {/* Date */}
                  <tr style={{ borderBottom: '1px solid #edf2f7' }}>
                    <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>Invoice Date</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right' }}>{viewingRecord.date || '-'}</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right' }}>{viewingRecord.gDate || viewingRecord.date || '-'}</td>
                    <td style={{ padding: '9px 12px', textAlign: 'right' }}>-</td>
                  </tr>

                  {/* Taxable Value */}
                  <tr style={{ borderBottom: '1px solid #edf2f7', background: !viewingRecord.isTaxableMatch ? '#fee2e2' : 'transparent' }}>
                    <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>Taxable Value</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700 }}>{formatAmount(viewingRecord.bTaxable)}</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700 }}>{formatAmount(viewingRecord.gTaxable)}</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: Math.abs(viewingRecord.dTaxable) > 1 ? '#dc2626' : '#15803d' }}>
                      {formatAmount(viewingRecord.dTaxable)}
                    </td>
                  </tr>

                  {/* IGST */}
                  <tr style={{ borderBottom: '1px solid #edf2f7', background: !viewingRecord.isIgstMatch ? '#fee2e2' : 'transparent' }}>
                    <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>IGST</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700 }}>{formatAmount(viewingRecord.bIgst)}</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700 }}>{formatAmount(viewingRecord.gIgst)}</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: Math.abs(viewingRecord.dIgst) > 1 ? '#dc2626' : '#15803d' }}>
                      {formatAmount(viewingRecord.dIgst)}
                    </td>
                  </tr>

                  {/* CGST */}
                  <tr style={{ borderBottom: '1px solid #edf2f7', background: !viewingRecord.isCgstMatch ? '#fee2e2' : 'transparent' }}>
                    <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>CGST</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700 }}>{formatAmount(viewingRecord.bCgst)}</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700 }}>{formatAmount(viewingRecord.gCgst)}</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: Math.abs(viewingRecord.dCgst) > 1 ? '#dc2626' : '#15803d' }}>
                      {formatAmount(viewingRecord.dCgst)}
                    </td>
                  </tr>

                  {/* SGST */}
                  <tr style={{ borderBottom: '1px solid #edf2f7', background: !viewingRecord.isSgstMatch ? '#fee2e2' : 'transparent' }}>
                    <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>SGST</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700 }}>{formatAmount(viewingRecord.bSgst)}</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700 }}>{formatAmount(viewingRecord.gSgst)}</td>
                    <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: Math.abs(viewingRecord.dSgst) > 1 ? '#dc2626' : '#15803d' }}>
                      {formatAmount(viewingRecord.dSgst)}
                    </td>
                  </tr>

                  {/* Total */}
                  <tr style={{ background: '#f8fafc', fontWeight: 800 }}>
                    <td style={{ padding: '11px 12px', color: '#1e293b' }}>TOTAL AMOUNT</td>
                    <td className="num-val" style={{ padding: '11px 12px', textAlign: 'right', color: '#350b5e', fontSize: '14px' }}>
                      {formatAmount(viewingRecord.bTotal)}
                    </td>
                    <td className="num-val" style={{ padding: '11px 12px', textAlign: 'right', color: '#1e1b4b', fontSize: '14px' }}>
                      {formatAmount(viewingRecord.gTotal)}
                    </td>
                    <td className="num-val" style={{ padding: '11px 12px', textAlign: 'right', color: Math.abs(viewingRecord.dTotal) > 1 ? '#dc2626' : '#15803d', fontSize: '14px' }}>
                      {formatAmount(viewingRecord.dTotal)}
                    </td>
                  </tr>
                </tbody>
              </table>

              {/* Audit History inside View Modal if available */}
              {viewingRecord.auditNote && (
                <div style={{ marginTop: '16px', padding: '12px 14px', background: '#f1f5f9', borderLeft: '3px solid #350b5e' }}>
                  <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#475569', textTransform: 'uppercase', marginBottom: '4px' }}>
                    Saved Audit Observation ({viewingRecord.auditStatus}):
                  </div>
                  <div style={{ fontSize: '13px', color: '#1e293b' }}>{viewingRecord.auditNote}</div>
                  {viewingRecord.auditUpdatedAt && (
                    <div style={{ fontSize: '11px', color: '#64748b', marginTop: '4px' }}>
                      Updated: {viewingRecord.auditUpdatedAt}
                    </div>
                  )}
                </div>
              )}

              {/* Audit Change Log in View Modal */}
              {(() => {
                const invoiceLogs = auditChangeLogs.filter(
                  (l) => cleanInv(l.invNo) === cleanInv(viewingRecord.invNo) || cleanStr(l.gstin) === cleanStr(viewingRecord.gstin)
                );
                if (invoiceLogs.length === 0) return null;
                return (
                  <div style={{ marginTop: '16px', background: '#f8fafc', border: '1px solid #e2e8f0', padding: '12px 14px' }}>
                    <div style={{ fontSize: '12px', fontWeight: 700, color: '#350b5e', display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' }}>
                      <span>📋</span>
                      <span>Audit Change Trail (Audit Log):</span>
                      <span style={{ fontSize: '11px', color: '#64748b', fontWeight: 500 }}>({invoiceLogs.length} changes recorded)</span>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '130px', overflowY: 'auto' }}>
                      {invoiceLogs.map((log) => (
                        <div key={log.id} style={{ fontSize: '12px', background: '#ffffff', border: '1px solid #e2e8f0', padding: '6px 10px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <div>
                            <strong style={{ color: '#0f172a' }}>{log.field}:</strong>{' '}
                            <span style={{ color: '#dc2626', textDecoration: 'line-through' }}>{log.oldValue}</span>{' '}
                            ➔{' '}
                            <span style={{ color: '#15803d', fontWeight: 700 }}>{log.newValue}</span>
                          </div>
                          <div style={{ fontSize: '11px', color: '#64748b', whiteSpace: 'nowrap', marginLeft: '12px' }}>
                            by <strong>{log.user}</strong> on {log.timestamp}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })()}
            </div>

            <div className="reco-modal-footer">
              <button
                type="button"
                className="btn-reco-action"
                onClick={() => setViewingRecord(null)}
              >
                Close View
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ======================================================== */}
      {/* 5. Audit Action Modal (Edit Mode Comparison Table)       */}
      {/* ======================================================== */}
      {auditingRecord && (() => {
        const curTaxable = Number(auditEditState.taxable) || 0;
        const curIgst = Number(auditEditState.igst) || 0;
        const curCgst = Number(auditEditState.cgst) || 0;
        const curSgst = Number(auditEditState.sgst) || 0;
        const curTotal = round2(curTaxable + curIgst + curCgst + curSgst);

        const diffTaxable = round2(curTaxable - auditingRecord.gTaxable);
        const diffIgst = round2(curIgst - auditingRecord.gIgst);
        const diffCgst = round2(curCgst - auditingRecord.gCgst);
        const diffSgst = round2(curSgst - auditingRecord.gSgst);
        const diffTotal = round2(curTotal - auditingRecord.gTotal);

        const isGstinMatch = cleanStr(auditEditState.gstin) === cleanStr(auditingRecord.gGstin || auditingRecord.gstin);
        const isInvMatch = cleanInv(auditEditState.invNo) === cleanInv(auditingRecord.gInvNo || auditingRecord.invNo);
        const isNameMatch = cleanStr(auditEditState.supplierName) === cleanStr(auditingRecord.gTradeName || auditingRecord.supplierName);
        const isDateMatch = !auditingRecord.gDate || auditEditState.date === auditingRecord.gDate;
        const isTaxableMatch = Math.abs(diffTaxable) <= 1.00;
        const isIgstMatch = Math.abs(diffIgst) <= 1.00;
        const isCgstMatch = Math.abs(diffCgst) <= 1.00;
        const isSgstMatch = Math.abs(diffSgst) <= 1.00;

        const all6Match = isGstinMatch && isInvMatch && isTaxableMatch && isIgstMatch && isCgstMatch && isSgstMatch;

        // Which fields were mismatched initially (Requirement 3: only mismatched fields are editable)
        // If invoice is 'Not in Purchase Register', all fields are missing in Books, so all are editable
        const isNotInBook = auditingRecord.status === 'Not in Purchase Register';
        const canEditGstin = isNotInBook || !auditingRecord.isGstinMatch;
        const canEditInv = isNotInBook || !auditingRecord.isInvMatch;
        const canEditName = isNotInBook || cleanStr(auditingRecord.supplierName) !== cleanStr(auditingRecord.gTradeName || auditingRecord.supplierName);
        const canEditDate = isNotInBook || (Boolean(auditingRecord.gDate) && auditingRecord.date !== auditingRecord.gDate);
        const canEditTaxable = isNotInBook || !auditingRecord.isTaxableMatch;
        const canEditIgst = isNotInBook || !auditingRecord.isIgstMatch;
        const canEditCgst = isNotInBook || !auditingRecord.isCgstMatch;
        const canEditSgst = isNotInBook || !auditingRecord.isSgstMatch;

        // Relevant logs for this invoice
        const invoiceLogs = auditChangeLogs.filter(
          (l) => cleanInv(l.invNo) === cleanInv(auditingRecord.invNo) || cleanStr(l.gstin) === cleanStr(auditingRecord.gstin)
        );

        const handleCopyAllFrom2b = () => {
          setAuditEditState(prev => ({
            ...prev,
            gstin: auditingRecord.gGstin || auditingRecord.gstin || prev.gstin,
            invNo: auditingRecord.gInvNo || auditingRecord.invNo || prev.invNo,
            supplierName: auditingRecord.gTradeName || auditingRecord.supplierName || prev.supplierName,
            date: auditingRecord.gDate || auditingRecord.date || prev.date,
            taxable: auditingRecord.gTaxable,
            igst: auditingRecord.gIgst,
            cgst: auditingRecord.gCgst,
            sgst: auditingRecord.gSgst
          }));
          showToast('Copied GSTR-2B values to Books inputs.');
        };

        return (
          <div className="reco-modal-backdrop" onClick={() => setAuditingRecord(null)}>
            <div
              className="reco-modal-card"
              style={{ maxWidth: '850px', width: '96%' }}
              onClick={(e) => e.stopPropagation()}
            >
              {/* Header */}
              <div className="reco-modal-header">
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <h3 style={{ margin: 0, color: '#350b5e' }}>Invoice Reconciliation Audit – {auditingRecord.invNo}</h3>
                  <span className={`reco-status-pill ${all6Match ? 'pill-matched' :
                    auditingRecord.status === 'Matched' ? 'pill-matched' :
                    auditingRecord.status === 'Not Match' ? 'pill-not-match' :
                    auditingRecord.status === 'Not in GSTR-2B' ? 'pill-not-in-2b' :
                    auditingRecord.status === 'Not in Purchase Register' ? 'pill-not-in-book' : 'pill-debit-note'
                  }`}>
                    ● {all6Match ? 'Matched (Auto)' : auditingRecord.status}
                  </span>
                </div>
                <button
                  type="button"
                  className="reco-modal-close"
                  onClick={() => setAuditingRecord(null)}
                >
                  ✕
                </button>
              </div>

              <div className="reco-modal-body" style={{ maxHeight: '72vh', overflowY: 'auto' }}>
                {/* Match Summary & Sync Bar */}
                <div
                  style={{
                    padding: '12px 16px',
                    marginBottom: '16px',
                    border: all6Match ? '1px solid #86efac' : '1px solid #fca5a5',
                    background: all6Match ? '#f0fdf4' : '#fef2f2',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    flexWrap: 'wrap',
                    gap: '10px'
                  }}
                >
                  <div>
                    <div style={{ fontWeight: 700, color: all6Match ? '#15803d' : '#b91c1c' }}>
                      {all6Match
                        ? '✓ All 6 Fields Match! On Save, status will automatically become "Matched" (ITC Verified).'
                        : 'Audit Edit Mode – Only mismatched fields in Books are editable. Govt GSTR-2B is read-only.'}
                    </div>
                    <div style={{ fontSize: '11.5px', color: '#64748b', marginTop: '3px' }}>
                      2B Month: <strong>{auditingRecord.month2b}</strong> • Matched fields are locked 🔒.
                    </div>
                  </div>

                  {!all6Match && (
                    <button
                      type="button"
                      onClick={handleCopyAllFrom2b}
                      style={{
                        background: '#350b5e',
                        color: '#ffffff',
                        border: 'none',
                        padding: '6px 12px',
                        fontSize: '11.5px',
                        fontWeight: 700,
                        cursor: 'pointer',
                        borderRadius: 0,
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '5px'
                      }}
                      title="Auto-fill Books fields with Govt GSTR-2B values to match"
                    >
                      ⚡ Copy 2B Values to Books
                    </button>
                  )}
                </div>

                {/* Side-by-Side Comparison & Edit Table */}
                <table className="reco-view-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                  <thead>
                    <tr style={{ background: '#350b5e', color: '#ffffff' }}>
                      <th style={{ padding: '9px 12px', textAlign: 'left', width: '22%' }}>FIELD NAME</th>
                      <th style={{ padding: '9px 12px', textAlign: 'right', width: '30%' }}>BOOKS (EDITABLE)</th>
                      <th style={{ padding: '9px 12px', textAlign: 'right', width: '28%' }}>GOVT GSTR-2B (READ-ONLY)</th>
                      <th style={{ padding: '9px 12px', textAlign: 'right', width: '20%' }}>DIFFERENCE</th>
                    </tr>
                  </thead>
                  <tbody>
                    {/* Supplier GSTIN */}
                    <tr style={{ borderBottom: '1px solid #edf2f7', background: !isGstinMatch ? '#fee2e2' : 'transparent' }}>
                      <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>
                        Supplier GSTIN
                      </td>
                      <td style={{ padding: '6px 10px', textAlign: 'right' }}>
                        {canEditGstin ? (
                          <input
                            type="text"
                            value={auditEditState.gstin}
                            onChange={(e) => setAuditEditState(prev => ({ ...prev, gstin: e.target.value.toUpperCase() }))}
                            maxLength={15}
                            style={{
                              width: '100%',
                              padding: '5px 8px',
                              border: '1.5px solid #7c3aed',
                              borderRadius: 0,
                              fontFamily: 'monospace',
                              fontWeight: 700,
                              fontSize: '12px',
                              outline: 'none',
                              textAlign: 'right'
                            }}
                          />
                        ) : (
                          <span style={{ fontWeight: 600 }}>
                            {auditEditState.gstin || '-'} <span title="Locked: Field Matches 2B" style={{ color: '#15803d', fontSize: '11px', marginLeft: '3px' }}>🔒</span>
                          </span>
                        )}
                      </td>
                      <td style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 600, color: '#1e293b' }}>
                        {auditingRecord.gGstin || auditingRecord.gstin || '-'}
                      </td>
                      <td style={{ padding: '9px 12px', textAlign: 'right' }}>
                        <span className={`mini-match-pill ${isGstinMatch ? 'match' : 'mismatch'}`}>
                          {isGstinMatch ? 'Match' : 'Not Match'}
                        </span>
                      </td>
                    </tr>

                    {/* Invoice No. */}
                    <tr style={{ borderBottom: '1px solid #edf2f7', background: !isInvMatch ? '#fee2e2' : 'transparent' }}>
                      <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>
                        Invoice No.
                      </td>
                      <td style={{ padding: '6px 10px', textAlign: 'right' }}>
                        {canEditInv ? (
                          <input
                            type="text"
                            value={auditEditState.invNo}
                            onChange={(e) => setAuditEditState(prev => ({ ...prev, invNo: e.target.value }))}
                            style={{
                              width: '100%',
                              padding: '5px 8px',
                              border: '1.5px solid #7c3aed',
                              borderRadius: 0,
                              fontWeight: 700,
                              fontSize: '12px',
                              outline: 'none',
                              textAlign: 'right'
                            }}
                          />
                        ) : (
                          <span style={{ fontWeight: 600 }}>
                            {auditEditState.invNo || '-'} <span title="Locked: Field Matches 2B" style={{ color: '#15803d', fontSize: '11px', marginLeft: '3px' }}>🔒</span>
                          </span>
                        )}
                      </td>
                      <td style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 600, color: '#1e293b' }}>
                        {auditingRecord.gInvNo || auditingRecord.invNo || '-'}
                      </td>
                      <td style={{ padding: '9px 12px', textAlign: 'right' }}>
                        <span className={`mini-match-pill ${isInvMatch ? 'match' : 'mismatch'}`}>
                          {isInvMatch ? 'Match' : 'Not Match'}
                        </span>
                      </td>
                    </tr>

                    {/* Supplier Name */}
                    <tr style={{ borderBottom: '1px solid #edf2f7' }}>
                      <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>
                        Supplier Name
                      </td>
                      <td style={{ padding: '6px 10px', textAlign: 'right' }}>
                        {canEditName ? (
                          <input
                            type="text"
                            value={auditEditState.supplierName}
                            onChange={(e) => setAuditEditState(prev => ({ ...prev, supplierName: e.target.value }))}
                            style={{
                              width: '100%',
                              padding: '5px 8px',
                              border: '1.5px solid #7c3aed',
                              borderRadius: 0,
                              fontSize: '12px',
                              outline: 'none',
                              textAlign: 'right'
                            }}
                          />
                        ) : (
                          <span>
                            {auditEditState.supplierName || '-'} <span title="Locked: Matches 2B" style={{ color: '#15803d', fontSize: '11px', marginLeft: '3px' }}>🔒</span>
                          </span>
                        )}
                      </td>
                      <td style={{ padding: '9px 12px', textAlign: 'right', color: '#1e293b' }}>
                        {auditingRecord.gTradeName || auditingRecord.supplierName || '-'}
                      </td>
                      <td style={{ padding: '9px 12px', textAlign: 'right' }}>
                        <span className={`mini-match-pill ${isNameMatch ? 'match' : 'mismatch'}`}>
                          {isNameMatch ? 'Match' : 'Not Match'}
                        </span>
                      </td>
                    </tr>

                    {/* Invoice Date */}
                    <tr style={{ borderBottom: '1px solid #edf2f7' }}>
                      <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>
                        Invoice Date
                      </td>
                      <td style={{ padding: '6px 10px', textAlign: 'right' }}>
                        {canEditDate ? (
                          <input
                            type="text"
                            value={auditEditState.date}
                            onChange={(e) => setAuditEditState(prev => ({ ...prev, date: e.target.value }))}
                            placeholder="YYYY-MM-DD"
                            style={{
                              width: '100%',
                              padding: '5px 8px',
                              border: '1.5px solid #7c3aed',
                              borderRadius: 0,
                              fontSize: '12px',
                              outline: 'none',
                              textAlign: 'right'
                            }}
                          />
                        ) : (
                          <span>
                            {auditEditState.date || '-'} <span title="Locked: Matches 2B" style={{ color: '#15803d', fontSize: '11px', marginLeft: '3px' }}>🔒</span>
                          </span>
                        )}
                      </td>
                      <td style={{ padding: '9px 12px', textAlign: 'right', color: '#1e293b' }}>
                        {auditingRecord.gDate || auditingRecord.date || '-'}
                      </td>
                      <td style={{ padding: '9px 12px', textAlign: 'right' }}>
                        <span className={`mini-match-pill ${isDateMatch ? 'match' : 'mismatch'}`}>
                          {isDateMatch ? 'Match' : 'Not Match'}
                        </span>
                      </td>
                    </tr>

                    {/* Taxable Value */}
                    <tr style={{ borderBottom: '1px solid #edf2f7', background: !isTaxableMatch ? '#fee2e2' : 'transparent' }}>
                      <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>
                        Taxable Value
                      </td>
                      <td style={{ padding: '6px 10px', textAlign: 'right' }}>
                        {canEditTaxable ? (
                          <input
                            type="number"
                            step="0.01"
                            value={auditEditState.taxable}
                            onChange={(e) => setAuditEditState(prev => ({ ...prev, taxable: parseFloat(e.target.value) || 0 }))}
                            style={{
                              width: '100%',
                              padding: '5px 8px',
                              border: '1.5px solid #7c3aed',
                              borderRadius: 0,
                              fontWeight: 700,
                              fontSize: '12.5px',
                              textAlign: 'right',
                              outline: 'none'
                            }}
                          />
                        ) : (
                          <span className="num-val" style={{ fontWeight: 700 }}>
                            {formatAmount(curTaxable)} <span title="Locked: Matches 2B" style={{ color: '#15803d', fontSize: '11px', marginLeft: '3px' }}>🔒</span>
                          </span>
                        )}
                      </td>
                      <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: '#1e293b' }}>
                        {formatAmount(auditingRecord.gTaxable)}
                      </td>
                      <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: isTaxableMatch ? '#15803d' : '#dc2626' }}>
                        {formatAmount(diffTaxable)}
                      </td>
                    </tr>

                    {/* IGST */}
                    <tr style={{ borderBottom: '1px solid #edf2f7', background: !isIgstMatch ? '#fee2e2' : 'transparent' }}>
                      <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>
                        IGST
                      </td>
                      <td style={{ padding: '6px 10px', textAlign: 'right' }}>
                        {canEditIgst ? (
                          <input
                            type="number"
                            step="0.01"
                            value={auditEditState.igst}
                            onChange={(e) => setAuditEditState(prev => ({ ...prev, igst: parseFloat(e.target.value) || 0 }))}
                            style={{
                              width: '100%',
                              padding: '5px 8px',
                              border: '1.5px solid #7c3aed',
                              borderRadius: 0,
                              fontWeight: 700,
                              fontSize: '12.5px',
                              textAlign: 'right',
                              outline: 'none'
                            }}
                          />
                        ) : (
                          <span className="num-val" style={{ fontWeight: 700 }}>
                            {formatAmount(curIgst)} <span title="Locked: Matches 2B" style={{ color: '#15803d', fontSize: '11px', marginLeft: '3px' }}>🔒</span>
                          </span>
                        )}
                      </td>
                      <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: '#1e293b' }}>
                        {formatAmount(auditingRecord.gIgst)}
                      </td>
                      <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: isIgstMatch ? '#15803d' : '#dc2626' }}>
                        {formatAmount(diffIgst)}
                      </td>
                    </tr>

                    {/* CGST */}
                    <tr style={{ borderBottom: '1px solid #edf2f7', background: !isCgstMatch ? '#fee2e2' : 'transparent' }}>
                      <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>
                        CGST
                      </td>
                      <td style={{ padding: '6px 10px', textAlign: 'right' }}>
                        {canEditCgst ? (
                          <input
                            type="number"
                            step="0.01"
                            value={auditEditState.cgst}
                            onChange={(e) => setAuditEditState(prev => ({ ...prev, cgst: parseFloat(e.target.value) || 0 }))}
                            style={{
                              width: '100%',
                              padding: '5px 8px',
                              border: '1.5px solid #7c3aed',
                              borderRadius: 0,
                              fontWeight: 700,
                              fontSize: '12.5px',
                              textAlign: 'right',
                              outline: 'none'
                            }}
                          />
                        ) : (
                          <span className="num-val" style={{ fontWeight: 700 }}>
                            {formatAmount(curCgst)} <span title="Locked: Matches 2B" style={{ color: '#15803d', fontSize: '11px', marginLeft: '3px' }}>🔒</span>
                          </span>
                        )}
                      </td>
                      <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: '#1e293b' }}>
                        {formatAmount(auditingRecord.gCgst)}
                      </td>
                      <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: isCgstMatch ? '#15803d' : '#dc2626' }}>
                        {formatAmount(diffCgst)}
                      </td>
                    </tr>

                    {/* SGST */}
                    <tr style={{ borderBottom: '1px solid #edf2f7', background: !isSgstMatch ? '#fee2e2' : 'transparent' }}>
                      <td style={{ padding: '9px 12px', fontWeight: 600, color: '#475569' }}>
                        SGST
                      </td>
                      <td style={{ padding: '6px 10px', textAlign: 'right' }}>
                        {canEditSgst ? (
                          <input
                            type="number"
                            step="0.01"
                            value={auditEditState.sgst}
                            onChange={(e) => setAuditEditState(prev => ({ ...prev, sgst: parseFloat(e.target.value) || 0 }))}
                            style={{
                              width: '100%',
                              padding: '5px 8px',
                              border: '1.5px solid #7c3aed',
                              borderRadius: 0,
                              fontWeight: 700,
                              fontSize: '12.5px',
                              textAlign: 'right',
                              outline: 'none'
                            }}
                          />
                        ) : (
                          <span className="num-val" style={{ fontWeight: 700 }}>
                            {formatAmount(curSgst)} <span title="Locked: Matches 2B" style={{ color: '#15803d', fontSize: '11px', marginLeft: '3px' }}>🔒</span>
                          </span>
                        )}
                      </td>
                      <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: '#1e293b' }}>
                        {formatAmount(auditingRecord.gSgst)}
                      </td>
                      <td className="num-val" style={{ padding: '9px 12px', textAlign: 'right', fontWeight: 700, color: isSgstMatch ? '#15803d' : '#dc2626' }}>
                        {formatAmount(diffSgst)}
                      </td>
                    </tr>

                    {/* Total Amount Row */}
                    <tr style={{ background: '#f8fafc', fontWeight: 800 }}>
                      <td style={{ padding: '11px 12px', color: '#1e293b' }}>TOTAL AMOUNT</td>
                      <td className="num-val" style={{ padding: '11px 12px', textAlign: 'right', color: '#350b5e', fontSize: '14px' }}>
                        {formatAmount(curTotal)}
                      </td>
                      <td className="num-val" style={{ padding: '11px 12px', textAlign: 'right', color: '#1e1b4b', fontSize: '14px' }}>
                        {formatAmount(auditingRecord.gTotal)}
                      </td>
                      <td className="num-val" style={{ padding: '11px 12px', textAlign: 'right', color: Math.abs(diffTotal) <= 1 ? '#15803d' : '#dc2626', fontSize: '14px' }}>
                        {formatAmount(diffTotal)}
                      </td>
                    </tr>
                  </tbody>
                </table>

                {/* Audit Observation Remarks */}
                <div style={{ marginTop: '16px' }}>
                  <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, color: '#334155', marginBottom: '4px' }}>
                    Audit Observation / Remark (Optional):
                  </label>
                  <input
                    type="text"
                    value={auditEditState.note}
                    onChange={(e) => setAuditEditState(prev => ({ ...prev, note: e.target.value }))}
                    placeholder="Enter audit explanation or reference..."
                    style={{
                      width: '100%',
                      padding: '8px 12px',
                      border: '1.5px solid #cbd5e1',
                      borderRadius: 0,
                      fontSize: '12.5px',
                      outline: 'none'
                    }}
                  />
                </div>

                {/* Audit Log Trail (Requirement 7) */}
                {invoiceLogs.length > 0 && (
                  <div style={{ marginTop: '16px', background: '#f8fafc', border: '1px solid #e2e8f0', padding: '12px 14px' }}>
                    <div style={{ fontSize: '12px', fontWeight: 700, color: '#350b5e', display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' }}>
                      <span>📋</span>
                      <span>Audit Change Trail (Audit Log History):</span>
                      <span style={{ fontSize: '11px', color: '#64748b', fontWeight: 500 }}>({invoiceLogs.length} changes recorded)</span>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '130px', overflowY: 'auto' }}>
                      {invoiceLogs.map(log => (
                        <div key={log.id} style={{ fontSize: '12px', background: '#ffffff', border: '1px solid #e2e8f0', padding: '6px 10px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <div>
                            <strong style={{ color: '#0f172a' }}>{log.field}:</strong>{' '}
                            <span style={{ color: '#dc2626', textDecoration: 'line-through' }}>{log.oldValue}</span>{' '}
                            ➔{' '}
                            <span style={{ color: '#15803d', fontWeight: 700 }}>{log.newValue}</span>
                          </div>
                          <div style={{ fontSize: '11px', color: '#64748b', whiteSpace: 'nowrap', marginLeft: '12px' }}>
                            by <strong>{log.user}</strong> on {log.timestamp}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Footer */}
              <div className="reco-modal-footer">
                <button
                  type="button"
                  className="btn-reco-cancel"
                  onClick={() => setAuditingRecord(null)}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  className="btn-reco-action"
                  style={{ background: '#16a34a' }}
                  onClick={() => handleSaveAudit()}
                >
                  Save Audit
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* Confirmation Dialog (matching ITC & GSTR-2B Sheet Modal) */}
      {confirmDialog && (
        <div className="mo" onClick={() => setConfirmDialog(null)}>
          <div className="mb" style={{ maxWidth: '440px' }} onClick={(e) => e.stopPropagation()}>
            <h3 style={{ margin: 0, color: '#350B5E', fontSize: '18px', fontWeight: 700 }}>
              {confirmDialog.title}
            </h3>
            <p style={{ margin: '14px 0 22px', color: '#475569', fontSize: '14px', lineHeight: 1.5 }}>
              {confirmDialog.message}
            </p>
            <div className="fr" style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px', margin: 0 }}>
              <button
                type="button"
                style={{
                  background: '#ffffff',
                  border: '1px solid #cbd5e1',
                  color: '#334155',
                  padding: '8px 18px',
                  fontWeight: 600,
                  borderRadius: 0,
                  cursor: 'pointer',
                  fontSize: '13px'
                }}
                onClick={() => setConfirmDialog(null)}
              >
                Cancel
              </button>
              <button
                type="button"
                className="pri"
                style={{
                  background: '#350B5E',
                  color: '#ffffff',
                  border: '1px solid #350B5E',
                  padding: '8px 18px',
                  fontWeight: 600,
                  borderRadius: 0,
                  cursor: 'pointer',
                  fontSize: '13px'
                }}
                onClick={() => {
                  confirmDialog.onConfirm();
                  setConfirmDialog(null);
                }}
              >
                {confirmDialog.confirmLabel || 'Confirm'}
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
