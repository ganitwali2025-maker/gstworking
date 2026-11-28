import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { useLiveReconciliation } from '../services/recoSyncService';
import { loadRegisterRows, REGISTER_TABS } from '../services/registerService';
import { isWorkingMonthLocked, GST_LOCK_CHANGED_EVENT } from '../services/lockService';
import { getOldItcPool } from '../services/oldItcService';
import defaultRegisters from '../data/defaultRegisters.json';

export type WorkingDetailCategory = 'matched' | 'not-in-book' | 'not-in-2b' | 'not-match';

interface WorkingStepsDetailSheetProps {
  category: WorkingDetailCategory;
  onBack?: () => void;
  showToast: (msg: string) => void;
  workingMonth?: string;
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
  month: string;
}

export interface DetailSheetRow {
  id: string;
  status: string;
  month2b: string;
  date: string;
  invNo: string;
  gstin: string;
  supplierName: string;
  taxable: number;
  igst: number;
  cgst: number;
  sgst: number;
  total: number;
  gDate?: string;
  gInvNo?: string;
  gGstin?: string;
  gTradeName?: string;
  gTaxable?: number;
  gIgst?: number;
  gCgst?: number;
  gSgst?: number;
  gTotal?: number;
  reason?: string;
  mismatchFields?: string[];
  auditStatus?: string;
  auditNote?: string;
}

function cleanStr(val: any): string {
  if (!val) return '';
  return String(val).trim().toUpperCase();
}

function cleanInv(val: any): string {
  if (!val) return '';
  return String(val).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function round2(num: number): number {
  return Math.round((num + Number.EPSILON) * 100) / 100;
}

function formatInr(val: number): string {
  const rounded = round2(val || 0);
  return '₹' + rounded.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function renderAmountCell(val: number, isTotal = false) {
  const isZero = Math.abs(val) < 0.005;
  const isNegative = val < -0.005;
  const formatted = Math.abs(val).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  const displayStr = isNegative ? `-${formatted}` : formatted;

  let className = 'amt-cell';
  if (isNegative) className += ' amt-cell-neg';
  else if (isZero) className += ' amt-cell-zero';

  return (
    <span className={className} style={isTotal ? { fontWeight: 800, fontSize: '14px' } : undefined}>
      {displayStr}
    </span>
  );
}

function parseRecordMonth(item: any, defaultMonthPrefix: string = 'Oct'): string {
  if (item.month && typeof item.month === 'string' && item.month.includes('-')) {
    return item.month.trim();
  }
  const dateStr = item.d || item.date || '';
  if (dateStr) {
    const parts = dateStr.split(/[-/]/);
    if (parts.length >= 3) {
      let mNum = 0;
      let yNum = '2026';
      if (parts[0].length === 4) {
        yNum = parts[0];
        mNum = parseInt(parts[1], 10);
      } else {
        mNum = parseInt(parts[1], 10);
        yNum = parts[2].length === 2 ? `20${parts[2]}` : parts[2];
      }
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      if (mNum >= 1 && mNum <= 12) {
        return `${months[mNum - 1]}-${yNum}`;
      }
    }
  }
  return `${defaultMonthPrefix}-2026`;
}

function getMonthPriority(recordMonth: string, currentMonth: string): number {
  if (!recordMonth || recordMonth === '-') return 999;
  if (recordMonth === currentMonth) return 0;
  const monthOrder = ['Apr-2026', 'May-2026', 'Jun-2026', 'Jul-2026', 'Aug-2026', 'Sep-2026', 'Oct-2026', 'Nov-2026', 'Dec-2026', 'Jan-2027', 'Feb-2027', 'Mar-2027'];
  const curIdx = monthOrder.indexOf(currentMonth);
  const recIdx = monthOrder.indexOf(recordMonth);
  if (curIdx === -1 || recIdx === -1) return 100;
  if (recIdx > curIdx) return recIdx - curIdx;
  return 50 + (curIdx - recIdx);
}

export default function WorkingStepsDetailSheet({
  category,
  onBack,
  showToast,
  workingMonth = 'Oct-2026'
}: WorkingStepsDetailSheetProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [isLocked, setIsLocked] = useState<boolean>(() => isWorkingMonthLocked(workingMonth));

  // Modals for not-in-2b
  const [viewItem, setViewItem] = useState<DetailSheetRow | null>(null);
  const [auditItem, setAuditItem] = useState<DetailSheetRow | null>(null);
  const [auditStatus, setAuditStatus] = useState<string>('Pending');
  const [auditNote, setAuditNote] = useState<string>('');
  const [dataVersion, setDataVersion] = useState(0);

  useEffect(() => {
    setIsLocked(isWorkingMonthLocked(workingMonth));
    const handleLockEvt = () => setIsLocked(isWorkingMonthLocked(workingMonth));
    window.addEventListener(GST_LOCK_CHANGED_EVENT, handleLockEvt);
    return () => window.removeEventListener(GST_LOCK_CHANGED_EVENT, handleLockEvt);
  }, [workingMonth]);

  // Sync fresh register rows on mount and listen to changes
  useEffect(() => {
    let isMounted = true;
    async function syncRegisters() {
      try {
        await Promise.all([
          loadRegisterRows(REGISTER_TABS.ITC_PURCHASE),
          loadRegisterRows(REGISTER_TABS.GSTR2B)
        ]);
        if (isMounted) {
          setDataVersion((v) => v + 1);
        }
      } catch (err) {
        console.warn('Live register sync note:', err);
      }
    }
    syncRegisters();

    const handleSync = () => setDataVersion((v) => v + 1);
    window.addEventListener('register-synced', handleSync);
    window.addEventListener('storage', handleSync);

    return () => {
      isMounted = false;
      window.removeEventListener('register-synced', handleSync);
      window.removeEventListener('storage', handleSync);
    };
  }, [workingMonth]);

  // Live one-way sync with Reco sheet engine
  const recoData = useLiveReconciliation(workingMonth);

  // Compute Data per Category
  const { categoryRows, totalInvoicesContext, categoryCount, matchPercentage } = useMemo(() => {
    let selectedRows: DetailSheetRow[] = [];
    let totalContext = 0;
    let count = 0;

    if (category === 'matched') {
      selectedRows = recoData.matched;
      totalContext = recoData.summary.totalCount;
      count = recoData.summary.matchedCount;
    } else if (category === 'not-in-book') {
      // Load all Books records across all months to check if an entry was made in later months
      let allBooks: any[] = [];
      try {
        const raw = localStorage.getItem('itc_register_v1');
        if (raw) {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed) && parsed.length > 0) allBooks = parsed;
        }
      } catch {}
      if (!allBooks.length) {
        allBooks = (defaultRegisters as any)?.itc || [];
      }

      // Filter out debit notes
      const regularBooks = allBooks.filter((b) => {
        const inv = String(b.invNo || b.n || '').toUpperCase();
        const tax = Number(b.taxable || b.t || 0);
        return !inv.includes('DBN') && tax >= 0;
      });

      const matchedBookIds = new Set<string>();

      selectedRows = recoData.notInBook.map((r) => {
        const rGstinClean = cleanStr(r.gstin);
        const rInvClean = cleanInv(r.invNo);
        const rTaxable = round2(Number(r.taxable || 0));
        const rIgst = round2(Number(r.igst || 0));
        const rCgst = round2(Number(r.cgst || 0));
        const rSgst = round2(Number(r.sgst || 0));

        // Check against regular books (strictly 1-to-1 matching to prevent double-counting)
        let isEntryDone = false;
        for (const b of regularBooks) {
          const bId = b.id || `${cleanStr(b.gstin || b.g)}_${cleanInv(b.invNo || b.n)}`;
          if (matchedBookIds.has(bId)) continue;

          const bGstinClean = cleanStr(b.gstin || b.g);
          const bInvClean = cleanInv(b.invNo || b.n);
          if (bGstinClean !== rGstinClean || bInvClean !== rInvClean) continue;

          let rawTaxable = typeof b.taxable === 'number' ? b.taxable : (typeof b.t === 'number' ? b.t : 0);
          let rawIgst = typeof b.igst === 'number' ? b.igst : (typeof b.i === 'number' ? b.i : 0);
          let rawCgst = typeof b.cgst === 'number' ? b.cgst : (typeof b.c === 'number' ? b.c : 0);
          let rawSgst = typeof b.sgst === 'number' ? b.sgst : (typeof b.s === 'number' ? b.s : 0);

          // Support amounts stored either directly in rupees or in paise
          let bTaxable = round2(rawTaxable);
          let bIgst = round2(rawIgst);
          let bCgst = round2(rawCgst);
          let bSgst = round2(rawSgst);

          if (Math.abs(bTaxable - rTaxable) > 1.00 && Math.abs(round2(rawTaxable / 100) - rTaxable) <= 1.00) {
            bTaxable = round2(rawTaxable / 100);
            bIgst = round2(rawIgst / 100);
            bCgst = round2(rawCgst / 100);
            bSgst = round2(rawSgst / 100);
          } else if (Math.abs(rawTaxable) >= 100000 && Math.abs(bTaxable - rTaxable) > 1.00) {
            bTaxable = round2(rawTaxable / 100);
            bIgst = round2(rawIgst / 100);
            bCgst = round2(rawCgst / 100);
            bSgst = round2(rawSgst / 100);
          }

          const taxMatch = Math.abs(bTaxable - rTaxable) <= 1.00;
          const igstMatch = Math.abs(bIgst - rIgst) <= 1.00;
          const cgstMatch = Math.abs(bCgst - rCgst) <= 1.00;
          const sgstMatch = Math.abs(bSgst - rSgst) <= 1.00;

          if (taxMatch && igstMatch && cgstMatch && sgstMatch) {
            matchedBookIds.add(bId);
            isEntryDone = true;
            break;
          }
        }

        return {
          ...r,
          status: isEntryDone ? 'Entry Done' : 'Not in Purchase Register'
        };
      });

      totalContext = recoData.summary.notInBookCount + recoData.summary.matchedCount;
      count = recoData.summary.notInBookCount;
    } else if (category === 'not-match') {
      selectedRows = recoData.notMatch.map((r) => ({
        ...r,
        status: 'Not Match'
      }));
      totalContext = recoData.summary.totalCount;
      count = recoData.summary.notMatchCount;
    } else {
      // category === 'not-in-2b': Single Source of Truth from Reco
      selectedRows = recoData.notIn2b.map((r) => ({
        ...r,
        status: 'Not in GSTR-2B'
      }));
      totalContext = recoData.summary.totalCount;
      count = recoData.summary.notIn2bCount;
    }

    const pct = totalContext > 0 ? ((count / totalContext) * 100).toFixed(1) : '0.0';

    return {
      categoryRows: selectedRows,
      totalInvoicesContext: totalContext,
      categoryCount: count,
      matchPercentage: pct
    };
  }, [recoData, category, dataVersion]);

  // Filter by search term
  const filteredRows = useMemo(() => {
    if (!searchTerm.trim()) return categoryRows;
    const term = searchTerm.toLowerCase().trim();
    return categoryRows.filter(
      (r) =>
        r.supplierName.toLowerCase().includes(term) ||
        r.gstin.toLowerCase().includes(term) ||
        r.invNo.toLowerCase().includes(term)
    );
  }, [categoryRows, searchTerm]);

  // Compute 5 summary values for this sheet
  const summaryValues = useMemo(() => {
    let taxable = 0;
    let igst = 0;
    let cgst = 0;
    let sgst = 0;
    filteredRows.forEach((r) => {
      taxable += r.taxable;
      igst += r.igst;
      cgst += r.cgst;
      sgst += r.sgst;
    });
    const totalTax = igst + cgst + sgst;
    const grandTotal = taxable + totalTax;
    return {
      taxable: round2(taxable),
      igst: round2(igst),
      cgst: round2(cgst),
      sgst: round2(sgst),
      totalTax: round2(totalTax),
      grandTotal: round2(grandTotal)
    };
  }, [filteredRows]);

  // Open Audit Modal
  const handleOpenAudit = (item: DetailSheetRow) => {
    setAuditItem(item);
    try {
      const stored = localStorage.getItem('gst_reco_audit_notes');
      const notes = stored ? JSON.parse(stored) : {};
      const key = `${cleanStr(item.gstin)}_${cleanInv(item.invNo)}`;
      const note = notes[item.id] || notes[key];
      setAuditStatus(note?.status || 'Pending');
      setAuditNote(note?.note || '');
    } catch {
      setAuditStatus('Pending');
      setAuditNote('');
    }
  };

  // Save Audit Note
  const handleSaveAudit = () => {
    if (isLocked) {
      showToast('This month is locked. Audit notes cannot be changed.');
      return;
    }
    if (!auditItem) return;
    try {
      const stored = localStorage.getItem('gst_reco_audit_notes');
      const notes = stored ? JSON.parse(stored) : {};
      const key = `${cleanStr(auditItem.gstin)}_${cleanInv(auditItem.invNo)}`;
      const val = { status: auditStatus, note: auditNote, updatedAt: new Date().toLocaleString() };
      notes[auditItem.id] = val;
      notes[key] = val;
      localStorage.setItem('gst_reco_audit_notes', JSON.stringify(notes));
      showToast(`Audit updated for Invoice ${auditItem.invNo}`);
    } catch (e) {
      console.error(e);
      showToast('Failed to save audit note.');
    }
    setAuditItem(null);
  };

  // Visual metadata per category
  const config = useMemo(() => {
    if (category === 'matched') {
      return {
        title: 'Matched Invoices – ITC Verified',
        subTitle: 'Found in both purchase register and GSTR-2B. ITC may be claimed.',
        badgeColor: '#15803d',
        borderColor: '#15803d',
        metric1Label: 'Total Books Invoices',
        metric2Label: 'Matched Invoices',
        metric3Label: 'Match Rate',
        statusPillBg: '#dcfce7',
        statusPillColor: '#15803d',
        statusPillBorder: '#86efac'
      };
    }
    if (category === 'not-in-book') {
      return {
        title: 'Not in Purchase Register',
        subTitle: 'Appears in GSTR-2B but not recorded in your books. Record it.',
        badgeColor: '#1d4ed8',
        borderColor: '#2563eb',
        metric1Label: 'Total 2B Invoices',
        metric2Label: 'Not in Purchase Register',
        metric3Label: '% of 2B Invoices',
        statusPillBg: '#dbeafe',
        statusPillColor: '#1d4ed8',
        statusPillBorder: '#93c5fd'
      };
    }
    if (category === 'not-match') {
      return {
        title: 'Not Match ITC',
        subTitle: 'Found in both Books and GSTR-2B, but with mismatch in tax, taxable, or invoice details.',
        badgeColor: '#dc2626',
        borderColor: '#dc2626',
        metric1Label: 'Total Books Invoices',
        metric2Label: 'Not Match Invoices',
        metric3Label: '% of Books Invoices',
        statusPillBg: '#fee2e2',
        statusPillColor: '#dc2626',
        statusPillBorder: '#fca5a5'
      };
    }
    return {
      title: 'Not in GSTR-2B',
      subTitle: 'Recorded in your books but not reflected in GSTR-2B. Hold ITC.',
      badgeColor: '#c2410c',
      borderColor: '#ea580c',
      metric1Label: 'Total Books Invoices',
      metric2Label: 'Not in GSTR-2B',
      metric3Label: '% of Books Invoices',
      statusPillBg: '#fef3c7',
      statusPillColor: '#b45309',
      statusPillBorder: '#fde68a'
    };
  }, [category]);

  const handleRefresh = async () => {
    setIsRefreshing(true);
    try {
      await Promise.all([
        loadRegisterRows(REGISTER_TABS.ITC_PURCHASE),
        loadRegisterRows(REGISTER_TABS.GSTR2B),
        new Promise((resolve) => setTimeout(resolve, 500))
      ]);
      setDataVersion((v) => v + 1);
      window.dispatchEvent(new Event('storage'));
      window.dispatchEvent(new CustomEvent('register-synced', { detail: { tab: 'all' } }));
      showToast(`${config.title} refreshed`);
    } catch {
      showToast('Data refreshed');
    } finally {
      setIsRefreshing(false);
    }
  };

  const handleExportExcel = () => {
    try {
      const exportData = filteredRows.map((r, i) => {
        if (category === 'not-in-2b') {
          return {
            'S.No': i + 1,
            'Book Month': workingMonth,
            '2B Month': '-',
            'Claim Month': '-',
            'Invoice No.': r.invNo,
            'Supplier GSTIN': r.gstin,
            'Supplier Name': r.supplierName,
            'Reason': 'Not in 2B: supplier ne return file nahi kiya (possible quarterly filer)',
            'Status': 'Not in GSTR-2B',
            'Books Taxable (₹)': r.taxable,
            'Books IGST (₹)': r.igst,
            'Books CGST (₹)': r.cgst,
            'Books SGST (₹)': r.sgst,
            'Books Total (₹)': r.total,
            '2B Taxable (₹)': '-',
            '2B IGST (₹)': '-',
            '2B CGST (₹)': '-',
            '2B SGST (₹)': '-',
            '2B Total (₹)': '-'
          };
        }
        return {
          '#': i + 1,
          'Status': r.status,
          '2B Month': r.month2b,
          'Date': r.date,
          'Invoice No.': r.invNo,
          'Supplier GSTIN': r.gstin,
          'Supplier Name': r.supplierName,
          'Taxable (₹)': r.taxable,
          'IGST (₹)': r.igst,
          'CGST (₹)': r.cgst,
          'SGST (₹)': r.sgst,
          'Total (₹)': r.total
        };
      });

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, config.title.slice(0, 31));
      const filename = `${config.title.replace(/[^A-Za-z0-9_-]/g, '_')}_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(wb, filename);
      showToast(`${config.title} Excel downloaded successfully!`);
    } catch (e) {
      console.error(e);
      showToast('Failed to export Excel.');
    }
  };

  return (
    <section className="detail-sheet-page" style={{ padding: '4px 0 24px 0' }}>
      {/* Top Verification Check Strip */}
      {(category === 'not-in-2b' || category === 'not-match') && (
        <div
          style={{
            background: '#ffffff',
            border: '1.5px solid #e2e8f0',
            borderLeft: `4px solid ${categoryRows.length === (category === 'not-match' ? recoData.summary.notMatchCount : recoData.summary.notIn2bCount) ? '#16a34a' : '#dc2626'}`,
            padding: '8px 16px',
            marginBottom: '12px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            fontSize: '12.5px',
            fontWeight: 700,
            color: '#1e293b'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span>
              Reco: Not Match {recoData.summary.notMatchCount} + Not in 2B {recoData.summary.notIn2bCount} = {recoData.summary.notMatchCount + recoData.summary.notIn2bCount} | Is sheet: {categoryRows.length}
            </span>
            {categoryRows.length === (category === 'not-match' ? recoData.summary.notMatchCount : recoData.summary.notIn2bCount) ? (
              <span style={{ color: '#16a34a', fontSize: '13.5px' }}>✔</span>
            ) : (
              <span style={{ color: '#dc2626', fontSize: '13.5px' }}>⚠️</span>
            )}
          </div>
          <span style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>
            Working Month: {workingMonth}
          </span>
        </div>
      )}

      {/* 2. Total Card Box (with left colored border & 4 metrics) */}
      <div
        style={{
          background: '#ffffff',
          border: '1.5px solid #e2e8f0',
          borderLeft: `5px solid ${config.borderColor}`,
          padding: '16px 20px',
          marginBottom: '14px',
          borderRadius: 0,
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
          gap: '16px'
        }}
      >
        <div>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
            {config.metric1Label}
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            {totalInvoicesContext}
          </div>
        </div>

        <div>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
            {config.metric2Label}
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: config.badgeColor, marginTop: '4px' }}>
            {categoryCount}
          </div>
        </div>

        <div>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
            {config.metric3Label}
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            {matchPercentage}%
          </div>
        </div>

        <div>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
            Total Taxable Value
          </div>
          <div style={{ fontSize: '20px', fontWeight: 800, color: '#350b5e', marginTop: '6px' }}>
            {formatInr(summaryValues.taxable)}
          </div>
        </div>
      </div>

      {/* 3. 5 Summary Metric Boxes */}
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))',
          gap: '10px',
          marginBottom: '16px'
        }}
      >
        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>Taxable Value</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            {formatInr(summaryValues.taxable)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>IGST</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            {formatInr(summaryValues.igst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>CGST</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            {formatInr(summaryValues.cgst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>SGST</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            {formatInr(summaryValues.sgst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>Total Tax Amount</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#350b5e', marginTop: '4px' }}>
            {formatInr(summaryValues.totalTax)}
          </div>
        </div>
      </div>

      {/* 4. Main Category Table Card Container with Controls & Table */}
      <div className="card gstr1-card" style={{ width: '100%', background: '#ffffff', border: '1px solid var(--line)', padding: '20px', borderRadius: 0, boxShadow: '0 1px 3px rgba(0, 0, 0, 0.05)', boxSizing: 'border-box' }}>
        {/* Controls: Search Input & Excel Export Button */}
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            gap: '12px',
            marginBottom: '16px',
            flexWrap: 'wrap'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flex: 1, minWidth: '240px' }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="11" cy="11" r="8"></circle>
              <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
            </svg>
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Search Supplier Name, GSTIN, Invoice No..."
              style={{
                width: '100%',
                maxWidth: '380px',
                border: '1px solid #cbd5e1',
                padding: '6px 12px',
                fontSize: '12.5px',
                outline: 'none',
                borderRadius: 0
              }}
            />
            {searchTerm && (
              <button
                type="button"
                onClick={() => setSearchTerm('')}
                style={{
                  border: 'none',
                  background: 'transparent',
                  cursor: 'pointer',
                  color: '#64748b',
                  fontSize: '13px'
                }}
              >
                ✕ Clear
              </button>
            )}
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <span style={{ fontSize: '12px', color: '#64748b', fontWeight: 600 }}>
              Showing {filteredRows.length} of {categoryRows.length} records
            </span>
            <button
              type="button"
              className="btn-gstr1 btn-refresh"
              onClick={handleRefresh}
              disabled={isRefreshing || isLocked}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px',
                padding: '6px 14px',
                fontSize: '12px',
                fontWeight: 700,
                cursor: (isRefreshing || isLocked) ? 'not-allowed' : 'pointer',
                opacity: isLocked ? 0.6 : 1
              }}
              title={isLocked ? 'Month is locked - refresh disabled' : 'Reload latest reconciliation data'}
            >
              <svg
                width="15"
                height="15"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.2"
                strokeLinecap="round"
                strokeLinejoin="round"
                className={isRefreshing ? 'spin-icon' : ''}
              >
                <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"></path>
              </svg>
              <span>{isRefreshing ? 'Refreshing...' : 'Refresh'}</span>
            </button>
            <button
              type="button"
              onClick={handleExportExcel}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px',
                background: '#15803d',
                color: '#ffffff',
                border: 'none',
                padding: '6px 14px',
                fontSize: '12px',
                fontWeight: 700,
                cursor: 'pointer',
                borderRadius: 0
              }}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="7 10 12 15 17 10" />
                <line x1="12" y1="15" x2="12" y2="3" />
              </svg>
              Export Excel
            </button>
          </div>
        </div>

        {/* 5. Main Category Table */}
        <div className="detail-sheet-table-wrapper" style={{ overflowX: 'auto', border: '1px solid #e2e8f0', background: '#ffffff', width: '100%' }}>
          {category === 'not-in-2b' || category === 'not-match' ? (
            /* ======================================================== */
            /* 20 COLUMNS TABLE FOR NOT IN GSTR-2B & NOT MATCH SHEETS     */
            /* ======================================================== */
            <table className="detail-sheet-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', color: '#1E1747' }}>
              <thead>
                {/* Group Row */}
                <tr style={{ background: '#2B0A4F', color: '#ffffff', textAlign: 'center' }}>
                  <th
                    colSpan={9}
                    style={{
                      padding: '7px 10px',
                      fontSize: '13px',
                      fontWeight: 800,
                      letterSpacing: '0.05em',
                      textTransform: 'uppercase',
                      textAlign: 'center',
                      borderRight: '1px solid rgba(255, 255, 255, 0.25)',
                      borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                      background: '#2B0A4F',
                      fontFamily: "'Bricolage Grotesque', sans-serif"
                    }}
                  >
                    Invoice &amp; Status Details
                  </th>
                  <th
                    colSpan={5}
                    style={{
                      padding: '7px 10px',
                      fontSize: '13px',
                      fontWeight: 800,
                      letterSpacing: '0.05em',
                      textTransform: 'uppercase',
                      textAlign: 'center',
                      background: '#240742',
                      borderRight: '1px solid rgba(255, 255, 255, 0.25)',
                      borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                      fontFamily: "'Bricolage Grotesque', sans-serif"
                    }}
                  >
                    Books (Purchase Register)
                  </th>
                  <th
                    colSpan={5}
                    style={{
                      padding: '7px 10px',
                      fontSize: '13px',
                      fontWeight: 800,
                      letterSpacing: '0.05em',
                      textTransform: 'uppercase',
                      textAlign: 'center',
                      background: '#155E75',
                      borderRight: '1px solid rgba(255, 255, 255, 0.25)',
                      borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                      fontFamily: "'Bricolage Grotesque', sans-serif"
                    }}
                  >
                    Govt GSTR-2B
                  </th>
                  <th
                    rowSpan={2}
                    style={{
                      padding: '10px 12px',
                      fontSize: '12px',
                      fontWeight: 800,
                      letterSpacing: '0.05em',
                      textTransform: 'uppercase',
                      textAlign: 'center',
                      minWidth: '130px',
                      width: '130px',
                      background: '#2B0A4F',
                      color: '#ffffff',
                      borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                      fontFamily: "'Bricolage Grotesque', sans-serif"
                    }}
                  >
                    Actions
                  </th>
                </tr>

                {/* Column Headers */}
                <tr
                  style={{
                    background: '#4A1A7A',
                    color: '#ffffff',
                    fontSize: '11.5px',
                    fontWeight: 700,
                    textTransform: 'uppercase',
                    letterSpacing: '0.03em',
                    whiteSpace: 'nowrap',
                    fontFamily: "'Bricolage Grotesque', sans-serif"
                  }}
                >
                  <th style={{ width: '45px', minWidth: '45px', padding: '6px 8px', textAlign: 'center', background: '#4A1A7A', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>#</th>
                  <th style={{ width: '95px', minWidth: '95px', padding: '6px 8px', textAlign: 'center', background: '#4A1A7A', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Book Month</th>
                  <th style={{ width: '85px', minWidth: '85px', padding: '6px 8px', textAlign: 'center', background: '#4A1A7A', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>2B Month</th>
                  <th style={{ width: '95px', minWidth: '95px', padding: '6px 8px', textAlign: 'center', background: '#4A1A7A', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Claim Month</th>
                  <th style={{ width: '130px', minWidth: '130px', padding: '6px 10px', textAlign: 'left', background: '#4A1A7A', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Invoice No.</th>
                  <th style={{ width: '145px', minWidth: '145px', padding: '6px 10px', textAlign: 'left', background: '#4A1A7A', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Supplier GSTIN</th>
                  <th style={{ width: '200px', minWidth: '200px', padding: '6px 10px', textAlign: 'left', background: '#4A1A7A', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Supplier Name</th>
                  <th style={{ width: '220px', minWidth: '220px', padding: '6px 10px', textAlign: 'left', background: '#4A1A7A', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Reason</th>
                  <th style={{ width: '110px', minWidth: '110px', padding: '6px 8px', textAlign: 'center', background: '#4A1A7A', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Status</th>

                  {/* Books 5 Columns */}
                  <th style={{ width: '110px', minWidth: '110px', padding: '6px 8px', textAlign: 'right', background: '#3b1263', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Taxable</th>
                  <th style={{ width: '95px', minWidth: '95px', padding: '6px 8px', textAlign: 'right', background: '#3b1263', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>IGST</th>
                  <th style={{ width: '95px', minWidth: '95px', padding: '6px 8px', textAlign: 'right', background: '#3b1263', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>CGST</th>
                  <th style={{ width: '95px', minWidth: '95px', padding: '6px 8px', textAlign: 'right', background: '#3b1263', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>SGST</th>
                  <th style={{ width: '115px', minWidth: '115px', padding: '6px 8px', textAlign: 'right', background: '#3b1263', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Total</th>

                  {/* Govt 2B 5 Columns */}
                  <th style={{ width: '110px', minWidth: '110px', padding: '6px 8px', textAlign: 'right', background: '#0E7490', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Taxable</th>
                  <th style={{ width: '95px', minWidth: '95px', padding: '6px 8px', textAlign: 'right', background: '#0E7490', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>IGST</th>
                  <th style={{ width: '95px', minWidth: '95px', padding: '6px 8px', textAlign: 'right', background: '#0E7490', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>CGST</th>
                  <th style={{ width: '95px', minWidth: '95px', padding: '6px 8px', textAlign: 'right', background: '#0E7490', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>SGST</th>
                  <th style={{ width: '115px', minWidth: '115px', padding: '6px 8px', textAlign: 'right', background: '#0E7490', borderBottom: '1px solid rgba(255, 255, 255, 0.12)' }}>Total</th>
                </tr>
              </thead>
              <tbody>
                {filteredRows.length === 0 ? (
                  <tr>
                    <td colSpan={20} style={{ padding: '42px 16px', textAlign: 'center', color: '#6B6690', fontSize: '14px', fontWeight: 600 }}>
                      No invoices found matching selected criteria.
                    </td>
                  </tr>
                ) : (
                  filteredRows.map((row, idx) => (
                    <tr
                      key={row.id || idx}
                      style={{
                        background: idx % 2 === 0 ? '#FAF8FD' : '#ffffff',
                        borderBottom: '1px solid #e2e8f0',
                        height: '42px'
                      }}
                      onMouseEnter={(e) => (e.currentTarget.style.background = '#F3EBFE')}
                      onMouseLeave={(e) => (e.currentTarget.style.background = idx % 2 === 0 ? '#FAF8FD' : '#ffffff')}
                    >
                      <td style={{ width: '45px', minWidth: '45px', textAlign: 'center', fontWeight: 600, color: '#64748b' }}>
                        {idx + 1}
                      </td>
                      <td style={{ width: '95px', minWidth: '95px', textAlign: 'center', fontWeight: 700, color: '#334155' }}>
                        {workingMonth}
                      </td>
                      <td style={{ width: '85px', minWidth: '85px', textAlign: 'center', color: category === 'not-match' ? '#1E1747' : '#94a3b8', fontWeight: category === 'not-match' ? 600 : 400 }}>
                        {category === 'not-match' ? (row.month2b || workingMonth) : '-'}
                      </td>
                      <td style={{ width: '95px', minWidth: '95px', textAlign: 'center', color: '#94a3b8' }}>
                        -
                      </td>
                      <td style={{ width: '130px', minWidth: '130px', textAlign: 'left', fontWeight: 700, color: '#1E1747', fontSize: '13px' }}>
                        {row.invNo}
                      </td>
                      <td style={{ width: '145px', minWidth: '145px', textAlign: 'left', fontFamily: 'monospace', fontSize: '12px' }}>
                        {row.gstin}
                      </td>
                      <td style={{ width: '200px', minWidth: '200px', textAlign: 'left', fontWeight: 600, color: '#1E1747' }}>
                        {row.supplierName}
                      </td>
                      <td style={{ width: '220px', minWidth: '220px', textAlign: 'left', fontSize: '12px', color: '#b91c1c', fontWeight: 600, whiteSpace: 'normal', wordBreak: 'break-word', lineHeight: 1.4 }}>
                        {row.reason || (category === 'not-in-2b' ? 'Not in 2B: supplier ne return file nahi kiya (possible quarterly filer)' : 'Not Match: Tax Mismatch')}
                      </td>
                      <td style={{ width: '110px', minWidth: '110px', textAlign: 'center' }}>
                        <span
                          style={{
                            display: 'inline-block',
                            padding: '3px 8px',
                            fontSize: '11px',
                            fontWeight: 700,
                            background: category === 'not-match' ? '#fee2e2' : '#fef3c7',
                            color: category === 'not-match' ? '#dc2626' : '#b45309',
                            border: `1.5px solid ${category === 'not-match' ? '#fca5a5' : '#fde68a'}`,
                            borderRadius: 0,
                            whiteSpace: 'nowrap'
                          }}
                        >
                          {category === 'not-match' ? 'Not Match' : 'Not in GSTR-2B'}
                        </span>
                      </td>

                      {/* Books Amounts */}
                      <td style={{ width: '110px', minWidth: '110px', textAlign: 'right', background: '#F0F7FF', fontWeight: 600, color: '#1E1747' }}>
                        {renderAmountCell(row.taxable)}
                      </td>
                      <td style={{ width: '95px', minWidth: '95px', textAlign: 'right', background: '#F0F7FF', fontWeight: 600, color: '#1E1747' }}>
                        {renderAmountCell(row.igst)}
                      </td>
                      <td style={{ width: '95px', minWidth: '95px', textAlign: 'right', background: '#F0F7FF', fontWeight: 600, color: '#1E1747' }}>
                        {renderAmountCell(row.cgst)}
                      </td>
                      <td style={{ width: '95px', minWidth: '95px', textAlign: 'right', background: '#F0F7FF', fontWeight: 600, color: '#1E1747' }}>
                        {renderAmountCell(row.sgst)}
                      </td>
                      <td style={{ width: '115px', minWidth: '115px', textAlign: 'right', background: '#F0F7FF', fontWeight: 700, color: '#7F1D1D' }}>
                        {renderAmountCell(row.total, true)}
                      </td>

                      {/* Govt 2B Amounts */}
                      {category === 'not-match' ? (
                        <>
                          <td style={{ width: '110px', minWidth: '110px', textAlign: 'right', background: Math.abs(row.taxable - (row.gTaxable ?? row.taxable)) > 1 ? '#fee2e2' : '#ECFEFF', fontWeight: 600, color: Math.abs(row.taxable - (row.gTaxable ?? row.taxable)) > 1 ? '#dc2626' : '#1E1747' }}>
                            {renderAmountCell(row.gTaxable ?? 0)}
                            {Math.abs(row.taxable - (row.gTaxable ?? row.taxable)) > 1 && ((row.gTaxable ?? 0) > row.taxable ? ' ▲' : ' ▼')}
                          </td>
                          <td style={{ width: '95px', minWidth: '95px', textAlign: 'right', background: Math.abs(row.igst - (row.gIgst ?? row.igst)) > 1 ? '#fee2e2' : '#ECFEFF', fontWeight: 600, color: Math.abs(row.igst - (row.gIgst ?? row.igst)) > 1 ? '#dc2626' : '#1E1747' }}>
                            {renderAmountCell(row.gIgst ?? 0)}
                            {Math.abs(row.igst - (row.gIgst ?? row.igst)) > 1 && ((row.gIgst ?? 0) > row.igst ? ' ▲' : ' ▼')}
                          </td>
                          <td style={{ width: '95px', minWidth: '95px', textAlign: 'right', background: Math.abs(row.cgst - (row.gCgst ?? row.cgst)) > 1 ? '#fee2e2' : '#ECFEFF', fontWeight: 600, color: Math.abs(row.cgst - (row.gCgst ?? row.cgst)) > 1 ? '#dc2626' : '#1E1747' }}>
                            {renderAmountCell(row.gCgst ?? 0)}
                            {Math.abs(row.cgst - (row.gCgst ?? row.cgst)) > 1 && ((row.gCgst ?? 0) > row.cgst ? ' ▲' : ' ▼')}
                          </td>
                          <td style={{ width: '95px', minWidth: '95px', textAlign: 'right', background: Math.abs(row.sgst - (row.gSgst ?? row.sgst)) > 1 ? '#fee2e2' : '#ECFEFF', fontWeight: 600, color: Math.abs(row.sgst - (row.gSgst ?? row.sgst)) > 1 ? '#dc2626' : '#1E1747' }}>
                            {renderAmountCell(row.gSgst ?? 0)}
                            {Math.abs(row.sgst - (row.gSgst ?? row.sgst)) > 1 && ((row.gSgst ?? 0) > row.sgst ? ' ▲' : ' ▼')}
                          </td>
                          <td style={{ width: '115px', minWidth: '115px', textAlign: 'right', background: Math.abs(row.total - (row.gTotal ?? row.total)) > 1 ? '#fee2e2' : '#ECFEFF', fontWeight: 700, color: Math.abs(row.total - (row.gTotal ?? row.total)) > 1 ? '#dc2626' : '#7F1D1D' }}>
                            {renderAmountCell(row.gTotal ?? 0, true)}
                            {Math.abs(row.total - (row.gTotal ?? row.total)) > 1 && ((row.gTotal ?? 0) > row.total ? ' ▲' : ' ▼')}
                          </td>
                        </>
                      ) : (
                        <>
                          <td style={{ width: '110px', minWidth: '110px', textAlign: 'right', background: '#ECFEFF', color: '#94a3b8' }}>-</td>
                          <td style={{ width: '95px', minWidth: '95px', textAlign: 'right', background: '#ECFEFF', color: '#94a3b8' }}>-</td>
                          <td style={{ width: '95px', minWidth: '95px', textAlign: 'right', background: '#ECFEFF', color: '#94a3b8' }}>-</td>
                          <td style={{ width: '95px', minWidth: '95px', textAlign: 'right', background: '#ECFEFF', color: '#94a3b8' }}>-</td>
                          <td style={{ width: '115px', minWidth: '115px', textAlign: 'right', background: '#ECFEFF', color: '#94a3b8' }}>-</td>
                        </>
                      )}

                      {/* Actions: View and Audit */}
                      <td style={{ width: '130px', minWidth: '130px', padding: '6px 10px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                        <div style={{ display: 'inline-flex', gap: '4px' }}>
                          <button
                            type="button"
                            onClick={() => setViewItem(row)}
                            style={{
                              background: '#f1f5f9',
                              color: '#334155',
                              border: '1px solid #cbd5e1',
                              padding: '3px 7px',
                              fontSize: '11px',
                              fontWeight: 700,
                              cursor: 'pointer',
                              borderRadius: 0
                            }}
                            title="View Invoice Details"
                          >
                            View
                          </button>
                          <button
                            type="button"
                            onClick={() => handleOpenAudit(row)}
                            style={{
                              background: '#f1f5f9',
                              color: '#334155',
                              border: '1px solid #cbd5e1',
                              padding: '3px 7px',
                              fontSize: '11px',
                              fontWeight: 700,
                              cursor: 'pointer',
                              borderRadius: 0
                            }}
                            title="Add / View Audit Note"
                          >
                            Audit
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
              {filteredRows.length > 0 && (
                <tfoot>
                  <tr style={{ background: '#2B0A4F', color: '#ffffff', fontWeight: 800, fontSize: '12px' }}>
                    <td colSpan={9} style={{ padding: '8px 12px', textAlign: 'right', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                      TOTAL ({filteredRows.length} INVOICES)
                    </td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#240742' }}>
                      {renderAmountCell(summaryValues.taxable)}
                    </td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#240742' }}>
                      {renderAmountCell(summaryValues.igst)}
                    </td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#240742' }}>
                      {renderAmountCell(summaryValues.cgst)}
                    </td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#240742' }}>
                      {renderAmountCell(summaryValues.sgst)}
                    </td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#350B5E', color: '#F3E8FF' }}>
                      {renderAmountCell(summaryValues.grandTotal, true)}
                    </td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#155E75', color: '#94a3b8' }}>-</td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#155E75', color: '#94a3b8' }}>-</td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#155E75', color: '#94a3b8' }}>-</td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#155E75', color: '#94a3b8' }}>-</td>
                    <td style={{ padding: '8px 8px', textAlign: 'right', background: '#0E7490', color: '#94a3b8' }}>-</td>
                    <td style={{ padding: '8px 8px', textAlign: 'center' }}>-</td>
                  </tr>
                </tfoot>
              )}
            </table>
          ) : (
            /* ======================================================== */
            /* STANDARD TABLE FOR MATCHED & NOT-IN-BOOK SHEETS          */
            /* ======================================================== */
            <table className="detail-sheet-table">
              <thead>
                <tr>
                  <th style={{ width: '45px', minWidth: '45px', textAlign: 'center' }}>#</th>
                  <th style={{ width: '120px', minWidth: '120px', textAlign: 'center' }}>STATUS</th>
                  <th style={{ width: '100px', minWidth: '100px', textAlign: 'center' }}>2B MONTH</th>
                  <th style={{ width: '110px', minWidth: '110px', textAlign: 'left' }}>DATE</th>
                  <th style={{ width: '165px', minWidth: '165px', textAlign: 'left' }}>INVOICE NO.</th>
                  <th style={{ width: '160px', minWidth: '160px', textAlign: 'left' }}>SUPPLIER GSTIN</th>
                  <th style={{ width: '250px', minWidth: '250px', textAlign: 'left' }}>SUPPLIER NAME</th>
                  <th style={{ width: '135px', minWidth: '135px', textAlign: 'right' }}>TAXABLE (₹)</th>
                  <th style={{ width: '115px', minWidth: '115px', textAlign: 'right' }}>IGST (₹)</th>
                  <th style={{ width: '115px', minWidth: '115px', textAlign: 'right' }}>CGST (₹)</th>
                  <th style={{ width: '115px', minWidth: '115px', textAlign: 'right' }}>SGST (₹)</th>
                  <th style={{ width: '135px', minWidth: '135px', textAlign: 'right' }}>TOTAL (₹)</th>
                </tr>
              </thead>
              <tbody>
                {filteredRows.length === 0 ? (
                  <tr>
                    <td colSpan={12} style={{ padding: '42px 16px', textAlign: 'center', color: '#6B6690', fontSize: '14px', fontWeight: 600 }}>
                      No invoices found matching selected criteria.
                    </td>
                  </tr>
                ) : (
                  filteredRows.map((row, idx) => (
                    <tr key={row.id || idx}>
                      <td style={{ width: '45px', minWidth: '45px', textAlign: 'center', fontWeight: 700, color: '#1E1747' }}>
                        {idx + 1}
                      </td>
                      <td style={{ width: '120px', minWidth: '120px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                        <span
                          style={{
                            display: 'inline-block',
                            padding: '4px 9px',
                            fontSize: '11.5px',
                            fontWeight: 700,
                            background: row.status === 'Entry Done' ? '#dcfce7' : config.statusPillBg,
                            color: row.status === 'Entry Done' ? '#15803d' : config.statusPillColor,
                            border: `1.5px solid ${row.status === 'Entry Done' ? '#86efac' : config.statusPillBorder}`,
                            borderRadius: 0,
                            whiteSpace: 'nowrap'
                          }}
                        >
                          {row.status}
                        </span>
                      </td>
                      <td style={{ width: '100px', minWidth: '100px', textAlign: 'center', color: '#1E1747', fontWeight: 600 }}>
                        {row.month2b}
                      </td>
                      <td style={{ width: '110px', minWidth: '110px', color: '#1E1747', whiteSpace: 'nowrap', fontWeight: 500 }}>
                        {row.date}
                      </td>
                      <td style={{ width: '165px', minWidth: '165px', textAlign: 'left', fontWeight: 700, color: '#1E1747', fontSize: '14px' }}>
                        {row.invNo}
                      </td>
                      <td style={{ width: '160px', minWidth: '160px', fontFamily: 'monospace, monospace', fontSize: '12.5px', color: '#1E1747', fontWeight: 600 }}>
                        {row.gstin}
                      </td>
                      <td style={{ width: '250px', minWidth: '250px', textAlign: 'left', fontWeight: 600, color: '#1E1747', wordBreak: 'break-word', whiteSpace: 'normal' }}>
                        {row.supplierName}
                      </td>
                      <td className="amt-cell" style={{ width: '135px', minWidth: '135px' }}>
                        {renderAmountCell(row.taxable)}
                      </td>
                      <td className="amt-cell" style={{ width: '115px', minWidth: '115px' }}>
                        {renderAmountCell(row.igst)}
                      </td>
                      <td className="amt-cell" style={{ width: '115px', minWidth: '115px' }}>
                        {renderAmountCell(row.cgst)}
                      </td>
                      <td className="amt-cell" style={{ width: '115px', minWidth: '115px' }}>
                        {renderAmountCell(row.sgst)}
                      </td>
                      <td className="amt-cell" style={{ width: '135px', minWidth: '135px' }}>
                        {renderAmountCell(row.total, true)}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {/* VIEW MODAL (FOR NOT IN 2B) */}
      {viewItem && (
        <div
          style={{
            position: 'fixed',
            inset: 0,
            background: 'rgba(0,0,0,0.5)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 1000
          }}
          onClick={() => setViewItem(null)}
        >
          <div
            style={{
              background: '#ffffff',
              width: '90%',
              maxWidth: '650px',
              padding: '24px',
              borderRadius: 0,
              boxShadow: '0 10px 25px rgba(0,0,0,0.2)'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1.5px solid #e2e8f0', paddingBottom: '12px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', color: '#241a3d', fontWeight: 800 }}>
                Invoice Detail: {viewItem.invNo}
              </h3>
              <button
                type="button"
                onClick={() => setViewItem(null)}
                style={{ background: 'transparent', border: 'none', fontSize: '18px', cursor: 'pointer', color: '#64748b' }}
              >
                ✕
              </button>
            </div>

            <div style={{ marginTop: '16px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', fontSize: '13px' }}>
              <div><strong>Supplier Name:</strong> {viewItem.supplierName}</div>
              <div><strong>Supplier GSTIN:</strong> <code style={{ color: '#4c1d95' }}>{viewItem.gstin}</code></div>
              <div><strong>Invoice Date:</strong> {viewItem.date}</div>
              <div><strong>Book Month:</strong> {workingMonth}</div>
              <div><strong>Status:</strong> <span style={{ color: viewItem.status === 'Not Match' ? '#dc2626' : '#b45309', fontWeight: 700 }}>{viewItem.status}</span></div>
              <div><strong>ITC Value:</strong> <span style={{ fontWeight: 800, color: '#c2410c' }}>{formatInr(viewItem.igst + viewItem.cgst + viewItem.sgst)}</span></div>
            </div>

            <div style={{ marginTop: '14px', background: viewItem.status === 'Not Match' ? '#fef2f2' : '#fff7ed', padding: '10px 12px', border: `1px solid ${viewItem.status === 'Not Match' ? '#fecaca' : '#fed7aa'}`, fontSize: '12.5px' }}>
              <strong>Status Reason:</strong>
              <div style={{ color: viewItem.status === 'Not Match' ? '#dc2626' : '#c2410c', marginTop: '4px', fontWeight: 700 }}>
                {viewItem.reason || (viewItem.status === 'Not Match' ? 'Not Match: Tax Mismatch' : 'Not in 2B: supplier ne return file nahi kiya (possible quarterly filer)')}
              </div>
            </div>

            {/* Books Tax Details */}
            <div style={{ marginTop: '16px', border: '1px solid #e2e8f0' }}>
              <div style={{ background: '#2B0A4F', color: '#ffffff', padding: '8px 12px', fontWeight: 700, fontSize: '12px', letterSpacing: '0.04em' }}>
                BOOKS (PURCHASE REGISTER) TAX DETAILS
              </div>
              <div style={{ padding: '14px', fontSize: '13px', background: '#FAF5FF' }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '12px' }}>
                  <div>Taxable: <strong>{formatInr(viewItem.taxable)}</strong></div>
                  <div>IGST: <strong>{formatInr(viewItem.igst)}</strong></div>
                  <div>CGST: <strong>{formatInr(viewItem.cgst)}</strong></div>
                  <div>SGST: <strong>{formatInr(viewItem.sgst)}</strong></div>
                </div>
                <div style={{ marginTop: '10px', borderTop: '1px solid #cbd5e1', paddingTop: '8px', color: '#7F1D1D', fontWeight: 800, fontSize: '14px' }}>
                  Total Invoice Value: {formatInr(viewItem.total)}
                </div>
              </div>
            </div>

            {/* Govt GSTR-2B Details */}
            <div style={{ marginTop: '14px', border: '1px solid #e2e8f0' }}>
              <div style={{ background: '#155E75', color: '#ffffff', padding: '8px 12px', fontWeight: 700, fontSize: '12px', letterSpacing: '0.04em' }}>
                GOVT GSTR-2B DETAILS
              </div>
              {viewItem.status === 'Not Match' ? (
                <div style={{ padding: '14px', fontSize: '13px', background: '#ECFEFF' }}>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '12px' }}>
                    <div>Taxable: <strong>{formatInr(viewItem.gTaxable ?? 0)}</strong></div>
                    <div>IGST: <strong>{formatInr(viewItem.gIgst ?? 0)}</strong></div>
                    <div>CGST: <strong>{formatInr(viewItem.gCgst ?? 0)}</strong></div>
                    <div>SGST: <strong>{formatInr(viewItem.gSgst ?? 0)}</strong></div>
                  </div>
                  <div style={{ marginTop: '10px', borderTop: '1px solid #cbd5e1', paddingTop: '8px', color: '#0E7490', fontWeight: 800, fontSize: '14px' }}>
                    Total 2B Value: {formatInr(viewItem.gTotal ?? 0)}
                  </div>
                </div>
              ) : (
                <div style={{ padding: '14px', fontSize: '13px', background: '#ECFEFF', color: '#0E7490', fontWeight: 600 }}>
                  ⚠️ Not found in Govt GSTR-2B. Ineligible for ITC claim until reflected in Govt portal.
                </div>
              )}
            </div>

            <div style={{ marginTop: '20px', textAlign: 'right' }}>
              <button
                type="button"
                onClick={() => setViewItem(null)}
                style={{
                  background: '#2B0A4F',
                  color: '#ffffff',
                  border: 'none',
                  padding: '8px 20px',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  borderRadius: 0
                }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* AUDIT MODAL (FOR NOT IN 2B) */}
      {auditItem && (
        <div
          style={{
            position: 'fixed',
            inset: 0,
            background: 'rgba(0,0,0,0.5)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 1000
          }}
          onClick={() => setAuditItem(null)}
        >
          <div
            style={{
              background: '#ffffff',
              width: '90%',
              maxWidth: '500px',
              padding: '24px',
              borderRadius: 0,
              boxShadow: '0 10px 25px rgba(0,0,0,0.2)'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1.5px solid #e2e8f0', paddingBottom: '12px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', color: '#241a3d', fontWeight: 800 }}>
                Audit Invoice: {auditItem.invNo}
              </h3>
              <button
                type="button"
                onClick={() => setAuditItem(null)}
                style={{ background: 'transparent', border: 'none', fontSize: '18px', cursor: 'pointer', color: '#64748b' }}
              >
                ✕
              </button>
            </div>

            <div style={{ marginTop: '16px', fontSize: '13px', color: '#475569' }}>
              <div><strong>Supplier:</strong> {auditItem.supplierName} ({auditItem.gstin})</div>
              <div style={{ marginTop: '4px' }}><strong>Amount:</strong> {formatInr(auditItem.total)}</div>
            </div>

            <div style={{ marginTop: '16px' }}>
              <label style={{ display: 'block', fontSize: '12.5px', fontWeight: 700, color: '#334155', marginBottom: '6px' }}>
                Audit Status:
              </label>
              <select
                value={auditStatus}
                onChange={(e) => setAuditStatus(e.target.value)}
                style={{
                  width: '100%',
                  border: '1px solid #cbd5e1',
                  padding: '8px 12px',
                  fontSize: '13px',
                  borderRadius: 0,
                  outline: 'none',
                  background: '#f8fafc'
                }}
              >
                <option value="Pending">Pending</option>
                <option value="Verified">Verified</option>
                <option value="Query">Query (Supplier Follow-up)</option>
              </select>
            </div>

            <div style={{ marginTop: '14px' }}>
              <label style={{ display: 'block', fontSize: '12.5px', fontWeight: 700, color: '#334155', marginBottom: '6px' }}>
                Audit Observation / Note:
              </label>
              <textarea
                value={auditNote}
                onChange={(e) => setAuditNote(e.target.value)}
                placeholder="Enter audit remarks or follow-up note..."
                rows={4}
                style={{
                  width: '100%',
                  border: '1px solid #cbd5e1',
                  padding: '8px 12px',
                  fontSize: '13px',
                  borderRadius: 0,
                  outline: 'none',
                  boxSizing: 'border-box'
                }}
              />
            </div>

            <div style={{ marginTop: '20px', display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
              <button
                type="button"
                onClick={() => setAuditItem(null)}
                style={{
                  background: '#f1f5f9',
                  color: '#334155',
                  border: '1px solid #cbd5e1',
                  padding: '8px 18px',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  borderRadius: 0
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleSaveAudit}
                style={{
                  background: '#2B0A4F',
                  color: '#ffffff',
                  border: 'none',
                  padding: '8px 22px',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  borderRadius: 0
                }}
              >
                Save Audit
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
