import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  syncOldItcPool,
  moveToNotClaim,
  updateOldItcAudit,
  OldItcItem
} from '../services/oldItcService';
import { cleanInv } from '../services/recoSyncService';

interface OldItcDetailSheetProps {
  onBack: () => void;
  showToast: (msg: string) => void;
  workingMonth?: string;
}

function formatInr(val: number): string {
  if (isNaN(val)) return '0.00';
  return val.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

export default function OldItcDetailSheet({
  onBack,
  showToast,
  workingMonth = 'Oct-2026'
}: OldItcDetailSheetProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [monthFilter, setMonthFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'Pending' | 'Matched - Claim'>('ALL');
  const [poolVersion, setPoolVersion] = useState(0);

  // Selected invoice for View modal
  const [viewItem, setViewItem] = useState<OldItcItem | null>(null);

  // Selected invoice for Audit modal
  const [auditItem, setAuditItem] = useState<OldItcItem | null>(null);
  const [auditStatus, setAuditStatus] = useState<string>('Pending');
  const [auditNote, setAuditNote] = useState<string>('');

  // Sync Old ITC pool
  const syncResult = useMemo(() => {
    return syncOldItcPool(workingMonth);
  }, [workingMonth, poolVersion]);

  // Listen to old-itc-updated events
  useEffect(() => {
    const handleUpdate = () => setPoolVersion((v) => v + 1);
    window.addEventListener('old-itc-updated', handleUpdate);
    return () => window.removeEventListener('old-itc-updated', handleUpdate);
  }, []);

  // Filter pool items
  const filteredItems = useMemo(() => {
    return syncResult.oldItcItems.filter((item) => {
      // Status filter
      if (statusFilter !== 'ALL' && item.status !== statusFilter) {
        return false;
      }

      // Month filter
      if (monthFilter !== 'ALL' && item.originalMonth !== monthFilter) {
        return false;
      }

      // Search term
      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        const matchName = item.supplierName.toLowerCase().includes(q);
        const matchGst = item.gstin.toLowerCase().includes(q);
        const matchInv = item.invNo.toLowerCase().includes(q);
        const matchReason = (item.reason || '').toLowerCase().includes(q);
        const matchStatus = (item.status || '').toLowerCase().includes(q);
        const matchBookMonth = (item.originalMonth || '').toLowerCase().includes(q);
        const match2bMonth = (item.month2b || '').toLowerCase().includes(q);
        const matchClaimMonth = (item.claimMonth || '').toLowerCase().includes(q);
        if (!matchName && !matchGst && !matchInv && !matchReason && !matchStatus && !matchBookMonth && !match2bMonth && !matchClaimMonth) {
          return false;
        }
      }

      return true;
    });
  }, [syncResult.oldItcItems, statusFilter, monthFilter, searchTerm]);

  // Unique months available
  const availableMonths = useMemo(() => {
    const set = new Set<string>();
    syncResult.oldItcItems.forEach((i) => {
      if (i.originalMonth) set.add(i.originalMonth);
    });
    return Array.from(set);
  }, [syncResult.oldItcItems]);

  // Totals of filtered items
  const summaryValues = useMemo(() => {
    let taxable = 0;
    let igst = 0;
    let cgst = 0;
    let sgst = 0;
    let totalTax = 0;
    let grandTotal = 0;

    filteredItems.forEach((r) => {
      taxable += r.bTaxable;
      igst += r.bIgst;
      cgst += r.bCgst;
      sgst += r.bSgst;
    });

    totalTax = igst + cgst + sgst;
    grandTotal = taxable + totalTax;

    return { taxable, igst, cgst, sgst, totalTax, grandTotal };
  }, [filteredItems]);

  // Percentage calculations for Pending vs Matched & Claimed (always summing to 100%)
  const { pendingPctStr, matchedPctStr } = useMemo(() => {
    const total = syncResult.oldItcItems.length;
    const pCount = syncResult.pendingCount;
    const mCount = syncResult.matchedClaimItems.length;

    if (total <= 0) {
      return { pendingPctStr: '0%', matchedPctStr: '0%' };
    }
    if (pCount <= 0 && mCount <= 0) {
      return { pendingPctStr: '0%', matchedPctStr: '0%' };
    }
    if (pCount === total) {
      return { pendingPctStr: '100%', matchedPctStr: '0%' };
    }
    if (mCount === total) {
      return { pendingPctStr: '0%', matchedPctStr: '100%' };
    }

    const rawPending = Math.round(((pCount / total) * 100) * 10) / 10;
    const rawMatched = Math.round((100 - rawPending) * 10) / 10;

    const pStr = Number.isInteger(rawPending) ? `${rawPending}%` : `${rawPending.toFixed(1)}%`;
    const mStr = Number.isInteger(rawMatched) ? `${rawMatched}%` : `${rawMatched.toFixed(1)}%`;

    return { pendingPctStr: pStr, matchedPctStr: mStr };
  }, [syncResult.oldItcItems.length, syncResult.pendingCount, syncResult.matchedClaimItems.length]);

  // Move to Not Claim handler
  const handleMoveToNotClaim = (item: OldItcItem) => {
    if (window.confirm(`Move Invoice "${item.invNo}" (${item.supplierName}) to Not Claim ITC?`)) {
      moveToNotClaim(item.id, 'Manually moved by user to Not Claim ITC');
      showToast(`Invoice "${item.invNo}" moved to Not Claim ITC.`);
      setPoolVersion((v) => v + 1);
    }
  };

  // Open Audit Modal
  const handleOpenAudit = (item: OldItcItem) => {
    setAuditItem(item);
    setAuditStatus(item.auditStatus || 'Pending');
    setAuditNote(item.auditNote || '');
  };

  // Save Audit Note
  const handleSaveAudit = () => {
    if (!auditItem) return;
    updateOldItcAudit(auditItem.id, auditStatus, auditNote);
    showToast(`Audit updated for Invoice ${auditItem.invNo}`);
    setAuditItem(null);
    setPoolVersion((v) => v + 1);
  };

  // Export to Excel
  const handleExportExcel = () => {
    try {
      const exportData = filteredItems.map((r, i) => ({
        'S.No': i + 1,
        'Book Month': r.originalMonth,
        '2B Month': r.month2b || '-',
        'Claim Month': r.status === 'Matched - Claim' ? (r.claimMonth || '-') : '-',
        'Invoice No.': r.invNo,
        'Supplier GSTIN': r.gstin,
        'Supplier Name': r.supplierName,
        'Reason': r.reason,
        'Status': r.status,
        'Books Taxable (₹)': r.bTaxable,
        'Books IGST (₹)': r.bIgst,
        'Books CGST (₹)': r.bCgst,
        'Books SGST (₹)': r.bSgst,
        'Books Total (₹)': r.bTotal,
        '2B Taxable (₹)': r.gTaxable,
        '2B IGST (₹)': r.gIgst,
        '2B CGST (₹)': r.gCgst,
        '2B SGST (₹)': r.gSgst,
        '2B Total (₹)': r.gTotal
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Old_ITC_Carried_Forward');
      XLSX.writeFile(wb, `Old_ITC_Carried_Forward_${new Date().toISOString().slice(0, 10)}.xlsx`);
      showToast('Old ITC Excel downloaded successfully!');
    } catch (e) {
      console.error(e);
      showToast('Failed to export Excel.');
    }
  };

  return (
    <section className="detail-sheet-page" style={{ padding: '4px 0 24px 0' }}>
      {/* 1. Unified Single Header & Metrics Template */}
      <div
        style={{
          background: '#ffffff',
          border: '1.5px solid #e2e8f0',
          padding: '12px 20px',
          marginBottom: '14px',
          borderRadius: 0,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: '20px',
          flexWrap: 'wrap'
        }}
      >
        {/* Left: Title & Working Month info */}
        <div style={{ minWidth: '220px', borderRight: '1.5px solid #e2e8f0', paddingRight: '20px' }}>
          <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: '#241a3d', letterSpacing: '-0.01em', whiteSpace: 'nowrap' }}>
            Old ITC – Carried Forward
          </h2>
          <div style={{ fontSize: '12px', color: '#64748b', marginTop: '3px', fontWeight: 600, whiteSpace: 'nowrap' }}>
            {filteredItems.length} invoices | Working Month: {workingMonth}
          </div>
        </div>

        {/* Right: 4 Metrics in single row */}
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flex: 1,
            gap: '16px',
            minWidth: '520px'
          }}
        >
          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              TOTAL INVOICES
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#241a3d', marginTop: '3px' }}>
              {syncResult.oldItcItems.length}
            </div>
          </div>

          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              Pending Invoices
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#6D28D9', marginTop: '3px', display: 'flex', alignItems: 'baseline', gap: '5px' }}>
              <span>{syncResult.pendingCount}</span>
              <span style={{ fontSize: '12px', fontWeight: 700, color: '#6D28D9' }}>({pendingPctStr})</span>
            </div>
          </div>

          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              Matched &amp; Claimed
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#15803d', marginTop: '3px', display: 'flex', alignItems: 'baseline', gap: '5px' }}>
              <span>{syncResult.matchedClaimItems.length}</span>
              <span style={{ fontSize: '12px', fontWeight: 700, color: '#15803d' }}>({matchedPctStr})</span>
            </div>
          </div>

          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              Total Pending ITC
            </div>
            <div style={{ fontSize: '19px', fontWeight: 800, color: '#6D28D9', marginTop: '4px' }}>
              ₹{formatInr(syncResult.pendingTotalItc)}
            </div>
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
            ₹{formatInr(summaryValues.taxable)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>IGST</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.igst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>CGST</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.cgst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>SGST</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.sgst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>Total Tax Amount</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#6D28D9', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.totalTax)}
          </div>
        </div>
      </div>

      {/* 4. Main Old ITC Table Card Container with Controls & Table */}
      <div className="card gstr1-card" style={{ width: '100%', background: '#ffffff', border: '1px solid var(--line)', padding: '20px', borderRadius: 0, boxShadow: '0 1px 3px rgba(0, 0, 0, 0.05)', boxSizing: 'border-box' }}>
        {/* Controls: Search, Filters & Excel Export */}
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
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flex: 1, minWidth: '240px', flexWrap: 'wrap' }}>
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
                maxWidth: '300px',
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
                style={{ border: 'none', background: 'transparent', cursor: 'pointer', color: '#64748b', fontSize: '13px' }}
              >
                ✕ Clear
              </button>
            )}

            {/* Month Filter */}
            <select
              value={monthFilter}
              onChange={(e) => setMonthFilter(e.target.value)}
              style={{
                border: '1px solid #cbd5e1',
                padding: '6px 10px',
                fontSize: '12px',
                fontWeight: 600,
                color: '#334155',
                borderRadius: 0,
                background: '#f8fafc'
              }}
            >
              <option value="ALL">All Months</option>
              {availableMonths.map((m) => (
                <option key={m} value={m}>{m}</option>
              ))}
            </select>

            {/* Status Filter */}
            <select
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value as any)}
              style={{
                border: '1px solid #cbd5e1',
                padding: '6px 10px',
                fontSize: '12px',
                fontWeight: 600,
                color: '#334155',
                borderRadius: 0,
                background: '#f8fafc'
              }}
            >
              <option value="ALL">All Status</option>
              <option value="Pending">Pending</option>
              <option value="Matched - Claim">Matched - Claim</option>
            </select>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <span style={{ fontSize: '12px', color: '#64748b', fontWeight: 600 }}>
              Showing {filteredItems.length} of {syncResult.oldItcItems.length} records
            </span>
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

        {/* 5. Main Old ITC Table */}
        <div className="detail-sheet-table-wrapper" style={{ overflowX: 'auto', border: '1px solid #e2e8f0', background: '#ffffff', width: '100%' }}>
          <table className="detail-sheet-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', color: '#1E1747' }}>
          <thead>
            {/* Top Multi-Header Row */}
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
                  borderRight: '1px solid rgba(255, 255, 255, 0.12)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                  background: '#2B0A4F',
                  fontFamily: "'Bricolage Grotesque', sans-serif"
                }}
              >
                Invoice &amp; Status Details
              </th>
              <th
                colSpan={4}
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
                colSpan={4}
                style={{
                  padding: '7px 10px',
                  fontSize: '13px',
                  fontWeight: 800,
                  letterSpacing: '0.05em',
                  textTransform: 'uppercase',
                  textAlign: 'center',
                  background: '#1d0535',
                  borderRight: '1px solid rgba(255, 255, 255, 0.12)',
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
                  minWidth: '180px',
                  width: '180px',
                  background: '#2B0A4F',
                  color: '#ffffff',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                  borderRight: '1px solid rgba(255, 255, 255, 0.12)',
                  fontFamily: "'Bricolage Grotesque', sans-serif"
                }}
              >
                Actions
              </th>
            </tr>

            {/* Column Header Row */}
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
              {/* 1. S.No */}
              <th
                style={{
                  width: '55px',
                  minWidth: '55px',
                  padding: '6px 8px',
                  textAlign: 'center',
                  background: '#4A1A7A',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                S.No
              </th>

              {/* 2. Book Month */}
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#4A1A7A',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Book Month
              </th>

              {/* 3. 2B Month */}
              <th
                style={{
                  width: '100px',
                  minWidth: '100px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                2B Month
              </th>

              {/* 4. Claim Month */}
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Claim Month
              </th>

              {/* 5. Invoice No. */}
              <th
                style={{
                  width: '140px',
                  minWidth: '140px',
                  padding: '6px 12px',
                  textAlign: 'center',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Invoice No.
              </th>

              {/* 6. Supplier GSTIN */}
              <th
                style={{
                  width: '155px',
                  minWidth: '155px',
                  padding: '6px 12px',
                  textAlign: 'center',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Supplier GSTIN
              </th>

              {/* 7. Supplier Name */}
              <th
                style={{
                  width: '220px',
                  minWidth: '220px',
                  padding: '6px 14px',
                  textAlign: 'center',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Supplier Name
              </th>

              {/* 8. Reason */}
              <th
                style={{
                  width: '200px',
                  minWidth: '200px',
                  padding: '6px 12px',
                  textAlign: 'center',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Reason
              </th>

              {/* 9. Status */}
              <th
                style={{
                  width: '130px',
                  minWidth: '130px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Status
              </th>

              {/* BOOKS (Purchase Register) 5 Columns */}
              <th
                style={{
                  width: '115px',
                  minWidth: '115px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#3b1263',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Taxable
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#3b1263',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                IGST
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#3b1263',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                CGST
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#3b1263',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                SGST
              </th>

              {/* GOVT GSTR-2B 4 Columns */}
              <th
                style={{
                  width: '115px',
                  minWidth: '115px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#2c0d4a',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Taxable
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#2c0d4a',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                IGST
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#2c0d4a',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                CGST
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#2c0d4a',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                SGST
              </th>
            </tr>
          </thead>
          <tbody>
            {filteredItems.length === 0 ? (
              <tr>
                <td colSpan={18} style={{ textAlign: 'center', padding: '36px', color: '#64748b', fontSize: '14px' }}>
                  No Old ITC invoices match the search or filter criteria.
                </td>
              </tr>
            ) : (
              filteredItems.map((item, idx) => {
                const isClaimed = item.status === 'Matched - Claim';

                // Comparison indicators for Taxable & Tax amounts
                const renderAmt = (bAmt: number, gAmt: number, isBooks: boolean) => {
                  const val = isBooks ? bAmt : gAmt;
                  const other = isBooks ? gAmt : bAmt;
                  let arrow = null;
                  if (other > 0 && Math.abs(val - other) > 0.01) {
                    if (val > other) {
                      arrow = <span style={{ color: '#15803d', fontSize: '10px', marginLeft: '3px' }}>▲</span>;
                    } else {
                      arrow = <span style={{ color: '#b91c1c', fontSize: '10px', marginLeft: '3px' }}>▼</span>;
                    }
                  }
                  return (
                    <span>
                      {val === 0 ? <span style={{ color: '#94a3b8' }}>0.00</span> : formatInr(val)}
                      {arrow}
                    </span>
                  );
                };

                return (
                  <tr
                    key={item.id}
                    style={{
                      background: idx % 2 === 0 ? '#FAF8FD' : '#ffffff',
                      borderBottom: '1px solid #e2e8f0',
                      height: '42px'
                    }}
                    onMouseEnter={(e) => (e.currentTarget.style.background = '#F3EBFE')}
                    onMouseLeave={(e) => (e.currentTarget.style.background = idx % 2 === 0 ? '#FAF8FD' : '#ffffff')}
                  >
                    {/* 1. S.No */}
                    <td
                      style={{
                        width: '55px',
                        minWidth: '55px',
                        padding: '8px 8px',
                        textAlign: 'center',
                        color: '#64748b',
                        fontWeight: 600
                      }}
                    >
                      {idx + 1}
                    </td>

                    {/* 2. Book Month */}
                    <td
                      style={{
                        width: '105px',
                        minWidth: '105px',
                        padding: '8px 10px',
                        textAlign: 'center',
                        fontWeight: 700,
                        color: '#334155'
                      }}
                    >
                      {item.originalMonth}
                    </td>

                    {/* 3. 2B Month */}
                    <td
                      style={{
                        width: '100px',
                        minWidth: '100px',
                        padding: '8px 10px',
                        textAlign: 'center',
                        fontWeight: 600,
                        color: item.month2b ? '#334155' : '#94a3b8'
                      }}
                    >
                      {item.month2b || '-'}
                    </td>

                    {/* 4. Claim Month */}
                    <td
                      style={{
                        width: '105px',
                        minWidth: '105px',
                        padding: '8px 10px',
                        textAlign: 'center',
                        fontWeight: 700,
                        color: isClaimed ? '#15803d' : '#94a3b8'
                      }}
                    >
                      {isClaimed ? (item.claimMonth || '-') : '-'}
                    </td>

                    {/* 5. Invoice No. */}
                    <td
                      style={{
                        width: '140px',
                        minWidth: '140px',
                        padding: '8px 12px',
                        textAlign: 'left',
                        fontWeight: 700,
                        color: '#1E1747',
                        fontSize: '13px'
                      }}
                    >
                      {item.invNo}
                    </td>

                    {/* 6. Supplier GSTIN */}
                    <td
                      style={{
                        width: '155px',
                        minWidth: '155px',
                        padding: '8px 12px',
                        textAlign: 'left',
                        fontFamily: 'monospace',
                        fontSize: '12px'
                      }}
                    >
                      {item.gstin}
                    </td>

                    {/* 7. Supplier Name */}
                    <td
                      style={{
                        width: '220px',
                        minWidth: '220px',
                        padding: '8px 14px',
                        textAlign: 'left',
                        fontWeight: 600,
                        color: '#1E1747'
                      }}
                    >
                      {item.supplierName}
                    </td>

                    {/* 8. Reason */}
                    <td
                      style={{
                        width: '200px',
                        minWidth: '200px',
                        padding: '8px 12px',
                        textAlign: 'left',
                        fontSize: '12px',
                        color: '#475569'
                      }}
                      title={item.reason}
                    >
                      {item.reason}
                    </td>

                    {/* 9. Status */}
                    <td
                      style={{
                        width: '130px',
                        minWidth: '130px',
                        padding: '8px 10px',
                        textAlign: 'center'
                      }}
                    >
                      <span
                        style={{
                          display: 'inline-block',
                          padding: '2px 8px',
                          fontSize: '11px',
                          fontWeight: 700,
                          borderRadius: '2px',
                          background: isClaimed ? '#dcfce7' : '#fef3c7',
                          color: isClaimed ? '#15803d' : '#b45309',
                          border: isClaimed ? '1px solid #86efac' : '1px solid #fde68a'
                        }}
                      >
                        {item.status}
                      </span>
                    </td>

                    {/* BOOKS 4 Columns (Same Purple tint #F6F2FC) */}
                    <td style={{ width: '115px', minWidth: '115px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600 }}>
                      {renderAmt(item.bTaxable, item.gTaxable, true)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600 }}>
                      {renderAmt(item.bIgst, item.gIgst, true)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600 }}>
                      {renderAmt(item.bCgst, item.gCgst, true)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600 }}>
                      {renderAmt(item.bSgst, item.gSgst, true)}
                    </td>

                    {/* GOVT 2B 4 Columns (Lavender tint #F6F2FC) */}
                    <td style={{ width: '115px', minWidth: '115px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600 }}>
                      {renderAmt(item.bTaxable, item.gTaxable, false)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600 }}>
                      {renderAmt(item.bIgst, item.gIgst, false)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600 }}>
                      {renderAmt(item.bCgst, item.gCgst, false)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600 }}>
                      {renderAmt(item.bSgst, item.gSgst, false)}
                    </td>

                    {/* Action Column: View, Audit, Move to Not Claim */}
                    <td
                      style={{
                        width: '180px',
                        minWidth: '180px',
                        padding: '8px 12px',
                        textAlign: 'center',
                        whiteSpace: 'nowrap'
                      }}
                    >
                      <div style={{ display: 'inline-flex', gap: '4px' }}>
                        <button
                          type="button"
                          className="btn-action-view"
                          onClick={() => setViewItem(item)}
                          title="View Side-by-Side Comparison"
                          style={{
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: '4px',
                            background: '#f1f5f9',
                            color: '#334155',
                            border: '1px solid #cbd5e1',
                            padding: '3px 8px',
                            fontSize: '11px',
                            fontWeight: 700,
                            cursor: 'pointer',
                            borderRadius: 0
                          }}
                        >
                          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                            <circle cx="12" cy="12" r="3"></circle>
                          </svg>
                          <span>View</span>
                        </button>
                        <button
                          type="button"
                          className="btn-action-audit"
                          onClick={() => handleOpenAudit(item)}
                          title="Add Audit Note & Mark Verified/Query"
                          style={{
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: '4px',
                            background: item.auditStatus === 'Verified' ? '#dcfce7' : '#FAF5FF',
                            color: item.auditStatus === 'Verified' ? '#15803d' : '#6B21A8',
                            border: item.auditStatus === 'Verified' ? '1px solid #86efac' : '1px solid #D8B4FE',
                            padding: '3px 8px',
                            fontSize: '11px',
                            fontWeight: 700,
                            cursor: 'pointer',
                            borderRadius: 0
                          }}
                        >
                          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                            <path d="M12 20h9"></path>
                            <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
                          </svg>
                          <span>Audit</span>
                        </button>
                        {!isClaimed && (
                          <button
                            type="button"
                            onClick={() => handleMoveToNotClaim(item)}
                            style={{
                              background: '#fef2f2',
                              color: '#b91c1c',
                              border: '1px solid #fecaca',
                              padding: '3px 7px',
                              fontSize: '11px',
                              fontWeight: 700,
                              cursor: 'pointer',
                              borderRadius: 0
                            }}
                            title="Move to Not Claim ITC"
                          >
                            Not Claim
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>

      {/* ======================================================== */}
      {/* 4. Side-by-Side View Modal (GST RecoManager Style)        */}
      {/* ======================================================== */}
      {viewItem && (() => {
        const isMatched = viewItem.status === 'Matched - Claim';
        const bTotalGst = viewItem.bIgst + viewItem.bCgst + viewItem.bSgst;
        const gTotalGst = viewItem.gIgst + viewItem.gCgst + viewItem.gSgst;
        const dTotalGst = bTotalGst - gTotalGst;

        const has2b = Boolean(viewItem.month2b || viewItem.gTaxable || viewItem.gIgst || viewItem.gCgst || viewItem.gSgst || viewItem.gInvNo);
        const isTaxableMatch = has2b && Math.abs(viewItem.bTaxable - viewItem.gTaxable) <= 1;
        const isIgstMatch = has2b && Math.abs(viewItem.bIgst - viewItem.gIgst) <= 1;
        const isCgstMatch = has2b && Math.abs(viewItem.bCgst - viewItem.gCgst) <= 1;
        const isSgstMatch = has2b && Math.abs(viewItem.bSgst - viewItem.gSgst) <= 1;
        const isInvMatch = has2b ? (viewItem.gInvNo ? cleanInv(viewItem.invNo) === cleanInv(viewItem.gInvNo) : true) : false;
        const isGstinMatch = true;
        const isAllFieldsMatch = isGstinMatch && isInvMatch && isTaxableMatch && isIgstMatch && isCgstMatch && isSgstMatch;

        return (
          <div
            className="reco-modal-backdrop"
            style={{
              position: 'fixed',
              inset: 0,
              background: 'rgba(15, 23, 42, 0.55)',
              backdropFilter: 'blur(3px)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              zIndex: 1050,
              padding: '16px'
            }}
            onClick={() => setViewItem(null)}
          >
            <div
              className="reco-modal-card"
              style={{
                maxWidth: '920px',
                width: '96%',
                background: '#ffffff',
                borderRadius: '16px',
                boxShadow: '0 25px 50px -12px rgba(42, 18, 80, 0.25), 0 0 0 1px rgba(0, 0, 0, 0.05)',
                fontFamily: "'Bricolage Grotesque', system-ui, -apple-system, 'Segoe UI', sans-serif",
                overflow: 'hidden',
                display: 'flex',
                flexDirection: 'column'
              }}
              onClick={(e) => e.stopPropagation()}
            >
              {/* Top Title Header */}
              <div
                style={{
                  padding: '20px 24px 16px 24px',
                  background: '#ffffff',
                  borderBottom: '1px solid #F0ECF8',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  gap: '16px'
                }}
              >
                <div>
                  <h2 style={{ margin: 0, fontSize: '21px', fontWeight: 700, color: '#350B5E', letterSpacing: '-0.01em' }}>
                    GST RecoManager
                  </h2>
                  <div style={{ margin: '6px 0 0 0', display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '8px 12px' }}>
                    <div style={{ fontSize: '13px', color: (isMatched || isAllFieldsMatch) ? '#15803D' : '#DC2626', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
                      <span style={{ fontSize: '14px' }}>{(isMatched || isAllFieldsMatch) ? '✓' : '⚠'}</span>
                      <span>
                        {(isMatched || isAllFieldsMatch)
                          ? `6/6 Fields Match – Verified & Eligible for ITC Claim (${viewItem.claimMonth || viewItem.month2b || workingMonth})`
                          : `Pending Old ITC (${viewItem.originalMonth}) – ${viewItem.reason || 'ITC Verification Required'}`}
                      </span>
                    </div>

                    {/* Field Match / Mismatch Tags */}
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                      {[
                        { label: 'GST No', match: isGstinMatch },
                        { label: 'Invoice', match: isInvMatch },
                        { label: 'Taxable', match: isTaxableMatch },
                        { label: 'IGST', match: isIgstMatch },
                        { label: 'CGST', match: isCgstMatch },
                        { label: 'SGST', match: isSgstMatch }
                      ].map((f) => (
                        <span
                          key={f.label}
                          style={{
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: '4px',
                            padding: '2px 8px',
                            borderRadius: '4px',
                            fontSize: '11px',
                            fontWeight: 600,
                            whiteSpace: 'nowrap',
                            background: f.match ? '#DCFCE7' : '#FEE2E2',
                            color: f.match ? '#15803D' : '#B42318',
                            border: f.match ? '1px solid #86EFAC' : '1px solid #FCA5A5'
                          }}
                        >
                          <span>{f.label}</span>
                          <strong style={{ fontWeight: 800 }}>{f.match ? '✓' : '✕'}</strong>
                        </span>
                      ))}
                    </div>
                  </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexShrink: 0 }}>
                  <span
                    style={{
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '5px',
                      padding: '5px 12px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: 600,
                      whiteSpace: 'nowrap',
                      flexShrink: 0,
                      background: isMatched ? '#DCFCE7' : '#FEF3C7',
                      color: isMatched ? '#14532D' : '#B45309',
                      border: isMatched ? '1px solid #86EFAC' : '1px solid #FDE68A'
                    }}
                  >
                    <span>●</span>
                    <span>{viewItem.status}</span>
                  </span>
                  <button
                    type="button"
                    onClick={() => setViewItem(null)}
                    style={{
                      width: '36px',
                      height: '36px',
                      minWidth: '36px',
                      borderRadius: '8px',
                      background: '#F4EFFC',
                      border: 'none',
                      color: '#350B5E',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      cursor: 'pointer',
                      fontSize: '16px',
                      fontWeight: 700,
                      transition: 'all 0.15s ease'
                    }}
                    onMouseEnter={(e) => (e.currentTarget.style.background = '#E8DDF7')}
                    onMouseLeave={(e) => (e.currentTarget.style.background = '#F4EFFC')}
                    title="Close"
                  >
                    ✕
                  </button>
                </div>
              </div>

              {/* Modal Body */}
              <div style={{ padding: '20px 24px', maxHeight: '76vh', overflowY: 'auto' }}>
                {/* 1. Two Party Cards (Side by Side) */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '16px', marginBottom: '18px' }}>
                  {/* Books Party */}
                  <div style={{ background: '#FAF5FE', border: '1.5px solid #DDD6FE', borderRadius: '12px', padding: '16px 20px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                      <div style={{ fontSize: '11px', fontWeight: 700, color: '#5B21B6', letterSpacing: '0.04em', textTransform: 'uppercase' }}>
                        BOOKS PARTY (AS PER PURCHASE REGISTER)
                      </div>
                    </div>
                    <div style={{ fontSize: '17px', fontWeight: 700, color: '#1E1B4B', lineHeight: 1.25 }}>
                      {viewItem.supplierName || 'Unknown Supplier'}
                    </div>
                    <div style={{ fontSize: '13px', color: '#4B5563', marginTop: '10px', fontWeight: 500, display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <span>GSTIN:</span>
                      <strong style={{ color: '#0F172A', fontWeight: 600, letterSpacing: '0.02em' }}>{viewItem.gstin || '-'}</strong>
                      <span
                        style={{
                          display: 'inline-flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          width: '18px',
                          height: '18px',
                          borderRadius: '50%',
                          background: '#DCFCE7',
                          color: '#15803D',
                          fontSize: '11px',
                          fontWeight: 800,
                          lineHeight: 1
                        }}
                        title="GSTIN Matched"
                      >
                        ✓
                      </span>
                    </div>
                  </div>

                  {/* GST Portal Party */}
                  <div style={{ background: '#F0FDFA', border: '1.5px solid #A7F3D0', borderRadius: '12px', padding: '16px 20px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                      <div style={{ fontSize: '11px', fontWeight: 700, color: '#0E7490', letterSpacing: '0.04em', textTransform: 'uppercase' }}>
                        GST PORTAL PARTY (AS PER GSTR-2B)
                      </div>
                    </div>
                    <div style={{ fontSize: '17px', fontWeight: 700, color: '#0F172A', lineHeight: 1.25 }}>
                      {(viewItem.supplierName || 'Unknown Supplier').toUpperCase()}
                    </div>
                    <div style={{ fontSize: '13px', color: '#4B5563', marginTop: '10px', fontWeight: 500, display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <span>GSTIN:</span>
                      <strong style={{ color: '#0F172A', fontWeight: 600, letterSpacing: '0.02em' }}>{viewItem.gstin || '-'}</strong>
                      <span
                        style={{
                          display: 'inline-flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          width: '18px',
                          height: '18px',
                          borderRadius: '50%',
                          background: '#DCFCE7',
                          color: '#15803D',
                          fontSize: '11px',
                          fontWeight: 800,
                          lineHeight: 1
                        }}
                        title="Govt Portal GSTIN"
                      >
                        ✓
                      </span>
                    </div>
                  </div>
                </div>

                {/* 2. Card 1: BOOKS (PURCHASE REGISTER) */}
                <div style={{ border: '1.5px solid #DDD6FE', borderRadius: '12px', overflow: 'hidden', background: '#ffffff', marginBottom: '16px', boxShadow: '0 2px 8px rgba(53, 11, 94, 0.04)' }}>
                  <div style={{ padding: '10px 16px', background: '#F3E8FF', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #E9D5FF' }}>
                    <div style={{ fontSize: '12px', fontWeight: 700, color: '#4C1D95', letterSpacing: '0.03em' }}>
                      BOOKS (PURCHASE REGISTER)
                    </div>
                    <div style={{ fontSize: '11.5px', color: '#581C87', fontWeight: 600 }}>
                      {viewItem.supplierName || '-'} • {viewItem.gstin || '-'}
                    </div>
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                    <thead>
                      <tr style={{ background: '#350B5E' }}>
                        <th style={{ padding: '10px 12px', textAlign: 'left', width: '22%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>SOURCE / INV NO.</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '15%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>TAXABLE AMOUNT (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>IGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>CGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>SGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '16%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>TOTAL GST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '17%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>TOTAL VALUE (₹)</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr style={{ height: '52px', background: '#ffffff' }}>
                        <td style={{ padding: '10px 12px', textAlign: 'left' }}>
                          <div style={{ fontSize: '13.5px', fontWeight: 600, color: '#0F172A', display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <span>{viewItem.invNo || '-'}</span>
                            <span style={{ fontSize: '12px', fontWeight: 800, color: isInvMatch ? '#15803D' : '#DC2626' }}>
                              {isInvMatch ? '✓' : '✕'}
                            </span>
                          </div>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, fontVariantNumeric: 'tabular-nums', color: '#0F172A' }}>
                          <span>{formatInr(viewItem.bTaxable)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: isTaxableMatch ? '#15803D' : '#DC2626' }}>
                            {isTaxableMatch ? '✓' : '✕'}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, fontVariantNumeric: 'tabular-nums', color: '#0F172A' }}>
                          <span>{formatInr(viewItem.bIgst)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: isIgstMatch ? '#15803D' : '#DC2626' }}>
                            {isIgstMatch ? '✓' : '✕'}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, fontVariantNumeric: 'tabular-nums', color: '#0F172A' }}>
                          <span>{formatInr(viewItem.bCgst)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: isCgstMatch ? '#15803D' : '#DC2626' }}>
                            {isCgstMatch ? '✓' : '✕'}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, fontVariantNumeric: 'tabular-nums', color: '#0F172A' }}>
                          <span>{formatInr(viewItem.bSgst)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: isSgstMatch ? '#15803D' : '#DC2626' }}>
                            {isSgstMatch ? '✓' : '✕'}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: '#350B5E', fontSize: '13.5px', fontVariantNumeric: 'tabular-nums' }}>
                          <span>{formatInr(bTotalGst)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: Math.abs(dTotalGst) <= 1 ? '#15803D' : '#DC2626' }}>
                            {Math.abs(dTotalGst) <= 1 ? '✓' : '✕'}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: '#350B5E', fontSize: '13.5px', fontVariantNumeric: 'tabular-nums' }}>
                          <span>{formatInr(viewItem.bTotal)}</span>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 3. Card 2: GOVT GSTR-2B */}
                <div style={{ border: '1.5px solid #A5F3FC', borderRadius: '12px', overflow: 'hidden', background: '#ffffff', marginBottom: '18px', boxShadow: '0 2px 8px rgba(14, 116, 144, 0.04)' }}>
                  <div style={{ padding: '10px 16px', background: '#CFFAFE', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #A5F3FC' }}>
                    <div style={{ fontSize: '12px', fontWeight: 700, color: '#0E7490', letterSpacing: '0.03em' }}>
                      GOVT GSTR-2B
                    </div>
                    <div style={{ fontSize: '11.5px', color: '#0369A1', fontWeight: 600 }}>
                      {(viewItem.supplierName || '-').toUpperCase()} • {viewItem.gstin || '-'}
                    </div>
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                    <thead>
                      <tr style={{ background: '#155E75' }}>
                        <th style={{ padding: '10px 12px', textAlign: 'left', width: '22%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>SOURCE / INV NO.</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '15%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>TAXABLE AMOUNT (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>IGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>CGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>SGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '16%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>TOTAL GST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '17%', fontSize: '11px', fontWeight: 600, letterSpacing: '0.04em', color: '#ffffff' }}>TOTAL VALUE (₹)</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr style={{ height: '52px', background: '#ffffff' }}>
                        <td style={{ padding: '10px 12px', textAlign: 'left' }}>
                          <div style={{ fontSize: '13.5px', fontWeight: 600, color: '#0F172A', display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <span>{viewItem.gInvNo || viewItem.invNo || '-'}</span>
                            <span style={{ fontSize: '12px', fontWeight: 800, color: '#15803D' }}>
                              ✓
                            </span>
                          </div>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, fontVariantNumeric: 'tabular-nums', color: '#0F172A' }}>
                          <span>{formatInr(viewItem.gTaxable)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: '#15803D' }}>
                            ✓
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, fontVariantNumeric: 'tabular-nums', color: '#0F172A' }}>
                          <span>{formatInr(viewItem.gIgst)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: '#15803D' }}>
                            ✓
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, fontVariantNumeric: 'tabular-nums', color: '#0F172A' }}>
                          <span>{formatInr(viewItem.gCgst)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: '#15803D' }}>
                            ✓
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, fontVariantNumeric: 'tabular-nums', color: '#0F172A' }}>
                          <span>{formatInr(viewItem.gSgst)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: '#15803D' }}>
                            ✓
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: '#0E7490', fontSize: '13.5px', fontVariantNumeric: 'tabular-nums' }}>
                          <span>{formatInr(gTotalGst)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: '#15803D' }}>
                            ✓
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: '#0E7490', fontSize: '13.5px', fontVariantNumeric: 'tabular-nums' }}>
                          <span>{formatInr(viewItem.gTotal)}</span>
                          <span style={{ marginLeft: '5px', fontSize: '12px', fontWeight: 800, color: '#15803D' }}>
                            ✓
                          </span>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 4. Action Guidance / Observation */}
                <div
                  style={{
                    marginTop: '12px',
                    fontSize: '13px',
                    fontWeight: 500,
                    color: (isMatched || isAllFieldsMatch) ? '#15803D' : '#DC2626',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px'
                  }}
                >
                  <span style={{ fontSize: '14px', fontWeight: 700 }}>
                    {(isMatched || isAllFieldsMatch) ? '✓' : '⚠'}
                  </span>
                  <span>
                    {(isMatched || isAllFieldsMatch)
                      ? `Books & GSTR-2B matched. Verified & claimed in ${viewItem.claimMonth || viewItem.month2b || workingMonth}.`
                      : viewItem.reason || 'Books GST and 2B GST not match. Please verify.'}
                  </span>
                </div>

                {/* Audit Observation Display if exists */}
                {viewItem.auditNote && (
                  <div style={{ marginTop: '16px', padding: '12px 16px', background: '#F8F6FC', borderRadius: '10px', borderLeft: '4px solid #3B1A6B' }}>
                    <div style={{ fontSize: '11.5px', fontWeight: 600, color: '#3B1A6B', textTransform: 'uppercase', marginBottom: '4px' }}>
                      Saved Audit Observation ({viewItem.auditStatus || 'Pending'}):
                    </div>
                    <div style={{ fontSize: '13px', color: '#1E293B', fontWeight: 500 }}>{viewItem.auditNote}</div>
                    {viewItem.auditUpdatedAt && (
                      <div style={{ fontSize: '11px', color: '#64748B', marginTop: '4px', fontWeight: 500 }}>
                        Updated: {viewItem.auditUpdatedAt}
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* Footer */}
              <div
                style={{
                  padding: '14px 24px',
                  background: '#FBFAFD',
                  borderTop: '1px solid #EBE6F3',
                  display: 'flex',
                  justifyContent: 'flex-end',
                  alignItems: 'center'
                }}
              >
                <button
                  type="button"
                  onClick={() => setViewItem(null)}
                  style={{
                    height: '42px',
                    padding: '0 24px',
                    borderRadius: '8px',
                    background: '#350B5E',
                    color: '#ffffff',
                    border: 'none',
                    fontSize: '13px',
                    fontWeight: 700,
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    boxShadow: '0 2px 8px rgba(53, 11, 94, 0.25)',
                    transition: 'all 0.15s ease'
                  }}
                  onMouseEnter={(e) => (e.currentTarget.style.background = '#250742')}
                  onMouseLeave={(e) => (e.currentTarget.style.background = '#350B5E')}
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ======================================================== */}
      {/* 5. Audit & Action Modal (GST RecoManager Style)          */}
      {/* ======================================================== */}
      {auditItem && (() => {
        const bTotalGst = auditItem.bIgst + auditItem.bCgst + auditItem.bSgst;
        const gTotalGst = auditItem.gIgst + auditItem.gCgst + auditItem.gSgst;

        return (
          <div
            className="reco-modal-backdrop"
            style={{
              position: 'fixed',
              inset: 0,
              background: 'rgba(15, 23, 42, 0.55)',
              backdropFilter: 'blur(3px)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              zIndex: 1050,
              padding: '16px'
            }}
            onClick={() => setAuditItem(null)}
          >
            <div
              className="reco-modal-card"
              style={{
                maxWidth: '920px',
                width: '96%',
                background: '#ffffff',
                borderRadius: '16px',
                boxShadow: '0 25px 50px -12px rgba(42, 18, 80, 0.25), 0 0 0 1px rgba(0, 0, 0, 0.05)',
                fontFamily: "'Bricolage Grotesque', system-ui, -apple-system, 'Segoe UI', sans-serif",
                overflow: 'hidden',
                display: 'flex',
                flexDirection: 'column'
              }}
              onClick={(e) => e.stopPropagation()}
            >
              {/* Header */}
              <div
                style={{
                  padding: '20px 24px 16px 24px',
                  background: '#ffffff',
                  borderBottom: '1px solid #F0ECF8',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  gap: '16px'
                }}
              >
                <div>
                  <h2 style={{ margin: 0, fontSize: '21px', fontWeight: 700, color: '#350B5E', letterSpacing: '-0.01em' }}>
                    Audit &amp; Verify Invoice – {auditItem.invNo}
                  </h2>
                  <div style={{ margin: '6px 0 0 0', fontSize: '13px', color: '#64748B', fontWeight: 500 }}>
                    Supplier: <strong style={{ color: '#0F172A' }}>{auditItem.supplierName}</strong> ({auditItem.gstin}) • Book Month: <strong>{auditItem.originalMonth}</strong>
                  </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexShrink: 0 }}>
                  <span
                    style={{
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '5px',
                      padding: '5px 12px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: 600,
                      whiteSpace: 'nowrap',
                      flexShrink: 0,
                      background: auditStatus === 'Verified' ? '#DCFCE7' : auditStatus === 'Query' ? '#FEE2E2' : '#FAF5FF',
                      color: auditStatus === 'Verified' ? '#15803D' : auditStatus === 'Query' ? '#B42318' : '#6B21A8',
                      border: auditStatus === 'Verified' ? '1px solid #86EFAC' : auditStatus === 'Query' ? '1px solid #FCA5A5' : '1px solid #D8B4FE'
                    }}
                  >
                    <span>●</span>
                    <span>Status: {auditStatus}</span>
                  </span>
                  <button
                    type="button"
                    onClick={() => setAuditItem(null)}
                    style={{
                      width: '36px',
                      height: '36px',
                      minWidth: '36px',
                      borderRadius: '8px',
                      background: '#F4EFFC',
                      border: 'none',
                      color: '#350B5E',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      cursor: 'pointer',
                      fontSize: '16px',
                      fontWeight: 700,
                      transition: 'all 0.15s ease'
                    }}
                    onMouseEnter={(e) => (e.currentTarget.style.background = '#E8DDF7')}
                    onMouseLeave={(e) => (e.currentTarget.style.background = '#F4EFFC')}
                    title="Close"
                  >
                    ✕
                  </button>
                </div>
              </div>

              {/* Modal Body */}
              <div style={{ padding: '20px 24px', maxHeight: '76vh', overflowY: 'auto' }}>
                {/* 1. Two Party Cards */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '16px', marginBottom: '18px' }}>
                  {/* Books Party */}
                  <div style={{ background: '#FAF5FE', border: '1.5px solid #DDD6FE', borderRadius: '12px', padding: '14px 18px' }}>
                    <div style={{ fontSize: '11px', fontWeight: 700, color: '#5B21B6', letterSpacing: '0.04em', textTransform: 'uppercase', marginBottom: '6px' }}>
                      BOOKS PARTY (AS PER PURCHASE REGISTER)
                    </div>
                    <div style={{ fontSize: '16px', fontWeight: 700, color: '#1E1B4B' }}>
                      {auditItem.supplierName}
                    </div>
                    <div style={{ fontSize: '13px', color: '#4B5563', marginTop: '6px', fontWeight: 500 }}>
                      GSTIN: <strong style={{ color: '#0F172A' }}>{auditItem.gstin}</strong>
                    </div>
                  </div>

                  {/* GST Portal Party */}
                  <div style={{ background: '#F0FDFA', border: '1.5px solid #A7F3D0', borderRadius: '12px', padding: '14px 18px' }}>
                    <div style={{ fontSize: '11px', fontWeight: 700, color: '#0E7490', letterSpacing: '0.04em', textTransform: 'uppercase', marginBottom: '6px' }}>
                      GST PORTAL PARTY (AS PER GSTR-2B)
                    </div>
                    <div style={{ fontSize: '16px', fontWeight: 700, color: '#0F172A' }}>
                      {(auditItem.supplierName).toUpperCase()}
                    </div>
                    <div style={{ fontSize: '13px', color: '#4B5563', marginTop: '6px', fontWeight: 500 }}>
                      GSTIN: <strong style={{ color: '#0F172A' }}>{auditItem.gstin}</strong>
                    </div>
                  </div>
                </div>

                {/* 2. Side-by-Side Comparison Tables */}
                {/* Books Card */}
                <div style={{ border: '1.5px solid #DDD6FE', borderRadius: '12px', overflow: 'hidden', background: '#ffffff', marginBottom: '16px', boxShadow: '0 2px 8px rgba(53, 11, 94, 0.04)' }}>
                  <div style={{ padding: '10px 16px', background: '#F3E8FF', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #E9D5FF' }}>
                    <div style={{ fontSize: '12px', fontWeight: 700, color: '#4C1D95', letterSpacing: '0.03em' }}>
                      BOOKS (PURCHASE REGISTER)
                    </div>
                    <div style={{ fontSize: '11.5px', color: '#581C87', fontWeight: 600 }}>
                      {auditItem.invNo} • Date: {auditItem.date}
                    </div>
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                    <thead>
                      <tr style={{ background: '#350B5E' }}>
                        <th style={{ padding: '10px 12px', textAlign: 'left', width: '22%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>SOURCE / INV NO.</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '15%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>TAXABLE AMOUNT (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>IGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>CGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>SGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '16%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>TOTAL GST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '17%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>TOTAL VALUE (₹)</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr style={{ height: '48px', background: '#ffffff' }}>
                        <td style={{ padding: '10px 12px', textAlign: 'left', fontWeight: 600, color: '#0F172A' }}>{auditItem.invNo}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, color: '#0F172A' }}>{formatInr(auditItem.bTaxable)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, color: '#0F172A' }}>{formatInr(auditItem.bIgst)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, color: '#0F172A' }}>{formatInr(auditItem.bCgst)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, color: '#0F172A' }}>{formatInr(auditItem.bSgst)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: '#350B5E' }}>{formatInr(bTotalGst)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: '#350B5E' }}>{formatInr(auditItem.bTotal)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 2B Card */}
                <div style={{ border: '1.5px solid #A5F3FC', borderRadius: '12px', overflow: 'hidden', background: '#ffffff', marginBottom: '18px', boxShadow: '0 2px 8px rgba(14, 116, 144, 0.04)' }}>
                  <div style={{ padding: '10px 16px', background: '#CFFAFE', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #A5F3FC' }}>
                    <div style={{ fontSize: '12px', fontWeight: 700, color: '#0E7490', letterSpacing: '0.03em' }}>
                      GOVT GSTR-2B
                    </div>
                    <div style={{ fontSize: '11.5px', color: '#0369A1', fontWeight: 600 }}>
                      {auditItem.month2b || 'Reflected in 2B'}
                    </div>
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                    <thead>
                      <tr style={{ background: '#155E75' }}>
                        <th style={{ padding: '10px 12px', textAlign: 'left', width: '22%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>SOURCE / INV NO.</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '15%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>TAXABLE AMOUNT (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>IGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>CGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '10%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>SGST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '16%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>TOTAL GST (₹)</th>
                        <th style={{ padding: '10px 12px', textAlign: 'right', width: '17%', fontSize: '11px', fontWeight: 600, color: '#ffffff' }}>TOTAL VALUE (₹)</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr style={{ height: '48px', background: '#ffffff' }}>
                        <td style={{ padding: '10px 12px', textAlign: 'left', fontWeight: 600, color: '#0F172A' }}>{auditItem.gInvNo || auditItem.invNo}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, color: '#0F172A' }}>{formatInr(auditItem.gTaxable)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, color: '#0F172A' }}>{formatInr(auditItem.gIgst)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, color: '#0F172A' }}>{formatInr(auditItem.gCgst)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 500, color: '#0F172A' }}>{formatInr(auditItem.gSgst)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: '#0E7490' }}>{formatInr(gTotalGst)}</td>
                        <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: '#0E7490' }}>{formatInr(auditItem.gTotal)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 3. Audit Controls */}
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '16px', marginTop: '16px' }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '12.5px', fontWeight: 700, color: '#350B5E', marginBottom: '6px' }}>
                      Audit Status
                    </label>
                    <select
                      value={auditStatus}
                      onChange={(e) => setAuditStatus(e.target.value)}
                      style={{
                        width: '100%',
                        height: '42px',
                        border: '1.5px solid #DCCFF3',
                        borderRadius: '8px',
                        padding: '0 12px',
                        fontSize: '13px',
                        fontWeight: 600,
                        background: '#ffffff',
                        color: '#1E293B',
                        outline: 'none',
                        cursor: 'pointer'
                      }}
                    >
                      <option value="Verified">Verified / Approved</option>
                      <option value="Query">Query / Follow-up Sent</option>
                      <option value="Pending">Pending Review</option>
                    </select>
                  </div>

                  <div>
                    <label style={{ display: 'block', fontSize: '12.5px', fontWeight: 700, color: '#350B5E', marginBottom: '6px' }}>
                      Audit Observation / Remark (Optional)
                    </label>
                    <textarea
                      value={auditNote}
                      onChange={(e) => setAuditNote(e.target.value)}
                      rows={3}
                      placeholder="Enter audit explanation or observation..."
                      style={{
                        width: '100%',
                        padding: '10px 12px',
                        border: '1.5px solid #DCCFF3',
                        borderRadius: '8px',
                        background: '#ffffff',
                        color: '#1E293B',
                        fontSize: '13px',
                        fontFamily: 'inherit',
                        outline: 'none',
                        resize: 'none',
                        boxSizing: 'border-box'
                      }}
                    />
                  </div>
                </div>
              </div>

              {/* Footer */}
              <div
                style={{
                  padding: '14px 24px',
                  background: '#FBFAFD',
                  borderTop: '1px solid #EBE6F3',
                  display: 'flex',
                  justifyContent: 'flex-end',
                  alignItems: 'center',
                  gap: '12px'
                }}
              >
                <button
                  type="button"
                  onClick={() => setAuditItem(null)}
                  style={{
                    height: '42px',
                    padding: '0 20px',
                    borderRadius: '8px',
                    background: '#ffffff',
                    color: '#475569',
                    border: '1.5px solid #CBD5E1',
                    fontSize: '13px',
                    fontWeight: 700,
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    transition: 'all 0.15s ease'
                  }}
                  onMouseEnter={(e) => (e.currentTarget.style.background = '#F8FAFC')}
                  onMouseLeave={(e) => (e.currentTarget.style.background = '#ffffff')}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleSaveAudit}
                  style={{
                    height: '42px',
                    padding: '0 22px',
                    borderRadius: '8px',
                    background: '#15803D',
                    color: '#ffffff',
                    border: 'none',
                    fontSize: '13px',
                    fontWeight: 700,
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    boxShadow: '0 2px 8px rgba(21, 128, 61, 0.25)',
                    transition: 'all 0.15s ease'
                  }}
                  onMouseEnter={(e) => (e.currentTarget.style.background = '#166534')}
                  onMouseLeave={(e) => (e.currentTarget.style.background = '#15803D')}
                >
                  Save Audit
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </section>
  );
}
