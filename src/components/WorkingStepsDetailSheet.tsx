import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { useLiveReconciliation } from '../services/recoSyncService';

export type WorkingDetailCategory = 'matched' | 'not-in-book' | 'not-in-2b';

interface WorkingStepsDetailSheetProps {
  category: WorkingDetailCategory;
  onBack: () => void;
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
      selectedRows = recoData.notInBook;
      totalContext = recoData.summary.notInBookCount + recoData.summary.matchedCount;
      count = recoData.summary.notInBookCount;
    } else {
      selectedRows = recoData.notIn2b;
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
  }, [recoData, category]);

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

  const handleExportExcel = () => {
    try {
      const exportData = filteredRows.map((r, i) => ({
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
      }));

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
      {/* 1. Header Bar with Title and Back to Steps Button */}
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          background: '#ffffff',
          border: '1.5px solid #e2e8f0',
          padding: '14px 20px',
          marginBottom: '16px',
          borderRadius: 0
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
          <div
            style={{
              width: '42px',
              height: '42px',
              background: config.statusPillBg,
              color: config.badgeColor,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              border: `1.5px solid ${config.statusPillBorder}`,
              borderRadius: 0
            }}
          >
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
              <polyline points="14 2 14 8 20 8" />
              <polyline points="9 15 11 17 15 12" />
            </svg>
          </div>
          <div>
            <h1 style={{ margin: 0, fontSize: '19px', fontWeight: 800, color: '#241a3d', letterSpacing: '-0.02em' }}>
              {config.title}
            </h1>
            <p style={{ margin: '3px 0 0 0', fontSize: '12.5px', color: '#64748b' }}>
              {config.subTitle} ({workingMonth})
            </p>
          </div>
        </div>

        {/* ONLY "Back to Steps" button */}
        <button
          type="button"
          onClick={onBack}
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: '8px',
            background: '#ffffff',
            color: '#350b5e',
            border: '1.5px solid #350b5e',
            padding: '8px 18px',
            fontSize: '13px',
            fontWeight: 700,
            cursor: 'pointer',
            borderRadius: 0,
            transition: 'all 0.15s ease'
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.background = '#f5f3ff';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.background = '#ffffff';
          }}
        >
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
            <line x1="19" y1="12" x2="5" y2="12"></line>
            <polyline points="12 19 5 12 12 5"></polyline>
          </svg>
          Back to Steps
        </button>
      </div>

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

      {/* 4. Controls: Search Input & Excel Export Button */}
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          gap: '12px',
          background: '#ffffff',
          border: '1.5px solid #e2e8f0',
          padding: '10px 16px',
          marginBottom: '12px',
          borderRadius: 0,
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
      <div
        style={{
          background: '#ffffff',
          border: '1.5px solid #e2e8f0',
          overflowX: 'auto',
          borderRadius: 0
        }}
      >
        <table
          style={{
            width: '100%',
            borderCollapse: 'collapse',
            fontSize: '12px',
            textAlign: 'left'
          }}
        >
          <thead>
            <tr style={{ background: '#350b5e', color: '#ffffff' }}>
              <th style={{ padding: '8px 10px', fontWeight: 700, width: '40px', borderRight: '1px solid rgba(255,255,255,0.2)', textAlign: 'center' }}>#</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, borderRight: '1px solid rgba(255,255,255,0.2)' }}>STATUS</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, borderRight: '1px solid rgba(255,255,255,0.2)' }}>2B MONTH</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, borderRight: '1px solid rgba(255,255,255,0.2)' }}>DATE</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, borderRight: '1px solid rgba(255,255,255,0.2)' }}>INVOICE NO.</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, borderRight: '1px solid rgba(255,255,255,0.2)' }}>SUPPLIER GSTIN</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, borderRight: '1px solid rgba(255,255,255,0.2)' }}>SUPPLIER NAME</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, textAlign: 'right', borderRight: '1px solid rgba(255,255,255,0.2)' }}>TAXABLE (₹)</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, textAlign: 'right', borderRight: '1px solid rgba(255,255,255,0.2)' }}>IGST (₹)</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, textAlign: 'right', borderRight: '1px solid rgba(255,255,255,0.2)' }}>CGST (₹)</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, textAlign: 'right', borderRight: '1px solid rgba(255,255,255,0.2)' }}>SGST (₹)</th>
              <th style={{ padding: '8px 12px', fontWeight: 700, textAlign: 'right' }}>TOTAL (₹)</th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.length === 0 ? (
              <tr>
                <td colSpan={12} style={{ padding: '36px 16px', textAlign: 'center', color: '#64748b' }}>
                  No invoices found matching selected criteria.
                </td>
              </tr>
            ) : (
              filteredRows.map((row, idx) => (
                <tr
                  key={row.id || idx}
                  style={{
                    background: idx % 2 === 0 ? '#ffffff' : '#f8fafc',
                    borderBottom: '1px solid #e2e8f0'
                  }}
                >
                  <td style={{ padding: '8px 10px', textAlign: 'center', color: '#64748b', fontWeight: 600 }}>
                    {idx + 1}
                  </td>
                  <td style={{ padding: '8px 12px', whiteSpace: 'nowrap' }}>
                    <span
                      style={{
                        display: 'inline-block',
                        padding: '3px 8px',
                        fontSize: '11px',
                        fontWeight: 700,
                        background: config.statusPillBg,
                        color: config.statusPillColor,
                        border: `1px solid ${config.statusPillBorder}`,
                        borderRadius: 0
                      }}
                    >
                      {row.status}
                    </span>
                  </td>
                  <td style={{ padding: '8px 12px', color: '#334155', fontWeight: 600 }}>
                    {row.month2b}
                  </td>
                  <td style={{ padding: '8px 12px', color: '#334155', whiteSpace: 'nowrap' }}>
                    {row.date}
                  </td>
                  <td style={{ padding: '8px 12px', color: '#0f172a', fontWeight: 700 }}>
                    {row.invNo}
                  </td>
                  <td style={{ padding: '8px 12px', fontFamily: 'monospace', fontSize: '11.5px', color: '#334155' }}>
                    {row.gstin}
                  </td>
                  <td style={{ padding: '8px 12px', color: '#0f172a', fontWeight: 600 }}>
                    {row.supplierName}
                  </td>
                  <td style={{ padding: '8px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: '#0f172a', fontWeight: 600 }}>
                    {row.taxable.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </td>
                  <td style={{ padding: '8px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: '#334155' }}>
                    {row.igst.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </td>
                  <td style={{ padding: '8px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: '#334155' }}>
                    {row.cgst.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </td>
                  <td style={{ padding: '8px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: '#334155' }}>
                    {row.sgst.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </td>
                  <td style={{ padding: '8px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: '#350b5e', fontWeight: 800 }}>
                    {row.total.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </td>
                </tr>
              ))
            )}
          </tbody>
          {filteredRows.length > 0 && (
            <tfoot>
              <tr style={{ background: '#240742', color: '#ffffff', fontWeight: 800 }}>
                <td colSpan={7} style={{ padding: '10px 14px', textAlign: 'left', letterSpacing: '0.04em' }}>
                  TOTAL ({filteredRows.length} INVOICES)
                </td>
                <td style={{ padding: '10px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>
                  {summaryValues.taxable.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </td>
                <td style={{ padding: '10px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>
                  {summaryValues.igst.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </td>
                <td style={{ padding: '10px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>
                  {summaryValues.cgst.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </td>
                <td style={{ padding: '10px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>
                  {summaryValues.sgst.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </td>
                <td style={{ padding: '10px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: '#fef08a' }}>
                  {summaryValues.grandTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </section>
  );
}
