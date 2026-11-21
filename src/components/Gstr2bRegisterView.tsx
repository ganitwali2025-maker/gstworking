import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { fm, fmtVal, dmySlash, inr, pp, iso } from '../utils/gstFormat';

export interface Gstr2bRecord {
  id: string;
  m: string;
  q: string;
  fy: string;
  d: string; // Date
  gstin: string; // GSTIN of Supplier
  tradeName: string; // Trade / Legal Name
  invNo: string; // Invoice No
  invVal: number; // Invoice Value
  taxable: number; // Taxable Value
  igst: number; // IGST
  cgst: number; // CGST
  sgst: number; // SGST
  cess: number; // CESS
}

interface Gstr2bRegisterViewProps {
  showToast?: (msg: string) => void;
}

export default function Gstr2bRegisterView({ showToast }: Gstr2bRegisterViewProps) {
  const [records, setRecords] = useState<Gstr2bRecord[]>(() => {
    try {
      const saved = localStorage.getItem('gstr2b_register_v1');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const [isLocked, setIsLocked] = useState(false);
  const [editingRecord, setEditingRecord] = useState<Gstr2bRecord | null>(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [pastRecords, setPastRecords] = useState<Gstr2bRecord[][]>([]);
  const [futureRecords, setFutureRecords] = useState<Gstr2bRecord[][]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const pushHistory = (prev: Gstr2bRecord[]) => {
    setPastRecords((p) => [...p.slice(-25), prev]);
    setFutureRecords([]);
  };

  const handleUndo = () => {
    if (pastRecords.length === 0) {
      notify('No more actions to undo');
      return;
    }
    const prev = pastRecords[pastRecords.length - 1];
    setFutureRecords((f) => [records, ...f]);
    setPastRecords((p) => p.slice(0, p.length - 1));
    setRecords(prev);
    try { localStorage.setItem('gstr2b_register_v1', JSON.stringify(prev)); } catch {}
    notify('Undo successful');
  };

  const handleRedo = () => {
    if (futureRecords.length === 0) {
      notify('No more actions to redo');
      return;
    }
    const next = futureRecords[0];
    setPastRecords((p) => [...p, records]);
    setFutureRecords((f) => f.slice(1));
    setRecords(next);
    try { localStorage.setItem('gstr2b_register_v1', JSON.stringify(next)); } catch {}
    notify('Redo successful');
  };

  // Totals for top KPI cards
  const totalTaxable = records.reduce((acc, r) => acc + (r.taxable || 0), 0);
  const totalIgst = records.reduce((acc, r) => acc + (r.igst || 0), 0);
  const totalCgst = records.reduce((acc, r) => acc + (r.cgst || 0), 0);
  const totalSgst = records.reduce((acc, r) => acc + (r.sgst || 0), 0);
  const totalTax = totalIgst + totalCgst + totalSgst + records.reduce((acc, r) => acc + (r.cess || 0), 0);

  const notify = (msg: string) => {
    if (showToast) showToast(msg);
    else alert(msg);
  };

  const persistRecords = (recs: Gstr2bRecord[]) => {
    pushHistory(records);
    setRecords(recs);
    try {
      localStorage.setItem('gstr2b_register_v1', JSON.stringify(recs));
    } catch {}
  };

  const handleOpenEdit = (rec: Gstr2bRecord) => {
    if (isLocked) {
      notify('Sheet is locked. Unlock to make changes.');
      return;
    }
    setEditingRecord(rec);
    setModalOpen(true);
  };

  const handleSaveRecord = (updated: Gstr2bRecord) => {
    persistRecords(records.map((r) => (r.id === updated.id ? updated : r)));
    setModalOpen(false);
    notify('Record updated successfully');
  };

  const handleClearData = () => {
    if (isLocked) {
      notify('Sheet is locked.');
      return;
    }
    if (window.confirm('Clear all GSTR-2B portal records?')) {
      persistRecords([]);
      notify('GSTR-2B data cleared');
    }
  };

  const handleDelete = (id: string) => {
    if (isLocked) {
      notify('Sheet is locked.');
      return;
    }
    persistRecords(records.filter((r) => r.id !== id));
    notify('Record deleted');
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array' });
      const firstSheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<any>(firstSheet, { defval: '' });

      const imported: Gstr2bRecord[] = [];
      for (const row of rows) {
        const d = iso(row['Date'] || row['Invoice Date'] || '');
        const gstin = String(row['GSTIN of Supplier'] || row['GSTIN'] || row['GST No.'] || '').trim();
        const tradeName = String(row['Trade / Legal Name'] || row['Legal Name'] || row['Supplier'] || '').trim();
        const invNo = String(row['Invoice No'] || row['Invoice'] || '').trim();
        const invVal = pp(row['Invoice Value'] || row['Total Invoice Value'] || 0);
        const taxable = pp(row['Taxable Value'] || row['Taxable'] || 0);
        const igst = pp(row['IGST'] || 0);
        const cgst = pp(row['CGST'] || 0);
        const sgst = pp(row['SGST'] || 0);
        const cess = pp(row['CESS'] || row['Cess'] || 0);

        if (!tradeName && !gstin) continue;

        imported.push({
          id: 'g2b-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
          m: 'Aug',
          q: 'Q2 (Jul-Sep)',
          fy: '2026-27',
          d: d || '2026-08-01',
          gstin,
          tradeName,
          invNo,
          invVal: invVal || (taxable + igst + cgst + sgst + cess),
          taxable: taxable || 0,
          igst: igst || 0,
          cgst: cgst || 0,
          sgst: sgst || 0,
          cess: cess || 0
        });
      }

      if (imported.length > 0) {
        persistRecords([...imported, ...records]);
        notify(`${imported.length} GSTR-2B portal records imported`);
      } else {
        notify('No valid rows found to import.');
      }
    } catch (err: any) {
      notify('Import error: ' + err.message);
    }
  };

  return (
    <section id="v-g1">
      <input
        type="file"
        ref={fileInputRef}
        accept=".xlsx,.xls,.csv,.json"
        style={{ display: 'none' }}
        onChange={handleFileUpload}
      />

      {/* 5 KPI Cards Grid */}
      <div className="gstr1-kpis-grid">
        {/* Card 1: Total Taxable */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total Taxable</div>
          <div className="gstr1-kpi-val">{inr(totalTaxable / 100)}</div>
        </div>

        {/* Card 2: Total IGST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="2" y1="12" x2="22" y2="12"></line>
                <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"></path>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total IGST</div>
          <div className="gstr1-kpi-val">{inr(totalIgst / 100)}</div>
        </div>

        {/* Card 3: Total CGST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="3" y1="21" x2="21" y2="21"></line>
                <line x1="3" y1="10" x2="21" y2="10"></line>
                <polyline points="5 6 12 3 19 6"></polyline>
                <line x1="4" y1="10" x2="4" y2="21"></line>
                <line x1="20" y1="10" x2="20" y2="21"></line>
                <line x1="8" y1="14" x2="8" y2="17"></line>
                <line x1="12" y1="14" x2="12" y2="17"></line>
                <line x1="16" y1="14" x2="16" y2="17"></line>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total CGST</div>
          <div className="gstr1-kpi-val">{inr(totalCgst / 100)}</div>
        </div>

        {/* Card 4: Total SGST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path>
                <circle cx="12" cy="10" r="3"></circle>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total SGST</div>
          <div className="gstr1-kpi-val">{inr(totalSgst / 100)}</div>
        </div>

        {/* Card 5: Total Tax */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <span style={{ fontSize: '16px', fontWeight: 700 }}>₹</span>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total Tax</div>
          <div className="gstr1-kpi-val">{inr(totalTax / 100)}</div>
        </div>
      </div>

      <div className="card gstr1-card">
        {/* Top Header Card */}
        <div className="gstr1-header">
          <div className="gstr1-title-box">
            <div className="gstr1-icon-badge">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
              </svg>
            </div>
            <div>
              <h2 className="gstr1-title">GSTR-2B – Govt Portal ITC Statement</h2>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="gstr1-btn-group">
            <button
              type="button"
              className="btn-gstr1 btn-lock"
              onClick={() => {
                setIsLocked(!isLocked);
                notify(isLocked ? 'Sheet unlocked' : 'Sheet locked');
              }}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                {isLocked ? (
                  <>
                    <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
                    <path d="M7 11V7a5 5 0 0 1 9.9-1"></path>
                  </>
                ) : (
                  <>
                    <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
                    <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
                  </>
                )}
              </svg>
              {isLocked ? 'Unlock Sheet' : 'Lock Sheet'}
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-clear"
              disabled={isLocked}
              onClick={handleClearData}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="3 6 5 6 21 6"></polyline>
                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
              </svg>
              Clear Data
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-refresh"
              onClick={() => notify('Register refreshed')}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"></path>
              </svg>
              Refresh
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-import"
              disabled={isLocked}
              onClick={() => fileInputRef.current?.click()}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="17 8 12 3 7 8"></polyline>
                <line x1="12" y1="3" x2="12" y2="15"></line>
              </svg>
              Import
            </button>
          </div>
        </div>

        {/* Table matching requested heading row */}
        <div className="gstr1-table-wrapper">
          <table className="gstr1-table">
            <thead>
              <tr>
                <th style={{ textAlign: 'center' }}>MONTH</th>
                <th style={{ textAlign: 'center' }}>QUARTER</th>
                <th style={{ textAlign: 'center' }}>DATE</th>
                <th style={{ textAlign: 'center' }}>GSTIN OF SUPPLIER</th>
                <th className="col-party" style={{ minWidth: '220px', textAlign: 'left' }}>TRADE / LEGAL NAME</th>
                <th className="col-inv" style={{ textAlign: 'left' }}>INVOICE NO</th>
                <th style={{ textAlign: 'center' }}>INVOICE VALUE</th>
                <th style={{ textAlign: 'center' }}>TAXABLE VALUE</th>
                <th style={{ textAlign: 'center' }}>IGST</th>
                <th style={{ textAlign: 'center' }}>CGST</th>
                <th style={{ textAlign: 'center' }}>SGST</th>
                <th style={{ textAlign: 'center' }}>CESS</th>
                <th style={{ textAlign: 'center' }}>TOTAL TAX</th>
                <th style={{ textAlign: 'center' }}>ACTION</th>
              </tr>
            </thead>
            <tbody>
              {records.length > 0 ? (
                records.map((r) => {
                  const totalTax = (r.igst || 0) + (r.cgst || 0) + (r.sgst || 0) + (r.cess || 0);

                  return (
                    <tr key={r.id}>
                      <td style={{ textAlign: 'center' }}>{r.m}</td>
                      <td style={{ textAlign: 'center' }}>{r.q}</td>
                      <td style={{ textAlign: 'center' }}>{dmySlash(r.d)}</td>
                      <td className="gst-no" style={{ textAlign: 'center' }}>{r.gstin || '—'}</td>
                      <td className="party-name" style={{ textAlign: 'left' }} title={r.tradeName}>
                        {r.tradeName && r.tradeName.length > 34 ? r.tradeName.slice(0, 34) + '...' : r.tradeName}
                      </td>
                      <td className="inv-no" style={{ textAlign: 'left' }}>{r.invNo}</td>
                      <td className="num-val" style={{ fontWeight: 600, textAlign: 'center' }}>{fmtVal(r.invVal)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.taxable)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.igst)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.cgst)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.sgst)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.cess)}</td>
                      <td className="num-val total-tax" style={{ textAlign: 'center' }}>{fmtVal(totalTax)}</td>
                      <td className="act-col" style={{ textAlign: 'center', whiteSpace: 'nowrap' }}>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={() => handleOpenEdit(r)}
                          className="row-edit-btn"
                          title="Edit record"
                        >
                          Edit
                        </button>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={() => handleDelete(r.id)}
                          className="row-delete-btn"
                          title="Delete record"
                        >
                          Delete
                        </button>
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={14} style={{ textAlign: 'center', padding: '36px', color: '#64748b' }}>
                    No rows found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Edit GSTR-2B Modal */}
      {modalOpen && editingRecord && (
        <Gstr2bEditModal
          isOpen={modalOpen}
          record={editingRecord}
          isLocked={isLocked}
          onClose={() => setModalOpen(false)}
          onSave={handleSaveRecord}
        />
      )}
    </section>
  );
}

interface Gstr2bEditModalProps {
  isOpen: boolean;
  record: Gstr2bRecord;
  isLocked: boolean;
  onClose: () => void;
  onSave: (record: Gstr2bRecord) => void;
}

function Gstr2bEditModal({ isOpen, record, isLocked, onClose, onSave }: Gstr2bEditModalProps) {
  if (!isOpen) return null;

  const [d, setD] = useState(record.d);
  const [gstin, setGstin] = useState(record.gstin);
  const [tradeName, setTradeName] = useState(record.tradeName);
  const [invNo, setInvNo] = useState(record.invNo);
  const [taxable, setTaxable] = useState((record.taxable / 100).toString());
  const [igst, setIgst] = useState((record.igst / 100).toString());
  const [cgst, setCgst] = useState((record.cgst / 100).toString());
  const [sgst, setSgst] = useState((record.sgst / 100).toString());
  const [cess, setCess] = useState((record.cess / 100).toString());
  const [error, setError] = useState('');

  const tVal = pp(taxable) || 0;
  const iVal = pp(igst) || 0;
  const cVal = pp(cgst) || 0;
  const sVal = pp(sgst) || 0;
  const cessVal = pp(cess) || 0;
  const totalTax = iVal + cVal + sVal + cessVal;
  const totalInv = tVal + totalTax;

  const handleSubmit = () => {
    if (isLocked) {
      setError('Sheet is locked.');
      return;
    }
    if (!d) return setError('Enter date.');
    if (!invNo.trim()) return setError('Enter invoice number.');
    if (!tradeName.trim()) return setError('Enter supplier trade/legal name.');

    onSave({
      ...record,
      d,
      tradeName: tradeName.trim(),
      gstin: (gstin || '').trim().toUpperCase(),
      invNo: invNo.trim(),
      taxable: tVal,
      igst: iVal,
      cgst: cVal,
      sgst: sVal,
      cess: cessVal,
      invVal: totalInv
    });
  };

  return (
    <div className="mo">
      <div className="mb">
        <h3>Edit GSTR-2B Record</h3>
        <div style={{ color: 'var(--mute)', fontSize: '13px', marginBottom: '12px' }}>
          Period: {record.m} | {record.fy}
        </div>

        <div className="fg">
          <div>
            <label>Date</label>
            <input type="date" value={d} onChange={(e) => setD(e.target.value)} />
          </div>
          <div>
            <label>GSTIN of Supplier</label>
            <input maxLength={15} value={gstin} onChange={(e) => setGstin(e.target.value.toUpperCase())} />
          </div>
          <div>
            <label>Trade / Legal Name</label>
            <input value={tradeName} onChange={(e) => setTradeName(e.target.value)} />
          </div>
          <div>
            <label>Invoice No</label>
            <input value={invNo} onChange={(e) => setInvNo(e.target.value)} />
          </div>
          <div>
            <label>Taxable Value (₹)</label>
            <input inputMode="decimal" value={taxable} onChange={(e) => setTaxable(e.target.value)} />
          </div>
          <div>
            <label>IGST (₹)</label>
            <input inputMode="decimal" value={igst} onChange={(e) => setIgst(e.target.value)} />
          </div>
          <div>
            <label>CGST (₹)</label>
            <input inputMode="decimal" value={cgst} onChange={(e) => setCgst(e.target.value)} />
          </div>
          <div>
            <label>SGST (₹)</label>
            <input inputMode="decimal" value={sgst} onChange={(e) => setSgst(e.target.value)} />
          </div>
          <div>
            <label>CESS (₹)</label>
            <input inputMode="decimal" value={cess} onChange={(e) => setCess(e.target.value)} />
          </div>
        </div>

        <div style={{ marginTop: '14px', display: 'flex', gap: '24px', flexWrap: 'wrap' }}>
          <span>Total Tax: <b>{fm(totalTax)}</b></span>
          <span>Invoice Value: <b>{fm(totalInv)}</b></span>
        </div>

        {error && <div className="fe">{error}</div>}

        <div className="fr">
          <button type="button" onClick={onClose}>Cancel</button>
          <button type="button" className="pri" onClick={handleSubmit}>Save Changes</button>
        </div>
      </div>
    </div>
  );
}
