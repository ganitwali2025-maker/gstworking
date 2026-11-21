import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { fm, fmtVal, dmySlash, inr, pp, iso } from '../utils/gstFormat';

export interface RcmRecord {
  id: string;
  m: string;
  q: string;
  fy: string;
  d: string; // Entry Date
  transporter: string; // Transporter Name
  lrNo: string; // Transporter L.R. No.
  amount: number; // Amount
  igst: number; // IGST 5%
  cgst: number; // CGST 2.5%
  sgst: number; // SGST 2.5%
}

const SEED_RCM_ROWS: RcmRecord[] = [
  {
    id: 'rcm-1',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: '2026-27',
    d: '2026-08-03',
    transporter: 'VRL Logistics Limited',
    lrNo: 'LR-78901',
    amount: 15000000,
    igst: 0,
    cgst: 375000,
    sgst: 375000
  },
  {
    id: 'rcm-2',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: '2026-27',
    d: '2026-08-07',
    transporter: 'Gati KWE Express Surface',
    lrNo: 'LR-44219',
    amount: 8000000,
    igst: 400000,
    cgst: 0,
    sgst: 0
  },
  {
    id: 'rcm-3',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: '2026-27',
    d: '2026-08-12',
    transporter: 'SafeXpress Surface Logistics',
    lrNo: 'LR-66102',
    amount: 24000000,
    igst: 0,
    cgst: 600000,
    sgst: 600000
  },
  {
    id: 'rcm-4',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: '2026-27',
    d: '2026-08-18',
    transporter: 'TCI Freight Logistics',
    lrNo: 'LR-99341',
    amount: 19000000,
    igst: 950000,
    cgst: 0,
    sgst: 0
  }
];

interface RcmRegisterViewProps {
  showToast?: (msg: string) => void;
}

export default function RcmRegisterView({ showToast }: RcmRegisterViewProps) {
  const [records, setRecords] = useState<RcmRecord[]>(() => {
    try {
      const saved = localStorage.getItem('rcm_register_v1');
      return saved ? JSON.parse(saved) : SEED_RCM_ROWS;
    } catch {
      return SEED_RCM_ROWS;
    }
  });

  const [isLocked, setIsLocked] = useState(false);
  const [editingRecord, setEditingRecord] = useState<RcmRecord | null>(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [pastRecords, setPastRecords] = useState<RcmRecord[][]>([]);
  const [futureRecords, setFutureRecords] = useState<RcmRecord[][]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const pushHistory = (prev: RcmRecord[]) => {
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
    try { localStorage.setItem('rcm_register_v1', JSON.stringify(prev)); } catch {}
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
    try { localStorage.setItem('rcm_register_v1', JSON.stringify(next)); } catch {}
    notify('Redo successful');
  };

  // Totals for top KPI cards
  const totalAmount = records.reduce((acc, r) => acc + (r.amount || 0), 0);
  const totalIgst = records.reduce((acc, r) => acc + (r.igst || 0), 0);
  const totalCgst = records.reduce((acc, r) => acc + (r.cgst || 0), 0);
  const totalSgst = records.reduce((acc, r) => acc + (r.sgst || 0), 0);
  const totalTax = totalIgst + totalCgst + totalSgst;

  const notify = (msg: string) => {
    if (showToast) showToast(msg);
    else alert(msg);
  };

  const persistRecords = (recs: RcmRecord[]) => {
    pushHistory(records);
    setRecords(recs);
    try {
      localStorage.setItem('rcm_register_v1', JSON.stringify(recs));
    } catch {}
  };

  const handleOpenEdit = (rec: RcmRecord) => {
    if (isLocked) {
      notify('Sheet is locked. Unlock to make changes.');
      return;
    }
    setEditingRecord(rec);
    setModalOpen(true);
  };

  const handleSaveRecord = (updated: RcmRecord) => {
    persistRecords(records.map((r) => (r.id === updated.id ? updated : r)));
    setModalOpen(false);
    notify('Record updated successfully');
  };

  const handleClearData = () => {
    if (isLocked) {
      notify('Sheet is locked.');
      return;
    }
    if (window.confirm('Clear all RCM register records?')) {
      persistRecords([]);
      notify('RCM data cleared');
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

      const imported: RcmRecord[] = [];
      for (const row of rows) {
        const d = iso(row['Entry Date'] || row['Date'] || '');
        const transporter = String(row['Transporter Name'] || row['Transporter'] || '').trim();
        const lrNo = String(row['Transporter L.R. No.'] || row['L.R. No.'] || row['LR No'] || '').trim();
        const amount = pp(row['Amount'] || 0);
        const igst = pp(row['IGST 5%'] || row['IGST'] || 0);
        const cgst = pp(row['CGST 2.5%'] || row['CGST'] || 0);
        const sgst = pp(row['SGST 2.5%'] || row['SGST'] || 0);

        if (!transporter) continue;

        imported.push({
          id: 'rcm-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
          m: 'Aug',
          q: 'Q2 (Jul-Sep)',
          fy: '2026-27',
          d: d || '2026-08-01',
          transporter,
          lrNo,
          amount: amount || 0,
          igst: igst || 0,
          cgst: cgst || 0,
          sgst: sgst || 0
        });
      }

      if (imported.length > 0) {
        persistRecords([...imported, ...records]);
        notify(`${imported.length} RCM records imported`);
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
        accept=".xlsx,.xls,.csv"
        style={{ display: 'none' }}
        onChange={handleFileUpload}
      />

      {/* 5 KPI Cards Grid */}
      <div className="gstr1-kpis-grid">
        {/* Card 1: Total Amount */}
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
          <div className="gstr1-kpi-label">Total Amount</div>
          <div className="gstr1-kpi-val">{inr(totalAmount / 100)}</div>
        </div>

        {/* Card 2: Total IGST 5% */}
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

        {/* Card 3: Total CGST 2.5% */}
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

        {/* Card 4: Total SGST 2.5% */}
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
                <rect x="1" y="3" width="15" height="13"></rect>
                <polygon points="16 8 20 8 23 11 23 16 16 16 16 8"></polygon>
                <circle cx="5.5" cy="18.5" r="2.5"></circle>
                <circle cx="18.5" cy="18.5" r="2.5"></circle>
              </svg>
            </div>
            <div>
              <h2 className="gstr1-title">Reverse Charge Mechanism (RCM) Register</h2>
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
                <th style={{ textAlign: 'center' }}>ENTRY DATE</th>
                <th className="col-party" style={{ minWidth: '220px', textAlign: 'left' }}>TRANSPORTER NAME</th>
                <th className="col-inv" style={{ textAlign: 'left' }}>TRANSPORTER L.R. NO.</th>
                <th style={{ textAlign: 'center' }}>AMOUNT</th>
                <th style={{ textAlign: 'center' }}>IGST 5%</th>
                <th style={{ textAlign: 'center' }}>CGST 2.5%</th>
                <th style={{ textAlign: 'center' }}>SGST 2.5%</th>
                <th style={{ textAlign: 'center' }}>TOTAL TAX</th>
                <th style={{ textAlign: 'center' }}>ACTION</th>
              </tr>
            </thead>
            <tbody>
              {records.length > 0 ? (
                records.map((r) => {
                  const totalTax = (r.igst || 0) + (r.cgst || 0) + (r.sgst || 0);

                  return (
                    <tr key={r.id}>
                      <td style={{ textAlign: 'center' }}>{r.m}</td>
                      <td style={{ textAlign: 'center' }}>{r.q}</td>
                      <td style={{ textAlign: 'center' }}>{dmySlash(r.d)}</td>
                      <td className="party-name" style={{ textAlign: 'left' }} title={r.transporter}>
                        {r.transporter}
                      </td>
                      <td className="inv-no" style={{ textAlign: 'left' }}>{r.lrNo || '—'}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.amount)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.igst)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.cgst)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.sgst)}</td>
                      <td className="num-val total-tax" style={{ textAlign: 'center' }}>{fmtVal(totalTax)}</td>
                      <td className="act-col" style={{ textAlign: 'center', whiteSpace: 'nowrap' }}>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={() => handleOpenEdit(r)}
                          className="row-edit-btn"
                          title="Edit entry"
                        >
                          Edit
                        </button>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={() => handleDelete(r.id)}
                          className="row-delete-btn"
                          title="Delete entry"
                        >
                          Delete
                        </button>
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={11} style={{ textAlign: 'center', padding: '36px', color: '#64748b' }}>
                    No rows found. Click “Import” to load records.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Edit RCM Modal */}
      {modalOpen && editingRecord && (
        <RcmEditModal
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

interface RcmEditModalProps {
  isOpen: boolean;
  record: RcmRecord;
  isLocked: boolean;
  onClose: () => void;
  onSave: (record: RcmRecord) => void;
}

function RcmEditModal({ isOpen, record, isLocked, onClose, onSave }: RcmEditModalProps) {
  if (!isOpen) return null;

  const [d, setD] = useState(record.d);
  const [transporter, setTransporter] = useState(record.transporter);
  const [lrNo, setLrNo] = useState(record.lrNo);
  const [amount, setAmount] = useState((record.amount / 100).toString());
  const [igst, setIgst] = useState((record.igst / 100).toString());
  const [cgst, setCgst] = useState((record.cgst / 100).toString());
  const [sgst, setSgst] = useState((record.sgst / 100).toString());
  const [error, setError] = useState('');

  const amtVal = pp(amount) || 0;
  const iVal = pp(igst) || 0;
  const cVal = pp(cgst) || 0;
  const sVal = pp(sgst) || 0;
  const totalTax = iVal + cVal + sVal;

  const handleSubmit = () => {
    if (isLocked) {
      setError('Sheet is locked.');
      return;
    }
    if (!d) return setError('Enter entry date.');
    if (!transporter.trim()) return setError('Enter transporter name.');

    onSave({
      ...record,
      d,
      transporter: transporter.trim(),
      lrNo: lrNo.trim(),
      amount: amtVal,
      igst: iVal,
      cgst: cVal,
      sgst: sVal
    });
  };

  return (
    <div className="mo">
      <div className="mb">
        <h3>Edit RCM Entry</h3>
        <div style={{ color: 'var(--mute)', fontSize: '13px', marginBottom: '12px' }}>
          Period: {record.m} | {record.fy}
        </div>

        <div className="fg">
          <div>
            <label>Entry Date</label>
            <input type="date" value={d} onChange={(e) => setD(e.target.value)} />
          </div>
          <div>
            <label>Transporter Name</label>
            <input value={transporter} onChange={(e) => setTransporter(e.target.value)} />
          </div>
          <div>
            <label>Transporter L.R. No.</label>
            <input value={lrNo} onChange={(e) => setLrNo(e.target.value)} />
          </div>
          <div>
            <label>Amount (₹)</label>
            <input inputMode="decimal" value={amount} onChange={(e) => setAmount(e.target.value)} />
          </div>
          <div>
            <label>IGST 5% (₹)</label>
            <input inputMode="decimal" value={igst} onChange={(e) => setIgst(e.target.value)} />
          </div>
          <div>
            <label>CGST 2.5% (₹)</label>
            <input inputMode="decimal" value={cgst} onChange={(e) => setCgst(e.target.value)} />
          </div>
          <div>
            <label>SGST 2.5% (₹)</label>
            <input inputMode="decimal" value={sgst} onChange={(e) => setSgst(e.target.value)} />
          </div>
        </div>

        <div style={{ marginTop: '14px', display: 'flex', gap: '24px', flexWrap: 'wrap' }}>
          <span>Total Tax: <b>{fm(totalTax)}</b></span>
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
