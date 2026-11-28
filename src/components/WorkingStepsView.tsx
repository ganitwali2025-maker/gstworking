import React, { useMemo, useState, useEffect } from 'react';
import { StepItem, ViewType } from '../types/gst';
import RegisterPeriodBar from './RegisterPeriodBar';
import { useLiveReconciliation } from '../services/recoSyncService';
import { syncOldItcPool } from '../services/oldItcService';
import { parseWorkingMonth } from '../services/lockService';

interface WorkingStepsViewProps {
  steps: StepItem[];
  toggleStepDone?: (id: number) => void;
  setView: (view: ViewType) => void;
  showToast: (msg: string) => void;
  fy?: string;
  month?: string;
}

function formatInr(val: number): string {
  if (val === null || val === undefined || isNaN(val)) return '0.00';
  return val.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

export default function WorkingStepsView({
  steps,
  setView,
  showToast,
  fy = 'FY 2026–27',
  month
}: WorkingStepsViewProps) {
  const [oldItcVersion, setOldItcVersion] = useState(0);

  const workingMonth = useMemo(() => {
    return parseWorkingMonth(month, fy);
  }, [fy, month]);

  // Live one-way sync with Reco sheet
  const recoData = useLiveReconciliation(workingMonth);

  // Live sync with Old ITC pool
  const oldItcData = useMemo(() => {
    try {
      return syncOldItcPool(workingMonth);
    } catch (e) {
      console.error('Failed to sync Old ITC pool:', e);
      return {
        oldItcItems: [],
        pendingItems: [],
        matchedClaimItems: [],
        notClaimItems: [],
        supplierRiskSummary: [],
        pendingCount: 0,
        pendingTotalItc: 0,
        notClaimCount: 0,
        notClaimTotalItc: 0
      };
    }
  }, [workingMonth, oldItcVersion]);

  // Listen to old-itc-updated events
  useEffect(() => {
    const handleUpdate = () => setOldItcVersion((v) => v + 1);
    window.addEventListener('old-itc-updated', handleUpdate);
    return () => window.removeEventListener('old-itc-updated', handleUpdate);
  }, []);

  const handleOpenModule = (index: number, stepTitle: string) => {
    if (index === 0) {
      setView('g1');
    } else if (index === 1) {
      setView('itc-reg');
    } else if (index === 2) {
      setView('rcm');
    } else if (index === 3) {
      setView('g2b');
    } else if (index === 4) {
      setView('rec');
    } else {
      showToast(`${stepTitle} module opened.`);
    }
  };

  const getStepIcon = (index: number) => {
    switch (index) {
      case 0:
        return (
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <polyline points="14 2 14 8 20 8" />
            <line x1="16" y1="13" x2="8" y2="13" />
            <line x1="16" y1="17" x2="8" y2="17" />
            <line x1="10" y1="9" x2="8" y2="9" />
          </svg>
        );
      case 1:
        return (
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z" />
            <line x1="3" y1="6" x2="21" y2="6" />
            <path d="M16 10a4 4 0 0 1-8 0" />
          </svg>
        );
      case 2:
        return (
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="17 1 21 5 17 9" />
            <path d="M3 11V9a4 4 0 0 1 4-4h14" />
            <polyline points="7 23 3 19 7 15" />
            <path d="M21 13v2a4 4 0 0 1-4 4H3" />
          </svg>
        );
      case 3:
        return (
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <polyline points="14 2 14 8 20 8" />
            <polyline points="9 15 11 17 15 12" />
          </svg>
        );
      case 4:
      default:
        return (
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="20 6 9 17 4 12" />
            <polyline points="22 10 13 19 11 17" />
          </svg>
        );
    }
  };

  // Steps 1 to 4
  const primarySteps = (steps || []).slice(0, 4);

  // Step 5 (ITC Reconciliation)
  const step5 = (steps || [])[4] || {
    id: 5,
    titleEn: 'ITC Reconciliation',
    titleHi: 'आईटीसी मिलान',
    descEn: 'Compare Purchase Register with GSTR-2B to maximize verified ITC.',
    descHi: 'सटीक आईटीसी क्लेम के लिए खरीद रजिस्टर का GSTR-2B से मिलान करें।',
    done: false
  };

  return (
    <>
      <RegisterPeriodBar view="steps" />
      <section id="v-steps" className="steps-container">
        <div className="steps-cards-grid">
          {/* Steps 1 to 4 */}
          {primarySteps.map((s, index) => (
            <div
              key={s.id}
              className={`step-card ${s.done ? 'step-done' : ''}`}
              onClick={() => handleOpenModule(index, s.titleEn)}
              style={{ cursor: 'pointer' }}
            >
              <div className="step-card-header">
                <div className="step-card-icon-box" title={s.titleEn}>
                  {getStepIcon(index)}
                </div>
                <span className="step-card-badge">Step {index + 1}</span>
              </div>

              <div className="step-card-body">
                <h3 className="step-card-title">{s.titleEn}</h3>
                <p className="step-card-desc-en">{s.descEn}</p>
                <p className="step-card-desc-hi">{s.descHi}</p>
                {s.notes && s.notes.length > 0 && (
                  <div className="step-notes-list">
                    {s.notes.map((note, nIdx) => (
                      <div key={nIdx} className="step-note-item">
                        <span>• {note}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div className="step-card-footer">
                <span
                  className="step-workplace-link"
                  onClick={(e) => {
                    e.stopPropagation();
                    handleOpenModule(index, s.titleEn);
                  }}
                  role="button"
                  tabIndex={0}
                >
                  Open your workplace →
                </span>
                <button
                  type="button"
                  className="step-action-btn btn-add-data"
                  onClick={(e) => {
                    e.stopPropagation();
                    handleOpenModule(index, s.titleEn);
                  }}
                  style={{
                    background: 'linear-gradient(135deg, #4c1d95 0%, #3b0764 100%)',
                    color: '#ffffff',
                    boxShadow: '0 4px 14px rgba(76, 29, 149, 0.45)',
                    border: 'none',
                    fontWeight: 600,
                    cursor: 'pointer'
                  }}
                >
                  Add GST Data
                </button>
              </div>
            </div>
          ))}

          {/* =========================================================
              THE 6 RECONCILIATION CARDS IN EXACT ORDER (Left to Right)
              1. ITC Reconciliation (purana, Step 5)
              2. Old ITC (NAYA, violet #6D28D9)
              3. Matched Invoices - ITC Verified (purana)
              4. Not Claim ITC (NAYA, gehra laal #7F1D1D)
              5. Not in Purchase Register (purana)
              6. Not in GSTR-2B (purana)
             ========================================================= */}

          {/* CARD 1: ITC Reconciliation (Step 5) */}
          <div
            className={`step-card ${step5.done ? 'step-done' : ''}`}
            onClick={() => setView('rec')}
            style={{ cursor: 'pointer' }}
          >
            <div className="step-card-header">
              <div className="step-card-icon-box" title={step5.titleEn}>
                {getStepIcon(4)}
              </div>
              <span className="step-card-badge">Step 5</span>
            </div>
            <div className="step-card-body">
              <h3 className="step-card-title">{step5.titleEn}</h3>
              <p className="step-card-desc-en">{step5.descEn}</p>
              <p className="step-card-desc-hi">{step5.descHi}</p>
            </div>
            <div className="step-card-footer">
              <span
                className="step-workplace-link"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('rec');
                }}
                role="button"
                tabIndex={0}
              >
                Open your workplace →
              </span>
              <button
                type="button"
                className="step-action-btn btn-reconcile"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('rec');
                }}
                style={{
                  background: 'linear-gradient(135deg, #4c1d95 0%, #3b0764 100%)',
                  color: '#ffffff',
                  boxShadow: '0 4px 14px rgba(76, 29, 149, 0.45)',
                  border: 'none',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                Start Reconciliation
              </button>
            </div>
          </div>

          {/* CARD 2: Old ITC - Carried Forward (NAYA, violet #6D28D9) */}
          <div
            className="step-card step-card-violet"
            onClick={() => setView('old-itc')}
            style={{ cursor: 'pointer' }}
          >
            <div className="step-card-header">
              <div className="step-card-icon-box icon-violet" title="Old ITC - Carried Forward">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#6D28D9" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="10" />
                  <polyline points="12 6 12 12 16 14" />
                </svg>
              </div>
              <span
                className="step-card-badge"
                style={{ color: '#6D28D9', fontWeight: 700, fontSize: '13px' }}
              >
                {oldItcData.matchedCount} Invoices
              </span>
            </div>
            <div className="step-card-body">
              <h3 className="step-card-title">Old ITC - Carried Forward</h3>
              <p className="step-card-desc-en">
                Earlier months' unmatched ITC, matched and verified in a later month.
              </p>
              <p className="step-card-desc-hi">
                पिछले महीनों का अनमैच ITC, बाद के महीने में मिलान होकर सत्यापित।
              </p>
            </div>
            <div className="step-card-footer">
              <span
                className="step-workplace-link"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('old-itc');
                }}
                role="button"
                tabIndex={0}
              >
                Open your workplace →
              </span>
              <button
                type="button"
                className="step-action-btn"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('old-itc');
                }}
                style={{
                  background: 'linear-gradient(135deg, #4c1d95 0%, #3b0764 100%)',
                  color: '#ffffff',
                  boxShadow: '0 4px 14px rgba(76, 29, 149, 0.45)',
                  border: 'none',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                Open Old ITC
              </button>
            </div>
          </div>

          {/* CARD 3: Matched Invoices - ITC Verified (purana, dark green) */}
          <div
            className="step-card step-card-green"
            onClick={() => setView('reco-matched')}
            style={{ cursor: 'pointer' }}
          >
            <div className="step-card-header">
              <div className="step-card-icon-box icon-green" title="Matched Invoices – ITC Verified">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
                  <polyline points="22 4 12 14.01 9 11.01" />
                </svg>
              </div>
              <span
                className="step-card-badge"
                style={{ color: '#14532d', fontWeight: 700, fontSize: '13px' }}
              >
                {recoData.summary.matchedCount} Invoices
              </span>
            </div>
            <div className="step-card-body">
              <h3 className="step-card-title">Matched Invoices – ITC Verified</h3>
              <p className="step-card-desc-en">
                Found in both purchase register and GSTR-2B. ITC may be claimed.
              </p>
              <p className="step-card-desc-hi">
                बुक रजिस्टर एवं GSTR-2B दोनों में उपलब्ध। ITC का दावा किया जा सकता है।
              </p>
            </div>
            <div className="step-card-footer">
              <span
                className="step-workplace-link"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('reco-matched');
                }}
                role="button"
                tabIndex={0}
              >
                Open your workplace →
              </span>
              <button
                type="button"
                className="step-action-btn btn-green"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('reco-matched');
                }}
                style={{
                  background: 'linear-gradient(135deg, #14532d 0%, #166534 100%)',
                  color: '#ffffff',
                  boxShadow: '0 4px 14px rgba(20, 83, 45, 0.45)',
                  border: 'none',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                Open Matched ITC
              </button>
            </div>
          </div>

          {/* CARD 4: Not Claim ITC (NAYA, gehra laal #7F1D1D) */}
          <div
            className="step-card step-card-darkred"
            onClick={() => setView('not-claim-itc')}
            style={{ cursor: 'pointer' }}
          >
            <div className="step-card-header">
              <div className="step-card-icon-box icon-darkred" title="Not Claim / Not Match ITC">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#7F1D1D" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <polygon points="7.86 2 16.14 2 22 7.86 22 16.14 16.14 22 7.86 22 2 16.14 2 7.86 7.86 2" />
                  <line x1="12" y1="8" x2="12" y2="12" />
                  <line x1="12" y1="16" x2="12.01" y2="16" />
                </svg>
              </div>
              <span
                className="step-card-badge"
                style={{ color: '#7F1D1D', fontWeight: 700, fontSize: '13px' }}
              >
                {oldItcData.notClaimCount} Invoices
              </span>
            </div>
            <div className="step-card-body">
              <h3 className="step-card-title">Not Claim / Not Match ITC</h3>
              <p className="step-card-desc-en">
                Not Match and Not in 2B invoices, still unmatched after re-check. Follow up with supplier.
              </p>
              <p className="step-card-desc-hi">
                नॉट मैच एवं 2B में अनुपलब्ध इनवॉइस, दोबारा जाँच के बाद भी मिलान नहीं। सप्लायर से संपर्क करें।
              </p>
              <div style={{ marginTop: '8px', fontSize: '12.5px', fontWeight: 700, color: '#7F1D1D' }}>
                ITC at risk: ₹{formatInr(oldItcData.notClaimTotalItc)}
              </div>
            </div>
            <div className="step-card-footer">
              <span
                className="step-workplace-link"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('not-claim-itc');
                }}
                role="button"
                tabIndex={0}
              >
                Open your workplace →
              </span>
              <button
                type="button"
                className="step-action-btn btn-darkred"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('not-claim-itc');
                }}
                style={{
                  background: 'linear-gradient(135deg, #7f1d1d 0%, #991b1b 100%)',
                  color: '#ffffff',
                  boxShadow: '0 4px 14px rgba(127, 29, 29, 0.45)',
                  border: 'none',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                Open Not Claim / Not Match
              </button>
            </div>
          </div>

          {/* CARD 5: Not in Purchase Register (purana, dark red) */}
          <div
            className="step-card step-card-red"
            onClick={() => setView('reco-missing-books')}
            style={{ cursor: 'pointer' }}
          >
            <div className="step-card-header">
              <div className="step-card-icon-box icon-red" title="Not in Purchase Register">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#7f1d1d" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20" />
                  <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z" />
                  <line x1="10" y1="10" x2="14" y2="14" />
                  <line x1="14" y1="10" x2="10" y2="14" />
                </svg>
              </div>
              <span
                className="step-card-badge"
                style={{ color: '#7f1d1d', fontWeight: 700, fontSize: '13px' }}
              >
                {recoData.summary.notInBookCount} Invoices
              </span>
            </div>
            <div className="step-card-body">
              <h3 className="step-card-title">Not in Purchase Register</h3>
              <p className="step-card-desc-en">
                Appears in GSTR-2B but not recorded in your books. Record it.
              </p>
              <p className="step-card-desc-hi">
                GSTR-2B में उपलब्ध, परंतु बुक रजिस्टर में दर्ज नहीं। बुक में प्रविष्टि करें।
              </p>
            </div>
            <div className="step-card-footer">
              <span
                className="step-workplace-link"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('reco-missing-books');
                }}
                role="button"
                tabIndex={0}
              >
                Open your workplace →
              </span>
              <button
                type="button"
                className="step-action-btn btn-red"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('reco-missing-books');
                }}
                style={{
                  background: 'linear-gradient(135deg, #7f1d1d 0%, #991b1b 100%)',
                  color: '#ffffff',
                  boxShadow: '0 4px 14px rgba(127, 29, 29, 0.45)',
                  border: 'none',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                Open Missing in Books
              </button>
            </div>
          </div>

          {/* CARD 6: Not in GSTR-2B (purana, dark red) */}
          <div
            className="step-card step-card-red"
            onClick={() => setView('reco-missing-2b')}
            style={{ cursor: 'pointer' }}
          >
            <div className="step-card-header">
              <div className="step-card-icon-box icon-red" title="Not in GSTR-2B">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#7f1d1d" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                  <polyline points="14 2 14 8 20 8" />
                  <line x1="9" y1="14" x2="15" y2="14" />
                </svg>
              </div>
              <span
                className="step-card-badge"
                style={{ color: '#7f1d1d', fontWeight: 700, fontSize: '13px' }}
              >
                {recoData.summary.notIn2bCount} Invoices
              </span>
            </div>
            <div className="step-card-body">
              <h3 className="step-card-title">Not in GSTR-2B</h3>
              <p className="step-card-desc-en">
                Recorded in your books but not reflected in GSTR-2B. Hold ITC.
              </p>
              <p className="step-card-desc-hi">
                बुक रजिस्टर में दर्ज, परंतु GSTR-2B में अनुपलब्ध। ITC रोककर रखें।
              </p>
            </div>
            <div className="step-card-footer">
              <span
                className="step-workplace-link"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('reco-missing-2b');
                }}
                role="button"
                tabIndex={0}
              >
                Open your workplace →
              </span>
              <button
                type="button"
                className="step-action-btn btn-red"
                onClick={(e) => {
                  e.stopPropagation();
                  setView('reco-missing-2b');
                }}
                style={{
                  background: 'linear-gradient(135deg, #7f1d1d 0%, #991b1b 100%)',
                  color: '#ffffff',
                  boxShadow: '0 4px 14px rgba(127, 29, 29, 0.45)',
                  border: 'none',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                Open Missing in 2B
              </button>
            </div>
          </div>
        </div>
      </section>
    </>
  );
}
