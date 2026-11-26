import React, { useState } from 'react';
import { StepItem, ViewType } from '../types/gst';
import WorkingStepsDetailSheet, { WorkingDetailCategory } from './WorkingStepsDetailSheet';
import RegisterPeriodBar from './RegisterPeriodBar';

interface WorkingStepsViewProps {
  steps: StepItem[];
  toggleStepDone?: (id: number) => void;
  setView: (view: ViewType) => void;
  showToast: (msg: string) => void;
}

interface ExtraRecoCard {
  id: string;
  title: string;
  descEn: string;
  descHi: string;
  btnLabel: string;
  btnColor: 'green' | 'red';
  targetTab: WorkingDetailCategory;
}

const EXTRA_RECO_CARDS: ExtraRecoCard[] = [
  {
    id: 'reco-matched',
    title: 'Matched Invoices – ITC Verified',
    descEn: 'Found in both purchase register and GSTR-2B. ITC may be claimed.',
    descHi: 'बुक रजिस्टर एवं GSTR-2B दोनों में उपलब्ध। ITC का दावा किया जा सकता है।',
    btnLabel: 'Open Matched ITC',
    btnColor: 'green',
    targetTab: 'matched'
  },
  {
    id: 'reco-missing-books',
    title: 'Not in Purchase Register',
    descEn: 'Appears in GSTR-2B but not recorded in your books. Record it.',
    descHi: 'GSTR-2B में उपलब्ध, परंतु बुक रजिस्टर में दर्ज नहीं। बुक में प्रविष्टि करें।',
    btnLabel: 'Open Missing in Books',
    btnColor: 'red',
    targetTab: 'not-in-book'
  },
  {
    id: 'reco-missing-2b',
    title: 'Not in GSTR-2B',
    descEn: 'Recorded in your books but not reflected in GSTR-2B. Hold ITC.',
    descHi: 'बुक रजिस्टर में दर्ज, परंतु GSTR-2B में अनुपलब्ध। ITC रोककर रखें।',
    btnLabel: 'Open Missing in 2B',
    btnColor: 'red',
    targetTab: 'not-in-2b'
  }
];

export default function WorkingStepsView({ steps, setView, showToast }: WorkingStepsViewProps) {
  const [activeDetailCategory, setActiveDetailCategory] = useState<WorkingDetailCategory | null>(null);

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

  const handleOpenCategory = (targetTab: WorkingDetailCategory) => {
    setActiveDetailCategory(targetTab);
  };

  const getStepIcon = (index: number) => {
    switch (index) {
      case 0:
        // GSTR-1: Document with lines
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
        // ITC-Purchase: Book / Ledger
        return (
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20" />
            <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z" />
            <line x1="8" y1="7" x2="16" y2="7" />
            <line x1="8" y1="11" x2="14" y2="11" />
          </svg>
        );
      case 2:
        // RCM: Bidirectional exchange arrows ⇄
        return (
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
            <path d="M7 10h14l-4-4" />
            <path d="M17 14H3l4 4" />
          </svg>
        );
      case 3:
        // GSTR-2B: Document with checkmark
        return (
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <polyline points="14 2 14 8 20 8" />
            <polyline points="9 15 11 17 15 12" />
          </svg>
        );
      case 4:
      default:
        // Reconciliation: Double checkmark
        return (
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="20 6 9 17 4 12" />
            <polyline points="22 10 13 19 11 17" />
          </svg>
        );
    }
  };

  const getExtraCardIcon = (id: string) => {
    if (id === 'reco-matched') {
      return (
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
          <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
          <polyline points="22 4 12 14.01 9 11.01" />
        </svg>
      );
    }
    if (id === 'reco-missing-books') {
      return (
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
          <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20" />
          <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z" />
          <line x1="10" y1="10" x2="14" y2="14" />
          <line x1="14" y1="10" x2="10" y2="14" />
        </svg>
      );
    }
    return (
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
        <polyline points="14 2 14 8 20 8" />
        <line x1="9" y1="14" x2="15" y2="14" />
      </svg>
    );
  };

  if (activeDetailCategory) {
    return (
      <WorkingStepsDetailSheet
        category={activeDetailCategory}
        onBack={() => setActiveDetailCategory(null)}
        showToast={showToast}
      />
    );
  }

  return (
    <>
      <RegisterPeriodBar view="steps" />
      <section id="v-steps" className="steps-container">
      <div className="steps-cards-grid">
        {/* Steps 1 to 5 */}
        {(steps || []).map((s, index) => {
          const isRecon = index === 4 || s.titleEn.toLowerCase().includes('reconciliation');

          return (
            <div key={s.id} className={`step-card ${s.done ? 'step-done' : ''}`}>
              {/* Top Header: Left Icon Box, Right Step Number */}
              <div className="step-card-header">
                <div className="step-card-icon-box" title={s.titleEn}>
                  {getStepIcon(index)}
                </div>
                <span className="step-card-badge">Step {index + 1}</span>
              </div>

              {/* Middle Body: Title, English Detail, Hindi Detail */}
              <div className="step-card-body">
                <h3 className="step-card-title">{s.titleEn}</h3>
                <p className="step-card-desc-en">{s.descEn}</p>
                <p className="step-card-desc-hi">{s.descHi}</p>

                {/* Notes if any */}
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

              {/* Bottom Footer: Workplace Link & Dark Purple Glow Action Button */}
              <div className="step-card-footer">
                <span
                  className="step-workplace-link"
                  onClick={() => handleOpenModule(index, s.titleEn)}
                  role="button"
                  tabIndex={0}
                >
                  Open your workplace →
                </span>
                <button
                  type="button"
                  className={`step-action-btn ${isRecon ? 'btn-reconcile' : 'btn-add-data'}`}
                  onClick={() => handleOpenModule(index, s.titleEn)}
                  style={{
                    background: 'linear-gradient(135deg, #4c1d95 0%, #3b0764 100%)',
                    color: '#ffffff',
                    boxShadow: '0 4px 14px rgba(76, 29, 149, 0.45)',
                    border: 'none',
                    fontWeight: 600,
                    cursor: 'pointer'
                  }}
                >
                  {isRecon ? 'Start Reconciliation' : 'Add GST Data'}
                </button>
              </div>
            </div>
          );
        })}

        {/* 3 Extra Reconciliation Status Cards (Without "Step" badge) */}
        {EXTRA_RECO_CARDS.map((card) => {
          const isGreen = card.btnColor === 'green';

          return (
            <div
              key={card.id}
              className={`step-card ${isGreen ? 'step-card-green' : 'step-card-red'}`}
            >
              {/* Top Header: Icon only, NO Step text */}
              <div className="step-card-header">
                <div
                  className={`step-card-icon-box ${isGreen ? 'icon-green' : 'icon-red'}`}
                  title={card.title}
                >
                  {getExtraCardIcon(card.id)}
                </div>
              </div>

              {/* Middle Body: Title, English Detail, Hindi Detail */}
              <div className="step-card-body">
                <h3 className="step-card-title">{card.title}</h3>
                <p className="step-card-desc-en">{card.descEn}</p>
                <p className="step-card-desc-hi">{card.descHi}</p>
              </div>

              {/* Bottom Footer: Workplace Link & Colored Action Button */}
              <div className="step-card-footer">
                <span
                  className="step-workplace-link"
                  onClick={() => handleOpenCategory(card.targetTab)}
                  role="button"
                  tabIndex={0}
                >
                  Open your workplace →
                </span>
                <button
                  type="button"
                  className={`step-action-btn ${isGreen ? 'btn-green' : 'btn-red'}`}
                  onClick={() => handleOpenCategory(card.targetTab)}
                  style={
                    isGreen
                      ? {
                          background: 'linear-gradient(135deg, #15803d 0%, #16a34a 100%)',
                          color: '#ffffff',
                          boxShadow: '0 4px 14px rgba(22, 163, 74, 0.40)',
                          border: 'none',
                          fontWeight: 600,
                          cursor: 'pointer'
                        }
                      : {
                          background: 'linear-gradient(135deg, #b91c1c 0%, #dc2626 100%)',
                          color: '#ffffff',
                          boxShadow: '0 4px 14px rgba(220, 38, 38, 0.40)',
                          border: 'none',
                          fontWeight: 600,
                          cursor: 'pointer'
                        }
                  }
                >
                  {card.btnLabel}
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </section>
  </>
  );
}
