import { ViewType } from '../types/gst';

interface RegisterPeriodBarProps {
  view: ViewType;
}

const TITLES: Record<string, [string, string]> = {
  itc: ['GST ITC', 'Input tax credit available, reconcile aur use kaise hua'],
  liab: ['Liability', 'Output tax, ITC adjustment aur cash payable'],
  steps: [
    'GST Return Preparation Steps - Systematic working every figure is verified before filing FY 2026-27',
    ''
  ],
  set: ['Settings', 'Business profile, filing aur alerts ki settings'],
  rcm: ['Reverse Charge Mechanism', 'RCM liability and eligible ITC tracking'],
  rec: ['ITC Reconciliation', 'GSTR-2B vs Books reconciliation statement']
};

export default function RegisterPeriodBar({
  view
}: RegisterPeriodBarProps) {
  if (['g1', 'itc-reg', 'rcm', 'g2b', 'rec', 'reco-matched', 'reco-missing-books', 'reco-missing-2b', 'old-itc', 'not-claim-itc'].includes(view)) return null;

  const titleInfo = TITLES[view] || TITLES.itc;

  return (
    <div className={`banner ${view === 'steps' ? 'banner-steps' : ''}`}>
      <div style={{ width: '100%', overflow: 'hidden' }}>
        <h1
          style={
            view === 'steps'
              ? {
                  fontSize: 'clamp(16px, 1.6vw, 20px)',
                  fontWeight: 700,
                  whiteSpace: 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  margin: 0,
                  color: '#ffffff',
                  letterSpacing: '0.015em',
                  lineHeight: '1.2'
                }
              : {}
          }
          title={titleInfo[0]}
        >
          {titleInfo[0]}
        </h1>
        {titleInfo[1] ? <p>{titleInfo[1]}</p> : null}
      </div>
    </div>
  );
}
