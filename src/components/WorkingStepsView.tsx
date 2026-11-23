import { StepItem, ViewType } from '../types/gst';

interface WorkingStepsViewProps {
  steps: StepItem[];
  toggleStepDone?: (id: number) => void;
  setView: (view: ViewType) => void;
  showToast: (msg: string) => void;
}

export default function WorkingStepsView({ steps, setView, showToast }: WorkingStepsViewProps) {
  const handleAddData = (index: number, stepTitle: string) => {
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

  return (
    <section id="v-steps">
      <div className="card">
        {(steps || []).map((s, index) => (
          <div key={s.id} className={`step ${s.done ? 'done' : ''}`}>
            <div className="num">{s.done ? '✓' : index + 1}</div>

            <div style={{ flex: 1, minWidth: 0, maxWidth: '760px' }}>
              <b>{s.titleEn}</b>
              <br />
              <small style={{ color: 'var(--mute)' }}>{s.descEn}</small>
              <br />
              <small style={{ color: 'var(--ink)' }}>{s.descHi}</small>

              {/* Notes list if any */}
              {s.notes && s.notes.length > 0 && s.notes.map((note, nIdx) => (
                <div key={nIdx} className="nr">
                  <span>• {note}</span>
                </div>
              ))}
            </div>

            <div className="act">
              <span
                className="workplace-text"
                onClick={() => handleAddData(index, s.titleEn)}
                role="button"
                tabIndex={0}
              >
                Open your workplace ➔
              </span>
              <button
                type="button"
                className="add"
                onClick={() => handleAddData(index, s.titleEn)}
              >
                {index === 4 || s.titleEn.toLowerCase().includes('reconciliation')
                  ? 'Start Reconciliation'
                  : 'Add GST Data'}
              </button>
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}
