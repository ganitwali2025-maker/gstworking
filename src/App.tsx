import { useState, useEffect, useMemo } from 'react';
import { useGstStore } from './hooks/useGstStore';
import { initAutoSyncRetryListener } from './services/registerService';
import { parseWorkingMonth } from './services/lockService';
import Sidebar from './components/Sidebar';
import TopBar from './components/TopBar';
import RegisterPeriodBar from './components/RegisterPeriodBar';
import ItcDashboardView from './components/ItcDashboardView';
import LiabilityDashboardView from './components/LiabilityDashboardView';
import WorkingStepsView from './components/WorkingStepsView';
import TaxRegisterView from './components/TaxRegisterView';
import SettingsView from './components/SettingsView';
import RcmRegisterView from './components/RcmRegisterView';
import ReconciliationView from './components/ReconciliationView';
import ItcRegisterView from './components/ItcRegisterView';
import Gstr2bRegisterView from './components/Gstr2bRegisterView';
import WorkingStepsDetailSheet from './components/WorkingStepsDetailSheet';
import OldItcDetailSheet from './components/OldItcDetailSheet';
import NotClaimItcDetailSheet from './components/NotClaimItcDetailSheet';

export default function App() {
  const {
    view,
    setView,
    fy,
    setFy,
    month,
    setMonth,
    monthlyData,
    invoices,
    isLocked,
    selectedInvoiceId,
    setSelectedInvoiceId,
    steps,
    settings,
    toastMessage,
    showToast,
    isLoggedOut,
    setIsLoggedOut,
    addInvoice,
    updateInvoice,
    deleteInvoice,
    clearInvoices,
    toggleLock,
    toggleStepDone,
    saveSettings,
    refreshInvoices,
    undoInvoices,
    redoInvoices,
    batchImportInvoices
  } = useGstStore();

  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(false);
  const toggleSidebarCollapse = () => setIsSidebarCollapsed((prev) => !prev);

  const workingMonth = useMemo(() => {
    return parseWorkingMonth(month, fy);
  }, [fy, month]);

  useEffect(() => {
    initAutoSyncRetryListener();
  }, []);

  return (
    <div className={`app ${isSidebarCollapsed ? 'sidebar-collapsed' : ''}`}>
      <Sidebar
        view={view}
        setView={setView}
        isCollapsed={isSidebarCollapsed}
      />

      <main className="main">
        <TopBar
          view={view}
          setView={setView}
          fy={fy}
          setFy={setFy}
          month={month}
          setMonth={setMonth}
          onLogout={() => setIsLoggedOut(true)}
          onToggleSidebar={toggleSidebarCollapse}
          isSidebarCollapsed={isSidebarCollapsed}
          showToast={showToast}
        />

        <div className={`scroll ${view === 'rec' ? 'scroll-rec' : ''} ${view === 'steps' ? 'scroll-steps' : ''}`}>
          {view !== 'steps' && <RegisterPeriodBar view={view} />}

          {view === 'itc' && (
            <ItcDashboardView
              month={month}
              setMonth={setMonth}
              monthlyData={monthlyData}
            />
          )}

          {view === 'liab' && (
            <LiabilityDashboardView
              month={month}
              monthlyData={monthlyData}
            />
          )}

          {view === 'steps' && (
            <WorkingStepsView
              steps={steps}
              toggleStepDone={toggleStepDone}
              setView={setView}
              showToast={showToast}
              fy={fy}
              month={month}
            />
          )}

          {view === 'g1' && (
            <TaxRegisterView
              invoices={invoices}
              isLocked={isLocked}
              onAddInvoice={addInvoice}
              onUpdateInvoice={updateInvoice}
              onDeleteInvoice={deleteInvoice}
              onClearInvoices={clearInvoices}
              onToggleLock={toggleLock}
              onRefresh={refreshInvoices}
              onUndo={undoInvoices}
              onRedo={redoInvoices}
              onBatchImport={batchImportInvoices}
              showToast={showToast}
              selectedInvoiceId={selectedInvoiceId}
              setSelectedInvoiceId={setSelectedInvoiceId}
              fy={fy}
              setFy={setFy}
              month={month}
              setMonth={setMonth}
            />
          )}

          {view === 'set' && (
            <SettingsView
              settings={settings}
              onSave={saveSettings}
            />
          )}

          {view === 'itc-reg' && <ItcRegisterView showToast={showToast} fy={fy} setFy={setFy} month={month} setMonth={setMonth} />}
          {view === 'rcm' && <RcmRegisterView showToast={showToast} fy={fy} setFy={setFy} month={month} setMonth={setMonth} />}
          {view === 'g2b' && <Gstr2bRegisterView showToast={showToast} fy={fy} setFy={setFy} month={month} setMonth={setMonth} />}
          {view === 'rec' && <ReconciliationView showToast={showToast} fy={fy} month={month} monthlyData={monthlyData} />}

          {view === 'reco-matched' && (
            <WorkingStepsDetailSheet
              category="matched"
              onBack={() => setView('steps')}
              showToast={showToast}
              workingMonth={workingMonth}
            />
          )}
          {view === 'reco-missing-books' && (
            <WorkingStepsDetailSheet
              category="not-in-book"
              onBack={() => setView('steps')}
              showToast={showToast}
              workingMonth={workingMonth}
            />
          )}
          {view === 'reco-missing-2b' && (
            <WorkingStepsDetailSheet
              category="not-in-2b"
              onBack={() => setView('steps')}
              showToast={showToast}
              workingMonth={workingMonth}
            />
          )}

          {view === 'old-itc' && (
            <OldItcDetailSheet
              onBack={() => setView('steps')}
              showToast={showToast}
              workingMonth={workingMonth}
            />
          )}

          {view === 'not-claim-itc' && (
            <NotClaimItcDetailSheet
              onBack={() => setView('steps')}
              showToast={showToast}
              workingMonth={workingMonth}
            />
          )}

          {!['g1', 'itc-reg', 'rcm', 'g2b', 'rec', 'reco-matched', 'reco-missing-books', 'reco-missing-2b', 'old-itc', 'not-claim-itc'].includes(view) && (
            <p
              className="note"
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '8px',
                fontSize: '12.5px',
                color: '#64748b',
                letterSpacing: '0.02em',
                margin: '18px 0 8px',
                borderTop: '1px solid #e2e8f0',
                paddingTop: '10px'
              }}
            >
              <span>GST Intelligence Suite&trade; | Designed &amp; Developed by <strong style={{ color: '#350B5E' }}>Lokesh Rajak</strong>, Accountant</span>
            </p>
          )}
        </div>
      </main>

      {/* Global Toast Alert */}
      {toastMessage && <div className="gt">{toastMessage}</div>}

      {/* Logout Dialog */}
      {isLoggedOut && (
        <div className="login">
          <div className="card">
            <h3>Aap logout ho gaye</h3>
            <p style={{ color: 'var(--mute)', margin: '10px 0 16px' }}>
              Dobara dashboard dekhne ke liye login karein.
            </p>
            <button
              type="button"
              className="btn"
              onClick={() => setIsLoggedOut(false)}
            >
              Login again
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
