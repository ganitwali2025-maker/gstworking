import { useState, useEffect } from 'react';
import { useGstStore } from './hooks/useGstStore';
import { initAutoSyncRetryListener } from './services/registerService';
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

        <div className="scroll">
          <RegisterPeriodBar view={view} />

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

          {!['g1', 'itc-reg', 'rcm', 'g2b', 'rec'].includes(view) && (
            <p className="note">Yeh sample data hai. Real data aane par numbers yahin update honge.</p>
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
