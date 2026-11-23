import { ViewType } from '../types/gst';

interface TopBarProps {
  view: ViewType;
  setView: (view: ViewType) => void;
  fy?: string;
  setFy?: (fy: string) => void;
  onLogout: () => void;
  onToggleSidebar?: () => void;
  isSidebarCollapsed?: boolean;
  showToast?: (msg: string) => void;
}

export default function TopBar({
  view,
  setView,
  onLogout,
  onToggleSidebar,
  isSidebarCollapsed
}: TopBarProps) {
  return (
    <div className="top">
      <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
        {onToggleSidebar && (
          <button
            type="button"
            className="clean-hamburger-btn"
            onClick={onToggleSidebar}
            title={isSidebarCollapsed ? 'Expand Sidebar' : 'Collapse Sidebar'}
            aria-label="Toggle Sidebar"
          >
            <svg
              width="24"
              height="24"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
              strokeLinecap="round"
            >
              <line x1="3" y1="6" x2="21" y2="6"></line>
              <line x1="3" y1="12" x2="21" y2="12"></line>
              <line x1="3" y1="18" x2="21" y2="18"></line>
            </svg>
          </button>
        )}
        <h2>GST ITC Liability Dashboard</h2>
      </div>
      <div className="right">
        {/* Inner Register Pages: Show colorful, compact '← Back to Steps' */}
        {['g1', 'itc-reg', 'rcm', 'g2b', 'rec'].includes(view) && (
          <>
            <button
              type="button"
              className="btn-back-steps"
              onClick={() => setView('steps')}
              title="Back to GST Working Steps"
            >
              <svg
                width="15"
                height="15"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <line x1="19" y1="12" x2="5" y2="12"></line>
                <polyline points="12 19 5 12 12 5"></polyline>
              </svg>
              <span>Back to Steps</span>
            </button>
            <div className="topbar-divider" />
          </>
        )}

        <div className="topbar-user-profile">
          <div className="topbar-user-avatar">
            <svg
              width="17"
              height="17"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
              <circle cx="12" cy="7" r="4" />
            </svg>
          </div>
          <span className="topbar-user-name">User</span>
        </div>

        <button
          type="button"
          className="topbar-logout-btn"
          onClick={onLogout}
          title="Logout"
          aria-label="Logout"
        >
          <svg
            width="17"
            height="17"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2.2"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
            <polyline points="16 17 21 12 16 7" />
            <line x1="21" y1="12" x2="9" y2="12" />
          </svg>
        </button>
      </div>
    </div>
  );
}
