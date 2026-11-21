import React from 'react';
import { ViewType } from '../types/gst';

interface SidebarProps {
  view: ViewType;
  setView: (view: ViewType) => void;
  isCollapsed: boolean;
}

export default function Sidebar({ view, setView, isCollapsed }: SidebarProps) {
  const navItems: { id: ViewType; label: string; icon: React.ReactNode }[] = [
    {
      id: 'itc',
      label: 'GST ITC',
      icon: (
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <path d="M4 19V9m6 10V5m6 14v-7m4 7H2" />
        </svg>
      )
    },
    {
      id: 'liab',
      label: 'Liability',
      icon: (
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <rect x="3" y="5" width="18" height="14" rx="2" />
          <path d="M3 10h18M7 15h4" />
        </svg>
      )
    },
    {
      id: 'steps',
      label: 'GST Working Steps',
      icon: (
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <path d="M9 6h11M9 12h11M9 18h11M4 6l1 1 2-2M4 12l1 1 2-2M4 18l1 1 2-2" />
        </svg>
      )
    },
    {
      id: 'set',
      label: 'Settings',
      icon: (
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <circle cx="12" cy="12" r="3" />
          <path d="M19 12a7 7 0 0 0-.1-1.2l2-1.5-2-3.4-2.3 1a7 7 0 0 0-2-1.2L14 3h-4l-.4 2.7a7 7 0 0 0-2 1.2l-2.3-1-2 3.4 2 1.5a7 7 0 0 0 0 2.4l-2 1.5 2 3.4 2.3-1a7 7 0 0 0 2 1.2L10 21h4l.4-2.7a7 7 0 0 0 2-1.2l2.3 1 2-3.4-2-1.5c.1-.4.1-.8.1-1.2z" />
        </svg>
      )
    }
  ];

  return (
    <aside className={isCollapsed ? 'collapsed' : ''}>
      <div className="brand" style={isCollapsed ? { justifyContent: 'center', padding: '4px 0 16px' } : {}}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <div className="logo">₹</div>
          {!isCollapsed && <span className="brand-name">GST Panel</span>}
        </div>
      </div>

      <div className="nav-items" style={{ display: 'flex', flexDirection: 'column', gap: '6px', width: '100%' }}>
        {navItems.map((item) => {
          const isActive = view === item.id || (view === 'g1' && item.id === 'steps');
          return (
            <button
              key={item.id}
              type="button"
              className={`nav ${isActive ? 'on' : ''}`}
              onClick={() => setView(item.id)}
              title={item.label}
            >
              {item.icon}
              {!isCollapsed && <span className="nav-label">{item.label}</span>}
            </button>
          );
        })}
      </div>
    </aside>
  );
}
