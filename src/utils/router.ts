import { ViewType } from '../types/gst';

export const ROUTE_PATH_MAP: Record<ViewType, string> = {
  itc: '/dashboard',
  liab: '/liability',
  steps: '/working-steps',
  g1: '/gstr1',
  'itc-reg': '/purchase',
  rcm: '/rcm',
  g2b: '/gstr2b',
  rec: '/reconciliation',
  set: '/settings',
  'reco-matched': '/reco-matched',
  'reco-missing-books': '/missing-in-books',
  'reco-missing-2b': '/missing-in-2b'
};

export const PATH_VIEW_MAP: Record<string, ViewType> = {
  '/dashboard': 'itc',
  '/liability': 'liab',
  '/working-steps': 'steps',
  '/gstr1': 'g1',
  '/purchase': 'itc-reg',
  '/rcm': 'rcm',
  '/g2b': 'g2b',
  '/reconciliation': 'rec',
  '/settings': 'set',
  '/reco-matched': 'reco-matched',
  '/missing-in-books': 'reco-missing-books',
  '/missing-in-2b': 'reco-missing-2b',
  // Direct aliases for convenience
  '/itc': 'itc',
  '/liab': 'liab',
  '/steps': 'steps',
  '/g1': 'g1',
  '/itc-reg': 'itc-reg',
  '/gstr-1': 'g1',
  '/gstr2b': 'g2b',
  '/gstr-2b': 'g2b',
  '/rec': 'rec',
  '/set': 'set',
  '/matched': 'reco-matched',
  '/not-in-books': 'reco-missing-books',
  '/not-in-2b': 'reco-missing-2b'
};

/**
 * Returns ViewType matching current pathname, or null if unknown
 */
export function getViewFromPath(pathname: string): ViewType | null {
  if (!pathname) return null;
  const clean = pathname.replace(/\/+$/, '') || '/';
  if (clean === '/') return 'itc';
  return PATH_VIEW_MAP[clean.toLowerCase()] || null;
}

/**
 * Returns canonical URL path for a ViewType
 */
export function getPathForView(view: ViewType): string {
  return ROUTE_PATH_MAP[view] || '/dashboard';
}

/**
 * Reads query parameters from window.location.search
 */
export function getQueryParams(): { fy?: string; month?: string; tab?: string } {
  if (typeof window === 'undefined') return {};
  try {
    const params = new URLSearchParams(window.location.search);
    const fy = params.get('fy')?.trim() || undefined;
    const month = params.get('month')?.trim() || undefined;
    const tab = params.get('tab')?.trim() || undefined;
    return { fy, month, tab };
  } catch {
    return {};
  }
}

/**
 * Builds URL from view and optional fy/month query params
 */
export function buildUrl(view: ViewType, fy?: string, month?: string): string {
  const path = getPathForView(view);
  const params = new URLSearchParams();
  if (fy && fy !== 'ALL') {
    params.set('fy', fy);
  }
  if (month && month !== 'ALL') {
    params.set('month', month);
  }
  if (typeof window !== 'undefined') {
    const existingTab = new URLSearchParams(window.location.search).get('tab');
    if (existingTab && view === 'rec') {
      params.set('tab', existingTab);
    }
  }
  const queryString = params.toString();
  return queryString ? `${path}?${queryString}` : path;
}

/**
 * Programmatically navigate to a view, updating window.history without full page reload
 */
export function navigateTo(view: ViewType, fy?: string, month?: string, replace = false) {
  if (typeof window === 'undefined') return;
  const targetUrl = buildUrl(view, fy, month);
  const currentUrl = window.location.pathname + window.location.search;

  if (currentUrl !== targetUrl) {
    if (replace) {
      window.history.replaceState({ view, fy, month }, '', targetUrl);
    } else {
      window.history.pushState({ view, fy, month }, '', targetUrl);
    }
  }
}

/**
 * Update query params (fy, month) on current route without changing pathname
 */
export function updateQueryParams(params: { fy?: string; month?: string; tab?: string }) {
  if (typeof window === 'undefined') return;
  try {
    const searchParams = new URLSearchParams(window.location.search);

    if (params.fy !== undefined) {
      if (params.fy && params.fy !== 'ALL') {
        searchParams.set('fy', params.fy);
      } else {
        searchParams.delete('fy');
      }
    }

    if (params.month !== undefined) {
      if (params.month && params.month !== 'ALL') {
        searchParams.set('month', params.month);
      } else {
        searchParams.delete('month');
      }
    }

    if (params.tab !== undefined) {
      if (params.tab) {
        searchParams.set('tab', params.tab);
      } else {
        searchParams.delete('tab');
      }
    }

    const queryString = searchParams.toString();
    const newUrl = window.location.pathname + (queryString ? `?${queryString}` : '');
    const currentUrl = window.location.pathname + window.location.search;

    if (newUrl !== currentUrl) {
      window.history.replaceState(window.history.state, '', newUrl);
    }
  } catch {}
}
