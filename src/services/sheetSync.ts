export interface SheetSyncPayload {
  action: 'upsert' | 'delete' | 'clear' | 'lock' | 'unlock';
  sheet: string;
  headers: string[];
  rows?: { id: string; values: (string | number)[] }[];
  ids?: string[];
}

export async function syncModuleToSheets(
  payload: SheetSyncPayload
): Promise<{ success: boolean; count?: number; error?: string | null }> {
  try {
    const res = await fetch('/api/sheets/sync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (!res.ok) {
      return { success: false, error: `HTTP ${res.status}` };
    }

    const data = await res.json();
    return data;
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Network error';
    return { success: false, error: msg };
  }
}
