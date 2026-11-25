const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const pad = (n: number): string => (n < 10 ? '0' + n : String(n));

/** SharePoint date-only values arrive as "2026-11-25T00:00:00Z"; read the Y-M-D part only (no time-zone shifting). */
export function parseDate(s?: string | null): Date | undefined {
  if (!s) { return undefined; }
  const p = s.substr(0, 10).split('-').map(Number);
  if (p.length !== 3 || isNaN(p[0])) { return undefined; }
  return new Date(p[0], p[1] - 1, p[2]);
}

export function toDateOnly(d?: Date | null): string | null {
  return d ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` : null;
}

/** UK display format dd MMM yyyy */
export function fmt(d?: Date | string | null): string {
  if (!d) { return '—'; }
  const x = typeof d === 'string' ? new Date(d) : d;
  if (isNaN(x.getTime())) { return '—'; }
  return `${pad(x.getDate())} ${MONTHS[x.getMonth()]} ${x.getFullYear()}`;
}

export function fmtDateTime(s?: string): string {
  if (!s) { return '—'; }
  const d = new Date(s);
  return `${fmt(d)} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

export function today(): Date {
  const n = new Date();
  return new Date(n.getFullYear(), n.getMonth(), n.getDate());
}

export function daysTo(d: Date): number {
  return Math.round((d.getTime() - today().getTime()) / 86400000);
}

export function sameOrAfter(a: Date, b: Date): boolean {
  return toDateOnly(a)! >= toDateOnly(b)!;
}

export function nextCode(prefix: string, existing: string[], width: number = 4): string {
  let max = 0;
  existing.forEach(c => {
    if (c && c.toUpperCase().indexOf(prefix) === 0) {
      const n = parseInt(c.substr(prefix.length), 10);
      if (!isNaN(n) && n > max) { max = n; }
    }
  });
  let num = String(max + 1);
  while (num.length < width) { num = '0' + num; }
  return prefix + num;
}

export function money(a: number | undefined, cur: string): string {
  if (a === undefined || a === null) { return '—'; }
  return `${cur} ${a.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

export function errMsg(e: any): string {
  if (!e) { return 'Unknown error'; }
  if (typeof e === 'string') { return e; }
  if (e.status === 403 || e.status === 401) { return 'Access denied: you do not have permission to perform this action (V12).'; }
  const m: string = e.message || String(e);
  const idx = m.indexOf('{');
  if (idx >= 0) {
    try {
      const j = JSON.parse(m.substr(idx));
      const inner = j['odata.error'] && j['odata.error'].message && j['odata.error'].message.value;
      if (inner) { return inner; }
    } catch (x) { /* fall through */ }
  }
  return m;
}

export function guid(): string {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = (Math.random() * 16) | 0;
    return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16);
  });
}

export const esc = (s: string): string => s.replace(/'/g, "''");
export const isEmail = (s: string): boolean => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
