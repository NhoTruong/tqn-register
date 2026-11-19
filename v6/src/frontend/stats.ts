export const sorted = (a: number[]) => [...a].sort((x, y) => x - y);
export function median(a: number[]) { if (!a.length) return 0; const s = sorted(a), m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; }
/** p95 theo cách nearest-rank */
export function p95(a: number[]) { if (!a.length) return 0; const s = sorted(a); return s[Math.min(s.length - 1, Math.ceil(0.95 * s.length) - 1)]; }
