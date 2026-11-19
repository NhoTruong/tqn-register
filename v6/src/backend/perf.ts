// Đo từng chặng trong 1 request Apps Script: đọc Sheets, ghi Sheets, khóa, ghi sổ, Drive, cache. Trả về cho web hiển thị.
export const P = {
  ms: {} as Record<string, number>, n: {} as Record<string, number>, t0: 0,
  reset() { this.ms = {}; this.n = {}; this.t0 = Date.now(); },
  time<T>(k: string, fn: () => T): T { const a = Date.now(); try { return fn(); } finally { this.ms[k] = (this.ms[k] || 0) + Date.now() - a; this.n[k] = (this.n[k] || 0) + 1; } },
  out() { return { total: Date.now() - this.t0, ...this.ms, calls: { ...this.n } }; },
};
