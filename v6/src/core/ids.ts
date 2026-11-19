// ULID (26 ký tự Crockford base32): sắp theo thời gian, không trùng, không phụ thuộc số dòng.
const B32 = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';

export type Rand = () => number;

export function ulid(nowMs: number, rand: Rand = Math.random): string {
  let t = nowMs, time = '';
  for (let i = 0; i < 10; i++) { time = B32[t % 32] + time; t = Math.floor(t / 32); }
  let r = '';
  for (let i = 0; i < 16; i++) r += B32[Math.floor(rand() * 32) % 32];
  return time + r;
}

export function isUlid(s: string): boolean { return /^[0-9A-HJKMNP-TV-Z]{26}$/.test(s); }
