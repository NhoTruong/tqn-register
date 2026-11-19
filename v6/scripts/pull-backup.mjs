// Sao lưu nguyên trạng code đang chạy (PROD) về thư mục backup/<ngày>, kèm SHA-256 – chạy TRƯỚC mọi lần push.
import { execSync as x } from 'node:child_process'; import { mkdirSync, writeFileSync, readdirSync, readFileSync } from 'node:fs'; import { createHash } from 'node:crypto';
const env = process.argv[2] || 'prod', id = JSON.parse(readFileSync(`.clasp.${env}.json`, 'utf8')).scriptId;
const dir = `backup/${env}-${new Date().toISOString().replace(/[:.]/g, '-')}`; mkdirSync(dir, { recursive: true });
writeFileSync(`${dir}/.clasp.json`, JSON.stringify({ scriptId: id, rootDir: '.' }));
x('npx clasp pull', { cwd: dir, stdio: 'inherit' });
const sums = readdirSync(dir).filter(f => !f.startsWith('.')).map(f => createHash('sha256').update(readFileSync(`${dir}/${f}`)).digest('hex') + '  ' + f).join('\n');
writeFileSync(`${dir}/SHA256SUMS`, sums + '\n'); console.log(sums); console.log('✔ Sao lưu tại', dir);
