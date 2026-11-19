// Quay lại phiên bản cũ KHÔNG cần sửa code: trỏ deployment về version trước. node scripts/rollback.mjs <env> <version>
import { execSync as x } from 'node:child_process'; import { readFileSync, copyFileSync, appendFileSync } from 'node:fs';
const [env, ver] = process.argv.slice(2); const cfg = `.clasp.${env}.json`; const dep = JSON.parse(readFileSync(cfg, 'utf8')).deploymentId;
if (!dep || !ver) throw new Error('Cần deploymentId trong ' + cfg + ' và số version');
copyFileSync(cfg, '.clasp.json');
x(`npx clasp create-deployment -i ${dep} -V ${ver} -d "ROLLBACK to v${ver}"`, { stdio: 'inherit' });
appendFileSync('RELEASES.md', `| ${new Date().toISOString()} | ${env.toUpperCase()} | ROLLBACK | v${ver} | ${dep} |\n`);
