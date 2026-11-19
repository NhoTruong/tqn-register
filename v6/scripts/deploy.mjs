// Triển khai có kiểm soát: node scripts/deploy.mjs <dev|staging|prod>
// - prod: bắt buộc git sạch, đang ở nhánh main, đã có tag v*, đã chạy test; ghi RELEASES.md (commit ↔ version ↔ deployment)
import { execSync as x } from 'node:child_process';
import { existsSync, copyFileSync, appendFileSync, readFileSync } from 'node:fs';
const env = process.argv[2]; const sh = c => x(c, { stdio: ['ignore', 'pipe', 'inherit'] }).toString().trim();
if (!['dev', 'staging', 'prod'].includes(env)) throw new Error('Dùng: node scripts/deploy.mjs dev|staging|prod');
const cfg = `.clasp.${env}.json`; if (!existsSync(cfg)) throw new Error(`Thiếu ${cfg} (copy từ ${cfg}.example, điền scriptId)`);
const commit = sh('git rev-parse --short HEAD'), dirty = sh('git status --porcelain');
if (env !== 'dev' && dirty) throw new Error('Còn thay đổi chưa commit – không deploy ' + env);
if (env === 'prod') {
  if (sh('git rev-parse --abbrev-ref HEAD') !== 'main') throw new Error('PROD chỉ deploy từ nhánh main');
  if (!sh(`git tag --points-at HEAD`).startsWith('v')) throw new Error('PROD cần tag phiên bản (vd v0.6.0) trên commit này');
  if (process.env.TQN_CONFIRM_PROD !== commit) throw new Error(`Đặt TQN_CONFIRM_PROD=${commit} để xác nhận deploy PROD`);
}
x('npm test', { stdio: 'inherit' });
x(`npm run build`, { stdio: 'inherit', env: { ...process.env, TQN_ENV: env.toUpperCase() } });
copyFileSync(cfg, '.clasp.json');
x('npx clasp push -f', { stdio: 'inherit' });
const ver = sh(`npx clasp create-version "${env} ${commit}"`).match(/(\d+)/)?.[1] || '?';
const depId = (JSON.parse(readFileSync(cfg, 'utf8')).deploymentId) || '';
if (depId) x(`npx clasp create-deployment -i ${depId} -V ${ver} -d "${env} ${commit}"`, { stdio: 'inherit' });
appendFileSync('RELEASES.md', `| ${new Date().toISOString()} | ${env.toUpperCase()} | ${commit} | v${ver} | ${depId || '(chưa có web deployment)'} |\n`);
console.log(`✔ ${env}: commit ${commit} → Apps Script version ${ver}`);
