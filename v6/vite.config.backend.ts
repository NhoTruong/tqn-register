import { defineConfig } from 'vite';
import gas from 'rollup-plugin-google-apps-script';
import { execSync } from 'node:child_process';
// Đóng dấu phiên bản vào code chạy: commit + thời điểm build + môi trường → trả lời "Web App đang chạy bản nào".
const sh = (c: string) => { try { return execSync(c).toString().trim(); } catch { return 'nogit'; } };
const BUILD = { commit: sh('git rev-parse --short HEAD') + (sh('git status --porcelain') ? '-dirty' : ''), builtAt: new Date().toISOString(), env: process.env.TQN_ENV || 'DEV', version: process.env.npm_package_version || '0' };
export default defineConfig({
  plugins: [gas()],
  define: { __BUILD__: JSON.stringify(BUILD) },
  build: { rollupOptions: { input: 'src/backend/main.ts', output: { dir: 'gas/dist', entryFileNames: 'main.js' } }, emptyOutDir: false, minify: false },
});
