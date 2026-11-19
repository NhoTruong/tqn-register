import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';
// Từ starter gas-webapp-boilerplate (MIT): React → 1 file HTML để HtmlService phục vụ.
export default defineConfig({ plugins: [react(), viteSingleFile()], build: { outDir: 'gas/dist', emptyOutDir: false } });
