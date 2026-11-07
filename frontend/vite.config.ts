import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Production build notes:
// - No source maps are shipped (nothing about the source is exposed to the browser).
// - assetsInlineLimit 0 keeps every asset a real file, so the CSP never needs data: for
//   scripts, fonts or workers.
// - Each editor is a separate dynamic import, so the hub loads with almost nothing.
// - base './' makes every asset URL relative, so the same build works at the domain root
//   or under any subpath (domain.com/something/) without rebuilding.
export default defineConfig({
  plugins: [react()],
  base: './',
  build: {
    target: 'es2022',
    outDir: 'dist',
    sourcemap: false,
    assetsInlineLimit: 0,
    chunkSizeWarningLimit: 8000,
    cssCodeSplit: true,
    reportCompressedSize: false,
  },
  worker: {
    format: 'es',
  },
  server: {
    host: '127.0.0.1',
    port: 5173,
  },
  preview: {
    host: '127.0.0.1',
    port: 4173,
  },
});
