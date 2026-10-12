import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { fileURLToPath } from 'node:url';

/*
 * The desktop window is a React app; the Excel taskpane stays plain JS.
 *
 * Output goes to dist/app/web/ui/ alongside the compiled taskpane, because the
 * server serves dist/app/web/ statically. `/ui/` is therefore the window's URL
 * and needs no separate build server. `base` must match that path or the asset
 * URLs will 404.
 */
export default defineConfig({
  root: fileURLToPath(new URL('./ui/', import.meta.url)),
  base: '/ui/',
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: { '@': fileURLToPath(new URL('./ui/src/', import.meta.url)) },
  },
  build: {
    outDir: fileURLToPath(new URL('./dist/app/web/ui/', import.meta.url)),
    emptyOutDir: true,
    // A local app: readable output beats shaving bytes.
    minify: false,
    sourcemap: true,
  },
});
