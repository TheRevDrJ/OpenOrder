import fs from 'fs'
import path from 'path'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

export default defineConfig(({ command }) => ({
  // ⛔⛔ THE BUILD NUMBER IS COMPILED IN, AND ONLY FOR A REAL BUILD. Under `vite`
  // (the dev server) it is deliberately EMPTY: a dev box has no build to number,
  // and a stale counter shown there is worse than none. The UI renders the line
  // only when this is non-empty. @decision:gold 2026-09-24
  //
  // ⚠ IT COMES FROM THE ENVIRONMENT, NOT FROM THE FILE. build.sh claims the next
  // number and exports OO_BUILD_NUMBER BEFORE calling npm run build; reading
  // build-number.txt here instead would compile in the PREVIOUS build's number,
  // because the counter on disk has not been bumped yet at that moment.
  define: {
    __APP_BUILD__: JSON.stringify(
      command === 'serve' ? '' : (process.env.OO_BUILD_NUMBER ?? '').trim(),
    ),
    __APP_VERSION__: JSON.stringify(
      JSON.parse(fs.readFileSync(path.resolve(__dirname, 'package.json'), 'utf8')).version,
    ),
  },
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
  server: {
    port: 6800,
    strictPort: true,
    proxy: {
      // ⛔ 127.0.0.1, NOT localhost. uvicorn binds 0.0.0.0, which is IPv4 ONLY, while
      // Node 17+ resolves `localhost` to ::1 first — so every /api call came back 502
      // with a healthy backend sitting right there on IPv4. Naming the address removes
      // the resolver from the path entirely.
      '/api': 'http://127.0.0.1:6801',
    },
  },
}))
