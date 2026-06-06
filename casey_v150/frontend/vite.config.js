import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    host: '0.0.0.0',
  },
  optimizeDeps: {
    force: true,
  },
  cacheDir: 'node_modules/.vite_casey',
})
