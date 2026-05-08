import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig(({ mode }) => ({
  plugins: [react()],
  /** Electron loads dist/index.html from disk — needs relative asset paths */
  base: mode === 'electron' ? './' : '/',
}))
