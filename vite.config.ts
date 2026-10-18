import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), VitePWA({
    registerType: 'prompt',
    injectRegister: 'auto',
    manifest: { name: 'AKXR Editor', short_name: 'AKXR', start_url: '.', display: 'standalone', theme_color: '#171224', background_color: '#171224', icons: [] },
    workbox: { globPatterns: ['**/*.{js,mjs,css,html,png,svg,woff2}'], maximumFileSizeToCacheInBytes: 6 * 1024 * 1024, navigateFallback: 'index.html', cleanupOutdatedCaches: true },
  })],
})
