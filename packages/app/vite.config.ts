import { fileURLToPath, URL } from 'node:url'

import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vueJsx from '@vitejs/plugin-vue-jsx'
import vueDevTools from 'vite-plugin-vue-devtools'

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    vue(),
    vueJsx(),
    vueDevTools(),
  ],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
      '@pptx-player/renderer-canvas': fileURLToPath(new URL('../renderer-canvas/src/index.ts', import.meta.url)),
      '@pptx-player/renderer-svg': fileURLToPath(new URL('../renderer-svg/src/index.ts', import.meta.url)),
      'readable-stream': fileURLToPath(new URL('./src/browserStreamShim.ts', import.meta.url)),
    },
  },
})
