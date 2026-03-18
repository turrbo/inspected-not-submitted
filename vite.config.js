import { defineConfig } from 'vite'

export default defineConfig({
  base: '/inspected-not-submitted/',
  esbuild: {
    jsxInject: `import React from 'react'`,
  },
  server: {
    host: '0.0.0.0',
    port: 5199,
    allowedHosts: true,
  },
})
