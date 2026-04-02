import { defineConfig } from 'vite'
import * as path from 'path'

// https://vite.dev/config/
export default defineConfig({
  base: "./",
  server: {
    host: true,
    port: 3000,
  },
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
  },
});
