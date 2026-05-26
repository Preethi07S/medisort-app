import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],

  // Explicitly expose all VITE_* env vars used by MediSort AI.
  // Vite only inlines variables that are referenced as import.meta.env.VITE_*
  // in source — listing them here makes the contract explicit and prevents
  // accidental omission during build on Vercel.
  envPrefix: 'VITE_',

  build: {
    outDir: 'dist',        // Vercel expects output in /dist by default
    sourcemap: false,      // disable in production — enable locally if needed
    minify: 'esbuild',     // fastest minifier, ships with Vite
    target: 'es2020',      // modern browsers; supports BigInt, optional chaining, etc.
    chunkSizeWarningLimit: 1000, // MediSort bundles xlsx + lucide; raise the warning threshold
  },

  server: {
    port: 5173,            // local dev port
    open: true,            // auto-open browser on `npm run dev`
    cors: true,            // allow cross-origin requests in dev (Gemini / Google APIs)
  },

  preview: {
    port: 4173,            // `vite preview` port (used to test the production build locally)
  },
})