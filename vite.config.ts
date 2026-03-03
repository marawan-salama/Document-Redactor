// vite.config.ts
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "path";
import { getHttpsServerOptions } from "office-addin-dev-certs";

export default defineConfig(async () => ({
  plugins: [react()],
  root: "./",
  base: "/",
  server: {
    port: 3000,
    strictPort: true,
    host: "localhost",

    // Serve HTTPS using the trusted Office dev cert
    https: await getHttpsServerOptions(),

    // Word + HTTPS needs WSS for HMR
    hmr: {
      protocol: "wss",
      host: "localhost",
      port: 3000,
      clientPort: 3000,
      overlay: false
    },

    cors: true,
    watch: {
      usePolling: true
    }
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "index.html")
      },
      output: {
        entryFileNames: "assets/[name].js",
        chunkFileNames: "assets/[name].js",
        assetFileNames: "assets/[name].[ext]"
      }
    }
  },
  resolve: {
    alias: {
      "@": resolve(__dirname, "src")
    }
  }
}));
