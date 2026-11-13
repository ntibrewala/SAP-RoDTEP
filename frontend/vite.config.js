import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

export default defineConfig({
  base: "./",
  plugins: [react(), tailwindcss()],
  server: {
    host: "0.0.0.0",
    port: 5178,
    proxy: {
      "/api": {
        target: "http://127.0.0.1:5065",
        changeOrigin: true
      },
      "/rod/api": {
        target: "http://127.0.0.1:5065",
        changeOrigin: true
      }
    }
  }
});
