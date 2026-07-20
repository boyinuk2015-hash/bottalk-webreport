import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Deployed to GitHub Pages at https://boyinuk2015-hash.github.io/bottalk-webreport/
// so the base path must match the repository name.
export default defineConfig({
  base: "/bottalk-webreport/",
  plugins: [react()],
});
