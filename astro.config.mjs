import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

// BASE_PATH setzt der Deploy-Workflow: leer bei eigener Domain,
// "/schmal-unternehmensberatung-website" auf github.io ohne Domain.
export default defineConfig({
  // Immer https: GitHub meldet die eigene Domain vor „Enforce HTTPS“ noch als http.
  site: (process.env.SITE_URL || 'https://schmal-unternehmensberatung.de').replace(/^http:/, 'https:'),
  base: process.env.BASE_PATH || '/',
  trailingSlash: 'always',
  integrations: [sitemap()],
});
