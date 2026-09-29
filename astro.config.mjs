import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

// BASE_PATH setzt der Deploy-Workflow: leer bei eigener Domain,
// "/schmal-unternehmensberatung-website" auf github.io ohne Domain.
export default defineConfig({
  site: process.env.SITE_URL || 'https://schmal-unternehmensberatung.de',
  base: process.env.BASE_PATH || '/',
  trailingSlash: 'always',
  integrations: [sitemap()],
});
