import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

export default defineConfig({
  site: 'https://schmal-unternehmensberatung.de',
  trailingSlash: 'always',
  integrations: [sitemap()],
});
