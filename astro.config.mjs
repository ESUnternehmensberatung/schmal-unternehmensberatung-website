import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';
import { readdirSync, readFileSync } from 'node:fs';

// Blog erst in die Sitemap, wenn ein Beitrag veröffentlicht ist (kein „entwurf: true“).
const blogDir = new URL('./src/content/blog/', import.meta.url);
const hatBlog = readdirSync(blogDir)
  .filter((f) => /\.mdx?$/.test(f))
  .some((f) => !/^entwurf:\s*true/m.test(readFileSync(new URL(f, blogDir), 'utf8')));

// BASE_PATH setzt der Deploy-Workflow: leer bei eigener Domain,
// "/schmal-unternehmensberatung-website" auf github.io ohne Domain.
export default defineConfig({
  // Immer https: GitHub meldet die eigene Domain vor „Enforce HTTPS“ noch als http.
  site: (process.env.SITE_URL || 'https://schmal-unternehmensberatung.de').replace(/^http:/, 'https:'),
  base: process.env.BASE_PATH || '/',
  trailingSlash: 'always',
  // Alte WordPress-Adressen, die Google noch kennt, auf die neuen Seiten umleiten.
  redirects: {
    '/sample-page/': '/#ueber-mich',
    '/privacy-policy/': '/datenschutz/',
    '/kontakt/': '/quick-win/',
  },
  integrations: [sitemap({ filter: (page) => hatBlog || !page.includes('/blog/') })],
});
