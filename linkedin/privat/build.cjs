// Erzeugt das LinkedIn-Banner für das Privatprofil (1584 x 396) aus banner.html.
// Aufruf: NPM_ROOT=$(npm root -g) node linkedin/privat/build.cjs
const path = require('path');
const { chromium } = require((process.env.NPM_ROOT ? process.env.NPM_ROOT + '/' : '') + 'playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1584, height: 396 }, deviceScaleFactor: 2 });
  await p.goto('file://' + path.join(__dirname, 'banner.html'), { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
  await p.locator('.banner').screenshot({ path: path.join(__dirname, 'LinkedIn-Banner-Eduard-Schmal-privat.png') });
  await b.close();
})();
