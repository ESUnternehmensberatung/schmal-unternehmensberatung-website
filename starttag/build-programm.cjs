// Erzeugt das einseitige Programm zum KI-Starttag als PDF (A4).
// Aufruf: NPM_ROOT=$(npm root -g) node starttag/build-programm.cjs
const path = require('path');
const { chromium } = require((process.env.NPM_ROOT ? process.env.NPM_ROOT + '/' : '') + 'playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 794, height: 1123 }, deviceScaleFactor: 2 });
  await p.goto('file://' + path.join(__dirname, 'programm.html'), { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
  await p.pdf({ path: path.join(__dirname, 'KI-Starttag-Programm.pdf'), preferCSSPageSize: true, printBackground: true });
  await p.screenshot({ path: path.join(__dirname, 'KI-Starttag-Programm-Vorschau.png'), fullPage: true });
  await b.close();
})();
