// Erzeugt die Visitenkarte als Druck-PDF (2 Seiten, 91 × 61 mm inkl. 3 mm Beschnitt)
// und Vorschaubilder im Endformat 85 × 55 mm.
// Aufruf: NPM_ROOT=$(npm root -g) node visitenkarte/build.cjs
const path = require('path');
const { chromium } = require((process.env.NPM_ROOT ? process.env.NPM_ROOT + '/' : '') + 'playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 344, height: 231 }, deviceScaleFactor: 4 });
  await p.goto('file://' + path.join(__dirname, 'visitenkarte.html'), { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
  await p.pdf({ path: path.join(__dirname, 'Visitenkarte-Eduard-Schmal-Druck.pdf'), preferCSSPageSize: true, printBackground: true });
  // Vorschau: auf Endformat beschnitten (3 mm = 11,34 px bei 96 dpi)
  const mm = 96 / 25.4;
  for (const [cls, name] of [['vorne', 'Vorderseite'], ['hinten', 'Rueckseite']]) {
    const box = await p.locator('.' + cls).boundingBox();
    await p.screenshot({
      path: path.join(__dirname, `Visitenkarte-${name}-Vorschau.png`),
      fullPage: true,
      clip: { x: box.x + 3 * mm, y: box.y + 3 * mm, width: 85 * mm, height: 55 * mm },
    });
  }
  await b.close();
})();
