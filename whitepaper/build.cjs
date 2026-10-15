// Erzeugt das Whitepaper-PDF aus whitepaper.html.
// Aufruf: NPM_ROOT=$(npm root -g) node whitepaper/build.cjs
// Seite 1 (Cover) wird ohne Fußzeile erzeugt; merge.py setzt beide Teile zusammen:
//   python3 whitepaper/merge.py   (benötigt pymupdf)
const path = require('path');
const fs = require('fs');
const { chromium } = require((process.env.NPM_ROOT ? process.env.NPM_ROOT + '/' : '') + 'playwright');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('file://' + path.join(__dirname, 'whitepaper.html'), { waitUntil: 'networkidle' });
  await page.evaluate(() => document.fonts.ready);
  const fuss = `<div style="width:100%;font-family:Helvetica,Arial,sans-serif;font-size:7.5pt;color:#5a6572;padding:0 21mm;display:flex;justify-content:space-between">
    <span>Schmal Unternehmensberatung · 5 Gründe, warum KI-Projekte scheitern</span><span class="pageNumber"></span></div>`;
  const basis = { format: 'A4', printBackground: true, preferCSSPageSize: true };
  const cover = await page.pdf({ ...basis, pageRanges: '1' });
  const rest = await page.pdf({
    ...basis,
    pageRanges: '2-',
    displayHeaderFooter: true,
    headerTemplate: '<span></span>',
    footerTemplate: fuss,
  });
  await browser.close();

  fs.writeFileSync(path.join(__dirname, '.cover.pdf'), cover);
  fs.writeFileSync(path.join(__dirname, '.rest.pdf'), rest);
})();
