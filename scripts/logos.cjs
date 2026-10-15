// Erzeugt aus den hochgeladenen Logos (Hintergrund #0F2438) freigestellte Web-Dateien,
// Favicons und das Vorschaubild für soziale Netzwerke. Aufruf: node scripts/logos.cjs
const sharp = require('sharp');
const path = require('path');

const BG = [15, 36, 56]; // #0F2438
const D = 200; // Farbabstand, ab dem ein Pixel als voll deckend gilt
const quelle = (f) => path.join(__dirname, '..', 'assets', 'logo', f);
const ziel = (f) => path.join(__dirname, '..', f);

// Hintergrund herausrechnen: Alpha aus dem Abstand zum Navy, Farbe zurückgewinnen.
async function freistellen(datei) {
  const { data, info } = await sharp(quelle(datei)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  for (let i = 0; i < data.length; i += 4) {
    const d = Math.hypot(data[i] - BG[0], data[i + 1] - BG[1], data[i + 2] - BG[2]);
    const a = Math.min(1, d / D);
    if (a < 0.02) {
      data[i + 3] = 0;
      continue;
    }
    for (let c = 0; c < 3; c++) {
      data[i + c] = Math.max(0, Math.min(255, Math.round((data[i + c] - (1 - a) * BG[c]) / a)));
    }
    data[i + 3] = Math.round(a * 255);
  }
  return sharp(data, { raw: info }).trim({ threshold: 1 }).png().toBuffer();
}

(async () => {
  const quer = await freistellen('Schmalo S Unternehmen.png');
  const zeichen = await freistellen('Schmalo S Unternehmen (1).png');

  // Freigestellte Master-Dateien (für Website, Whitepaper, Druck)
  await sharp(quer).toFile(ziel('assets/logo/logo-quer-transparent.png'));
  await sharp(zeichen).toFile(ziel('assets/logo/logo-zeichen-transparent.png'));

  // Favicons: Bildzeichen auf Navy mit Rand
  const icon = async (groesse, datei, rand = 0.14) => {
    const innen = Math.round(groesse * (1 - 2 * rand));
    const z = await sharp(zeichen).resize(innen, innen, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).toBuffer();
    await sharp({ create: { width: groesse, height: groesse, channels: 4, background: { r: BG[0], g: BG[1], b: BG[2], alpha: 1 } } })
      .composite([{ input: z, gravity: 'center' }])
      .png()
      .toFile(ziel(datei));
  };
  await icon(32, 'public/favicon-32.png', 0.08);
  await icon(48, 'public/favicon-48.png', 0.1);
  await icon(180, 'public/apple-touch-icon.png');
  await icon(192, 'public/icon-192.png');
  await icon(512, 'public/icon-512.png');

  // Vorschaubild für LinkedIn, WhatsApp & Co. (1200 x 630)
  const q = await sharp(quer).resize(760).toBuffer();
  await sharp({ create: { width: 1200, height: 630, channels: 4, background: { r: BG[0], g: BG[1], b: BG[2], alpha: 1 } } })
    .composite([{ input: q, gravity: 'center' }])
    .png()
    .toFile(ziel('public/og-image.png'));
})();
