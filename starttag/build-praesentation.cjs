// Erzeugt die Präsentation zum KI-Starttag (PowerPoint, 16:9).
// Aufruf: NODE_PATH=<Ordner mit pptxgenjs, sharp, react, react-dom, react-icons> node starttag/build-praesentation.cjs
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const lu = require('react-icons/lu');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(__dirname, 'KI-Starttag-Praesentation.pptx');
const SKILL = process.env.PPTX_SKILL; // Ordner des pptx-Skills (für apply_theme.js)

const HEX = {
  ink: '1A2330', navy: '0F2438', navy2: '1A3A5C', gold: 'C9A87C', goldDeep: '7A5C33',
  stone: '5A6572', light: 'EEF2F6', white: 'FFFFFF', green: '2E7D4F', red: 'B23B3B', muted: 'B4C0CF',
};

const THEME = {
  name: 'Schmal Unternehmensberatung',
  headFontFace: 'Cambria',
  bodyFontFace: 'Calibri',
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.light,
    accent1: HEX.gold, accent2: HEX.navy2, accent3: HEX.goldDeep, accent4: HEX.stone,
    accent5: HEX.green, accent6: HEX.red, hlink: HEX.navy2, folHlink: HEX.goldDeep,
  },
};

async function icon(name, hex) {
  const Cmp = lu[name] || lu.LuCircle;
  const svg = renderToStaticMarkup(React.createElement(Cmp, { size: 256, color: '#' + hex }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.333 × 7.5 in
  pres.author = 'Eduard Schmal';
  pres.company = 'Schmal Unternehmensberatung';
  pres.title = 'KI-Starttag';
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  const logoQuer = path.join(ROOT, 'assets/logo/logo-quer-transparent.png');
  const logoZeichen = path.join(ROOT, 'assets/logo/logo-zeichen-transparent.png');
  const hero = path.join(ROOT, 'assets/hero.png');
  const about = path.join(ROOT, 'assets/about.png');
  const qrPng = 'image/png;base64,' + (await sharp(path.join(ROOT, 'visitenkarte/qr-vcard.svg'), { density: 600 })
    .resize(800, 800).flatten({ background: '#ffffff' }).extend({ top: 40, bottom: 40, left: 40, right: 40, background: '#ffffff' })
    .png().toBuffer()).toString('base64');

  // ---------- Layouts ----------
  pres.defineSlideMaster({
    title: 'Titel',
    background: { color: C.text2 },
    objects: [],
  });
  pres.defineSlideMaster({
    title: 'Kapitel',
    background: { color: C.text2 },
    objects: [
      { image: { path: logoZeichen, x: 12.17, y: 6.55, w: 0.51, h: 0.55 } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 3.0, w: 10.5, h: 1.4, fontFace: THEME.headFontFace, fontSize: 44, bold: false, color: C.background1, valign: 'top', align: 'left', margin: 0 }, text: '' } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.45, w: 10.5, h: 0.9, fontFace: THEME.bodyFontFace, fontSize: 18, color: C.accent1, valign: 'top', margin: 0 }, text: '' } },
    ],
  });
  pres.defineSlideMaster({
    title: 'Inhalt',
    background: { color: C.background1 },
    objects: [
      { image: { path: logoZeichen, x: 12.38, y: 6.85, w: 0.35, h: 0.38 } },
      { text: { text: 'KI-Starttag · Schmal Unternehmensberatung', options: { x: 0.6, y: 6.9, w: 6, h: 0.3, fontFace: THEME.bodyFontFace, fontSize: 10, color: C.accent4, margin: 0 } } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.45, w: 12.1, h: 0.9, fontFace: THEME.headFontFace, fontSize: 34, color: C.text2, valign: 'top', align: 'left', margin: 0 }, text: '' } },
    ],
    slideNumber: { x: 11.6, y: 6.9, w: 0.6, h: 0.3, fontFace: THEME.bodyFontFace, fontSize: 10, color: C.accent4, align: 'right' },
  });
  pres.defineSlideMaster({
    title: 'Abschluss',
    background: { color: C.text2 },
    objects: [],
  });

  // ---------- Bausteine ----------
  const txt = (slide, text, opts) => slide.addText(text, { isTextBox: true, margin: 0, fontFace: THEME.bodyFontFace, color: C.text1, valign: 'top', ...opts });

  async function card(slide, { x, y, w, h, ic, title, text, fill = C.background2, iconColor = HEX.navy, titleColor = C.text2, textColor = C.text1 }) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Karte ' + title });
    let ty = y + 0.3;
    if (ic) {
      slide.addImage({ data: await icon(ic, iconColor), x: x + 0.3, y: y + 0.3, w: 0.5, h: 0.5, objectName: 'Symbol ' + title });
      ty = y + 1.0;
    }
    txt(slide, title, { x: x + 0.3, y: ty, w: w - 0.6, h: 0.45, fontFace: THEME.headFontFace, fontSize: 20, color: titleColor, bold: true });
    if (text) txt(slide, text, { x: x + 0.3, y: ty + 0.5, w: w - 0.6, h: h - (ty - y) - 0.7, fontSize: 16, color: textColor });
  }

  function badge(slide, n, x, y, d = 0.6, fill = C.accent1, color = C.text2) {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' }, objectName: 'Nummer ' + n });
    txt(slide, String(n), { x, y, w: d, h: d, align: 'center', valign: 'middle', fontFace: THEME.headFontFace, fontSize: 20, bold: true, color });
  }

  const content = (sectionTitle, title) => {
    const s = pres.addSlide({ masterName: 'Inhalt', sectionTitle });
    s.addText(title, { placeholder: 'title' });
    return s;
  };
  const chapter = (sectionTitle, nr, title, sub) => {
    const s = pres.addSlide({ masterName: 'Kapitel', sectionTitle });
    txt(s, `Teil ${nr}`, { x: 0.8, y: 2.35, w: 4, h: 0.5, fontSize: 16, bold: true, color: C.accent1, charSpacing: 3 });
    s.addText(title, { placeholder: 'title' });
    s.addText(sub, { placeholder: 'body' });
    return s;
  };

  // ======================================================================
  // Einstieg
  // ======================================================================
  pres.addSection({ title: 'Einstieg' });
  {
    const s = pres.addSlide({ masterName: 'Titel', sectionTitle: 'Einstieg' });
    s.addImage({ path: hero, x: 6.4, y: 0, w: 6.933, h: 7.5, sizing: { type: 'cover', w: 6.933, h: 7.5 }, altText: 'Eduard Schmal', objectName: 'Foto' });
    s.addImage({ path: logoQuer, x: 0.8, y: 0.7, w: 3.2, h: 3.2 * (342 / 1299), objectName: 'Logo' });
    txt(s, 'KI-STARTTAG', { x: 0.8, y: 2.55, w: 5.2, h: 0.4, fontSize: 16, bold: true, color: C.accent1, charSpacing: 4 });
    txt(s, 'KI sicher und wirksam im Arbeitsalltag', { x: 0.8, y: 3.0, w: 5.3, h: 1.9, fontFace: THEME.headFontFace, fontSize: 40, color: C.background1 });
    txt(s, 'Schulung mit Nachweis und Bordmittel-Check', { x: 0.8, y: 5.0, w: 5.3, h: 0.5, fontSize: 18, color: C.background1 });
    txt(s, 'Eduard Schmal · Schmal Unternehmensberatung', { x: 0.8, y: 6.55, w: 5.3, h: 0.4, fontSize: 13, color: C.accent1 });
    s.addNotes('Begrüßung. Kurz Rahmen setzen: Heute geht es nicht um Technik, sondern darum, dass jede und jeder nach vier Stunden KI sicher für die eigene Arbeit nutzen kann. Handys dürfen heute ausnahmsweise auf dem Tisch bleiben, wir probieren aus.');
  }
  {
    const s = content('Einstieg', 'Wer heute vor Ihnen steht');
    s.addImage({ path: about, x: 0.6, y: 1.55, w: 4.2, h: 5.0, sizing: { type: 'cover', w: 4.2, h: 5.0 }, rounding: false, altText: 'Eduard Schmal am Schreibtisch', objectName: 'Foto Eduard Schmal' });
    const punkte = [
      ['LuBriefcase', '20 Jahre Führungs- und Prozesserfahrung', 'Unter anderem Aufbau eines zentralen Servicebereichs mit fast 50 Mitarbeitenden.'],
      ['LuGraduationCap', 'Manager für angewandte KI-Transformation (IHK)', 'Mit Auszeichnung. Dozent an der IHK.'],
      ['LuTrendingUp', 'Ergebnisse aus der Praxis', '35 % weniger Zeitaufwand in der Buchhaltung, 1,5 Vollzeitstellen in 2 Monaten eingespart, mit Bordmitteln.'],
    ];
    let y = 1.6;
    for (const [ic, t, d] of punkte) {
      s.addShape(pres.shapes.OVAL, { x: 5.4, y, w: 0.7, h: 0.7, fill: { color: C.background2 }, line: { type: 'none' }, objectName: 'Kreis ' + t });
      s.addImage({ data: await icon(ic, HEX.navy), x: 5.55, y: y + 0.15, w: 0.4, h: 0.4, objectName: 'Symbol ' + t });
      txt(s, t, { x: 6.35, y: y - 0.02, w: 6.3, h: 0.45, fontFace: THEME.headFontFace, fontSize: 20, bold: true, color: C.text2 });
      txt(s, d, { x: 6.35, y: y + 0.45, w: 6.3, h: 0.9, fontSize: 15 });
      y += 1.6;
    }
    s.addNotes('Kurz vorstellen, nicht länger als zwei Minuten. Wichtigste Botschaft: Ich komme aus der Praxis, nicht aus der Theorie. Ein konkretes Beispiel aus einem Projekt erzählen, z. B. der Rechnungsworkflow.');
  }
  {
    const s = content('Einstieg', 'So läuft der Tag');
    const bloecke = [
      ['0:00', 'Begrüßung und Ziele', 'Wo steht Ihr Team heute?'],
      ['0:15', 'Wie KI arbeitet', 'Stärken, Grenzen, erfundene Antworten'],
      ['0:45', 'Gute Aufträge', 'Formel, Live-Demo, erste Übung'],
      ['1:45', 'Sicher arbeiten', 'Daten, Prüfroutine, Regeln, EU AI Act'],
      ['2:30', 'Eigene Fälle', 'Jede Person nimmt einen Fall mit'],
      ['3:10', 'Bordmittel-Check', 'Mit der Geschäftsführung'],
    ];
    const w = 1.9, gap = 0.12, x0 = 0.6, y = 1.9;
    bloecke.forEach(([zeit, t, d], i) => {
      const x = x0 + i * (w + gap);
      const dunkel = i === 5;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 3.6, fill: { color: dunkel ? C.text2 : C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Block ' + t });
      txt(s, zeit, { x: x + 0.25, y: y + 0.3, w: w - 0.5, h: 0.6, fontFace: THEME.headFontFace, fontSize: 28, color: dunkel ? C.accent1 : C.accent3 });
      txt(s, t, { x: x + 0.25, y: y + 1.05, w: w - 0.5, h: 0.95, fontFace: THEME.headFontFace, fontSize: 18, bold: true, color: dunkel ? C.background1 : C.text2 });
      txt(s, d, { x: x + 0.25, y: y + 2.05, w: w - 0.5, h: 1.3, fontSize: 14, color: dunkel ? C.background1 : C.text1 });
    });
    txt(s, 'Pause: 1:30 bis 1:45 · Ende Schulungsteil: 3:10', { x: 0.6, y: 5.8, w: 12, h: 0.4, fontSize: 15, color: C.accent4 });
    s.addNotes('Zeiten sind Richtwerte. Die Übungen haben Vorrang vor Theorie: Lieber eine Folie überspringen als eine Übung kürzen. Der Bordmittel-Check ab 3:10 läuft nur mit der Geschäftsführung, die übrigen Teilnehmenden sind dann fertig.');
  }
  {
    const s = content('Einstieg', 'Ihr Ziel für heute');
    await card(s, { x: 0.6, y: 1.7, w: 3.9, h: 2.9, ic: 'LuBrain', title: 'Verstehen', text: 'Sie wissen, wie KI zu ihren Antworten kommt und wo sie sich irrt.' });
    await card(s, { x: 4.72, y: 1.7, w: 3.9, h: 2.9, ic: 'LuMessagesSquare', title: 'Anwenden', text: 'Sie formulieren Arbeitsaufträge, die verlässliche Ergebnisse liefern.' });
    await card(s, { x: 8.84, y: 1.7, w: 3.9, h: 2.9, ic: 'LuShieldCheck', title: 'Sicher arbeiten', text: 'Sie wissen, welche Daten in welches Werkzeug dürfen, und prüfen jedes Ergebnis.' });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.95, w: 12.14, h: 1.1, fill: { color: C.text2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Kernziel' });
    txt(s, 'Am Ende hat jede Person einen eigenen Anwendungsfall, den sie ab Montag nutzt.', { x: 0.9, y: 4.95, w: 11.5, h: 1.1, valign: 'middle', fontFace: THEME.headFontFace, fontSize: 22, color: C.background1 });
    s.addNotes('Ziele vorlesen und fragen: Was möchten Sie persönlich heute mitnehmen? Zwei, drei Antworten an Flipchart oder Whiteboard notieren und am Ende darauf zurückkommen.');
  }
  {
    const s = content('Einstieg', 'Wo stehen Sie heute?');
    const fragen = [
      ['LuHand', 'Wer hat KI schon einmal genutzt, beruflich oder privat?'],
      ['LuListChecks', 'Wofür? Was hat gut funktioniert, was nicht?'],
      ['LuHelpCircle', 'Was hält Sie bisher davon ab, KI mehr zu nutzen?'],
    ];
    let y = 1.75;
    for (const [ic, f] of fragen) {
      s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.8, h: 0.8, fill: { color: C.accent1 }, line: { type: 'none' }, objectName: 'Kreis Frage' });
      s.addImage({ data: await icon(ic, HEX.navy), x: 0.78, y: y + 0.18, w: 0.44, h: 0.44, objectName: 'Symbol Frage' });
      txt(s, f, { x: 1.7, y, w: 6.2, h: 0.8, valign: 'middle', fontFace: THEME.headFontFace, fontSize: 22, color: C.text2 });
      y += 1.3;
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.5, y: 1.75, w: 4.24, h: 3.4, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Hinweis Runde' });
    txt(s, 'Kurze Runde', { x: 8.8, y: 2.05, w: 3.7, h: 0.45, fontFace: THEME.headFontFace, fontSize: 20, bold: true, color: C.text2 });
    txt(s, 'Jede Person antwortet in einem Satz. Es gibt kein Richtig oder Falsch.\n\nDie Antworten aus dem Vorab-Fragebogen fließen in die Übungen ein.', { x: 8.8, y: 2.6, w: 3.7, h: 2.3, fontSize: 15 });
    s.addNotes('Diese Runde zeigt Ihnen, wo die Gruppe steht. Bei sehr gemischtem Stand: erfahrene Teilnehmende in den Übungen als Tandempartner einsetzen. Antworten zu Hürden notieren, sie sind wertvoll für den Kurzbericht an die Geschäftsführung.');
  }

  // ======================================================================
  // Teil 1: Wie KI arbeitet
  // ======================================================================
  pres.addSection({ title: 'Wie KI arbeitet' });
  chapter('Wie KI arbeitet', 1, 'Wie KI arbeitet', 'Was ein Sprachmodell tut, was es kann und wo es sich irrt').addNotes('Übergang: Bevor wir mit KI arbeiten, schauen wir kurz unter die Haube. Nur so viel Technik, wie Sie für gute Ergebnisse brauchen.');
  {
    const s = content('Wie KI arbeitet', 'Drei Begriffe, die Sie kennen sollten');
    await card(s, { x: 0.6, y: 1.7, w: 3.9, h: 3.6, ic: 'LuCpu', title: 'Künstliche Intelligenz', text: 'Oberbegriff für Systeme, die Aufgaben lösen, für die sonst menschliches Denken nötig ist.' });
    await card(s, { x: 4.72, y: 1.7, w: 3.9, h: 3.6, ic: 'LuLineChart', title: 'Maschinelles Lernen', text: 'Systeme lernen Muster aus großen Datenmengen, statt fest programmiert zu werden. Beispiel: Spamfilter.' });
    await card(s, { x: 8.84, y: 1.7, w: 3.9, h: 3.6, ic: 'LuSparkles', title: 'Generative KI', text: 'Erzeugt neue Inhalte: Texte, Bilder, Zusammenfassungen. Dazu gehören ChatGPT, Copilot und die Modelle in Langdock.', fill: C.text2, iconColor: HEX.gold, titleColor: C.accent1, textColor: C.background1 });
    s.addNotes('Ziel: Begriffe sortieren, nicht vertiefen. Heute geht es nur um generative KI, also Werkzeuge, die Texte erzeugen. Beispiel aus dem Alltag nennen: Viele nutzen KI schon unbewusst, etwa in der Buchhaltungssoftware, die Belege vorsortiert.');
  }
  {
    const s = content('Wie KI arbeitet', 'Wahrscheinlichkeit statt Wissen');
    const schritte = [
      ['LuKeyboard', 'Ihre Eingabe', 'Text wird in kleine Bausteine (Token) zerlegt'],
      ['LuShuffle', 'Berechnung', 'Das Modell berechnet das wahrscheinlichste nächste Wort'],
      ['LuRepeat', 'Wort für Wort', 'Die Antwort entsteht Schritt für Schritt'],
      ['LuFileText', 'Antwort', 'Klingt immer überzeugend, ist aber nicht geprüft'],
    ];
    const w = 2.75, gap = 0.38;
    for (let i = 0; i < schritte.length; i++) {
      const [ic, t, d] = schritte[i];
      const x = 0.6 + i * (w + gap);
      await card(s, { x, y: 1.75, w, h: 3.2, ic, title: t, text: d });
      if (i < schritte.length - 1) s.addImage({ data: await icon('LuArrowRight', HEX.gold), x: x + w + 0.04, y: 3.2, w: 0.3, h: 0.3, objectName: 'Pfeil' });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.3, w: 12.14, h: 1.1, fill: { color: C.text2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Merksatz' });
    txt(s, [
      { text: 'Merksatz: ', options: { bold: true, color: C.accent1 } },
      { text: 'Ein Sprachmodell weiß nichts, es rechnet. Es kennt nur, was in seinen Trainingsdaten und in Ihrem Auftrag steht (Kontextfenster).', options: { color: C.background1 } },
    ], { x: 0.9, y: 5.3, w: 11.6, h: 1.1, valign: 'middle', fontSize: 17 });
    s.addNotes('Bild: eine extrem belesene Autovervollständigung. Daraus folgt alles Weitere: Die KI klingt immer sicher, auch wenn sie falsch liegt. Und sie weiß nichts über Ihr Unternehmen, außer Sie geben es ihr im Auftrag mit. Kontextfenster = wie viel Text das Modell gleichzeitig berücksichtigen kann.');
  }
  {
    const s = content('Wie KI arbeitet', 'Wofür KI taugt und wofür nicht');
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.7, w: 5.95, h: 4.7, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Stärken' });
    s.addImage({ data: await icon('LuThumbsUp', HEX.green), x: 0.9, y: 2.0, w: 0.5, h: 0.5, objectName: 'Symbol Stärken' });
    txt(s, 'Stark bei', { x: 1.6, y: 2.0, w: 4.5, h: 0.5, valign: 'middle', fontFace: THEME.headFontFace, fontSize: 22, bold: true, color: C.accent5 });
    txt(s, [
      { text: 'Entwürfe für Mails, Angebote, Berichte', options: { bullet: true, breakLine: true } },
      { text: 'Zusammenfassen langer Dokumente', options: { bullet: true, breakLine: true } },
      { text: 'Umformulieren, kürzen, übersetzen', options: { bullet: true, breakLine: true } },
      { text: 'Ideen sammeln, Gliederungen erstellen', options: { bullet: true, breakLine: true } },
      { text: 'Protokolle aus Notizen', options: { bullet: true } },
    ], { x: 0.9, y: 2.75, w: 5.4, h: 3.4, fontSize: 16, paraSpaceAfter: 8 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.79, y: 1.7, w: 5.95, h: 4.7, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Grenzen' });
    s.addImage({ data: await icon('LuAlertTriangle', HEX.red), x: 7.09, y: 2.0, w: 0.5, h: 0.5, objectName: 'Symbol Grenzen' });
    txt(s, 'Vorsicht bei', { x: 7.79, y: 2.0, w: 4.5, h: 0.5, valign: 'middle', fontFace: THEME.headFontFace, fontSize: 22, bold: true, color: C.accent6 });
    txt(s, [
      { text: 'Fakten, Zahlen, Paragrafen ohne Quelle', options: { bullet: true, breakLine: true } },
      { text: 'Rechnen und Zählen', options: { bullet: true, breakLine: true } },
      { text: 'Aktuellen Ereignissen', options: { bullet: true, breakLine: true } },
      { text: 'Entscheidungen über Menschen', options: { bullet: true, breakLine: true } },
      { text: 'Allem, was ungeprüft zum Kunden geht', options: { bullet: true } },
    ], { x: 7.09, y: 2.75, w: 5.4, h: 3.4, fontSize: 16, paraSpaceAfter: 8 });
    s.addNotes('Faustregel: Je mehr eine Aufgabe mit Sprache und Form zu tun hat, desto besser. Je mehr es auf exakte Fakten ankommt, desto mehr Prüfung braucht es. Frage an die Gruppe: Welche Ihrer Aufgaben gehören links hin?');
  }
  {
    const s = content('Wie KI arbeitet', 'Erfundene Antworten erkennen');
    const zeichen = [
      ['LuQuote', 'Sehr präzise Details ohne Quelle', 'Exakte Zahlen, Daten, Namen oder Paragrafen, die Sie nirgends wiederfinden.'],
      ['LuLink2', 'Quellen, die es nicht gibt', 'Links, Studien oder Urteile, die beim Nachprüfen ins Leere laufen.'],
      ['LuBadgeCheck', 'Übertriebene Sicherheit', 'Keine Einschränkung, kein „vermutlich“, auch bei schwierigen Fragen.'],
      ['LuArrowLeftRight', 'Widersprüche', 'Auf Nachfrage ändert sich die Antwort oder widerspricht sich selbst.'],
    ];
    for (let i = 0; i < 4; i++) {
      const [ic, t, d] = zeichen[i];
      const x = i % 2 === 0 ? 0.6 : 6.79;
      const y = i < 2 ? 1.7 : 3.95;
      await card(s, { x, y, w: 5.95, h: 2.05, ic: null, title: t, text: d });
      s.addImage({ data: await icon(ic, HEX.goldDeep), x: x + 5.15, y: y + 0.3, w: 0.45, h: 0.45, objectName: 'Symbol ' + t });
    }
    txt(s, [
      { text: 'Grundregel: ', options: { bold: true, color: C.accent3 } },
      { text: 'Übernehmen Sie nur, was Sie selbst prüfen können oder was eine nachprüfbare Quelle hat.' },
    ], { x: 0.6, y: 6.15, w: 12.1, h: 0.45, fontSize: 17 });
    s.addNotes('Die KI erfindet nicht aus Böswilligkeit, sondern weil sie das Wahrscheinlichste schreibt. In der Live-Demo gleich zeigen: nach einer Quelle fragen und prüfen, ob es sie gibt.');
  }

  // ======================================================================
  // Teil 2: Gute Arbeitsaufträge
  // ======================================================================
  pres.addSection({ title: 'Gute Arbeitsaufträge' });
  chapter('Gute Arbeitsaufträge', 2, 'Gute Arbeitsaufträge', 'Eine Formel, eine Live-Demo und Ihre erste Übung').addNotes('Übergang: Jetzt wird es praktisch. Die Qualität der Antwort hängt zu 80 % an der Qualität des Auftrags.');
  {
    const s = content('Gute Arbeitsaufträge', 'Die Formel: Rolle, Kontext, Aufgabe, Format');
    const teile = [
      ['LuUser', 'Rolle', 'Wer soll die KI sein?', '„Du bist Sachbearbeiterin im Kundenservice.“'],
      ['LuInfo', 'Kontext', 'Was muss sie wissen?', '„Der Kunde hat zum zweiten Mal eine falsche Lieferung erhalten.“'],
      ['LuTarget', 'Aufgabe', 'Was genau soll sie tun?', '„Schreibe eine Antwort, die sich entschuldigt und eine Lösung anbietet.“'],
      ['LuLayoutList', 'Format', 'Wie soll das Ergebnis aussehen?', '„Höflich, per Sie, höchstens 120 Wörter.“'],
    ];
    const w = 2.9, gap = 0.18;
    for (let i = 0; i < 4; i++) {
      const [ic, t, q, b] = teile[i];
      const x = 0.6 + i * (w + gap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.7, w, h: 4.6, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Formel ' + t });
      badge(s, i + 1, x + 0.3, 2.0);
      s.addImage({ data: await icon(ic, HEX.navy), x: x + w - 0.8, y: 2.05, w: 0.5, h: 0.5, objectName: 'Symbol ' + t });
      txt(s, t, { x: x + 0.3, y: 2.85, w: w - 0.6, h: 0.5, fontFace: THEME.headFontFace, fontSize: 24, bold: true, color: C.text2 });
      txt(s, q, { x: x + 0.3, y: 3.4, w: w - 0.6, h: 0.7, fontSize: 15, bold: true, color: C.accent3 });
      txt(s, b, { x: x + 0.3, y: 4.15, w: w - 0.6, h: 1.9, fontSize: 15, italic: true });
    }
    s.addNotes('Die vier Bausteine einzeln erklären und dann zusammen vorlesen. Nicht jeder Auftrag braucht alle vier, aber je wichtiger das Ergebnis, desto vollständiger. Tipp: eigene Unterlagen als Kontext mitgeben, statt alles abzutippen.');
  }
  {
    const s = content('Gute Arbeitsaufträge', 'Vorher und nachher');
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.7, w: 5.95, h: 4.7, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Vorher' });
    txt(s, 'Vorher', { x: 0.9, y: 1.95, w: 5.3, h: 0.5, fontFace: THEME.headFontFace, fontSize: 22, bold: true, color: C.accent6 });
    txt(s, '„Schreib eine Antwort auf die Beschwerde.“', { x: 0.9, y: 2.6, w: 5.3, h: 1.0, fontSize: 18, italic: true });
    txt(s, 'Ergebnis: allgemein, zu lang, falscher Ton, Sie schreiben fast alles neu.', { x: 0.9, y: 4.6, w: 5.3, h: 1.4, fontSize: 15, color: C.accent4 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.79, y: 1.7, w: 5.95, h: 4.7, fill: { color: C.text2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Nachher' });
    txt(s, 'Nachher', { x: 7.09, y: 1.95, w: 5.3, h: 0.5, fontFace: THEME.headFontFace, fontSize: 22, bold: true, color: C.accent1 });
    txt(s, '„Du bist Sachbearbeiterin im Kundenservice eines Handwerksbetriebs. Ein Stammkunde hat zum zweiten Mal falsches Material erhalten, seine Mail hängt an. Schreibe eine Antwort, die sich entschuldigt, eine Lieferung bis Freitag zusagt und einen Rabatt von 5 % anbietet. Höflich, per Sie, höchstens 120 Wörter.“', { x: 7.09, y: 2.6, w: 5.35, h: 3.0, fontSize: 15, italic: true, color: C.background1 });
    txt(s, 'Ergebnis: passt fast sofort, nur noch prüfen.', { x: 7.09, y: 5.75, w: 5.3, h: 0.45, fontSize: 15, color: C.accent1 });
    s.addNotes('Beide Aufträge gleich live in Langdock eingeben und die Ergebnisse vergleichen. Die Gruppe abstimmen lassen, welches Ergebnis sie so verschicken würde.');
  }
  {
    const s = content('Gute Arbeitsaufträge', 'Nachschärfen statt neu anfangen');
    const schritte = [
      ['LuMessageSquare', 'Erster Entwurf', 'Mit der Formel starten'],
      ['LuPencilLine', 'Gezielt korrigieren', '„Kürzer“, „förmlicher“, „ohne Fachbegriffe“, „mit Aufzählung“'],
      ['LuCopyCheck', 'Prüfen und übernehmen', 'Fakten prüfen, persönlich machen, verschicken'],
    ];
    for (let i = 0; i < 3; i++) {
      const [ic, t, d] = schritte[i];
      const x = 0.6 + i * 4.12;
      await card(s, { x, y: 1.75, w: 3.9, h: 3.0, ic, title: t, text: d });
      badge(s, i + 1, x + 3.1, 2.0, 0.5);
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.05, w: 12.14, h: 1.35, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Tipp' });
    s.addImage({ data: await icon('LuLightbulb', HEX.goldDeep), x: 0.9, y: 5.4, w: 0.6, h: 0.6, objectName: 'Symbol Tipp' });
    txt(s, 'Gute Aufträge speichern: Was einmal funktioniert hat, legen Sie als Vorlage ab. In Langdock geht das als eigener Assistent für das ganze Team.', { x: 1.75, y: 5.05, w: 10.7, h: 1.35, valign: 'middle', fontSize: 16 });
    s.addNotes('Häufigster Fehler: Nach einem mittelmäßigen Ergebnis aufgeben. Besser zwei, drei gezielte Korrekturen. Hinweis auf die Vorlagensammlung, die alle nach dem Starttag erhalten.');
  }
  {
    const s = content('Gute Arbeitsaufträge', 'Live-Demo: Langdock');
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.7, w: 4.6, h: 4.7, fill: { color: C.text2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Langdock Info' });
    s.addImage({ data: await icon('LuMonitorPlay', HEX.gold), x: 0.95, y: 2.05, w: 0.7, h: 0.7, objectName: 'Symbol Demo' });
    txt(s, 'Warum Langdock?', { x: 0.95, y: 3.0, w: 4.0, h: 0.5, fontFace: THEME.headFontFace, fontSize: 22, bold: true, color: C.accent1 });
    txt(s, [
      { text: 'Datenschutzkonforme KI-Plattform aus Deutschland', options: { bullet: true, breakLine: true } },
      { text: 'Mehrere Sprachmodelle an einem Ort', options: { bullet: true, breakLine: true } },
      { text: 'Eigene Assistenten mit Unternehmenswissen', options: { bullet: true } },
    ], { x: 0.95, y: 3.6, w: 4.0, h: 2.6, fontSize: 15, color: C.background1, paraSpaceAfter: 8 });
    const demo = [
      'Vorher und nachher: die Beschwerde-Antwort',
      'Ein Dokument hochladen und zusammenfassen lassen',
      'Nach einer Quelle fragen und sie prüfen',
      'Aus einer Notiz ein Protokoll erstellen',
      'Einen Assistenten mit Vorlage für das Team anlegen',
    ];
    demo.forEach((d, i) => {
      const y = 1.75 + i * 0.93;
      badge(s, i + 1, 5.6, y, 0.55);
      txt(s, d, { x: 6.4, y, w: 6.3, h: 0.55, valign: 'middle', fontSize: 18, color: C.text2 });
    });
    s.addNotes('Vor dem Termin prüfen: Internetzugang beim Kunden, Bildschirmübertragung, Langdock-Zugang mit Demo-Arbeitsbereich. Nur Beispieldaten verwenden, keine echten Kundendaten des Unternehmens. Transparenz: Ich arbeite mit Langdock als Partner zusammen, das offen sagen. Die Grundsätze gelten genauso für Copilot oder ChatGPT.');
  }
  {
    const s = content('Gute Arbeitsaufträge', 'Übung 1: Ihr erster Auftrag');
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.7, w: 3.3, h: 4.7, fill: { color: C.accent1 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Übungszeit' });
    s.addImage({ data: await icon('LuTimer', HEX.navy), x: 0.95, y: 2.05, w: 0.7, h: 0.7, objectName: 'Symbol Zeit' });
    txt(s, '20', { x: 0.95, y: 2.95, w: 2.6, h: 1.3, fontFace: THEME.headFontFace, fontSize: 72, bold: true, color: C.text2 });
    txt(s, 'Minuten, allein oder zu zweit', { x: 0.95, y: 4.35, w: 2.7, h: 1.0, fontSize: 16, color: C.text2 });
    const schritte = [
      'Wählen Sie eine Aufgabe, die Sie diese Woche wirklich erledigen müssen.',
      'Schreiben Sie einen Auftrag nach der Formel: Rolle, Kontext, Aufgabe, Format.',
      'Schärfen Sie das Ergebnis mindestens zweimal nach.',
      'Prüfen Sie: Würden Sie das so verschicken? Was mussten Sie ändern?',
    ];
    schritte.forEach((d, i) => {
      const y = 1.8 + i * 1.12;
      badge(s, i + 1, 4.3, y, 0.6);
      txt(s, d, { x: 5.15, y: y - 0.05, w: 7.5, h: 0.75, valign: 'middle', fontSize: 18 });
    });
    txt(s, 'Keine echten Kunden- oder Personaldaten verwenden.', { x: 4.3, y: 6.15, w: 8.4, h: 0.4, fontSize: 14, color: C.accent6, bold: true });
    s.addNotes('Herumgehen und helfen. Zwei, drei gute Beispiele am Ende kurz zeigen lassen. Wer schnell fertig ist: dieselbe Aufgabe mit einem anderen Format (Tabelle, Stichpunkte) wiederholen.');
  }

  // ======================================================================
  // Teil 3: Sicher arbeiten
  // ======================================================================
  pres.addSection({ title: 'Sicher arbeiten' });
  chapter('Sicher arbeiten', 3, 'Sicher arbeiten', 'Daten, Prüfroutine, Regeln und der rechtliche Rahmen').addNotes('Nach der Pause. Übergang: KI nutzen ist einfach, sie sicher nutzen braucht ein paar klare Regeln.');
  {
    const s = content('Sicher arbeiten', 'Darf das hier rein?');
    const ampel = [
      [C.accent5, 'Grün', 'Unbedenklich', 'Öffentliche Informationen, allgemeine Texte, eigene Formulierungen ohne Namen, anonymisierte Beispiele.'],
      [C.accent1, 'Gelb', 'Nur in freigegebenen Werkzeugen', 'Interne Unterlagen, Angebote, Prozessbeschreibungen. Nur in Werkzeugen mit Vertrag zur Auftragsverarbeitung, z. B. Langdock.'],
      [C.accent6, 'Rot', 'Nicht ohne Freigabe', 'Gesundheits-, Personal- und Mandantendaten, Passwörter, Geschäftsgeheimnisse, Bankdaten.'],
    ];
    for (let i = 0; i < 3; i++) {
      const [col, t, sub, d] = ampel[i];
      const y = 1.7 + i * 1.6;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.15, w: 1.1, h: 1.1, fill: { color: col }, line: { type: 'none' }, objectName: 'Ampel ' + t });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.0, y, w: 10.74, h: 1.4, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Ampeltext ' + t });
      txt(s, [{ text: t + ': ', options: { bold: true, color: C.text2 } }, { text: sub, options: { bold: true, color: C.text2 } }], { x: 2.3, y: y + 0.15, w: 10.2, h: 0.45, fontFace: THEME.headFontFace, fontSize: 19 });
      txt(s, d, { x: 2.3, y: y + 0.65, w: 10.2, h: 0.7, fontSize: 15 });
    }
    s.addNotes('Diese Ampel bekommen alle als Karte für den Arbeitsplatz. Wichtig für Kanzleien und Praxen: Mandanten- und Patientendaten unterliegen zusätzlich der Schweigepflicht und gehören immer in Rot. Frage an die Gruppe: Nennen Sie je ein Beispiel aus Ihrem Alltag für jede Farbe.');
  }
  {
    const s = content('Sicher arbeiten', 'Die Prüfroutine vor dem Absenden');
    const pruef = [
      ['LuSearchCheck', 'Fakten', 'Stimmen Namen, Daten, Zusagen?'],
      ['LuCalculator', 'Zahlen', 'Selbst nachrechnen, nie blind übernehmen'],
      ['LuBookOpen', 'Quellen', 'Gibt es die Quelle wirklich? Sagt sie das?'],
      ['LuMessageCircle', 'Ton', 'Passt es zu uns und zum Empfänger?'],
    ];
    const w = 2.9, gap = 0.18;
    for (let i = 0; i < 4; i++) {
      const [ic, t, d] = pruef[i];
      await card(s, { x: 0.6 + i * (w + gap), y: 1.75, w, h: 3.1, ic, title: t, text: d });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.15, w: 12.14, h: 1.25, fill: { color: C.text2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Verantwortung' });
    txt(s, 'Die Verantwortung bleibt bei Ihnen. Was mit Ihrem Namen rausgeht, haben Sie geprüft.', { x: 0.9, y: 5.15, w: 11.6, h: 1.25, valign: 'middle', fontFace: THEME.headFontFace, fontSize: 20, color: C.background1 });
    s.addNotes('Vier Fragen, die in einer Minute erledigt sind. Mit der Übung von vorhin verknüpfen: Hätte Ihr Ergebnis alle vier Prüfungen bestanden?');
  }
  {
    const s = content('Sicher arbeiten', 'Der rechtliche Rahmen in Kürze');
    await card(s, { x: 0.6, y: 1.7, w: 3.9, h: 4.0, ic: 'LuScale', title: 'EU AI Act, Art. 4', text: 'Unternehmen, die KI einsetzen, müssen Maßnahmen ergreifen, um die KI-Kompetenz ihrer Mitarbeitenden zu fördern. Der heutige Tag ist eine solche Maßnahme, dokumentiert mit Nachweis je Person.' });
    await card(s, { x: 4.72, y: 1.7, w: 3.9, h: 4.0, ic: 'LuLock', title: 'Datenschutz (DSGVO)', text: 'Personenbezogene Daten nur in Werkzeugen mit Vertrag zur Auftragsverarbeitung. Im Zweifel anonymisieren oder weglassen.' });
    await card(s, { x: 8.84, y: 1.7, w: 3.9, h: 4.0, ic: 'LuCopyright', title: 'Urheberrecht', text: 'Fremde Texte und Bilder nicht ungeprüft hochladen. KI-Ergebnisse prüfen, bevor Sie sie veröffentlichen, und bei Bedarf als KI-erstellt kennzeichnen.' });
    s.addNotes('Stand November 2026: Art. 4 wurde durch die Verordnung (EU) 2026/1744 (Digital Omnibus, in Kraft seit 27.07.2026) geändert. Unternehmen müssen Maßnahmen zur Förderung der KI-Kompetenz ergreifen, ein bestimmtes Kompetenzniveau wird nicht mehr verlangt. Vor jedem Termin den aktuellen Wortlaut prüfen. Keine Rechtsberatung geben, bei Detailfragen auf Fachanwälte oder Datenschutzbeauftragte verweisen.');
  }
  {
    const s = content('Sicher arbeiten', 'Fünf Regeln für Ihr Unternehmen');
    const regeln = [
      ['Werkzeuge', 'Welche KI-Werkzeuge sind freigegeben?'],
      ['Daten', 'Was darf hinein? Grundlage ist die Ampel.'],
      ['Prüfung', 'Wer prüft Ergebnisse, bevor sie nach außen gehen?'],
      ['Kennzeichnung', 'Wann sagen wir, dass KI beteiligt war?'],
      ['Ansprechperson', 'Wer beantwortet Fragen und sammelt Ideen?'],
    ];
    regeln.forEach(([t, d], i) => {
      const y = 1.7 + i * 0.95;
      badge(s, i + 1, 0.6, y + 0.08, 0.6);
      txt(s, t, { x: 1.45, y, w: 3.0, h: 0.75, valign: 'middle', fontFace: THEME.headFontFace, fontSize: 20, bold: true, color: C.text2 });
      txt(s, d, { x: 4.5, y, w: 4.2, h: 0.75, valign: 'middle', fontSize: 16 });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.0, y: 1.7, w: 3.74, h: 4.7, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Vorlage' });
    s.addImage({ data: await icon('LuClipboardCheck', HEX.navy), x: 9.3, y: 2.0, w: 0.6, h: 0.6, objectName: 'Symbol Vorlage' });
    txt(s, 'Vorlage inklusive', { x: 9.3, y: 2.8, w: 3.2, h: 0.5, fontFace: THEME.headFontFace, fontSize: 20, bold: true, color: C.text2 });
    txt(s, 'Sie erhalten eine einseitige KI-Richtlinie als Vorlage. Die Geschäftsführung passt sie an und gibt sie frei.', { x: 9.3, y: 3.35, w: 3.2, h: 2.8, fontSize: 15 });
    s.addNotes('Mit der Gruppe kurz sammeln, wie die Antworten heute aussehen würden. Oft gibt es noch keine Regeln, und Mitarbeitende nutzen private Zugänge. Das ist ein wichtiger Punkt für den Bordmittel-Check und den Kurzbericht.');
  }

  // ======================================================================
  // Teil 4: Eigene Anwendungsfälle
  // ======================================================================
  pres.addSection({ title: 'Eigene Anwendungsfälle' });
  chapter('Eigene Anwendungsfälle', 4, 'Eigene Anwendungsfälle', 'Jede Person nimmt einen Fall mit, der ab Montag Zeit spart').addNotes('Übergang: Jetzt bringen wir alles zusammen, an Ihrer eigenen Arbeit.');
  {
    const s = content('Eigene Anwendungsfälle', 'Wo KI im Büro am meisten bringt');
    const felder = [
      ['LuMail', 'Schriftverkehr', 'Antworten, Angebote, Erinnerungen'],
      ['LuFileText', 'Zusammenfassen', 'Verträge, Berichte, lange Mails'],
      ['LuClipboardList', 'Protokolle', 'Aus Notizen oder Mitschrift'],
      ['LuLanguages', 'Übersetzen', 'Korrespondenz, Anleitungen'],
      ['LuSearch', 'Recherche', 'Mit Quellenprüfung'],
      ['LuTable', 'Daten aufbereiten', 'Listen ordnen, Tabellen erklären'],
    ];
    for (let i = 0; i < 6; i++) {
      const [ic, t, d] = felder[i];
      const x = 0.6 + (i % 3) * 4.12;
      const y = i < 3 ? 1.7 : 4.15;
      await card(s, { x, y, w: 3.9, h: 2.25, ic, title: t, text: d });
    }
    s.addNotes('Beispiele aus dem Vorab-Fragebogen des Unternehmens hier ergänzen. Frage an die Gruppe: Wo verlieren Sie in Ihrer Woche die meiste Zeit mit Routine?');
  }
  {
    const s = content('Eigene Anwendungsfälle', 'Übung 2: Ihr Anwendungsfall');
    const felder = [
      ['Aufgabe', 'Was genau mache ich heute, und wie oft?'],
      ['Zeit heute', 'Wie lange dauert das pro Woche?'],
      ['KI-Schritt', 'Welchen Teil übernimmt die KI? Mit welchem Auftrag?'],
      ['Prüfung', 'Was prüfe ich, bevor es weitergeht?'],
      ['Ziel', 'Wie viel Zeit will ich sparen?'],
    ];
    felder.forEach(([t, d], i) => {
      const y = 1.7 + i * 0.92;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 8.2, h: 0.78, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.06, objectName: 'Steckbrief ' + t });
      txt(s, t, { x: 0.9, y, w: 2.2, h: 0.78, valign: 'middle', fontFace: THEME.headFontFace, fontSize: 18, bold: true, color: C.text2 });
      txt(s, d, { x: 3.2, y, w: 5.4, h: 0.78, valign: 'middle', fontSize: 16 });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.1, y: 1.7, w: 3.64, h: 4.6, fill: { color: C.accent1 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Übungszeit 2' });
    s.addImage({ data: await icon('LuTimer', HEX.navy), x: 9.4, y: 2.0, w: 0.7, h: 0.7, objectName: 'Symbol Zeit 2' });
    txt(s, '25', { x: 9.4, y: 2.85, w: 2.8, h: 1.3, fontFace: THEME.headFontFace, fontSize: 72, bold: true, color: C.text2 });
    txt(s, 'Minuten: Steckbrief ausfüllen und den KI-Schritt einmal ausprobieren', { x: 9.4, y: 4.25, w: 3.1, h: 1.6, fontSize: 16, color: C.text2 });
    s.addNotes('Steckbrief als Handout austeilen. Die Steckbriefe einsammeln oder fotografieren (mit Einverständnis): Sie sind die Grundlage für den Kurzbericht und zeigen der Geschäftsführung, wo die größten Hebel liegen.');
  }

  // ======================================================================
  // Abschluss
  // ======================================================================
  pres.addSection({ title: 'Abschluss' });
  {
    const s = content('Abschluss', 'So geht es weiter');
    const schritte = [
      ['LuMonitor', 'Lernplattform', 'Inhalte vertiefen, Vorlagen herunterladen'],
      ['LuListChecks', 'Abschlusstest', 'Kurzer Test innerhalb von 14 Tagen'],
      ['LuAward', 'Nachweis', 'Nachweis KI-Kompetenz je Person für Ihre Unterlagen'],
      ['LuFileBarChart', 'Kurzbericht', 'Die 3 bis 5 größten Hebel für die Geschäftsführung'],
    ];
    const w = 2.75, gap = 0.38;
    for (let i = 0; i < 4; i++) {
      const [ic, t, d] = schritte[i];
      const x = 0.6 + i * (w + gap);
      await card(s, { x, y: 1.7, w, h: 3.0, ic, title: t, text: d });
      if (i < 3) s.addImage({ data: await icon('LuArrowRight', HEX.gold), x: x + w + 0.04, y: 3.05, w: 0.3, h: 0.3, objectName: 'Pfeil' });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.0, w: 5.95, h: 1.4, fill: { color: C.text2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Option Umsetzung' });
    txt(s, [{ text: 'Umsetzung: ', options: { bold: true, color: C.accent1 } }, { text: 'Ich setze die Hebel aus dem Kurzbericht gemeinsam mit Ihnen um.', options: { color: C.background1 } }], { x: 0.9, y: 5.0, w: 5.4, h: 1.4, valign: 'middle', fontSize: 16 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.79, y: 5.0, w: 5.95, h: 1.4, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Option Weiterbildung' });
    txt(s, [{ text: 'Weiterbildung: ', options: { bold: true, color: C.text2 } }, { text: 'Für Schlüsselpersonen gibt es längere, förderfähige Programme über einen Partner.' }], { x: 7.09, y: 5.0, w: 5.4, h: 1.4, valign: 'middle', fontSize: 16 });
    s.addNotes('Zugangsdaten zur Lernplattform verteilen und das Datum für den Abschlusstest nennen. Zur Weiterbildung: Förderung nur in Aussicht stellen, nicht zusagen. Ob und wie viel gefördert wird, entscheidet die Agentur für Arbeit.');
  }
  {
    const s = content('Abschluss', 'Teil 2: Bordmittel-Check mit der Geschäftsführung');
    const fragen = [
      ['LuLayers', 'Was ist schon da?', 'Microsoft 365, ERP, Branchensoftware, Telefonanlage: Wo steckt bereits KI oder Automatisierung drin, die niemand nutzt?'],
      ['LuClock', 'Wo geht Zeit verloren?', 'Die Steckbriefe aus Übung 2 und die Erfahrungen der Geschäftsführung zusammenführen.'],
      ['LuShieldCheck', 'Was ist geregelt?', 'Freigegebene Werkzeuge, Datenschutz, Verantwortlichkeiten, Dokumentation der Schulung.'],
      ['LuRocket', 'Was zuerst?', 'Die 3 bis 5 Hebel mit dem besten Verhältnis von Aufwand und Nutzen.'],
    ];
    for (let i = 0; i < 4; i++) {
      const [ic, t, d] = fragen[i];
      const x = i % 2 === 0 ? 0.6 : 6.79;
      const y = i < 2 ? 1.7 : 4.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.95, h: 2.2, fill: { color: C.background2 }, line: { type: 'none' }, rectRadius: 0.08, objectName: 'Check ' + t });
      s.addImage({ data: await icon(ic, HEX.navy), x: x + 0.3, y: y + 0.3, w: 0.55, h: 0.55, objectName: 'Symbol ' + t });
      txt(s, t, { x: x + 1.1, y: y + 0.3, w: 4.6, h: 0.55, valign: 'middle', fontFace: THEME.headFontFace, fontSize: 20, bold: true, color: C.text2 });
      txt(s, d, { x: x + 0.3, y: y + 1.0, w: 5.35, h: 1.1, fontSize: 15 });
    }
    s.addNotes('Ca. 50 Minuten, nur mit der Geschäftsführung (ggf. plus IT-Ansprechperson). Ergebnis ist die Grundlage für den Kurzbericht, der innerhalb von 5 Werktagen folgt. Hier auch klären: Gibt es Mitarbeitende, die tiefer einsteigen sollen? Dann auf die geförderte Weiterbildung über den Partner hinweisen.');
  }
  {
    const s = pres.addSlide({ masterName: 'Abschluss', sectionTitle: 'Abschluss' });
    s.addImage({ path: logoQuer, x: 0.8, y: 0.8, w: 3.2, h: 3.2 * (342 / 1299), objectName: 'Logo' });
    txt(s, 'Vielen Dank', { x: 0.8, y: 2.3, w: 7, h: 1.0, fontFace: THEME.headFontFace, fontSize: 44, color: C.background1 });
    txt(s, 'Fragen, Ideen, nächste Schritte: Ich bin für Sie da.', { x: 0.8, y: 3.35, w: 7, h: 0.5, fontSize: 18, color: C.accent1 });
    const kontakt = [
      ['LuUser', 'Eduard Schmal'],
      ['LuMail', 'es@schmal-unternehmensberatung.de'],
      ['LuPhone', '+49 1522 685 95 24'],
      ['LuGlobe', 'schmal-unternehmensberatung.de'],
    ];
    for (let i = 0; i < kontakt.length; i++) {
      const [ic, t] = kontakt[i];
      const y = 4.3 + i * 0.6;
      s.addImage({ data: await icon(ic, HEX.gold), x: 0.8, y: y + 0.05, w: 0.35, h: 0.35, objectName: 'Symbol ' + t });
      txt(s, t, { x: 1.35, y, w: 6, h: 0.45, valign: 'middle', fontSize: 17, color: C.background1 });
    }
    s.addImage({ data: qrPng, x: 9.4, y: 2.3, w: 2.9, h: 2.9, altText: 'QR-Code: Kontakt speichern', objectName: 'QR-Code' });
    txt(s, 'Kontakt speichern', { x: 9.4, y: 5.35, w: 2.9, h: 0.4, align: 'center', fontSize: 14, color: C.accent1 });
    s.addNotes('Feedback einholen: Was nehmen Sie mit? Auf die Ziele vom Anfang zurückkommen. Hinweis auf Abschlusstest und Nachweis wiederholen.');
  }

  await pres.writeFile({ fileName: OUT });
  if (SKILL) {
    const { applyTheme } = require(path.join(SKILL, 'scripts/apply_theme.js'));
    await applyTheme(OUT, THEME);
  }
  console.log('geschrieben:', OUT);
})();
