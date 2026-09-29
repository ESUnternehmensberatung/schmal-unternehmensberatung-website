# Design

Status: **gebaut** (Astro, `src/`). Dieses Dokument beschreibt das umgesetzte
System. Änderungen am Look zuerst hier festhalten, dann im Code.

## Design Read
Persuasive OnePager für Inhaber familiengeführter KMU (45+), ruhig-editorial,
Schweizer Ordnung, klassische Serif-Headlines, echte Fotografie, harte Zahlen.

Dials (Taste): Variance 5 · Motion 4 · Density 4. Keine Design-Library, natives
CSS mit Tokens (Aesthetic „Editorial“, kein offizielles System).

## Farben (`src/styles/global.css`)

| Token | Hex | Einsatz |
|---|---|---|
| `--navy` | `#1A3A5C` | Headlines auf hell, Primär-Link, Outline-Button |
| `--navy-deep` | `#0F2438` | Header, Hero-Overlay, Ablauf, Abschluss-CTA, Footer |
| `--gold` | `#C9A87C` | Primär-Button, Kennzahlen, Akzentwort – nur auf dunkel |
| `--gold-deep` | `#7A5C33` | Gold für Text/Links/Icons auf hell |
| `--ivory` | `#F7F4EE` | Seitenhintergrund |
| `--white` | `#FFFFFF` | Leistungen, Ergebnisse, Karten |
| `--ink` | `#1A2330` | Fließtext |
| `--stone` | `#5A6572` | Sekundärtext |
| `--line` | `#E3DDD2` | Trennlinien auf hell |
| `--on-dark` / `--on-dark-muted` | `#EEF1F5` / `#B4C0CF` | Text auf Navy-Deep |
| `--line-dark` | Gold 28 % | Trennlinien auf dunkel |

Regeln: Primär-Button Gold mit Navy-Deep-Text (7,1 : 1). Gold nie als Text auf
hell. Einziger Verlauf: Foto-Overlay im Hero. Nur Light-Theme (Markenvorgabe);
dunkle Sektionen sind feste Navy-Deep-Flächen, kein Theme-Wechsel.

## Typografie
- Headlines: **Source Serif 4** (variabel, optische Größen), Gewicht 500,
  `text-wrap: balance`. Akzentwort kursiv in Gold – nur im Hero und im
  Abschluss-CTA.
- Text: **Hanken Grotesk** (variabel), Grundgröße 19 px, Zeilenhöhe 1,6.
- Beide selbst gehostet über `@fontsource-variable`.
- Skala: `--step--1` 15 px … `--step-4` bis 72 px (fluid via `clamp`).
- Zahlen mit geschütztem Leerzeichen vor `%`.

## Layout
- Container max. 1320 px, Gutter `clamp(1.25rem, 4vw, 3rem)`, Sektionsabstand
  `clamp(5rem, 9vw, 9rem)`.
- Ecken: einheitlich 2 px (`--radius`). Karten nur für Business-Concierge.
- Sektionen (Leseweg aus `briefing/seitenstruktur.md`), jede mit eigener Form:
  1. Hero – Vollbild-Foto, Overlay links, Kennzahlenleiste unten
  2. Die Lage – großes Zitat-Statement mit Quellenfußnote, 2-spaltige Schmerz→System-Liste
  3. Leistungen – nummerierte Zeilen (5/7-Raster) mit Beleg-Box
  4. Ablauf – dunkel, 5-Schritte-Zeitleiste (mobil vertikal)
  5. Über mich – Porträt sticky links, Prinzipien, Change-Zitat, Qualifikationen
  6. Ergebnisse – Methodenleiste, 3 Vorher/Nachher-Spalten, 2 Zusatzzahlen,
     Google-Bewertungen (erscheinen automatisch, sobald eingetragen)
  7. Business-Concierge – Text + Empfehlungskarte mit Provisionshinweis
  8. FAQ – sticky Überschrift + Accordion (`details`)
  9. Whitepaper – schmales Band
  10. Abschluss-CTA – dunkel, beide Handlungen + Direktkontakt
- Mobil: alle Raster einspaltig ab 900–960 px, Hero-Foto oben (4 : 3).

## Komponenten
- Buttons: `.btn-primary` (Gold), `.btn-ghost` (auf dunkel), `.btn-outline`
  (auf hell), `.btn-navy`. Ein Label pro Absicht: „Quick-Win-Analyse anfragen“,
  „Angebot anfordern“, „Whitepaper vormerken“.
- Icons: Phosphor (regular, `src/icons`), eine Familie.
- Fokus: 2 px Gold-Deep-Outline (auf dunkel Gold).

## Bewegung (Emil Kowalski)
- Easing `--ease-out: cubic-bezier(0.23, 1, 0.32, 1)`.
- Hero beim Laden: Foto setzt sich (Scale 1,045 → 1, 1,6 s), Text steigt
  gestaffelt (80 ms) auf – einmaliger Moment.
- Scroll-Reveal: `[data-reveal]`, 18 px / 700 ms, Stagger 70 ms, per
  IntersectionObserver; ohne JS sofort sichtbar.
- Buttons: `:active` scale 0,97 (160 ms), Hover nur bei `hover: hover`.
- FAQ: Plus dreht auf 45° (250 ms), Antwort blendet ein. Mobilmenü 200 ms.
- `prefers-reduced-motion`: alle Bewegungen aus.

## Verboten
Laut, verspielt, billig, AI-Slop: keine Lila-Verläufe, kein Glassmorphism,
keine drei gleichen Icon-Karten, keine Buzzword-Überschriften, keine
Platzhalterzahlen, keine Gedankenstriche im Seitentext.
