# Homepage Schmal Unternehmensberatung

OnePager + Blog, gebaut mit [Astro](https://astro.build) auf Basis des Briefings
in `briefing/`, `PRODUCT.md` und `DESIGN.md`.

## Lokal starten

```bash
npm install
npm run dev      # Entwicklung auf http://localhost:4321
npm run build    # statische Seite nach dist/
```

## Inhalte pflegen

| Was | Wo |
|---|---|
| Links für Terminbuchung, Angebot, Whitepaper, Google-Profil, Kontakt | `src/config/site.ts` |
| Texte der Startseite (je Sektion) | `src/components/sections/*.astro` |
| FAQ | `src/content/faq/*.md` (eine Datei pro Frage) |
| Google-Bewertungen | `src/data/bewertungen.ts` (leer = Sektion ausgeblendet) |
| Blogartikel (Business-Concierge) | `src/content/blog/*.md` – Vorlage in `README.txt` |
| Impressum, Datenschutz | `src/content/rechtliches/*.md` |
| Fotos | `assets/hero.png`, `assets/about.png` (werden beim Build als AVIF/WebP optimiert) |

Die Buttons öffnen vorerst eine vorbereitete E-Mail an
es@schmal-unternehmensberatung.de. Sobald Jotform und die Google-Terminbuchung
stehen, nur die URLs in `src/config/site.ts` ersetzen.

## Veröffentlichen

`.github/workflows/deploy.yml` baut die Seite bei jedem Push auf `main` und
veröffentlicht sie über GitHub Pages (Repository → Settings → Pages → Source:
**GitHub Actions**, nicht „Deploy from a branch“).

Ohne eigene Domain liegt die Seite unter
`https://esunternehmensberatung.github.io/schmal-unternehmensberatung-website/`,
der Workflow setzt die Pfade automatisch passend. Für die eigene Domain unter
Settings → Pages → Custom domain `schmal-unternehmensberatung.de` eintragen und
den DNS-Eintrag umstellen; beim nächsten Deploy stimmen die Pfade automatisch.

## Ablauf

| Schritt | Ergebnis | Status |
|---|---|---|
| 1. Fragebogen Runden 1–3 | `briefing/fragebogen.md` | erledigt |
| 2. Positionierung → Seitenstruktur, Hero, Leistungen | `briefing/*.md` | fertig |
| 3. Design-Signale, FAQ, Technik | `briefing/*.md` | fertig |
| 4. Produktgrundlage für Impeccable | `PRODUCT.md` | fertig |
| 5. Design-Richtung (Taste) | `DESIGN.md` | fertig |
| 6. Umsetzung mit Skill `homepage-design` | `src/` | gebaut |

Herkunft: Briefing aus `ESUnternehmensberatung/Skills`, Branch `homepage-briefing` (Historie übernommen).
