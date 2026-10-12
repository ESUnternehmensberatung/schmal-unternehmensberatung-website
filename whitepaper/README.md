# Whitepaper

„5 Gründe, warum die meisten KI-Projekte bei Mittelständlern und Freiberuflern scheitern“

- Quelle: `whitepaper.html` (Design der Website, Schriften aus `node_modules`)
- Fertiges PDF: `Whitepaper-KI-Projekte-Schmal-Unternehmensberatung.pdf`
- Das PDF liegt bewusst **nicht** in `public/`, damit es nur über das Jotform-Formular
  verschickt wird (Autoresponder mit Anhang).

Neu erzeugen:

```bash
NPM_ROOT=$(npm root -g) node whitepaper/build.cjs
python3 whitepaper/merge.py   # benötigt: pip install pymupdf
```
