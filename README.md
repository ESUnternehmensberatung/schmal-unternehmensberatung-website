# Homepage Unternehmensberatung – Neuaufbau

Kompletter Neuaufbau (Inhalt und Optik) auf Basis der geschärften Positionierung.
Einzige feste Vorgabe: die Markenfarben.

## Ablauf

| Schritt | Ergebnis | Status |
|---|---|---|
| 1. Fragebogen Runden 1–3 | `briefing/fragebogen.md` | erledigt |
| 2. Positionierung → Seitenstruktur, Hero, Leistungen | `briefing/seitenstruktur.md`, `hero.md`, `leistungen.md`, `profil.md` | Entwurf v3 |
| 3. Design-Signale, FAQ, Technik | `briefing/design-signale.md`, `faq.md`, `technik.md` | Entwurf |
| 4. Produktgrundlage für Impeccable | `PRODUCT.md` | offen |
| 5. Design-Richtung (Taste) | `DESIGN.md` | offen |
| 6. Umsetzung mit Skill `homepage-design` | Code | offen |

## In ein anderes Repo übernehmen

Dieser Branch hat eine eigene Historie (orphan) und enthält nur Homepage-Dateien:

```bash
git remote add briefing https://github.com/ESUnternehmensberatung/Skills
git fetch briefing homepage-briefing
git merge --allow-unrelated-histories briefing/homepage-briefing
```

Oder einfach den Inhalt des Branches in das Ziel-Repo kopieren.
