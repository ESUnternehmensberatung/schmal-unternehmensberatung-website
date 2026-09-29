# Homepage Unternehmensberatung – Neuaufbau

Kompletter Neuaufbau (Inhalt und Optik) auf Basis der geschärften Positionierung.
Einzige feste Vorgabe: die Markenfarben.

## Ablauf

| Schritt | Ergebnis | Status |
|---|---|---|
| 1. Fragebogen Runde 1 – Positionierung | `briefing/fragebogen.md` | erledigt |
| 2. Positionierung → Seitenstruktur + Hero-Botschaft | `briefing/seitenstruktur.md` | Entwurf v1 |
| 3. Fragebogen Runde 2 – Beweise, Runde 3 – Design-Signale | `briefing/fragebogen.md` | offen |
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
