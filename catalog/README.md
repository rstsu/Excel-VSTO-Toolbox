# Interaktiver Demo-Katalog

Der Katalog wird bei passenden Änderungen und bei neuen Releases automatisch über GitHub Pages veröffentlicht.

Die Quelldaten stammen direkt aus:

- `RegExPQ/DemoCatalog.PowerQuery.vb`
- `RegExPQ/DemoCatalog.Regex.vb`
- `RegExPQ/DemoCatalog.Vba.vb`
- `RegExPQ/DemoCatalog.Formula.vb`
- `RegExPQ/DemoRunner.vb`
- `RegExPQ/Demos/*.zip`

`catalog/scripts/import-vsto-catalog.js` erzeugt daraus `catalog/site/demo-data.js`. Als neu gelten automatisch die seit dem vorherigen Release hinzugekommenen Demo-IDs. Vor einem neuen Release wird dafür mit dem letzten vorhandenen Release-Tag verglichen; nach dem Release bleiben dessen neue Einträge markiert. ZIP-Zuordnungen werden aus `DemoRunner.vb` übernommen.

Der Workflow `.github/workflows/demo-catalog-pages.yml` prüft den erzeugten JavaScript-Code und veröffentlicht anschließend `catalog/site`.

## Lokal aktualisieren

```powershell
node catalog/scripts/import-vsto-catalog.js . catalog/site/demo-data.js v1.0.2.20
```
