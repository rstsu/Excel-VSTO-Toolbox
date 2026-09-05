/* Automatisch aus dem VSTO-Katalog erzeugt.
 * Quelle: https://github.com/rstsu/Excel-VSTO-Toolbox
 * Erneut erzeugen: node scripts/import-vsto-catalog.js <Repository-Ordner> dist/demo-data.js <Version>
 */
window.CATALOG_META = {
  "version": "v1.0.2.20",
  "sourceCommit": "998335f38db2ea37c90ad350a4c5b9a7eb17fe68",
  "repositoryUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox",
  "counts": {
    "Power Query": 21,
    "Regex": 21,
    "VBA": 11,
    "Formeln": 11
  },
  "packageDemoCount": 26,
  "uniquePackageCount": 21
};

window.DEMO_CATALOG = [
  {
    "id": "PQ-001",
    "sourceId": "pq_001",
    "title": "Wörter je Nummer gruppieren",
    "category": "Power Query",
    "summary": "Text/Wörter werden je gleicher Nummer in einer Spalte gruppiert bzw. mit Leerzeichen zusammengefasst. Also Spalte A sind Nummern und Spalte B Wörter. Sind die Nummern in Spalte A gleich, …",
    "description": "Text/Wörter werden je gleicher Nummer in einer Spalte gruppiert bzw. mit Leerzeichen zusammengefasst.\nAlso Spalte A sind Nummern und Spalte B Wörter. Sind die Nummern in Spalte A gleich, werden die Wörter gruppiert.\n\nNach Änderungen in der Grundtabelle wird die Abfrage mit STRG+ALT+F5 aktualisiert!",
    "tags": [
      "regex",
      "m-code",
      "gruppieren",
      "table.group",
      "text",
      "text.combine"
    ],
    "codeText": "let\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_1\"]}[Content],\n    Gruppe = Table.Group(Quelle, {\"Nr\"}, {{\"Wörter\", each _, type table [Nr=nullable number, Wörter=nullable text]}}),\n    HinzuSpalte = Table.AddColumn(Gruppe, \"Alle\", each Text.Combine(List.Distinct(List.Transform([Wörter][Wörter], each Text.Trim(_))), \" \"))\nin\n    HinzuSpalte",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L8"
  },
  {
    "id": "PQ-002",
    "sourceId": "pq_002",
    "title": "Zollangaben aus Text auslesen",
    "category": "Power Query",
    "summary": "In den Zellen A2:Ax sind Texte mit mehreren Zollangaben (z. B. 1/2\" oder dann auch 1 3/4\"). Die Zollangaben sollen ausgelesen werden. Auch mehrere. Überschriften (Z1, Z2...) automatisch …",
    "description": "In den Zellen A2:Ax sind Texte mit mehreren Zollangaben (z. B. 1/2\" oder dann auch 1 3/4\").\nDie Zollangaben sollen ausgelesen werden. Auch mehrere.\nÜberschriften (Z1, Z2...) automatisch generieren.\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.",
    "tags": [
      "regex",
      "m-code",
      "zahl",
      "zoll",
      "text",
      "extrahieren"
    ],
    "codeText": "let\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_2\"]}[Content],\n    AddZollListe = Table.AddColumn(Quelle, \"ZollListe\", each\n        let\n            v = [Beschreibung],\n            protected =\n                List.Accumulate(\n                    {\"1 1/2\"\"\",\"2 1/2\"\"\",\"3 1/2\"\"\",\"4 1/2\"\"\",\"5 1/2\"\"\",\"6 1/2\"\"\",\"7 1/2\"\"\",\"8 1/2\"\"\",\"9 1/2\"\"\"},\n                    v,\n                    (w,x) => Text.Replace(w, x, Text.Replace(x, \" \", \"~\"))\n                ),\n            y = Text.SplitAny(protected, \" ,.-:;()[]\"),\n            z = List.Transform(\n                List.Select(y,\n                    each Text.EndsWith(_, \"\"\"\") and Text.Length(Text.Select(_, {\"0\"..\"9\"})) > 0\n                ),\n                each Text.Replace(_, \"~\", \" \")\n            )\n        in\n            z),\n    MaxAnzahl = List.Max(List.Transform(AddZollListe[ZollListe], each List.Count(_))),\n    ZNames = List.Transform({1..MaxAnzahl}, each \"Z\" & Text.From(_)),\n    AddRecord = Table.AddColumn(AddZollListe, \"Z\", each Record.FromList(List.FirstN([ZollListe] & List.Repeat({null}, MaxAnzahl), MaxAnzahl), ZNames)),\n    Expand = Table.ExpandRecordColumn(AddRecord, \"Z\", ZNames, ZNames),\n    Erg = Table.SelectColumns(Expand, List.Transform(ZNames, each Text.From(_)))\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L36"
  },
  {
    "id": "PQ-003",
    "sourceId": "pq_003",
    "title": "Inhalt des TEMP-Ordners auflisten",
    "category": "Power Query",
    "summary": "Der Inhalt des lokalen TEMP-Ordners... (C:\\Users\\USERNAME\\AppData\\Local\\Temp) ...wird aufgelistet. Der USERNAME wird über eine Formel ausgelesen. …",
    "description": "Der Inhalt des lokalen TEMP-Ordners...\n(C:\\Users\\USERNAME\\AppData\\Local\\Temp)\n...wird aufgelistet.\nDer USERNAME wird über eine Formel ausgelesen.\n=TEXTVOR(ZELLE(\"filename\");\"\\\";3)&\"\\AppData\\Local\\Temp\"\n!!!!!!!!WICHTIG!!!!!!!!\nDie Datei, in der das probiert wird MUSS gespeichert sein.\nIn einer neuen UNGESPEICHERTEN Datei geht das NICHT!\nDa erscheint dann der Fehler #WERT! in Zelle G1.\n!!!!!!!!WICHTIG!!!!!!!!\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.",
    "tags": [
      "regex",
      "m-code",
      "tmp",
      "temp",
      "text",
      "auflisten"
    ],
    "codeText": "let\n    Quelle = Folder.Files(Excel.CurrentWorkbook(){[Name=\"Benutzername\"]}[Content]{0}[Column1]),\n    Hinzu = Table.AddColumn(Quelle, \"Datei\", each [Folder Path]&[Name]),\n    Erg = Table.SelectColumns(Hinzu,{\"Datei\", \"Date accessed\", \"Date modified\", \"Date created\"})\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L86"
  },
  {
    "id": "PQ-004",
    "sourceId": "pq_004",
    "title": "Sonderzeichen entfernen",
    "category": "Power Query",
    "summary": "In A2:A8 stehen Texte mit Sonderzeichen - auch mehrere hintereinander. Diese sollen entfernt werden. Zwischen den Wörtern darf nur ein Leerzeichen übrig bleiben.",
    "description": "In A2:A8 stehen Texte mit Sonderzeichen - auch mehrere hintereinander.\nDiese sollen entfernt werden.\nZwischen den Wörtern darf nur ein Leerzeichen übrig bleiben.\n\nFalls es nicht nur Leerzeichen sind, könnte man auch das nehmen:\nClean = Text.Combine(List.Select(Text.Split(DelS, \" \"), each _ <> \"\"),\n\nEs würde auch über eine Funktion gehen (mit Namen \"fncCleanText\"):\n(Text as text, optional RemoveChars as list) =>\nlet\n    Chars = if RemoveChars = null then {\":\",\"?\",\"*\",\"/\",\"\\\"} else RemoveChars,\n    Clean = List.Accumulate(Chars, Text, (txt, c) => Text.Replace(txt, c, \"\")),\n    Result = Text.Combine(List.RemoveItems(Text.Split(Clean, \" \"), {\"\"}), \" \")\nin\n    Result\n\nAufruf:\nfncCleanText([Text])\n\nOder mit eigenen Zeichen:\nfncCleanText([Text], {\":\",\"?\",\"*\",\"/\",\"\\\",\".\",\",\",\";\"})\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.",
    "tags": [
      "sonderzeichen",
      "m-code",
      "entfernen",
      "mehrere",
      "text"
    ],
    "codeText": "let\n    Erg =\n        Table.TransformColumns(\n            Excel.CurrentWorkbook(){[Name=\"Demo_PQ_4\"]}[Content],\n            {\n                {\"Text\", each\n                    let\n                        DelS = List.Accumulate(\n                            {\":\",\"?\",\"*\",\"/\",\"\\\"},_,\n                            (txt, char) => Text.Replace(txt, char, \"\")\n                        ),\n                        Clean = Text.Combine(List.RemoveItems(Text.Split(DelS, \" \"), {\"\"}), \" \")\n                    in\n                        Clean\n                }\n            }\n        )\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L122"
  },
  {
    "id": "PQ-005",
    "sourceId": "pq_005",
    "title": "Kreuztabelle aus Liste erste 3 Buchstaben",
    "category": "Power Query",
    "summary": "Aus einer Liste (A2:A15) wird eine Kreuztabelle erstellt. Grundlage sind die ersten 3 gleichen Buchstaben. Es ist auch in Formeln gelöst. Mit der gleichen Bezeichnung. Für Informationen, …",
    "description": "Aus einer Liste (A2:A15) wird eine Kreuztabelle erstellt.\nGrundlage sind die ersten 3 gleichen Buchstaben.\nEs ist auch in Formeln gelöst. Mit der gleichen Bezeichnung.\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.\n\nNach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!",
    "tags": [
      "power query",
      "text",
      "m-code",
      "tabelle",
      "liste"
    ],
    "codeText": "let\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_5\"]}[Content],\n    Kuerzel = Table.AddColumn(Quelle, \"Kürzel\", each Text.Start([Wert], 3), type text),\n    Gruppe = Table.Group(Kuerzel, {\"Kürzel\"}, {{\"Daten\", each Table.AddIndexColumn(_, \"Pos\", 1, 1, Int64.Type)}}),\n    Expand = Table.ExpandTableColumn(Gruppe, \"Daten\", {\"Wert\", \"Pos\"}),\n    Pos = Table.TransformColumns(Expand, {{\"Pos\", each \"Pos\" & Text.From(_), type text}}),\n    Pivot = Table.Pivot(Pos, List.Distinct(Pos[Pos]), \"Pos\", \"Wert\"),\n    Erg = Table.RemoveColumns(Pivot, {\"Kürzel\"})\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L182"
  },
  {
    "id": "PQ-006",
    "sourceId": "pq_006",
    "title": "String zwischen 2 Zahlen auslesen",
    "category": "Power Query",
    "summary": "Aus einer Liste (A2:A12) wird ein Text zwischen 2 Zahlem ausgelesen. Z. B. \"1. Vom Bodensee in den Schwarzwald 21 May 2001\". Es wird nur der \"Titel\" in der Mitte ausgelesen bzw. der Rest …",
    "description": "Aus einer Liste (A2:A12) wird ein Text zwischen 2 Zahlem ausgelesen.\nZ. B. \"1. Vom Bodensee in den Schwarzwald 21 May 2001\".\nEs wird nur der \"Titel\" in der Mitte ausgelesen bzw. der Rest ersetzt.\n\nNach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!",
    "tags": [
      "power query",
      "text",
      "m-code",
      "tabelle",
      "liste"
    ],
    "codeText": "let\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_6\"]}[Content],\n    Ausgabe = Table.AddColumn(Quelle, \"Ausgabe\", each\n        let\n            ON = Text.AfterDelimiter([Daten], \". \"),\n            TE = Text.Split(Text.Trim(ON), \" \"),\n            TI = Text.Combine(List.RemoveLastN(TE, 3), \" \")\n        in\n            TI,\n        type text\n    ),\n    Erg = Table.SelectColumns(Ausgabe, {\"Ausgabe\"})\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L217"
  },
  {
    "id": "PQ-007",
    "sourceId": "pq_007",
    "title": "Datumbrereich in Liste auflösen",
    "category": "Power Query",
    "summary": "Aus einer Liste (A2:A6 - Ferientermine 2027 in Baden Württemberg) wird eine laufende Liste erstellt.",
    "description": "Aus einer Liste (A2:A6 - Ferientermine 2027 in Baden Württemberg) wird eine laufende Liste erstellt.\n\nDer auskommentierte Code ist eine andere Herangehensweise ( also alles zwischen /* ...M-Code... */).\n\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.\n\nNach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!",
    "tags": [
      "power query",
      "datum",
      "m-code",
      "tabelle",
      "ferien"
    ],
    "codeText": "/*\nlet\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_7\"]}[Content],\n    Typen = Table.TransformColumnTypes(\n        Quelle,\n        {{\"Ferien BW\", type text}, {\"Von\", type date}, {\"Bis\", type date}}\n    ),\n    Erg = Table.Sort(\n        Table.SelectColumns(\n            Table.AddColumn(\n                Table.ExpandListColumn(\n                    Table.AddColumn(\n                        Typen,\n                        \"Datum\",\n                        each List.Dates([Von], Duration.Days([Bis] - [Von]) + 1, #duration(1,0,0,0)),\n                        type list\n                    ),\n                    \"Datum\"\n                ),\n                \"Tag\",\n                each Date.DayOfWeekName([Datum]),\n                type text\n            ),\n            {\"Datum\", \"Tag\"}\n        ),\n        {{\"Datum\", Order.Ascending}}\n    )\nin\n    Erg\n*/\nlet\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_7\"]}[Content],\n    TypeG = Table.TransformColumnTypes(Quelle,{{\"Ferien BW\", type text}, {\"Von\", type date}, {\"Bis\", type date}}),\n    ErwS = Table.AddColumn(TypeG, \"Datum\", each List.Dates([Von], Duration.Days([Bis] - [Von]) + 1, #duration(1,0,0,0))),\n    ExpandT = Table.TransformColumnTypes(Table.ExpandListColumn(ErwS, \"Datum\"),{{\"Datum\", type date}}),\n    TagN = Table.AddColumn(ExpandT, \"Tag\", each Date.DayOfWeekName([Datum]), type text),\n    Erg = Table.Sort(Table.SelectColumns(TagN,{\"Datum\", \"Tag\"}), {{\"Datum\", Order.Ascending}})\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L254"
  },
  {
    "id": "PQ-008",
    "sourceId": "pq_008",
    "title": "Kreuztabelle aus Liste Materialnummern",
    "category": "Power Query",
    "summary": "Aus einer Liste (A2:E11) wird eine Kreuztabelle erstellt. Grundlage sind die Materialnummern in der ersten Spalte. Spalten (Text) werden so viel wie nötig erstellt. Für Informationen, …",
    "description": "Aus einer Liste (A2:E11) wird eine Kreuztabelle erstellt.\nGrundlage sind die Materialnummern in der ersten Spalte.\nSpalten (Text) werden so viel wie nötig erstellt.\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.\n\nNach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!",
    "tags": [
      "power query",
      "material",
      "m-code",
      "tabelle",
      "kreuztabelle"
    ],
    "codeText": "let\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_8\"]}[Content],\n    Gruppe = Table.Group(Quelle, {\"Materialnummer\"}, {\"Rows\", each Table.AddIndexColumn(_, \"Zeile\", 1)}),\n    Expand = Table.ExpandTableColumn(Gruppe, \"Rows\", {\"Werk\",\"Materialart\",\"TextID\",\"Text\",\"Zeile\"}),\n    TypG = Table.TransformColumnTypes(Expand, {{\"Zeile\", type text}}),\n    Pivot = Table.Pivot(TypG, List.Distinct(TypG[Zeile]), \"Zeile\", \"Text\"),\n    Erg = Table.RenameColumns(Pivot, List.Transform(List.Distinct(TypG[Zeile]), each {_, \"Text\" & _}))\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L319"
  },
  {
    "id": "PQ-009",
    "sourceId": "pq_009",
    "title": "Auflistung aus Gruppe - Von Bis erstellen",
    "category": "Power Query",
    "summary": "Aus einer Liste (A2:C12) wird eine Auflistung erstellt. Grundlage ist die Gruppe in Spalte A. Dann kommt Von und Bis. Berücksichtigt wird, ob Text statt Zahl eingegeben ist und ob es …",
    "description": "Aus einer Liste (A2:C12) wird eine Auflistung erstellt.\nGrundlage ist die Gruppe in Spalte A. Dann kommt Von und Bis.\nBerücksichtigt wird, ob Text statt Zahl eingegeben ist und ob es eine leere Eingabe ist.\nAuch wird darauf geachtet, ob Von größer Bis ist.\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.\n\nNach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!\n\nMan kann auch eine eigene Funktion erstellen (Name der Funktion: fncGetWert):\nlet\n    Quelle = (g as any, von as any, bis as any) as nullable list =>\nlet\n    VonZahl = try Int64.From(von) otherwise null,\n    BisZahl = try Int64.From(bis) otherwise null,\n    VonIstText = von <> null and VonZahl = null,\n    BisIstText = bis <> null and BisZahl = null\nin\n    if g = null then\n        null\n    else if VonIstText and BisZahl <> null then {BisZahl}\n    else if BisIstText and VonZahl <> null then {VonZahl}\n    else if VonZahl <> null and BisZahl = null then {1..VonZahl}\n    else if VonZahl <> null and BisZahl <> null then\n        if VonZahl <= BisZahl\n        then {VonZahl..BisZahl}\n        else {BisZahl..VonZahl}\n    else\n        null\nin\n    Quelle\n\nDann eine Abfrage (Name: tblErgfnc):\nlet\n    Quelle = Table.Sort(Excel.CurrentWorkbook(){[Name=\"Demo_PQ_9\"]}[Content],{{\"Gruppe\", Order.Ascending}}),\n    OhneDuplikate = Table.Distinct(Quelle, {\"Gruppe\",\"Von\",\"Bis\"}),\n    MitListe = Table.AddColumn(OhneDuplikate, \"Wert\", each fncGetWert([Gruppe],[Von],[Bis]), type list),\n    Gefiltert = Table.SelectRows(MitListe, each [Gruppe] <> null and [Wert] <> null),\n    Erg = Table.SelectColumns(Table.ExpandListColumn(Gefiltert, \"Wert\"), {\"Gruppe\",\"Wert\"})\nin\n    Erg",
    "tags": [
      "power query",
      "auflistung",
      "m-code",
      "von",
      "bis"
    ],
    "codeText": "let\n    // Mit der Funktion ermittle ich die Werteliste\n    GetWert = (g as any, von as any, bis as any) as nullable list =>\n        let\n            VonZahl = try Int64.From(von) otherwise null,\n            BisZahl = try Int64.From(bis) otherwise null,\n            VonIstText = von <> null and VonZahl = null,\n            BisIstText = bis <> null and BisZahl = null\n        in\n            if g = null then\n                null\n            else if VonIstText and BisZahl <> null then {BisZahl}\n            else if BisIstText and VonZahl <> null then {VonZahl}\n            else if VonZahl <> null and BisZahl = null then {1..VonZahl}\n            else if VonZahl <> null and BisZahl <> null then\n                if VonZahl <= BisZahl\n                then {VonZahl..BisZahl}\n                else {BisZahl..VonZahl}\n            else null,\n    Quelle = Table.Sort(Excel.CurrentWorkbook(){[Name=\"Demo_PQ_9\"]}[Content],{{\"Gruppe\", Order.Ascending}}),\n    OhneDuplikate = Table.Distinct(Quelle, {\"Gruppe\",\"Von\",\"Bis\"}),\n    MitListe = Table.AddColumn(OhneDuplikate, \"Wert\", each GetWert([Gruppe],[Von],[Bis]), type list),\n    Gefiltert = Table.SelectRows(MitListe, each [Gruppe] <> null and [Wert] <> null),\n    Erg = Table.SelectColumns(Table.ExpandListColumn(Gefiltert, \"Wert\"), {\"Gruppe\",\"Wert\"})\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L353"
  },
  {
    "id": "PQ-010",
    "sourceId": "pq_0010",
    "title": "Blockbildung mit Zwischensummen und Gesamtsumme",
    "category": "Power Query",
    "summary": "Aus einer Liste (A2:C16) werden Blöcke je Code mit Zwischensummen gebildet. Die Gesamtsumme kann auch über \"Tabellenentwurf - Tabellenformatoptionen - Haken bei Ergebniszeile setzen\" …",
    "description": "Aus einer Liste (A2:C16) werden Blöcke je Code mit Zwischensummen gebildet.\nDie Gesamtsumme kann auch über \"Tabellenentwurf - Tabellenformatoptionen - Haken bei Ergebniszeile setzen\" angzeigt werden.\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.\n\nNach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!\n\nMit folgendem M-Code wird die Gesamtsumme direkt am Ende ausgegeben:\n\nlet\n    Quelle = Table.TransformColumnTypes(Excel.CurrentWorkbook(){[Name=\"Demo_PQ_10\"]}[Content], {{\"Datum\", type date}, {\"Code\", Int64.Type}, {\"Betrag\", type number}}),\n    // Falls \"Erhebniszeile\" in der Grundtabelle eingeblendet ist\n    EntfF = Table.RemoveRowsWithErrors(Quelle, {\"Datum\"}),\n    Gruppe = Table.Sort(Table.Group(EntfF, {\"Code\"}, {{\"x\", each _, type table}}),{{\"Code\", Order.Ascending}}),\n    Block = List.Transform(Gruppe[x], (x) =>\n        let\n            y = List.Sum(x[Betrag]),\n            z = #table({\"Datum\", \"Code\", \"Betrag\"}, {{null, null, y}})\n        in\n            Table.Combine({x, z})\n        ),\n    // Nach Ausgabe im Tabellenblatt die Spalte Datum formatieren!\n    Erg = Table.InsertRows(Table.Combine(Block), Table.RowCount(EntfF)+Table.RowCount(Gruppe), {[Datum = null, Code = null, Betrag = List.Sum(EntfF[Betrag])]})\nin\n    Erg\n\nMan kann die Gesamtsumme auch per Formel (Teilergebnis und Bereich.Verschieben) ausgeben. Hier zwei Möglichkeiten (Tabelle ist über Beispiel Demo schon im Tabellenblatt ausgegeben und Abfrage1 wurde NICHT umbenannt):\n\n=SUMMENPRODUKT((ISTZAHL(Abfrage1[Datum]))*TEILERGEBNIS(109;BEREICH.VERSCHIEBEN(Abfrage1[Betrag];ZEILE(Abfrage1[Betrag])-MIN(ZEILE(Abfrage1[Betrag]));0;1)))\n\nUnd mit LET:\n=LET(x;Abfrage1[Datum];y;Abfrage1[Betrag];z;TEILERGEBNIS(109;BEREICH.VERSCHIEBEN(y;ZEILE(y)-MIN(ZEILE(y));0;1));SUMMENPRODUKT((x<>\"\")*z))",
    "tags": [
      "power query",
      "zwischensumme",
      "m-code",
      "gesamtsumme",
      "blockbildung"
    ],
    "codeText": "let\n    Quelle = Table.TransformColumnTypes(Excel.CurrentWorkbook(){[Name=\"Demo_PQ_10\"]}[Content], {{\"Datum\", type date}, {\"Code\", Int64.Type}, {\"Betrag\", type number}}),\n    // Falls \"Erhebniszeile\" in der Grundtabelle eingeblendet ist\n    EntfF = Table.RemoveRowsWithErrors(Quelle, {\"Datum\"}),\n    Gruppe = Table.Sort(Table.Group(EntfF, {\"Code\"}, {{\"x\", each _, type table}}),{{\"Code\", Order.Ascending}}),\n    Block = List.Transform(Gruppe[x], (x) =>\n        let\n            y = List.Sum(x[Betrag]),\n            z = #table({\"Datum\", \"Code\", \"Betrag\"}, {{null, null, y}})\n        in\n            Table.Combine({x, z})\n        ),\n    // Nach Ausgabe im Tabellenblatt die Spalte Datum formatieren!\n    Erg = Table.Combine(Block)\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L438"
  },
  {
    "id": "PQ-011",
    "sourceId": "pq_0011",
    "title": "Anmeldename bilden aus Vor- und Nachname",
    "category": "Power Query",
    "summary": "Aus einer Liste (A2:B10) werden Anmeldenamen aus Vorname und Nachname gebildet. Der Vorname nur 3 Buchstaben. Aus \"Paul Huber\" wird \"Huber.Pau\". Bei doppeltem Vor- und Nachname wird eine …",
    "description": "Aus einer Liste (A2:B10) werden Anmeldenamen aus Vorname und Nachname gebildet.\nDer Vorname nur 3 Buchstaben. Aus \"Paul Huber\" wird \"Huber.Pau\". Bei doppeltem Vor- und Nachname wird eine Zahl angehängt.\n\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.\n\nNach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!",
    "tags": [
      "power query",
      "vorname",
      "m-code",
      "nachname",
      "anmeldename"
    ],
    "codeText": "let\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_11\"]}[Content],\n    HinzuVN = Table.AddColumn(Quelle, \"Nachname.Vorname\", each [Nachname] & \".\" & Text.Start([Vorname], 3)),\n    Gruppe = Table.Group(HinzuVN, {\"Nachname.Vorname\"}, {{\"Alle Zeilen\", each Table.AddIndexColumn(_, \"Index\", 1, 1, Int64.Type), type table [Vorname=nullable text, Nachname=nullable text, Nachname.Vor=nullable text, Index=Int64.Type]}}),\n    Expand = Table.ExpandTableColumn(Gruppe, \"Alle Zeilen\", {\"Vorname\", \"Nachname\", \"Index\"}),\n    Final = Table.AddColumn(Expand, \"FinalerNachnameVor\", each if [Index] > 1 then [Nachname.Vorname] & Text.From([Index]) else [Nachname.Vorname]),\n    Erg = Table.SelectColumns(Table.RenameColumns(Final,{{\"FinalerNachnameVor\", \"NachnameVorneme3\"}}), {\"Vorname\", \"Nachname\", \"NachnameVorneme3\"})\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L504"
  },
  {
    "id": "PQ-012",
    "sourceId": "pq_0012",
    "title": "2 Wohnungen - Vermietung - nicht belegte Tage",
    "category": "Power Query",
    "summary": "Aus einer Liste (A2:C30) werden aus der Vermietung von 2 Wohnungen (Wohnung, Anreise, Abreise) die Tage abgebildet, welche NICHT belegt sind. Wie der Abreisetag gerechnet werden …",
    "description": "Aus einer Liste (A2:C30) werden aus der Vermietung von 2 Wohnungen (Wohnung, Anreise, Abreise) die Tage abgebildet, welche NICHT belegt sind.\nWie der Abreisetag gerechnet werden kann/soll, ist im Kommentar des M-Codes gezeigt.\n\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.\n\nNach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!",
    "tags": [
      "power query",
      "vermietung",
      "m-code",
      "datum",
      "belegung"
    ],
    "codeText": "let\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_12\"]}[Content],\n    TypG = Table.TransformColumnTypes(Quelle, {{\"Anreise\", type date}, {\"Abreise\", type date}, {\"Wohnung\", type text}}),\n    MinAnreise = if List.Count(TypG[Anreise]) > 0 then List.Min(TypG[Anreise]) else null,\n    MaxAbreise = if List.Count(TypG[Abreise]) > 0 then List.Max(TypG[Abreise]) else null,\n    DayB = if MinAnreise <> null and MaxAbreise <> null then Duration.Days(MaxAbreise - MinAnreise) else 0,\n    DateT = if MinAnreise <> null then Table.FromList(List.Dates(MinAnreise, DayB + 1, #duration(1, 0, 0, 0)), Splitter.SplitByNothing(), {\"Nicht Belegt\"}, null, ExtraValues.Error) else #table({\"Nicht Belegt\"}, {}),\n    Whg1 = Table.AddColumn(DateT, \"Whg1\", each if Table.RowCount(Table.SelectRows(TypG, (row) => row[Wohnung] = \"Whg. 1\" and row[Anreise] <= [Nicht Belegt] and row[Abreise] > [Nicht Belegt])) > 0 then 1 else 0),\n    // Wie der Abreisetag gerechnet werden soll liegt an > oder >=\n    //Whg1 = Table.AddColumn(DateT, \"Whg1\", each if Table.RowCount(Table.SelectRows(TypG, (row) => row[Wohnung] = \"Whg. 1\" and row[Anreise] <= [Nicht Belegt] and row[Abreise] >= [Nicht Belegt])) > 0 then 1 else 0),\n    Whg2 = Table.AddColumn(Whg1, \"Whg2\", each if Table.RowCount(Table.SelectRows(TypG, (row) => row[Wohnung] = \"Whg. 2\" and row[Anreise] <= [Nicht Belegt] and row[Abreise] > [Nicht Belegt])) > 0 then 1 else 0),\n    // Wie der Abreisetag gerechnet werden soll liegt an > oder >= bei row[Abreise] >= [Nicht Belegt]\n    //Whg2 = Table.AddColumn(Whg1, \"Whg2\", each if Table.RowCount(Table.SelectRows(TypG, (row) => row[Wohnung] = \"Whg. 2\" and row[Anreise] <= [Nicht Belegt] and row[Abreise] >= [Nicht Belegt])) > 0 then 1 else 0),\n    //Result = Table.SelectColumns(Table.SelectRows(Whg2, each ([Whg1] = 0) and ([Whg2] = 0)), {\"Nicht Belegt\"}),\n    Result = Table.TransformColumnTypes(Table.SelectColumns(Table.SelectRows(Whg2, each ([Whg1] = 0) and ([Whg2] = 0)), {\"Nicht Belegt\"}),{{\"Nicht Belegt\", type date}}, \"de-DE\"),\n    Monat = Table.AddColumn(Result, \"Monat\", each Date.MonthName([Nicht Belegt]), type text),\n    KW = Table.AddColumn(Monat, \"KW\", each let d = [Nicht Belegt], shifted = Date.AddDays(d, 3 - Date.DayOfWeek(d, Day.Monday)) in Date.WeekOfYear(shifted, Day.Monday), Int64.Type),\n    Jahr = Table.AddColumn(KW, \"Jahr\", each let d = [Nicht Belegt], shifted = Date.AddDays(d, 3 - Date.DayOfWeek(d, Day.Monday)) in Date.Year(shifted), Int64.Type),\n    KWText = Table.AddColumn(Jahr, \"KWText\", each Text.From([Jahr]) & \"-KW \" & Text.PadStart(Text.From([KW]), 2, \"0\"), type text),\n    Erg = Table.Sort(KWText, {{\"Nicht Belegt\", Order.Ascending}})\nin\n    Erg",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L538"
  },
  {
    "id": "PQ-013",
    "sourceId": "pq_0013",
    "title": "Items zusammenfassen - auf 2 Arten",
    "category": "Power Query",
    "summary": "Aus einer Liste (A2:D11) werden die \"items\" (Spalte B:D) nach dem Namen (Spalte A) zusammengefasst. Einmal pivotiert und einmal als Liste mit zusammengefassten \"items\".",
    "description": "Aus einer Liste (A2:D11) werden die \"items\" (Spalte B:D) nach dem Namen (Spalte A) zusammengefasst.\nEinmal pivotiert und einmal als Liste mit zusammengefassten \"items\".\n\n!!!!!!!!WICHTIG!!!!!!!!\nBei diesem Beispiel wird die Query UND die Ausgabe im Tabellenblatt DIREKT erzeugt!\nEs muss also kein M-Code kopiert werden. Die Ausgabe dauert einen Augenblick, da alles generiert wird!\n!!!!!!!!WICHTIG!!!!!!!!\n\nFür Informationen, wie mit dem M-Code umzugehen ist, auf den Button \"PQ M-Code Info\" klicken.\nUm die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.\n\nNach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!\n\nDas ist der zweite M-Code (Beide M-Codes sind im Beispiel schon drin):\nlet\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_13\"]}[Content],\n    Entpivotieren = Table.UnpivotOtherColumns(Quelle, {\"Name\"}, \"Attribut\", \"Wert\"),\n    EntferneD = Table.Distinct(Entpivotieren, {\"Name\", \"Wert\"}),\n    GruppeZ = Table.Group(EntferneD, {\"Name\"}, {{\"items\", each _, type table [Name=text, Attribut=text, Wert=text]}}),\n    HinzuS = Table.AddColumn(GruppeZ, \"itemsN\", each Text.Combine(List.Sort([items][Wert]), \", \")),\n    EntferneS = Table.RemoveColumns(HinzuS, {\"items\"})\nin\n    EntferneS",
    "tags": [
      "power query",
      "pivotieren",
      "m-code",
      "entpivotieren",
      "gruppieren"
    ],
    "codeText": "let\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Demo_PQ_13\"]}[Content],\n    Entpivotieren = Table.UnpivotOtherColumns(Quelle, {\"Name\"}, \"Attribut\", \"Wert\"),\n    EntferneD = Table.Distinct(Entpivotieren, {\"Name\", \"Wert\"}),\n    PivotS = Table.Pivot(EntferneD, List.Distinct(EntferneD[Attribut]), \"Attribut\", \"Wert\")\nin\n    PivotS",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L585"
  },
  {
    "id": "PQ-014",
    "sourceId": "pq_0014",
    "title": "Geburtstagsliste der nächsten X Tage anzeigen",
    "category": "Power Query",
    "summary": "Dieses Beispiel besteht aus einer Excel-Datei. PQ_Geburtstagsliste_Anzahl_Tage_vorgeben.xlsx",
    "description": "Dieses Beispiel besteht aus einer Excel-Datei.\nPQ_Geburtstagsliste_Anzahl_Tage_vorgeben.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv\nin folgenden Ordner entpackt:\n\n%TEMP%\\Excel-VSTO-Toolbox\\PQ_014\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nEs sind zwei Arten des Umgangs im M-Code. Es geht einmal darum,\nwie mit den \"TageBisGeburtstag\" umgegangen wird. Entweder\nbei heute Geburtstag 0 oder 365/366. Auch der 29.2. wird anders\nbehandelt.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "power query",
      "beispieldatei",
      "m-code",
      "geburtstag",
      "zip"
    ],
    "codeText": "'Bitte auf \"Demo erzeugen\" klicken um die Beispieldatei\n'zu entpacken und die XLSX zu öffnen.\n'Mit folgendem VBA-Code kann die Aktualisierung\n'automatisch erfolgen:\nOption Explicit\n'Excel -VSTO - Toolbox\n'Power Query - Demo\n'Ralf Stolzenburg (Case)\n'https://github.com/rstsu/Excel-VSTO-Toolbox\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n    On Error GoTo Fin\n    If Target.Address = \"$E$1\" Then\n        Application.EnableEvents = False\n        Me.ListObjects(\"tblErg_1\").QueryTable.Refresh\n        Me.ListObjects(\"tblErg_2\").QueryTable.Refresh\n    End If\nFin:\n    Application.Goto Range(\"E1\"), False\n    Application.EnableEvents = True\nEnd Sub",
    "isNew": false,
    "package": {
      "fileName": "Demo_PQ_14.zip",
      "folderName": "PQ_014",
      "sizeBytes": 28652,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_PQ_14.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L633"
  },
  {
    "id": "PQ-015",
    "sourceId": "pq_0015",
    "title": "Tage und Feiertage über Kontrollkästchen ausgeben",
    "category": "Power Query",
    "summary": "Dieses Beispiel besteht aus einer Excel-Datei. PQ_Power_Query_Feiertage_und_Tage_von_Montag_bis_Sonntag_ueber_Kontrollkaestchen_ausgeben.xlsb",
    "description": "Dieses Beispiel besteht aus einer Excel-Datei.\nPQ_Power_Query_Feiertage_und_Tage_von_Montag_bis_Sonntag_ueber_Kontrollkaestchen_ausgeben.xlsb\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv\nin folgenden Ordner entpackt:\n\n%TEMP%\\Excel-VSTO-Toolbox\\PQ_015\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nTage und Feiertage (mit KW und Quartal) werden als Liste ausgegeben.\nAuswahl der Tage über Kontrollkästchen -\n(die \"neuen\" über Einfügen - Kontrollkästchen).\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "power query",
      "beispieldatei",
      "m-code",
      "feiertage",
      "zip"
    ],
    "codeText": "'Bitte auf \"Demo erzeugen\" klicken um die Beispieldatei\n'zu entpacken und die XLSB zu öffnen.\n'Folgender VBA-Code ist schon in der Datei.\n'Hier nochmal zum kopieren:\nOption Explicit\n'Excel -VSTO - Toolbox\n'Power Query - Demo\n'Ralf Stolzenburg (Case)\n'https://github.com/rstsu/Excel-VSTO-Toolbox\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n    On Error GoTo Fin\n    If Not Intersect(Target, Range(\"B2, J2:Q2\")) Is Nothing Then\n        Application.EnableEvents = False\n        Me.ListObjects(\"tblErg\").QueryTable.Refresh\n    End If\nFin:\n    Application.Goto Range(\"B2\"), False\n    Application.EnableEvents = True\nEnd Sub",
    "isNew": false,
    "package": {
      "fileName": "Demo_PQ_15.zip",
      "folderName": "PQ_015",
      "sizeBytes": 22886,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_PQ_15.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L692"
  },
  {
    "id": "PQ-016",
    "sourceId": "pq_0016",
    "title": "Nummern mit 0 rechts auffüllen - auf 5 Stellen",
    "category": "Power Query",
    "summary": "Dieses Beispiel besteht aus einer Excel-Datei. PQ_UND_Formel_immer_auf_5_Stellen_RECHTS_mit_0_Null_auffuellen_Groesser_5_rechts_abtrennen_ODER_lassen.xlsx",
    "description": "Dieses Beispiel besteht aus einer Excel-Datei.\nPQ_UND_Formel_immer_auf_5_Stellen_RECHTS_mit_0_Null_auffuellen_Groesser_5_rechts_abtrennen_ODER_lassen.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv\nin folgenden Ordner entpackt:\n\n%TEMP%\\Excel-VSTO-Toolbox\\PQ_016\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nNummern werden mit 0 von rechts auf 5 Stellen aufgefüllt.\nSind mehr als 5 Stellen vorhanden, wird einmal gekürzt und einmal nicht.\nEs ist auch mit Formeln in der Beispieldatei gelöst.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "power query",
      "beispieldatei",
      "m-code",
      "nummern",
      "auffüllen"
    ],
    "codeText": "'Bitte auf \"Demo erzeugen\" klicken um die Beispieldatei\n'zu entpacken und die XLSX zu öffnen.\n'Mit folgendem VBA-Code kann die Aktualisierung\n'automatisch erfolgen:\nOption Explicit\n'Excel -VSTO - Toolbox\n'Power Query - Demo\n'Ralf Stolzenburg (Case)\n'https://github.com/rstsu/Excel-VSTO-Toolbox\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n    On Error GoTo Fin\n    If Target.Column = 1 And Target.Row > 1 Then\n        Application.EnableEvents = False\n        Me.ListObjects(\"tblErg_1\").QueryTable.Refresh\n        Me.ListObjects(\"tblErg_2\").QueryTable.Refresh\n    End If\n    Application.Goto Range(Target.Address), False\nFin:\n    Application.EnableEvents = True\nEnd Sub",
    "isNew": false,
    "package": {
      "fileName": "Demo_PQ_16.zip",
      "folderName": "PQ_016",
      "sizeBytes": 18993,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_PQ_16.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L749"
  },
  {
    "id": "PQ-017",
    "sourceId": "pq_0017",
    "title": "Zeitstempel bei bestimmtem Text eintragen oder löschen",
    "category": "Power Query",
    "summary": "Dieses Beispiel besteht aus einer Excel-Datei. PQ_Zeitstempel_TT_MM_JJJJ_hh_mm_ss_setzen_wenn_Spalte_C_leer_und_bestimmter_Text_in_Spalte_A.xlsx",
    "description": "Dieses Beispiel besteht aus einer Excel-Datei.\nPQ_Zeitstempel_TT_MM_JJJJ_hh_mm_ss_setzen_wenn_Spalte_C_leer_und_bestimmter_Text_in_Spalte_A.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv\nin folgenden Ordner entpackt:\n\n%TEMP%\\Excel-VSTO-Toolbox\\PQ_017\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nWenn in Spalte \"Wert_1\"  \"Neu\" steht und in der Spalte \"Zeitstempel\" kein Eintrag vorhanden ist, wird das aktuelle Datum mit Zeit eingetragen.\nWenn in Spalte \"Wert_1\"  \"Alt\" steht, wird der Zeitstempel geleert.\ntblErg_1 - Der Zeitstempel wird getrennt in Datum und Zeit.\ntblErg_2 - mit \"Replacer.ReplaceValue\" andere Herangehensweise.\ntblErg_2_1 - Der Zeitstempel wird nicht getrennt.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "power query",
      "uhrzeit",
      "m-code",
      "zeitstempel",
      "datum"
    ],
    "codeText": "'Bitte auf \"Demo erzeugen\" klicken um die Beispieldatei\n'zu entpacken und die XLSX zu öffnen.\n'Mit folgendem VBA-Code kann die Aktualisierung\n'automatisch erfolgen:\nOption Explicit\n'Excel -VSTO - Toolbox\n'Power Query - Demo\n'Ralf Stolzenburg (Case)\n'https://github.com/rstsu/Excel-VSTO-Toolbox\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n    On Error GoTo Fin\n    If Target.Column = 1 And Target.Row > 1 Then\n        Application.EnableEvents = False\n        Tabelle2.ListObjects(\"tblErg_1\").QueryTable.Refresh\n        Tabelle3.ListObjects(\"tblErg_2\").QueryTable.Refresh\n        Tabelle4.ListObjects(\"tblErg_2_1\").QueryTable.Refresh\n    End If\n    Application.Goto Range(Target.Address), False\nFin:\n    Application.EnableEvents = True\nEnd Sub",
    "isNew": false,
    "package": {
      "fileName": "Demo_PQ_17.zip",
      "folderName": "PQ_017",
      "sizeBytes": 150090,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_PQ_17.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L807"
  },
  {
    "id": "PQ-018",
    "sourceId": "pq_0018",
    "title": "Datum umwandeln",
    "category": "Power Query",
    "summary": "March 6 2020 8:00:12 AM February 28 2020 11:17:26 AM January 24 2020 3:22:44 PM December 30 2023 1:38 PM November 22 2024 7:29 AM 8/31/2025 8:15:57 AM 4/24/2022 9:31:16 AM 9/24/2026 …",
    "description": "March 6 2020 8:00:12 AM\nFebruary 28 2020 11:17:26 AM\nJanuary 24 2020 3:22:44 PM\nDecember 30 2023 1:38 PM\nNovember 22 2024 7:29 AM\n8/31/2025 8:15:57 AM\n4/24/2022 9:31:16 AM\n9/24/2026 9:31:16 PM\n2026-11-22T14:30:00Z\n2026-08-12T14:30:00Z\n\nDiese Daten werden umgewandelt. Auch UTC nach MEZ/MESZ.\nVBA_Power_Query_Formel_Datum_umwandeln.xlsb\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Demo_VBA_PQ_Formel\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "formel",
      "datum",
      "utc",
      "mez",
      "mesz",
      "vba",
      "power query"
    ],
    "codeText": "/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Tabelle1\"]}[Content],\n    Erg = Table.TransformColumnTypes(Quelle, {{\"Daten\", type datetime}}, \"en-US\")\nin\n    Erg\n\n/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Tabelle1\"]}[Content],\n    Parsen = Table.TransformColumnTypes(Quelle, {{\"Daten\", type datetime}}, \"en-US\"),\n    Erg = Table.TransformColumnTypes(Table.SplitColumn(Table.TransformColumnTypes(Parsen, {{\"Daten\", type text}}, \"de-DE\"), \"Daten\", Splitter.SplitTextByDelimiter(\" \", QuoteStyle.Csv), {\"Datum\", \"Uhrzeit\"}),{{\"Datum\", type date}, {\"Uhrzeit\", type time}})\nin\n    Erg\n\n/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    Quelle = Table.FromRows({\n        {\"March 6 2020 5:00 AM\"},          // Englisch\n        {\"Février 28 2020 5:00 AM\"},       // Französisch\n        {\"Enero 15 2021 14:20\"},           // Spanisch\n        {\"Giugno 5 2022 8:15\"},            // Italienisch\n        {\"November 22 2019 7:29 AM\"},      // Deutsch/Englisch\n        {\"8/31/2012 8:15:57 AM\"},          // US numerisch\n        {\"4/24/2012 9:31:16 AM\"},          // US numerisch\n        {\"22/08/2026\"},                    // DE numerisch\n        {\"15.09.2025 14:45\"},              // DE Punktformat\n        {\"2026-11-22T14:30:00Z\"},          // ISO UTC\n        {\"2026-08-12T14:30:00Z\"},          // ISO UTC\n        {\"2026-08-17\"}                     // ISO Datum\n    }, {\"Original\"}),\n    Lokal = {\"en-US\", \"de-DE\", \"fr-FR\", \"es-ES\", \"it-IT\", \"nl-NL\", \"pt-PT\"},\n    UTCLokal = (dt as datetime) as datetime =>\n        let\n            Zone = DateTimeZone.From(dt),\n            ZoneL = DateTimeZone.ToLocal(Zone),\n            ZoneN = DateTimeZone.RemoveZone(ZoneL)\n        in\n            ZoneN,\n    Parse = (val as any) as nullable datetime =>\n        let\n            DateOK = if Value.Is(val, type datetime) then val else null,\n            txt = if DateOK = null and Value.Is(val, type text) then Text.Trim(val) else null,\n            ParseISO = if txt <> null and Text.Length(txt) >= 4 and Text.Range(txt,0,4) >= \"0001\" and Text.Range(txt,0,4) <= \"9999\" and Text.Middle(txt,4,1) = \"-\" \n                        then try DateTime.FromText(txt, \"en-US\") otherwise null\n                        else null,\n            parsedLocales = if DateOK = null and ParseISO = null and txt <> null then\n                                let\n                                    tryLocales = List.First(\n                                        List.RemoveNulls(\n                                            List.Transform(Lokal, each try DateTime.FromText(txt, _) otherwise null)\n                                        ),\n                                        null\n                                    )\n                                in\n                                    tryLocales\n                            else null,\n            ParseA = List.First(List.RemoveNulls({DateOK, ParseISO, parsedLocales}), null),\n            result = if txt <> null and Text.EndsWith(txt, \"Z\") and ParseA <> null\n                    then UTCLokal(ParseA)\n                    else ParseA\n        in\n            result,\n    MitDatumZeit = Table.AddColumn(\n        Quelle,\n        \"DatumZeit\",\n        each Parse([Original]),\n        type nullable datetime\n    ),\n    MitDatum = Table.AddColumn(\n        MitDatumZeit,\n        \"Datum\",\n        each if [DatumZeit] <> null then Date.From([DatumZeit]) else null,\n        type date\n    ),\n    MitUhrzeit = Table.AddColumn(\n        MitDatum,\n        \"Uhrzeit\",\n        each if [DatumZeit] <> null then Time.From([DatumZeit]) else null,\n        type time\n    ),\n    Erg = Table.RemoveColumns(MitUhrzeit, {\"DatumZeit\"})\nin\n    Erg",
    "isNew": false,
    "package": {
      "fileName": "Demo_VBA_PQ_Formel.zip",
      "folderName": "Demo_VBA_PQ_Formel",
      "sizeBytes": 37133,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_VBA_PQ_Formel.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L868"
  },
  {
    "id": "PQ-019",
    "sourceId": "pq_0019",
    "title": "Biorhythmus mit Diagramm...",
    "category": "Power Query",
    "summary": "Auf dem Tabellenblatt \"Biorhythmus\" werden im Diagramm alle Linien angezeigt (die Punkte der \"kritischen Tage\" können ausgewählt werden. Im Tabellenblatt \"Biorhythmus_Part\" können die …",
    "description": "Auf dem Tabellenblatt \"Biorhythmus\" werden im Diagramm alle Linien angezeigt (die Punkte der \"kritischen Tage\" können ausgewählt werden.\nIm Tabellenblatt \"Biorhythmus_Part\" können die anzuzeigenden Linien ausgewählt werden.\nAktualisiert wird über VBA.\n\nPQ_Biorhythmus_Worksheet_Change_Diagramm_Chart.xlsb\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\PQ_019\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "pq",
      "power query",
      "chart",
      "diagramm",
      "biorhythmus",
      "vba"
    ],
    "codeText": "/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    GebTagRaw  = Excel.CurrentWorkbook(){[Name=\"GebTag\"]}[Content]{0}[Column1],\n    TageRaw    = Excel.CurrentWorkbook(){[Name=\"Tage\"]}[Content]{0}[Column1],\n    KKurvenRaw = Excel.CurrentWorkbook(){[Name=\"KKurven\"]}[Content]{0}[Column1],\n    GebTag = Date.From(GebTagRaw),\n    Tage   = Number.From(TageRaw),\n    NextDays = List.Dates(Date.From(DateTime.LocalNow()), Tage, #duration(1,0,0,0)),\n    Alter = List.Transform(NextDays, each Duration.Days(_ - GebTag)),\n    PI = 3.141592653589793,\n    Bio = (periode as number) as list =>\n        List.Transform(\n            Alter,\n            each Number.Round(\n                Number.Sin(2 * PI * _ / periode),\n                2\n            )\n        ),\n    Physisch      = Bio(23),\n    Emotional     = Bio(28),\n    Intellektuell = Bio(33),\n    Intuitiv      = Bio(38),\n    Spirituell    = Bio(53),\n    Kreativ       = Bio(41),\n    Perioden = [\n        physisch      = 23,\n        emotional     = 28,\n        intellektuell = 33,\n        intuitiv      = 38,\n        spirituell    = 53,\n        kreativ       = 41\n    ],\n    KritischeKurven =\n        let\n            txt = Text.Lower(Text.Trim(Text.From(KKurvenRaw)))\n        in\n            if txt = \"alle\" then\n                Record.FieldNames(Perioden)\n            else\n                List.Transform(\n                    Text.Split(txt, \",\"),\n                    each Text.Trim(_)\n                ),\n    IstKritisch =\n        (alterHeute as number, periode as number) as logical =>\n            Number.RoundDown(2 * (alterHeute - 1) / periode)\n            <>\n            Number.RoundDown(2 * alterHeute / periode),\n    Kritisch =\n    List.Transform(\n        Alter,\n        (a) =>\n            if\n                List.AnyTrue(\n                    List.Transform(\n                        KritischeKurven,\n                        (k) =>\n                            IstKritisch(\n                                a,\n                                Record.Field(Perioden, k)\n                            )\n                    )\n                )\n            then\n                0\n            else\n                null\n    ),\n        KritischeKurve =\n        List.Transform(\n            Alter,\n            (a) =>\n                let\n                    Treffer =\n                        List.Select(\n                            KritischeKurven,\n                            (k) =>\n                                IstKritisch(\n                                    a,\n                                    Record.Field(Perioden, k)\n                                )\n                        )\n                in\n                    if List.IsEmpty(Treffer) then\n                        null\n                    else\n                        Text.Combine(\n                            List.Transform(Treffer, each Text.Proper(_)),\n                            \", \"\n                        )\n        ),\n    Tabelle =\n        Table.FromColumns(\n            {\n                NextDays,\n                Physisch,\n                Emotional,\n                Intellektuell,\n                Intuitiv,\n                Spirituell,\n                Kreativ,\n                Kritisch,\n                KritischeKurve\n            },\n            {\n                \"Datum\",\n                \"Physisch\",\n                \"Emotional\",\n                \"Intellektuell\",\n                \"Intuitiv\",\n                \"Spirituell\",\n                \"Kreativ\",\n                \"Kritisch\",\n                \"Kritische Kurve\"\n            }\n        ),\n    Erg =\n        Table.TransformColumns(\n            Tabelle,\n            {\n                {\n                    \"Datum\",\n                    each Date.ToText(_, \"dd.MM.yyyy\"),\n                    type text\n                }\n            }\n        )\nin\n    Erg\n\n/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    GebTagRaw = Excel.CurrentWorkbook(){[Name=\"GebTagP\"]}[Content]{0}[Column1],\n    TageRaw   = Excel.CurrentWorkbook(){[Name=\"TageP\"]}[Content]{0}[Column1],\n    GebTag = Date.From(GebTagRaw),\n    Tage   = Number.From(TageRaw),\n    AuswahlRaw    = Excel.CurrentWorkbook(){[Name=\"tblKurve\"]}[Content],\n    AuswahlRecord = AuswahlRaw{0},\n    AlleSpalten = {\n        \"Physisch\",\n        \"Emotional\",\n        \"Intellektuell\",\n        \"Intuitiv\",\n        \"Spirituell\",\n        \"Kreativ\",\n        \"Kritisch\"\n    },\n    Angehaakt = List.Select(Record.FieldNames(AuswahlRecord), each Record.HasFields(AuswahlRecord, _) and Record.Field(AuswahlRecord, _) = true),\n    Auswahl = if List.IsEmpty(Angehaakt) then AlleSpalten else Angehaakt,\n    BioNamen = {\n        \"Physisch\",\n        \"Emotional\",\n        \"Intellektuell\",\n        \"Intuitiv\",\n        \"Spirituell\",\n        \"Kreativ\"\n    },\n    AusgewaehlteBioKurven = List.Intersect({Auswahl, BioNamen}),\n    KritischeKurven = if List.IsEmpty(AusgewaehlteBioKurven) then BioNamen else AusgewaehlteBioKurven,\n    NextDays = List.Dates(Date.From(DateTime.LocalNow()), Tage, #duration(1, 0, 0, 0)),\n    Alter = List.Transform(NextDays, each Duration.Days(_ - GebTag)),\n    PI = 3.141592653589793,\n    Bio = (periode as number) as list =>\n        List.Transform(\n            Alter,\n            each Number.Round(\n                Number.Sin(2 * PI * _ / periode),\n                2\n            )\n        ),\n    Physisch      = Bio(23),\n    Emotional     = Bio(28),\n    Intellektuell = Bio(33),\n    Intuitiv      = Bio(38),\n    Spirituell    = Bio(53),\n    Kreativ       = Bio(41),\n    Perioden = [\n        Physisch      = 23,\n        Emotional     = 28,\n        Intellektuell = 33,\n        Intuitiv      = 38,\n        Spirituell    = 53,\n        Kreativ       = 41\n    ],\n    IstKritisch =\n        (alterHeute as number, periode as number) as logical =>\n            Number.RoundDown(\n                2 * (alterHeute - 1) / periode\n            )\n            <>\n            Number.RoundDown(\n                2 * alterHeute / periode\n            ),\n    Kritisch =\n        List.Transform(\n            Alter,\n            (a) =>\n                if\n                    List.AnyTrue(\n                        List.Transform(\n                            KritischeKurven,\n                            (k) =>\n                                IstKritisch(\n                                    a,\n                                    Record.Field(Perioden, k)\n                                )\n                        )\n                    )\n                then\n                    0\n                else\n                    null\n        ),\n    NullListe = List.Repeat({null}, Tage),\n    PhysischOut = if List.Contains(Auswahl, \"Physisch\") then Physisch else NullListe,\n    EmotionalOut = if List.Contains(Auswahl, \"Emotional\") then Emotional else NullListe,\n    IntellektuellOut = if List.Contains(Auswahl, \"Intellektuell\") then Intellektuell else NullListe,\n    IntuitivOut = if List.Contains(Auswahl, \"Intuitiv\") then Intuitiv else NullListe,\n    SpirituellOut = if List.Contains(Auswahl, \"Spirituell\") then Spirituell else NullListe,\n    KreativOut = if List.Contains(Auswahl, \"Kreativ\") then Kreativ else NullListe,\n    KritischOut = if List.Contains(Auswahl, \"Kritisch\") then Kritisch else NullListe,\n    Tabelle =\n        Table.FromColumns(\n            {\n                NextDays,\n                PhysischOut,\n                EmotionalOut,\n                IntellektuellOut,\n                IntuitivOut,\n                SpirituellOut,\n                KreativOut,\n                KritischOut\n            },\n            {\n                \"Datum\",\n                \"Physisch\",\n                \"Emotional\",\n                \"Intellektuell\",\n                \"Intuitiv\",\n                \"Spirituell\",\n                \"Kreativ\",\n                \"Kritisch\"\n            }\n        ),\n    Erg =\n        Table.TransformColumnTypes(\n            Tabelle,\n            {\n                {\"Datum\", type date},\n                {\"Physisch\", type number},\n                {\"Emotional\", type number},\n                {\"Intellektuell\", type number},\n                {\"Intuitiv\", type number},\n                {\"Spirituell\", type number},\n                {\"Kreativ\", type number},\n                {\"Kritisch\", type number}\n            }\n        )\nin\n    Erg",
    "isNew": false,
    "package": {
      "fileName": "Demo_PQ_19.zip",
      "folderName": "PQ_019",
      "sizeBytes": 43593,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_PQ_19.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L1010"
  },
  {
    "id": "PQ-020",
    "sourceId": "pq_0020",
    "title": "Freitag der 13te...",
    "category": "Power Query",
    "summary": "Jeder Freitag der 13te wird ausgegeben. Nach Start- und Enddatum. Zusätzlich kann noch der Monat gewählt werden. Aktualisiert wird über VBA. …",
    "description": "Jeder Freitag der 13te wird ausgegeben. Nach Start- und Enddatum. Zusätzlich kann noch der Monat gewählt werden.\nAktualisiert wird über VBA.\nPQ_Freitag_der_13te_Startdatum_Enddatum_Monat.xlsb\n\nAuch mit Formeln gelöst (Formeln in O1, Q1 und S1).\n=LET(x;SEQUENZ(A2-A1+1;1;A1;1);VSTAPELN(\"Freitag, der 13te...\";FILTER(x;(x>=DATUM(1900;3;1))*(TAG(x)=13)*(WOCHENTAG(x;2)=5))))\n=LET(x;SEQUENZ(A2-A1+1;1;A1;1);m;MONAT(DATWERT(\"1. \"&A4));VSTAPELN(\"Freitag, der 13te von A4\";FILTER(x;(x>=DATUM(1900;3;1))*(TAG(x)=13)*(WOCHENTAG(x;2)=5)*(MONAT(x)=m))))\n=LET(x;SEQUENZ(Ende-Start+1;1;Start;1);m;MONAT(DATWERT(\"1. \"&Monat));VSTAPELN(\"Freitag, der 13te ohne A4\";FILTER(x;(x>=DATUM(1900;3;1))*(TAG(x)=13)*(WOCHENTAG(x;2)=5)*(MONAT(x)=m))))\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\PQ_020\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "pq",
      "power query",
      "freitag",
      "13",
      "formel",
      "vba"
    ],
    "codeText": "/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    Startdatum = Date.From(Excel.CurrentWorkbook(){[Name=\"Start\"]}[Content]{0}[Column1]),\n    Enddatum = Date.From(Excel.CurrentWorkbook(){[Name=\"Ende\"]}[Content]{0}[Column1]),\n    StartMonat = Date.StartOfMonth(Startdatum),\n    EndeMonat = Date.StartOfMonth(Enddatum),\n    AnzahlMonate = (Date.Year(EndeMonat) - Date.Year(StartMonat)) * 12 + Date.Month(EndeMonat) - Date.Month(StartMonat) + 1,\n    Monate = List.Transform({0..AnzahlMonate - 1}, each Date.AddMonths(StartMonat, _)),\n    Dreizehnte = List.Transform(Monate, each #date(Date.Year(_), Date.Month(_), 13)),\n    Freitag13 = List.Select(Dreizehnte, each _ >= Startdatum and _ <= Enddatum and Date.DayOfWeek(_, Day.Monday) = 4),\n    Tabelle = Table.FromList(Freitag13, Splitter.SplitByNothing(), {\"Datum\"}),\n    MitWochentag = Table.AddColumn(Tabelle, \"Wochentag\", each Date.DayOfWeekName([Datum], \"de-DE\"), type text),\n    MitMonat = Table.AddColumn(MitWochentag, \"Monat\", each Date.MonthName([Datum], \"de-DE\"), type text),\n    MitJahr = Table.AddColumn(MitMonat, \"Jahr\", each Date.Year([Datum]), Int64.Type)\nin\n    MitJahr\n\n/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    Startdatum = Date.From(Excel.CurrentWorkbook(){[Name=\"Start\"]}[Content]{0}[Column1]),\n    Enddatum = Date.From(Excel.CurrentWorkbook(){[Name=\"Ende\"]}[Content]{0}[Column1]),\n    AuswahlMonat = Text.From(Excel.CurrentWorkbook(){[Name=\"Monat\"]}[Content]{0}[Column1]),\n    StartMonat = Date.StartOfMonth(Startdatum),\n    EndeMonat = Date.StartOfMonth(Enddatum),\n    AnzahlMonate = (Date.Year(EndeMonat) - Date.Year(StartMonat)) * 12 + Date.Month(EndeMonat) - Date.Month(StartMonat) + 1,\n    Monate = List.Transform({0..AnzahlMonate - 1}, each Date.AddMonths(StartMonat, _)),\n    Dreizehnte = List.Transform(Monate, each #date(Date.Year(_), Date.Month(_), 13)),\n    Freitag13 = List.Select(Dreizehnte, each _ >= Startdatum and _ <= Enddatum and Date.DayOfWeek(_, Day.Monday) = 4 and Date.MonthName(_, \"de-DE\") = AuswahlMonat),\n    Tabelle = Table.FromList(Freitag13, Splitter.SplitByNothing(), {\"Datum\"}),\n    MitWochentag = Table.AddColumn(Tabelle, \"Wochentag\", each Date.DayOfWeekName([Datum], \"de-DE\"), type text),\n    MitMonat = Table.AddColumn(MitWochentag, \"Monat\", each Date.MonthName([Datum], \"de-DE\"), type text),\n    MitJahr = Table.AddColumn(MitMonat, \"Jahr\", each Date.Year([Datum]), Int64.Type)\nin\n    MitJahr",
    "isNew": false,
    "package": {
      "fileName": "Demo_PQ_20.zip",
      "folderName": "PQ_020",
      "sizeBytes": 28827,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_PQ_20.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L1315"
  },
  {
    "id": "PQ-021",
    "sourceId": "pq_0021",
    "title": "Langer Text in Zellen aufteilen...",
    "category": "Power Query",
    "summary": "Lange Texte in Zellen werden nach vorgegebener Anzahl von Zeichen aufgetrennt. Nicht mitten im Wort. Verschiedene Möglichkeiten (in Spalten, Zeilen, Blöcken...). Mit RegEx, Power Query, …",
    "description": "Lange Texte in Zellen werden nach vorgegebener Anzahl von Zeichen aufgetrennt.\nNicht mitten im Wort.\nVerschiedene Möglichkeiten (in Spalten, Zeilen, Blöcken...).\nMit RegEx, Power Query, Formeln und VBA (UDF und Sub).\n\nRegEx_Power_Query_Formel_VBA_Text_auftrennen.xlsb\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Demo_RegEx_PQ_Formel_VBA\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "auftrennenl",
      "text",
      "aufteilen",
      "formel",
      "regex",
      "vba",
      "power query"
    ],
    "codeText": "/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Tabelle1\"]}[Content],\n    QuelleN = Excel.CurrentWorkbook(){[Name=\"AnzZ\"]}[Content]{0}[Column1],\n    TextTrennen = (TextT as text, LenT as number) as list =>\n        let\n            Worte = Text.Split(TextT, \" \"),\n            Result = List.Accumulate(Worte, {\"\"}, (Liste, Worte) =>\n                let\n                    ZeileL = List.Last(Liste),\n                    ZeileN = if Text.Length(ZeileL & \" \" & Worte) <= LenT then Text.Trim(ZeileL & \" \" & Worte) else Worte,\n                    ListeL = if Text.Length(ZeileL & \" \" & Worte) <= LenT then List.RemoveLastN(Liste,1) & {ZeileN} else Liste & {ZeileN}\n                in\n                    ListeL)\n        in\n            Result,\n    ListeM = Table.AddColumn(Quelle, \"Getrennt_nach_C1\", each TextTrennen([Daten], QuelleN)),\n    Erg = Table.RemoveColumns(Table.ExpandListColumn(ListeM, \"Getrennt_nach_C1\"),{\"Daten\"})\nin\n    Erg\n\n/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Tabelle3\"]}[Content],\n    TextTrennen = (TextT as text, LenT as number) as list =>\n        let\n            Worte = Text.Split(TextT, \" \"),\n            Result = List.Accumulate(Worte, {\"\"}, (Liste, Worte) =>\n                let\n                    ZeileL = List.Last(Liste),\n                    ZeileN = if Text.Length(ZeileL & \" \" & Worte) <= LenT then Text.Trim(ZeileL & \" \" & Worte) else Worte,\n                    ListeL = if Text.Length(ZeileL & \" \" & Worte) <= LenT then List.RemoveLastN(Liste,1) & {ZeileN} else Liste & {ZeileN}\n                in\n                    ListeL)\n        in\n            Result,\n    MitListen = Table.AddColumn(Quelle, \"Getrennt_nach_Spalte_B\", each TextTrennen([Daten], [Länge])),\n    Erg = Table.RemoveColumns(Table.ExpandListColumn(MitListen, \"Getrennt_nach_Spalte_B\"),{\"Daten\", \"Länge\"})\nin\n    Erg\n\n/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Tabelle1\"]}[Content],\n    QuelleN = Excel.CurrentWorkbook(){[Name=\"AnzZ\"]}[Content]{0}[Column1],\n    TextTrennen = (TextT as text, LenT as number) as text =>\n        let\n            Worte = Text.Split(TextT, \" \"),\n            Result = List.Accumulate(Worte, {\"\"}, (Liste, Worte) =>\n                let\n                    ZeileL = List.Last(Liste),\n                    ZeileN = if Text.Length(ZeileL & \" \" & Worte) <= LenT then Text.Trim(ZeileL & \" \" & Worte) else Worte,\n                    ListeL = if Text.Length(ZeileL & \" \" & Worte) <= LenT then List.RemoveLastN(Liste,1) & {ZeileN} else Liste & {ZeileN}\n                in\n                    ListeL),\n            Erg = Text.Combine(Result, \"#(lf)\")\n        in\n            Erg,\n    Erg1 = Table.RemoveColumns(Table.AddColumn(Quelle, \"ZeilenText\", each TextTrennen([Daten], QuelleN)),{\"Daten\"})\nin\n    Erg1\n\n/*\nExcel-VSTO-Toolbox\nPower Query-Demo\nRalf Stolzenburg (Case)\nhttps://github.com/rstsu/Excel-VSTO-Toolbox\n*/\nlet\n    Quelle = Excel.CurrentWorkbook(){[Name=\"Tabelle1\"]}[Content],\n    QuelleN = Excel.CurrentWorkbook(){[Name=\"AnzZ\"]}[Content]{0}[Column1],\n    TextTrennen = (TextT as text, LenT as number) as list =>\n        let\n            Worte = Text.Split(Text.Trim(TextT), \" \"),\n            Result =\n                List.Accumulate(\n                    Worte,\n                    {\"\"},\n                    (Liste, Wort) =>\n                        let\n                            ZeileL = List.Last(Liste),\n                            Neu = Text.Trim(ZeileL & \" \" & Wort),\n                            Passt = Text.Length(Neu) <= LenT,\n                            ListeN =\n                                if Passt\n                                then List.RemoveLastN(Liste, 1) & {Neu}\n                                else Liste & {Wort}\n                        in\n                            ListeN\n                )\n        in\n            Result,\n    MitListen = Table.AddColumn(Quelle, \"Liste\", each TextTrennen([Daten], QuelleN)),\n    MitRecords = Table.AddColumn(MitListen, \"Getrennt\", each Record.FromList([Liste], List.Transform({1..List.Count([Liste])}, each \"Teil_\" & Text.From(_)))),\n    AlleSpalten = List.Union(List.Transform(MitRecords[Getrennt], each Record.FieldNames(_))),\n    Expandiert = Table.ExpandRecordColumn(MitRecords, \"Getrennt\", AlleSpalten, AlleSpalten),\n    Erg = Table.RemoveColumns(Expandiert, {\"Daten\", \"Liste\"})\nin\n    Erg",
    "isNew": true,
    "package": {
      "fileName": "Demo_RegEx_PQ_Formel_VBA.zip",
      "folderName": "Demo_RegEx_PQ_Formel_VBA",
      "sizeBytes": 73867,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_RegEx_PQ_Formel_VBA.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.PowerQuery.vb#L1397"
  },
  {
    "id": "RX-001",
    "sourceId": "regex_001",
    "title": "Text und Zahlen trennen",
    "category": "Regex",
    "summary": "Text und Zahlen trennen - Zahlen mit und ohne Klammern ausgeben. Mit 3 Formeln und mit einer Spill-Formel.",
    "description": "Text und Zahlen trennen - Zahlen mit und ohne Klammern ausgeben.\nMit 3 Formeln und mit einer Spill-Formel.",
    "tags": [
      "regex",
      "zahl",
      "klammer",
      "text",
      "extrahieren"
    ],
    "codeText": "=REGEXEXTRAHIEREN(A2:.A999;\"^(.+?)\\s*\\(?\\d+\\)?$\";2)\n=REGEXEXTRAHIEREN(A2:.A999;\"\\(?(\\d+)\\)?$\")\n=REGEXEXTRAHIEREN(A2:.A999;\"\\((\\d+)\\)\";2)\n\n=LET(w;A2:.A999;x;REGEXEXTRAHIEREN(w;\"\\((\\d+)\\)\";2);y;REGEXEXTRAHIEREN(w;\"(\\(\\d+\\))\");z;REGEXEXTRAHIEREN(w;\"^(.+?)\\s*\\(\\d+\\)$\";2);HSTAPELN(z;x;y))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L8"
  },
  {
    "id": "RX-002",
    "sourceId": "regex_002",
    "title": "Ordner am Backslash kürzen",
    "category": "Regex",
    "summary": "Ordner in Spalte A am Backslash (von rechts) kürzen nach Vorgabe/Länge in Spalte B. Mit REGEXEXTRAHIEREN oder auch TEXTVOR. Die unteren beiden Formeln sind mit Überschrift.",
    "description": "Ordner in Spalte A am Backslash (von rechts) kürzen nach Vorgabe/Länge in Spalte B.\nMit REGEXEXTRAHIEREN oder auch TEXTVOR.\nDie unteren beiden Formeln sind mit Überschrift.",
    "tags": [
      "regex",
      "ordner",
      "kürzen",
      "text",
      "backslash"
    ],
    "codeText": "=REGEXEXTRAHIEREN(A2:.A999;\"^(.*)(?:\\\\[^\\\\]+){\"&B2:.B100&\"}$\";2)\n=TEXTVOR(A2:.A999;\"\"\\\"\";-B2:.B999)\n\n=VSTAPELN(A1;REGEXEXTRAHIEREN(A2:.A999;\"^(.*)(?:\\\\[^\\\\]+){\"&B2:.B100&\"}$\";2))\n=VSTAPELN(A1;TEXTVOR(A2:.A999;\"\\\";-B2:.B999))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L33"
  },
  {
    "id": "RX-003",
    "sourceId": "regex_003",
    "title": "Zollangaben aus Text auslesen",
    "category": "Regex",
    "summary": "In den Zellen A2:Ax sind Texte mit mehreren Zollangaben (z. B. 1/2\" oder dann auch 1 3/4\"). Die Zollangaben sollen ausgelesen werden. Auch mehrere. Überschriften (Z1, Z2...) automatisch …",
    "description": "In den Zellen A2:Ax sind Texte mit mehreren Zollangaben (z. B. 1/2\" oder dann auch 1 3/4\").\nDie Zollangaben sollen ausgelesen werden. Auch mehrere.\nÜberschriften (Z1, Z2...) automatisch generieren.\nDie zweite Formel zeigt den Weg, wenn IMMER maximal 3 Zollangaben vorhanden sind.",
    "tags": [
      "regex",
      "zoll",
      "auslesen",
      "text",
      "mehrere"
    ],
    "codeText": "=LET(x;A2:.A999;p;\"(?:\\d+\\s+)?\\d+(?:/\\d+)?\"\"\";n;MAX(MAP(x;LAMBDA(a;ANZAHL2(REGEXEXTRAHIEREN(a;p;1)))));VSTAPELN(\"Z\"&SEQUENZ(;n);MATRIXERSTELLEN(ZEILEN(x);n;LAMBDA(r;c;WENNFEHLER(INDEX(REGEXEXTRAHIEREN(INDEX(x;r);p;1);c);\"\")))))\n\n=VSTAPELN({\"Z1\".\"Z2\".\"Z3\"};MATRIXERSTELLEN(ZEILEN(A2:.A999);3;LAMBDA(r;c;WENNFEHLER(INDEX(REGEXEXTRAHIEREN(INDEX(A2:.A999;r);\"(?:\\d+\\s+)?\\d+(?:/\\d+)?\"\"\";1);c);\"\"))))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L59"
  },
  {
    "id": "RX-004",
    "sourceId": "regex_004",
    "title": "Sonderzeichen entfernen",
    "category": "Regex",
    "summary": "In A2:A8 stehen Texte mit Sonderzeichen - auch mehrere hintereinander. Diese sollen entfernt werden. Zwischen den Wörtern darf nur ein Leerzeichen übrig bleiben. Einmal wird das Pattern …",
    "description": "In A2:A8 stehen Texte mit Sonderzeichen - auch mehrere hintereinander.\nDiese sollen entfernt werden.\nZwischen den Wörtern darf nur ein Leerzeichen übrig bleiben.\nEinmal wird das Pattern direkt in die Formel geschrieben und einmal aus einer Zelle (J2) in die Formel übernommen.\n\nDie dritte und vierte Formel wieder mit Überschrift.",
    "tags": [
      "regex",
      "sonderzeichen",
      "entfernen",
      "text",
      "mehrere"
    ],
    "codeText": "=GLÄTTEN(REGEXERSETZEN(B2:.B999;\"[:?*/\\\\]\";\"\"))\n=GLÄTTEN(REGEXERSETZEN(B2:.B999;A10;\"\"))\n\n=VSTAPELN(B1;GLÄTTEN(REGEXERSETZEN(B2:.B999;\"[:?*/\\\\]\";\"\")))\n=VSTAPELN(B1;GLÄTTEN(REGEXERSETZEN(B2:.B999;A10;\"\")))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L84"
  },
  {
    "id": "RX-005",
    "sourceId": "regex_005",
    "title": "Bestimmte Textteile auslesen",
    "category": "Regex",
    "summary": "In A2:A7 stehen Texte in der Art: \\NH-DATA\\Public\\CAD\\Project_1\\UV_A_1_GL401 A.pro\\ Der Bereich zwischen dem vorletzten \\ und dem .pro soll ausgelsesen werden. In dem Fall - UV_A_1_GL401 …",
    "description": "In A2:A7 stehen Texte in der Art: \\NH-DATA\\Public\\CAD\\Project_1\\UV_A_1_GL401 A.pro\\\nDer Bereich zwischen dem vorletzten \\ und dem .pro soll ausgelsesen werden.\nIn dem Fall - UV_A_1_GL401 A - ohne die beiden Zeichen, an denen getrennt wird.\nGroß- Kleinschreibung soll keine Rolle spielen.",
    "tags": [
      "regex",
      "auslesen",
      "entfernen",
      "text",
      "mehrere"
    ],
    "codeText": "=VSTAPELN(A1;REGEXEXTRAHIEREN(A2:.A999;\"\\\\([^\\\\]+)\\.pro\\\\\";2;1))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L113"
  },
  {
    "id": "RX-006",
    "sourceId": "regex_006",
    "title": "Straße und Hausnummer trennen",
    "category": "Regex",
    "summary": "Straße und Hausnummer in A2:A11. Diese sollen getrennt werden. Exoten wie \"Strasse des 17. Juni\" bleiben komplett bestehen - es sei denn, es gibt eine Hausnummer am Ende.",
    "description": "Straße und Hausnummer in A2:A11.\nDiese sollen getrennt werden.\nExoten wie \"Strasse des 17. Juni\" bleiben komplett bestehen - es sei denn, es gibt eine Hausnummer am Ende.",
    "tags": [
      "regex",
      "straße",
      "hausnummer",
      "trennen",
      "mehrere"
    ],
    "codeText": "=LET(v;A2:.A999;w;HSTAPELN(\"Straße\";\"Hausnummer\");hn;WENNFEHLER(REGEXEXTRAHIEREN(v;\"\\d+\\s*[A-Za-z]?$\");\"\");str;WENN(hn=\"\";v;GLÄTTEN(REGEXERSETZEN(v;\"\\s*\\d+\\s*[A-Za-z]?$\";\"\")));VSTAPELN(w;HSTAPELN(str;hn)))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L136"
  },
  {
    "id": "RX-007",
    "sourceId": "regex_007",
    "title": "Datumbereich aus String auslesen",
    "category": "Regex",
    "summary": "Strings mit Datumbereich in A2:A7 (z. B. Text und 12345 dann12.02.-14.07.2026 Text). Der Datumbereich soll ausgelesen werden. Einmal in eine Zelle - und dann noch getrennt in zwei …",
    "description": "Strings mit Datumbereich in A2:A7 (z. B. Text und 12345 dann12.02.-14.07.2026 Text).\nDer Datumbereich soll ausgelesen werden.\nEinmal in eine Zelle - und dann noch getrennt in zwei Zellen.\nAusgegeben werden auch die Arbeits- und Kalendertage zwischen den Daten.\nFormeln in B1, C1, D1, E1 und H1.",
    "tags": [
      "regex",
      "datum",
      "bereich",
      "zahlen",
      "mehrere"
    ],
    "codeText": "=VSTAPELN(\"Datum auslesen\";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;\"\\d{2}\\.\\d{2}\\.?-?\\s*-\\s*\\d{2}\\.\\d{2}\\.\\d{4}\");\"???\"))\n\n=VSTAPELN(A1&\" - Richtig\";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;\"\\d{2}\\.\\d{2}(?:\\.\\d{4})?\\.?\\s*-\\s*(?:\\d{2}\\.\\d{2}|\\d{2})\\.\\d{4}|\\d{2}\\.\\d{2}\\.\\d{4}|\\d{2}\\.\\d{4}\");\"\"))\n\n=VSTAPELN(A1&\" - Oder\";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;\"\\d{2}\\.\\d{2}(?:\\.\\d{4})?\\.?\\s*-\\s*\\d{2}\\.\\d{2}\\.\\d{4}\");\"\"))\n\n=LET(a;A2:.A999;\n    b;WENNFEHLER(\n        REGEXEXTRAHIEREN(\n            a;\n            \"\\d{2}\\.\\d{2}(?:\\.\\d{4})?\\.?\\s*-\\s*(?:\\d{2}\\.\\d{2}|\\d{2})\\.\\d{4}|\\d{2}\\.\\d{2}\\.\\d{4}|\\d{2}\\.\\d{4}|\\d{2}\\.\\d{2}\";\n            1\n        );\n        \"\"\n    );\n    c;ISTZAHL(SUCHEN(\"-\";b));\n    d;WENN(\n        c;\n        REGEXEXTRAHIEREN(b;\"^\\d{2}\\.\\d{2}(?:\\.\\d{4})?\");\n        b\n    );\n    e;WENN(\n        c;\n        GLÄTTEN(TEXTNACH(b;\"-\"));\n        \"\"\n    );\n    f;LAMBDA(s;g;\n        WENN(s=\"\";\n            \"\";\n            WENN(\n                UND(LÄNGE(s)=5;ISTZAHL(WERT(TEIL(s;1;2))));\n                DATUM(g;TEIL(s;4;2);TEIL(s;1;2));\n                WENN(\n                    LÄNGE(s)=7;\n                    DATUM(TEIL(s;4;4);TEIL(s;1;2);1);\n                    WENN(\n                        LÄNGE(s)=10;\n                        DATUM(TEIL(s;7;4);TEIL(s;4;2);TEIL(s;1;2));\n                        WENN(\n                            LÄNGE(s)=6;\n                            DATUM(g;TEIL(s;4;2);TEIL(s;1;2));\n                            s\n                        )\n                    )\n                )\n            )\n        )\n    );\n    h;f(d;JAHR(HEUTE()));\n    i;WENN(\n        e<>\"\";\n        LAMBDA(t;\n            WENN(\n                ODER(LÄNGE(e)=5;LÄNGE(e)=6);\n                LET(\n                    j;f(e;JAHR(h));\n                    WENN(j<h;DATUM(JAHR(j)+1;MONAT(j);TAG(j));j)\n                );\n                f(e;JAHR(HEUTE()))\n            )\n        )(e);\n        \"\"\n    );\n    k;HSTAPELN(h;i);\n    WENN(ANZAHL2(k)=0;\"\";\n    VSTAPELN({\"Startdatum\".\"Enddatum\"};k)))\n\n=LET(\n    a;A2:.A999;\n    b;WENNFEHLER(\n        REGEXEXTRAHIEREN(\n            a;\n            \"\\d{2}\\.\\d{2}(?:\\.\\d{4})?\\.?\\s*-\\s*(?:\\d{2}\\.\\d{2}|\\d{2})\\.\\d{4}|\\d{2}\\.\\d{2}\\.\\d{4}|\\d{2}\\.\\d{4}|\\d{2}\\.\\d{2}\";\n            1\n        );\n        \"\"\n    );\n    c;ISTZAHL(SUCHEN(\"-\";b));\n    d;WENN(\n        c;\n        REGEXEXTRAHIEREN(b;\"^\\d{2}\\.\\d{2}(?:\\.\\d{4})?\");\n        b\n    );\n    e;WENN(\n        c;\n        GLÄTTEN(TEXTNACH(b;\"-\"));\n        \"\"\n    );\n    f; LAMBDA(s;g;\n        WENN(s=\"\";\n            \"\";\n            WENNFEHLER(\n                WENN(\n                    LÄNGE(s)=5;\n                    DATUM(g;TEIL(s;4;2);TEIL(s;1;2));\n                    WENN(\n                        LÄNGE(s)=7;\n                        DATUM(TEIL(s;4;4);TEIL(s;1;2);1);\n                        WENN(\n                            LÄNGE(s)=10;\n                            DATUM(TEIL(s;7;4);TEIL(s;4;2);TEIL(s;1;2));\n                            WENN(\n                                LÄNGE(s)=6;\n                                DATUM(g;TEIL(s;4;2);TEIL(s;1;2));\n                                \"\"\n                            )\n                        )\n                    )\n                );\n                \"\"\n            )\n        )\n    );\n    h; f(d;JAHR(HEUTE()));\n    i; WENN(\n        e<>\"\";\n        LAMBDA(t;\n            WENN(\n                ODER(LÄNGE(e)=5;LÄNGE(e)=6);\n                LET(\n                    j;f(e;JAHR(h));\n                    WENN(j<h;DATUM(JAHR(j)+1;MONAT(j);TAG(j));j)\n                );\n                f(e;JAHR(HEUTE()))\n            )\n        )(e);\n        \"\"\n    );\n    k; WENN(\n        UND(ISTZAHL(h);ISTZAHL(i));\n        NETTOARBEITSTAGE(h;i);\n        \"\"\n    );\n    l; WENN(\n        UND(ISTZAHL(h);ISTZAHL(i));\n        i-h;\n        \"\"\n    );\n    m;HSTAPELN(h;i;k;l);VSTAPELN({\"Startdatum\".\"Enddatum\".\"Arbeitstage\".\"Tage\"};m))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L158"
  },
  {
    "id": "RX-008",
    "sourceId": "regex_008",
    "title": "String zwischen 2 Zahlen auslesen",
    "category": "Regex",
    "summary": "Aus einer Liste (A2:A12) wird ein Text zwischen 2 Zahlem ausgelesen. Z. B. \"1. Vom Bodensee in den Schwarzwald 21 May 2001\". Es wird nur der \"Titel\" in der Mitte ausgelesen bzw. der Rest …",
    "description": "Aus einer Liste (A2:A12) wird ein Text zwischen 2 Zahlem ausgelesen.\nZ. B. \"1. Vom Bodensee in den Schwarzwald 21 May 2001\".\nEs wird nur der \"Titel\" in der Mitte ausgelesen bzw. der Rest ersetzt.\n\nFormel in C1:\nDer Titel bleibt übrig, da Anfang (Nummer mit Punkt) und Ende (Datum am Schluss) entfernt werden.\n\nFormel in E1:\nHier wird der Titel gezielt als \"Gruppe\" herausgenommen und der gesamte Text am Ende damit ersetzt (\"$1\").",
    "tags": [
      "regex",
      "string",
      "bereich",
      "zahlen",
      "mehrere"
    ],
    "codeText": "=VSTAPELN(A1;REGEXERSETZEN(A2:.A999;\"^\\s*\\d+\\.\\s*|\\s+\\d{1,2}\\s+\\S+\\s+\\d{4}\\s*$\";\"\"))\n\n=VSTAPELN(\"Ausgabe\";REGEXERSETZEN(A2:.A999;\"^\\s*\\d+\\.\\s*(.*?)\\s+\\d{1,2}\\s+\\S+\\s+\\d{4}\\s*$\";\"$1\"))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L320"
  },
  {
    "id": "RX-009",
    "sourceId": "regex_009",
    "title": "String mit Text, Zahl und Datum/Zeit bearbeiten",
    "category": "Regex",
    "summary": "Aus einer Liste (A2:A10) wird ein Text mit Zahlen und Datum/Zeit ausgelesen.",
    "description": "Aus einer Liste (A2:A10) wird ein Text mit Zahlen und Datum/Zeit ausgelesen.\n\nZ. B. \"Dateiname Beschreibung Kuerzel 31_05_02_2025_12_13_09.pdf\".\nWird zu \"Dateiname Beschreibung Kuerzel 31.pdf\".\n\nDie Zahl nach \"Kuerzel\" soll erhalten bleiben - nur Datum und Zeit entfernen.\n\nFormel in C1 und E2 (der Formeltext).",
    "tags": [
      "regex",
      "string",
      "datum",
      "zahlen",
      "text"
    ],
    "codeText": "=VSTAPELN(\"Name\";REGEXERSETZEN(A2:.A999;\"_\\d{2}_\\d{2}_\\d{4}_\\d{2}_\\d{2}_\\d{2}\";\"\"))\n\n=FORMELTEXT(C1)",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L350"
  },
  {
    "id": "RX-010",
    "sourceId": "regex_010",
    "title": "Dateipfade - Name auslesen",
    "category": "Regex",
    "summary": "Aus einer Liste (A2:A7) mit Dateipfaden wird der Vor- und Nachname ausgelesen. Verschiedenen Pattern (Herangehensweisen).",
    "description": "Aus einer Liste (A2:A7) mit Dateipfaden wird der Vor- und Nachname ausgelesen. Verschiedenen Pattern (Herangehensweisen).\n\nZ. B. \"https://d.docs.live.net/t56bk74834x63rgu/05 - Verwaltung/01 - Arbeitszeit/Erfassung/Angestellte/2023/[2023 - Max Mustermann - Arbeitszeiten - Dokumentation.xlsx]Januar\".\nWird zu \"Max Mustermann\".\n\nDie RegEx-Patten hier mal angedeutet:\n\\[ → das öffnende [ (muss maskiert werden)\n\\d{4} → vierstellige Jahreszahl\n- → Trennzeichen\n(.+?) → möglichst kurzer Treffer (= der Name)\n- Arbeitszeiten → Ende des Namens\n\n\\[ → Beginn des Dateinamens\n[^-]+ → erster Teil (Jahr oder was auch immer)\n(.+?) → der Name\n[^]]+ → Beschreibung bis zur schließenden ]\n\\.xlsx → Dateiendung\n\n(?:.+?) → nicht-speichernde Gruppe für den ersten Teil\n(.+?) → einzige Capture-Gruppe = Name\n[^]]+ → Rest des Dateinamens bis ]\n\nWeitere Infos zu RegEx und Pattern (dort können die Pattern auch getestet werden):\nhttps://regex101.com/",
    "tags": [
      "regex",
      "string",
      "pfad",
      "datei",
      "text"
    ],
    "codeText": "=VSTAPELN(\"Name\";REGEXEXTRAHIEREN(A2:.A999;\"\\[\\d{4} - (.+?) - Arbeitszeiten\";2))\n=VSTAPELN(\"Name\";REGEXEXTRAHIEREN(A2:.A999;\"\\[.*? - (.+?) - Arbeitszeiten\";2))\n=VSTAPELN(\"Name\";REGEXEXTRAHIEREN(A2:.A999;\"\\[\\d{4} - (.+?) - .*?\\.xlsx\";2))\n=VSTAPELN(\"Name\";REGEXEXTRAHIEREN(A2:.A999;\"\\[.*? - (.+?) - .*?\\.xlsx\";2))\n=VSTAPELN(\"Name\";REGEXEXTRAHIEREN(A2:.A999;\"\\[[^-]+ - (.+?) - [^]]+\\.xlsx\";2))\n=VSTAPELN(\"Name\";REGEXEXTRAHIEREN(A2:.A999;\"\\[[^-]+ - (.+?) - [^]]+\\]\";2))\n=VSTAPELN(\"Name\";REGEXEXTRAHIEREN(A2:.A999;\"\\[(?:.+?) - (.+?) - [^]]+\\]\";2))\n=VSTAPELN(\"Name\";REGEXEXTRAHIEREN(A2:.A999;\"\\[(?:.+?) - (.+?) -\";2))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L379"
  },
  {
    "id": "RX-011",
    "sourceId": "regex_011",
    "title": "Zahlen - letzte Zahl mit Punkt darstellen",
    "category": "Regex",
    "summary": "Aus einer Liste (A2:A11) mit Zahlen wird die letzte Zahl mit Punkt angezeigt.",
    "description": "Aus einer Liste (A2:A11) mit Zahlen wird die letzte Zahl mit Punkt angezeigt.\n\nZ. B. \"172411\".\nWird zu \"17241.1\".\n\nWeitere Infos zu RegEx und Pattern (dort können die Pattern auch getestet werden):\nhttps://regex101.com/",
    "tags": [
      "regex",
      "zahl",
      "punkt",
      "zahlen",
      "spill"
    ],
    "codeText": "=VSTAPELN(A1;REGEXERSETZEN(TEXT(A2:.A999;\"0\");\"(.*)(\\d)\";\"$1.$2\"))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L429"
  },
  {
    "id": "RX-012",
    "sourceId": "regex_012",
    "title": "Umgang mit Sonderzeichen und Umlauten",
    "category": "Regex",
    "summary": "Aus einer Liste (A2:A7) mit Texten werden die Sonderzeichen getestet und entfernt. Mit den Funktionen REGEXTESTEN und REGEXERSETZEN.",
    "description": "Aus einer Liste (A2:A7) mit Texten werden die Sonderzeichen getestet und entfernt.\nMit den Funktionen REGEXTESTEN und REGEXERSETZEN.\n\nIn Spalte C:E Möglichkeiten mit unterschiedlichen Pattern zum TESTEN der Inhalte von Spalte A.\nIn Spalte G:H Möglichkeiten mit unterschiedlichen Pattern zum ERSETZEN der Inhalte von Spalte A.\n\nEinmal bleiben die Umlaute vorhanden - dann werden sie entfernt.\nDie Formeln werden in die jeweiligen Zellen in die Kommentare geschrieben.\nUnd ab L1:Lx werden die Formeln in die Zellen - mit z. B. =FORMELTEXT(C1) - geschrieben.\n\nWeitere Infos zu RegEx und Pattern (dort können die Pattern auch getestet werden):\nhttps://regex101.com/",
    "tags": [
      "regex",
      "zahl",
      "umlaute",
      "zahlen",
      "sonderzeichen"
    ],
    "codeText": "=VSTAPELN(\"Test 1\";REGEXTESTEN(A2:.A999;\"^[äöüßA-Za-z0-9\\s-]+$\"))\n=VSTAPELN(\"Test 2\";REGEXTESTEN(A2:.A999;\"[^a-zA-Z0-9\\s-]\"))\n=VSTAPELN(\"Test 3\";REGEXTESTEN(A2:.A999;\"^[a-zA-Z0-9\\s-]+$\"))\n=VSTAPELN(\"Ohne Umlaute\";GLÄTTEN(REGEXERSETZEN(A2:.A999;\"[^a-zA-Z0-9\\s-]\";\" \")))\n=VSTAPELN(\"Mit Umlaute\";GLÄTTEN(REGEXERSETZEN(A2:.A999;\"[^a-zA-ZÄÖÜäöüß0-9\\s-]\";\" \")))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L455"
  },
  {
    "id": "RX-013",
    "sourceId": "regex_013",
    "title": "Zahl - an bestimmten Stellen ein Minus einfügen",
    "category": "Regex",
    "summary": "In einer Liste (A2:A10) mit Zahlen werden an bestimmten Stellen in der Zahl ein Minus eingefügt. Mit den Funktionen REGEXTESTEN und REGEXERSETZEN.",
    "description": "In einer Liste (A2:A10) mit Zahlen werden an bestimmten Stellen in der Zahl ein Minus eingefügt.\nMit den Funktionen REGEXTESTEN und REGEXERSETZEN.\n\nAlso aus z. B. 5467121030655 wird 5467-12-103-0655.\n\nDie Formeln werden in die jeweiligen Zellen in die Kommentare geschrieben.\nUnd ab K1:Kx werden die Formeln in die Zellen - mit z. B. =FORMELTEXT(C1) - geschrieben.\n\nWeitere Infos zu RegEx und Pattern (dort können die Pattern auch getestet werden):\nhttps://regex101.com/",
    "tags": [
      "regex",
      "zahl",
      "minus",
      "zahlen",
      "einfügen"
    ],
    "codeText": "=VSTAPELN(\"Falsch:\";REGEXERSETZEN(A2:.A999;\"(\\d{4})(\\d{2})(\\d{3})(\\d{4})\";\"$1-$2-$3-$4\"))\n=VSTAPELN(\"Fast richtig:\";REGEXERSETZEN(REGEXERSETZEN(A2:.A999;\"[^\\d]\";\"\");\"(\\d{4})(\\d{2})(\\d{3})(\\d{4})\";\"$1-$2-$3-$4\"))\n=VSTAPELN(\"Möglichkeit 1:\";LET(x;REGEXERSETZEN(A2:.A999;\"[^\\d]\";\"\");WENN(REGEXTESTEN(x;\"^\\d{13}$\");REGEXERSETZEN(x;\"(\\d{4})(\\d{2})(\\d{3})(\\d{4})\";\"$1-$2-$3-$4\");\"?\")))\n=VSTAPELN(\"Möglichkeit 2:\";WENN(REGEXTESTEN(A2:.A999;\"^\\d{13}$\");REGEXERSETZEN(A2:.A999;\"(\\d{4})(\\d{2})(\\d{3})(\\d{4})\";\"$1-$2-$3-$4\");\"?\"))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L490"
  },
  {
    "id": "RX-014",
    "sourceId": "regex_014",
    "title": "Telefonnummern aus Text extrahieren",
    "category": "Regex",
    "summary": "Mehrere Telefonnummern in zwei Formaten werden aus Spalte A extrahiert. REGEXEXTRAHIEREN_Telefonnummern_mehrere_aus_Zelle_zu_extrahieren.xlsx",
    "description": "Mehrere Telefonnummern in zwei Formaten werden aus Spalte A extrahiert.\nREGEXEXTRAHIEREN_Telefonnummern_mehrere_aus_Zelle_zu_extrahieren.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Regex_14\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in C2, I2, M2.\nPattern in I2/M2 ist besser.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "regex",
      "zahl",
      "telefonnummer",
      "zahlen",
      "extrahieren"
    ],
    "codeText": "=REGEXEXTRAHIEREN(A2;\"\\(?([\\d \\-\\)\\–\\+\\/\\(]{8,})\\)?([ .\\-–\\/]?)([\\d]+)\";1)\n=REGEXEXTRAHIEREN(A2;$H$1;1)\n=EINDEUTIG(REGEXEXTRAHIEREN(A2;\"(?<!\\w)\\+?\\d(?:[\\s()./–-]*\\d){7,14}(?!\\w)\";1);WAHR)",
    "isNew": false,
    "package": {
      "fileName": "Demo_Regex_14.zip",
      "folderName": "Regex_14",
      "sizeBytes": 10646,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Regex_14.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L522"
  },
  {
    "id": "RX-015",
    "sourceId": "regex_015",
    "title": "Materialnummer - wenn in Liste - aus Text auslesen",
    "category": "Regex",
    "summary": "Materialnummern (Spalte B) werden im Text in Spalte A gesucht und wenn vorhanden ausgelesen. REGEXEXTRAHIEREN_Zahl_Materialnummer_nur_extrahieren_wenn_in_Liste_vorhanden_Vergleich.xlsx",
    "description": "Materialnummern (Spalte B) werden im Text in Spalte A gesucht und wenn vorhanden ausgelesen.\nREGEXEXTRAHIEREN_Zahl_Materialnummer_nur_extrahieren_wenn_in_Liste_vorhanden_Vergleich.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Regex_15\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in D2, E2, F2, H2, I1.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "regex",
      "zahl",
      "materialnummer",
      "zahlen",
      "auslesen"
    ],
    "codeText": "=NACHZEILE(A2:.A999;LAMBDA(t;WENNFEHLER(VERWEIS(2;1/SUCHEN(B2:.B999;t);B2:.B999);\"---\")))\n=LET(m;B2:.B999;NACHZEILE(A2:.A999;LAMBDA(t;WENNFEHLER(VERWEIS(2;1/SUCHEN(m;t);m);\"---\"))))\n=LET(m;B2:.B999;NACHZEILE(A2:.A999;LAMBDA(t;XVERWEIS(WAHR;REGEXTESTEN(t;\"(^|[^0-9])\"&m&\"([^0-9]|$)\");m;\"---\";0;-1))))\n=LET(x;WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;\"\\d+\")*1;NV());y;ISTZAHL(VERGLEICH(x;B2:.B999;0));WENN(y;x;\"---\"))\n=VSTAPELN(\"Formel 5:\";LET(x;WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;\"\\d+\")*1;NV());y;ISTZAHL(VERGLEICH(x;B2:.B999;0));WENN(y;x;\"---\")))",
    "isNew": false,
    "package": {
      "fileName": "Demo_Regex_15.zip",
      "folderName": "Regex_15",
      "sizeBytes": 10624,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Regex_15.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L560"
  },
  {
    "id": "RX-016",
    "sourceId": "regex_016",
    "title": "Großbuchstaben aus String extrahieren",
    "category": "Regex",
    "summary": "Aus Texten in Spalte A werden nur die Großbuchstaben ausgelesen. Auch Umlaute und das ß werden berücksichtigt. Formel_REGEXERSETZEN_Nur_Grossbuchstaben_ausgeben_Rest_mit_Nichts_ersetzen.xlsx",
    "description": "Aus Texten in Spalte A werden nur die Großbuchstaben ausgelesen. Auch Umlaute und das ß werden berücksichtigt.\nFormel_REGEXERSETZEN_Nur_Grossbuchstaben_ausgeben_Rest_mit_Nichts_ersetzen.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Regex_16\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in C1, D1.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "regex",
      "buchstaben",
      "sonderzeichen",
      "klein",
      "auslesen"
    ],
    "codeText": "=VSTAPELN(\"Nur Großbuchstaben, Umlaute und ß ausgeben…\";REGEXERSETZEN(A2:.A999;\"[^A-ZÄÖÜß]\";\"\"))\n=VSTAPELN(\"Gib das andere aus…\";REGEXERSETZEN(A2:.A999;\"[A-ZÄÖÜß]\";\"\"))",
    "isNew": false,
    "package": {
      "fileName": "Demo_Regex_16.zip",
      "folderName": "Regex_16",
      "sizeBytes": 10680,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Regex_16.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L599"
  },
  {
    "id": "RX-017",
    "sourceId": "regex_017",
    "title": "Letzte Zahl in mehreren Variationen auslesen",
    "category": "Regex",
    "summary": "Aus Texten in Spalte A werden nur die letzten Zahlen - die in mehreren Variationen vorkommen können - extrahiert. REGEXEXTRAHIEREN_und_andere_Formeln_letzte_Zahl_mit_oder …",
    "description": "Aus Texten in Spalte A werden nur die letzten Zahlen - die in mehreren Variationen vorkommen können - extrahiert.\nREGEXEXTRAHIEREN_und_andere_Formeln_letzte_Zahl_mit_oder ohne_Punkt_davor_Optional_in_Klammern.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Regex_17\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in C1, D1, E1, F1, G1, H1.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "regex",
      "zahl",
      "punkt",
      "klammer",
      "extrahieren"
    ],
    "codeText": "=VSTAPELN(\"REGEX\";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;\"(?:\\(\\s*)?([.,]?\\d+(?:[.,]\\d+)?)(?:\\s*\\))?\\s*$\";2);\"\"))\n=VSTAPELN(\"REGEX_1\";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;\"mm\\s*\\(?\\s*([.,]?\\d+(?:[.,]\\d+)?)\\s*\\)?\\s*$\";2);\"\"))\n=VSTAPELN(\"TEXTNACH\";GLÄTTEN(WENNFEHLER(LET(x;TEXTNACH(A2:.A999;\"mm \";-1);WECHSELN(WECHSELN(x;\"(\";\"\");\")\";\"\"));\"\")))\n=VSTAPELN(\"TEXTTEILEN\";MAP(A2:.A999;LAMBDA(a;WENN(a=\"\";\"\";LET(x;GLÄTTEN(TEXTNACH(a;\"mm\"));y;TEXTTEILEN(WECHSELN(WECHSELN(x;\"(\";\"\");\")\";\"\");\" \";;WAHR);WENNFEHLER(INDEX(y;1;1);\"\"))))))\n=VSTAPELN(\"TEXTTEILEN_1\";MAP(A2:.A999;LAMBDA(a;WENN(a=\"\";\"\";LET(x;GLÄTTEN(TEXTNACH(a;\"mm\"));y;WECHSELN(WECHSELN(x;\"(\";\"\");\")\";\"\");z;TEXTTEILEN(y;\" \";;WAHR);WENNFEHLER(INDEX(z;1);\"\"))))\n=VSTAPELN(\"XMLFILTERN\";MAP(A2:.A999;LAMBDA(a;WENN(a=\"\";\"\";LET(x;GLÄTTEN(TEXTNACH(a;\"mm\"));y;WECHSELN(WECHSELN(x;\"(\";\"\");\")\";\"\");WENN(GLÄTTEN(y)=\"\";\"\";LET(z;WECHSELN(GLÄTTEN(y);\" \";\"</s><s>_\");r;XMLFILTERN(\"<t><s>_\"&z&\"</s></t>\";\"//s[last()]\");WECHSELN(r;\"_\";\"\"))))))))",
    "isNew": false,
    "package": {
      "fileName": "Demo_Regex_17.zip",
      "folderName": "Regex_17",
      "sizeBytes": 11476,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Regex_17.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L635"
  },
  {
    "id": "RX-018",
    "sourceId": "regex_018",
    "title": "Datum UTC - MEZ - MESZ...",
    "category": "Regex",
    "summary": "Aus Daten in Spalte A (UTC) werden die Daten und Zeiten ausgelesen. UTC_MEZ_MESZ_REGEXEXTRAHIEREN_und_mehr.xlsx",
    "description": "Aus Daten in Spalte A (UTC) werden die Daten und Zeiten ausgelesen.\nUTC_MEZ_MESZ_REGEXEXTRAHIEREN_und_mehr.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Regex_18\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in C2, D2, F2, G2, I2, L2 und N2.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "regex",
      "zahl",
      "mez",
      "mesz",
      "datum",
      "utc"
    ],
    "codeText": "=REGEXEXTRAHIEREN(A2:.A989;\"\\d{4}-\\d{2}-\\d{2}\")\n=REGEXEXTRAHIEREN(A2:.A989;\"\\d{2}:\\d{2}:\\d{2}\")\n=--REGEXEXTRAHIEREN(A2:.A989;\"\\d{4}-\\d{2}-\\d{2}\")\n=--REGEXEXTRAHIEREN(A2:.A989;\"\\d{2}:\\d{2}:\\d{2}\")\n=--REGEXEXTRAHIEREN(A2;\"(\\d{4}-\\d{2}-\\d{2})T(\\d{2}:\\d{2}:\\d{2})\";2)\n=LET(x;--WECHSELN(LINKS(A2;19);\"T\";\" \");j;JAHR(x);SoM;DATUM(j;4;1)-WOCHENTAG(DATUM(j;4;1);2);SoO;DATUM(j;11;1)-WOCHENTAG(DATUM(j;11;1);2);x+WENN(UND(x>=SoM+1/24;x<SoO+1/24);2/24;1/24))\n=LET(d;REGEXEXTRAHIEREN(A2;\"\\d{4}-\\d{2}-\\d{2}\");t;REGEXEXTRAHIEREN(A2;\"\\d{2}:\\d{2}:\\d{2}\");x;--d+--t;j;JAHR(x);SoM;DATUM(j;4;1)-WOCHENTAG(DATUM(j;4;1);2);SoO;DATUM(j;11;1)-WOCHENTAG(DATUM(j;11;1);2);x+WENN(UND(x>=SoM+1/24;x<SoO+1/24);2/24;1/24))",
    "isNew": false,
    "package": {
      "fileName": "Demo_Regex_18.zip",
      "folderName": "Regex_18",
      "sizeBytes": 10517,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Regex_18.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L675"
  },
  {
    "id": "RX-019",
    "sourceId": "regex_019",
    "title": "Name - Alter - Status auslesen",
    "category": "Regex",
    "summary": "Aus Daten in Spalte A (z. B. Bruno Graf (von) Bayern (72) - Abwesend) werden Name, Alter und Status ausgelesen. REGEXEXTRAHIEREN_Name_Alter_Status_auslesen.xlsx",
    "description": "Aus Daten in Spalte A (z. B. Bruno Graf (von) Bayern (72) - Abwesend) werden Name, Alter und Status ausgelesen.\nREGEXEXTRAHIEREN_Name_Alter_Status_auslesen.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Regex_19\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in C2, D2, E2, C15, C18, C19, C22, C36, C50 und C64.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "regex",
      "zahl",
      "name",
      "alter",
      "status",
      "auslesen"
    ],
    "codeText": "=REGEXEXTRAHIEREN(A2;\"^.*?(?=\\s*\\(\\s*\\d+)\")\n=--REGEXEXTRAHIEREN(A2;\"\\d+\")\n=WENNFEHLER(REGEXEXTRAHIEREN(A2;\"(?<=[-–]\\s).*$\");\"\")\n=REGEXEXTRAHIEREN(A2;\"^\\s*(.*?)\\s*\\(\\s*(\\d+)(?:\\s+Jahre)?\\s*\\)\\s*(?:[-–]\\s*)?(.*)$\";2)\n=LET(x;REGEXEXTRAHIEREN(A2;\"^\\s*(.*?)\\s*\\(\\s*(\\d+)(?:\\s+Jahre)?\\s*\\)\\s*(?:[-–]\\s*)?(.*)$\";2);HSTAPELN(INDEX(x;1);--INDEX(x;2);INDEX(x;3)))\n=LET(x;REGEXEXTRAHIEREN(A2;\"^\\s*(.*?)\\s*\\(\\s*(\\d+)(?:\\s+Jahre)?\\s*\\)\\s*(?:[-–]\\s*)?(.*)$\";2);HSTAPELN(SPALTENWAHL(x;1);--SPALTENWAHL(x;2);SPALTENWAHL(x;3)))\n=VSTAPELN(HSTAPELN(\"Name\";\"Alter\";\"Status\");LET(a;A2:.A1000;WEGLASSEN(REDUCE(\"\";a;LAMBDA(x;y;VSTAPELN(x;REGEXEXTRAHIEREN(y;\"^\\s*(.*?)\\s*\\(\\s*(\\d+)(?:\\s+Jahre)?\\s*\\)\\s*(?:[-–]\\s*)?(.*)$\";2))));1)))\n=LET(x;A2:.A1000;r;WEGLASSEN(REDUCE(\"\";x;LAMBDA(a;z;VSTAPELN(a;REGEXEXTRAHIEREN(z;\"^\\s*(.*?)\\s*\\(\\s*(\\d+)(?:\\s+Jahre)?\\s*\\)\\s*(?:[-–]\\s*)?(.*)$\";2))));1);VSTAPELN(HSTAPELN(\"Name\";\"Alter\";\"Status\");HSTAPELN(SPALTENWAHL(r;1);--SPALTENWAHL(r;2);SPALTENWAHL(r;3))))\n=LET(x;A2:.A1000;r;WEGLASSEN(REDUCE(\"\";x;LAMBDA(a;z;VSTAPELN(a;REGEXEXTRAHIEREN(z;\"^\\s*(.*?)\\s*\\(\\s*(\\d+)(?:\\s+Jahre)?\\s*\\)\\s*(?:[-–]\\s*)?(.*)$\";2))));1);VSTAPELN({\"Name\".\"Alter\".\"Status\"};HSTAPELN(INDEX(r;;1);--INDEX(r;;2);INDEX(r;;3))))\n=LET(x;A2:.A1000;VSTAPELN(HSTAPELN(\"Name\";\"Alter\";\"Status\");HSTAPELN(MAP(x;LAMBDA(z;GLÄTTEN(REGEXEXTRAHIEREN(z;\"^(.*?)\\s*\\(\\s*\\d+(?:\\s*Jahre)?\\s*\\)\";2))));MAP(x;LAMBDA(z;--REGEXEXTRAHIEREN(z;\"^.*\\(\\s*(\\d+)(?:\\s*Jahre)?\\s*\\)\";2)));MAP(x;LAMBDA(z;WENNFEHLER(GLÄTTEN(REGEXEXTRAHIEREN(z;\"^.*\\)\\s*[-–]\\s*(.*)$\";2));\"\"))))))",
    "isNew": false,
    "package": {
      "fileName": "Demo_Regex_19.zip",
      "folderName": "Regex_19",
      "sizeBytes": 13045,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Regex_19.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L716"
  },
  {
    "id": "RX-020",
    "sourceId": "regex_020",
    "title": "Zahlen mit 1.000er Punkt und Text am Ende",
    "category": "Regex",
    "summary": "Daten in A1:F8. In Spalte G werden die Werte aus den jeweiligen Zeilen berechnet. Aus H (Haben) und S (Soll).",
    "description": "Daten in A1:F8.\nIn Spalte G werden die Werte aus den jeweiligen Zeilen berechnet.\nAus H (Haben) und S (Soll).\n\nIn I1 folgende (Spill-Formel) werden die Daten ausgegeben.\nMit Minus (-) wenn S (Soll) am Ende steht.\n\n1.000er Punkt, Positiv, Negativ und 0 über Benutzerdefiniertes Format.\n#.##0,00;-#.##0,00;0\nREGEXEXTRAHIEREN_Zahlen_mit_Punkt_und_Text_am_Ende_auslesen_bearbeiten.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Regex_19\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in G1 und I1.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "regex",
      "zahl",
      "punkt",
      "zahlen",
      "text",
      "auslesen"
    ],
    "codeText": "=VSTAPELN(\"Ergebnis\";NACHZEILE(C2:.F970;LAMBDA(y;LET(z;WENNNV(--REGEXEXTRAHIEREN(y;\"\\d{1,3}(?:[.,]\\d{3})*(?:[.,]\\d+)?\");0);e;SUMME(z*WENN(REGEXTESTEN(y;\"S\\s*$\");-1;1));WENN(e=0;0;ABS(e)&\" \"&WENN(e<0;\"S\";\"H\"))))))\n=LET(x;C2:.G970;z;WENNNV(--REGEXEXTRAHIEREN(x;\"\\d+(?:\\.\\d{3})*(?:,\\d+)?\");\"\");VSTAPELN(C1:G1;WENN(z=\"\";\"\";z*WENN(REGEXTESTEN(x;\"S\\s*$\");-1;1))))",
    "isNew": false,
    "package": {
      "fileName": "Demo_Regex_20.zip",
      "folderName": "Regex_20",
      "sizeBytes": 11482,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Regex_20.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L760"
  },
  {
    "id": "RX-021",
    "sourceId": "regex_021",
    "title": "Langer Text in Zellen aufteilen...",
    "category": "Regex",
    "summary": "Lange Texte in Zellen werden nach vorgegebener Anzahl von Zeichen aufgetrennt. Nicht mitten im Wort. Verschiedene Möglichkeiten (in Spalten, Zeilen, Blöcken...). Mit RegEx, Power Query, …",
    "description": "Lange Texte in Zellen werden nach vorgegebener Anzahl von Zeichen aufgetrennt.\nNicht mitten im Wort.\nVerschiedene Möglichkeiten (in Spalten, Zeilen, Blöcken...).\nMit RegEx, Power Query, Formeln und VBA (UDF und Sub).\n\nRegEx_Power_Query_Formel_VBA_Text_auftrennen.xlsb\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Demo_RegEx_PQ_Formel_VBA\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "auftrennenl",
      "text",
      "aufteilen",
      "formel",
      "regex",
      "vba",
      "power query"
    ],
    "codeText": "=LET(x;A2:.A999;WENNFEHLER(WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;T;VSTAPELN(A;GLÄTTEN(REGEXEXTRAHIEREN(T;\".{1,25}(?=\\s|$)\";1;1)))));1);\"\"))\n=LET(x;A2:.A999;lg;D1;WENNFEHLER(WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;T;VSTAPELN(A;GLÄTTEN(REGEXEXTRAHIEREN(T;\".{1,\"&lg&\"}(?=\\s|$)\";1;1)))));1);\"\"))\n=LET(x;A2:.A999;lg;D1;WENNFEHLER(WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;T;VSTAPELN(A;GLÄTTEN(REGEXEXTRAHIEREN(T;\"\\S{\"&(lg+1)&\",}|.{1,\"&lg&\"}(?=\\s|$)\";1;1)))));1);\"\"))\n\n=LET(x;A2;y;$C$1;z;\"\\S{\"&(y+1)&\",}|.{1,\"&y&\"}(?=\\s|$)\";MTRANS(GLÄTTEN(REGEXEXTRAHIEREN(x;z;1;1))))\n=LET(x;A2:.A999;lg;C1;WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;z;VSTAPELN(A;MTRANS(GLÄTTEN(REGEXEXTRAHIEREN(z;\"\\S{\"&(lg+1)&\",}|.{1,\"&lg&\"}(?=\\s|$)\";1;1))))));1))\n=LET(x;A2:.A999;y;$C$1;z;\"\\S{\"&(y+1)&\",}|.{1,\"&y&\"}(?=\\s|$)\";WENNNV(WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;T;HSTAPELN(A;MTRANS(GLÄTTEN(REGEXEXTRAHIEREN(T;z;1;1))))));;1);\"\"))",
    "isNew": true,
    "package": {
      "fileName": "Demo_RegEx_PQ_Formel_VBA.zip",
      "folderName": "Demo_RegEx_PQ_Formel_VBA",
      "sizeBytes": 73867,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_RegEx_PQ_Formel_VBA.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Regex.vb#L804"
  },
  {
    "id": "VB-001",
    "sourceId": "vba_001",
    "title": "Alle ENVIRON Variablen ausgeben",
    "category": "VBA",
    "summary": "Alle ENVIRON Variablen mit Inhalt werden in einer neuen Datei aufgelistet. Wie z. B. USERNAME, USERPROFILE, TEMP...",
    "description": "Alle ENVIRON Variablen mit Inhalt werden in einer neuen Datei aufgelistet.\nWie z. B. USERNAME, USERPROFILE, TEMP...\n\n!!!!!!!!WICHTIG!!!!!!!!\nDen Code über den Button \"Code kopieren\" in die Zwischenablage übernehmen und dann im VBA-Editor ein Modul einfügen und dort den Code hineinkopieren.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "vba",
      "environ",
      "variablen",
      "ausgeben",
      "text"
    ],
    "codeText": "Option Explicit\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' https://github.com/rstsu/Excel-VSTO-Toolbox\nPublic Sub Main()\n    Dim wkbBook As Workbook\n    Dim strTMP() As String\n    Dim lngTMP As Long\n    On Error GoTo Fin\n    lngTMP = 1\n    Set wkbBook = Workbooks.Add(1)\n    With wkbBook.Worksheets(lngTMP)\n        Do\n            strTMP = Split(Environ(lngTMP), \"=\")\n            If Join(strTMP) <> \"\" Then\n                .Cells(lngTMP, 1).Value = strTMP(0)\n                .Cells(lngTMP, 2).Value = strTMP(1)\n                lngTMP = lngTMP + 1\n            End If\n        Loop Until Join(strTMP) = \"\"\n        .Columns(\"A:B\").AutoFit\n    End With\nFin:\n    Set wkbBook = Nothing\n    If Err.Number <> 0 Then MsgBox \"Fehler: \" & Err.Number & \" \" & Err.Description\nEnd Sub",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L6"
  },
  {
    "id": "VB-002",
    "sourceId": "vba_002",
    "title": "Alle Filter aller Tabellenblätter zurücksetzen",
    "category": "VBA",
    "summary": "Alle gesetzten Filter werden beim beenden der Datei zurückgesetzt. Das passiert auf allen Tabellenblättern.",
    "description": "Alle gesetzten Filter werden beim beenden der Datei zurückgesetzt.\nDas passiert auf allen Tabellenblättern.\n\n!!!!!!!!WICHTIG!!!!!!!!\nDer Code gehört unter \"DieseArbeitsmappe\" - NICHT in ein Modul.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "vba",
      "filter",
      "tabellenblätter",
      "zurücksetzen",
      "tabellenblatt"
    ],
    "codeText": "Option Explicit\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' https://github.com/rstsu/Excel-VSTO-Toolbox\nPrivate Sub Workbook_BeforeClose(Cancel As Boolean)\n    Dim wksSheet As Worksheet\n    For Each wksSheet In ThisWorkbook.Worksheets\n        With wksSheet\n            ' Falls die Tabellenblätter mit einem Paaswort geschützt sind\n            '.Unprotect Password:=\"DEINPASSWORT\"\n            If .AutoFilterMode Then\n                If .FilterMode Then\n                    .ShowAllData\n                End If\n            End If\n            ' Passwortschutz des Tabellenblattes wieder setzen\n            ' UserInterfaceOnly auf True bedeutet - VBA kann Änderungen vornehmen, ohne dass der Blattschutz entfernt werden muss\n            '.Protect Password:=\"DEINPASSWORT\", UserInterfaceOnly:=True\n        End With\n    Next wksSheet\n    ThisWorkbook.Save\nEnd Sub",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L55"
  },
  {
    "id": "VB-003",
    "sourceId": "vba_003",
    "title": "UserForm zur Laufzeit erstellen",
    "category": "VBA",
    "summary": "Es wird zur Laufzeit eine UserForm mit \"OptionButtons\" und \"CommansButton\" erstellt und gleich angezeigt. Dort kann eine Auswahl getroffen werden. Dann schließt sich die UserForm (Klick …",
    "description": "Es wird zur Laufzeit eine UserForm mit \"OptionButtons\" und \"CommansButton\" erstellt und gleich angezeigt.\nDort kann eine Auswahl getroffen werden.\nDann schließt sich die UserForm (Klick auf CommandButton \"OK\")und wird auch wieder entfernt, also keine UserForm vorhanden.\n\n!!!!!!!!WICHTIG!!!!!!!!\nDatei - Optionen - Trust Center - Einstellungen für das Trust Center... - Makroeinstellungen - \"Zugriff auf das VBA-Projektmodell vertrauen\" - dieser Haken MUSS gesetzt sein, damit der Code funktioniert.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "vba",
      "userform",
      "laufzeit",
      "erstellen",
      "ausführen"
    ],
    "codeText": "Option Explicit\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' https://github.com/rstsu/Excel-VSTO-Toolbox\nPublic Sub Main()\n  Dim strC As String\n  strC = strC & vbLf & \"Private WithEvents Opt1 as MSForms.OptionButton\"\n  strC = strC & vbLf & \"Private WithEvents Opt2 as MSForms.OptionButton\"\n  strC = strC & vbLf & \"Private WithEvents Opt3 as MSForms.OptionButton\"\n  strC = strC & vbLf & \"Private WithEvents Cmd1 as MSForms.CommandButton\"\n  strC = strC & vbLf\n  strC = strC & vbLf & \"Private Sub UserForm_Initialize\"\n  strC = strC & vbLf & \"  With Me\"\n  strC = strC & vbLf & \"    .Caption = \"\"Abfrage\"\"\"\n  strC = strC & vbLf & \"    .Width = 190\"\n  strC = strC & vbLf & \"    With .Controls.Add(\"\"Forms.Label.1\"\")\"\n  strC = strC & vbLf & \"      .Top = 10\"\n  strC = strC & vbLf & \"      .Left = 20\"\n  strC = strC & vbLf & \"      .Width = 200\"\n  strC = strC & vbLf & \"      .Font.Size = 10\"\n  strC = strC & vbLf & \"      .Caption = \"\"Bitte wählen Sie aus den Optionen\"\"\"\n  strC = strC & vbLf & \"    End With\"\n  strC = strC & vbLf & \"    Set Opt1 = .Controls.Add(\"\"Forms.OptionButton.1\"\")\"\n  strC = strC & vbLf & \"    With  Opt1\"\n  strC = strC & vbLf & \"     .Top = 40\"\n  strC = strC & vbLf & \"     .Left = 40\"\n  strC = strC & vbLf & \"     .Caption = \"\"Drucken\"\"\"\n  strC = strC & vbLf & \"    End With\"\n  strC = strC & vbLf & \"    Set Opt2 = .Controls.Add(\"\"Forms.OptionButton.1\"\")\"\n  strC = strC & vbLf & \"    With Opt2\"\n  strC = strC & vbLf & \"      .Top = 60\"\n  strC = strC & vbLf & \"      .Left = 40\"\n  strC = strC & vbLf & \"      .Caption = \"\"Speichern\"\"\"\n  strC = strC & vbLf & \"    End With\"\n  strC = strC & vbLf & \"    Set Opt3 = .Controls.Add(\"\"Forms.OptionButton.1\"\")\"\n  strC = strC & vbLf & \"    With Opt3\"\n  strC = strC & vbLf & \"      .Top = 80\"\n  strC = strC & vbLf & \"      .Left = 40\"\n  strC = strC & vbLf & \"      .Caption = \"\"Drucken und Speichern\"\"\"\n  strC = strC & vbLf & \"    End With\"\n  strC = strC & vbLf & \"    Set Cmd1 = .Controls.Add(\"\"Forms.CommandButton.1\"\")\"\n  strC = strC & vbLf & \"    With Cmd1\"\n  strC = strC & vbLf & \"      .Top = 120\"\n  strC = strC & vbLf & \"      .Left = 35\"\n  strC = strC & vbLf & \"      .Width = 115\"\n  strC = strC & vbLf & \"      .Caption = \"\"OK\"\"\"\n  strC = strC & vbLf & \"    End With\"\n  strC = strC & vbLf & \"  End With\"\n  strC = strC & vbLf & \"End Sub\"\n  strC = strC & vbLf\n  strC = strC & vbLf & \"Private Sub Cmd1_Click\"\n  strC = strC & vbLf & \"  If Opt1 Then\"\n  strC = strC & vbLf & \"    Call Drucken\"\n  strC = strC & vbLf & \"    Unload Me\"\n  strC = strC & vbLf & \"  ElseIf Opt2 Then\"\n  strC = strC & vbLf & \"    Call Speichern\"\n  strC = strC & vbLf & \"    Unload Me\"\n  strC = strC & vbLf & \"  ElseIf Opt3 Then\"\n  strC = strC & vbLf & \"    Call Drucken\"\n  strC = strC & vbLf & \"    Call Speichern\"\n  strC = strC & vbLf & \"    Unload Me\"\n  strC = strC & vbLf & \"  Else\"\n  strC = strC & vbLf & \"    MsgBox \"\"Bitte wählen Sie eine Option\"\", vbExclamation, \"\"Nochmal\"\"\"\n  strC = strC & vbLf & \"  End If\"\n  strC = strC & vbLf & \"End Sub\"\n  With Application.VBE.ActiveVBProject\n    With .VBComponents.Add(3)\n      .CodeModule.InsertLines .CodeModule.CountOfLines + 1, strC\n      VBA.UserForms.Add(.Name).Show\n      .Collection.Remove .CodeModule.Parent\n    End With\n  End With\nEnd Sub\nPublic Sub Drucken()\n'Code zum drucken\n  MsgBox \"Drucken\"\nEnd Sub\nPublic Sub Speichern()\n'Code zum speichern\n  MsgBox \"Speichern\"\nEnd Sub",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L100"
  },
  {
    "id": "VB-004",
    "sourceId": "vba_004",
    "title": "Kreuztabelle aus Liste erste 3 Buchstaben",
    "category": "VBA",
    "summary": "Aus einer Liste (A2:A15) wird eine Kreuztabelle erstellt. Grundlage sind die ersten 3 gleichen Buchstaben. Es ist auch in Formeln und Power Query gelöst. Mit der gleichen Bezeichnung. …",
    "description": "Aus einer Liste (A2:A15) wird eine Kreuztabelle erstellt.\nGrundlage sind die ersten 3 gleichen Buchstaben.\nEs ist auch in Formeln und Power Query gelöst. Mit der gleichen Bezeichnung.\nDie Aufgabe wurde mit einfachen Werten und einer \"intelligenten Tabelle\" gelöst.\n\n!!!!!!!!WICHTIG!!!!!!!!\nDen Code über den Button \"Code kopieren\" in die Zwischenablage übernehmen und dann im VBA-Editor ein Modul einfügen und dort den Code hineinkopieren.\nErst dann funktionieren die beiden Buttons im Tabellenblatt.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "vba",
      "text",
      "kreuztabelle",
      "intelligente tabelle",
      "liste"
    ],
    "codeText": "Option Explicit\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' https://github.com/rstsu/Excel-VSTO-Toolbox\nPublic Sub Main_1()\n    Dim strPrefix As String\n    Dim vaArrZ() As Variant\n    Dim varArrQ As Variant\n    Dim lngMaxCol As Long\n    Dim lngCount As Long\n    Dim lngGroup As Long\n    Dim strTMP As String\n    Dim lngRow As Long\n    Dim lngCol As Long\n    With ThisWorkbook.Worksheets(\"Demo_VBA_4\")\n        .Range(\"G1:M\" & .Rows.Count).Clear\n        varArrQ = Range(\"C2\", Cells(Rows.Count, \"C\").End(xlUp)).Value2\n        lngGroup = 1\n        lngCol = 1\n        strTMP = Left$(varArrQ(1, 1), 3)\n        For lngCount = 2 To UBound(varArrQ, 1)\n            strPrefix = Left$(varArrQ(lngCount, 1), 3)\n            If strPrefix = strTMP Then\n                lngCol = lngCol + 1\n            Else\n                If lngCol > lngMaxCol Then lngMaxCol = lngCol\n                lngGroup = lngGroup + 1\n                lngCol = 1\n                strTMP = strPrefix\n            End If\n        Next lngCount\n        If lngCol > lngMaxCol Then lngMaxCol = lngCol\n        ReDim vaArrZ(1 To lngGroup + 1, 1 To lngMaxCol)\n        For lngCount = 1 To lngMaxCol\n            vaArrZ(1, lngCount) = \"Pos\" & lngCount\n        Next\n        lngRow = 2\n        lngCol = 1\n        strTMP = Left$(varArrQ(1, 1), 3)\n        vaArrZ(lngRow, lngCol) = varArrQ(1, 1)\n        For lngCount = 2 To UBound(varArrQ, 1)\n            strPrefix = Left$(varArrQ(lngCount, 1), 3)\n            If strPrefix = strTMP Then\n                lngCol = lngCol + 1\n            Else\n                lngRow = lngRow + 1\n                lngCol = 1\n                strTMP = strPrefix\n            End If\n            vaArrZ(lngRow, lngCol) = varArrQ(lngCount, 1)\n        Next lngCount\n        .Range(\"G1\").Resize(UBound(vaArrZ, 1), UBound(vaArrZ, 2)).Value = vaArrZ\n    End With\nEnd Sub\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' https://github.com/rstsu/Excel-VSTO-Toolbox\nPublic Sub Main_2()\n    Dim objList As ListObject\n    Dim strPrefix As String\n    Dim vaArrZ() As Variant\n    Dim varArrQ As Variant\n    Dim lngMaxCol As Long\n    Dim lngCount As Long\n    Dim lngGroup As Long\n    Dim strTMP As String\n    Dim lngRow As Long\n    Dim lngCol As Long\n    With ThisWorkbook.Worksheets(\"Demo_VBA_4\")\n        Set objList = .ListObjects(\"Demo_VBA_4\")\n        .Range(\"G1:M\" & .Rows.Count).Clear\n        varArrQ = objList.ListColumns(1).DataBodyRange.Value2\n        lngGroup = 1\n        lngCol = 1\n        strTMP = Left$(varArrQ(1, 1), 3)\n        For lngCount = 2 To UBound(varArrQ, 1)\n            strPrefix = Left$(varArrQ(lngCount, 1), 3)\n            If strPrefix = strTMP Then\n                lngCol = lngCol + 1\n            Else\n                If lngCol > lngMaxCol Then lngMaxCol = lngCol\n                lngGroup = lngGroup + 1\n                lngCol = 1\n                strTMP = strPrefix\n            End If\n        Next lngCount\n        If lngCol > lngMaxCol Then lngMaxCol = lngCol\n        ReDim vaArrZ(1 To lngGroup + 1, 1 To lngMaxCol)\n        For lngCount = 1 To lngMaxCol\n            vaArrZ(1, lngCount) = \"Pos\" & lngCount\n        Next\n        lngRow = 2\n        lngCol = 1\n        strTMP = Left$(varArrQ(1, 1), 3)\n        vaArrZ(lngRow, lngCol) = varArrQ(1, 1)\n        For lngCount = 2 To UBound(varArrQ, 1)\n            strPrefix = Left$(varArrQ(lngCount, 1), 3)\n            If strPrefix = strTMP Then\n                lngCol = lngCol + 1\n            Else\n                lngRow = lngRow + 1\n                lngCol = 1\n                strTMP = strPrefix\n            End If\n            vaArrZ(lngRow, lngCol) = varArrQ(lngCount, 1)\n        Next\n        .Range(\"G1\").Resize(UBound(vaArrZ, 1), UBound(vaArrZ, 2)).Value = vaArrZ\n    End With\nEnd Sub",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L205"
  },
  {
    "id": "VB-005",
    "sourceId": "vba_005",
    "title": "Links in klickbare Hyperlinks umwandeln",
    "category": "VBA",
    "summary": "Eine Liste (A2:A7) mit Links wird in klickbare Hyperlinks umgewandelt (Main_1). Die Texte werden in Spalte C als klickbare Hyperlinks ausgegeben (Main_2). Die Hyperlinks (nicht der Text) …",
    "description": "Eine Liste (A2:A7) mit Links wird in klickbare Hyperlinks umgewandelt (Main_1).\nDie Texte werden in Spalte C als klickbare Hyperlinks ausgegeben (Main_2).\nDie Hyperlinks (nicht der Text) in Spalte A werden entfernt. Und die Links in Spalte C komplett (Main_3).\n\n!!!!!!!!WICHTIG!!!!!!!!\nDen Code über den Button \"Code kopieren\" in die Zwischenablage übernehmen und dann im VBA-Editor ein Modul einfügen und dort den Code hineinkopieren.\nErst dann funktionieren die drei Buttons im Tabellenblatt.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "vba",
      "link",
      "hyperlink",
      "umwandeln",
      "liste"
    ],
    "codeText": "Option Explicit\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' https://github.com/rstsu/Excel-VSTO-Toolbox\n\n' Die Links in Spalte A werden in klickbare Hyperlinks umgewandelt\n' Es wird das aktive Tabellenblatt genommen\nPublic Sub Main_1_1()\n    Dim rngCell As Range\n    For Each rngCell In Range(Range(\"A2\"), Cells(Rows.Count, 1).End(xlUp))\n        rngCell.Parent.Hyperlinks.Add Anchor:=rngCell, Address:=rngCell.Value, TextToDisplay:=rngCell.Value\n    Next rngCell\nEnd Sub\n' Oder trage die Formmel per VBA ein - hier in Spalte C, Links stehen in Spalte A\n' Es wird das aktive Tabellenblatt genutzt\nPublic Sub Main_2_1()\n    Range(\"C2:C\" & Cells(Rows.Count, 1).End(xlUp).Row).Formula2 = \"=HYPERLINK(A2,A2)\"\n    Columns(\"C\").Autofit\nEnd Sub\n' Hyperlinks in Spalte A werden entfernt\n' Auch hier das aktive Tabellenblatt\nPublic Sub Main_3_1()\n    Range(\"A2:A\" & Cells(Rows.Count, 1).End(xlUp).Row).Hyperlinks.Delete\n    Range(\"C2:C\" & Cells(Rows.Count, 1).End(xlUp).Row).ClearContents\nEnd Sub",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L340"
  },
  {
    "id": "VB-006",
    "sourceId": "vba_006",
    "title": "Access Datenbank erstellen...",
    "category": "VBA",
    "summary": "Eine Access-Datenbank komplett per VBA erzeugen – inklusive Tabelle, Autowert-Primärschlüssel und Index. Die Datenbank wird am Ende geöffnet. Die Datei wird im lokalenTemp-Ordner erstellt.",
    "description": "Eine Access-Datenbank komplett per VBA erzeugen – inklusive Tabelle, Autowert-Primärschlüssel und Index. Die Datenbank wird am Ende geöffnet.\nDie Datei wird im lokalenTemp-Ordner erstellt.\n\n!!!!!!!!WICHTIG!!!!!!!!\nDen Code über den Button \"Code kopieren\" in die Zwischenablage übernehmen und dann im VBA-Editor ein Modul einfügen und dort den Code hineinkopieren.\nErst dann funktioniert der Button im Tabellenblatt.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "vba",
      "access",
      "adox",
      "ace",
      "oledb"
    ],
    "codeText": "Option Explicit\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' https://github.com/rstsu/Excel-VSTO-Toolbox\nPrivate Declare PtrSafe Function ShellExecute Lib \"shell32.dll\" Alias \"ShellExecuteA\" ( _\n        ByVal hwnd As LongPtr, _\n        ByVal lpOperation As String, _\n        ByVal lpFile As String, _\n        ByVal lpParameters As String, _\n        ByVal lpDirectory As String, _\n        ByVal nShowCmd As Long) As LongPtr\nPrivate Const SW_MAXIMIZE = 3\nPublic Sub Main_Access_1()\n    Dim strDescription As String\n    Dim varProvider As Variant\n    Dim strFileName As String\n    Dim strProvider As String\n    Dim catCatalog As Object\n    Dim objTable As Object\n    Dim objIndex As Object\n    Dim objConn As Object\n    Dim lngError As Long\n    strFileName = Environ$(\"TEMP\") & \"\\VBA_Demo_6.accdb\"\n    On Error GoTo Fin\n    If Len(Dir$(strFileName)) > 0 Then Kill strFileName\n    Set catCatalog = CreateObject(\"ADOX.Catalog\")\n    For Each varProvider In Array(\"Microsoft.ACE.OLEDB.16.0\", \"Microsoft.ACE.OLEDB.12.0\")\n        Err.Clear\n        On Error Resume Next\n        catCatalog.Create \"Provider=\" & CStr(varProvider) & \";Data Source=\" & strFileName & \";\"\n        If Err.Number = 0 Then\n            strProvider = CStr(varProvider)\n            On Error GoTo Fin\n            Exit For\n        End If\n        lngError = Err.Number\n        strDescription = Err.Description\n        On Error GoTo Fin\n    Next varProvider\n    If Len(strProvider) = 0 Then\n        Err.Raise vbObjectError + 1000, _\n            \"CreateDataBase_1\", _\n            \"Es wurde kein geeigneter ACE-OLEDB-Provider gefunden.\" & _\n            vbCrLf & vbCrLf & _\n            \"Letzter Fehler: \" & lngError & _\n            vbCrLf & strDescription\n    End If\n    Set objConn = CreateObject(\"ADODB.Connection\")\n    With objConn\n        .CursorLocation = 3       'adUseClient\n        .Provider = strProvider\n        .Properties(\"Data Source\") = strFileName\n        .Open\n    End With\n    Set objTable = CreateObject(\"ADOX.Table\")\n    With objTable\n        .Name = \"Lieferanten\"\n        .ParentCatalog = catCatalog\n        .Columns.Append \"Primaer\", 3   'adInteger\n        With .Columns(\"Primaer\")\n            .Properties(\"Description\") = \"Schluessel\"\n            .Properties(\"Autoincrement\") = True\n        End With\n        .Columns.Append \"Name\", 202, 60   'adVarWChar\n        With .Columns(\"Name\")\n            .Properties(\"Description\") = \"Nachname\"\n            .Properties(\"Jet OLEDB:Allow Zero Length\") = True\n            .Properties(\"Nullable\") = True\n        End With\n    End With\n    catCatalog.Tables.Append objTable\n    Set objIndex = CreateObject(\"ADOX.Index\")\n    With objIndex\n        .Name = \"PrimaryKey\"\n        .Columns.Append \"Primaer\"\n        .PrimaryKey = True\n        .Unique = True\n    End With\n    objTable.Indexes.Append objIndex\n    MsgBox \"Datenbank wurde erstellt:\" & vbCrLf & _\n        strFileName & vbCrLf & vbCrLf & _\n        \"Provider: \" & strProvider, vbInformation\n    If objConn.State = 1 Then objConn.Close\n    If ShellExecute(Application.hwnd, \"Open\", strFileName, vbNullString, vbNullString, SW_MAXIMIZE) <= 32 Then\n        MsgBox \"Die Datenbank wurde erstellt, konnte aber nicht geöffnet werden.\", vbExclamation\n    End If\nFin:\n    If Err.Number <> 0 Then MsgBox \"Fehler: \" & Err.Number & vbCrLf & Err.Description, vbExclamation\n    If Not objConn Is Nothing Then\n        If objConn.State = 1 Then objConn.Close\n    End If\n    Set objIndex = Nothing\n    Set objTable = Nothing\n    Set catCatalog = Nothing\n    Set objConn = Nothing\nEnd Sub",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L390"
  },
  {
    "id": "VB-007",
    "sourceId": "vba_007",
    "title": "Spalte nach Vorgabe aufteilen - mit UDF und Formeln",
    "category": "VBA",
    "summary": "Spalte A (mit Überschrift) und Spalte C (ohne Überschrift) werden nach Vorgaben aufgeteilt. Mit einer UDF (Tabellenblatt UDF) und mit Formeln (Tabellenblatt Formeln). \"Alte\" UDF mit …",
    "description": "Spalte A (mit Überschrift) und Spalte C (ohne Überschrift) werden nach Vorgaben aufgeteilt.\nMit einer UDF (Tabellenblatt UDF) und mit Formeln (Tabellenblatt Formeln).\n\"Alte\" UDF mit \"neuen\" Formeln ersetzt. ;-)\n\n=fncSplit(\"A\")          3er-Gruppen pro Zeile, keine Überschrift\n=fncSplit(\"A\";4)        4er-Gruppen pro Zeile, keine Überschrift\n=fncSplit(\"A\";;0)       3 Zeilen pro Spalte, Überschrift\n=fncSplit(\"A\"; 3; 1)    3er-Gruppen pro Zeile, mit Überschrift\n=fncSplit(\"C\";;0;0)     3 Zeilen pro Spalte, keine Überschrift\n=fncSplit(\"C\";4;0;0)    4 Zeilen pro Spalte, keine Überschrift\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv\nin folgenden Ordner entpackt:\n\n%TEMP%\\Excel-VSTO-Toolbox\\VBA_07\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "vba",
      "udf",
      "aufteilen",
      "spalte",
      "formeln"
    ],
    "codeText": "Option Explicit\n'Excel -VSTO - Toolbox\n'Power Query - Demo\n'Ralf Stolzenburg (Case)\n'https://github.com/rstsu/Excel-VSTO-Toolbox\n\n'=fncSplit(\"A\")         3er-Gruppen pro Zeile, keine Überschrift\n'=fncSplit(\"A\";4)       4er-Gruppen pro Zeile, keine Überschrift\n'=fncSplit(\"A\";;0)      3 Zeilen pro Spalte, Überschrift\n'=fncSplit(\"A\"; 3; 1)   3er-Gruppen pro Zeile, mit Überschrift (erste Zeile ignoriert)\n'=fncSplit(\"C\";;0;0)    3 Zeilen pro Spalte, keine Überschrift\n'=fncSplit(\"C\";4;0;0)   4 Zeilen pro Spalte, keine Überschrift\nPublic Function fncSplit(strColumn As String, Optional lngCount As Long = 3, Optional lngArt As Long = 1, Optional lngHead As Long = 1) As Variant\n    Dim varResult() As Variant\n    Dim lngColumn As Long\n    Dim rngRange As Range\n    Dim varArr As Variant\n    Dim lngRCount As Long\n    Dim lngCCount As Long\n    Dim lngStart As Long\n    Dim lngRow As Long\n    Dim lngRC As Long\n    Set rngRange = Range(strColumn & \"1\", Cells(Rows.Count, strColumn).End(xlUp))\n    lngStart = 1 + lngHead\n    varArr = rngRange.Resize(rngRange.Rows.Count - lngHead, 1).Offset(lngHead, 0).Value\n    lngRC = UBound(varArr, 1)\n    Do While lngRC Mod lngCount <> 0\n        lngRC = lngRC + 1\n    Loop\n    If lngArt = 1 Then\n        lngRCount = lngRC \\ lngCount\n        lngCCount = lngCount\n        ReDim varResult(1 To lngRCount, 1 To lngCCount)\n        For lngRow = 1 To lngRCount\n            For lngColumn = 1 To lngCCount\n                If (lngRow - 1) * lngCount + lngColumn <= UBound(varArr, 1) Then\n                    varResult(lngRow, lngColumn) = varArr((lngRow - 1) * lngCount + lngColumn, 1)\n                Else\n                    varResult(lngRow, lngColumn) = vbNullString\n                End If\n            Next lngColumn\n        Next lngRow\n    Else\n        lngRCount = lngCount\n        lngCCount = lngRC \\ lngCount\n        ReDim varResult(1 To lngRCount, 1 To lngCCount)\n        For lngColumn = 1 To lngCCount\n            For lngRow = 1 To lngRCount\n                If (lngColumn - 1) * lngCount + lngRow <= UBound(varArr, 1) Then\n                    varResult(lngRow, lngColumn) = varArr((lngColumn - 1) * lngCount + lngRow, 1)\n                Else\n                    varResult(lngRow, lngColumn) = vbNullString\n                End If\n            Next lngRow\n        Next lngColumn\n    End If\n    fncSplit = varResult\nEnd Function",
    "isNew": false,
    "package": {
      "fileName": "Demo_VBA_7.zip",
      "folderName": "VBA_07",
      "sizeBytes": 19034,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_VBA_7.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L510"
  },
  {
    "id": "VB-008",
    "sourceId": "vba_008",
    "title": "Access Datenbank erstellen - die Zweite...",
    "category": "VBA",
    "summary": "Es wird eine Accessdatenbank im lokalen Temp-Ordner mit Namen \"VBA_Demo_8.accdb\" aus Excel erstellt. Access_Datenbank_erstellen_VBA_Demo_8_accdb.xlsb",
    "description": "Es wird eine Accessdatenbank im lokalen Temp-Ordner mit Namen \"VBA_Demo_8.accdb\" aus Excel erstellt.\nAccess_Datenbank_erstellen_VBA_Demo_8_accdb.xlsb\n\nDann werden folgende Tabellen angelegt:\n1. Kunden\n2. Artikel\n3. Bestellungen\n4. Bestellpositionen\n\nIn den entsprechenden Tabellen werden Daten zu Kunden, Artikel, Bestellungen und  Bestellpositionen erstellt.\nNachdem noch Beziehungen eingerichtet sind, wird die Datei geöffnet.\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv\nin folgenden Ordner entpackt:\n\n%TEMP%\\Excel-VSTO-Toolbox\\VBA_08\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "vba",
      "access",
      "adox",
      "ace",
      "oledb"
    ],
    "codeText": "Option Explicit\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' Ralf Stolzenburg (Case)\n' https://github.com/rstsu/Excel-VSTO-Toolbox\nPrivate Declare PtrSafe Function ShellExecute Lib \"shell32.dll\" Alias \"ShellExecuteA\" ( _\n        ByVal hwnd As LongPtr, _\n        ByVal lpOperation As String, _\n        ByVal lpFile As String, _\n        ByVal lpParameters As String, _\n        ByVal lpDirectory As String, _\n        ByVal nShowCmd As Long) As LongPtr\nPrivate Const SW_MAXIMIZE = 3\nPublic Sub Main_Access_2()\n    Dim strDescription As String\n    Dim varProvider As Variant\n    Dim strFileName As String\n    Dim strProvider As String\n    Dim catCatalog As Object\n    Dim objTable As Object\n    Dim objIndex As Object\n    Dim objConn As Object\n    Dim lngError As Long\n    strFileName = Environ$(\"TEMP\") & \"\\VBA_Demo_8.accdb\"\n    On Error GoTo Fin\n    If Len(Dir$(strFileName)) > 0 Then Kill strFileName\n    Set catCatalog = CreateObject(\"ADOX.Catalog\")\n    For Each varProvider In Array(\"Microsoft.ACE.OLEDB.16.0\", \"Microsoft.ACE.OLEDB.12.0\")\n        Err.Clear\n        On Error Resume Next\n        catCatalog.Create \"Provider=\" & CStr(varProvider) & \";Data Source=\" & strFileName & \";\"\n        If Err.Number = 0 Then\n            strProvider = CStr(varProvider)\n            On Error GoTo Fin\n            Exit For\n        End If\n        lngError = Err.Number\n        strDescription = Err.Description\n        On Error GoTo Fin\n    Next varProvider\n    If Len(strProvider) = 0 Then\n        Err.Raise vbObjectError + 1000, _\n            \"CreateDataBase_1\", _\n            \"Es wurde kein geeigneter ACE-OLEDB-Provider gefunden.\" & _\n            vbCrLf & vbCrLf & _\n            \"Letzter Fehler: \" & lngError & _\n            vbCrLf & strDescription\n    End If\n    Set objConn = CreateObject(\"ADODB.Connection\")\n    With objConn\n        .CursorLocation = 3\n        .Open \"Provider=\" & strProvider & \";Data Source=\" & strFileName & \";\"\n        .Execute _\n            \"CREATE TABLE Kunden (\" & _\n            \"KundenID AUTOINCREMENT CONSTRAINT PK_Kunden PRIMARY KEY, \" & _\n            \"Firma VARCHAR(100), \" & _\n            \"Ansprechpartner VARCHAR(100), \" & _\n            \"Ort VARCHAR(50), \" & _\n            \"Land VARCHAR(50))\"\n        .Execute _\n            \"CREATE TABLE Artikel (\" & _\n            \"ArtikelID AUTOINCREMENT CONSTRAINT PK_Artikel PRIMARY KEY, \" & _\n            \"Bezeichnung VARCHAR(100), \" & _\n            \"Preis CURRENCY, \" & _\n            \"Bestand INTEGER)\"\n        .Execute _\n            \"CREATE TABLE Bestellungen (\" & _\n            \"BestellID AUTOINCREMENT CONSTRAINT PK_Bestellungen PRIMARY KEY, \" & _\n            \"KundenID LONG, \" & _\n            \"Bestelldatum DATETIME)\"\n        .Execute _\n            \"CREATE TABLE Bestellpositionen (\" & _\n            \"PositionsID AUTOINCREMENT CONSTRAINT PK_Bestellpositionen PRIMARY KEY, \" & _\n            \"BestellID LONG, \" & _\n            \"ArtikelID LONG, \" & _\n            \"Menge INTEGER, \" & _\n            \"Einzelpreis CURRENCY)\"\n        .Execute _\n            \"ALTER TABLE Bestellungen \" & _\n            \"ADD CONSTRAINT FK_Bestellungen_Kunden \" & _\n            \"FOREIGN KEY (KundenID) REFERENCES Kunden (KundenID)\"\n        .Execute _\n            \"ALTER TABLE Bestellpositionen \" & _\n            \"ADD CONSTRAINT FK_Positionen_Bestellungen \" & _\n            \"FOREIGN KEY (BestellID) REFERENCES Bestellungen (BestellID)\"\n        .Execute _\n            \"ALTER TABLE Bestellpositionen \" & _\n            \"ADD CONSTRAINT FK_Positionen_Artikel \" & _\n            \"FOREIGN KEY (ArtikelID) REFERENCES Artikel (ArtikelID)\"\n        .Execute \"INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Nordstern GmbH', 'Anna Berger', 'Hamburg', 'Deutschland')\"\n        .Execute \"INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Rheinblick AG', 'Michael Weber', 'Köln', 'Deutschland')\"\n        .Execute \"INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Alpenhandel KG', 'Sabine Keller', 'München', 'Deutschland')\"\n        .Execute \"INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Spree Büroservice', 'Thomas Richter', 'Berlin', 'Deutschland')\"\n        .Execute \"INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('MainTech GmbH', 'Julia Hartmann', 'Frankfurt', 'Deutschland')\"\n        .Execute \"INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Elbe Logistik', 'Martin Schulze', 'Dresden', 'Deutschland')\"\n        .Execute \"INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Westfalen Bürobedarf', 'Petra König', 'Dortmund', 'Deutschland')\"\n        .Execute \"INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Donau Consulting', 'Daniel Fischer', 'Ulm', 'Deutschland')\"\n        .Execute \"INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Notebookständer', 29.9, 35)\"\n        .Execute \"INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('USB-C Hub', 39.5, 28)\"\n        .Execute \"INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Funkmaus', 24.9, 42)\"\n        .Execute \"INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Tastatur', 49.9, 21)\"\n        .Execute \"INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Webcam', 59.0, 17)\"\n        .Execute \"INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Headset', 69.9, 14)\"\n        .Execute \"INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Monitorarm', 84.5, 9)\"\n        .Execute \"INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Laptop-Tasche', 44.0, 31)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (1, #01/12/2026#)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (2, #01/18/2026#)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (1, #02/03/2026#)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (4, #02/15/2026#)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (5, #03/02/2026#)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (3, #03/21/2026#)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (6, #04/09/2026#)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (2, #04/28/2026#)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (7, #05/10/2026#)\"\n        .Execute \"INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (8, #05/23/2026#)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (1, 1, 2, 29.9)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (1, 3, 2, 24.9)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (2, 2, 3, 39.5)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (2, 4, 1, 49.9)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (3, 5, 2, 59)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (3, 6, 1, 69.9)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (4, 3, 4, 24.9)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (4, 8, 2, 44)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (5, 7, 2, 84.5)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (5, 2, 2, 39.5)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (6, 4, 3, 49.9)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (7, 6, 2, 69.9)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (8, 1, 5, 29.9)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (8, 5, 2, 59)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (9, 8, 3, 44)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (10, 2, 1, 39.5)\"\n        .Execute \"INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (10, 3, 2, 24.9)\"\n        MsgBox \"Datenbank wurde erstellt:\" & vbCrLf & _\n            strFileName & vbCrLf & vbCrLf & _\n            \"Provider: \" & strProvider, vbInformation\n        If .State = 1 Then .Close\n    End With\n    If ShellExecute(Application.hwnd, \"Open\", strFileName, vbNullString, vbNullString, SW_MAXIMIZE) <= 32 Then\n        MsgBox \"Die Datenbank wurde erstellt, konnte aber nicht geöffnet werden.\", vbExclamation\n    End If\nFin:\n    If Err.Number <> 0 Then MsgBox \"Fehler: \" & Err.Number & vbCrLf & Err.Description, vbExclamation\n    If Not objConn Is Nothing Then\n        If objConn.State = 1 Then objConn.Close\n    End If\n    Set objIndex = Nothing\n    Set objTable = Nothing\n    Set catCatalog = Nothing\n    Set objConn = Nothing\nEnd Sub",
    "isNew": false,
    "package": {
      "fileName": "Demo_VBA_8.zip",
      "folderName": "VBA_08",
      "sizeBytes": 20047,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_VBA_8.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L606"
  },
  {
    "id": "VB-009",
    "sourceId": "vba_009",
    "title": "Klassenprogrammierung - UserForm - TextBox - Text markieren",
    "category": "VBA",
    "summary": "UserForm - TextBox - Klassenprogrammierung. Klassenprogrammierung_UserForm_TextBox_Text_markieren_Farbe_wechseln_und_zurueck.xlsb",
    "description": "UserForm - TextBox - Klassenprogrammierung.\nKlassenprogrammierung_UserForm_TextBox_Text_markieren_Farbe_wechseln_und_zurueck.xlsb\n\nIn einer UserForm werden alle TextBoxen in einem Array gesammelt und der Klasse zugewiesen.\nDamit wird bei jeder TextBox beim aktivieren der gesamte Text markiert.\nBei einem Doppleklick in die TextBox wird Schrift- und Hintergrundfarbe geändert. Nächster Doppelklick alles wieder zurück.\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv\nin folgenden Ordner entpackt:\n\n%TEMP%\\Excel-VSTO-Toolbox\\VBA_09\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "vba",
      "userform",
      "klassenprogrammierung",
      "textbox",
      "markieren",
      "farbe",
      "doppelklick"
    ],
    "codeText": "DieseArbeitsmappe:\nOption Explicit\nPrivate Sub Workbook_Open()\n    UserForm1.Show\nEnd Sub\n\nUserForm1:\nOption Explicit\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' Ralf Stolzenburg (Case)\n' https://github.com/rstsu/Excel-VSTO-Toolbox\n' Klasse initialisieren. Nur TextBoxen über Array sammeln.\n' Text in TextBox1 markieren. Bei Doppelklick Farbe wechseln und zurück.\nPrivate mobjTextBoxClass() As clsTextBox\nPrivate Sub UserForm_Initialize()\n    Dim objControl As Control\n    Dim lngIndex As Long\n    For Each objControl In Controls\n        If TypeOf objControl Is MSForms.TextBox Then\n            ReDim Preserve mobjTextBoxClass(lngIndex)\n            Set mobjTextBoxClass(lngIndex) = New clsTextBox\n            Set mobjTextBoxClass(lngIndex).prpTextBox = objControl\n            lngIndex = lngIndex + 1\n        End If\n    Next objControl\n    With TextBox1\n        .SelStart = 0\n        .SelLength = .TextLength\n    End With\nEnd Sub\nPrivate Sub UserForm_Terminate()\n    Dim lngIndex As Long\n    If CBool(Not Not mobjTextBoxClass) Then\n        For lngIndex = LBound(mobjTextBoxClass) To UBound(mobjTextBoxClass)\n            Set mobjTextBoxClass(lngIndex) = Nothing\n        Next lngIndex\n    End If\nEnd Sub\nPrivate Sub CommandButton1_Click()\n    Unload Me\nEnd Sub\n\nModul1:\nOption Explicit\nPublic Sub Main()\n    UserForm1.Show\nEnd Sub\n\nKlassenmodul (Name = clsTextBox):\nOption Explicit\n' Excel-VSTO-Toolbox\n' VBA-Demo\n' Ralf Stolzenburg (Case)\n' https://github.com/rstsu/Excel-VSTO-Toolbox\n' Klasse initialisieren. Nur TextBoxen über Array sammeln.\n' Text in TextBox1 markieren. Bei Doppelklick Farbe wechseln und zurück.\n' https://learn.microsoft.com/de-de/dotnet/visual-basic/language-reference/modifiers/withevents\nPrivate WithEvents mobjTextBox As MSForms.TextBox\n' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/property-get-statement\nPrivate Property Get TextBox() As MSForms.TextBox\n    Set TextBox = mobjTextBox\nEnd Property\n' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/property-set-statement\nFriend Property Set prpTextBox(objTextBox As MSForms.TextBox)\n    Set mobjTextBox = objTextBox\nEnd Property\n' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/terminate-event-visual-basic-for-applications\nPrivate Sub Class_Terminate()\n    Set mobjTextBox = Nothing\nEnd Sub\n' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/mousedown-mouseup-events\nPrivate Sub mobjTextBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)\n    With mobjTextBox\n        .SelStart = 0\n        .SelLength = .TextLength\n    End With\nEnd Sub\n' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/dblclick-event\nPrivate Sub mobjTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)\n    With TextBox\n        .BackColor = IIf(.BackColor = &H80000005, &HC0FFC0, &H80000005)\n        .ForeColor = IIf(.ForeColor = -2147483640, &HFF&, -2147483640)\n    End With\nEnd Sub",
    "isNew": false,
    "package": {
      "fileName": "Demo_VBA_9.zip",
      "folderName": "VBA_09",
      "sizeBytes": 23096,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_VBA_9.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L795"
  },
  {
    "id": "VB-010",
    "sourceId": "vba_0010",
    "title": "Datum umwandeln",
    "category": "VBA",
    "summary": "March 6 2020 8:00:12 AM February 28 2020 11:17:26 AM January 24 2020 3:22:44 PM December 30 2023 1:38 PM November 22 2024 7:29 AM 8/31/2025 8:15:57 AM 4/24/2022 9:31:16 AM 9/24/2026 …",
    "description": "March 6 2020 8:00:12 AM\nFebruary 28 2020 11:17:26 AM\nJanuary 24 2020 3:22:44 PM\nDecember 30 2023 1:38 PM\nNovember 22 2024 7:29 AM\n8/31/2025 8:15:57 AM\n4/24/2022 9:31:16 AM\n9/24/2026 9:31:16 PM\n2026-11-22T14:30:00Z\n2026-08-12T14:30:00Z\n\nDiese Daten werden umgewandelt. Auch UTC nach MEZ/MESZ.\nVBA_Power_Query_Formel_Datum_umwandeln.xlsb\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Demo_VBA_PQ_Formel\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "formel",
      "datum",
      "utc",
      "mez",
      "mesz",
      "vba",
      "power query"
    ],
    "codeText": "Option Explicit\n' Excel -VSTO - Toolbox\n' Power Query - Demo\n' Ralf Stolzenburg (Case)\n' https://github.com/rstsu/Excel-VSTO-Toolbox\nPrivate Declare PtrSafe Function VarDateFromStr Lib \"oleaut32.dll\" _\n    (ByVal strIn As LongPtr, _\n    ByVal lcid As Long, _\n    ByVal dwFlags As Long, _\n    ByRef pDateOut As Date) As Long\n' https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-vardatefromstr\nFunction fncDate(ByVal strTMP As String, Optional ByVal lcid As Long = 1033) As Date\n    Call VarDateFromStr(StrPtr(strTMP), lcid, &H80000000, fncDate)\nEnd Function\nPublic Sub Main()\n    Dim lngLastRow As Long\n    Dim dateTMP As Date\n    On Error GoTo Fin\n    With Tabelle1\n        For lngLastRow = 2 To IIf(Len(.Cells(.Rows.Count, 1)), .Rows.Count, .Cells(.Rows.Count, 1).End(xlUp).Row)\n            If .Cells(lngLastRow, 1).Value <> \"\" Then\n                dateTMP = fncDate(.Cells(lngLastRow, 1).Text)\n                If dateTMP = \"00:00:00\" Then dateTMP = fncUTCtoDE(.Cells(lngLastRow, 1).Text)\n                .Cells(lngLastRow, 4).NumberFormat = \"dd/mm/yyyy hh:mm:ss\"\n                .Cells(lngLastRow, 4).Value = dateTMP\n                .Cells(lngLastRow, 6).NumberFormat = \"dd/mm/yyyy\"\n                .Cells(lngLastRow, 6).Value = DateValue(dateTMP)\n                .Cells(lngLastRow, 7).NumberFormat = \"hh:mm:ss\"\n                .Cells(lngLastRow, 7).Value = TimeValue(dateTMP)\n            End If\n        Next lngLastRow\n    End With\nFin:\n    If Err.Number <> 0 Then MsgBox \"Fehler: \" & Err.Number & \" \" & Err.Description\nEnd Sub\nPublic Function fncUTCtoDE(utcString As String) As Date\n    Dim lngOffset As Long\n    Dim dateStart As Date\n    Dim dateEnd As Date\n    Dim dateUTC As Date\n    dateUTC = DateSerial(CInt(Mid$(utcString, 1, 4)), CInt(Mid$(utcString, 6, 2)), CInt(Mid$(utcString, 9, 2))) + _\n        TimeSerial(CInt(Mid$(utcString, 12, 2)), CInt(Mid$(utcString, 15, 2)), CInt(Mid$(utcString, 18, 2)))\n    dateStart = DateSerial(CInt(Mid$(utcString, 1, 4)), 3, 31 - Weekday(DateSerial(CInt(Mid$(utcString, 1, 4)), 3, 31), vbMonday) + 1)\n    dateEnd = DateSerial(CInt(Mid$(utcString, 1, 4)), 10, 31 - Weekday(DateSerial(CInt(Mid$(utcString, 1, 4)), 10, 31), vbMonday) + 1)\n    lngOffset = 1\n    If dateUTC >= dateStart + TimeSerial(1, 0, 0) And dateUTC < dateEnd + TimeSerial(1, 0, 0) Then\n        lngOffset = lngOffset + 1\n    End If\n    fncUTCtoDE = dateUTC + lngOffset / 24\nEnd Function",
    "isNew": false,
    "package": {
      "fileName": "Demo_VBA_PQ_Formel.zip",
      "folderName": "Demo_VBA_PQ_Formel",
      "sizeBytes": 37133,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_VBA_PQ_Formel.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L914"
  },
  {
    "id": "VB-011",
    "sourceId": "vba_0011",
    "title": "Langer Text in Zellen aufteilen...",
    "category": "VBA",
    "summary": "Lange Texte in Zellen werden nach vorgegebener Anzahl von Zeichen aufgetrennt. Nicht mitten im Wort. Verschiedene Möglichkeiten (in Spalten, Zeilen, Blöcken...). Mit RegEx, Power Query, …",
    "description": "Lange Texte in Zellen werden nach vorgegebener Anzahl von Zeichen aufgetrennt.\nNicht mitten im Wort.\nVerschiedene Möglichkeiten (in Spalten, Zeilen, Blöcken...).\nMit RegEx, Power Query, Formeln und VBA (UDF und Sub).\n\nRegEx_Power_Query_Formel_VBA_Text_auftrennen.xlsb\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Demo_RegEx_PQ_Formel_VBA\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "auftrennenl",
      "text",
      "aufteilen",
      "formel",
      "regex",
      "vba",
      "power query"
    ],
    "codeText": "Option Explicit\n'Excel -VSTO - Toolbox\n'Regex, Power Query, Formel und VBA - Demo\n'Ralf Stolzenburg (Case)\n'https://github.com/rstsu/Excel-VSTO-Toolbox\nPublic Sub Main()\n    Dim varArr As Variant\n    Dim lngLoop As Long\n    Dim lngTMP As Long\n    Dim lngRow As Long\n    Range(\"D2:I14\").ClearContents\n    lngRow = 2\n    For lngLoop = 2 To Cells(Rows.Count, \"A\").End(xlUp).Row\n        If Len(Cells(lngLoop, \"A\").Value) > 0 Then\n            varArr = SplitTextByLength(CStr(Cells(lngLoop, \"A\").Value), Range(\"C1\").Value)\n            For lngTMP = 0 To UBound(varArr)\n                Cells(lngRow, 4 + lngTMP).Value = varArr(lngTMP)\n            Next lngTMP\n            lngRow = lngRow + 1\n        End If\n    Next lngLoop\nEnd Sub\nPrivate Function SplitTextByLength(ByVal strText As String, ByVal lngMaxLen As Long) As Variant\n    Dim strTestText As String\n    Dim strCurrTExt As String\n    Dim strResult() As String\n    Dim strWords() As String\n    Dim lngLoop As Long\n    Dim lngTMP As Long\n    strText = Application.Trim(strText)\n    If Len(strText) = 0 Or lngMaxLen <= 0 Then\n        SplitTextByLength = Array(\"\")\n        Exit Function\n    End If\n    strWords = Split(strText, \" \")\n    ReDim strResult(0 To UBound(strWords))\n    strCurrTExt = strWords(0)\n    For lngLoop = 1 To UBound(strWords)\n        strTestText = strCurrTExt & \" \" & strWords(lngLoop)\n        If Len(strTestText) <= lngMaxLen Then\n            strCurrTExt = strTestText\n        Else\n            strResult(lngTMP) = strCurrTExt\n            lngTMP = lngTMP + 1\n            strCurrTExt = strWords(lngLoop)\n        End If\n    Next lngLoop\n    strResult(lngTMP) = strCurrTExt\n    ReDim Preserve strResult(0 To lngTMP)\n    SplitTextByLength = strResult\nEnd Function\nPublic Function TextTrennen(ByVal strText As String, ByVal lngMaxLen As Long) As Variant\n    Dim varResult() As Variant\n    Dim varArr As Variant\n    Dim lngLoop As Long\n    varArr = SplitTextByLength(strText, lngMaxLen)\n    ReDim varResult(1 To UBound(varArr) + 1, 1 To 1)\n    For lngLoop = 0 To UBound(varArr)\n        varResult(lngLoop + 1, 1) = varArr(lngLoop)\n    Next lngLoop\n    TextTrennen = varResult\nEnd Function\nPublic Function TextTrennen_A(ByVal strText As String, ByVal lngMaxLen As Long) As Variant\n    Dim varResult() As Variant\n    Dim varArr As Variant\n    Dim lngLoop As Long\n    varArr = SplitTextByLength(strText, lngMaxLen)\n    ReDim varResult(1 To 1, 1 To UBound(varArr) + 1)\n    For lngLoop = 0 To UBound(varArr)\n        varResult(1, lngLoop + 1) = varArr(lngLoop)\n    Next lngLoop\n    TextTrennen_A = varResult\nEnd Function",
    "isNew": true,
    "package": {
      "fileName": "Demo_RegEx_PQ_Formel_VBA.zip",
      "folderName": "Demo_RegEx_PQ_Formel_VBA",
      "sizeBytes": 73867,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_RegEx_PQ_Formel_VBA.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Vba.vb#L1006"
  },
  {
    "id": "FX-001",
    "sourceId": "formula_001",
    "title": "Eindeutige Werte sortieren",
    "category": "Formeln",
    "summary": "Beispiel mit EINDEUTIG und SORTIEREN. Es werden nur eindeutige Werte aus Spalte A sortiert wiedergegeben. Die zweite Formel (D1) ist mit Überschrift.",
    "description": "Beispiel mit EINDEUTIG und SORTIEREN.\nEs werden nur eindeutige Werte aus Spalte A sortiert wiedergegeben.\nDie zweite Formel (D1) ist mit Überschrift.",
    "tags": [
      "formel",
      "eindeutig",
      "sortieren",
      "text",
      "formula2"
    ],
    "codeText": "=SORTIEREN(EINDEUTIG(A2:.A999))\n\n=VSTAPELN(A1;SORTIEREN(EINDEUTIG(A2:.A999)))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L7"
  },
  {
    "id": "FX-002",
    "sourceId": "formula_002",
    "title": "Zahlen \"richtig\" sortieren",
    "category": "Formeln",
    "summary": "Zahlen die als Text vorliegen (3.10.20) richtig sortieren. Es soll also nicht 12.52.63 vor 2.11.74 einsortiert werden sondern in der richtigen Reihenfolge. Die zweite Formel (G1) ist mit …",
    "description": "Zahlen die als Text vorliegen (3.10.20) richtig sortieren.\nEs soll also nicht 12.52.63 vor 2.11.74 einsortiert werden sondern in der richtigen Reihenfolge.\nDie zweite Formel (G1) ist mit Überschriften.",
    "tags": [
      "formel",
      "zahl",
      "sortieren",
      "text",
      "formula2"
    ],
    "codeText": "=SORTIERENNACH(A2:.B999;--TEXTVOR(A2:.A999;\".\");1;--TEIL(TEXTNACH(A2:.A999;\".\");1;2);1)\n\n=VSTAPELN(A1:B1;SORTIERENNACH(A2:.B999;--TEXTVOR(A2:.A999;\".\");1;--TEIL(TEXTNACH(A2:.A999;\".\");1;2);1))\n\n=LET(w;A2:.A999;x;A2:.B999;y;--TEXTVOR(w;\".\");z;--TEIL(TEXTNACH(w;\".\");1;2);SORTIERENNACH(x;y;1;z;1))\n\n=LET(w;A2:.A999;x;A2:.B999;y;--TEXTVOR(w;\".\");z;--TEIL(TEXTNACH(w;\".\");1;2);VSTAPELN(A1:B1;SORTIERENNACH(x;y;1;z;1)))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L31"
  },
  {
    "id": "FX-003",
    "sourceId": "formula_003",
    "title": "Kalender erstellen",
    "category": "Formeln",
    "summary": "Es wird ein Blockkalender erstellt (6 Zeilen und 7 Spalten). Einmal nur der Kalender, dann mit Tagen als Überschrift und schließlich noch mit Monat als Überschrift. Jahr und Monat kann …",
    "description": "Es wird ein Blockkalender erstellt (6 Zeilen und 7 Spalten).\nEinmal nur der Kalender, dann mit Tagen als Überschrift und schließlich noch mit Monat als Überschrift.\nJahr und Monat kann jeweils aus \"Daten - Gültigkeit - Liste\" ausgewählt werden.\nMarkiert werden Feiertage, Besondere Tage, Wochenenden und der aktuelle Tag.\nBundeseinheitliche Feiertage in I2:I10.\nDas färben (Schrift und Hintergrund) geht auch über die Bedingte Formatiereung:\nSchrift grau und kursiv, wenn nicht im aktuellen Monat - Formel: =MONAT(I23)<>$F$2 und dann entsprechend formatieren.\nFeiertage/Besondere Tage - Formel: =ISTZAHL(VERGLEICH(I23;$Q$2:$Q$10;0)) und dann entsprechend formatieren.\nHeutiger Tag - Formel: =I23=HEUTE() und dann entsprechend formatieren.\nWochenenden - Formel: =WOCHENTAG(I23;2)>5 und dann entsprechend formatieren.\nAdressen anpassen und anwenden auf den jeweiligen Kalenderbereich.",
    "tags": [
      "formel",
      "zahl",
      "kalender",
      "bedingte formatierung",
      "datum"
    ],
    "codeText": "=LET(Jahr;$B$1;Monat;$B$2;Erster;DATUM(Jahr;Monat;1);Start;Erster-WOCHENTAG(Erster;2)+1;SEQUENZ(6;7;Start;1))\n\n=LET(Jahr;$B$1;Monat;$B$2;Erster;DATUM(Jahr;Monat;1);Start;Erster-WOCHENTAG(Erster;2)+1;VSTAPELN(TEXT(SEQUENZ(1;7;2;1);\"TTT\");SEQUENZ(6;7;Start;1)))\n\n=LET(Jahr;$B$1;Monat;$B$2;Erster;DATUM(Jahr;Monat;1);Start;Erster-WOCHENTAG(Erster;2)+1;VSTAPELN(HSTAPELN(TEXT(Erster;\"MMMM JJJJ\");\"\";\"\";\"\";\"\";\"\";\"\");TEXT(SEQUENZ(1;7;2;1);\"TTT\");SEQUENZ(6;7;Start;1)))\n\n=LET(Jahr;$B$1;Ostersonntag;RUNDEN((TAG(MINUTE(Jahr/38)/2+55)&\".4.\"&Jahr)/7;0)*7-6;Daten;VSTAPELN(DATUM(Jahr;{1;5;10;12;12};{1;1;3;25;26});Ostersonntag+{-2;1;39;50});Namen;{\"Neujahr\";\"Tag der Arbeit\";\"Tag der Deutschen Einheit\";\"1. Weihnachtstag\";\"2. Weihnachtstag\";\"Karfreitag\";\"Ostermontag\";\"Christi Himmelfahrt\";\"Pfingstmontag\"};SORTIERENNACH(HSTAPELN(Daten;Namen);Daten))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L59"
  },
  {
    "id": "FX-004",
    "sourceId": "formula_004",
    "title": "Kreuztabelle aus Liste erste 3 Buchstaben",
    "category": "Formeln",
    "summary": "Aus einer Liste (A2:A15) wird eine Kreuztabelle erstellt. Grundlage sind die ersten 3 gleichen Buchstaben. Einmal bezogen auf einfache Werte (Spalte C). Und dann auf eine \"intelligente …",
    "description": "Aus einer Liste (A2:A15) wird eine Kreuztabelle erstellt.\nGrundlage sind die ersten 3 gleichen Buchstaben.\nEinmal bezogen auf einfache Werte (Spalte C).\nUnd dann auf eine \"intelligente Tabelle\" (Spalte A).\nFormeln in E1, M1 und U1.",
    "tags": [
      "formel",
      "text",
      "kreuztabelle",
      "intelligente tabelle",
      "liste"
    ],
    "codeText": "=LET(t;Demo_Formel_4[Wert];u;LINKS(t;3);v;EINDEUTIG(u);w;MAX(MAP(v;LAMBDA(x;SUMME(--(u=x)))));VSTAPELN(\"Pos\"&SEQUENZ(;w);MATRIXERSTELLEN(ZEILEN(v);w;LAMBDA(y;z;WENNFEHLER(INDEX(FILTER(t;u=INDEX(v;y));z);\"\")))))\n\n=LET(m;Demo_Formel_4[Wert];k;EINDEUTIG(LINKS(m;3));n;MAX(MAP(k;LAMBDA(x;SUMME(--(LINKS(m;3)=x)))));VSTAPELN(\"Pos\"&SEQUENZ(;n);MATRIXERSTELLEN(ZEILEN(k);n;LAMBDA(z;p;WENNFEHLER(INDEX(FILTER(m;LINKS(m;3)=INDEX(k;z));p);\"\")))))\n\n=LET(t;$A$2:.$A$999;u;LINKS(t;3);v;EINDEUTIG(u);w;MAX(MAP(v;LAMBDA(x;SUMME(--(u=x)))));VSTAPELN(\"Pos\"&SEQUENZ(;w);MATRIXERSTELLEN(ZEILEN(v);w;LAMBDA(y;z;WENNFEHLER(INDEX(FILTER(t;u=INDEX(v;y));z);\"\")))))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L95"
  },
  {
    "id": "FX-005",
    "sourceId": "formula_005",
    "title": "Datumbrereich in Liste auflösen",
    "category": "Formeln",
    "summary": "Aus einer Liste (A2:A6 - Ferientermine 2027 in Baden Württemberg) wird eine laufende Liste erstellt.",
    "description": "Aus einer Liste (A2:A6 - Ferientermine 2027 in Baden Württemberg) wird eine laufende Liste erstellt.\n\nFormeln in E2 und H1 (diese ist mit Überschrift).",
    "tags": [
      "formel",
      "ferien",
      "Datum",
      "intelligente tabelle",
      "liste"
    ],
    "codeText": "=LET(v;FBW[Von];b;FBW[Bis];n;b-v+1;m;MAX(n);d;ZUSPALTE(WENN(SEQUENZ(m;;0)<=MTRANS(n-1);MTRANS(v)+SEQUENZ(m;;0);NV());3;WAHR);HSTAPELN(d;TEXT(d;\"TTTT\")))\n\n=VSTAPELN({\"Datum\".\"Tag\"};LET(v;FBW[Von];b;FBW[Bis];n;b-v+1;m;MAX(n);d;ZUSPALTE(WENN(SEQUENZ(m;;0)<=MTRANS(n-1);MTRANS(v)+SEQUENZ(m;;0);NV());3;WAHR);HSTAPELN(d;TEXT(d;\"TTTT\"))))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L123"
  },
  {
    "id": "FX-006",
    "sourceId": "formula_006",
    "title": "Primzahlen auflisten - Von Bis",
    "category": "Formeln",
    "summary": "Aus einer Liste (B2:C5) werden die Primzahlen aufgelistet.",
    "description": "Aus einer Liste (B2:C5) werden die Primzahlen aufgelistet.\n\nAufgeteilt, da mit einer Spill-Formel - je nach Größe der vorgegebenen Zahlen - die Berechnung schnell zu umfangreich wird.\n\nFormeln in E2, F2, G2, H2, J1 und K1.",
    "tags": [
      "formel",
      "primzahlen",
      "start",
      "ende",
      "zahlen"
    ],
    "codeText": "=LET(Z;SEQUENZ(C2-B2+1;;B2);t;SEQUENZ(1;MAX(Z));FILTER(Z;(Z>1)*(MMULT(--(REST(Z;t)=0);MTRANS(t^0))=2)))\n=LET(Z;SEQUENZ(C3-B3+1;;B3);t;SEQUENZ(1;MAX(Z));FILTER(Z;(Z>1)*(MMULT(--(REST(Z;t)=0);MTRANS(t^0))=2)))\n=LET(Z;SEQUENZ(C4-B4+1;;B4);t;SEQUENZ(1;MAX(Z));FILTER(Z;(Z>1)*(MMULT(--(REST(Z;t)=0);MTRANS(t^0))=2)))\n=LET(Z;SEQUENZ(C5-B5+1;;B5);t;SEQUENZ(1;MAX(Z));FILTER(Z;(Z>1)*(MMULT(--(REST(Z;t)=0);MTRANS(t^0))=2)))\n\n=VSTAPELN(\"Gestapelt 1\";VSTAPELN(E2#;F2#;G2#))\n\n=VSTAPELN(\"Gestapelt 2\";ZUSPALTE(E2:.G999;1;WAHR))",
    "isNew": false,
    "package": null,
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L147"
  },
  {
    "id": "FX-007",
    "sourceId": "formula_007",
    "title": "Filtern bestimmter Tage aus Liste mit Datum",
    "category": "Formeln",
    "summary": "Aus einer Liste (A2:C368 - Geburtstagsliste) werden vorgegebene Tage aufgelistet. Formel_aus_einer_Geburtstagsliste_oder_anderen_Listen_mit_Datum_bestimmte_Tage_ausgeben_FILTER.xlsx",
    "description": "Aus einer Liste (A2:C368 - Geburtstagsliste) werden vorgegebene Tage aufgelistet.\nFormel_aus_einer_Geburtstagsliste_oder_anderen_Listen_mit_Datum_bestimmte_Tage_ausgeben_FILTER.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Formel_07\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in E2:H2, J2:M2, O2:R2, T1:W1.\nIn X1 ist eine Liste der Tage (Daten Gültigkeit).\nDie Zellen in E2, G2, J2, L2, O2, Q2 sind \"Benutzerdefiniertes Format\". Je nach Aufbau der Formel gibt man 0 bis 6 oder 1 bis 7 ein.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "formel",
      "filter",
      "datum",
      "liste",
      "zahlen"
    ],
    "codeText": "=FILTER(C2:.C999;REST(C2:.C999;7)=E1;\"----\")\n=LET(x;FILTER(C2:.C999;REST(C2:.C999;7)=E1;\"----\");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))\n=FILTER(Tabelle1[Geburtsdatum];REST(Tabelle1[Geburtsdatum];7)=G1;\"----\")\n=LET(x;FILTER(Tabelle1[Geburtsdatum];REST(Tabelle1[Geburtsdatum];7)=G1;\"----\");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))\n\n=FILTER(C2:.C999;TEXT(C2:.C999;\"tttt\")=TEXT(J1;\"tttt\");\"---\")\n=LET(x;FILTER(C2:.C999;TEXT(C2:.C999;\"tttt\")=TEXT(J1;\"tttt\");\"---\");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))\n=FILTER(Tabelle1[Geburtsdatum];TEXT(Tabelle1[Geburtsdatum];\"tttt\")=TEXT(L1;\"tttt\");\"---\")\n=LET(x;FILTER(Tabelle1[Geburtsdatum];TEXT(Tabelle1[Geburtsdatum];\"tttt\")=TEXT(L1;\"tttt\");\"---\");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))\n\n=FILTER(C2:.C999;WOCHENTAG(C2:.C999;1)=O1;\"----\")\n=LET(x;FILTER(C2:.C999;WOCHENTAG(C2:.C999;1)=O1;\"----\");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))\n=FILTER(Tabelle1[Geburtsdatum];WOCHENTAG(Tabelle1[Geburtsdatum];1)=Q1;\"----\")\n=LET(x;FILTER(Tabelle1[Geburtsdatum];WOCHENTAG(Tabelle1[Geburtsdatum];1)=Q1;\"----\");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))\n\n=VSTAPELN(X1;FILTER(C2:.C999;TEXT(C2:.C999;\"tttt\")=TEXT(X1;\"tttt\");\"---\"))\n=VSTAPELN(X1;LET(x;FILTER(C2:.C999;TEXT(C2:.C999;\"tttt\")=TEXT(X1;\"tttt\");\"---\");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1)))\n=VSTAPELN(X1;FILTER(Tabelle1[Geburtsdatum];TEXT(Tabelle1[Geburtsdatum];\"tttt\")=TEXT(X1;\"tttt\");\"---\"))\n=VSTAPELN(X1;LET(x;FILTER(Tabelle1[Geburtsdatum];TEXT(Tabelle1[Geburtsdatum];\"tttt\")=TEXT(X1;\"tttt\");\"---\");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1)))",
    "isNew": false,
    "package": {
      "fileName": "Demo_Formel_7.zip",
      "folderName": "Formel_07",
      "sizeBytes": 33315,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Formel_7.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L178"
  },
  {
    "id": "FX-008",
    "sourceId": "formula_008",
    "title": "Aus Jahr und KW - Woche Von Bis in einer Zelle(n) darstellen",
    "category": "Formeln",
    "summary": "Aus Jahr und KW wird eine Woche in einer Zelle erstellt - z. B.: KW 34 - 17.08.2026 - 23.08.2026 Formel_KW_Jahr_Datum_Woche_Von_Bis_in_einer_Zelle_ausgeben_mit_KW_am_Anfang.xlsx",
    "description": "Aus Jahr und KW wird eine Woche in einer Zelle erstellt - z. B.:\nKW 34 - 17.08.2026 - 23.08.2026\nFormel_KW_Jahr_Datum_Woche_Von_Bis_in_einer_Zelle_ausgeben_mit_KW_am_Anfang.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Formel_08\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in I1, K1, M1.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "formel",
      "kw",
      "datum",
      "woche",
      "zahlen"
    ],
    "codeText": "=VSTAPELN(\"Jahr A2 - KW B2\";TEXT(7*KÜRZEN((2&-1&-$A$2)/7+C2)-5;\"TT.MM.JJJJ - \")&TEXT(7*KÜRZEN((2&-1&-$A$2)/7+C2)+1;\"TT.MM.JJJJ\"))\n=LET(k;SEQUENZ($C$2);z;7*KÜRZEN((2&-1&-$A$2)/7+k);e;WENN($E$2;z-1;WENN($F$2;z;z+1));VSTAPELN(\"Jahr \"&$A$2&\" - KW bis \"&$C$2;\"KW \"&k&\" - \"&TEXT(z-5;\"TT.MM.JJJJ - \")&TEXT(e;\"TT.MM.JJJJ\")))\n=VSTAPELN(\"Woche - Heute\";LET(x;ISOKALENDERWOCHE(HEUTE());y;JAHR(HEUTE());z;7*KÜRZEN((2&-1&-y)/7+x);\"KW \"&x&\" - \"&TEXT(z-5;\"TT.MM.JJJJ - \")&TEXT(z+1;\"TT.MM.JJJJ\")))",
    "isNew": false,
    "package": {
      "fileName": "Demo_Formel_8.zip",
      "folderName": "Formel_08",
      "sizeBytes": 11895,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Formel_8.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L233"
  },
  {
    "id": "FX-009",
    "sourceId": "formula_009",
    "title": "Zwei Spalten in eine Spalte zusammenfassen",
    "category": "Formeln",
    "summary": "Spalte A (Kundennummer) und Spalte B (Produktnummer) werden in eine Spalte zusammengefasst. Dabei werden zu jeder Kundennummer (in Spalte D nur die EINDEUTIGEN) alle Produkte ausgegeben. …",
    "description": "Spalte A (Kundennummer) und Spalte B (Produktnummer) werden in eine Spalte zusammengefasst. Dabei werden zu jeder Kundennummer (in Spalte D nur die EINDEUTIGEN) alle Produkte ausgegeben.\nÜber die Bedingte Formatierung werden die Kundennummern in Spalte D und E farbig dargestellt (Formel =REST(ZEILE(A2)-2;20)=0).\nAus_zwei_Spalten_mit_Ueberschrift_EINE_Spalte_Namen_und_Produkte_untereinander.xlsx\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Formel_09\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in D1 und E1.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "formel",
      "spalte",
      "spalten",
      "kundennummer",
      "produktnummer"
    ],
    "codeText": "=LET(q;FILTER(A2:.A999;A2:.A999<>\"\");r;FILTER(B2:.B999;B2:.B999<>\"\");s;ZEILEN(q);t;ZEILEN(r);u;1+t;v;s*u;w;SEQUENZ(v);x;QUOTIENT(w-1;u)+1;y;REST(w-1;u);z;WENN(y=0;TEXT(INDEX(q;x);\"@\");INDEX(r;y));VSTAPELN({\"Ausgabe\"};z))\n=LET(q;EINDEUTIG(FILTER(A2:.A999;A2:.A999<>\"\"));r;FILTER(B2:.B999;B2:.B999<>\"\");s;ZEILEN(q);t;ZEILEN(r);u;1+t;v;s*u;w;SEQUENZ(v);x;QUOTIENT(w-1;u)+1;y;REST(w-1;u);z;WENN(y=0;TEXT(INDEX(q;x);\"@\");INDEX(r;y));VSTAPELN({\"Ausgabe EINDEUTIG\"};z))",
    "isNew": false,
    "package": {
      "fileName": "Demo_Formel_9.zip",
      "folderName": "Formel_09",
      "sizeBytes": 136635,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_Formel_9.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L271"
  },
  {
    "id": "FX-010",
    "sourceId": "formula_0010",
    "title": "Datum umwandeln",
    "category": "Formeln",
    "summary": "March 6 2020 8:00:12 AM February 28 2020 11:17:26 AM January 24 2020 3:22:44 PM December 30 2023 1:38 PM November 22 2024 7:29 AM 8/31/2025 8:15:57 AM 4/24/2022 9:31:16 AM 9/24/2026 …",
    "description": "March 6 2020 8:00:12 AM\nFebruary 28 2020 11:17:26 AM\nJanuary 24 2020 3:22:44 PM\nDecember 30 2023 1:38 PM\nNovember 22 2024 7:29 AM\n8/31/2025 8:15:57 AM\n4/24/2022 9:31:16 AM\n9/24/2026 9:31:16 PM\n2026-11-22T14:30:00Z\n2026-08-12T14:30:00Z\n\nDiese Daten werden umgewandelt. Auch UTC nach MEZ/MESZ.\nVBA_Power_Query_Formel_Datum_umwandeln.xlsb\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Demo_VBA_PQ_Formel\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\nFormeln in C1, E1, H1, K2 und K3.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "formel",
      "datum",
      "utc",
      "mez",
      "mesz",
      "vba",
      "power query"
    ],
    "codeText": "=VSTAPELN(\"Daten\";MAP(DatenENUS[#Daten];LAMBDA(x;LET(a;TEXTTEILEN(GLÄTTEN(x);\" \");n;SPALTEN(a);ds;INDEX(a;1;1);ts;INDEX(a;1;n-1);ap;INDEX(a;1;n);dp;WENN(ISTZAHL(SUCHEN(\"/\";ds));TEXTTEILEN(ds;\"/\");\"\");dt;WENN(ISTZAHL(SUCHEN(\"/\";ds));DATUM(--INDEX(dp;1;3);--INDEX(dp;1;1);--INDEX(dp;1;2));DATUM(--INDEX(a;1;3);XVERGLEICH(INDEX(a;1;1);{\"January\".\"February\".\"March\".\"April\".\"May\".\"June\".\"July\".\"August\".\"September\".\"October\".\"November\".\"December\"});--INDEX(a;1;2)));tp;TEXTTEILEN(ts;\":\");tm;ZEIT(REST(--INDEX(tp;1;1);12)+12*(ap=\"PM\");--INDEX(tp;1;2);WENNFEHLER(--INDEX(tp;1;3);0));dt+tm))))\n=LET(z;WEGLASSEN(C1#;1);VSTAPELN({\"Datum\".\"Uhrzeit\"};HSTAPELN(GANZZAHL(z);REST(z;1))))\n=VSTAPELN({\"Datum\".\"Uhrzeit\"};LET(z;MAP(DatenENUS[#Daten];LAMBDA(x;LET(a;TEXTTEILEN(GLÄTTEN(x);\" \");n;SPALTEN(a);ds;INDEX(a;1;1);ts;INDEX(a;1;n-1);ap;INDEX(a;1;n);dp;WENN(ISTZAHL(SUCHEN(\"/\";ds));TEXTTEILEN(ds;\"/\");\"\");dt;WENN(ISTZAHL(SUCHEN(\"/\";ds));DATUM(--INDEX(dp;1;3);--INDEX(dp;1;1);--INDEX(dp;1;2));DATUM(--INDEX(a;1;3);XVERGLEICH(INDEX(a;1;1);{\"January\".\"February\".\"March\".\"April\".\"May\".\"June\".\"July\".\"August\".\"September\".\"October\".\"November\".\"December\"});--INDEX(a;1;2)));tp;TEXTTEILEN(ts;\":\");tm;ZEIT(REST(--INDEX(tp;1;1);12)+12*(ap=\"PM\");--INDEX(tp;1;2);WENNFEHLER(--INDEX(tp;1;3);0));dt+tm)));HSTAPELN(GANZZAHL(z);REST(z;1))))\n=LET(v;A10;w;WERT(WECHSELN(TEIL(v;1;19);\"T\";\" \"));x;DATUM(JAHR(w);3;31-WOCHENTAG(DATUM(JAHR(w);3;31)+1));y;DATUM(JAHR(w);10;31-WOCHENTAG(DATUM(JAHR(w);10;31)+1));z;1+ WENN(UND(w>=x;w<y);1;0);w+z/24)\n=LET(v;A11;w;WERT(WECHSELN(TEIL(v;1;19);\"T\";\" \"));x;DATUM(JAHR(w);3;31-WOCHENTAG(DATUM(JAHR(w);3;31)+1));y;DATUM(JAHR(w);10;31-WOCHENTAG(DATUM(JAHR(w);10;31)+1));z;1+ WENN(UND(w>=x;w<y);1;0);w+z/24)",
    "isNew": false,
    "package": {
      "fileName": "Demo_VBA_PQ_Formel.zip",
      "folderName": "Demo_VBA_PQ_Formel",
      "sizeBytes": 37133,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_VBA_PQ_Formel.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L308"
  },
  {
    "id": "FX-011",
    "sourceId": "formula_0011",
    "title": "Langer Text in Zellen aufteilen...",
    "category": "Formeln",
    "summary": "Lange Texte in Zellen werden nach vorgegebener Anzahl von Zeichen aufgetrennt. Nicht mitten im Wort. Verschiedene Möglichkeiten (in Spalten, Zeilen, Blöcken...). Mit RegEx, Power Query, …",
    "description": "Lange Texte in Zellen werden nach vorgegebener Anzahl von Zeichen aufgetrennt.\nNicht mitten im Wort.\nVerschiedene Möglichkeiten (in Spalten, Zeilen, Blöcken...).\nMit RegEx, Power Query, Formeln und VBA (UDF und Sub).\n\nRegEx_Power_Query_Formel_VBA_Text_auftrennen.xlsb\n\nBeim Klick auf \"Demo erzeugen\" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:\n%TEMP%\\Excel-VSTO-Toolbox\\Demo_RegEx_PQ_Formel_VBA\n\nEin bereits vorhandener Demo-Ordner wird vorher gelöscht.\nAnschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.\n\n!!!!!!!!WICHTIG!!!!!!!!\nFalls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der\nvorhandene Ordner nicht gelöscht und das Beispiel nicht erneut\nbereitgestellt werden.\n!!!!!!!!WICHTIG!!!!!!!!",
    "tags": [
      "auftrennenl",
      "text",
      "aufteilen",
      "formel",
      "regex",
      "vba",
      "power query"
    ],
    "codeText": "=LET(x;A2:.A999;lg;$D$1;WENNNV(WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;txt;VSTAPELN(A;LET(w;TEXTTEILEN(GLÄTTEN(txt);\" \");erg;REDUCE(\"\";w;LAMBDA(acc;z;LET(V;TEXTTEILEN(acc;\"♦\");lst;INDEX(V;1;SPALTEN(V));WENN(acc=\"\";z;WENN(LÄNGE(lst)+1+LÄNGE(z)<=lg;acc&\" \"&z;acc&\"♦\"&z)))));TEXTTEILEN(erg;\"♦\")))));1);\"\"))\n\n=LET(txt;GLÄTTEN(A2);lg;$C$1;w;TEXTTEILEN(txt;\" \");erg;REDUCE(\"\";w;LAMBDA(A;z;LET(V;TEXTTEILEN(A;\"♦\");letzte;INDEX(V;1;SPALTEN(V));WENN(A=\"\";z;WENN(LÄNGE(letzte)+1+LÄNGE(z)<=lg;A&\" \"&z;A&\"♦\"&z)))));MTRANS(TEXTTEILEN(erg;\"♦\")))\n=LET(x;A2:.A999;lg;C1;WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;txt;VSTAPELN(A;LET(w;TEXTTEILEN(GLÄTTEN(txt);\" \");erg;REDUCE(\"\";w;LAMBDA(acc;z;LET(V;TEXTTEILEN(acc;\"♦\");lst;INDEX(V;1;SPALTEN(V));WENN(acc=\"\";z;WENN(LÄNGE(lst)+1+LÄNGE(z)<=lg;acc&\" \"&z;acc&\"♦\"&z)))));MTRANS(TEXTTEILEN(erg;\"♦\"))))));1))\n=LET(x;A2:.A999;lg;$C$1;WENNNV(WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;txt;HSTAPELN(A;MTRANS(LET(w;TEXTTEILEN(GLÄTTEN(txt);\" \");erg;REDUCE(\"\";w;LAMBDA(acc;z;LET(V;TEXTTEILEN(acc;\"♦\");lst;INDEX(V;1;SPALTEN(V));WENN(acc=\"\";z;WENN(LÄNGE(lst)+1+LÄNGE(z)<=lg;acc&\" \"&z;acc&\"♦\"&z)))));TEXTTEILEN(erg;\"♦\"))))));;1);\"\"))\n\n=MTRANS(fncTextSplit(A2;25))\n=LET(x;A2:.A999;WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;T;VSTAPELN(A;MTRANS(fncTextSplit(T;$C$1)))));1))\n=fncTexteSplit(A2;C1)\n=fncTextSplit(A2;25)\n=fncTexteSplitZ(A2:A14;C1)\n=LET(x;A2:A14;WENNNV(WEGLASSEN(REDUCE(\"\";x;LAMBDA(A;T;HSTAPELN(A;MTRANS(fncTextSplit(T;$C$1)))));;1);\"\"))",
    "isNew": true,
    "package": {
      "fileName": "Demo_RegEx_PQ_Formel_VBA.zip",
      "folderName": "Demo_RegEx_PQ_Formel_VBA",
      "sizeBytes": 73867,
      "downloadUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/Demo_RegEx_PQ_Formel_VBA.zip"
    },
    "sourceUrl": "https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/RegExPQ/DemoCatalog.Formula.vb#L358"
  }
];
