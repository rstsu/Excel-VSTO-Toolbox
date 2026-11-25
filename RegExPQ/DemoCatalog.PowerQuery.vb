Imports System.Xml.Linq

Partial Public Module DemoCatalog

    Private Function GetPowerQueryExamples() As List(Of DemoDefinition)
        Return New List(Of DemoDefinition) From {
            New DemoDefinition With {
                .Id = "pq_001",
                .Category = DemoCategory.PowerQuery,
                .Title = "Wörter je Nummer gruppieren",
                .Tags = {"regex", "m-code", "gruppieren", "table.group", "text", "text.combine"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Text/Wörter werden je gleicher Nummer in einer Spalte gruppiert bzw. mit Leerzeichen zusammengefasst.
Also Spalte A sind Nummern und Spalte B Wörter. Sind die Nummern in Spalte A gleich, werden die Wörter gruppiert.

Nach Änderungen in der Grundtabelle wird die Abfrage mit STRG+ALT+F5 aktualisiert!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_1"]}[Content],
    Gruppe = Table.Group(Quelle, {"Nr"}, {{"Wörter", each _, type table [Nr=nullable number, Wörter=nullable text]}}),
    HinzuSpalte = Table.AddColumn(Gruppe, "Alle", each Text.Combine(List.Distinct(List.Transform([Wörter][Wörter], each Text.Trim(_))), " "))
in
    HinzuSpalte
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_002",
                .Category = DemoCategory.PowerQuery,
                .Title = "Zollangaben aus Text auslesen",
                .Tags = {"regex", "m-code", "zahl", "zoll", "text", "extrahieren"},
                .Description = TextBlock(
    <text>
        <![CDATA[
In den Zellen A2:Ax sind Texte mit mehreren Zollangaben (z. B. 1/2" oder dann auch 1 3/4").
Die Zollangaben sollen ausgelesen werden. Auch mehrere.
Überschriften (Z1, Z2...) automatisch generieren.
Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_2"]}[Content],
    AddZollListe = Table.AddColumn(Quelle, "ZollListe", each
        let
            v = [Beschreibung],
            protected =
                List.Accumulate(
                    {"1 1/2""","2 1/2""","3 1/2""","4 1/2""","5 1/2""","6 1/2""","7 1/2""","8 1/2""","9 1/2"""},
                    v,
                    (w,x) => Text.Replace(w, x, Text.Replace(x, " ", "~"))
                ),
            y = Text.SplitAny(protected, " ,.-:;()[]"),
            z = List.Transform(
                List.Select(y,
                    each Text.EndsWith(_, """") and Text.Length(Text.Select(_, {"0".."9"})) > 0
                ),
                each Text.Replace(_, "~", " ")
            )
        in
            z),
    MaxAnzahl = List.Max(List.Transform(AddZollListe[ZollListe], each List.Count(_))),
    ZNames = List.Transform({1..MaxAnzahl}, each "Z" & Text.From(_)),
    AddRecord = Table.AddColumn(AddZollListe, "Z", each Record.FromList(List.FirstN([ZollListe] & List.Repeat({null}, MaxAnzahl), MaxAnzahl), ZNames)),
    Expand = Table.ExpandRecordColumn(AddRecord, "Z", ZNames, ZNames),
    Erg = Table.SelectColumns(Expand, List.Transform(ZNames, each Text.From(_)))
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_003",
                .Category = DemoCategory.PowerQuery,
                .Title = "Inhalt des TEMP-Ordners auflisten",
                .Tags = {"regex", "m-code", "tmp", "temp", "text", "auflisten"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Der Inhalt des lokalen TEMP-Ordners...
(C:\Users\USERNAME\AppData\Local\Temp)
...wird aufgelistet.
Der USERNAME wird über eine Formel ausgelesen.
=TEXTVOR(ZELLE("filename");"\";3)&"\AppData\Local\Temp"
!!!!!!!!WICHTIG!!!!!!!!
Die Datei, in der das probiert wird MUSS gespeichert sein.
In einer neuen UNGESPEICHERTEN Datei geht das NICHT!
Da erscheint dann der Fehler #WERT! in Zelle G1.
!!!!!!!!WICHTIG!!!!!!!!
Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Folder.Files(Excel.CurrentWorkbook(){[Name="Benutzername"]}[Content]{0}[Column1]),
    Hinzu = Table.AddColumn(Quelle, "Datei", each [Folder Path]&[Name]),
    Erg = Table.SelectColumns(Hinzu,{"Datei", "Date accessed", "Date modified", "Date created"})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_004",
                .Category = DemoCategory.PowerQuery,
                .Title = "Sonderzeichen entfernen",
                .Tags = {"sonderzeichen", "m-code", "entfernen", "mehrere", "text"},
                .Description = TextBlock(
    <text>
        <![CDATA[
In A2:A8 stehen Texte mit Sonderzeichen - auch mehrere hintereinander.
Diese sollen entfernt werden.
Zwischen den Wörtern darf nur ein Leerzeichen übrig bleiben.

Falls es nicht nur Leerzeichen sind, könnte man auch das nehmen:
Clean = Text.Combine(List.Select(Text.Split(DelS, " "), each _ <> ""),

Es würde auch über eine Funktion gehen (mit Namen "fncCleanText"):
(Text as text, optional RemoveChars as list) =>
let
    Chars = if RemoveChars = null then {":","?","*","/","\"} else RemoveChars,
    Clean = List.Accumulate(Chars, Text, (txt, c) => Text.Replace(txt, c, "")),
    Result = Text.Combine(List.RemoveItems(Text.Split(Clean, " "), {""}), " ")
in
    Result

Aufruf:
fncCleanText([Text])

Oder mit eigenen Zeichen:
fncCleanText([Text], {":","?","*","/","\",".",",",";"})
Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Erg =
        Table.TransformColumns(
            Excel.CurrentWorkbook(){[Name="Demo_PQ_4"]}[Content],
            {
                {"Text", each
                    let
                        DelS = List.Accumulate(
                            {":","?","*","/","\"},_,
                            (txt, char) => Text.Replace(txt, char, "")
                        ),
                        Clean = Text.Combine(List.RemoveItems(Text.Split(DelS, " "), {""}), " ")
                    in
                        Clean
                }
            }
        )
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_005",
                .Category = DemoCategory.PowerQuery,
                .Title = "Kreuztabelle aus Liste erste 3 Buchstaben",
                .Tags = {"power query", "text", "m-code", "tabelle", "liste"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A15) wird eine Kreuztabelle erstellt.
Grundlage sind die ersten 3 gleichen Buchstaben.
Es ist auch in Formeln gelöst. Mit der gleichen Bezeichnung.
Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.

Nach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_5"]}[Content],
    Kuerzel = Table.AddColumn(Quelle, "Kürzel", each Text.Start([Wert], 3), type text),
    Gruppe = Table.Group(Kuerzel, {"Kürzel"}, {{"Daten", each Table.AddIndexColumn(_, "Pos", 1, 1, Int64.Type)}}),
    Expand = Table.ExpandTableColumn(Gruppe, "Daten", {"Wert", "Pos"}),
    Pos = Table.TransformColumns(Expand, {{"Pos", each "Pos" & Text.From(_), type text}}),
    Pivot = Table.Pivot(Pos, List.Distinct(Pos[Pos]), "Pos", "Wert"),
    Erg = Table.RemoveColumns(Pivot, {"Kürzel"})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_006",
                .Category = DemoCategory.PowerQuery,
                .Title = "String zwischen 2 Zahlen auslesen",
                .Tags = {"power query", "text", "m-code", "tabelle", "liste"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A12) wird ein Text zwischen 2 Zahlem ausgelesen.
Z. B. "1. Vom Bodensee in den Schwarzwald 21 May 2001".
Es wird nur der "Titel" in der Mitte ausgelesen bzw. der Rest ersetzt.

Nach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_6"]}[Content],
    Ausgabe = Table.AddColumn(Quelle, "Ausgabe", each
        let
            ON = Text.AfterDelimiter([Daten], ". "),
            TE = Text.Split(Text.Trim(ON), " "),
            TI = Text.Combine(List.RemoveLastN(TE, 3), " ")
        in
            TI,
        type text
    ),
    Erg = Table.SelectColumns(Ausgabe, {"Ausgabe"})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_007",
                .Category = DemoCategory.PowerQuery,
                .Title = "Datumbrereich in Liste auflösen",
                .Tags = {"power query", "datum", "m-code", "tabelle", "ferien"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A6 - Ferientermine 2027 in Baden Württemberg) wird eine laufende Liste erstellt.

Der auskommentierte Code ist eine andere Herangehensweise ( also alles zwischen /* ...M-Code... */).

Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.

Nach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_7"]}[Content],
    Typen = Table.TransformColumnTypes(
        Quelle,
        {{"Ferien BW", type text}, {"Von", type date}, {"Bis", type date}}
    ),
    Erg = Table.Sort(
        Table.SelectColumns(
            Table.AddColumn(
                Table.ExpandListColumn(
                    Table.AddColumn(
                        Typen,
                        "Datum",
                        each List.Dates([Von], Duration.Days([Bis] - [Von]) + 1, #duration(1,0,0,0)),
                        type list
                    ),
                    "Datum"
                ),
                "Tag",
                each Date.DayOfWeekName([Datum]),
                type text
            ),
            {"Datum", "Tag"}
        ),
        {{"Datum", Order.Ascending}}
    )
in
    Erg
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_7"]}[Content],
    TypeG = Table.TransformColumnTypes(Quelle,{{"Ferien BW", type text}, {"Von", type date}, {"Bis", type date}}),
    ErwS = Table.AddColumn(TypeG, "Datum", each List.Dates([Von], Duration.Days([Bis] - [Von]) + 1, #duration(1,0,0,0))),
    ExpandT = Table.TransformColumnTypes(Table.ExpandListColumn(ErwS, "Datum"),{{"Datum", type date}}),
    TagN = Table.AddColumn(ExpandT, "Tag", each Date.DayOfWeekName([Datum]), type text),
    Erg = Table.Sort(Table.SelectColumns(TagN,{"Datum", "Tag"}), {{"Datum", Order.Ascending}})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_008",
                .Category = DemoCategory.PowerQuery,
                .Title = "Kreuztabelle aus Liste Materialnummern",
                .Tags = {"power query", "material", "m-code", "tabelle", "kreuztabelle"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:E11) wird eine Kreuztabelle erstellt.
Grundlage sind die Materialnummern in der ersten Spalte.
Spalten (Text) werden so viel wie nötig erstellt.
Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.

Nach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_8"]}[Content],
    Gruppe = Table.Group(Quelle, {"Materialnummer"}, {"Rows", each Table.AddIndexColumn(_, "Zeile", 1)}),
    Expand = Table.ExpandTableColumn(Gruppe, "Rows", {"Werk","Materialart","TextID","Text","Zeile"}),
    TypG = Table.TransformColumnTypes(Expand, {{"Zeile", type text}}),
    Pivot = Table.Pivot(TypG, List.Distinct(TypG[Zeile]), "Zeile", "Text"),
    Erg = Table.RenameColumns(Pivot, List.Transform(List.Distinct(TypG[Zeile]), each {_, "Text" & _}))
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_009",
                .Category = DemoCategory.PowerQuery,
                .Title = "Auflistung aus Gruppe - Von Bis erstellen",
                .Tags = {"power query", "auflistung", "m-code", "von", "bis"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:C12) wird eine Auflistung erstellt.
Grundlage ist die Gruppe in Spalte A. Dann kommt Von und Bis.
Berücksichtigt wird, ob Text statt Zahl eingegeben ist und ob es eine leere Eingabe ist.
Auch wird darauf geachtet, ob Von größer Bis ist.
Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.

Nach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!

Man kann auch eine eigene Funktion erstellen (Name der Funktion: fncGetWert):
let
    Quelle = (g as any, von as any, bis as any) as nullable list =>
let
    VonZahl = try Int64.From(von) otherwise null,
    BisZahl = try Int64.From(bis) otherwise null,
    VonIstText = von <> null and VonZahl = null,
    BisIstText = bis <> null and BisZahl = null
in
    if g = null then
        null
    else if VonIstText and BisZahl <> null then {BisZahl}
    else if BisIstText and VonZahl <> null then {VonZahl}
    else if VonZahl <> null and BisZahl = null then {1..VonZahl}
    else if VonZahl <> null and BisZahl <> null then
        if VonZahl <= BisZahl
        then {VonZahl..BisZahl}
        else {BisZahl..VonZahl}
    else
        null
in
    Quelle

Dann eine Abfrage (Name: tblErgfnc):
let
    Quelle = Table.Sort(Excel.CurrentWorkbook(){[Name="Demo_PQ_9"]}[Content],{{"Gruppe", Order.Ascending}}),
    OhneDuplikate = Table.Distinct(Quelle, {"Gruppe","Von","Bis"}),
    MitListe = Table.AddColumn(OhneDuplikate, "Wert", each fncGetWert([Gruppe],[Von],[Bis]), type list),
    Gefiltert = Table.SelectRows(MitListe, each [Gruppe] <> null and [Wert] <> null),
    Erg = Table.SelectColumns(Table.ExpandListColumn(Gefiltert, "Wert"), {"Gruppe","Wert"})
in
    Erg
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    // Mit der Funktion ermittle ich die Werteliste
    GetWert = (g as any, von as any, bis as any) as nullable list =>
        let
            VonZahl = try Int64.From(von) otherwise null,
            BisZahl = try Int64.From(bis) otherwise null,
            VonIstText = von <> null and VonZahl = null,
            BisIstText = bis <> null and BisZahl = null
        in
            if g = null then
                null
            else if VonIstText and BisZahl <> null then {BisZahl}
            else if BisIstText and VonZahl <> null then {VonZahl}
            else if VonZahl <> null and BisZahl = null then {1..VonZahl}
            else if VonZahl <> null and BisZahl <> null then
                if VonZahl <= BisZahl
                then {VonZahl..BisZahl}
                else {BisZahl..VonZahl}
            else null,
    Quelle = Table.Sort(Excel.CurrentWorkbook(){[Name="Demo_PQ_9"]}[Content],{{"Gruppe", Order.Ascending}}),
    OhneDuplikate = Table.Distinct(Quelle, {"Gruppe","Von","Bis"}),
    MitListe = Table.AddColumn(OhneDuplikate, "Wert", each GetWert([Gruppe],[Von],[Bis]), type list),
    Gefiltert = Table.SelectRows(MitListe, each [Gruppe] <> null and [Wert] <> null),
    Erg = Table.SelectColumns(Table.ExpandListColumn(Gefiltert, "Wert"), {"Gruppe","Wert"})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0010",
                .Category = DemoCategory.PowerQuery,
                .Title = "Blockbildung mit Zwischensummen und Gesamtsumme",
                .Tags = {"power query", "zwischensumme", "m-code", "gesamtsumme", "blockbildung"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:C16) werden Blöcke je Code mit Zwischensummen gebildet.
Die Gesamtsumme kann auch über "Tabellenentwurf - Tabellenformatoptionen - Haken bei Ergebniszeile setzen" angzeigt werden.
Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.

Nach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!

Mit folgendem M-Code wird die Gesamtsumme direkt am Ende ausgegeben:

let
    Quelle = Table.TransformColumnTypes(Excel.CurrentWorkbook(){[Name="Demo_PQ_10"]}[Content], {{"Datum", type date}, {"Code", Int64.Type}, {"Betrag", type number}}),
    // Falls "Erhebniszeile" in der Grundtabelle eingeblendet ist
    EntfF = Table.RemoveRowsWithErrors(Quelle, {"Datum"}),
    Gruppe = Table.Sort(Table.Group(EntfF, {"Code"}, {{"x", each _, type table}}),{{"Code", Order.Ascending}}),
    Block = List.Transform(Gruppe[x], (x) =>
        let
            y = List.Sum(x[Betrag]),
            z = #table({"Datum", "Code", "Betrag"}, {{null, null, y}})
        in
            Table.Combine({x, z})
        ),
    // Nach Ausgabe im Tabellenblatt die Spalte Datum formatieren!
    Erg = Table.InsertRows(Table.Combine(Block), Table.RowCount(EntfF)+Table.RowCount(Gruppe), {[Datum = null, Code = null, Betrag = List.Sum(EntfF[Betrag])]})
in
    Erg

Man kann die Gesamtsumme auch per Formel (Teilergebnis und Bereich.Verschieben) ausgeben. Hier zwei Möglichkeiten (Tabelle ist über Beispiel Demo schon im Tabellenblatt ausgegeben und Abfrage1 wurde NICHT umbenannt):

=SUMMENPRODUKT((ISTZAHL(Abfrage1[Datum]))*TEILERGEBNIS(109;BEREICH.VERSCHIEBEN(Abfrage1[Betrag];ZEILE(Abfrage1[Betrag])-MIN(ZEILE(Abfrage1[Betrag]));0;1)))

Und mit LET:
=LET(x;Abfrage1[Datum];y;Abfrage1[Betrag];z;TEILERGEBNIS(109;BEREICH.VERSCHIEBEN(y;ZEILE(y)-MIN(ZEILE(y));0;1));SUMMENPRODUKT((x<>"")*z))
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Table.TransformColumnTypes(Excel.CurrentWorkbook(){[Name="Demo_PQ_10"]}[Content], {{"Datum", type date}, {"Code", Int64.Type}, {"Betrag", type number}}),
    // Falls "Erhebniszeile" in der Grundtabelle eingeblendet ist
    EntfF = Table.RemoveRowsWithErrors(Quelle, {"Datum"}),
    Gruppe = Table.Sort(Table.Group(EntfF, {"Code"}, {{"x", each _, type table}}),{{"Code", Order.Ascending}}),
    Block = List.Transform(Gruppe[x], (x) =>
        let
            y = List.Sum(x[Betrag]),
            z = #table({"Datum", "Code", "Betrag"}, {{null, null, y}})
        in
            Table.Combine({x, z})
        ),
    // Nach Ausgabe im Tabellenblatt die Spalte Datum formatieren!
    Erg = Table.Combine(Block)
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0011",
                .Category = DemoCategory.PowerQuery,
                .Title = "Anmeldename bilden aus Vor- und Nachname",
                .Tags = {"power query", "vorname", "m-code", "nachname", "anmeldename"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:B10) werden Anmeldenamen aus Vorname und Nachname gebildet.
Der Vorname nur 3 Buchstaben. Aus "Paul Huber" wird "Huber.Pau". Bei doppeltem Vor- und Nachname wird eine Zahl angehängt.

Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.

Nach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_11"]}[Content],
    HinzuVN = Table.AddColumn(Quelle, "Nachname.Vorname", each [Nachname] & "." & Text.Start([Vorname], 3)),
    Gruppe = Table.Group(HinzuVN, {"Nachname.Vorname"}, {{"Alle Zeilen", each Table.AddIndexColumn(_, "Index", 1, 1, Int64.Type), type table [Vorname=nullable text, Nachname=nullable text, Nachname.Vor=nullable text, Index=Int64.Type]}}),
    Expand = Table.ExpandTableColumn(Gruppe, "Alle Zeilen", {"Vorname", "Nachname", "Index"}),
    Final = Table.AddColumn(Expand, "FinalerNachnameVor", each if [Index] > 1 then [Nachname.Vorname] & Text.From([Index]) else [Nachname.Vorname]),
    Erg = Table.SelectColumns(Table.RenameColumns(Final,{{"FinalerNachnameVor", "NachnameVorneme3"}}), {"Vorname", "Nachname", "NachnameVorneme3"})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0012",
                .Category = DemoCategory.PowerQuery,
                .Title = "2 Wohnungen - Vermietung - nicht belegte Tage",
                .Tags = {"power query", "vermietung", "m-code", "datum", "belegung"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:C30) werden aus der Vermietung von 2 Wohnungen (Wohnung, Anreise, Abreise) die Tage abgebildet, welche NICHT belegt sind.
Wie der Abreisetag gerechnet werden kann/soll, ist im Kommentar des M-Codes gezeigt.

Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.

Nach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_12"]}[Content],
    TypG = Table.TransformColumnTypes(Quelle, {{"Anreise", type date}, {"Abreise", type date}, {"Wohnung", type text}}),
    MinAnreise = if List.Count(TypG[Anreise]) > 0 then List.Min(TypG[Anreise]) else null,
    MaxAbreise = if List.Count(TypG[Abreise]) > 0 then List.Max(TypG[Abreise]) else null,
    DayB = if MinAnreise <> null and MaxAbreise <> null then Duration.Days(MaxAbreise - MinAnreise) else 0,
    DateT = if MinAnreise <> null then Table.FromList(List.Dates(MinAnreise, DayB + 1, #duration(1, 0, 0, 0)), Splitter.SplitByNothing(), {"Nicht Belegt"}, null, ExtraValues.Error) else #table({"Nicht Belegt"}, {}),
    Whg1 = Table.AddColumn(DateT, "Whg1", each if Table.RowCount(Table.SelectRows(TypG, (row) => row[Wohnung] = "Whg. 1" and row[Anreise] <= [Nicht Belegt] and row[Abreise] > [Nicht Belegt])) > 0 then 1 else 0),
    // Wie der Abreisetag gerechnet werden soll liegt an > oder >=
    //Whg1 = Table.AddColumn(DateT, "Whg1", each if Table.RowCount(Table.SelectRows(TypG, (row) => row[Wohnung] = "Whg. 1" and row[Anreise] <= [Nicht Belegt] and row[Abreise] >= [Nicht Belegt])) > 0 then 1 else 0),
    Whg2 = Table.AddColumn(Whg1, "Whg2", each if Table.RowCount(Table.SelectRows(TypG, (row) => row[Wohnung] = "Whg. 2" and row[Anreise] <= [Nicht Belegt] and row[Abreise] > [Nicht Belegt])) > 0 then 1 else 0),
    // Wie der Abreisetag gerechnet werden soll liegt an > oder >= bei row[Abreise] >= [Nicht Belegt]
    //Whg2 = Table.AddColumn(Whg1, "Whg2", each if Table.RowCount(Table.SelectRows(TypG, (row) => row[Wohnung] = "Whg. 2" and row[Anreise] <= [Nicht Belegt] and row[Abreise] >= [Nicht Belegt])) > 0 then 1 else 0),
    //Result = Table.SelectColumns(Table.SelectRows(Whg2, each ([Whg1] = 0) and ([Whg2] = 0)), {"Nicht Belegt"}),
    Result = Table.TransformColumnTypes(Table.SelectColumns(Table.SelectRows(Whg2, each ([Whg1] = 0) and ([Whg2] = 0)), {"Nicht Belegt"}),{{"Nicht Belegt", type date}}, "de-DE"),
    Monat = Table.AddColumn(Result, "Monat", each Date.MonthName([Nicht Belegt]), type text),
    KW = Table.AddColumn(Monat, "KW", each let d = [Nicht Belegt], shifted = Date.AddDays(d, 3 - Date.DayOfWeek(d, Day.Monday)) in Date.WeekOfYear(shifted, Day.Monday), Int64.Type),
    Jahr = Table.AddColumn(KW, "Jahr", each let d = [Nicht Belegt], shifted = Date.AddDays(d, 3 - Date.DayOfWeek(d, Day.Monday)) in Date.Year(shifted), Int64.Type),
    KWText = Table.AddColumn(Jahr, "KWText", each Text.From([Jahr]) & "-KW " & Text.PadStart(Text.From([KW]), 2, "0"), type text),
    Erg = Table.Sort(KWText, {{"Nicht Belegt", Order.Ascending}})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0013",
                .Category = DemoCategory.PowerQuery,
                .Title = "Items zusammenfassen - auf 2 Arten",
                .Tags = {"power query", "pivotieren", "m-code", "entpivotieren", "gruppieren"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:D11) werden die "items" (Spalte B:D) nach dem Namen (Spalte A) zusammengefasst.
Einmal pivotiert und einmal als Liste mit zusammengefassten "items".

!!!!!!!!WICHTIG!!!!!!!!
Bei diesem Beispiel wird die Query UND die Ausgabe im Tabellenblatt DIREKT erzeugt!
Es muss also kein M-Code kopiert werden. Die Ausgabe dauert einen Augenblick, da alles generiert wird!
!!!!!!!!WICHTIG!!!!!!!!

Für Informationen, wie mit dem M-Code umzugehen ist, auf den Button "PQ M-Code Info" klicken.
Um die Beschreibung wieder zu sehen, auf die Bezeichnung klicken.

Nach Änderungen in der Grundtabelle - Abfrage mit STRG+ALT+F5 aktualisieren!

Das ist der zweite M-Code (Beide M-Codes sind im Beispiel schon drin):
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_13"]}[Content],
    Entpivotieren = Table.UnpivotOtherColumns(Quelle, {"Name"}, "Attribut", "Wert"),
    EntferneD = Table.Distinct(Entpivotieren, {"Name", "Wert"}),
    GruppeZ = Table.Group(EntferneD, {"Name"}, {{"items", each _, type table [Name=text, Attribut=text, Wert=text]}}),
    HinzuS = Table.AddColumn(GruppeZ, "itemsN", each Text.Combine(List.Sort([items][Wert]), ", ")),
    EntferneS = Table.RemoveColumns(HinzuS, {"items"})
in
    EntferneS
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
let
    Quelle = Excel.CurrentWorkbook(){[Name="Demo_PQ_13"]}[Content],
    Entpivotieren = Table.UnpivotOtherColumns(Quelle, {"Name"}, "Attribut", "Wert"),
    EntferneD = Table.Distinct(Entpivotieren, {"Name", "Wert"}),
    PivotS = Table.Pivot(EntferneD, List.Distinct(EntferneD[Attribut]), "Attribut", "Wert")
in
    PivotS
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0014",
                .Category = DemoCategory.PowerQuery,
                .Title = "Geburtstagsliste der nächsten X Tage anzeigen",
                .Tags = {"power query", "beispieldatei", "m-code", "geburtstag", "zip"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Dieses Beispiel besteht aus einer Excel-Datei.
PQ_Geburtstagsliste_Anzahl_Tage_vorgeben.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv
in folgenden Ordner entpackt:

%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Es sind zwei Arten des Umgangs im M-Code. Es geht einmal darum,
wie mit den "TageBisGeburtstag" umgegangen wird. Entweder
bei heute Geburtstag 0 oder 365/366. Auch der 29.2. wird anders
behandelt.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
'Bitte auf "Demo erzeugen" klicken um die Beispieldatei
'zu entpacken und die XLSX zu öffnen.
'Mit folgendem VBA-Code kann die Aktualisierung
'automatisch erfolgen:
Option Explicit
'Excel -VSTO - Toolbox
'Power Query - Demo
'Ralf Stolzenburg (Case)
'https://github.com/rstsu/Excel-VSTO-Toolbox
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Fin
    If Target.Address = "$E$1" Then
        Application.EnableEvents = False
        Me.ListObjects("tblErg_1").QueryTable.Refresh
        Me.ListObjects("tblErg_2").QueryTable.Refresh
    End If
Fin:
    Application.Goto Range("E1"), False
    Application.EnableEvents = True
End Sub
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0015",
                .Category = DemoCategory.PowerQuery,
                .Title = "Tage und Feiertage über Kontrollkästchen ausgeben",
                .Tags = {"power query", "beispieldatei", "m-code", "feiertage", "zip"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Dieses Beispiel besteht aus einer Excel-Datei.
PQ_Power_Query_Feiertage_und_Tage_von_Montag_bis_Sonntag_ueber_Kontrollkaestchen_ausgeben.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv
in folgenden Ordner entpackt:

%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Tage und Feiertage (mit KW und Quartal) werden als Liste ausgegeben.
Auswahl der Tage über Kontrollkästchen -
(die "neuen" über Einfügen - Kontrollkästchen).

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
'Bitte auf "Demo erzeugen" klicken um die Beispieldatei
'zu entpacken und die XLSB zu öffnen.
'Folgender VBA-Code ist schon in der Datei.
'Hier nochmal zum kopieren:
Option Explicit
'Excel -VSTO - Toolbox
'Power Query - Demo
'Ralf Stolzenburg (Case)
'https://github.com/rstsu/Excel-VSTO-Toolbox
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Fin
    If Not Intersect(Target, Range("B2, J2:Q2")) Is Nothing Then
        Application.EnableEvents = False
        Me.ListObjects("tblErg").QueryTable.Refresh
    End If
Fin:
    Application.Goto Range("B2"), False
    Application.EnableEvents = True
End Sub
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0016",
                .Category = DemoCategory.PowerQuery,
                .Title = "Nummern mit 0 rechts auffüllen - auf 5 Stellen",
                .Tags = {"power query", "beispieldatei", "m-code", "nummern", "auffüllen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Dieses Beispiel besteht aus einer Excel-Datei.
PQ_UND_Formel_immer_auf_5_Stellen_RECHTS_mit_0_Null_auffuellen_Groesser_5_rechts_abtrennen_ODER_lassen.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv
in folgenden Ordner entpackt:

%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Nummern werden mit 0 von rechts auf 5 Stellen aufgefüllt.
Sind mehr als 5 Stellen vorhanden, wird einmal gekürzt und einmal nicht.
Es ist auch mit Formeln in der Beispieldatei gelöst.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
'Bitte auf "Demo erzeugen" klicken um die Beispieldatei
'zu entpacken und die XLSX zu öffnen.
'Mit folgendem VBA-Code kann die Aktualisierung
'automatisch erfolgen:
Option Explicit
'Excel -VSTO - Toolbox
'Power Query - Demo
'Ralf Stolzenburg (Case)
'https://github.com/rstsu/Excel-VSTO-Toolbox
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Fin
    If Target.Column = 1 And Target.Row > 1 Then
        Application.EnableEvents = False
        Me.ListObjects("tblErg_1").QueryTable.Refresh
        Me.ListObjects("tblErg_2").QueryTable.Refresh
    End If
    Application.Goto Range(Target.Address), False
Fin:
    Application.EnableEvents = True
End Sub
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0017",
                .Category = DemoCategory.PowerQuery,
                .Title = "Zeitstempel bei bestimmtem Text eintragen oder löschen",
                .Tags = {"power query", "uhrzeit", "m-code", "zeitstempel", "datum"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Dieses Beispiel besteht aus einer Excel-Datei.
PQ_Zeitstempel_TT_MM_JJJJ_hh_mm_ss_setzen_wenn_Spalte_C_leer_und_bestimmter_Text_in_Spalte_A.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv
in folgenden Ordner entpackt:

%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Wenn in Spalte "Wert_1"  "Neu" steht und in der Spalte "Zeitstempel" kein Eintrag vorhanden ist, wird das aktuelle Datum mit Zeit eingetragen.
Wenn in Spalte "Wert_1"  "Alt" steht, wird der Zeitstempel geleert.
tblErg_1 - Der Zeitstempel wird getrennt in Datum und Zeit.
tblErg_2 - mit "Replacer.ReplaceValue" andere Herangehensweise.
tblErg_2_1 - Der Zeitstempel wird nicht getrennt.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
'Bitte auf "Demo erzeugen" klicken um die Beispieldatei
'zu entpacken und die XLSX zu öffnen.
'Mit folgendem VBA-Code kann die Aktualisierung
'automatisch erfolgen:
Option Explicit
'Excel -VSTO - Toolbox
'Power Query - Demo
'Ralf Stolzenburg (Case)
'https://github.com/rstsu/Excel-VSTO-Toolbox
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Fin
    If Target.Column = 1 And Target.Row > 1 Then
        Application.EnableEvents = False
        Tabelle2.ListObjects("tblErg_1").QueryTable.Refresh
        Tabelle3.ListObjects("tblErg_2").QueryTable.Refresh
        Tabelle4.ListObjects("tblErg_2_1").QueryTable.Refresh
    End If
    Application.Goto Range(Target.Address), False
Fin:
    Application.EnableEvents = True
End Sub
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0018",
                .Category = DemoCategory.PowerQuery,
                .Title = "Datum umwandeln",
                .Tags = {"formel", "datum", "utc", "mez", "mesz", "vba", "power query"},
                .Description = TextBlock(
    <text>
        <![CDATA[
March 6 2020 8:00:12 AM
February 28 2020 11:17:26 AM
January 24 2020 3:22:44 PM
December 30 2023 1:38 PM
November 22 2024 7:29 AM
8/31/2025 8:15:57 AM
4/24/2022 9:31:16 AM
9/24/2026 9:31:16 PM
2026-11-22T14:30:00Z
2026-08-12T14:30:00Z

Diese Daten werden umgewandelt. Auch UTC nach MEZ/MESZ.
VBA_Power_Query_Formel_Datum_umwandeln.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    Erg = Table.TransformColumnTypes(Quelle, {{"Daten", type datetime}}, "en-US")
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    Parsen = Table.TransformColumnTypes(Quelle, {{"Daten", type datetime}}, "en-US"),
    Erg = Table.TransformColumnTypes(Table.SplitColumn(Table.TransformColumnTypes(Parsen, {{"Daten", type text}}, "de-DE"), "Daten", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"Datum", "Uhrzeit"}),{{"Datum", type date}, {"Uhrzeit", type time}})
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Table.FromRows({
        {"March 6 2020 5:00 AM"},          // Englisch
        {"Février 28 2020 5:00 AM"},       // Französisch
        {"Enero 15 2021 14:20"},           // Spanisch
        {"Giugno 5 2022 8:15"},            // Italienisch
        {"November 22 2019 7:29 AM"},      // Deutsch/Englisch
        {"8/31/2012 8:15:57 AM"},          // US numerisch
        {"4/24/2012 9:31:16 AM"},          // US numerisch
        {"22/08/2026"},                    // DE numerisch
        {"15.09.2025 14:45"},              // DE Punktformat
        {"2026-11-22T14:30:00Z"},          // ISO UTC
        {"2026-08-12T14:30:00Z"},          // ISO UTC
        {"2026-08-17"}                     // ISO Datum
    }, {"Original"}),
    Lokal = {"en-US", "de-DE", "fr-FR", "es-ES", "it-IT", "nl-NL", "pt-PT"},
    UTCLokal = (dt as datetime) as datetime =>
        let
            Zone = DateTimeZone.From(dt),
            ZoneL = DateTimeZone.ToLocal(Zone),
            ZoneN = DateTimeZone.RemoveZone(ZoneL)
        in
            ZoneN,
    Parse = (val as any) as nullable datetime =>
        let
            DateOK = if Value.Is(val, type datetime) then val else null,
            txt = if DateOK = null and Value.Is(val, type text) then Text.Trim(val) else null,
            ParseISO = if txt <> null and Text.Length(txt) >= 4 and Text.Range(txt,0,4) >= "0001" and Text.Range(txt,0,4) <= "9999" and Text.Middle(txt,4,1) = "-" 
                        then try DateTime.FromText(txt, "en-US") otherwise null
                        else null,
            parsedLocales = if DateOK = null and ParseISO = null and txt <> null then
                                let
                                    tryLocales = List.First(
                                        List.RemoveNulls(
                                            List.Transform(Lokal, each try DateTime.FromText(txt, _) otherwise null)
                                        ),
                                        null
                                    )
                                in
                                    tryLocales
                            else null,
            ParseA = List.First(List.RemoveNulls({DateOK, ParseISO, parsedLocales}), null),
            result = if txt <> null and Text.EndsWith(txt, "Z") and ParseA <> null
                    then UTCLokal(ParseA)
                    else ParseA
        in
            result,
    MitDatumZeit = Table.AddColumn(
        Quelle,
        "DatumZeit",
        each Parse([Original]),
        type nullable datetime
    ),
    MitDatum = Table.AddColumn(
        MitDatumZeit,
        "Datum",
        each if [DatumZeit] <> null then Date.From([DatumZeit]) else null,
        type date
    ),
    MitUhrzeit = Table.AddColumn(
        MitDatum,
        "Uhrzeit",
        each if [DatumZeit] <> null then Time.From([DatumZeit]) else null,
        type time
    ),
    Erg = Table.RemoveColumns(MitUhrzeit, {"DatumZeit"})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0019",
                .Category = DemoCategory.PowerQuery,
                .Title = "Biorhythmus mit Diagramm...",
                .Tags = {"pq", "power query", "chart", "diagramm", "biorhythmus", "vba"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Auf dem Tabellenblatt "Biorhythmus" werden im Diagramm alle Linien angezeigt (die Punkte der "kritischen Tage" können ausgewählt werden.
Im Tabellenblatt "Biorhythmus_Part" können die anzuzeigenden Linien ausgewählt werden.
Aktualisiert wird über VBA.

PQ_Biorhythmus_Worksheet_Change_Diagramm_Chart.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    GebTagRaw  = Excel.CurrentWorkbook(){[Name="GebTag"]}[Content]{0}[Column1],
    TageRaw    = Excel.CurrentWorkbook(){[Name="Tage"]}[Content]{0}[Column1],
    KKurvenRaw = Excel.CurrentWorkbook(){[Name="KKurven"]}[Content]{0}[Column1],
    GebTag = Date.From(GebTagRaw),
    Tage   = Number.From(TageRaw),
    NextDays = List.Dates(Date.From(DateTime.LocalNow()), Tage, #duration(1,0,0,0)),
    Alter = List.Transform(NextDays, each Duration.Days(_ - GebTag)),
    PI = 3.141592653589793,
    Bio = (periode as number) as list =>
        List.Transform(
            Alter,
            each Number.Round(
                Number.Sin(2 * PI * _ / periode),
                2
            )
        ),
    Physisch      = Bio(23),
    Emotional     = Bio(28),
    Intellektuell = Bio(33),
    Intuitiv      = Bio(38),
    Spirituell    = Bio(53),
    Kreativ       = Bio(41),
    Perioden = [
        physisch      = 23,
        emotional     = 28,
        intellektuell = 33,
        intuitiv      = 38,
        spirituell    = 53,
        kreativ       = 41
    ],
    KritischeKurven =
        let
            txt = Text.Lower(Text.Trim(Text.From(KKurvenRaw)))
        in
            if txt = "alle" then
                Record.FieldNames(Perioden)
            else
                List.Transform(
                    Text.Split(txt, ","),
                    each Text.Trim(_)
                ),
    IstKritisch =
        (alterHeute as number, periode as number) as logical =>
            Number.RoundDown(2 * (alterHeute - 1) / periode)
            <>
            Number.RoundDown(2 * alterHeute / periode),
    Kritisch =
    List.Transform(
        Alter,
        (a) =>
            if
                List.AnyTrue(
                    List.Transform(
                        KritischeKurven,
                        (k) =>
                            IstKritisch(
                                a,
                                Record.Field(Perioden, k)
                            )
                    )
                )
            then
                0
            else
                null
    ),
        KritischeKurve =
        List.Transform(
            Alter,
            (a) =>
                let
                    Treffer =
                        List.Select(
                            KritischeKurven,
                            (k) =>
                                IstKritisch(
                                    a,
                                    Record.Field(Perioden, k)
                                )
                        )
                in
                    if List.IsEmpty(Treffer) then
                        null
                    else
                        Text.Combine(
                            List.Transform(Treffer, each Text.Proper(_)),
                            ", "
                        )
        ),
    Tabelle =
        Table.FromColumns(
            {
                NextDays,
                Physisch,
                Emotional,
                Intellektuell,
                Intuitiv,
                Spirituell,
                Kreativ,
                Kritisch,
                KritischeKurve
            },
            {
                "Datum",
                "Physisch",
                "Emotional",
                "Intellektuell",
                "Intuitiv",
                "Spirituell",
                "Kreativ",
                "Kritisch",
                "Kritische Kurve"
            }
        ),
    Erg =
        Table.TransformColumns(
            Tabelle,
            {
                {
                    "Datum",
                    each Date.ToText(_, "dd.MM.yyyy"),
                    type text
                }
            }
        )
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    GebTagRaw = Excel.CurrentWorkbook(){[Name="GebTagP"]}[Content]{0}[Column1],
    TageRaw   = Excel.CurrentWorkbook(){[Name="TageP"]}[Content]{0}[Column1],
    GebTag = Date.From(GebTagRaw),
    Tage   = Number.From(TageRaw),
    AuswahlRaw    = Excel.CurrentWorkbook(){[Name="tblKurve"]}[Content],
    AuswahlRecord = AuswahlRaw{0},
    AlleSpalten = {
        "Physisch",
        "Emotional",
        "Intellektuell",
        "Intuitiv",
        "Spirituell",
        "Kreativ",
        "Kritisch"
    },
    Angehaakt = List.Select(Record.FieldNames(AuswahlRecord), each Record.HasFields(AuswahlRecord, _) and Record.Field(AuswahlRecord, _) = true),
    Auswahl = if List.IsEmpty(Angehaakt) then AlleSpalten else Angehaakt,
    BioNamen = {
        "Physisch",
        "Emotional",
        "Intellektuell",
        "Intuitiv",
        "Spirituell",
        "Kreativ"
    },
    AusgewaehlteBioKurven = List.Intersect({Auswahl, BioNamen}),
    KritischeKurven = if List.IsEmpty(AusgewaehlteBioKurven) then BioNamen else AusgewaehlteBioKurven,
    NextDays = List.Dates(Date.From(DateTime.LocalNow()), Tage, #duration(1, 0, 0, 0)),
    Alter = List.Transform(NextDays, each Duration.Days(_ - GebTag)),
    PI = 3.141592653589793,
    Bio = (periode as number) as list =>
        List.Transform(
            Alter,
            each Number.Round(
                Number.Sin(2 * PI * _ / periode),
                2
            )
        ),
    Physisch      = Bio(23),
    Emotional     = Bio(28),
    Intellektuell = Bio(33),
    Intuitiv      = Bio(38),
    Spirituell    = Bio(53),
    Kreativ       = Bio(41),
    Perioden = [
        Physisch      = 23,
        Emotional     = 28,
        Intellektuell = 33,
        Intuitiv      = 38,
        Spirituell    = 53,
        Kreativ       = 41
    ],
    IstKritisch =
        (alterHeute as number, periode as number) as logical =>
            Number.RoundDown(
                2 * (alterHeute - 1) / periode
            )
            <>
            Number.RoundDown(
                2 * alterHeute / periode
            ),
    Kritisch =
        List.Transform(
            Alter,
            (a) =>
                if
                    List.AnyTrue(
                        List.Transform(
                            KritischeKurven,
                            (k) =>
                                IstKritisch(
                                    a,
                                    Record.Field(Perioden, k)
                                )
                        )
                    )
                then
                    0
                else
                    null
        ),
    NullListe = List.Repeat({null}, Tage),
    PhysischOut = if List.Contains(Auswahl, "Physisch") then Physisch else NullListe,
    EmotionalOut = if List.Contains(Auswahl, "Emotional") then Emotional else NullListe,
    IntellektuellOut = if List.Contains(Auswahl, "Intellektuell") then Intellektuell else NullListe,
    IntuitivOut = if List.Contains(Auswahl, "Intuitiv") then Intuitiv else NullListe,
    SpirituellOut = if List.Contains(Auswahl, "Spirituell") then Spirituell else NullListe,
    KreativOut = if List.Contains(Auswahl, "Kreativ") then Kreativ else NullListe,
    KritischOut = if List.Contains(Auswahl, "Kritisch") then Kritisch else NullListe,
    Tabelle =
        Table.FromColumns(
            {
                NextDays,
                PhysischOut,
                EmotionalOut,
                IntellektuellOut,
                IntuitivOut,
                SpirituellOut,
                KreativOut,
                KritischOut
            },
            {
                "Datum",
                "Physisch",
                "Emotional",
                "Intellektuell",
                "Intuitiv",
                "Spirituell",
                "Kreativ",
                "Kritisch"
            }
        ),
    Erg =
        Table.TransformColumnTypes(
            Tabelle,
            {
                {"Datum", type date},
                {"Physisch", type number},
                {"Emotional", type number},
                {"Intellektuell", type number},
                {"Intuitiv", type number},
                {"Spirituell", type number},
                {"Kreativ", type number},
                {"Kritisch", type number}
            }
        )
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0020",
                .Category = DemoCategory.PowerQuery,
                .Title = "Freitag der 13te...",
                .Tags = {"pq", "power query", "freitag", "13", "formel", "vba"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Jeder Freitag der 13te wird ausgegeben. Nach Start- und Enddatum. Zusätzlich kann noch der Monat gewählt werden.
Aktualisiert wird über VBA.
PQ_Freitag_der_13te_Startdatum_Enddatum_Monat.xlsb

Auch mit Formeln gelöst (Formeln in O1, Q1 und S1).
=LET(x;SEQUENZ(A2-A1+1;1;A1;1);VSTAPELN("Freitag, der 13te...";FILTER(x;(x>=DATUM(1900;3;1))*(TAG(x)=13)*(WOCHENTAG(x;2)=5))))
=LET(x;SEQUENZ(A2-A1+1;1;A1;1);m;MONAT(DATWERT("1. "&A4));VSTAPELN("Freitag, der 13te von A4";FILTER(x;(x>=DATUM(1900;3;1))*(TAG(x)=13)*(WOCHENTAG(x;2)=5)*(MONAT(x)=m))))
=LET(x;SEQUENZ(Ende-Start+1;1;Start;1);m;MONAT(DATWERT("1. "&Monat));VSTAPELN("Freitag, der 13te ohne A4";FILTER(x;(x>=DATUM(1900;3;1))*(TAG(x)=13)*(WOCHENTAG(x;2)=5)*(MONAT(x)=m))))

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Startdatum = Date.From(Excel.CurrentWorkbook(){[Name="Start"]}[Content]{0}[Column1]),
    Enddatum = Date.From(Excel.CurrentWorkbook(){[Name="Ende"]}[Content]{0}[Column1]),
    StartMonat = Date.StartOfMonth(Startdatum),
    EndeMonat = Date.StartOfMonth(Enddatum),
    AnzahlMonate = (Date.Year(EndeMonat) - Date.Year(StartMonat)) * 12 + Date.Month(EndeMonat) - Date.Month(StartMonat) + 1,
    Monate = List.Transform({0..AnzahlMonate - 1}, each Date.AddMonths(StartMonat, _)),
    Dreizehnte = List.Transform(Monate, each #date(Date.Year(_), Date.Month(_), 13)),
    Freitag13 = List.Select(Dreizehnte, each _ >= Startdatum and _ <= Enddatum and Date.DayOfWeek(_, Day.Monday) = 4),
    Tabelle = Table.FromList(Freitag13, Splitter.SplitByNothing(), {"Datum"}),
    MitWochentag = Table.AddColumn(Tabelle, "Wochentag", each Date.DayOfWeekName([Datum], "de-DE"), type text),
    MitMonat = Table.AddColumn(MitWochentag, "Monat", each Date.MonthName([Datum], "de-DE"), type text),
    MitJahr = Table.AddColumn(MitMonat, "Jahr", each Date.Year([Datum]), Int64.Type)
in
    MitJahr

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Startdatum = Date.From(Excel.CurrentWorkbook(){[Name="Start"]}[Content]{0}[Column1]),
    Enddatum = Date.From(Excel.CurrentWorkbook(){[Name="Ende"]}[Content]{0}[Column1]),
    AuswahlMonat = Text.From(Excel.CurrentWorkbook(){[Name="Monat"]}[Content]{0}[Column1]),
    StartMonat = Date.StartOfMonth(Startdatum),
    EndeMonat = Date.StartOfMonth(Enddatum),
    AnzahlMonate = (Date.Year(EndeMonat) - Date.Year(StartMonat)) * 12 + Date.Month(EndeMonat) - Date.Month(StartMonat) + 1,
    Monate = List.Transform({0..AnzahlMonate - 1}, each Date.AddMonths(StartMonat, _)),
    Dreizehnte = List.Transform(Monate, each #date(Date.Year(_), Date.Month(_), 13)),
    Freitag13 = List.Select(Dreizehnte, each _ >= Startdatum and _ <= Enddatum and Date.DayOfWeek(_, Day.Monday) = 4 and Date.MonthName(_, "de-DE") = AuswahlMonat),
    Tabelle = Table.FromList(Freitag13, Splitter.SplitByNothing(), {"Datum"}),
    MitWochentag = Table.AddColumn(Tabelle, "Wochentag", each Date.DayOfWeekName([Datum], "de-DE"), type text),
    MitMonat = Table.AddColumn(MitWochentag, "Monat", each Date.MonthName([Datum], "de-DE"), type text),
    MitJahr = Table.AddColumn(MitMonat, "Jahr", each Date.Year([Datum]), Int64.Type)
in
    MitJahr
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0021",
                .Category = DemoCategory.PowerQuery,
                .Title = "Langer Text in Zellen aufteilen...",
                .Tags = {"auftrennenl", "text", "aufteilen", "formel", "regex", "vba", "power query"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Lange Texte in Zellen werden nach vorgegebener Anzahl von Zeichen aufgetrennt.
Nicht mitten im Wort.
Verschiedene Möglichkeiten (in Spalten, Zeilen, Blöcken...).
Mit RegEx, Power Query, Formeln und VBA (UDF und Sub).

RegEx_Power_Query_Formel_VBA_Text_auftrennen.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    QuelleN = Excel.CurrentWorkbook(){[Name="AnzZ"]}[Content]{0}[Column1],
    TextTrennen = (TextT as text, LenT as number) as list =>
        let
            Worte = Text.Split(TextT, " "),
            Result = List.Accumulate(Worte, {""}, (Liste, Worte) =>
                let
                    ZeileL = List.Last(Liste),
                    ZeileN = if Text.Length(ZeileL & " " & Worte) <= LenT then Text.Trim(ZeileL & " " & Worte) else Worte,
                    ListeL = if Text.Length(ZeileL & " " & Worte) <= LenT then List.RemoveLastN(Liste,1) & {ZeileN} else Liste & {ZeileN}
                in
                    ListeL)
        in
            Result,
    ListeM = Table.AddColumn(Quelle, "Getrennt_nach_C1", each TextTrennen([Daten], QuelleN)),
    Erg = Table.RemoveColumns(Table.ExpandListColumn(ListeM, "Getrennt_nach_C1"),{"Daten"})
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle3"]}[Content],
    TextTrennen = (TextT as text, LenT as number) as list =>
        let
            Worte = Text.Split(TextT, " "),
            Result = List.Accumulate(Worte, {""}, (Liste, Worte) =>
                let
                    ZeileL = List.Last(Liste),
                    ZeileN = if Text.Length(ZeileL & " " & Worte) <= LenT then Text.Trim(ZeileL & " " & Worte) else Worte,
                    ListeL = if Text.Length(ZeileL & " " & Worte) <= LenT then List.RemoveLastN(Liste,1) & {ZeileN} else Liste & {ZeileN}
                in
                    ListeL)
        in
            Result,
    MitListen = Table.AddColumn(Quelle, "Getrennt_nach_Spalte_B", each TextTrennen([Daten], [Länge])),
    Erg = Table.RemoveColumns(Table.ExpandListColumn(MitListen, "Getrennt_nach_Spalte_B"),{"Daten", "Länge"})
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    QuelleN = Excel.CurrentWorkbook(){[Name="AnzZ"]}[Content]{0}[Column1],
    TextTrennen = (TextT as text, LenT as number) as text =>
        let
            Worte = Text.Split(TextT, " "),
            Result = List.Accumulate(Worte, {""}, (Liste, Worte) =>
                let
                    ZeileL = List.Last(Liste),
                    ZeileN = if Text.Length(ZeileL & " " & Worte) <= LenT then Text.Trim(ZeileL & " " & Worte) else Worte,
                    ListeL = if Text.Length(ZeileL & " " & Worte) <= LenT then List.RemoveLastN(Liste,1) & {ZeileN} else Liste & {ZeileN}
                in
                    ListeL),
            Erg = Text.Combine(Result, "#(lf)")
        in
            Erg,
    Erg1 = Table.RemoveColumns(Table.AddColumn(Quelle, "ZeilenText", each TextTrennen([Daten], QuelleN)),{"Daten"})
in
    Erg1

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    QuelleN = Excel.CurrentWorkbook(){[Name="AnzZ"]}[Content]{0}[Column1],
    TextTrennen = (TextT as text, LenT as number) as list =>
        let
            Worte = Text.Split(Text.Trim(TextT), " "),
            Result =
                List.Accumulate(
                    Worte,
                    {""},
                    (Liste, Wort) =>
                        let
                            ZeileL = List.Last(Liste),
                            Neu = Text.Trim(ZeileL & " " & Wort),
                            Passt = Text.Length(Neu) <= LenT,
                            ListeN =
                                if Passt
                                then List.RemoveLastN(Liste, 1) & {Neu}
                                else Liste & {Wort}
                        in
                            ListeN
                )
        in
            Result,
    MitListen = Table.AddColumn(Quelle, "Liste", each TextTrennen([Daten], QuelleN)),
    MitRecords = Table.AddColumn(MitListen, "Getrennt", each Record.FromList([Liste], List.Transform({1..List.Count([Liste])}, each "Teil_" & Text.From(_)))),
    AlleSpalten = List.Union(List.Transform(MitRecords[Getrennt], each Record.FieldNames(_))),
    Expandiert = Table.ExpandRecordColumn(MitRecords, "Getrennt", AlleSpalten, AlleSpalten),
    Erg = Table.RemoveColumns(Expandiert, {"Daten", "Liste"})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0022",
                .Category = DemoCategory.PowerQuery,
                .Title = "Werte aus Spalten in einer Spalte ausgeben...",
                .Tags = {"auftrennen", "text", "aufteilen", "formel", "m-code", "vba", "power query"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Daten aus Spalten (im Beispiel A:F) werden in eine Spalte zusammengeführt.
Mit Power Query und Formeln gelöst. Aktualisierung der Power Query Abfragen mit VBA.

PQ_Power_Query_Formeln_mehrere_Spalten_in_EINE_Spalte.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    SoderZ = Excel.CurrentWorkbook(){[Name="ZoderS"]}[Content]{0}[Column1],
    Erg = Table.FromList(List.Combine( if SoderZ then Table.ToColumns(Quelle) else Table.ToRows(Quelle)), Splitter.SplitByNothing(), {"Wert"})
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Bereich"]}[Content],
    SoderZ = Excel.CurrentWorkbook(){[Name="SoderZ"]}[Content]{0}[Column1],
    Erg = Table.FromList(List.Combine( if SoderZ then Table.ToColumns(Quelle) else Table.ToRows(Quelle)), Splitter.SplitByNothing(), {"Wert"})
in
    Erg
//Mit List.Accumulate gelöst
/*
let
    Quelle = Excel.CurrentWorkbook(){[Name="Bereich"]}[Content],
    SoderZ = Excel.CurrentWorkbook(){[Name="SoderZ"]}[Content]{0}[Column1],
    Result =
        if SoderZ then
            List.Accumulate(
                {0 .. Table.ColumnCount(Quelle) - 1},
                {},
                (st, cur) => st & Table.Column(Quelle, Table.ColumnNames(Quelle){cur})
            )
        else
            List.Accumulate(
                {0 .. Table.RowCount(Quelle) - 1},
                {},
                (st, cur) => st & Record.ToList(Quelle{cur})
            ),
    Erg = Table.FromList(Result, Splitter.SplitByNothing(), {"Wert"})
in
    Erg
*/

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Bereich_1"]}[Content],
    SoderZ = Excel.CurrentWorkbook(){[Name="SoderZ"]}[Content]{0}[Column1],
    Erg = Table.FromList(List.Combine( if SoderZ then Table.ToColumns(Quelle) else Table.ToRows(Quelle)), Splitter.SplitByNothing(), {"Wert"})
in
    Erg

//Mit List.Accumulate gelöst
/*
let
    Quelle = Excel.CurrentWorkbook(){[Name="Bereich_1"]}[Content],
    SoderZ = Excel.CurrentWorkbook(){[Name="SoderZ"]}[Content]{0}[Column1],
    Result =
        if SoderZ then
            List.Accumulate(
                {0 .. Table.ColumnCount(Quelle) - 1},
                {},
                (st, cur) => st & Table.Column(Quelle, Table.ColumnNames(Quelle){cur})
            )
        else
            List.Accumulate(
                {0 .. Table.RowCount(Quelle) - 1},
                {},
                (st, cur) => st & Record.ToList(Quelle{cur})
            ),
    Erg = Table.FromList(Result, Splitter.SplitByNothing(), {"Wert"})
in
    Erg
*/
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0023",
                .Category = DemoCategory.PowerQuery,
                .Title = "Geburtstagsliste nach Monaten...",
                .Tags = {"geburtstag", "liste", "datum", "formel", "m-code", "vba", "power query", "pivot"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Über die Kontrollkästchen können die gewünschten Monate ausgewählt werden.
Beim Aktualisieren (geht über VBA) liest Power Query deren Status ein,
filtert die Geburtstagsliste und berechnet Wochentage, Alter sowie die Tage bis zum nächsten Geburtstag.
Sind keine Monate angehakt, werden automatisch alle Monate berücksichtigt.
Die Überschrift wird passend zur jeweiligen Monatsauswahl erzeugt.

Power_Query_Pivot_mit_Formel_Geburtstagsliste_Monate.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    Daten = Table.TransformColumnTypes(Quelle, {{"Name", type text}, {"Geburtstag", type date}}),
    Monatsnamen = {"Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"},
    Haken = Record.ToList(Excel.CurrentWorkbook(){[Name="Monate"]}[Content]{0}),
    Auswahl = List.Select({1..12}, each try Haken{_ - 1} = true otherwise false),
    Monate = if List.IsEmpty(Auswahl) then {1..12} else Auswahl,
    Gruppen = List.Accumulate(
        Monate,
        {},
        (s, m) =>
            if List.IsEmpty(s) or m <> List.Last(List.Last(s)) + 1
            then s & {{m}}
            else List.RemoveLastN(s, 1) &
                 {List.Last(s) & {m}}
    ),
    AnzahlGruppen = List.Count(Gruppen),
    Teile = List.Transform(
        Gruppen,
        each
            let
                Von = List.First(_),
                Bis = List.Last(_)
            in
                if Von = Bis then
                    Monatsnamen{Von - 1}
                else
                    (if AnzahlGruppen = 1 or AnzahlGruppen >= 3
                     then "von "
                     else "") &
                    Monatsnamen{Von - 1} & " bis " &
                    Monatsnamen{Bis - 1}
    ),
    Monatstext =
        if AnzahlGruppen = 1 then
            Teile{0}
        else if AnzahlGruppen = 2 then
            Teile{0} & " und " & Teile{1}
        else
            Text.Combine(List.RemoveLastN(Teile, 1), ", ") &
            " und " & List.Last(Teile),
    Überschrift = "Geburtstagsliste " & Monatstext,
    Gefiltert = Table.SelectRows(Daten, each List.Contains(Monate, Date.Month([Geburtstag]))),
    Heute = Date.From(DateTime.LocalNow()),
    Berechnet = Table.AddColumn(
        Gefiltert,
        "Daten",
        each
            let
                Geburtstag = [Geburtstag],
                GeburtstagDiesesJahr =
                    try #date(
                        Date.Year(Heute),
                        Date.Month(Geburtstag),
                        Date.Day(Geburtstag)
                    )
                    otherwise #date(Date.Year(Heute), 2, 28),
                NächsterGeburtstag =
                    if GeburtstagDiesesJahr < Heute
                    then Date.AddYears(GeburtstagDiesesJahr, 1)
                    else GeburtstagDiesesJahr,
                Alter =
                    Date.Year(Heute) - Date.Year(Geburtstag) -
                    (
                        if GeburtstagDiesesJahr > Heute
                        then 1
                        else 0
                    )
            in
                [
                    TagG = Date.DayOfWeekName(Geburtstag, "de-DE"),
                    TageBis = Duration.Days(NächsterGeburtstag - Heute),
                    TagH = Date.DayOfWeekName(NächsterGeburtstag, "de-DE"),
                    Alter = Alter
                ]
    ),
    Erweitert = Table.ExpandRecordColumn(Berechnet, "Daten", {"TagG", "TageBis", "TagH", "Alter"}),
    Sortiert = Table.Sort(Erweitert, {{"Name", Order.Ascending}}),
    /*
    Titelzeile = #table(
        Table.ColumnNames(Sortiert),
        {{Überschrift, null, null, null, null, null}}
    ),
    Erg = Table.Combine({Titelzeile, Sortiert})
    */
    Erg = Table.RenameColumns(Sortiert, {{"Name", Überschrift}})
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
// Die Überschrift der ersten Spalte ist Name. Die ausgewählten Monate werden in der ersten Datenzeile ausgegeben.
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    Daten = Table.TransformColumnTypes(Quelle, {{"Name", type text}, {"Geburtstag", type date}}),
    Monatsnamen = {"Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"},
    Haken = Record.ToList(Excel.CurrentWorkbook(){[Name="Monate"]}[Content]{0}),
    Auswahl = List.Select({1..12}, each try Haken{_ - 1} = true otherwise false),
    Monate = if List.IsEmpty(Auswahl) then {1..12} else Auswahl,
    Gruppen = List.Accumulate(
        Monate,
        {},
        (s, m) =>
            if List.IsEmpty(s) or m <> List.Last(List.Last(s)) + 1
            then s & {{m}}
            else List.RemoveLastN(s, 1) &
                 {List.Last(s) & {m}}
    ),
    AnzahlGruppen = List.Count(Gruppen),
    Teile = List.Transform(
        Gruppen,
        each
            let
                Von = List.First(_),
                Bis = List.Last(_)
            in
                if Von = Bis then
                    Monatsnamen{Von - 1}
                else
                    (if AnzahlGruppen = 1 or AnzahlGruppen >= 3
                     then "von "
                     else "") &
                    Monatsnamen{Von - 1} & " bis " &
                    Monatsnamen{Bis - 1}
    ),
    Monatstext =
        if AnzahlGruppen = 1 then
            Teile{0}
        else if AnzahlGruppen = 2 then
            Teile{0} & " und " & Teile{1}
        else
            Text.Combine(List.RemoveLastN(Teile, 1), ", ") &
            " und " & List.Last(Teile),
    Überschrift = "Geburtstagsliste " & Monatstext,
    Gefiltert = Table.SelectRows(Daten, each List.Contains(Monate, Date.Month([Geburtstag]))),
    Heute = Date.From(DateTime.LocalNow()),
    Berechnet = Table.AddColumn(
        Gefiltert,
        "Daten",
        each
            let
                Geburtstag = [Geburtstag],
                GeburtstagDiesesJahr =
                    try #date(
                        Date.Year(Heute),
                        Date.Month(Geburtstag),
                        Date.Day(Geburtstag)
                    )
                    otherwise #date(Date.Year(Heute), 2, 28),
                NächsterGeburtstag =
                    if GeburtstagDiesesJahr < Heute
                    then Date.AddYears(GeburtstagDiesesJahr, 1)
                    else GeburtstagDiesesJahr,
                Alter =
                    Date.Year(Heute) - Date.Year(Geburtstag) -
                    (
                        if GeburtstagDiesesJahr > Heute
                        then 1
                        else 0
                    )
            in
                [
                    TagG = Date.DayOfWeekName(Geburtstag, "de-DE"),
                    TageBis = Duration.Days(NächsterGeburtstag - Heute),
                    TagH = Date.DayOfWeekName(NächsterGeburtstag, "de-DE"),
                    Alter = Alter
                ]
    ),
    Erweitert = Table.ExpandRecordColumn(Berechnet, "Daten", {"TagG", "TageBis", "TagH", "Alter"}),
    Sortiert = Table.Sort(Erweitert, {{"Name", Order.Ascending}}),
    Titelzeile = #table(Table.ColumnNames(Sortiert), {{Überschrift, null, null, null, null, null}}),
    Erg = Table.Combine({Titelzeile, Sortiert})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0024",
                .Category = DemoCategory.PowerQuery,
                .Title = "Längste Reihe positiver Zahlen...",
                .Tags = {"zahl", "positiv", "reihe", "formel", "let", "längste"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Reihe von Zahlen wird die längste positive Zahlenreihe gesucht.
Ausgegeben werden - welche Zahl, Länge, Start, Ende und Summe.
Auch mehrere gleiche Zahlenreihen sind berücksichtigt.

PQ_Formel_Laengste_Reihe_positive_Zahlen.xlsx

In den Abfragen werden unterschiedliche Herangehensweisen gezeigt.

Auch mit Formeln gelöst.

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    Zahlen = Quelle[Zahlen],
    Gruppen =
        List.Accumulate(
            Zahlen,
            {},
            (s, x) =>
                let
                    Letzte = if List.IsEmpty(s) then null else List.Last(s),
                    Neu =
                        if x > 0 then
                            if Letzte <> null and Letzte[Zahl] = x then
                                List.RemoveLastN(s, 1) &
                                {[
                                    Zahl = x,
                                    Länge = Letzte[Länge] + 1
                                ]}
                            else
                                s & {[Zahl = x, Länge = 1]}
                        else
                            s & {[Zahl = null, Länge = 0]}
                in
                    Neu
        ),
    Positiv = List.Select(Gruppen, each [Länge] > 0),
    MaxLänge = List.Max(List.Transform(Positiv, each [Länge])),
    Treffer = List.First(List.Select(Positiv, each [Länge] = MaxLänge)),
    Erg =
        #table(
            {"Länge", "Summe"},
            {{Treffer[Länge], Treffer[Zahl] * Treffer[Länge]}}
        )
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    Zahlen = Quelle[Zahlen],
    Gruppen =
        List.Accumulate(
            Zahlen,
            {},
            (s, x) =>
                let
                    Letzte = if List.IsEmpty(s) then null else List.Last(s),
                    Neu =
                        if x > 0 then
                            if Letzte <> null and Letzte[Zahl] = x then
                                List.RemoveLastN(s, 1) &
                                {[
                                    Zahl = x,
                                    Länge = Letzte[Länge] + 1
                                ]}
                            else
                                s & {[Zahl = x, Länge = 1]}
                        else
                            s & {[Zahl = null, Länge = 0]}
                in
                    Neu
        ),
    Positiv = List.Select(Gruppen, each [Länge] > 0),
    MaxLänge = List.Max(List.Transform(Positiv, each [Länge])),
    Treffer = List.First(List.Select(Positiv, each [Länge] = MaxLänge)),
    Erg =
        #table(
            {"Zahl", "Länge", "Summe"},
            {{
                Treffer[Zahl],
                Treffer[Länge],
                Treffer[Zahl] * Treffer[Länge]
            }}
        )
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    Zahlen = Quelle[Zahlen],
    Gruppen =
        List.Accumulate(
            {0 .. List.Count(Zahlen) - 1},
            {},
            (s, i) =>
                let
                    x = Zahlen{i},
                    Letzte = if List.IsEmpty(s) then null else List.Last(s),
                    Neu =
                        if x > 0 then
                            if Letzte <> null and Letzte[Zahl] = x then
                                List.RemoveLastN(s, 1) &
                                {[
                                    Zahl = x,
                                    Länge = Letzte[Länge] + 1,
                                    Start = Letzte[Start],
                                    Ende = i + 1
                                ]}
                            else
                                s & {[
                                    Zahl = x,
                                    Länge = 1,
                                    Start = i + 1,
                                    Ende = i + 1
                                ]}
                        else
                            s
                in
                    Neu
        ),
    MaxLänge = List.Max(List.Transform(Gruppen, each [Länge])),
    Treffer = List.First(List.Select(Gruppen, each [Länge] = MaxLänge)),
    Erg =
        #table(
            {"Zahl", "Länge", "Start", "Ende", "Summe"},
            {{
                Treffer[Zahl],
                Treffer[Länge],
                Treffer[Start],
                Treffer[Ende],
                Treffer[Zahl] * Treffer[Länge]
            }}
        )
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    Zahlen = Quelle[Zahlen],
    Gruppen =
        List.Accumulate(
            {0 .. List.Count(Zahlen) - 1},
            {},
            (s, i) =>
                let
                    x = Zahlen{i},
                    Letzte = if List.IsEmpty(s) then null else List.Last(s),
                    Neu =
                        if x > 0 then
                            if Letzte <> null and Letzte[Zahl] = x then
                                List.RemoveLastN(s, 1) &
                                {[
                                    Zahl = x,
                                    Länge = Letzte[Länge] + 1,
                                    Start = Letzte[Start],
                                    Ende = i + 1
                                ]}
                            else
                                s & {[
                                    Zahl = x,
                                    Länge = 1,
                                    Start = i + 1,
                                    Ende = i + 1
                                ]}
                        else
                            s
                in
                    Neu
        ),
    MaxLänge = List.Max(List.Transform(Gruppen, each [Länge])),
    Treffer = List.Select(Gruppen, each [Länge] = MaxLänge),
    //Erg =Table.FromRecords(List.Select(Gruppen, each [Länge] = MaxLänge))
    Erg = Table.FromRecords(List.Transform(List.Select(Gruppen, each [Länge] = MaxLänge), each Record.AddField(_, "Summe", [Zahl] * [Länge])))
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Tabelle1"]}[Content],
    Zahlen = Quelle[Zahlen],
    Gruppen =
        List.Accumulate(
            {0 .. List.Count(Zahlen) - 1},
            {},
            (s, i) =>
                let
                    x = Zahlen{i},
                    Letzte = if List.IsEmpty(s) then null else List.Last(s),
                    Neu =
                        if x > 0 then
                            if Letzte <> null and Letzte[Zahl] = x then
                                List.RemoveLastN(s, 1) &
                                {[
                                    Zahl = x,
                                    Länge = Letzte[Länge] + 1,
                                    Start = Letzte[Start],
                                    Ende = i + 1,
                                    Summe = x * (Letzte[Länge] + 1)
                                    //Felder eines Records können bei der Berechnung weiterer Felder verwendet werden.
                                    //Summe = Zahl * Länge
                                ]}
                            else
                                s & {[
                                    Zahl = x,
                                    Länge = 1,
                                    Start = i + 1,
                                    Ende = i + 1,
                                    Summe = x
                                    //Felder eines Records können bei der Berechnung weiterer Felder verwendet werden.
                                    //Summe = Zahl * Länge
                                ]}
                        else
                            s
                in
                    Neu
        ),
    MaxLänge = List.Max(List.Transform(Gruppen, each [Länge])),
    Treffer = List.Select(Gruppen, each [Länge] = MaxLänge),
    Erg =Table.FromRecords(Treffer)
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0025",
                .Category = DemoCategory.PowerQuery,
                .Title = "Variationen mit Datum...",
                .Tags = {"tag", "datum", "februar", "formel", "monat", "jahr"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Variationen mit Datum. In Anlehnung an die Formellösung.
Erstes Tabellenblatt "Übersicht" mit Erklärungen.

PQ_Variationen_mit_Datum.xlsb

In den Abfragen werden unterschiedliche Herangehensweisen gezeigt.

Auch mit Formeln gelöst.

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
//Man könnte die ISO-KW (oder auch das Jahr) mit einer Funktion laden:

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
(Date as date) as number =>
let
    Donnerstag = Date.AddDays(Date, 3 - Date.DayOfWeek(Date, Day.Monday)),
    KW = Number.IntegerDivide(Date.DayOfYear(Donnerstag) - 1, 7) + 1
in
    KW

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
(Date as date) as record =>
let
    Donnerstag = Date.AddDays(Date, 3 - Date.DayOfWeek(Date, Day.Monday)),
    ISOJahr = Date.Year(Donnerstag),
    ISOKW = Number.IntegerDivide(Date.DayOfYear(Donnerstag) - 1, 7) + 1
in
    [ISOJahr = ISOJahr, ISOKW = ISOKW]
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0026",
                .Category = DemoCategory.PowerQuery,
                .Title = "Drei Runden/Tabellen zusammenführen...",
                .Tags = {"pivot", "gruppieren", "join", "unpivot", "outerjoin", "leftjoin"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Drei Tabellen werden zusammengeführt. Es gibt 5 Möglichkeiten.
Mit Gruppieren und Join.
Erstes Tabellenblatt "Übersicht" mit Erklärungen.

PQ_Gruppe_Join_mit_Uebersicht.xlsx

In den Abfragen werden unterschiedliche Herangehensweisen gezeigt.

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Runde1"]}[Content]
in
    Quelle

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Runde2"]}[Content]
in
    Quelle

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="Runde3"]}[Content]
in
    Quelle

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Table.Combine({Runde1, Runde2, Runde3}),
    Gruppe = Table.Group(Quelle, {"Name"}, {{"Runde1", each List.Sum([Runde1]), type number}, {"Runde2", each List.Sum([Runde2]), type number}, {"Runde3", each List.Sum([Runde3]), type number}}),
    Summe = Table.AddColumn(Gruppe, "Gesamt", each List.Sum({[Runde1], [Runde2], [Runde3]}), type number)
in
    Summe

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Table.Combine({Runde1, Runde2, Runde3}),
    Unpivot = Table.UnpivotOtherColumns(Quelle, {"Name"}, "Attribut", "Wert"),
    Pivot = Table.Pivot(Unpivot, List.Distinct(Unpivot[Attribut]), "Attribut", "Wert", List.Sum),
    Summe = Table.AddColumn(Pivot, "Gesamt", each List.Sum({[Runde1], [Runde2], [Runde3]}), type number)
in
    Summe

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Join12 = Table.NestedJoin(Runde1, {"Name"}, Runde2, {"Name"}, "R2", JoinKind.FullOuter),
    Expand12 = Table.ExpandTableColumn(Join12, "R2", {"Name", "Runde2"}, {"Name2", "Runde2"}),
    Name12 = Table.AddColumn(Expand12, "NameNeu", each List.First(List.RemoveNulls({[Name], [Name2]})), type text),
    Bereinigt12 = Table.RemoveColumns(Name12, {"Name", "Name2"}),
    Umbenannt12 = Table.RenameColumns(Bereinigt12, {{"NameNeu", "Name"}}),
    Join123 = Table.NestedJoin(Umbenannt12, {"Name"}, Runde3, {"Name"}, "R3", JoinKind.FullOuter),
    Expand123 = Table.ExpandTableColumn(Join123, "R3", {"Name", "Runde3"}, {"Name3", "Runde3"}),
    Name123 = Table.AddColumn(Expand123, "NameNeu", each List.First(List.RemoveNulls({[Name], [Name3]})), type text),
    Bereinigt123 = Table.RemoveColumns(Name123, {"Name", "Name3"}),
    Umbenannt123 = Table.RenameColumns(Bereinigt123, {{"NameNeu", "Name"}}),
    NullErsetzen = Table.ReplaceValue(Umbenannt123, null, 0, Replacer.ReplaceValue, {"Runde1", "Runde2", "Runde3"}),
    Summe = Table.AddColumn(NullErsetzen, "Gesamt", each [Runde1] + [Runde2] + [Runde3], type number),
    Erg = Table.ReorderColumns(Summe, {"Name", "Runde1", "Runde2", "Runde3", "Gesamt"})
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Namen = List.Distinct(List.Combine({Runde1[Name], Runde2[Name], Runde3[Name]})),
    Basis = Table.FromList(Namen, Splitter.SplitByNothing(), {"Name"}),
    Join1 = Table.NestedJoin(Basis, {"Name"}, Runde1, {"Name"}, "R1", JoinKind.LeftOuter),
    Expand1 = Table.ExpandTableColumn(Join1, "R1", {"Runde1"}),
    Join2 = Table.NestedJoin(Expand1, {"Name"}, Runde2, {"Name"}, "R2", JoinKind.LeftOuter),
    Expand2 = Table.ExpandTableColumn(Join2, "R2", {"Runde2"}),
    Join3 = Table.NestedJoin(Expand2, {"Name"}, Runde3, {"Name"}, "R3", JoinKind.LeftOuter),
    Expand3 = Table.ExpandTableColumn(Join3, "R3", {"Runde3"}),
    NullErsetzen = Table.ReplaceValue(Expand3, null, 0, Replacer.ReplaceValue, {"Runde1", "Runde2", "Runde3"}),
    Summe = Table.AddColumn(NullErsetzen, "Gesamt", each [Runde1] + [Runde2] + [Runde3], type number)
in
    Summe

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Namen = List.Distinct(List.Combine({Runde1[Name], Runde2[Name], Runde3[Name]})),
    Basis = Table.FromList(Namen, Splitter.SplitByNothing(), {"Name"}),
    Join1 = Table.Join(Basis, "Name", Table.RenameColumns(Runde1, {{"Name", "Name1"}}), "Name1", JoinKind.LeftOuter),
    Entferne1 = Table.RemoveColumns(Join1, {"Name1"}),
    Join2 = Table.Join(Entferne1, "Name", Table.RenameColumns(Runde2, {{"Name", "Name2"}}), "Name2", JoinKind.LeftOuter),
    Entferne2 = Table.RemoveColumns(Join2, {"Name2"}),
    Join3 = Table.Join(Entferne2, "Name", Table.RenameColumns(Runde3, {{"Name", "Name3"}}), "Name3", JoinKind.LeftOuter),
    Entferne3 = Table.RemoveColumns(Join3, {"Name3"}),
    NullErsetzen = Table.ReplaceValue(Entferne3, null, 0, Replacer.ReplaceValue, {"Runde1", "Runde2", "Runde3"}),
    Summe = Table.AddColumn(NullErsetzen, "Gesamt", each [Runde1] + [Runde2] + [Runde3], type number)
in
    Summe
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0027",
                .Category = DemoCategory.PowerQuery,
                .Title = "Zahlenfolgen - Von Fibonacci bis zur eigenen Folge...",
                .Tags = {"fibonacci", "zahlenfolge", "tribonacci", "power query", "startwerte", "gewichtung"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Von Fibonacci und Lucas bis zu Tribonacci und gewichteten Zahlenfolgen:
Diese Demo zeigt, wie sich mathematische Zahlenfolgen mit modernen Excel-Formeln (LET/LAMBDA/MAP/REDUCE),
VBA und Power Query berechnen lassen. Mit frei wählbaren Startwerten, Gewichtungen und interaktiver Auswahl.

Auch in VBA und Formeln gelöst.

Fibonacci_and_More.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
(n as number, Start as list, Gewichte as list) as list =>
let
    m = List.Count(Start),
    Pruefung =
        if n < 1
            or n <> Number.RoundDown(n)
            or m = 0
            or List.Count(Gewichte) <> m
        then
            error "Ungültige Vorgaben"
        else
            null,
    Erg =
        if Pruefung <> null then
            {}
        else if n <= m then
            List.FirstN(Start, n)
        else
            List.Accumulate(
                {m + 1 .. n},
                Start,
                (zustand, i) =>
                    let
                        Letzte =
                            List.Reverse(
                                List.LastN(zustand, m)
                            ),
                        Neu =
                            List.Sum(
                                List.Transform(
                                    List.Positions(Gewichte),
                                    (j) => Letzte{j} * Gewichte{j}
                                )
                            )
                    in
                        zustand & {Neu}
            )
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblVorgaben"]}[Content],
    Folge = Text.From(Quelle{0}[Wert]),
    Anzahl = Number.From(Quelle{1}[Wert]),
    Definitionen = [
        #"Fibonacci 0" = [Start={0,1}, Gewicht={1,1}],
        #"Fibonacci 1" = [Start={1,1}, Gewicht={1,1}],
        Lucas = [Start={2,1}, Gewicht={1,1}],
        #"Tribonacci 0" = [Start={0,0,1}, Gewicht={1,1,1}],
        #"Tribonacci 1" = [Start={0,1,1}, Gewicht={1,1,1}],
        Tetranacci = [Start={0,0,0,1}, Gewicht={1,1,1,1}],
        Pell = [Start={0,1}, Gewicht={2,1}],
        Jacobsthal = [Start={0,1}, Gewicht={1,2}]
    ],
    Parameter =
        if Folge = "Eigene Zahlenfolge" then
            let
                AnzahlStart =
                    Number.From(
                        Excel.CurrentWorkbook(){
                            [Name="tblStartwerte"]
                        }[Content]{0}[Anzahl Startwerte]
                    ),
                EigeneTabelle =
                    Excel.CurrentWorkbook(){
                        [Name="tblZahlenfolge"]
                    }[Content],
                PruefungAnzahl =
                    if AnzahlStart < 2
                        or AnzahlStart > 5
                        or AnzahlStart <> Number.RoundDown(AnzahlStart)
                        or Table.RowCount(EigeneTabelle) < AnzahlStart
                    then
                        error "Ungültige Anzahl der Startwerte."
                    else
                        Table.FirstN(EigeneTabelle, AnzahlStart),
                Start =
                    List.Transform(
                        PruefungAnzahl[Start],
                        each Number.From(_)
                    ),
                Gewicht =
                    List.Transform(
                        PruefungAnzahl[Gewicht],
                        each Number.From(_)
                    ),
                Result =
                    if List.Contains(Start, null)
                        or List.Contains(Gewicht, null)
                    then
                        error "Startwerte und Gewichtungen dürfen nicht leer sein."
                    else
                        [Start=Start, Gewicht=Gewicht]
            in
                Result
        else
            Record.Field(Definitionen, Folge),
    Werte = fncZahlenfolge(Anzahl, Parameter[Start], Parameter[Gewicht]),
    Erg = Table.FromColumns({List.Numbers(1, List.Count(Werte)), Werte}, {"Position", "Wert"})
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0028",
                .Category = DemoCategory.PowerQuery,
                .Title = "CSV-Dateien einlesen - Strukturen vereinheitlichen...",
                .Tags = {"csv", "einlesen", "anpassen", "power query", "bom", "datentypern"},
                .Description = TextBlock(
    <text>
        <![CDATA[
CSV-Dateien aus einem Arbeitsmappen-relativen Ordner einlesen, zusammenführen und unterschiedliche Dateistrukturen vereinheitlichen.
Der Pfad wird automatisch über eine Formel ermittelt.

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    BasisPfad = Excel.CurrentWorkbook(){[Name="FolderPath"]}[Content]{0}[Column1],
    CSVOrdner = BasisPfad & "01_Gleiche_Struktur",
    Quelle = Folder.Files(CSVOrdner),
    CDat = Table.SelectRows(Quelle, each Text.Lower([Extension]) = ".csv"),
    CSVInhalt =
        Table.AddColumn(
            CDat,
            "Daten",
            each
                Table.PromoteHeaders(
                    Csv.Document(
                        [Content],
                        [
                            Delimiter = ";",
                            Encoding = 65001,
                            QuoteStyle = QuoteStyle.Csv
                        ]
                    ),
                    [PromoteAllScalars = true]
                ),
            type table
        ),
    Kombiniert = Table.Combine(CSVInhalt[Daten]),
    Erg =
        Table.TransformColumnTypes(
            Kombiniert,
            {
                {"BestellNr", Int64.Type},
                {"Bestelldatum", type date},
                {"Kunde", type text},
                {"Artikel", type text},
                {"Menge", Int64.Type},
                {"Einzelpreis", type number},
                {"Region", type text}
            },
            "de-DE"
        )
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    BasisPfad = Excel.CurrentWorkbook(){[Name="FolderPath"]}[Content]{0}[Column1],
    CSVOrdner = BasisPfad & "01_Gleiche_Struktur",
    Quelle = Folder.Files(CSVOrdner),
    CDat = Table.SelectRows(Quelle, each Text.Lower([Extension]) = ".csv"),
    CSVInhalt =
    Table.AddColumn(
        CDat,
        "Daten",
        each
            let
                DateiName = [Name],
                CSV =
                    Table.PromoteHeaders(
                        Csv.Document(
                            [Content],
                            [
                                Delimiter = ";",
                                Encoding = 65001,
                                QuoteStyle = QuoteStyle.Csv
                            ]
                        ),
                        [PromoteAllScalars = true]
                    ),
                MitDateiname =
                    Table.AddColumn(
                        CSV,
                        "Quelldatei",
                        each DateiName,
                        type text
                    )
            in
                MitDateiname,
        type table
    ),
    Kombiniert = Table.Combine(CSVInhalt[Daten]),
    Erg =
        Table.TransformColumnTypes(
            Kombiniert,
            {
                {"BestellNr", Int64.Type},
                {"Bestelldatum", type date},
                {"Kunde", type text},
                {"Artikel", type text},
                {"Menge", Int64.Type},
                {"Einzelpreis", type number},
                {"Region", type text}
            },
            "de-DE"
        )
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    BasisPfad = Excel.CurrentWorkbook(){[Name="FolderPath"]}[Content]{0}[Column1],
    CSVOrdner = BasisPfad & "02_Unterschiedliche_Struktur",
    Quelle = Folder.Files(CSVOrdner),
    CDat = Table.SelectRows(Quelle, each Text.Lower([Extension]) = ".csv" and [Folder Path] = CSVOrdner & "\"),
    Spalten = {
        "BestellNr",
        "Bestelldatum",
        "Kunde",
        "Artikel",
        "Menge",
        "Einzelpreis",
        "Region",
        "Vertriebskanal"
    },
    CSVInhalt =
        Table.AddColumn(
            CDat,
            "Daten",
            each
                let
                    DateiName = [Name],
                    IstKomma = Text.StartsWith(DateiName, "05_"),
                    IstANSI = Text.StartsWith(DateiName, "06_"),
                    Trennzeichen = if IstKomma then "," else ";",
                    Codierung = if IstANSI then 1252 else 65001,
                    Kultur = if IstKomma then "en-US" else "de-DE",
                    CSV =
                        Table.PromoteHeaders(
                            Csv.Document(
                                [Content],
                                [
                                    Delimiter = Trennzeichen,
                                    Encoding = Codierung,
                                    QuoteStyle = QuoteStyle.Csv
                                ]
                            ),
                            [PromoteAllScalars = true]
                        ),
                    Normalisiert = Table.SelectColumns(CSV, Spalten, MissingField.UseNull),
                    Typisiert =
                        Table.TransformColumnTypes(
                            Normalisiert,
                            {
                                {"BestellNr", Int64.Type},
                                {"Bestelldatum", type date},
                                {"Kunde", type text},
                                {"Artikel", type text},
                                {"Menge", Int64.Type},
                                {"Einzelpreis", type number},
                                {"Region", type text},
                                {"Vertriebskanal", type text}
                            },
                            Kultur
                        ),
                    MitDateiname =
                        Table.AddColumn(
                            Typisiert,
                            "Quelldatei",
                            each DateiName,
                            type text
                        )
                in
                    MitDateiname,
            type table
        ),
    Erg = Table.Combine(CSVInhalt[Daten])
in
    Erg

/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    BasisPfad = Excel.CurrentWorkbook(){[Name="FolderPath"]}[Content]{0}[Column1],
    CSVOrdner = BasisPfad & "02_Unterschiedliche_Struktur",
    Quelle = Folder.Files(CSVOrdner),
    CDat = Table.SelectRows(Quelle, each Text.Lower([Extension]) = ".csv" and [Folder Path] = CSVOrdner & "\"),
    Spalten = {
        "BestellNr",
        "Bestelldatum",
        "Kunde",
        "Artikel",
        "Menge",
        "Einzelpreis",
        "Region",
        "Vertriebskanal"
    },
    PflichtSpalten = {
        "BestellNr",
        "Bestelldatum",
        "Kunde",
        "Artikel",
        "Menge",
        "Einzelpreis"
    },
    fnEncoding = (Inhalt as binary) as number =>
    let
        TextUTF8 = Text.FromBinary(Inhalt, TextEncoding.Utf8),
        BinaryUTF8 = Text.ToBinary(TextUTF8, TextEncoding.Utf8, false),
        OhneBOM =
            if Binary.Length(Inhalt) >= 3
                and Binary.ToList(Binary.Range(Inhalt, 0, 3))
                    = {239, 187, 191}
            then
                Binary.Range(Inhalt, 3)
            else
        Inhalt,
        IstUTF8 = Binary.ToList(OhneBOM) = Binary.ToList(BinaryUTF8),
        Result = if IstUTF8 then 65001 else 1252
    in
        Result,
    fnDelimiter = (Inhalt as binary, pEncoding as number) as text =>
        let
            Kandidaten = {";", ",", "#(tab)", "|"},
            Bewertung =
                List.Transform(
                    Kandidaten,
                    (Trenner) =>
                        let
                            Vorschau =
                                Csv.Document(
                                    Inhalt,
                                    [
                                        Delimiter = Trenner,
                                        Encoding = pEncoding,
                                        QuoteStyle = QuoteStyle.Csv
                                    ]
                                ),
                            Kopf =
                                if Table.RowCount(Vorschau) > 0 then
                                    List.Transform(
                                        Record.FieldValues(Vorschau{0}),
                                        each Text.Trim(Text.From(_))
                                    )
                                else
                                    {},
                            Treffer =
                                List.Count(
                                    List.Intersect({
                                        Kopf,
                                        PflichtSpalten
                                    })
                                )
                        in
                            [
                                Delimiter = Trenner,
                                Treffer = Treffer
                            ]
                ),
            Sortiert =
                List.Sort(
                    Bewertung,
                    (a, b) =>
                        Value.Compare(
                            b[Treffer],
                            a[Treffer]
                        )
                ),
            Ergebnis =
                if Sortiert{0}[Treffer] = 0 then
                    error "Trennzeichen konnte nicht erkannt werden."
                else
                    Sortiert{0}[Delimiter]
        in
            Ergebnis,
    CSVInhalt =
        Table.AddColumn(
            CDat,
            "Daten",
            each
                let
                    DateiName = [Name],
                    Inhalt = [Content],
                    DateiEncoding = fnEncoding(Inhalt),
                    DateiDelimiter = fnDelimiter(Inhalt, DateiEncoding),
                    CSV =
                        Table.PromoteHeaders(
                            Csv.Document(
                                Inhalt,
                                [
                                    Delimiter = DateiDelimiter,
                                    Encoding = DateiEncoding,
                                    QuoteStyle = QuoteStyle.Csv
                                ]
                            ),
                            [PromoteAllScalars = true]
                        ),
                    Normalisiert =
                        Table.SelectColumns(
                            CSV,
                            Spalten,
                            MissingField.UseNull
                        ),
                    Typisiert =
                        Table.TransformColumns(
                            Normalisiert,
                            {
                                {
                                    "BestellNr",
                                    each Int64.From(_),
                                    Int64.Type
                                },
                                {
                                    "Bestelldatum",
                                    each
                                        try Date.FromText(
                                            _,
                                            [Culture = "de-DE"]
                                        )
                                        otherwise
                                            Date.FromText(
                                                _,
                                                [Culture = "en-US"]
                                            ),
                                    type date
                                },
                                {
                                    "Menge",
                                    each Int64.From(_),
                                    Int64.Type
                                },
                                {
                                    "Einzelpreis",
                                    each
                                        if _ = null then
                                            null
                                        else if
                                            Text.Contains(
                                                Text.From(_), ","
                                            )
                                        then
                                            Number.FromText(
                                                Text.From(_),
                                                "de-DE"
                                            )
                                        else
                                            Number.FromText(
                                                Text.From(_),
                                                "en-US"
                                            ),
                                    type number
                                }
                            }
                        ),
                    MitDateiname =
                        Table.AddColumn(
                            Typisiert,
                            "Quelldatei",
                            each DateiName,
                            type text
                        ),
                    MitDelimiter =
                        Table.AddColumn(
                            MitDateiname,
                            "Trennzeichen",
                            each
                                if DateiDelimiter = "#(tab)" then
                                    "TAB"
                                else
                                    DateiDelimiter,
                            type text
                        ),
                    MitEncoding =
                        Table.AddColumn(
                            MitDelimiter,
                            "Encoding",
                            each DateiEncoding,
                            Int64.Type
                        )
                in
                    MitEncoding,
            type table
        ),
    Erg = Table.Combine(CSVInhalt[Daten])
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0029",
                .Category = DemoCategory.PowerQuery,
                .Title = "PDF-Dateien einlesen - Strukturen erkennen...",
                .Tags = {"pdf", "einlesen", "anpassen", "power query", "tabelle", "datentypern"},
                .Description = TextBlock(
    <text>
        <![CDATA[
PDF-Dateien aus einem Arbeitsmappen-relativen Ordner einlesen.
Sechs praxisnahe Beispiele - vollständig fiktive Daten - Möglichkeiten und Grenzen des PDF-Connectors.
Der Pfad wird automatisch über eine Formel ermittelt.

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
Es sind 17 Abfragen.
M-Code bitte in Beispieldatei anschauen.
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0030",
                .Category = DemoCategory.PowerQuery,
                .Title = "Variable Amerikanische Datumsangaben umwandeln...",
                .Tags = {"datum", "amerikanisch", "umwandeln", "power query", "regex", "daten"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Amerikanische Datumsangaben mit numerischem Monat oder englischer Monatsabkürzung werden in deutsche Datums- und Zeitangaben umgewandelt.
Optional entsteht ein echter Excel-Datumswert statt einer Textdarstellung.

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
// Abfrage tblErg
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblUSDatum"]}[Content],
    Monate = {
        "Jan", "Feb", "Mar", "Apr",
        "May", "Jun", "Jul", "Aug",
        "Sep", "Oct", "Nov", "Dec"
    },
    fnUSDatum = (Eingabe as nullable text) as nullable datetime =>
        let
            Text = if Eingabe = null then "" else Text.Trim(Eingabe),
            Teile = List.Select(
                Text.SplitAny(Text, " /-:"),
                each _ <> ""
            ),
            MonatsText = Teile{0},
            MonatsIndex = List.PositionOf(
                List.Transform(Monate, Text.Upper),
                Text.Upper(MonatsText)
            ),
            Monat =
                if MonatsIndex >= 0
                then MonatsIndex + 1
                else Number.FromText(MonatsText),
            Tag = Number.FromText(Teile{1}),
            JahrRoh = Number.FromText(Teile{2}),
            Jahr = if JahrRoh < 100 then 2000 + JahrRoh else JahrRoh,
            HatZeit = List.Count(Teile) > 3,
            StundeRoh = if HatZeit then Number.FromText(Teile{3}) else 0,
            Minute = if HatZeit then Number.FromText(Text.Select(Teile{4}, {"0".."9"})) else 0,
            IstPM = Text.EndsWith(Text.Upper(Text), "PM"),
            IstAM = Text.EndsWith(Text.Upper(Text), "AM"),
            Stunde =
                if IstPM then Number.Mod(StundeRoh, 12) + 12
                else if IstAM then Number.Mod(StundeRoh, 12)
                else StundeRoh,
            Ergebnis = #datetime(Jahr, Monat, Tag, Stunde, Minute, 0)
        in
            if Text = "" then null else Ergebnis,
    MitDatum = Table.AddColumn(Quelle, "Deutsches Datum", each fnUSDatum(Text.From([Daten])), type datetime),
    Erg = Table.AddColumn(
        MitDatum,
        "Formatiert",
        each
            DateTime.ToText(
                [Deutsches Datum],
                if Text.Contains([Daten], ":")
                then "dd.MM.yyyy HH:mm"
                else "dd.MM.yyyy",
                "de-DE"
            ),
        type text
    )
in
    Erg

// Funktion fncUSDatum
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
(Eingabe as nullable text) as nullable datetime =>
let
    t = if Eingabe = null then "" else Text.Trim(Eingabe),
    Monate = {
        "Jan", "Feb", "Mar", "Apr",
        "May", "Jun", "Jul", "Aug",
        "Sep", "Oct", "Nov", "Dec"
    },
    Teile = List.Select(Text.SplitAny(t, " /-:"), each _ <> ""),
    MonatsText = Teile{0},
    MonatsIndex = List.PositionOf(List.Transform(Monate, Text.Upper), Text.Upper(MonatsText)),
    Monat = if MonatsIndex >= 0 then MonatsIndex + 1 else Number.FromText(MonatsText),
    Tag = Number.FromText(Teile{1}),
    JahrRoh = Number.FromText(Teile{2}),
    Jahr = if JahrRoh < 100 then 2000 + JahrRoh else JahrRoh,
    HatZeit = List.Count(Teile) > 3,
    StundeRoh = if HatZeit then Number.FromText(Teile{3}) else 0,
    Minute = if HatZeit then Number.FromText(Text.Select(Teile{4}, {"0".."9"})) else 0,
    IstPM = Text.EndsWith(Text.Upper(t), "PM"),
    IstAM = Text.EndsWith(Text.Upper(t), "AM"),
    Stunde = if IstPM then Number.Mod(StundeRoh, 12) + 12 else if IstAM then Number.Mod(StundeRoh, 12) else StundeRoh,
    Erg = if t = "" then null else #datetime(Jahr, Monat, Tag, Stunde, Minute, 0)
in
    Erg

// Abfrage US_DE_Ergebnis basierend auf der Funktion fncUSDatum
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblUSDatum"]}[Content],
    Datentyp = Table.TransformColumnTypes(Quelle, {{"Daten", type text}}),
    MitDatum = Table.AddColumn(Datentyp, "Deutsches Datum", each fncUSDatum([Daten]), type datetime),
    Erg =
        Table.AddColumn(
            MitDatum,
            "Formatiert",
            each
                if [Deutsches Datum] = null then null
                else
                    DateTime.ToText(
                        [Deutsches Datum],
                        if Text.Contains([Daten], ":")
                        then "dd.MM.yyyy HH:mm"
                        else "dd.MM.yyyy",
                        "de-DE"
                    ),
            type text
        )
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0031",
                .Category = DemoCategory.PowerQuery,
                .Title = "Kalender - lang und quer - erstellen...",
                .Tags = {"kalender", "Power query", "table.fromlist", "nestedjoin", "bundesland", "varianten"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Jahreskalender für alle 16 Bundesländer mit gesetzlichen Feiertagen, ISO-Kalenderwochen, Schaltjahren und zwei Ansichten.
Gelöst mit Formeln, VBA und Power Query.

Kalender_Formel_VBA_Power_Query.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
// PQ_Vorgaben
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(),
    Wert = (Name as text) as any => Quelle{[Name = Name]}[Content]{0}[Column1],
    Jahr = Int64.From(Wert("pJahr")),
    Bundesland = Text.From(Wert("pBundesland")),
    KW = Logical.From(Wert("pKW")),
    Feiertage = Logical.From(Wert("pFeiertage")),
    Heute = Logical.From(Wert("pHeute")),
    Erg = [Jahr = Jahr, Bundesland = Bundesland, KW = KW, Feiertage = Feiertage, Heute = Heute]
in
    Erg

// PQ_Feiertage
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook() {[Name="pFeiertagsmatrix"]}[Content],
    Erg = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true])
in
    Erg

// PQ_Feiertage_Berechnet
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Jahr = PQ_Vorgaben[Jahr],
    Bundesland = PQ_Vorgaben[Bundesland],
    Matrix = PQ_Feiertage,
    LandesSpalten = List.RemoveItems(Table.ColumnNames(Matrix), {"Feiertag"}),
    a = Number.Mod(Jahr, 19),
    b = Number.IntegerDivide(Jahr, 100),
    c = Number.Mod(Jahr, 100),
    d = Number.IntegerDivide(b, 4),
    e = Number.Mod(b, 4),
    f = Number.IntegerDivide(b + 8, 25),
    g = Number.IntegerDivide(b - f + 1, 3),
    h = Number.Mod(19*a + b - d - g + 15, 30),
    i = Number.IntegerDivide(c, 4),
    k = Number.Mod(c, 4),
    l = Number.Mod(32 + 2*e + 2*i - h - k, 7),
    m = Number.IntegerDivide(a + 11*h + 22*l, 451),
    OsterMonat = Number.IntegerDivide(h + l - 7*m + 114, 31),
    OsterTag = Number.Mod(h + l - 7*m + 114, 31) + 1,
    Ostern = #date(Jahr, OsterMonat, OsterTag),
    November22 = #date(Jahr, 11, 22),
    BussUndBettag = Date.AddDays(November22, -Number.Mod(Date.DayOfWeek(November22, Day.Monday) - 2 + 7, 7)),
    Definitionen = [
        #"Neujahr" = [Datum = #date(Jahr, 1, 1), Art = "Fest"],
        #"Heilige Drei Könige" = [Datum = #date(Jahr, 1, 6), Art = "Fest"],
        #"Internationaler Frauentag" = [Datum = #date(Jahr, 3, 8), Art = "Fest"],
        #"Karfreitag" = [Datum = Date.AddDays(Ostern, -2), Art = "Beweglich"],
        #"Ostersonntag" = [Datum = Ostern, Art = "Beweglich"],
        #"Ostermontag" = [Datum = Date.AddDays(Ostern, 1), Art = "Beweglich"],
        #"Tag der Arbeit" = [Datum = #date(Jahr, 5, 1), Art = "Fest"],
        #"Christi Himmelfahrt" = [Datum = Date.AddDays(Ostern, 39), Art = "Beweglich"],
        #"Pfingstsonntag" = [Datum = Date.AddDays(Ostern, 49), Art = "Beweglich"],
        #"Pfingstmontag" = [Datum = Date.AddDays(Ostern, 50), Art = "Beweglich"],
        #"Fronleichnam" = [Datum = Date.AddDays(Ostern, 60), Art = "Beweglich"],
        #"Augsburger Friedensfest" = [Datum = #date(Jahr, 8, 8), Art = "Fest"],
        #"Mariä Himmelfahrt" = [Datum = #date(Jahr, 8, 15), Art = "Fest"],
        #"Weltkindertag" = [Datum = #date(Jahr, 9, 20), Art = "Fest"],
        #"Tag der Deutschen Einheit" = [Datum = #date(Jahr, 10, 3), Art = "Fest"],
        #"Reformationstag" = [Datum = #date(Jahr, 10, 31), Art = "Fest"],
        #"Allerheiligen" = [Datum = #date(Jahr, 11, 1), Art = "Fest"],
        #"Buß- und Bettag" = [Datum = BussUndBettag, Art = "Beweglich"],
        #"1. Weihnachtstag" = [Datum = #date(Jahr, 12, 25), Art = "Fest"],
        #"2. Weihnachtstag" = [Datum = #date(Jahr, 12, 26), Art = "Fest"]
    ],
    MitGueltigkeit = Table.AddColumn(
        Matrix,
        "Gültigkeit",
        each
            let
                Werte = Record.ToList(
                    Record.SelectFields(_, LandesSpalten)
                ),
                Anzahl = List.Sum(
                    List.Transform(Werte, each Number.From(_))
                )
            in
                if Anzahl = List.Count(LandesSpalten) then
                    "Bundesweit"
                else if Anzahl = 0 then
                    "Regionaler Sonderfall"
                else
                    "Regional",
        type text
    ),
    Auswahl = Table.SelectColumns(MitGueltigkeit, {"Feiertag", Bundesland, "Gültigkeit"}),
    MitDefinition = Table.AddColumn(
        Auswahl,
        "Definition",
        each
            let
                Name = [Feiertag]
            in
                if Name = null then
                    error Error.Record(
                        "Feiertagszuordnung",
                        "In der Matrix fehlt ein Feiertagsname.",
                        null
                    )
                else if Record.HasFields(Definitionen, Name) then
                    Record.Field(Definitionen, Name)
                else
                    error Error.Record(
                        "Feiertagszuordnung",
                        "Für diesen Feiertag fehlt eine Berechnungsregel.",
                        [Feiertag = Name]
                    ),
        type record
    ),
    MitDatumUndArt = Table.ExpandRecordColumn(MitDefinition, "Definition", {"Datum", "Art"}, {"Datum", "Art"}),
    TypDatum = Table.TransformColumnTypes(MitDatumUndArt, {{"Datum", type date}, {"Art", type text}}),
    MitWochentag = Table.AddColumn(TypDatum, "Wochentag", each Date.DayOfWeekName([Datum], "de-DE"), type text),
    Umbenannt = Table.RenameColumns(MitWochentag, {{Bundesland, "Gesetzlich"}}),
    Spalten = Table.SelectColumns(
        Umbenannt,
        {
            "Feiertag",
            "Datum",
            "Wochentag",
            "Art",
            "Gültigkeit",
            "Gesetzlich"
        }
    ),
    Typen = Table.TransformColumnTypes(Spalten, {{"Gesetzlich", Int64.Type}}),
    Erg = Table.Sort(Typen, {{"Datum", Order.Ascending}})
in
    Erg

// PQ_Kalender_Lang
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Jahr = PQ_Vorgaben[Jahr],
    MitKW = PQ_Vorgaben[KW],
    Jahresbeginn = #date(Jahr, 1, 1),
    Jahresende = #date(Jahr + 1, 1, 1),
    AnzahlTage = Duration.Days(Jahresende - Jahresbeginn),
    DatumsListe = List.Dates(Jahresbeginn, AnzahlTage, #duration(1, 0, 0, 0)),
    Kalender = Table.FromList(DatumsListe, Splitter.SplitByNothing(), {"Datum"}),
    TypDatum = Table.TransformColumnTypes(Kalender, {{"Datum", type date}}),
    MitWochentag = Table.AddColumn(TypDatum, "Wochentag", each Date.DayOfWeekName([Datum], "de-DE"), type text),
    ISO_KW = (Datum as date) as number =>
        let
            Donnerstag = Date.AddDays(
                Datum,
                3 - Date.DayOfWeek(Datum, Day.Monday)
            ),
            ISOJahr = Date.Year(Donnerstag),
            VierterJanuar = #date(ISOJahr, 1, 4),
            ErsterMontag = Date.StartOfWeek(
                VierterJanuar, Day.Monday
            ),
            KW = Number.IntegerDivide(
                Duration.Days(Donnerstag - ErsterMontag),
                7
            ) + 1
        in
            KW,
    MitKWSpalte = Table.AddColumn(MitWochentag, "KW", each if MitKW then ISO_KW([Datum]) else null, Int64.Type),
    GesetzlicheFeiertage = Table.SelectRows(PQ_Feiertage_Berechnet, each [Gesetzlich] = 1),
    FeiertagsAuswahl = Table.SelectColumns(GesetzlicheFeiertage, {"Datum", "Feiertag"}),
    Verbunden = Table.NestedJoin(MitKWSpalte, {"Datum"}, FeiertagsAuswahl, {"Datum"}, "FT", JoinKind.LeftOuter),
    MitFeiertag = Table.ExpandTableColumn(Verbunden, "FT", {"Feiertag"}, {"Feiertag"}),
    MitTagestyp = Table.AddColumn(
        MitFeiertag,
        "Tagestyp",
        each
            if [Feiertag] <> null then "Feiertag"
            else if Date.DayOfWeek([Datum], Day.Monday) = 5
                then "Samstag"
            else if Date.DayOfWeek([Datum], Day.Monday) = 6
                then "Sonntag"
            else "Werktag",
        type text
    ),
    Erg = Table.ReorderColumns(MitTagestyp, {"Datum", "Wochentag", "KW", "Feiertag", "Tagestyp"})
in
    Erg

// PQ_Kalender_Quer
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = PQ_Kalender_Lang,
    Auswahl = Table.SelectColumns(Quelle, {"Datum", "Wochentag", "KW", "Feiertag", "Tagestyp"}),
    DatumAlsText = Table.TransformColumns(Auswahl, {{"Datum", each Date.ToText(_, "dd.MM.yyyy", "de-DE"), type text}}),
    //OhneNull = Table.ReplaceValue(DatumAlsText, null, "", Replacer.ReplaceValue, {"KW", "Feiertag"}),
    AlsText = Table.TransformColumns(DatumAlsText, {{"KW", each Text.From(_, "de-DE"), type text}}),
    Gedreht = Table.Transpose(AlsText),
    Bezeichnungen = {"Datum", "Wochentag", "KW", "Feiertag", "Tagestyp"},
    MitIndex = Table.AddIndexColumn(Gedreht, "Index", 0, 1, Int64.Type),
    MitBezeichnung = Table.AddColumn(MitIndex, "Bezeichnung", each Bezeichnungen{[Index]}, type text),
    OhneIndex = Table.RemoveColumns(MitBezeichnung, {"Index"}),
    Erg = Table.ReorderColumns(OhneIndex, {"Bezeichnung"} & List.RemoveItems(Table.ColumnNames(OhneIndex), {"Bezeichnung"}))
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0032",
                .Category = DemoCategory.PowerQuery,
                .Title = "Freitext - PNR-Nummern auslesen...",
                .Tags = {"text", "power query", "let", "sequenz", "regex", "varianten"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus Freitext sollen unterschiedlich viele PNR-Nummern in dynamische Spalten PNR1, PNR2, … ausgegeben werden.
Gelöst mit Formeln, Regex, VBA und Power Query.

Demo_Regex_Formel_PQ_VBA.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
// tblErg_1
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblDaten"]}[Content],
    Daten = Table.SelectRows(Quelle, each [Daten] <> null and Text.Trim([Daten]) <> ""),
    fnPNR = (txt as text) as list =>
        let
            Teile = Text.SplitAny(txt, " ,;"),
            Treffer = List.Select(Teile, each
                let
                    p = Text.Upper(Text.Trim(_)),
                    n = if Text.StartsWith(p, "PNRNUMMER_")
                        then Text.Range(p, 10)
                        else "",
                    gueltig = n <> "" and
                        Text.Select(n, {"0".."9"}) = n
                in
                    gueltig
            )
        in
            Treffer,
    Listen = List.Transform(Daten[Daten], each fnPNR(_)),
    Breite = List.Max({0} & List.Transform(Listen, List.Count)),
    Spalten = if Breite = 0 then {} else List.Transform({1..Breite}, each "PNR" & Text.From(_)),
    Zeilen = List.Transform(Listen, each _ & List.Repeat({null}, Breite - List.Count(_))),
    Erg = Table.FromRows(Zeilen, Spalten)
in
    Erg

// tblErg_2
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblDaten"]}[Content],
    Daten = Table.SelectRows(Quelle, each [Daten] <> null and Text.Trim([Daten]) <> ""),
    Listen = List.Transform(Daten[Daten], each
        let
            Teile = Text.Split(_, ","),
            Werte = if List.Count(Teile) > 1
                then List.RemoveLastN(Teile, 1)
                else {},
            Bereinigt = List.Transform(Werte, Text.Trim)
        in
            Bereinigt
    ),
    Breite = List.Max({0} & List.Transform(Listen, List.Count)),
    Spalten = if Breite = 0 then {} else List.Transform({1..Breite}, each "PNR" & Text.From(_)),
    Zeilen = List.Transform(Listen, each _ & List.Repeat({null}, Breite - List.Count(_))),
    Erg = Table.FromRows(Zeilen, Spalten)
in
    Erg

// tblErg_3
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblDaten"]}[Content],
    Daten = Table.SelectRows(Quelle, each [Daten] <> null and Text.Trim([Daten]) <> ""),
    fnGueltig = (wert as text) as logical =>
        let
            t = Text.Upper(Text.Trim(wert)),
            n = if Text.StartsWith(t, "PNRNUMMER_")
                then Text.Range(t, 10)
                else "",
            ok = n <> "" and Text.Select(n, {"0".."9"}) = n
        in
            ok,
    Listen = List.Transform(Daten[Daten], each
        List.FirstN(
            List.Transform(Text.Split(_, ","), Text.Trim),
            each fnGueltig(_)
        )
    ),
    Breite = List.Max({0} & List.Transform(Listen, List.Count)),
    Spalten = if Breite = 0 then {} else List.Transform({1..Breite}, each "PNR" & Text.From(_)),
    Zeilen = List.Transform(Listen, each _ & List.Repeat({null}, Breite - List.Count(_))),
    Erg = Table.FromRows(Zeilen, Spalten)
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0033",
                .Category = DemoCategory.PowerQuery,
                .Title = "Berechnungen Moderne, Historische und Antike Daten...",
                .Tags = {"datum", "power query", "let", "historisch", "modern", "antik"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Drei getrennte PQ-Berechnungen für vollendete Lebensjahre. Moderne Daten - Historische Daten - Antike Daten - Hinweise, Fallen und Besonderheiten.

PowerQuery_Alter_Modern_Antik_Historisch_Demo.xlsx

Als "Schmankerl" für das Tabellenblatt "Moderne_Daten" auch mit Formel in z. B. M1:
=LET(
    G;C2:C37;
    T;D2:D37;
    Heute;HEUTE();
    Datum400;LAMBDA(d;
        WENNFEHLER(
            EDATUM(d;4800);
            WENNFEHLER(
                DATWERT(LINKS(d;6)&(RECHTS(d;4)+400));
                ""
            )
        )
    );
    Ergebnis;MAP(G;T;LAMBDA(g;t;
        LET(
            LeerG;g="";
            LeerT;t="";
            Geburt;WENN(LeerG;"";Datum400(g));
            Tod;WENN(LeerT;"";Datum400(t));
            Heute400;Datum400(Heute);
            WENN(UND(LeerG;LeerT);"Beide Datumsfelder fehlen";
                WENN(LeerG;"Geburtstag fehlt";
                    WENN(Geburt="";"Ungültiger Geburtstag";
                        WENN(UND(NICHT(LeerT);Tod="");"Ungültiger Sterbetag";
                            WENN(Geburt>Heute400;"Geburtstag in der Zukunft";
                                WENN(UND(NICHT(LeerT);Tod>Heute400);"Sterbetag in der Zukunft";
                                    WENN(UND(NICHT(LeerT);Tod<Geburt);"Sterbetag vor Geburtstag";
                                        WENNFEHLER(
                                            DATEDIF(Geburt;WENN(LeerT;Heute400;Tod);"y");
                                            "Keine Berechnung"
                                        )
                                    )
                                )
                            )
                        )
                    )
                )
            )
        )
    ));
    VSTAPELN("Alter (Formel)";Ergebnis)
)

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
// tblModerne_Daten_1
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblModerneDaten"]}[Content],
    Heute = Date.From(DateTime.FixedLocalNow()),
    AlterBerechnen = (Geburt as nullable date,
                      Tod as nullable date) as nullable number =>
        let
            Stichtag = if Tod = null then Heute else Tod,
            Alter =
                if Geburt = null or Stichtag < Geburt then
                    null
                else
                    Date.Year(Stichtag) - Date.Year(Geburt)
                    - (
                        if Date.Month(Stichtag) < Date.Month(Geburt)
                            or (
                                Date.Month(Stichtag) = Date.Month(Geburt)
                                and Date.Day(Stichtag) < Date.Day(Geburt)
                            )
                        then 1
                        else 0
                    )
        in
            Alter,
    Result = Table.AddColumn(
        Quelle,
        "Alter",
        each AlterBerechnen(
            try Date.From([Geburtstag]) otherwise null,
            try Date.From([Sterbetag]) otherwise null
        ),
        Int64.Type
    ),
Erg = Table.SelectColumns(Result, "Alter")    
in
    Erg

// tblModerne_Daten_2 - Nur Verbindung
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblModerneDaten"]}[Content],
    Heute = Date.From(DateTime.FixedLocalNow()),
    // Berechnung der vollendeten Lebensjahre
    AlterBerechnen = (Geburt as date, Stichtag as date) as number =>
        Date.Year(Stichtag) - Date.Year(Geburt)
        - (
            if Date.Month(Stichtag) < Date.Month(Geburt)
                or (
                    Date.Month(Stichtag) = Date.Month(Geburt)
                    and Date.Day(Stichtag) < Date.Day(Geburt)
                )
            then 1
            else 0
        ),
    // Datum prüfen, ohne fehlerhafte Werte zu verschlucken
    DatumPruefen = (Wert as any) as record =>
        let
            Leer = Wert = null or (Value.Is(Wert, type text) and Text.Trim(Wert) = ""),
            Versuch =
                if Leer then
                    [HasError = false, Value = null]
                else
                    try Date.From(Wert, "de-DE")
        in
            [
                Datum = if Versuch[HasError] then null else Versuch[Value],
                Ungueltig = Versuch[HasError]
            ],
    // Status und Alter gemeinsam ermitteln
    Berechnung = Table.AddColumn(
        Quelle,
        "Pruefung",
        each
            let
                G = DatumPruefen([Geburtstag]),
                T = DatumPruefen([Sterbetag]),
                Geburt = G[Datum],
                Tod = T[Datum],
                Status =
                    if G[Ungueltig] then "Ungültiger Geburtstag"
                    else if T[Ungueltig] then "Ungültiger Sterbetag"
                    else if Geburt = null and Tod = null then
                        "Beide Datumsfelder fehlen"
                    else if Geburt = null then "Geburtstag fehlt"
                    else if Geburt > Heute then "Geburtstag in der Zukunft"
                    else if Tod <> null and Tod < Geburt then
                        "Sterbetag vor Geburtstag"
                    else if Tod <> null and Tod > Heute then
                        "Sterbetag in der Zukunft"
                    else "OK",
                Alter =
                    if Status = "OK" then
                        AlterBerechnen(
                            Geburt,
                            if Tod = null then Heute else Tod
                        )
                    else
                        null
            in
                [Alter = Alter, Status = Status],
        type record
    ),
    Erweitert = Table.ExpandRecordColumn(
        Berechnung,
        "Pruefung",
        {"Alter", "Status"}
    ),
    Typen = Table.TransformColumnTypes(
        Erweitert,
        {{"Alter", Int64.Type}, {"Status", type text}}
    ),
    Erg = Table.SelectColumns(Typen, {"Alter", "Status"})
in
    Erg

// tblModerne_Daten_3
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblModerneDaten"]}[Content],
    Heute = Date.From(DateTime.FixedLocalNow()),
    // Berechnung der vollendeten Lebensjahre
    AlterBerechnen = (Geburt as date, Stichtag as date) as number =>
        Date.Year(Stichtag) - Date.Year(Geburt)
        - (
            if Date.Month(Stichtag) < Date.Month(Geburt)
                or (
                    Date.Month(Stichtag) = Date.Month(Geburt)
                    and Date.Day(Stichtag) < Date.Day(Geburt)
                )
            then 1
            else 0
        ),
    // Datum prüfen, ohne fehlerhafte Werte zu verschlucken
    DatumPruefen = (Wert as any) as record =>
        let
            Leer = Wert = null or (Value.Is(Wert, type text) and Text.Trim(Wert) = ""),
            Excel1900Fehler =
                if Leer then false
                else Value.Is(Wert, type number) and Wert >= 60 and Wert < 61,
            Versuch =
                if Leer or Excel1900Fehler then
                    [HasError = false, Value = null]
                else
                    try Date.From(Wert, "de-DE")
        in
            [
                Datum = if Versuch[HasError] then null else Versuch[Value],
                Ungueltig = Versuch[HasError],
                Excel1900Fehler = Excel1900Fehler
            ],
    // Status und Alter gemeinsam ermitteln
    Berechnung = Table.AddColumn(
        Quelle,
        "Pruefung",
        each
            let
                G = DatumPruefen([Geburtstag]),
                T = DatumPruefen([Sterbetag]),
                Geburt = G[Datum],
                Tod = T[Datum],
                Status =
                    if G[Excel1900Fehler] then "Ungültig: 29.02.1900 (Excel)"
                    else if T[Excel1900Fehler] then "Ungültig: 29.02.1900 (Excel)"
                    else if G[Ungueltig] then "Ungültiger Geburtstag"
                    else if T[Ungueltig] then "Ungültiger Sterbetag"
                    else if Geburt = null and Tod = null then "Beide Datumsfelder fehlen"
                    else if Geburt = null then "Geburtstag fehlt"
                    else if Geburt > Heute then "Geburtstag in der Zukunft"
                    else if Tod <> null and Tod < Geburt then "Sterbetag vor Geburtstag"
                    else if Tod <> null and Tod > Heute then "Sterbetag in der Zukunft"
                    else "OK",
                Alter =
                    if Status = "OK" then
                        AlterBerechnen(
                            Geburt,
                            if Tod = null then Heute else Tod
                        )
                    else
                        null
            in
                [Alter = Alter, Status = Status],
        type record
    ),
    Erweitert = Table.ExpandRecordColumn(
        Berechnung,
        "Pruefung",
        {"Alter", "Status"}
    ),
    Typen = Table.TransformColumnTypes(
        Erweitert,
        {{"Alter", Int64.Type}, {"Status", type text}}
    ),
    Erg = Table.SelectColumns(Typen, {"Alter", "Status"})
in
    Erg

// Pruefung_29_02_1900 - Nur Verbindung
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblModerneDaten"]}[Content],
    Wert = Quelle{17}[Geburtstag],
    Erg = [
        Wert = Wert,
        Typ = Value.Type(Wert),
        IstZahl = Value.Is(Wert, type number),
        AlsDatum = try Date.From(Wert) otherwise null,
        AlsZahl = try Number.From(Wert) otherwise null
    ]
in
    Erg

// tblHistorischeDaten_1
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblHistorischeDaten"]}[Content],
    // Kalender und Schaltjahre
    IstSchaltjahr = (J as number, K as text) as logical =>
        if K = "Julianisch" then
            Number.Mod(J, 4) = 0
        else
            Number.Mod(J, 400) = 0 or
            (Number.Mod(J, 4) = 0 and Number.Mod(J, 100) <> 0),
    // Textdatum zunächst in Bestandteile zerlegen
    DatumLesen = (Wert as any) as record =>
        let
            S = if Wert = null then "" else Text.Trim(Text.From(Wert)),
            Teile = Text.Split(S, "."),
            FormatOK =
                Text.Length(S) = 10
                and List.Count(Teile) = 3
                and Text.Length(Teile{0}) = 2
                and Text.Length(Teile{1}) = 2
                and Text.Length(Teile{2}) = 4
                and Text.Length(Text.Select(S, {"0".."9"})) = 8,
            Zahlen =
                if FormatOK then
                    List.Transform(
                        Teile,
                        each try Number.FromText(_, "en-US") otherwise null
                    )
                else {null, null, null},
            T = Zahlen{0},
            M = Zahlen{1},
            J = Zahlen{2},
            StrukturOK =
                T <> null and M <> null and J <> null
                and J >= 1 and J <= 9999
                and M >= 1 and M <= 12
                and T >= 1
        in
            [
                Leer = S = "",
                StrukturOK = StrukturOK,
                Tag = T,
                Monat = M,
                Jahr = J
            ],
    // Kalendertag auf Gültigkeit prüfen
    DatumGueltig = (D as record, K as text) as logical =>
        let
            M = D[Monat],
            J = D[Jahr],
            T = D[Tag],
            Tage =
                if not D[StrukturOK] then 0
                else if M = 2 then
                    if IstSchaltjahr(J, K) then 29 else 28
                else if List.Contains({4, 6, 9, 11}, M) then 30
                else 31
        in
            D[StrukturOK] and T <= Tage,
    // Kalender nach Region und Datum bestimmen
    KalenderBestimmen = (
        D as record,
        Kalender as any,
        Region as any
    ) as text =>
        let
            K = if Kalender = null then ""
                else Text.Trim(Text.From(Kalender)),
            R = if Region = null then ""
                else Text.Lower(Text.Trim(Text.From(Region))),
            Zahl =
                if D[StrukturOK] then
                    D[Jahr] * 10000 + D[Monat] * 100 + D[Tag]
                else null,
            Ergebnis =
                if K = "Julianisch" or K = "Gregorianisch" then K
                else if K <> "Historisch regional" then "Unbekannt"
                else if Zahl = null then "Unbekannt"
                else if Text.Contains(R, "rom / 1582") then
                    if Zahl <= 15821004 then "Julianisch"
                    else if Zahl >= 15821015 then "Gregorianisch"
                    else "Kalenderlücke"
                else if Text.Contains(R, "großbritannien / 1752") then
                    if Zahl <= 17520902 then "Julianisch"
                    else if Zahl >= 17520914 then "Gregorianisch"
                    else "Kalenderlücke"
                else "Unbekannt"
        in
            Ergebnis,
    // Julianische Tageszahl berechnen
    ZuJDN = (
        J as number,
        M as number,
        T as number,
        K as text
    ) as number =>
        let
            A = Number.IntegerDivide(14 - M, 12),
            Y = J + 4800 - A,
            MM = M + 12 * A - 3,
            Basis =
                T
                + Number.IntegerDivide(153 * MM + 2, 5)
                + 365 * Y
                + Number.IntegerDivide(Y, 4),
            Result =
                if K = "Julianisch" then
                    Basis - 32083
                else
                    Basis
                    - Number.IntegerDivide(Y, 100)
                    + Number.IntegerDivide(Y, 400)
                    - 32045
        in
            Result,
    // JDN in julianisches oder gregorianisches Datum
    VonJDN = (N as number, K as text) as record =>
        let
            A = N + 32044,
            B =
                if K = "Gregorianisch" then
                    Number.IntegerDivide(4 * A + 3, 146097)
                else 0,
            C =
                if K = "Gregorianisch" then
                    A - Number.IntegerDivide(146097 * B, 4)
                else
                    N + 32082,
            D = Number.IntegerDivide(4 * C + 3, 1461),
            E = C - Number.IntegerDivide(1461 * D, 4),
            M = Number.IntegerDivide(5 * E + 2, 153),
            Tag = E - Number.IntegerDivide(153 * M + 2, 5) + 1,
            Monat = M + 3 - 12 * Number.IntegerDivide(M, 10),
            Jahr = 100 * B + D - 4800
                + Number.IntegerDivide(M, 10)
        in
            [Jahr = Jahr, Monat = Monat, Tag = Tag],
    // Vollendete Lebensjahre
    AlterBerechnen = (
        Geburt as record,
        Ende as record
    ) as number =>
        let
            Jahre = Ende[Jahr] - Geburt[Jahr],
            GeburtstagErreicht =
                Ende[Monat] > Geburt[Monat]
                or (
                    Ende[Monat] = Geburt[Monat]
                    and Ende[Tag] >= Geburt[Tag]
                ),
            Alter = Jahre - (if GeburtstagErreicht then 0 else 1)
        in
            Alter,
    // Eine Tabellenzeile auswerten
    Pruefen = (Zeile as record) as record =>
        let
            G = DatumLesen(Zeile[#"Geburt (Text)"]),
            T = DatumLesen(Zeile[#"Tod (Text)"]),
            Region = Zeile[#"Region / Regel"],
            KG = KalenderBestimmen(
                G, Zeile[Kalender Geburt], Region
            ),
            KT = KalenderBestimmen(
                T, Zeile[Kalender Tod], Region
            ),
            Erlaubt = {"Julianisch", "Gregorianisch"},
            KalenderOK = List.Contains(Erlaubt, KG) and List.Contains(Erlaubt, KT),
            GeburtOK =
                if List.Contains(Erlaubt, KG) then
                    DatumGueltig(G, KG)
                else false,
            TodOK =
                if List.Contains(Erlaubt, KT) then
                    DatumGueltig(T, KT)
                else false,
            Berechenbar = GeburtOK and TodOK and KalenderOK,
            JG =
                if Berechenbar then
                    ZuJDN(G[Jahr], G[Monat], G[Tag], KG)
                else null,
            JT =
                if Berechenbar then
                    ZuJDN(T[Jahr], T[Monat], T[Tag], KT)
                else null,
            // Tod in den Geburtskalender umrechnen
            Ende =
                if Berechenbar then VonJDN(JT, KG)
                else null,
            Status =
                if G[Leer] and T[Leer] then
                    "Beide Datumsfelder fehlen"
                else if G[Leer] then
                    "Geburtstag fehlt"
                else if T[Leer] then
                    "Sterbetag fehlt"
                else if not G[StrukturOK] then
                    "Ungültiger Geburtstag"
                else if not T[StrukturOK] then
                    "Ungültiger Sterbetag"
                else if KG = "Kalenderlücke"
                    or KT = "Kalenderlücke" then
                    "Datum liegt in Kalenderlücke"
                else if not KalenderOK then
                    "Kalender nicht unterstützt"
                else if not GeburtOK then
                    "Ungültiger Geburtstag"
                else if not TodOK then
                    "Ungültiger Sterbetag"
                else if JT < JG then
                    "Sterbetag vor Geburtstag"
                else "OK",
            Alter =
                if Status = "OK" then
                    AlterBerechnen(G, Ende)
                else null
        in
            [Alter = Alter, Status = Status],
    Berechnung = Table.AddColumn(
        Quelle,
        "Pruefung",
        each Pruefen(_),
        type record
    ),
    Erweitert = Table.ExpandRecordColumn(
        Berechnung,
        "Pruefung",
        {"Alter", "Status"}
    ),
    Typen = Table.TransformColumnTypes(
        Erweitert,
        {
            {"Alter", Int64.Type},
            {"Status", type text}
        }
    ),
    Erg = Table.SelectColumns(
        Typen,
        {"Alter", "Status"}
    )
in
    Erg

//tblAntikeDaten_1
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblAntikeDaten"]}[Content],
    // Historische Jahreszählung:
    // 1 v. Chr. = astronomisches Jahr 0
    // 2 v. Chr. = astronomisches Jahr -1
    // 1 n. Chr. = astronomisches Jahr 1
    DatumLesen = (Wert as any) as record =>
        let
            Roh =
                if Wert = null then ""
                else Text.Trim(Text.From(Wert)),

            Normal = Text.Lower(
                Text.Remove(Roh, {" ", Character.FromNumber(160)})
            ),
            VorChr = Text.EndsWith(Normal, "v.chr."),
            NachChr = Text.EndsWith(Normal, "n.chr."),
            Kern =
                if VorChr or NachChr then
                    Text.Start(Normal, Text.Length(Normal) - 6)
                else
                    Normal,
            Teile = Text.Split(Kern, "."),
            FormatOK =
                List.Count(Teile) = 3
                and Text.Length(Teile{0}) = 2
                and Text.Length(Teile{1}) = 2
                and Text.Length(Teile{2}) = 4
                and Text.Length(Text.Select(Kern, {"0".."9"})) = 8,
            Zahlen =
                if FormatOK then
                    List.Transform(
                        Teile,
                        each try Number.FromText(_, "en-US") otherwise null
                    )
                else
                    {null, null, null},
            T = Zahlen{0},
            M = Zahlen{1},
            J = Zahlen{2},
            StrukturOK =
                FormatOK
                and T <> null and M <> null and J <> null
                and T >= 1
                and M >= 1 and M <= 12
                and J >= 1,
            JahrAstronomisch =
                if not StrukturOK then null
                else if VorChr then 1 - J
                else J
        in
            [
                Leer = Roh = "",
                StrukturOK = StrukturOK,
                Tag = T,
                Monat = M,
                Jahr = JahrAstronomisch
            ],
    // Schaltjahrregel für den proleptischen julianischen Kalender
    IstSchaltjahr = (J as number) as logical => Number.Mod(J, 4) = 0,
    DatumGueltig = (D as record) as logical =>
        let
            M = D[Monat],
            J = D[Jahr],
            T = D[Tag],
            Tage =
                if not D[StrukturOK] then 0
                else if M = 2 then
                    if IstSchaltjahr(J) then 29 else 28
                else if List.Contains({4, 6, 9, 11}, M) then
                    30
                else
                    31
        in
            D[StrukturOK] and T <= Tage,
    // Nur explizit proleptisch julianische Daten berechnen
    KalenderOK = (Wert as any) as logical =>
        Wert <> null
        and Text.Lower(Text.Trim(Text.From(Wert))) =
            "proleptisch julianisch",
    // Datum -> Julianische Tageszahl
    ZuJDN = (J as number, M as number, T as number) as number =>
        let
            A = Number.IntegerDivide(14 - M, 12),
            Y = J + 4800 - A,
            MM = M + 12 * A - 3,
            Result =
                T
                + Number.IntegerDivide(153 * MM + 2, 5)
                + 365 * Y
                + Number.IntegerDivide(Y, 4)
                - 32083
        in
            Result,
    // Julianische Tageszahl -> Datum
    VonJDN = (N as number) as record =>
        let
            C = N + 32082,
            D = Number.IntegerDivide(4 * C + 3, 1461),
            E = C - Number.IntegerDivide(1461 * D, 4),
            M = Number.IntegerDivide(5 * E + 2, 153),
            Tag = E - Number.IntegerDivide(153 * M + 2, 5) + 1,
            Monat = M + 3 - 12 * Number.IntegerDivide(M, 10),
            Jahr = D - 4800 + Number.IntegerDivide(M, 10)
        in
            [Jahr = Jahr, Monat = Monat, Tag = Tag],
    // Vollendete Lebensjahre
    AlterBerechnen = (Geburt as record, Ende as record) as number =>
        let
            Jahre = Ende[Jahr] - Geburt[Jahr],
            GeburtstagErreicht =
                Ende[Monat] > Geburt[Monat]
                or (
                    Ende[Monat] = Geburt[Monat]
                    and Ende[Tag] >= Geburt[Tag]
                ),
            Alter =
                Jahre - (if GeburtstagErreicht then 0 else 1)
        in
            Alter,
    // Genauigkeit der Datumsangaben prüfen
    GenauigkeitOK = (Wert as any) as logical =>
        let
            G =
                if Wert = null then ""
                else Text.Lower(Text.Trim(Text.From(Wert)))
        in
            List.Contains({"synthetisch", "exakt"}, G),
    // Prüfung einer Tabellenzeile
    Pruefen = (Zeile as record) as record =>
        let
            G = DatumLesen(Zeile[#"Geburt (Text)"]),
            T = DatumLesen(Zeile[#"Tod (Text)"]),
            KG = KalenderOK(Zeile[#"Kalender Geburt"]),
            KT = KalenderOK(Zeile[#"Kalender Tod"]),
            GenauOK = GenauigkeitOK(Record.FieldOrDefault(Zeile, "Genauigkeit", null)),
            Berechenbar =
                KG and KT
                and DatumGueltig(G)
                and DatumGueltig(T)
                and GenauOK,
            JG =
                if Berechenbar then
                    ZuJDN(G[Jahr], G[Monat], G[Tag])
                else
                    null,
            JT =
                if Berechenbar then
                    ZuJDN(T[Jahr], T[Monat], T[Tag])
                else
                    null,
            Ende =
                if Berechenbar then VonJDN(JT)
                else null,
            Status =
                if G[Leer] and T[Leer] then
                    "Beide Datumsfelder fehlen"
                else if G[Leer] then
                    "Geburtstag fehlt"
                else if T[Leer] then
                    "Sterbetag fehlt"
                else if not G[StrukturOK] then
                    "Ungültiger Geburtstag"
                else if not T[StrukturOK] then
                    "Ungültiger Sterbetag"
                else if not KG or not KT then
                    "Kalender nicht unterstützt"
                else if not DatumGueltig(G) then
                    "Ungültiger Geburtstag"
                else if not DatumGueltig(T) then
                    "Ungültiger Sterbetag"
                else if not GenauOK then
                    "Keine exakte Altersberechnung"
                else if JT < JG then
                    "Sterbetag vor Geburtstag"
                else
                    "OK",
            Alter =
                if Status = "OK" then
                    AlterBerechnen(G, Ende)
                else
                    null
        in
            [Alter = Alter, Status = Status],
    Berechnung = Table.AddColumn(
        Quelle,
        "Pruefung",
        each Pruefen(_),
        type record
    ),
    Erweitert = Table.ExpandRecordColumn(
        Berechnung,
        "Pruefung",
        {"Alter", "Status"}
    ),
    Typen = Table.TransformColumnTypes(
        Erweitert,
        {
            {"Alter", Int64.Type},
            {"Status", type text}
        }
    ),
    Erg = Table.SelectColumns(Typen, {"Alter", "Status"})
in
    Erg

// tblTest_1_bis_2100 - Nur Verbindung
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    // Testbereich
    VonJahr = 1,
    BisJahr = 2100,
    // Statt Referenzen Massentest wählen, dann wird komplett ausgewertet
    Testmodus = "Referenzen",
    //Testmodus = "Massentest",
    // Schaltjahrregel
    IstSchaltjahr = (J as number, K as text) as logical =>
        if K = "Julianisch" then
            Number.Mod(J, 4) = 0
        else
            Number.Mod(J, 400) = 0
            or (
                Number.Mod(J, 4) = 0
                and Number.Mod(J, 100) <> 0
            ),
    // Anzahl der Tage eines Monats
    Monatstage = (J as number, M as number, K as text) as number =>
        if M = 2 then
            if IstSchaltjahr(J, K) then 29 else 28
        else if List.Contains({4, 6, 9, 11}, M) then
            30
        else
            31,
    // Kalenderdatum -> JDN
    ZuJDN = (J as number, M as number, T as number, K as text) as number =>
        let
            A = Number.IntegerDivide(14 - M, 12),
            Y = J + 4800 - A,
            MM = M + 12 * A - 3,
            Basis =
                T
                + Number.IntegerDivide(153 * MM + 2, 5)
                + 365 * Y
                + Number.IntegerDivide(Y, 4),
            Result =
                if K = "Julianisch" then
                    Basis - 32083
                else
                    Basis
                    - Number.IntegerDivide(Y, 100)
                    + Number.IntegerDivide(Y, 400)
                    - 32045
        in
            Result,
    // JDN -> Kalenderdatum
    VonJDN = (N as number, K as text) as record =>
        let
            A = N + 32044,
            B =
                if K = "Gregorianisch" then
                    Number.IntegerDivide(4 * A + 3, 146097)
                else 0,
            C =
                if K = "Gregorianisch" then
                    A - Number.IntegerDivide(146097 * B, 4)
                else
                    N + 32082,
            D = Number.IntegerDivide(4 * C + 3, 1461),
            E = C - Number.IntegerDivide(1461 * D, 4),
            M = Number.IntegerDivide(5 * E + 2, 153),
            Tag = E - Number.IntegerDivide(153 * M + 2, 5) + 1,
            Monat = M + 3 - 12 * Number.IntegerDivide(M, 10),
            Jahr =
                100 * B + D - 4800
                + Number.IntegerDivide(M, 10)
        in
            [Jahr = Jahr, Monat = Monat, Tag = Tag],
    // Alle gültigen Testdaten erzeugen
    Testdaten =
        List.Combine(
            List.Transform(
                {"Julianisch", "Gregorianisch"},
                (K) =>
                    List.Combine(
                        List.Transform(
                            {VonJahr..BisJahr},
                            (J) =>
                                List.Combine(
                                    List.Transform(
                                        {1..12},
                                        (M) =>
                                            List.Transform(
                                                {1..Monatstage(J, M, K)},
                                                (T) =>
                                                    [
                                                        Kalender = K,
                                                        Jahr = J,
                                                        Monat = M,
                                                        Tag = T
                                                    ]
                                            )
                                    )
                                )
                        )
                    )
            )
        ),
    Tabelle = Table.FromRecords(
        Testdaten,
        type table [
            Kalender = text,
            Jahr = Int64.Type,
            Monat = Int64.Type,
            Tag = Int64.Type
        ]
    ),
    // JDN und Rückrechnung
    Berechnet = Table.AddColumn(
        Tabelle,
        "Pruefung",
        each
            let
                N = ZuJDN([Jahr], [Monat], [Tag], [Kalender]),
                R = VonJDN(N, [Kalender]),
                OK =
                    R[Jahr] = [Jahr]
                    and R[Monat] = [Monat]
                    and R[Tag] = [Tag]
            in
                [
                    JDN = N,
                    RueckJahr = R[Jahr],
                    RueckMonat = R[Monat],
                    RueckTag = R[Tag],
                    Status = if OK then "OK" else "FEHLER"
                ],
        type record
    ),
    Erweitert = Table.ExpandRecordColumn(
        Berechnet,
        "Pruefung",
        {"JDN", "RueckJahr", "RueckMonat", "RueckTag", "Status"}
    ),
    // Bekannte JDN-Referenzwerte
    Referenzen = #table(
        type table [
            Kalender = text,
            Jahr = Int64.Type,
            Monat = Int64.Type,
            Tag = Int64.Type,
            ErwarteteJDN = Int64.Type
        ],
        {
            {"Gregorianisch", 2000, 1, 1, 2451545},
            {"Gregorianisch", 1582, 10, 15, 2299161},
            {"Julianisch", 1582, 10, 4, 2299160},
            {"Gregorianisch", 1900, 3, 1, 2415080}
        }
    ),
    ReferenzPruefung = Table.AddColumn(
        Referenzen,
        "BerechneteJDN",
        each ZuJDN([Jahr], [Monat], [Tag], [Kalender]),
        Int64.Type
    ),
    ReferenzStatus = Table.AddColumn(
        ReferenzPruefung,
        "Status",
        each if [ErwarteteJDN] = [BerechneteJDN]
             then "OK"
             else "FEHLER",
        type text
    ),
    Erg =
        if Testmodus = "Referenzen" then
            ReferenzStatus
        else
            Table.SelectRows(Erweitert, each [Status] <> "OK")
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0034",
                .Category = DemoCategory.PowerQuery,
                .Title = "Zerlegung, Normalisierung und Plausibilität - Kennzeichen...",
                .Tags = {"kennzeichen", "power query", "zerlegen", "text", "replace", "varianten"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Vier Wege zur Zerlegung, Normalisierung und Plausibilitätsprüfung deutscher Kfz-Kennzeichen.
Gelöst mit Formeln, Regex, VBA und Power Query.

RegEx_PQ_Formel_VBA_Kennzeichen_Demo.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
// tblKennzeichen_1
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    // Zentrale Datenquelle
    Quelle = Excel.CurrentWorkbook(){[Name="tblKennzeichen"]}[Content],
    // Hilfsfunktion: Kennzeichen zerlegen
    fnZerlegen = (Eingabe as nullable text) as record =>
    let
        Roh = if Eingabe = null then "" else Eingabe,
        Text = Text.Upper(Text.Trim(Text.Clean(Roh))),
        // Leerzeichen vereinheitlichen
        Leerzeichen = Text.Replace(Text, Character.FromNumber(160), " "),
        Teile = List.Select(Text.Split(Leerzeichen, " "), each _ <> ""),
        Bereinigt = Text.Combine(Teile, " "),
        // Saisonangabe am Ende abtrennen
        SaisonTeile = Text.Split(Bereinigt, "/"),
        HatSaison = List.Count(SaisonTeile) = 2
            and Text.Length(Text.End(SaisonTeile{0}, 2)) = 2
            and Text.Length(SaisonTeile{1}) = 2
            and Text.Select(Text.End(SaisonTeile{0}, 2) & SaisonTeile{1}, {"0".."9"})
                = Text.End(SaisonTeile{0}, 2) & SaisonTeile{1},
        Von = if HatSaison then Text.End(SaisonTeile{0}, 2) else "",
        Bis = if HatSaison then SaisonTeile{1} else "",
        OhneSaison = if HatSaison then Text.Start(SaisonTeile{0}, Text.Length(SaisonTeile{0}) - 2) else Bereinigt,
        // E/H-Zusatz abtrennen
        Stamm = Text.Trim(OhneSaison),
        LetztesZeichen = Text.End(Stamm, 1),
        Zusatz = if List.Contains({"E", "H"}, LetztesZeichen) then LetztesZeichen else "",
        OhneZusatz = if Zusatz <> "" then Text.Trim(Text.Start(Stamm, Text.Length(Stamm) - 1)) else Stamm,
        // Bindestriche und Leerzeichen als Trenner
        MitLeerzeichen = Text.Replace(OhneZusatz, "-", " "),
        Wortliste = List.Select(Text.Split(MitLeerzeichen, " "), each _ <> ""),
        // Buchstaben und Ziffern bestimmen
        Kompakt = Text.Combine(Wortliste, ""),
        Buchstaben = Text.Select(Kompakt, {"A".."Z"}),
        Ziffern = Text.Select(Kompakt, {"0".."9"}),
        // Aufteilung der Buchstabengruppen
        Bezirk = if List.Count(Wortliste) >= 2
            then Wortliste{0}
            else if Text.Length(Buchstaben) >= 2
                then Text.Start(Buchstaben, Text.Length(Buchstaben) - 1)
                else "",
        Gruppe = if List.Count(Wortliste) >= 2
            then Wortliste{1}
            else if Text.Length(Buchstaben) >= 2
                then Text.End(Buchstaben, 1)
                else "",
        // Nur Buchstaben in den Gruppen
        BezirkOK = Text.Select(Bezirk, {"A".."Z"}),
        GruppeOK = Text.Select(Gruppe, {"A".."Z"}),
        Nummer = Ziffern,
        // Einfache Strukturprüfung
        Erkannt =
            Text.Length(BezirkOK) >= 1
            and Text.Length(BezirkOK) <= 3
            and Text.Length(GruppeOK) >= 1
            and Text.Length(GruppeOK) <= 2
            and Text.Length(Nummer) >= 1
            and Text.Length(Nummer) <= 4,
        Normalisiert = if Erkannt then
            BezirkOK & "-" & GruppeOK & " " & Nummer & Zusatz &
            (if HatSaison then " " & Von & "/" & Bis else "")
            else "",
        Result = [
            Bezirk = if Erkannt then BezirkOK else "",
            Buchstaben = if Erkannt then GruppeOK else "",
            Nummer = if Erkannt then Nummer else "",
            Zusatz = if Erkannt then Zusatz else "",
            Von = if Erkannt then Von else "",
            Bis = if Erkannt then Bis else "",
            Normalisiert = Normalisiert
        ]
    in
        Result,
    // Quelldaten übernehmen
    Typ = Table.TransformColumnTypes(Quelle, {{"Kennzeichen", type text}}),
    Umbenannt = Table.RenameColumns(Typ, {{"Kennzeichen", "Eingabe"}}),
    // Funktion auf alle Zeilen anwenden
    Zerlegt = Table.AddColumn(
        Umbenannt,
        "Ergebnis",
        each fnZerlegen([Eingabe]),
        type record
    ),
    // Ergebnisfelder aufklappen
    Erg = Table.ExpandRecordColumn(
        Zerlegt,
        "Ergebnis",
        {"Bezirk", "Buchstaben", "Nummer",
         "Zusatz", "Von", "Bis", "Normalisiert"}
    )
in
    Erg

// tblKennzeichen_2
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    // ==============================================================
    // tblKennzeichen_2 - Kontrollierte Kennzeichenprüfung ohne RegEx
    // Datenquelle: tblKennzeichen
    // ==============================================================
    Quelle = Excel.CurrentWorkbook(){[Name="tblKennzeichen"]}[Content],
    Buchstaben = {"A".."Z"},
    Ziffern = {"0".."9"},
    // Prüfen, ob ein Text ausschließlich erlaubte Zeichen hat
    fnNurZeichen = (t as text, erlaubt as list) as logical =>
        Text.Length(t) > 0
        and Text.Length(Text.Remove(t, erlaubt)) = 0,
    // Leere Ergebnisstruktur
    Leer = [
        Bezirk = "",
        Buchstaben = "",
        Nummer = "",
        Zusatz = "",
        Von = "",
        Bis = "",
        Normalisiert = "",
        Status = "Kein Treffer",
        Diagnose = "Struktur nicht erkannt"
    ],
    fnPruefen = (Eingabe as nullable text) as record =>
    let
        Roh = if Eingabe = null then "" else Eingabe,
        // Großschreibung und äußere Leerzeichen
        Gross = Text.Upper(Text.Trim(Roh)),
        // Mehrfache Leerzeichen und Bindestriche
        // zunächst nicht entfernen, sondern strukturiert behandeln
        Teile = List.Select(
            Text.SplitAny(Gross, " -"),
            each _ <> ""
        ),
        // -------------------------------------------------
        // 1. Saisonangabe abtrennen
        // -------------------------------------------------
        // Leerzeichen um den Schrägstrich erlauben
        SaisonText = Text.Replace(
            Text.Replace(Gross, " /", "/"),
            "/ ", "/"
        ),
        SlashTeile = Text.Split(SaisonText, "/"),
        HatSlash = List.Count(SlashTeile) = 2,
        Rechts = if HatSlash then SlashTeile{1} else "",
        Links = if HatSlash then SlashTeile{0} else SaisonText,
        BisText = Text.Trim(Rechts),
        // Die letzten beiden Zeichen vor "/" sind der Monat
        VonText = if HatSlash
            then Text.End(Text.Trim(Links), 2)
            else "",
        SaisonFormat =
            HatSlash
            and Text.Length(VonText) = 2
            and Text.Length(BisText) = 2
            and fnNurZeichen(VonText, Ziffern)
            and fnNurZeichen(BisText, Ziffern)
            and List.Contains(
                {"01","02","03","04","05","06",
                "07","08","09","10","11","12"},
                VonText
            )
            and List.Contains(
                {"01","02","03","04","05","06",
                "07","08","09","10","11","12"},
                BisText
            ),
        OhneSaison =
            if SaisonFormat then
                Text.Trim(
                    Text.Start(
                        Text.Trim(Links),
                        Text.Length(Text.Trim(Links)) - 2
                    )
                )
            else if HatSlash then ""
            else SaisonText,
        // -------------------------------------------------
        // 2. E/H-Zusatz abtrennen
        // -------------------------------------------------
        Stamm = Text.Trim(OhneSaison),
        Ende = Text.End(Stamm, 1),
        Zusatz =
            if List.Contains({"E", "H"}, Ende)
            then Ende
            else "",
        OhneZusatz =
            if Zusatz <> "" then
                Text.Trim(Text.Start(Stamm, Text.Length(Stamm) - 1))
            else Stamm,
        // -------------------------------------------------
        // 3. Buchstaben und Nummer trennen
        // -------------------------------------------------
        Zeichen = Text.ToList(OhneZusatz),
        ErsteZiffer =
            List.PositionOf(
                List.Transform(
                    Zeichen,
                    each List.Contains(Ziffern, _)
                ),
                true
            ),
        BuchstabenTeil =
            if ErsteZiffer >= 0
            then Text.Start(OhneZusatz, ErsteZiffer)
            else "",
        NummerTeil =
            if ErsteZiffer >= 0
            then Text.Trim(Text.Range(OhneZusatz, ErsteZiffer))
            else "",
        // Bindestriche und Leerzeichen zwischen
        // Buchstabengruppen auswerten
        Gruppen = List.Select(
            Text.SplitAny(Text.Trim(BuchstabenTeil), " -"),
            each _ <> ""
        ),
        AnzahlGruppen = List.Count(Gruppen),
        AlleBuchstaben = Text.Combine(Gruppen, ""),
        // Ohne Trenner wie RegEx:
        // möglichst langer Bezirk, mindestens ein
        // Buchstabe für die Erkennungsgruppe
        Bezirk =
            if AnzahlGruppen = 2 then Gruppen{0}
            else if AnzahlGruppen = 1
                and Text.Length(AlleBuchstaben) >= 2
            then
                Text.Start(
                    AlleBuchstaben,
                    List.Min({3, Text.Length(AlleBuchstaben) - 1})
                )
            else "",
        Gruppe =
            if AnzahlGruppen = 2 then Gruppen{1}
            else if AnzahlGruppen = 1
                and Text.Length(AlleBuchstaben) >= 2
            then
                Text.Range(AlleBuchstaben, Text.Length(Bezirk))
            else "",
        // -------------------------------------------------
        // 4. Strukturprüfung
        // -------------------------------------------------
        BezirkOK =
            Text.Length(Bezirk) >= 1
            and Text.Length(Bezirk) <= 3
            and fnNurZeichen(Bezirk, Buchstaben),
        GruppeOK =
            Text.Length(Gruppe) >= 1
            and Text.Length(Gruppe) <= 2
            and fnNurZeichen(Gruppe, Buchstaben),
        NummerOK =
            Text.Length(NummerTeil) >= 1
            and Text.Length(NummerTeil) <= 4
            and fnNurZeichen(NummerTeil, Ziffern),
        // Unerlaubte Zeichen im gesamten Stamm verhindern
        // eine versehentliche "Reparatur"
        StammZeichenOK =
            Text.Length(
                Text.Remove(
                    OhneZusatz,
                    List.Combine({Buchstaben, Ziffern, {" ", "-"}})
                )
            ) = 0,
        // Nur erlaubte Trennzeichen, keine mehrfachen
        // Bindestriche und keine Ziffern in Buchstabengruppen
        TrennerOK =
            not Text.Contains(OhneZusatz, "--")
            and not Text.Contains(OhneZusatz, "- -")
            and not Text.Contains(NummerTeil, "-")
            and AnzahlGruppen >= 1
            and AnzahlGruppen <= 2,
        StrukturOK =
            BezirkOK
            and GruppeOK
            and NummerOK
            and StammZeichenOK
            and TrennerOK
            and (not HatSlash or SaisonFormat),
        // -------------------------------------------------
        // 5. Normalisierung
        // -------------------------------------------------
        Normalisiert =
            if StrukturOK then
                Bezirk & "-" & Gruppe & " " &
                NummerTeil & Zusatz &
                (if HatSlash then
                    " " & VonText & "/" & BisText
                 else "")
            else "",
        // -------------------------------------------------
        // 6. Plausibilitätsprüfung
        // -------------------------------------------------
        NummerWert = if StrukturOK then Number.FromText(NummerTeil) else 0,
        Laenge = Text.Length(Bezirk & Gruppe & NummerTeil & Zusatz),
        VonWert = if SaisonFormat then Number.FromText(VonText) else 0,
        BisWert = if SaisonFormat then Number.FromText(BisText) else 0,
        Monate = if SaisonFormat then BisWert - VonWert + 1 else 0,
        // Ohne Trenner: mögliche Mehrdeutigkeit
        Kompakt = Text.Remove(BuchstabenTeil, {" ", "-"}),
        Mehrdeutig =
            StrukturOK
            and AnzahlGruppen = 1
            and Text.Length(Kompakt) >= 3
            and Text.Length(Kompakt) <= 4,
        fNummer = StrukturOK and NummerWert = 0,
        fLaenge = StrukturOK and Laenge > 8,
        fSaison = StrukturOK and SaisonFormat and BisWert < VonWert,
        fKurz = StrukturOK and SaisonFormat and BisWert >= VonWert and Monate < 2,
        fLang = StrukturOK and SaisonFormat and Monate > 11,
        fMehr = Mehrdeutig,
        fNorm = StrukturOK and Roh <> Normalisiert,
        // -------------------------------------------------
        // 7. Mehrfachdiagnose
        // -------------------------------------------------
        Fehler = List.RemoveNulls({
            if fNummer then "Nummer ist 0" else null,
            if fLaenge then "Mehr als 8 Zeichen" else null,
            if fSaison then "Saisonende vor Saisonbeginn" else null,
            if fKurz then "Saison kürzer als 2 Monate" else null,
            if fLang then "Saison länger als 11 Monate" else null,
            if fMehr then "Aufteilung möglicherweise mehrdeutig" else null
        }),
        Status =
            if not StrukturOK then "Kein Treffer"
            else if List.Count(Fehler) > 0 then "Prüfen"
            else if fNorm then "Normalisiert"
            else "OK",
        Diagnose =
            if not StrukturOK then "Struktur nicht erkannt"
            else if List.Count(Fehler) > 0
                then Text.Combine(Fehler, " | ")
            else if fNorm then "Schreibweise vereinheitlicht"
            else "Keine Auffälligkeiten",
        Result =
            if not StrukturOK then Leer
            else [
                Bezirk = Bezirk,
                Buchstaben = Gruppe,
                Nummer = NummerTeil,
                Zusatz = Zusatz,
                Von = VonText,
                Bis = BisText,
                Normalisiert = Normalisiert,
                Status = Status,
                Diagnose = Diagnose
            ]
    in
        Result,
    // =====================================================
    // Abfrage auf die gesamte Tabelle anwenden
    // =====================================================
    Typ = Table.TransformColumnTypes(Quelle, {{"Kennzeichen", type text}}),
    Umbenannt = Table.RenameColumns(Typ, {{"Kennzeichen", "Eingabe"}}),
    Geprueft = Table.AddColumn(
        Umbenannt,
        "Ergebnis",
        each fnPruefen([Eingabe]),
        type record
    ),
    Erg = Table.ExpandRecordColumn(
        Geprueft,
        "Ergebnis",
        {
            "Bezirk", "Buchstaben", "Nummer",
            "Zusatz", "Von", "Bis", "Normalisiert",
            "Status", "Diagnose"
        }
    )
in
    Erg
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "pq_0035",
                .Category = DemoCategory.PowerQuery,
                .Title = "Monatsauswertung - Formeln, Power Query und Power Pivot...",
                .Tags = {"monatsauswertung", "power query", "gruppieren", "tag", "stunden", "varianten"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Monatsauswertung von Mitarbeiterbuchungen.
Gelöst mit Formeln, Pivot und Power Query.

Demo_Power_Query_Formeln_Gruppieren.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
/*
Excel-VSTO-Toolbox
Power Query-Demo
Ralf Stolzenburg (Case)
https://github.com/rstsu/Excel-VSTO-Toolbox
*/
let
    Quelle = Excel.CurrentWorkbook(){[Name="tblDaten"]}[Content],
    Typen = Table.TransformColumnTypes(Quelle, {
        {"Pers.-Nr.", Int64.Type},
        {"Mitarbeiter", type text},
        {"Datum", type date},
        {"Std. Menge", type number},
        {"Lohnart Bezeichnung", type text}
    }),
    GruppeTag = Table.Group(Typen,
        {"Pers.-Nr.", "Mitarbeiter", "Datum"}, {
        {"Tagesstunden", each List.Sum([Std. Menge]), type number},
        {"Urlaub", each
            Number.From(List.Contains([Lohnart Bezeichnung], "Urlaub")),
            Int64.Type},
        {"Krankheit", each
            Number.From(List.Contains([Lohnart Bezeichnung], "Krankheitstage")),
            Int64.Type}
    }),
    GruppeMonat = Table.Group(GruppeTag,
        {"Pers.-Nr.", "Mitarbeiter"}, {
        {"Anzahl Monatsstunden", each List.Sum([Tagesstunden]), type number},
        {"Anzahl Tage ≥8h", each
            List.Count(List.Select([Tagesstunden], each _ >= 8)),
            Int64.Type},
        {"Anzahl-Url.-Tage", each List.Sum([Urlaub]), Int64.Type},
        {"Anzahl-Kranktage", each List.Sum([Krankheit]), Int64.Type}
    }),
    Erg = Table.Sort(GruppeMonat, {{"Pers.-Nr.", Order.Ascending}})
in
    Erg
        ]]>
    </code>
        )
            }
        }
    End Function
End Module