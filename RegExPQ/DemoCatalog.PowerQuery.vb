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
            }
        }
    End Function
End Module

