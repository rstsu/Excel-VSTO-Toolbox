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
            }
        }
    End Function
End Module

