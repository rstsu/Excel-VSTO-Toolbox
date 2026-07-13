Imports System.Xml.Linq

Partial Public Module DemoCatalog
    Private Function GetFormulaExamples() As List(Of DemoDefinition)
        Return New List(Of DemoDefinition) From {
            New DemoDefinition With {
                .Id = "formula_001",
                .Category = DemoCategory.Formula,
                .Title = "Eindeutige Werte sortieren",
                .Tags = {"formel", "eindeutig", "sortieren", "text", "formula2"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Beispiel mit EINDEUTIG und SORTIEREN.
Es werden nur eindeutige Werte aus Spalte A sortiert wiedergegeben.
Die zweite Formel (D1) ist mit Überschrift.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=SORTIEREN(EINDEUTIG(A2:.A999))

=VSTAPELN(A1;SORTIEREN(EINDEUTIG(A2:.A999)))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "formula_002",
                .Category = DemoCategory.Formula,
                .Title = "Zahlen ""richtig"" sortieren",
                .Tags = {"formel", "zahl", "sortieren", "text", "formula2"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Zahlen die als Text vorliegen (3.10.20) richtig sortieren.
Es soll also nicht 12.52.63 vor 2.11.74 einsortiert werden sondern in der richtigen Reihenfolge.
Die zweite Formel (G1) ist mit Überschriften.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=SORTIERENNACH(A2:.B999;--TEXTVOR(A2:.A999;".");1;--TEIL(TEXTNACH(A2:.A999;".");1;2);1)

=VSTAPELN(A1:B1;SORTIERENNACH(A2:.B999;--TEXTVOR(A2:.A999;".");1;--TEIL(TEXTNACH(A2:.A999;".");1;2);1))

Oder mit LET:
=LET(w;A2:.A999;x;A2:.B999;y;--TEXTVOR(w;".");z;--TEIL(TEXTNACH(w;".");1;2);SORTIERENNACH(x;y;1;z;1))

=LET(w;A2:.A999;x;A2:.B999;y;--TEXTVOR(w;".");z;--TEIL(TEXTNACH(w;".");1;2);VSTAPELN(A1:B1;SORTIERENNACH(x;y;1;z;1)))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "formula_003",
                .Category = DemoCategory.Formula,
                .Title = "Kalender erstellen",
                .Tags = {"formel", "zahl", "kalender", "bedingte formatierung", "datum"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Es wird ein Blockkalender erstellt (6 Zeilen und 7 Spalten).
Einmal nur der Kalender, dann mit Tagen als Überschrift und schließlich noch mit Monat als Überschrift.
Jahr und Monat kann jeweils aus "Daten - Gültigkeit - Liste" ausgewählt werden.
Markiert werden Feiertage, Besondere Tage, Wochenenden und der aktuelle Tag.
Bundeseinheitliche Feiertage in I2:I10.
Das färben (Schrift und Hintergrund) geht auch über die Bedingte Formatiereung:
Schrift grau und kursiv, wenn nicht im aktuellen Monat - Formel: =MONAT(I23)<>$F$2 und dann entsprechend formatieren.
Feiertage/Besondere Tage - Formel: =ISTZAHL(VERGLEICH(I23;$Q$2:$Q$10;0)) und dann entsprechend formatieren.
Heutiger Tag - Formel: =I23=HEUTE() und dann entsprechend formatieren.
Wochenenden - Formel: =WOCHENTAG(I23;2)>5 und dann entsprechend formatieren.
Adressen anpassen und anwenden auf den jeweiligen Kalenderbereich.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=LET(Jahr;$B$1;Monat;$B$2;Erster;DATUM(Jahr;Monat;1);Start;Erster-WOCHENTAG(Erster;2)+1;SEQUENZ(6;7;Start;1))

=LET(Jahr;$B$1;Monat;$B$2;Erster;DATUM(Jahr;Monat;1);Start;Erster-WOCHENTAG(Erster;2)+1;VSTAPELN(TEXT(SEQUENZ(1;7;2;1);"TTT");SEQUENZ(6;7;Start;1)))

=LET(Jahr;$B$1;Monat;$B$2;Erster;DATUM(Jahr;Monat;1);Start;Erster-WOCHENTAG(Erster;2)+1;VSTAPELN(HSTAPELN(TEXT(Erster;"MMMM JJJJ");"";"";"";"";"";"");TEXT(SEQUENZ(1;7;2;1);"TTT");SEQUENZ(6;7;Start;1)))

=LET(Jahr;$B$1;Ostersonntag;RUNDEN((TAG(MINUTE(Jahr/38)/2+55)&".4."&Jahr)/7;0)*7-6;Daten;VSTAPELN(DATUM(Jahr;{1;5;10;12;12};{1;1;3;25;26});Ostersonntag+{-2;1;39;50});Namen;{"Neujahr";"Tag der Arbeit";"Tag der Deutschen Einheit";"1. Weihnachtstag";"2. Weihnachtstag";"Karfreitag";"Ostermontag";"Christi Himmelfahrt";"Pfingstmontag"};SORTIERENNACH(HSTAPELN(Daten;Namen);Daten))
        ]]>
    </code>
        )
            }
        }
    End Function

End Module

