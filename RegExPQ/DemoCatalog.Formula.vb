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
            },
            New DemoDefinition With {
                .Id = "formula_004",
                .Category = DemoCategory.Formula,
                .Title = "Kreuztabelle aus Liste erste 3 Buchstaben",
                .Tags = {"formel", "text", "kreuztabelle", "intelligente tabelle", "liste"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A15) wird eine Kreuztabelle erstellt.
Grundlage sind die ersten 3 gleichen Buchstaben.
Einmal bezogen auf einfache Werte (Spalte C).
Und dann auf eine "intelligente Tabelle" (Spalte A).
Formeln in E1, M1 und U1.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=LET(t;Demo_Formel_4[Wert];u;LINKS(t;3);v;EINDEUTIG(u);w;MAX(MAP(v;LAMBDA(x;SUMME(--(u=x)))));VSTAPELN("Pos"&SEQUENZ(;w);MATRIXERSTELLEN(ZEILEN(v);w;LAMBDA(y;z;WENNFEHLER(INDEX(FILTER(t;u=INDEX(v;y));z);"")))))

=LET(m;Demo_Formel_4[Wert];k;EINDEUTIG(LINKS(m;3));n;MAX(MAP(k;LAMBDA(x;SUMME(--(LINKS(m;3)=x)))));VSTAPELN("Pos"&SEQUENZ(;n);MATRIXERSTELLEN(ZEILEN(k);n;LAMBDA(z;p;WENNFEHLER(INDEX(FILTER(m;LINKS(m;3)=INDEX(k;z));p);"")))))

=LET(t;$A$2:.$A$999;u;LINKS(t;3);v;EINDEUTIG(u);w;MAX(MAP(v;LAMBDA(x;SUMME(--(u=x)))));VSTAPELN("Pos"&SEQUENZ(;w);MATRIXERSTELLEN(ZEILEN(v);w;LAMBDA(y;z;WENNFEHLER(INDEX(FILTER(t;u=INDEX(v;y));z);"")))))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "formula_005",
                .Category = DemoCategory.Formula,
                .Title = "Datumbrereich in Liste auflösen",
                .Tags = {"formel", "ferien", "Datum", "intelligente tabelle", "liste"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A6 - Ferientermine 2027 in Baden Württemberg) wird eine laufende Liste erstellt.

Formeln in E2 und H1 (diese ist mit Überschrift).
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=LET(v;FBW[Von];b;FBW[Bis];n;b-v+1;m;MAX(n);d;ZUSPALTE(WENN(SEQUENZ(m;;0)<=MTRANS(n-1);MTRANS(v)+SEQUENZ(m;;0);NV());3;WAHR);HSTAPELN(d;TEXT(d;"TTTT")))

=VSTAPELN({"Datum"."Tag"};LET(v;FBW[Von];b;FBW[Bis];n;b-v+1;m;MAX(n);d;ZUSPALTE(WENN(SEQUENZ(m;;0)<=MTRANS(n-1);MTRANS(v)+SEQUENZ(m;;0);NV());3;WAHR);HSTAPELN(d;TEXT(d;"TTTT"))))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "formula_006",
                .Category = DemoCategory.Formula,
                .Title = "Primzahlen auflisten - Von Bis",
                .Tags = {"formel", "primzahlen", "start", "ende", "zahlen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (B2:C5) werden die Primzahlen aufgelistet.

Aufgeteilt, da mit einer Spill-Formel - je nach Größe der vorgegebenen Zahlen - die Berechnung schnell zu umfangreich wird.

Formeln in E2, F2, G2, H2, J1 und K1.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=LET(Z;SEQUENZ(C2-B2+1;;B2);t;SEQUENZ(1;MAX(Z));FILTER(Z;(Z>1)*(MMULT(--(REST(Z;t)=0);MTRANS(t^0))=2)))
=LET(Z;SEQUENZ(C3-B3+1;;B3);t;SEQUENZ(1;MAX(Z));FILTER(Z;(Z>1)*(MMULT(--(REST(Z;t)=0);MTRANS(t^0))=2)))
=LET(Z;SEQUENZ(C4-B4+1;;B4);t;SEQUENZ(1;MAX(Z));FILTER(Z;(Z>1)*(MMULT(--(REST(Z;t)=0);MTRANS(t^0))=2)))
=LET(Z;SEQUENZ(C5-B5+1;;B5);t;SEQUENZ(1;MAX(Z));FILTER(Z;(Z>1)*(MMULT(--(REST(Z;t)=0);MTRANS(t^0))=2)))

=VSTAPELN("Gestapelt 1";VSTAPELN(E2#;F2#;G2#))

=VSTAPELN("Gestapelt 2";ZUSPALTE(E2:.G999;1;WAHR))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "formula_007",
                .Category = DemoCategory.Formula,
                .Title = "Filtern bestimmter Tage aus Liste mit Datum",
                .Tags = {"formel", "filter", "datum", "liste", "zahlen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:C368 - Geburtstagsliste) werden vorgegebene Tage aufgelistet.
Formel_aus_einer_Geburtstagsliste_oder_anderen_Listen_mit_Datum_bestimmte_Tage_ausgeben_FILTER.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\Formel_07

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Formeln in E2:H2, J2:M2, O2:R2, T1:W1.
In X1 ist eine Liste der Tage (Daten Gültigkeit).
Die Zellen in E2, G2, J2, L2, O2, Q2 sind "Benutzerdefiniertes Format". Je nach Aufbau der Formel gibt man 0 bis 6 oder 1 bis 7 ein.

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
=FILTER(C2:.C999;REST(C2:.C999;7)=E1;"----")
=LET(x;FILTER(C2:.C999;REST(C2:.C999;7)=E1;"----");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))
=FILTER(Tabelle1[Geburtsdatum];REST(Tabelle1[Geburtsdatum];7)=G1;"----")
=LET(x;FILTER(Tabelle1[Geburtsdatum];REST(Tabelle1[Geburtsdatum];7)=G1;"----");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))

=FILTER(C2:.C999;TEXT(C2:.C999;"tttt")=TEXT(J1;"tttt");"---")
=LET(x;FILTER(C2:.C999;TEXT(C2:.C999;"tttt")=TEXT(J1;"tttt");"---");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))
=FILTER(Tabelle1[Geburtsdatum];TEXT(Tabelle1[Geburtsdatum];"tttt")=TEXT(L1;"tttt");"---")
=LET(x;FILTER(Tabelle1[Geburtsdatum];TEXT(Tabelle1[Geburtsdatum];"tttt")=TEXT(L1;"tttt");"---");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))

=FILTER(C2:.C999;WOCHENTAG(C2:.C999;1)=O1;"----")
=LET(x;FILTER(C2:.C999;WOCHENTAG(C2:.C999;1)=O1;"----");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))
=FILTER(Tabelle1[Geburtsdatum];WOCHENTAG(Tabelle1[Geburtsdatum];1)=Q1;"----")
=LET(x;FILTER(Tabelle1[Geburtsdatum];WOCHENTAG(Tabelle1[Geburtsdatum];1)=Q1;"----");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1))

=VSTAPELN(X1;FILTER(C2:.C999;TEXT(C2:.C999;"tttt")=TEXT(X1;"tttt");"---"))
=VSTAPELN(X1;LET(x;FILTER(C2:.C999;TEXT(C2:.C999;"tttt")=TEXT(X1;"tttt");"---");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1)))
=VSTAPELN(X1;FILTER(Tabelle1[Geburtsdatum];TEXT(Tabelle1[Geburtsdatum];"tttt")=TEXT(X1;"tttt");"---"))
=VSTAPELN(X1;LET(x;FILTER(Tabelle1[Geburtsdatum];TEXT(Tabelle1[Geburtsdatum];"tttt")=TEXT(X1;"tttt");"---");SORTIERENNACH(x;MONAT(x)*100+TAG(x);1)))
        ]]>
    </code>
        )
            }
        }
    End Function

End Module

