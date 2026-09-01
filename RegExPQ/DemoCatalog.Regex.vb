Imports System.Xml.Linq

Partial Public Module DemoCatalog

    Private Function GetRegexExamples() As List(Of DemoDefinition)
        Return New List(Of DemoDefinition) From {
            New DemoDefinition With {
                .Id = "regex_001",
                .Category = DemoCategory.Regex,
                .Title = "Text und Zahlen trennen",
                .Tags = {"regex", "zahl", "klammer", "text", "extrahieren"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Text und Zahlen trennen - Zahlen mit und ohne Klammern ausgeben.
Mit 3 Formeln und mit einer Spill-Formel.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=REGEXEXTRAHIEREN(A2:.A999;"^(.+?)\s*\(?\d+\)?$";2)
=REGEXEXTRAHIEREN(A2:.A999;"\(?(\d+)\)?$")
=REGEXEXTRAHIEREN(A2:.A999;"\((\d+)\)";2)

=LET(w;A2:.A999;x;REGEXEXTRAHIEREN(w;"\((\d+)\)";2);y;REGEXEXTRAHIEREN(w;"(\(\d+\))");z;REGEXEXTRAHIEREN(w;"^(.+?)\s*\(\d+\)$";2);HSTAPELN(z;x;y))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_002",
                .Category = DemoCategory.Regex,
                .Title = "Ordner am Backslash kürzen",
                .Tags = {"regex", "ordner", "kürzen", "text", "backslash"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Ordner in Spalte A am Backslash (von rechts) kürzen nach Vorgabe/Länge in Spalte B.
Mit REGEXEXTRAHIEREN oder auch TEXTVOR.
Die unteren beiden Formeln sind mit Überschrift.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=REGEXEXTRAHIEREN(A2:.A999;"^(.*)(?:\\[^\\]+){"&B2:.B100&"}$";2)
=TEXTVOR(A2:.A999;""\"";-B2:.B999)

=VSTAPELN(A1;REGEXEXTRAHIEREN(A2:.A999;"^(.*)(?:\\[^\\]+){"&B2:.B100&"}$";2))
=VSTAPELN(A1;TEXTVOR(A2:.A999;"\";-B2:.B999))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_003",
                .Category = DemoCategory.Regex,
                .Title = "Zollangaben aus Text auslesen",
                .Tags = {"regex", "zoll", "auslesen", "text", "mehrere"},
                .Description = TextBlock(
    <text>
        <![CDATA[
In den Zellen A2:Ax sind Texte mit mehreren Zollangaben (z. B. 1/2" oder dann auch 1 3/4").
Die Zollangaben sollen ausgelesen werden. Auch mehrere.
Überschriften (Z1, Z2...) automatisch generieren.
Die zweite Formel zeigt den Weg, wenn IMMER maximal 3 Zollangaben vorhanden sind.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=LET(x;A2:.A999;p;"(?:\d+\s+)?\d+(?:/\d+)?""";n;MAX(MAP(x;LAMBDA(a;ANZAHL2(REGEXEXTRAHIEREN(a;p;1)))));VSTAPELN("Z"&SEQUENZ(;n);MATRIXERSTELLEN(ZEILEN(x);n;LAMBDA(r;c;WENNFEHLER(INDEX(REGEXEXTRAHIEREN(INDEX(x;r);p;1);c);"")))))

=VSTAPELN({"Z1"."Z2"."Z3"};MATRIXERSTELLEN(ZEILEN(A2:.A999);3;LAMBDA(r;c;WENNFEHLER(INDEX(REGEXEXTRAHIEREN(INDEX(A2:.A999;r);"(?:\d+\s+)?\d+(?:/\d+)?""";1);c);""))))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_004",
                .Category = DemoCategory.Regex,
                .Title = "Sonderzeichen entfernen",
                .Tags = {"regex", "sonderzeichen", "entfernen", "text", "mehrere"},
                .Description = TextBlock(
    <text>
        <![CDATA[
In A2:A8 stehen Texte mit Sonderzeichen - auch mehrere hintereinander.
Diese sollen entfernt werden.
Zwischen den Wörtern darf nur ein Leerzeichen übrig bleiben.
Einmal wird das Pattern direkt in die Formel geschrieben und einmal aus einer Zelle (J2) in die Formel übernommen.

Die dritte und vierte Formel wieder mit Überschrift.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=GLÄTTEN(REGEXERSETZEN(B2:.B999;"[:?*/\\]";""))
=GLÄTTEN(REGEXERSETZEN(B2:.B999;A10;""))

=VSTAPELN(B1;GLÄTTEN(REGEXERSETZEN(B2:.B999;"[:?*/\\]";"")))
=VSTAPELN(B1;GLÄTTEN(REGEXERSETZEN(B2:.B999;A10;"")))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_005",
                .Category = DemoCategory.Regex,
                .Title = "Bestimmte Textteile auslesen",
                .Tags = {"regex", "auslesen", "entfernen", "text", "mehrere"},
                .Description = TextBlock(
    <text>
        <![CDATA[
In A2:A7 stehen Texte in der Art: \NH-DATA\Public\CAD\Project_1\UV_A_1_GL401 A.pro\
Der Bereich zwischen dem vorletzten \ und dem .pro soll ausgelsesen werden.
In dem Fall - UV_A_1_GL401 A - ohne die beiden Zeichen, an denen getrennt wird.
Groß- Kleinschreibung soll keine Rolle spielen.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=VSTAPELN(A1;REGEXEXTRAHIEREN(A2:.A999;"\\([^\\]+)\.pro\\";2;1))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_006",
                .Category = DemoCategory.Regex,
                .Title = "Straße und Hausnummer trennen",
                .Tags = {"regex", "straße", "hausnummer", "trennen", "mehrere"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Straße und Hausnummer in A2:A11.
Diese sollen getrennt werden.
Exoten wie "Strasse des 17. Juni" bleiben komplett bestehen - es sei denn, es gibt eine Hausnummer am Ende.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=LET(v;A2:.A999;w;HSTAPELN("Straße";"Hausnummer");hn;WENNFEHLER(REGEXEXTRAHIEREN(v;"\d+\s*[A-Za-z]?$");"");str;WENN(hn="";v;GLÄTTEN(REGEXERSETZEN(v;"\s*\d+\s*[A-Za-z]?$";"")));VSTAPELN(w;HSTAPELN(str;hn)))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_007",
                .Category = DemoCategory.Regex,
                .Title = "Datumbereich aus String auslesen",
                .Tags = {"regex", "datum", "bereich", "zahlen", "mehrere"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Strings mit Datumbereich in A2:A7 (z. B. Text und 12345 dann12.02.-14.07.2026 Text).
Der Datumbereich soll ausgelesen werden.
Einmal in eine Zelle - und dann noch getrennt in zwei Zellen.
Ausgegeben werden auch die Arbeits- und Kalendertage zwischen den Daten.
Formeln in B1, C1, D1, E1 und H1.
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=VSTAPELN("Datum auslesen";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;"\d{2}\.\d{2}\.?-?\s*-\s*\d{2}\.\d{2}\.\d{4}");"???"))

=VSTAPELN(A1&" - Richtig";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;"\d{2}\.\d{2}(?:\.\d{4})?\.?\s*-\s*(?:\d{2}\.\d{2}|\d{2})\.\d{4}|\d{2}\.\d{2}\.\d{4}|\d{2}\.\d{4}");""))

=VSTAPELN(A1&" - Oder";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;"\d{2}\.\d{2}(?:\.\d{4})?\.?\s*-\s*\d{2}\.\d{2}\.\d{4}");""))

=LET(a;A2:.A999;
    b;WENNFEHLER(
        REGEXEXTRAHIEREN(
            a;
            "\d{2}\.\d{2}(?:\.\d{4})?\.?\s*-\s*(?:\d{2}\.\d{2}|\d{2})\.\d{4}|\d{2}\.\d{2}\.\d{4}|\d{2}\.\d{4}|\d{2}\.\d{2}";
            1
        );
        ""
    );
    c;ISTZAHL(SUCHEN("-";b));
    d;WENN(
        c;
        REGEXEXTRAHIEREN(b;"^\d{2}\.\d{2}(?:\.\d{4})?");
        b
    );
    e;WENN(
        c;
        GLÄTTEN(TEXTNACH(b;"-"));
        ""
    );
    f;LAMBDA(s;g;
        WENN(s="";
            "";
            WENN(
                UND(LÄNGE(s)=5;ISTZAHL(WERT(TEIL(s;1;2))));
                DATUM(g;TEIL(s;4;2);TEIL(s;1;2));
                WENN(
                    LÄNGE(s)=7;
                    DATUM(TEIL(s;4;4);TEIL(s;1;2);1);
                    WENN(
                        LÄNGE(s)=10;
                        DATUM(TEIL(s;7;4);TEIL(s;4;2);TEIL(s;1;2));
                        WENN(
                            LÄNGE(s)=6;
                            DATUM(g;TEIL(s;4;2);TEIL(s;1;2));
                            s
                        )
                    )
                )
            )
        )
    );
    h;f(d;JAHR(HEUTE()));
    i;WENN(
        e<>"";
        LAMBDA(t;
            WENN(
                ODER(LÄNGE(e)=5;LÄNGE(e)=6);
                LET(
                    j;f(e;JAHR(h));
                    WENN(j<h;DATUM(JAHR(j)+1;MONAT(j);TAG(j));j)
                );
                f(e;JAHR(HEUTE()))
            )
        )(e);
        ""
    );
    k;HSTAPELN(h;i);
    WENN(ANZAHL2(k)=0;"";
    VSTAPELN({"Startdatum"."Enddatum"};k)))

=LET(
    a;A2:.A999;
    b;WENNFEHLER(
        REGEXEXTRAHIEREN(
            a;
            "\d{2}\.\d{2}(?:\.\d{4})?\.?\s*-\s*(?:\d{2}\.\d{2}|\d{2})\.\d{4}|\d{2}\.\d{2}\.\d{4}|\d{2}\.\d{4}|\d{2}\.\d{2}";
            1
        );
        ""
    );
    c;ISTZAHL(SUCHEN("-";b));
    d;WENN(
        c;
        REGEXEXTRAHIEREN(b;"^\d{2}\.\d{2}(?:\.\d{4})?");
        b
    );
    e;WENN(
        c;
        GLÄTTEN(TEXTNACH(b;"-"));
        ""
    );
    f; LAMBDA(s;g;
        WENN(s="";
            "";
            WENNFEHLER(
                WENN(
                    LÄNGE(s)=5;
                    DATUM(g;TEIL(s;4;2);TEIL(s;1;2));
                    WENN(
                        LÄNGE(s)=7;
                        DATUM(TEIL(s;4;4);TEIL(s;1;2);1);
                        WENN(
                            LÄNGE(s)=10;
                            DATUM(TEIL(s;7;4);TEIL(s;4;2);TEIL(s;1;2));
                            WENN(
                                LÄNGE(s)=6;
                                DATUM(g;TEIL(s;4;2);TEIL(s;1;2));
                                ""
                            )
                        )
                    )
                );
                ""
            )
        )
    );
    h; f(d;JAHR(HEUTE()));
    i; WENN(
        e<>"";
        LAMBDA(t;
            WENN(
                ODER(LÄNGE(e)=5;LÄNGE(e)=6);
                LET(
                    j;f(e;JAHR(h));
                    WENN(j<h;DATUM(JAHR(j)+1;MONAT(j);TAG(j));j)
                );
                f(e;JAHR(HEUTE()))
            )
        )(e);
        ""
    );
    k; WENN(
        UND(ISTZAHL(h);ISTZAHL(i));
        NETTOARBEITSTAGE(h;i);
        ""
    );
    l; WENN(
        UND(ISTZAHL(h);ISTZAHL(i));
        i-h;
        ""
    );
    m;HSTAPELN(h;i;k;l);VSTAPELN({"Startdatum"."Enddatum"."Arbeitstage"."Tage"};m))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_008",
                .Category = DemoCategory.Regex,
                .Title = "String zwischen 2 Zahlen auslesen",
                .Tags = {"regex", "string", "bereich", "zahlen", "mehrere"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A12) wird ein Text zwischen 2 Zahlem ausgelesen.
Z. B. "1. Vom Bodensee in den Schwarzwald 21 May 2001".
Es wird nur der "Titel" in der Mitte ausgelesen bzw. der Rest ersetzt.

Formel in C1:
Der Titel bleibt übrig, da Anfang (Nummer mit Punkt) und Ende (Datum am Schluss) entfernt werden.

Formel in E1:
Hier wird der Titel gezielt als "Gruppe" herausgenommen und der gesamte Text am Ende damit ersetzt ("$1").
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=VSTAPELN(A1;REGEXERSETZEN(A2:.A999;"^\s*\d+\.\s*|\s+\d{1,2}\s+\S+\s+\d{4}\s*$";""))

=VSTAPELN("Ausgabe";REGEXERSETZEN(A2:.A999;"^\s*\d+\.\s*(.*?)\s+\d{1,2}\s+\S+\s+\d{4}\s*$";"$1"))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_009",
                .Category = DemoCategory.Regex,
                .Title = "String mit Text, Zahl und Datum/Zeit bearbeiten",
                .Tags = {"regex", "string", "datum", "zahlen", "text"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A10) wird ein Text mit Zahlen und Datum/Zeit ausgelesen.

Z. B. "Dateiname Beschreibung Kuerzel 31_05_02_2025_12_13_09.pdf".
Wird zu "Dateiname Beschreibung Kuerzel 31.pdf".

Die Zahl nach "Kuerzel" soll erhalten bleiben - nur Datum und Zeit entfernen.

Formel in C1 und E2 (der Formeltext).
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=VSTAPELN("Name";REGEXERSETZEN(A2:.A999;"_\d{2}_\d{2}_\d{4}_\d{2}_\d{2}_\d{2}";""))

=FORMELTEXT(C1)
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_010",
                .Category = DemoCategory.Regex,
                .Title = "Dateipfade - Name auslesen",
                .Tags = {"regex", "string", "pfad", "datei", "text"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A7) mit Dateipfaden wird der Vor- und Nachname ausgelesen. Verschiedenen Pattern (Herangehensweisen).

Z. B. "https://d.docs.live.net/t56bk74834x63rgu/05 - Verwaltung/01 - Arbeitszeit/Erfassung/Angestellte/2023/[2023 - Max Mustermann - Arbeitszeiten - Dokumentation.xlsx]Januar".
Wird zu "Max Mustermann".

Die RegEx-Patten hier mal angedeutet:
\[ → das öffnende [ (muss maskiert werden)
\d{4} → vierstellige Jahreszahl
- → Trennzeichen
(.+?) → möglichst kurzer Treffer (= der Name)
- Arbeitszeiten → Ende des Namens

\[ → Beginn des Dateinamens
[^-]+ → erster Teil (Jahr oder was auch immer)
(.+?) → der Name
[^]]+ → Beschreibung bis zur schließenden ]
\.xlsx → Dateiendung

(?:.+?) → nicht-speichernde Gruppe für den ersten Teil
(.+?) → einzige Capture-Gruppe = Name
[^]]+ → Rest des Dateinamens bis ]

Weitere Infos zu RegEx und Pattern (dort können die Pattern auch getestet werden):
https://regex101.com/
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=VSTAPELN("Name";REGEXEXTRAHIEREN(A2:.A999;"\[\d{4} - (.+?) - Arbeitszeiten";2))
=VSTAPELN("Name";REGEXEXTRAHIEREN(A2:.A999;"\[.*? - (.+?) - Arbeitszeiten";2))
=VSTAPELN("Name";REGEXEXTRAHIEREN(A2:.A999;"\[\d{4} - (.+?) - .*?\.xlsx";2))
=VSTAPELN("Name";REGEXEXTRAHIEREN(A2:.A999;"\[.*? - (.+?) - .*?\.xlsx";2))
=VSTAPELN("Name";REGEXEXTRAHIEREN(A2:.A999;"\[[^-]+ - (.+?) - [^]]+\.xlsx";2))
=VSTAPELN("Name";REGEXEXTRAHIEREN(A2:.A999;"\[[^-]+ - (.+?) - [^]]+\]";2))
=VSTAPELN("Name";REGEXEXTRAHIEREN(A2:.A999;"\[(?:.+?) - (.+?) - [^]]+\]";2))
=VSTAPELN("Name";REGEXEXTRAHIEREN(A2:.A999;"\[(?:.+?) - (.+?) -";2))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_011",
                .Category = DemoCategory.Regex,
                .Title = "Zahlen - letzte Zahl mit Punkt darstellen",
                .Tags = {"regex", "zahl", "punkt", "zahlen", "spill"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A11) mit Zahlen wird die letzte Zahl mit Punkt angezeigt.

Z. B. "172411".
Wird zu "17241.1".

Weitere Infos zu RegEx und Pattern (dort können die Pattern auch getestet werden):
https://regex101.com/
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=VSTAPELN(A1;REGEXERSETZEN(TEXT(A2:.A999;"0");"(.*)(\d)";"$1.$2"))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_012",
                .Category = DemoCategory.Regex,
                .Title = "Umgang mit Sonderzeichen und Umlauten",
                .Tags = {"regex", "zahl", "umlaute", "zahlen", "sonderzeichen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A7) mit Texten werden die Sonderzeichen getestet und entfernt.
Mit den Funktionen REGEXTESTEN und REGEXERSETZEN.

In Spalte C:E Möglichkeiten mit unterschiedlichen Pattern zum TESTEN der Inhalte von Spalte A.
In Spalte G:H Möglichkeiten mit unterschiedlichen Pattern zum ERSETZEN der Inhalte von Spalte A.

Einmal bleiben die Umlaute vorhanden - dann werden sie entfernt.
Die Formeln werden in die jeweiligen Zellen in die Kommentare geschrieben.
Und ab L1:Lx werden die Formeln in die Zellen - mit z. B. =FORMELTEXT(C1) - geschrieben.

Weitere Infos zu RegEx und Pattern (dort können die Pattern auch getestet werden):
https://regex101.com/
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=VSTAPELN("Test 1";REGEXTESTEN(A2:.A999;"^[äöüßA-Za-z0-9\s-]+$"))
=VSTAPELN("Test 2";REGEXTESTEN(A2:.A999;"[^a-zA-Z0-9\s-]"))
=VSTAPELN("Test 3";REGEXTESTEN(A2:.A999;"^[a-zA-Z0-9\s-]+$"))
=VSTAPELN("Ohne Umlaute";GLÄTTEN(REGEXERSETZEN(A2:.A999;"[^a-zA-Z0-9\s-]";" ")))
=VSTAPELN("Mit Umlaute";GLÄTTEN(REGEXERSETZEN(A2:.A999;"[^a-zA-ZÄÖÜäöüß0-9\s-]";" ")))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_013",
                .Category = DemoCategory.Regex,
                .Title = "Zahl - an bestimmten Stellen ein Minus einfügen",
                .Tags = {"regex", "zahl", "minus", "zahlen", "einfügen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
In einer Liste (A2:A10) mit Zahlen werden an bestimmten Stellen in der Zahl ein Minus eingefügt.
Mit den Funktionen REGEXTESTEN und REGEXERSETZEN.

Also aus z. B. 5467121030655 wird 5467-12-103-0655.

Die Formeln werden in die jeweiligen Zellen in die Kommentare geschrieben.
Und ab K1:Kx werden die Formeln in die Zellen - mit z. B. =FORMELTEXT(C1) - geschrieben.

Weitere Infos zu RegEx und Pattern (dort können die Pattern auch getestet werden):
https://regex101.com/
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
=VSTAPELN("Falsch:";REGEXERSETZEN(A2:.A999;"(\d{4})(\d{2})(\d{3})(\d{4})";"$1-$2-$3-$4"))
=VSTAPELN("Fast richtig:";REGEXERSETZEN(REGEXERSETZEN(A2:.A999;"[^\d]";"");"(\d{4})(\d{2})(\d{3})(\d{4})";"$1-$2-$3-$4"))
=VSTAPELN("Möglichkeit 1:";LET(x;REGEXERSETZEN(A2:.A999;"[^\d]";"");WENN(REGEXTESTEN(x;"^\d{13}$");REGEXERSETZEN(x;"(\d{4})(\d{2})(\d{3})(\d{4})";"$1-$2-$3-$4");"?")))
=VSTAPELN("Möglichkeit 2:";WENN(REGEXTESTEN(A2:.A999;"^\d{13}$");REGEXERSETZEN(A2:.A999;"(\d{4})(\d{2})(\d{3})(\d{4})";"$1-$2-$3-$4");"?"))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_014",
                .Category = DemoCategory.Regex,
                .Title = "Telefonnummern aus Text extrahieren",
                .Tags = {"regex", "zahl", "telefonnummer", "zahlen", "extrahieren"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Mehrere Telefonnummern in zwei Formaten werden aus Spalte A extrahiert.
REGEXEXTRAHIEREN_Telefonnummern_mehrere_aus_Zelle_zu_extrahieren.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\Regex_14

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Formeln in C2, I2, M2.
Pattern in I2/M2 ist besser.

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
=REGEXEXTRAHIEREN(A2;"\(?([\d \-\)\–\+\/\(]{8,})\)?([ .\-–\/]?)([\d]+)";1)
=REGEXEXTRAHIEREN(A2;$H$1;1)
=EINDEUTIG(REGEXEXTRAHIEREN(A2;"(?<!\w)\+?\d(?:[\s()./–-]*\d){7,14}(?!\w)";1);WAHR)
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_015",
                .Category = DemoCategory.Regex,
                .Title = "Materialnummer - wenn in Liste - aus Text auslesen",
                .Tags = {"regex", "zahl", "materialnummer", "zahlen", "auslesen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Materialnummern (Spalte B) werden im Text in Spalte A gesucht und wenn vorhanden ausgelesen.
REGEXEXTRAHIEREN_Zahl_Materialnummer_nur_extrahieren_wenn_in_Liste_vorhanden_Vergleich.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\Regex_15

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Formeln in D2, E2, F2, H2, I1.

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
=NACHZEILE(A2:.A999;LAMBDA(t;WENNFEHLER(VERWEIS(2;1/SUCHEN(B2:.B999;t);B2:.B999);"---")))
=LET(m;B2:.B999;NACHZEILE(A2:.A999;LAMBDA(t;WENNFEHLER(VERWEIS(2;1/SUCHEN(m;t);m);"---"))))
=LET(m;B2:.B999;NACHZEILE(A2:.A999;LAMBDA(t;XVERWEIS(WAHR;REGEXTESTEN(t;"(^|[^0-9])"&m&"([^0-9]|$)");m;"---";0;-1))))
=LET(x;WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;"\d+")*1;NV());y;ISTZAHL(VERGLEICH(x;B2:.B999;0));WENN(y;x;"---"))
=VSTAPELN("Formel 5:";LET(x;WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;"\d+")*1;NV());y;ISTZAHL(VERGLEICH(x;B2:.B999;0));WENN(y;x;"---")))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_016",
                .Category = DemoCategory.Regex,
                .Title = "Großbuchstaben aus String extrahieren",
                .Tags = {"regex", "buchstaben", "sonderzeichen", "klein", "auslesen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus Texten in Spalte A werden nur die Großbuchstaben ausgelesen. Auch Umlaute und das ß werden berücksichtigt.
Formel_REGEXERSETZEN_Nur_Grossbuchstaben_ausgeben_Rest_mit_Nichts_ersetzen.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\Regex_16

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Formeln in C1, D1.

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
=VSTAPELN("Nur Großbuchstaben, Umlaute und ß ausgeben…";REGEXERSETZEN(A2:.A999;"[^A-ZÄÖÜß]";""))
=VSTAPELN("Gib das andere aus…";REGEXERSETZEN(A2:.A999;"[A-ZÄÖÜß]";""))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_017",
                .Category = DemoCategory.Regex,
                .Title = "Letzte Zahl in mehreren Variationen auslesen",
                .Tags = {"regex", "zahl", "punkt", "klammer", "extrahieren"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus Texten in Spalte A werden nur die letzten Zahlen - die in mehreren Variationen vorkommen können - extrahiert.
REGEXEXTRAHIEREN_und_andere_Formeln_letzte_Zahl_mit_oder ohne_Punkt_davor_Optional_in_Klammern.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\Regex_17

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Formeln in C1, D1, E1, F1, G1, H1.

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
=VSTAPELN("REGEX";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;"(?:\(\s*)?([.,]?\d+(?:[.,]\d+)?)(?:\s*\))?\s*$";2);""))
=VSTAPELN("REGEX_1";WENNFEHLER(REGEXEXTRAHIEREN(A2:.A999;"mm\s*\(?\s*([.,]?\d+(?:[.,]\d+)?)\s*\)?\s*$";2);""))
=VSTAPELN("TEXTNACH";GLÄTTEN(WENNFEHLER(LET(x;TEXTNACH(A2:.A999;"mm ";-1);WECHSELN(WECHSELN(x;"(";"");")";""));"")))
=VSTAPELN("TEXTTEILEN";MAP(A2:.A999;LAMBDA(a;WENN(a="";"";LET(x;GLÄTTEN(TEXTNACH(a;"mm"));y;TEXTTEILEN(WECHSELN(WECHSELN(x;"(";"");")";"");" ";;WAHR);WENNFEHLER(INDEX(y;1;1);""))))))
=VSTAPELN("TEXTTEILEN_1";MAP(A2:.A999;LAMBDA(a;WENN(a="";"";LET(x;GLÄTTEN(TEXTNACH(a;"mm"));y;WECHSELN(WECHSELN(x;"(";"");")";"");z;TEXTTEILEN(y;" ";;WAHR);WENNFEHLER(INDEX(z;1);""))))
=VSTAPELN("XMLFILTERN";MAP(A2:.A999;LAMBDA(a;WENN(a="";"";LET(x;GLÄTTEN(TEXTNACH(a;"mm"));y;WECHSELN(WECHSELN(x;"(";"");")";"");WENN(GLÄTTEN(y)="";"";LET(z;WECHSELN(GLÄTTEN(y);" ";"</s><s>_");r;XMLFILTERN("<t><s>_"&z&"</s></t>";"//s[last()]");WECHSELN(r;"_";""))))))))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_018",
                .Category = DemoCategory.Regex,
                .Title = "Datum UTC - MEZ - MESZ...",
                .Tags = {"regex", "zahl", "mez", "mesz", "datum", "utc"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus Daten in Spalte A (UTC) werden die Daten und Zeiten ausgelesen.
UTC_MEZ_MESZ_REGEXEXTRAHIEREN_und_mehr.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\Regex_18

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Formeln in C2, D2, F2, G2, I2, L2 und N2.

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
=REGEXEXTRAHIEREN(A2:.A989;"\d{4}-\d{2}-\d{2}")
=REGEXEXTRAHIEREN(A2:.A989;"\d{2}:\d{2}:\d{2}")
=--REGEXEXTRAHIEREN(A2:.A989;"\d{4}-\d{2}-\d{2}")
=--REGEXEXTRAHIEREN(A2:.A989;"\d{2}:\d{2}:\d{2}")
=--REGEXEXTRAHIEREN(A2;"(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})";2)
=LET(x;--WECHSELN(LINKS(A2;19);"T";" ");j;JAHR(x);SoM;DATUM(j;4;1)-WOCHENTAG(DATUM(j;4;1);2);SoO;DATUM(j;11;1)-WOCHENTAG(DATUM(j;11;1);2);x+WENN(UND(x>=SoM+1/24;x<SoO+1/24);2/24;1/24))
=LET(d;REGEXEXTRAHIEREN(A2;"\d{4}-\d{2}-\d{2}");t;REGEXEXTRAHIEREN(A2;"\d{2}:\d{2}:\d{2}");x;--d+--t;j;JAHR(x);SoM;DATUM(j;4;1)-WOCHENTAG(DATUM(j;4;1);2);SoO;DATUM(j;11;1)-WOCHENTAG(DATUM(j;11;1);2);x+WENN(UND(x>=SoM+1/24;x<SoO+1/24);2/24;1/24))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_019",
                .Category = DemoCategory.Regex,
                .Title = "Name - Alter - Status auslesen",
                .Tags = {"regex", "zahl", "name", "alter", "status", "auslesen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus Daten in Spalte A (z. B. Bruno Graf (von) Bayern (72) - Abwesend) werden Name, Alter und Status ausgelesen.
REGEXEXTRAHIEREN_Name_Alter_Status_auslesen.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\Regex_19

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Formeln in C2, D2, E2, C15, C18, C19, C22, C36, C50 und C64.

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
=REGEXEXTRAHIEREN(A2;"^.*?(?=\s*\(\s*\d+)")
=--REGEXEXTRAHIEREN(A2;"\d+")
=WENNFEHLER(REGEXEXTRAHIEREN(A2;"(?<=[-–]\s).*$");"")
=REGEXEXTRAHIEREN(A2;"^\s*(.*?)\s*\(\s*(\d+)(?:\s+Jahre)?\s*\)\s*(?:[-–]\s*)?(.*)$";2)
=LET(x;REGEXEXTRAHIEREN(A2;"^\s*(.*?)\s*\(\s*(\d+)(?:\s+Jahre)?\s*\)\s*(?:[-–]\s*)?(.*)$";2);HSTAPELN(INDEX(x;1);--INDEX(x;2);INDEX(x;3)))
=LET(x;REGEXEXTRAHIEREN(A2;"^\s*(.*?)\s*\(\s*(\d+)(?:\s+Jahre)?\s*\)\s*(?:[-–]\s*)?(.*)$";2);HSTAPELN(SPALTENWAHL(x;1);--SPALTENWAHL(x;2);SPALTENWAHL(x;3)))
=VSTAPELN(HSTAPELN("Name";"Alter";"Status");LET(a;A2:.A1000;WEGLASSEN(REDUCE("";a;LAMBDA(x;y;VSTAPELN(x;REGEXEXTRAHIEREN(y;"^\s*(.*?)\s*\(\s*(\d+)(?:\s+Jahre)?\s*\)\s*(?:[-–]\s*)?(.*)$";2))));1)))
=LET(x;A2:.A1000;r;WEGLASSEN(REDUCE("";x;LAMBDA(a;z;VSTAPELN(a;REGEXEXTRAHIEREN(z;"^\s*(.*?)\s*\(\s*(\d+)(?:\s+Jahre)?\s*\)\s*(?:[-–]\s*)?(.*)$";2))));1);VSTAPELN(HSTAPELN("Name";"Alter";"Status");HSTAPELN(SPALTENWAHL(r;1);--SPALTENWAHL(r;2);SPALTENWAHL(r;3))))
=LET(x;A2:.A1000;r;WEGLASSEN(REDUCE("";x;LAMBDA(a;z;VSTAPELN(a;REGEXEXTRAHIEREN(z;"^\s*(.*?)\s*\(\s*(\d+)(?:\s+Jahre)?\s*\)\s*(?:[-–]\s*)?(.*)$";2))));1);VSTAPELN({"Name"."Alter"."Status"};HSTAPELN(INDEX(r;;1);--INDEX(r;;2);INDEX(r;;3))))
=LET(x;A2:.A1000;VSTAPELN(HSTAPELN("Name";"Alter";"Status");HSTAPELN(MAP(x;LAMBDA(z;GLÄTTEN(REGEXEXTRAHIEREN(z;"^(.*?)\s*\(\s*\d+(?:\s*Jahre)?\s*\)";2))));MAP(x;LAMBDA(z;--REGEXEXTRAHIEREN(z;"^.*\(\s*(\d+)(?:\s*Jahre)?\s*\)";2)));MAP(x;LAMBDA(z;WENNFEHLER(GLÄTTEN(REGEXEXTRAHIEREN(z;"^.*\)\s*[-–]\s*(.*)$";2));""))))))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_020",
                .Category = DemoCategory.Regex,
                .Title = "Zahlen mit 1.000er Punkt und Text am Ende",
                .Tags = {"regex", "zahl", "punkt", "zahlen", "text", "auslesen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Daten in A1:F8.
In Spalte G werden die Werte aus den jeweiligen Zeilen berechnet.
Aus H (Haben) und S (Soll).

In I1 folgende (Spill-Formel) werden die Daten ausgegeben.
Mit Minus (-) wenn S (Soll) am Ende steht.

1.000er Punkt, Positiv, Negativ und 0 über Benutzerdefiniertes Format.
#.##0,00;-#.##0,00;0
REGEXEXTRAHIEREN_Zahlen_mit_Punkt_und_Text_am_Ende_auslesen_bearbeiten.xlsx

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\Regex_19

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

Formeln in G1 und I1.

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
=VSTAPELN("Ergebnis";NACHZEILE(C2:.F970;LAMBDA(y;LET(z;WENNNV(--REGEXEXTRAHIEREN(y;"\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?");0);e;SUMME(z*WENN(REGEXTESTEN(y;"S\s*$");-1;1));WENN(e=0;0;ABS(e)&" "&WENN(e<0;"S";"H"))))))
=LET(x;C2:.G970;z;WENNNV(--REGEXEXTRAHIEREN(x;"\d+(?:\.\d{3})*(?:,\d+)?");"");VSTAPELN(C1:G1;WENN(z="";"";z*WENN(REGEXTESTEN(x;"S\s*$");-1;1))))
        ]]>
    </code>
        )
            },
            New DemoDefinition With {
                .Id = "regex_021",
                .Category = DemoCategory.Regex,
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
%TEMP%\Excel-VSTO-Toolbox\Demo_RegEx_PQ_Formel_VBA

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
=LET(x;A2:.A999;WENNFEHLER(WEGLASSEN(REDUCE("";x;LAMBDA(A;T;VSTAPELN(A;GLÄTTEN(REGEXEXTRAHIEREN(T;".{1,25}(?=\s|$)";1;1)))));1);""))
=LET(x;A2:.A999;lg;D1;WENNFEHLER(WEGLASSEN(REDUCE("";x;LAMBDA(A;T;VSTAPELN(A;GLÄTTEN(REGEXEXTRAHIEREN(T;".{1,"&lg&"}(?=\s|$)";1;1)))));1);""))
=LET(x;A2:.A999;lg;D1;WENNFEHLER(WEGLASSEN(REDUCE("";x;LAMBDA(A;T;VSTAPELN(A;GLÄTTEN(REGEXEXTRAHIEREN(T;"\S{"&(lg+1)&",}|.{1,"&lg&"}(?=\s|$)";1;1)))));1);""))

=LET(x;A2;y;$C$1;z;"\S{"&(y+1)&",}|.{1,"&y&"}(?=\s|$)";MTRANS(GLÄTTEN(REGEXEXTRAHIEREN(x;z;1;1))))
=LET(x;A2:.A999;lg;C1;WEGLASSEN(REDUCE("";x;LAMBDA(A;z;VSTAPELN(A;MTRANS(GLÄTTEN(REGEXEXTRAHIEREN(z;"\S{"&(lg+1)&",}|.{1,"&lg&"}(?=\s|$)";1;1))))));1))
=LET(x;A2:.A999;y;$C$1;z;"\S{"&(y+1)&",}|.{1,"&y&"}(?=\s|$)";WENNNV(WEGLASSEN(REDUCE("";x;LAMBDA(A;T;HSTAPELN(A;MTRANS(GLÄTTEN(REGEXEXTRAHIEREN(T;z;1;1))))));;1);""))
        ]]>
    </code>
        )
            }
        }
    End Function
End Module

