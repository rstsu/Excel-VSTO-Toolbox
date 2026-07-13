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
            }
        }
    End Function

End Module

