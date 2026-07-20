Imports Excel = Microsoft.Office.Interop.Excel
Partial Public Class DemoRunner
    Private Sub CreateRegexDemo_1()
        Dim ws = CreateFreshSheet("Demo_Regex_1")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Daten"
            ws.Range("A2").Value = "Artikel (2234)"
            ws.Range("A3").Value = "Der Artikel (12345)"
            ws.Range("A4").Value = "Der tolle Artikel (124599)"
            ws.Range("A5").Value = "Artikel 1 (EUR) (0815)"
            ws.Range("C1").Value = "Text"
            ws.Range("D1").Value = "Zahlen mit Klammern"
            ws.Range("E1").Value = "Zahlen ohne Klammern"
            ws.Range("G1").Value = "Eine Spill-Formel"
            ws.Range("C2").Formula2 = "=REGEXEXTRACT(A2:.A999,""^(.+?)\s*\(?\d+\)?$"",2)"
            ws.Range("D2").Formula2 = "=REGEXEXTRACT(A2:.A999,""\(?(\d+)\)?$"")"
            ws.Range("E2").Formula2 = "=REGEXEXTRACT(A2:.A999,""\((\d+)\)"",2)"
            ws.Range("G2").Formula2 = "=LET(w,A2:.A999,x,REGEXEXTRACT(w,""\((\d+)\)"",2),y,REGEXEXTRACT(w,""(\(\d+\))""),z,REGEXEXTRACT(w,""^(.+?)\s*\(\d+\)$"",2),HSTACK(z,x,y))"
            ws.Range("B8").Value = "Excel-VSTO-Toolbox"
            ws.Range("B9").Value = "RegEx-Demo"
            ws.Range("B10").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("B8:B10").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_2()
        Dim ws = CreateFreshSheet("Demo_Regex_2")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Daten"
            ws.Range("A2").Value = "C:\Users\NutzerA\Desktop\Dokumente\Word\Vorlagen\Beispiele\Ordner1"
            ws.Range("A3").Value = "C:\Users\NutzerA\Desktop\Dokumente\Word\Vorlagen\Beispiele\Ordner2"
            ws.Range("A4").Value = "C:\Users\NutzerA\Desktop\Dokumente\Word\Vorlagen\Beispiele\Ordner3"
            ws.Range("A5").Value = "C:\Users\NutzerA\Desktop\Dokumente\Word\Vorlagen\Beispiele\Ordner4"
            ws.Range("A6").Value = "C:\Users\NutzerA\Desktop\Dokumente\Word\Vorlagen\Beispiele\Ordner5"
            ws.Range("A7").Value = "C:\Users\NutzerA\Desktop\Dokumente\Word\Vorlagen\Beispiele\Ordner6"
            ws.Range("B1").Value = "Abschneiden"
            ws.Range("B2").Value = "3"
            ws.Range("B3").Value = "1"
            ws.Range("B4").Value = "4"
            ws.Range("B5").Value = "2"
            ws.Range("B6").Value = "5"
            ws.Range("B7").Value = "7"
            ws.Range("D1").Value = "Ordner von rechts entfernen - Anzahl aus Spalte B"
            ws.Range("F1").Value = "Mit TEXTVOR"
            ws.Range("D2").Formula2 = "=REGEXEXTRACT(A2:.A999,""^(.*)(?:\\[^\\]+){""&B2:.B100&""}$"",2)"
            ws.Range("F2").Formula2 = "=TEXTBEFORE(A2:.A999,""\"",-B2:.B999)"
            ws.Range("H1").Formula2 = "=VSTACK(A1,REGEXEXTRACT(A2:.A999,""^(.*)(?:\\[^\\]+){""&B2:.B100&""}$"",2))"
            ws.Range("J1").Formula2 = "=VSTACK(A1,TEXTBEFORE(A2:.A999,""\"",-B2:.B999))"
            ws.Range("C10").Value = "Excel-VSTO-Toolbox"
            ws.Range("C11").Value = "RegEx-Demo"
            ws.Range("C12").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("C10:C12").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_3()
        Dim ws = CreateFreshSheet("Demo_Regex_3")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Beschreibung"
            ws.Range("A2").Value = "BUSHING RED. HEXAGONAL MANUFACTURING STD. ASTM A 182 F 316L PIPE-END PREPARATION THREADED NPT DIMENSIONAL STD. ASME B16.11 DN 1/2" & Chr(34) & "  X 1/4" & Chr(34)
            ws.Range("A3").Value = "BUSHING RED. HEXAGONAL MANUFACTURING STD. ASTM A 182 F 316L PIPE-END PREPARATION THREADED NPT DIMENSIONAL STD. ASME B16.11 DN 3/4" & Chr(34) & " X 1/2" & Chr(34)
            ws.Range("A4").Value = "BUSHING RED. HEXAGONAL MANUFACTURING STD. ASTM A 182 F 316L PIPE-END PREPARATION THREADED NPT DIMENSIONAL STD. ASME B16.11 DN 1/2" & Chr(34) & " X 1/4" & Chr(34) & " SCH 10"
            ws.Range("A5").Value = "BUSHING RED. HEXAGONAL MANUFACTURING STD. ASTM A 182 F 316L PIPE-END PREPARATION THREADED NPT DIMENSIONAL STD. ASME B16.11 DN 3/4" & Chr(34) & " X 1/2" & Chr(34) & " SCH 5S"
            ws.Range("A6").Value = "ELBOW R3xD 30° MANUFACTURING STD. ASTM A 403 GRADE WP MATERIAL 304L CLASS W DIMENSIONAL STD. ASME B16.9 PIPE-END PREPARATION ASME B16.25 BW DN 24" & Chr(34) & " THK. 15,88 mm"
            ws.Range("A7").Value = "ELBOW R3xD 45° MANUFACTURING STD. ASTM A 403 GRADE WP MATERIAL 316L CLASS W DIMENSIONAL STD. ASME B16.9 PIPE-END PREPARATION ASME B16.25 BW DN 1/2" & Chr(34) & " SCH 80S"
            ws.Range("A8").Value = "TEE RED. MANUFACTURING STD. ASTM A 182 F 304L CLASS 3000 LB PIPE-END PREPARATION SOCKET WELDING SW DIMENSIONAL STD. ASME B16.11 DN 1" & Chr(34) & " X 1" & Chr(34) & " X 1/2" & Chr(34)
            ws.Range("A9").Value = "TEE RED. MANUFACTURING STD. ASTM A 182 F 304L CLASS 3000 LB PIPE-END PREPARATION SOCKET WELDING SW DIMENSIONAL STD. ASME B16.11 DN 1" & Chr(34) & " X 1" & Chr(34) & " X 1/2" & Chr(34)
            ws.Range("A10").Value = "TEE RED. MANUFACTURING STD. ASTM A 182 F 304L CLASS 3000 LB PIPE-END PREPARATION SOCKET WELDING SW DIMENSIONAL STD. ASME B16.11 DN 1" & Chr(34) & " X 1" & Chr(34) & " X 1/2" & Chr(34) & " SCH 10"
            ws.Range("A11").Value = "TEE RED. MANUFACTURING STD. ASTM A 182 F 304L CLASS 3000 LB PIPE-END PREPARATION SOCKET WELDING SW DIMENSIONAL STD. ASME B16.11 DN 1" & Chr(34) & " X 1" & Chr(34) & " X 1/2" & Chr(34) & " SCH 5S"
            ws.Range("A12").Value = "ELBOW R3xD 30° MANUFACTURING STD. ASTM A 403 GRADE WP MATERIAL 304L CLASS W DIMENSIONAL STD. ASME B16.9 PIPE-END PREPARATION ASME B16.25 BW DN 24" & Chr(34)
            ws.Range("A13").Value = "ELBOW R3xD 45° MANUFACTURING STD. ASTM A 403 GRADE WP MATERIAL 316L CLASS W DIMENSIONAL STD. ASME B16.9 PIPE-END PREPARATION ASME B16.25 BW DN 1/2" & Chr(34)
            ws.Range("A14").Value = "DIMENSIONAL STD. ASME B16.11 DN 1" & Chr(34) & " X 1" & Chr(34) & " X 1 1/2" & Chr(34) & " SCH 10"
            ws.Range("C1").Formula2 = "=LET(x,A2:.A999,p,""(?:\d+\s+)?\d+(?:/\d+)?"""""",n,MAX(MAP(x,LAMBDA(a,COUNTA(REGEXEXTRACT(a,p,1))))),VSTACK(""Z""&SEQUENCE(,n),MAKEARRAY(ROWS(x),n,LAMBDA(r,c,IFERROR(INDEX(REGEXEXTRACT(INDEX(x,r),p,1),c),"""")))))"
            ws.Range("G1").Formula2 = "=VSTACK({""Z1"",""Z2"",""Z3""},MAKEARRAY(ROWS(A2:.A999),3,LAMBDA(r,c,IFERROR(INDEX(REGEXEXTRACT(INDEX(A2:.A999,r),""(?:\d+\s+)?\d+(?:/\d+)?"""""",1),c),""""))))"
            ws.Range("B17").Value = "Excel-VSTO-Toolbox"
            ws.Range("B18").Value = "RegEx-Demo"
            ws.Range("B19").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("B17:B19").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_4()
        Dim ws = CreateFreshSheet("Demo_Regex_4")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Text"
            ws.Range("A2").Value = "1 - Das :ist? ein Test* mit Sonderzeichen /"
            ws.Range("A3").Value = "2 - Das :ist?? ein Test* mit Sonderzeichen /"
            ws.Range("A4").Value = "3 - Das :ist? ein Test*** mit Sonderzeichen /"
            ws.Range("A5").Value = "4 - Das :ist? ein Test* mit Sonderzeichen //"
            ws.Range("A6").Value = "5 - Das ::ist? ein Test* mit Sonderzeichen \\"
            ws.Range("A7").Value = "6 - Das ist? ? ein Test ** mit Sonderzeichen \"
            ws.Range("A8").Value = "7 - Das ist? ?? ? ein Test * * * * mit Sonderzeichen \"
            ws.Range("J1").Formula2 = "Pattern"
            ws.Range("J2").Formula2 = "[:?*/\\]"
            ws.Range("B1").Formula2 = "Text"
            ws.Range("B2").Formula2 = "=TRIM(REGEXREPLACE(A2:.A999,""[:?*/\\]"",""""))"
            ws.Range("D1").Formula2 = "Text"
            ws.Range("D2").Formula2 = "=TRIM(REGEXREPLACE(A2:.A999,J2,""""))"
            ws.Range("F1").Formula2 = "=VSTACK(A1,TRIM(REGEXREPLACE(A2:.A999,""[:?*/\\]"","""")))"
            ws.Range("H1").Formula2 = "=VSTACK(A1,TRIM(REGEXREPLACE(A2:.A999,J2,"""")))"
            ws.Range("B10").Value = "Excel-VSTO-Toolbox"
            ws.Range("B11").Value = "RegEx-Demo"
            ws.Range("B12").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("B10:B12").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_5()
        Dim ws = CreateFreshSheet("Demo_Regex_5")
        ws.Range("A1").Value = "Daten"
        ws.Range("A2").Value = "\\NH-DATA\Public\CAD\Project_1\UV_A_1_GL401 A.pro\"
        ws.Range("A3").Value = "\\NH-DATA\Public\CAD\Project_2\ZG_Z_331_JJKL401 B.pRo\"
        ws.Range("A4").Value = "\\NH-DATA\Public\CAD\Project_3\KOPV_KKO_23451_BR401 ZZH.PRO\"
        ws.Range("A5").Value = "\\NH-DATA\Public\CAD\Project_7\GH_P_1_ZT701 K.pro\"
        ws.Range("A6").Value = "\\NH-DATA\Public\CAD\Project_5\BB_U_546_UHKL6701 ZZ.pRo\"
        ws.Range("A7").Value = "\\NH-DATA\Public\CAD\Project_3\PRAV_PUK_75345_ZG781 BRA.PRO\"
        ws.Range("C1").Formula2 = "=VSTACK(A1,REGEXEXTRACT(A2:.A999,""\\([^\\]+)\.pro\\"",2,1))"
        ws.Range("E1").Value = "Excel-VSTO-Toolbox"
        ws.Range("E2").Value = "RegEx-Demo"
        ws.Range("E3").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
        With ws.Range("E1:E3").Font
            .ColorIndex = 16
            .Size = 8
            .Italic = True
        End With
        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_6()
        Dim ws = CreateFreshSheet("Demo_Regex_6")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Straße und Hausnummer"
            ws.Range("A2").Value = "Friedrich-Ebert-Str.144"
            ws.Range("A3").Value = "Vennhofallee 57"
            ws.Range("A4").Value = "Am Landscheidbühl 12"
            ws.Range("A5").Value = "Am Gänsbühl 13a"
            ws.Range("A6").Value = "Von-Heide-Weg 17 c"
            ws.Range("A7").Value = "Marienstr. 3 a"
            ws.Range("A8").Value = "Strasse des 17. Juni"
            ws.Range("A9").Value = "Strasse des 7. Oktober"
            ws.Range("A10").Value = "Strasse des 17. Juni 8a"
            ws.Range("A11").Value = "Am Grasbödele 44 b"
            ws.Range("C1").Formula2 = "=LET(v,A2:.A999,w,HSTACK(""Straße"",""Hausnummer""),hn,IFERROR(REGEXEXTRACT(v,""\d+\s*[A-Za-z]?$""),""""),str,IF(hn="""",v,TRIM(REGEXREPLACE(v,""\s*\d+\s*[A-Za-z]?$"",""""))),VSTACK(w,HSTACK(str,hn)))"
            ws.Range("F1").Value = "Excel-VSTO-Toolbox"
            ws.Range("F2").Value = "RegEx-Demo"
            ws.Range("F3").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("F1:F3").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_7()
        Dim ws = CreateFreshSheet("Demo_Regex_7")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Daten"
            ws.Range("A2").Value = "Text und 12345 dann12.02.-14.07.2026 Text"
            ws.Range("A3").Value = "Text und 12345 dann12.02 - 14.07.2026 Text"
            ws.Range("A4").Value = "Text und 12345 dann12.02.2026-14.07.2026 Text"
            ws.Range("A5").Value = "Text und 12345 dann12.02.-14.07.2027 Text"
            ws.Range("A6").Value = "Text und 12345 dann12.02.2026-14.07.2027 Text"
            ws.Range("A7").Value = "Text und 12345 dann12.02.2027-14.07.2028 Text"
            ws.Range("B1").Formula2 = "=VSTACK(""Datum auslesen"",IFERROR(REGEXEXTRACT(A2:.A999,""\d{2}\.\d{2}\.?-?\s*-\s*\d{2}\.\d{2}\.\d{4}""),""???""))"
            ws.Range("C1").Formula2 = "=VSTACK(A1&"" - Richtig"",IFERROR(REGEXEXTRACT(A2:.A999,""\d{2}\.\d{2}(?:\.\d{4})?\.?\s*-\s*(?:\d{2}\.\d{2}|\d{2})\.\d{4}|\d{2}\.\d{2}\.\d{4}|\d{2}\.\d{4}""),""""))"
            ws.Range("D1").Formula2 = "=VSTACK(A1&"" - Oder"",IFERROR(REGEXEXTRACT(A2:.A999,""\d{2}\.\d{2}(?:\.\d{4})?\.?\s*-\s*\d{2}\.\d{2}\.\d{4}""),""""))"
            ws.Range("E1").Formula2 = "=LET(a,$A$2:.$A$999,b,IFERROR(REGEXEXTRACT(a,""\d{2}\.\d{2}(?:\.\d{4})?\.?\s*-\s*(?:\d{2}\.\d{2}|\d{2})\.\d{4}|\d{2}\.\d{2}\.\d{4}|d{2}\.\d{4}|\d{2}\.\d{2}"",1),""""),c,ISNUMBER(SEARCH(""-"",b)),d,IF(c,REGEXEXTRACT(b,""^\d{2}\.\d{2}(?:\.\d{4})?""),b),e,IF(c,TRIM(TEXTAFTER(b,""-"")),""""),f,LAMBDA(s,g,IF(s="""","""",IF(AND(LEN(s)=5,ISNUMBER(VALUE(MID(s,1,2)))),DATE(g,MID(s,4,2),MID(s,1,2)),IF(LEN(s)=7,DATE(MID(s,4,4),MID(s,1,2),1),IF(LEN(s)=10,DATE(MID(s,7,4),MID(s,4,2),MID(s,1,2)),IF(LEN(s)=6,DATE(g,MID(s,4,2),MID(s,1,2)),s)))))),h,f(d,YEAR(TODAY())),i,IF(e<>"""",LAMBDA(t,IF(OR(LEN(e)=5,LEN(e)=6),LET(j,f(e,YEAR(h)),IF(j<h,DATE(YEAR(j)+1,MONTH(j),DAY(j)),j)),f(e,YEAR(TODAY()))))(e),""""),k,HSTACK(h,i),IF(COUNTA(k)=0,"""",VSTACK({""Startdatum"",""Enddatum""},k)))"
            ws.Range("E1:F7").NumberFormat = "m/d/yyyy"
            ws.Range("H1").Formula2 = "=LET(a,$A$2:.$A$999,b,IFERROR(REGEXEXTRACT(a,""\d{2}\.\d{2}(?:\.\d{4})?\.?\s*-\s*(?:\d{2}\.\d{2}|\d{2})\.\d{4}|\d{2}\.\d{2}\.\d{4}|d{2}\.\d{4}|\d{2}\.\d{2}"",1),""""),c,ISNUMBER(SEARCH(""-"",b)),d,IF(c,REGEXEXTRACT(b,""^\d{2}\.\d{2}(?:\.\d{4})?""),b),e,IF(c,TRIM(TEXTAFTER(b,""-"")),""""),f,LAMBDA(s,g,IF(s="""","""",IFERROR(IF(LEN(s)=5,DATE(g,MID(s,4,2),MID(s,1,2)),IF(LEN(s)=7,DATE(MID(s,4,4),MID(s,1,2),1),IF(LEN(s)=10,DATE(MID(s,7,4),MID(s,4,2),MID(s,1,2)),IF(LEN(s)=6,DATE(g,MID(s,4,2),MID(s,1,2)),"""")))),""""))),h,f(d,YEAR(TODAY())),i,IF(e<>"""",LAMBDA(t,IF(OR(LEN(e)=5,LEN(e)=6),LET(j,f(e,YEAR(h)),IF(j<h,DATE(YEAR(j)+1,MONTH(j),DAY(j)),j)),f(e,YEAR(TODAY()))))(e),""""),k,IF(AND(ISNUMBER(h),ISNUMBER(i)),NETWORKDAYS(h,i),""""),l,IF(AND(ISNUMBER(h),ISNUMBER(i)),i-h,""""),m,HSTACK(h,i,k,l),VSTACK({""Startdatum"",""Enddatum"",""Arbeitstage"",""Tage""},m))"
            ws.Range("H1:I7").NumberFormat = "m/d/yyyy"
            ws.Range("B10").Value = "Excel-VSTO-Toolbox"
            ws.Range("B11").Value = "RegEx-Demo"
            ws.Range("B12").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("B10:B12").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_8()
        Dim ws = CreateFreshSheet("Demo_Regex_8")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Daten"
            ws.Range("A2").Value = "1. Vom Bodensee in den Schwarzwald 21 May 2001"
            ws.Range("A3").Value = "2. Von Baden-Baden nach Ludwigsburg 28 May 2003"
            ws.Range("A4").Value = "3. Von Schwäbisch Hall nach Stuttgart 4 Jun. 2004"
            ws.Range("A5").Value = "4. Von Heidelberg ins Neckartal 11 Jun. 2005"
            ws.Range("A6").Value = "5. Neckartal und Odenwald 18 Jun. 2001"
            ws.Range("A7").Value = "6. Von Speyer in die Weinpfalz 25 Jun. 2007"
            ws.Range("A8").Value = "7. Von Koblenz zu Mosel und Ahr 2 Jul. 2009"
            ws.Range("A9").Value = "8. Vom Waldecker Land nach Wiesbaden 9 Jul. 2008"
            ws.Range("A10").Value = "9. Von Fulda in den Rheingau 16 Jul. 2002"
            ws.Range("A11").Value = "10. Vom Thüringer Wald nach Erfurt 23 Jul. 2001"
            ws.Range("A12").Value = "11. Rund um München 30 Jul. 2006"
            ws.Range("C1").Formula2 = "=VSTACK(A1,REGEXREPLACE(A2:.A999,""^\s*\d+\.\s*|\s+\d{1,2}\s+\S+\s+\d{4}\s*$"",""""))"
            ws.Range("E1").Formula2 = "=VSTACK(""Ausgabe"",REGEXREPLACE(A2:.A999,""^\s*\d+\.\s*(.*?)\s+\d{1,2}\s+\S+\s+\d{4}\s*$"",""$1""))"
            ws.Range("B15").Value = "Excel-VSTO-Toolbox"
            ws.Range("B16").Value = "Power Query-Demo"
            ws.Range("B17").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("B15:B17").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_9()
        Dim ws = CreateFreshSheet("Demo_Regex_9")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Daten"
            ws.Range("A2").Value = "Dateiname Beschreibung Kuerzel 31_05_02_2025_12_13_09.pdf"
            ws.Range("A3").Value = "Dateiname Beschreibung Kuerzel 14_05_02_2025_12_13_09.pdf"
            ws.Range("A4").Value = "Dateiname Beschreibung Kuerzel 52_05_02_2025_12_13_09.pdf"
            ws.Range("A5").Value = "Dateiname Beschreibung Kuerzel 7_05_02_2025_12_13_09.pdf"
            ws.Range("A6").Value = "Dateiname Beschreibung Kuerzel 47_05_02_2025_12_13_09.pdf"
            ws.Range("A7").Value = "Dateiname Beschreibung Kuerzel 29_05_02_2025_12_13_09.pdf"
            ws.Range("A8").Value = "Dateiname Beschreibung Kuerzel 2_05_02_2025_12_13_09.pdf"
            ws.Range("A9").Value = "Dateiname Beschreibung Kuerzel 86_05_02_2025_12_13_09.pdf"
            ws.Range("A10").Value = "Dateiname Beschreibung Kuerzel 17_05_02_2025_12_13_09.pdf"
            ws.Range("C1").Formula2 = "=VSTACK(""Name"",REGEXREPLACE(A2:.A999,""_\d{2}_\d{2}_\d{4}_\d{2}_\d{2}_\d{2}"",""""))"
            ws.Range("E1").Value = "Formel"
            ws.Range("E2").Formula2 = "=FORMULATEXT(C1)"
            ws.Range("B13").Value = "Excel-VSTO-Toolbox"
            ws.Range("B14").Value = "Power Query-Demo"
            ws.Range("B15").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("B13:B15").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
End Class