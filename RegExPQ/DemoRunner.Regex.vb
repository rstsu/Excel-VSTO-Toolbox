Imports Excel = Microsoft.Office.Interop.Excel
Partial Public Class DemoRunner

    Private Sub CreateRegexDemo_1()
        Dim ws = CreateFreshSheet("Demo_Regex_1")

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

        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_2()
        Dim ws = CreateFreshSheet("Demo_Regex_2")

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

        FormatSheet(ws)
    End Sub

    Private Sub CreateRegexDemo_3()
        Dim ws = CreateFreshSheet("Demo_Regex_3")

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

        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_4()
        Dim ws = CreateFreshSheet("Demo_Regex_4")

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
        FormatSheet(ws)
    End Sub
    Private Sub CreateRegexDemo_6()
        Dim ws = CreateFreshSheet("Demo_Regex_6")
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
        FormatSheet(ws)
    End Sub
End Class