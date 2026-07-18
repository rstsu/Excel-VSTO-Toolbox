Imports Excel = Microsoft.Office.Interop.Excel
Partial Public Class DemoRunner
    Private Sub CreatePowerQueryDemo_1()
        Dim ws = CreateFreshSheet("Demo_PQ_1")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            Dim data(,) As Object = {
            {"Nr", "Wörter"},
            {12345, "Test"},
            {12358, "Test2"},
            {45784, "Hilfe"},
            {3456, "Hilfe2"},
            {12345, "Excel"},
            {456784, "Excel2"},
            {12345, "Internet"},
            {3456, "Microsoft"}
        }
            ws.Range("A1:B9").Value = data
            Dim lo = ws.ListObjects.Add(
                Excel.XlListObjectSourceType.xlSrcRange,
                ws.Range("A1:B9"),, Excel.XlYesNoGuess.xlYes)
            lo.Name = "Demo_PQ_1"
            lo.TableStyle = "TableStyleMedium2"
            ws.Range("A11").Value = "Excel-VSTO-Toolbox"
            ws.Range("A12").Value = "Power Query-Demo"
            ws.Range("A13").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("A11:A13").Font
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
    Private Sub CreatePowerQueryDemo_2()
        Dim ws = CreateFreshSheet("Demo_PQ_2")
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
            Dim lo = ws.ListObjects.Add(
                Excel.XlListObjectSourceType.xlSrcRange,
                ws.Range("A1:A14"),, Excel.XlYesNoGuess.xlYes)
            lo.Name = "Demo_PQ_2"
            lo.TableStyle = "TableStyleMedium2"
            ws.Range("A17").Value = "Excel-VSTO-Toolbox"
            ws.Range("A18").Value = "Power Query-Demo"
            ws.Range("A19").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("A17:A19").Font
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
    Private Sub CreatePowerQueryDemo_3()
        Dim ws = CreateFreshSheet("Demo_PQ_3")
        Dim App As Excel.Application = ws.Application
        Try
            App.ScreenUpdating = False
            App.EnableEvents = False
            ws.Range("G1").Formula2 = "=TEXTBEFORE(CELL(""filename""),""\"",3)&""\AppData\Local\Temp"""
            ws.Parent.Names.Add(Name:="Benutzername", RefersTo:="=Demo_PQ_3!G1")
            ws.Range("I1").Value = "Excel-VSTO-Toolbox"
            ws.Range("I2").Value = "Power Query-Demo"
            ws.Range("I3").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("I1:I3").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            App.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreatePowerQueryDemo_4()
        Dim ws = CreateFreshSheet("Demo_PQ_4")
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
            Dim lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, ws.Range("A1:A8"),, Excel.XlYesNoGuess.xlYes)
            lo.Name = "Demo_PQ_4"
            lo.TableStyle = "TableStyleMedium2"
            ws.Range("A11").Value = "Excel-VSTO-Toolbox"
            ws.Range("A12").Value = "Power Query-Demo"
            ws.Range("A13").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("A11:A13").Font
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
    Private Sub CreatePowerQueryDemo_5()
        Dim ws = CreateFreshSheet("Demo_PQ_5")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Wert"
            ws.Range("A2").Value = "FRD001"
            ws.Range("A3").Value = "FRD003"
            ws.Range("A4").Value = "FRD005"
            ws.Range("A5").Value = "FRD007"
            ws.Range("A6").Value = "FRD008"
            ws.Range("A7").Value = "FRD009"
            ws.Range("A8").Value = "FRD010"
            ws.Range("A9").Value = "JUG003"
            ws.Range("A10").Value = "JUG005"
            ws.Range("A11").Value = "JUG006"
            ws.Range("A12").Value = "KIM001"
            ws.Range("A13").Value = "STD005"
            ws.Range("A14").Value = "STD006"
            ws.Range("A15").Value = "STD007"
            Dim lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, ws.Range("A1:A15"),, Excel.XlYesNoGuess.xlYes)
            lo.Name = "Demo_PQ_5"
            lo.TableStyle = "TableStyleMedium2"
            ws.Range("B18").Value = "Excel-VSTO-Toolbox"
            ws.Range("B19").Value = "Power Query-Demo"
            ws.Range("B20").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("B18:B20").Font
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
    Private Sub CreatePowerQueryDemo_6()
        Dim ws = CreateFreshSheet("Demo_PQ_6")
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
            Dim lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, ws.Range("A1:A12"),, Excel.XlYesNoGuess.xlYes)
            lo.Name = "Demo_PQ_6"
            lo.TableStyle = "TableStyleMedium2"
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
End Class