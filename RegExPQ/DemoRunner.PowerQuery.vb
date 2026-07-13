Imports Excel = Microsoft.Office.Interop.Excel
Partial Public Class DemoRunner
    Private Sub CreatePowerQueryDemo_1()
        Dim ws = CreateFreshSheet("Demo_PQ_1")

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
        FormatSheet(ws)
    End Sub

    Private Sub CreatePowerQueryDemo_2()
        Dim ws = CreateFreshSheet("Demo_PQ_2")

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

        FormatSheet(ws)
    End Sub
    Private Sub CreatePowerQueryDemo_3()
        Dim ws = CreateFreshSheet("Demo_PQ_3")

        ws.Range("G1").Formula2 = "=TEXTBEFORE(CELL(""filename""),""\"",3)&""\AppData\Local\Temp"""
        ws.Parent.Names.Add(Name:="Benutzername", RefersTo:="=Demo_PQ_3!G1")
        FormatSheet(ws)
    End Sub
    Private Sub CreatePowerQueryDemo_4()
        Dim ws = CreateFreshSheet("Demo_PQ_4")

        ws.Range("A1").Value = "Text"
        ws.Range("A2").Value = "1 - Das :ist? ein Test* mit Sonderzeichen /"
        ws.Range("A3").Value = "2 - Das :ist?? ein Test* mit Sonderzeichen /"
        ws.Range("A4").Value = "3 - Das :ist? ein Test*** mit Sonderzeichen /"
        ws.Range("A5").Value = "4 - Das :ist? ein Test* mit Sonderzeichen //"
        ws.Range("A6").Value = "5 - Das ::ist? ein Test* mit Sonderzeichen \\"
        ws.Range("A7").Value = "6 - Das ist? ? ein Test ** mit Sonderzeichen \"
        ws.Range("A8").Value = "7 - Das ist? ?? ? ein Test * * * * mit Sonderzeichen \"

        Dim lo = ws.ListObjects.Add(
            Excel.XlListObjectSourceType.xlSrcRange,
            ws.Range("A1:A8"),, Excel.XlYesNoGuess.xlYes)

        lo.Name = "Demo_PQ_4"
        lo.TableStyle = "TableStyleMedium2"

        FormatSheet(ws)
    End Sub
End Class