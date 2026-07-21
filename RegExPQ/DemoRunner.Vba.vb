Imports Excel = Microsoft.Office.Interop.Excel
Partial Public Class DemoRunner
    Private Sub CreateVBADemo_4()
        Dim ws = CreateFreshSheet("Demo_VBA_4")
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
            lo.Name = "Demo_VBA_4"
            lo.TableStyle = "TableStyleMedium2"
            ws.Range("C1").Value = "Wert"
            ws.Range("C2").Value = "FRD001"
            ws.Range("C3").Value = "FRD003"
            ws.Range("C4").Value = "FRD005"
            ws.Range("C5").Value = "FRD007"
            ws.Range("C6").Value = "FRD008"
            ws.Range("C7").Value = "FRD009"
            ws.Range("C8").Value = "FRD010"
            ws.Range("C9").Value = "JUG003"
            ws.Range("C10").Value = "JUG005"
            ws.Range("C11").Value = "JUG006"
            ws.Range("C12").Value = "KIM001"
            ws.Range("C13").Value = "STD005"
            ws.Range("C14").Value = "STD006"
            ws.Range("C15").Value = "STD007"
            ws.Range("B18").Value = "Excel-VSTO-Toolbox"
            ws.Range("B19").Value = "VBA-Demo"
            ws.Range("B20").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("B18:B20").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
            With ws.Range("D18:E20")
                Dim fbu = ws.Buttons.Add(.Left, .Top, .Width, .Height)
                fbu.Caption = "Main_1..."
                fbu.OnAction = "Main_1"
            End With
            With ws.Range("D22:E24")
                Dim fbu = ws.Buttons.Add(.Left, .Top, .Width, .Height)
                fbu.Caption = "Main_2..."
                fbu.OnAction = "Main_2"
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateVBADemo_5()
        Dim ws = CreateFreshSheet("Demo_VBA_5")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "URLs"
            ws.Range("A2").Value = "https://www.wikipedia.de"
            ws.Range("A3").Value = "https://www.herber.de/forum"
            ws.Range("A4").Value = "https://github.com"
            ws.Range("A5").Value = "https://albknoblauch.de"
            ws.Range("A6").Value = "https://learn.microsoft.com/de-de/office/dev/scripts/overview/excel"
            ws.Range("A7").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            ws.Range("C1").Value = "Link"
            ws.Range("E1").Value = "Excel-VSTO-Toolbox"
            ws.Range("E2").Value = "VBA-Demo"
            ws.Range("E3").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("E1:E3").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
            With ws.Range("B9:B11")
                Dim fbu = ws.Buttons.Add(.Left, .Top, .Width, .Height)
                fbu.Caption = "Main_1_1..."
                fbu.OnAction = "Main_1_1"
            End With
            With ws.Range("B13:B15")
                Dim fbu = ws.Buttons.Add(.Left, .Top, .Width, .Height)
                fbu.Caption = "Main_2_1..."
                fbu.OnAction = "Main_2_1"
            End With
            With ws.Range("B17:B19")
                Dim fbu = ws.Buttons.Add(.Left, .Top, .Width, .Height)
                fbu.Caption = "Main_3_1..."
                fbu.OnAction = "Main_3_1"
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
End Class