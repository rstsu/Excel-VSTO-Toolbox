Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Partial Public Class DemoRunner
    Private ReadOnly App As Excel.Application
    Public Sub New(application As Excel.Application)
        App = application
    End Sub
    Public Sub CreateDemo(demo As DemoDefinition)
        If demo Is Nothing Then Return
        Select Case demo.Id
            Case "regex_001"
                CreateRegexDemo_1()
            Case "regex_002"
                CreateRegexDemo_2()
            Case "regex_003"
                CreateRegexDemo_3()
            Case "regex_004"
                CreateRegexDemo_4()
            Case "regex_005"
                CreateRegexDemo_5()
            Case "regex_006"
                CreateRegexDemo_6()
            Case "regex_007"
                CreateRegexDemo_7()
            Case "regex_008"
                CreateRegexDemo_8()
            Case "regex_009"
                CreateRegexDemo_9()
            Case "regex_010"
                CreateRegexDemo_10()
            Case "regex_011"
                CreateRegexDemo_11()
            Case "regex_012"
                CreateRegexDemo_12()
            Case "regex_013"
                CreateRegexDemo_13()
            Case "regex_014"
                CreateRegexDemo_14()
            Case "regex_015"
                CreateRegexDemo_15()
            Case "regex_016"
                CreateRegexDemo_16()
            Case "regex_017"
                CreateRegexDemo_17()
            Case "regex_018"
                CreateRegexDemo_18()
            Case "regex_019"
                CreateRegexDemo_19()
            Case "regex_020"
                CreateRegexDemo_20()
            Case "pq_001"
                CreatePowerQueryDemo_1()
            Case "pq_002"
                CreatePowerQueryDemo_2()
            Case "pq_003"
                CreatePowerQueryDemo_3()
            Case "pq_004"
                CreatePowerQueryDemo_4()
            Case "pq_005"
                CreatePowerQueryDemo_5()
            Case "pq_006"
                CreatePowerQueryDemo_6()
            Case "pq_007"
                CreatePowerQueryDemo_7()
            Case "pq_008"
                CreatePowerQueryDemo_8()
            Case "pq_009"
                CreatePowerQueryDemo_9()
            Case "pq_0010"
                CreatePowerQueryDemo_10()
            Case "pq_0011"
                CreatePowerQueryDemo_11()
            Case "pq_0012"
                CreatePowerQueryDemo_12()
            Case "pq_0013"
                CreatePowerQueryDemo_13()
            Case "pq_0014"
                CreatePowerQueryDemo_14()
            Case "pq_0015"
                CreatePowerQueryDemo_15()
            Case "pq_0016"
                CreatePowerQueryDemo_16()
            Case "pq_0017"
                CreatePowerQueryDemo_17()
            Case "pq_0018"
                CreateVBA_PQ_FormelDemo()
            Case "pq_0019"
                CreatePowerQueryDemo_19()
            Case "pq_0020"
                CreatePowerQueryDemo_20()
            Case "formula_001"
                CreateFormulaDemo_1()
            Case "formula_002"
                CreateFormulaDemo_2()
            Case "formula_003"
                CreateFormulaDemo_3()
            Case "formula_004"
                CreateFormulaDemo_4()
            Case "formula_005"
                CreateFormulaDemo_5()
            Case "formula_006"
                CreateFormulaDemo_6()
            Case "formula_007"
                CreateFormelDemo_7()
            Case "formula_008"
                CreateFormelDemo_8()
            Case "formula_009"
                CreateFormelDemo_9()
            Case "formula_0010"
                CreateVBA_PQ_FormelDemo()
            Case "vba_001"
                MessageBox.Show("Dieses VBA-Beispiel bitte über 'Code kopieren' im VBA-Editor in das vorgegebene Modul einfügen.", "VBA Beispiel")
            Case "vba_002"
                MessageBox.Show("Dieses VBA-Beispiel bitte über 'Code kopieren' im VBA-Editor in das vorgegebene Modul einfügen.", "VBA Beispiel")
            Case "vba_003"
                MessageBox.Show("Dieses VBA-Beispiel bitte über 'Code kopieren' im VBA-Editor in das vorgegebene Modul einfügen.", "VBA Beispiel")
            Case "vba_004"
                CreateVBADemo_4()
            Case "vba_005"
                CreateVBADemo_5()
            Case "vba_006"
                CreateVBADemo_6()
            Case "vba_007"
                CreateVBADemo_7()
            Case "vba_008"
                CreateVBADemo_8()
            Case "vba_009"
                CreateVBADemo_9()
            Case "vba_0010"
                CreateVBA_PQ_FormelDemo()
        End Select
    End Sub
    Public Sub DeleteDemoSheets()
        Dim wb = App.ActiveWorkbook
        If wb Is Nothing Then Return
        Dim oldAlerts = App.DisplayAlerts
        App.DisplayAlerts = False
        Try
            For i = wb.Worksheets.Count To 1 Step -1
                Dim ws = CType(wb.Worksheets(i), Excel.Worksheet)

                If ws.Name.StartsWith("Demo_", StringComparison.OrdinalIgnoreCase) Then
                    ws.Delete()
                End If
            Next
        Finally
            App.DisplayAlerts = oldAlerts
        End Try
    End Sub
    Public Sub DeleteDemoSheetEins()
        Dim wb = App.ActiveWorkbook
        If wb Is Nothing Then Return
        Dim oldAlerts = App.DisplayAlerts
        App.DisplayAlerts = False
        Try
            Dim ws = CType(wb.Worksheets(wb.ActiveSheet.Name.ToString), Excel.Worksheet)
            If ws.Name.Equals(wb.ActiveSheet.Name.ToString, StringComparison.OrdinalIgnoreCase) Then
                If ws.Name.StartsWith("Demo_", StringComparison.OrdinalIgnoreCase) Then
                    ws.Delete()
                End If
            End If
        Finally
            App.DisplayAlerts = oldAlerts
        End Try
    End Sub
    Private Function CreateFreshSheet(sheetName As String) As Excel.Worksheet
        Dim wb = App.ActiveWorkbook
        If wb Is Nothing Then
            wb = App.Workbooks.Add()
        End If
        DeleteSheetIfExists(wb, sheetName)
        Dim ws = CType(
        wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)),
        Excel.Worksheet)
        ws.Name = sheetName
        Return ws
    End Function
    Private Sub DeleteSheetIfExists(wb As Excel.Workbook, sheetName As String)
        Dim oldAlerts = App.DisplayAlerts
        App.DisplayAlerts = False
        Try
            For Each ws As Excel.Worksheet In wb.Worksheets
                If ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase) Then
                    ws.Delete()
                    Exit For
                End If
            Next
        Finally
            App.DisplayAlerts = oldAlerts
        End Try
    End Sub
    Private Sub FormatSheet(ws As Excel.Worksheet)
        Dim used = ws.UsedRange
        If used IsNot Nothing Then
            used.Rows(1).Font.Bold = True
            used.Columns.AutoFit()
        End If
    End Sub
    Public Shared Sub FormatDemoSheet(ws As Excel.Worksheet)
        Dim app As Excel.Application = ws.Application
        Dim adr As New List(Of String) From {
        "A4:G9",
        "A12:G18",
        "A21:G28"
    }
        Try
            app.ScreenUpdating = False
            Dim rng As Excel.Range = ws.Range(String.Join(";", adr))
            rng.NumberFormat = "m/d/yyyy"
            ws.Range("I2:I12").NumberFormat = "m/d/yyyy"
            Dim aktuellerMonat As Integer = CInt(ws.Range("B2").Value)
            ' Feiertage aus I2:J12 in Dictionary laden (Datum -> Name)
            Dim feiertage As New Dictionary(Of Date, String)
            For Each row As Excel.Range In ws.Range("I2:I12").Rows
                Dim datCell As Excel.Range = row.Cells(1, 1) ' Spalte I
                Dim nameCell As Excel.Range = row.Offset(0, 1) ' Spalte J
                If datCell.Value2 IsNot Nothing AndAlso TypeOf datCell.Value2 Is Double Then
                    Dim dat As Date = DateTime.FromOADate(CDbl(datCell.Value2))
                    Dim name As String = ""
                    If nameCell.Value2 IsNot Nothing Then
                        name = CStr(nameCell.Value2)
                    End If
                    If Not feiertage.ContainsKey(dat) Then
                        feiertage.Add(dat, name)
                    End If
                End If
            Next
            ' Farben definieren
            Dim grau As Long = 10921638 ' Hellgrau
            Dim hellgruen As Long = RGB(198, 239, 206)
            Dim hellblau As Long = RGB(189, 215, 238)
            ' Über alle Zellen iterieren
            For Each c As Excel.Range In rng.Cells
                If c.Value2 IsNot Nothing AndAlso TypeOf c.Value2 Is Double Then
                    Dim datum As Date = DateTime.FromOADate(CDbl(c.Value2))
                    ' Standardformatierung zurücksetzen
                    c.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
                    c.Font.Italic = False
                    c.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone
                    If c.Comment IsNot Nothing Then c.Comment.Delete()
                    ' 1️⃣ Feiertage prüfen
                    If feiertage.ContainsKey(datum) Then
                        c.Interior.Color = hellblau
                        Dim feiertagsName As String = feiertage(datum)
                        If Not String.IsNullOrWhiteSpace(feiertagsName) Then
                            c.AddComment(feiertagsName)
                        End If
                        ' 2️⃣ Tage außerhalb des Monats → grau + kursiv
                    ElseIf datum.Month <> aktuellerMonat Then
                        c.Font.Color = grau
                        c.Font.Italic = True
                        c.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone
                        ' 3️⃣ Wochenenden im aktuellen Monat → hellgrün
                    ElseIf c.Column = 6 OrElse c.Column = 7 Then
                        c.Interior.Color = hellgruen
                    End If
                End If
            Next
            ' Überschriften fett + Spaltenbreite anpassen
            Dim used = ws.UsedRange
            If used IsNot Nothing Then
                used.Rows(1).Font.Bold = True
                used.Columns.AutoFit()
            End If
        Finally
            app.ScreenUpdating = True
        End Try
    End Sub
    Public Sub SetupYearMonthValidation(ws As Excel.Worksheet)
        Dim currentYear As Integer = DateTime.Now.Year
        Dim currentMonth As Integer = DateTime.Now.Month
        Dim app As Excel.Application = ws.Application
        ' --- B1: Jahr ---
        Dim yearList As New List(Of String)
        For y As Integer = currentYear - 5 To currentYear + 10
            yearList.Add(y.ToString())
        Next
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            With ws.Range("B1").Validation
                .Delete()
                .Add(Excel.XlDVType.xlValidateList,
                 Excel.XlDVAlertStyle.xlValidAlertStop,
                 Excel.XlFormatConditionOperator.xlBetween,
                 String.Join(";", yearList))
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
                .InputTitle = "Jahr auswählen"
                .InputMessage = "Bitte ein Jahr aus der Liste wählen"
                .ErrorTitle = "Ungültige Eingabe"
                .ErrorMessage = "Nur Jahre aus der Liste sind erlaubt."
            End With
            ws.Range("B1").Value = currentYear
            ' --- B2: Monat ---
            Dim monthList As New List(Of String)
            For m As Integer = 1 To 12
                monthList.Add(m.ToString())
            Next
            With ws.Range("B2").Validation
                .Delete()
                .Add(Excel.XlDVType.xlValidateList,
                 Excel.XlDVAlertStyle.xlValidAlertStop,
                 Excel.XlFormatConditionOperator.xlBetween,
                 String.Join(";", monthList))
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
                .InputTitle = "Monat auswählen"
                .InputMessage = "Bitte Monat (1-12) wählen"
                .ErrorTitle = "Ungültige Eingabe"
                .ErrorMessage = "Nur Werte von 1 bis 12 sind erlaubt."
            End With
            ws.Range("B2").Value = currentMonth
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
    End Sub
    Private Sub CreatePowerQueryDemo_14()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_PQ_14.zip",
                    "PQ_014")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Power-Query-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo PQ 014",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreatePowerQueryDemo_15()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_PQ_15.zip",
                    "PQ_015")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Power-Query-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo PQ 015",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreatePowerQueryDemo_16()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_PQ_16.zip",
                    "PQ_016")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Power-Query-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo PQ 016",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreatePowerQueryDemo_17()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_PQ_17.zip",
                    "PQ_017")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Power-Query-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo PQ 017",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreatePowerQueryDemo_19()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_PQ_19.zip",
                    "PQ_019")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Power-Query-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo PQ 019",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreatePowerQueryDemo_20()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_PQ_20.zip",
                    "PQ_020")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Power-Query-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo PQ 020",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateVBADemo_7()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_VBA_7.zip",
                    "VBA_07")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das VBA-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo VBA 07",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateVBADemo_8()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_VBA_8.zip",
                    "VBA_08")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das VBA-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo VBA 08",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateVBADemo_9()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_VBA_9.zip",
                    "VBA_09")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das VBA-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo VBA 09",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateFormelDemo_7()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Formel_7.zip",
                    "Formel_07")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Formel-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Formel 07",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateFormelDemo_8()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Formel_8.zip",
                    "Formel_08")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Formel-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Formel 08",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateFormelDemo_9()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Formel_9.zip",
                    "Formel_09")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Formel-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Formel 09",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateRegexDemo_14()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Regex_14.zip",
                    "Regex_14")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Regex-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Regex 14",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateRegexDemo_15()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Regex_15.zip",
                    "Regex_15")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Regex-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Regex 15",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateRegexDemo_16()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Regex_16.zip",
                    "Regex_16")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Regex-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Regex 16",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateRegexDemo_17()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Regex_17.zip",
                    "Regex_17")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Regex-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Regex 17",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateRegexDemo_18()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Regex_18.zip",
                    "Regex_18")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Regex-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Regex 18",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateRegexDemo_19()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Regex_19.zip",
                    "Regex_19")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Regex-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Regex 19",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateRegexDemo_20()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_Regex_20.zip",
                    "Regex_20")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Regex-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo Regex 20",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CreateVBA_PQ_FormelDemo()
        Try
            Dim workbook As Excel.Workbook =
                DemoPackageManager.ExtractAndOpen(
                    App,
                    "Demo_VBA_PQ_Formel.zip",
                    "Demo_VBA_PQ_Formel")
            workbook.Activate()
        Catch ex As Exception
            MessageBox.Show(
                "Das Power-Query-Beispiel konnte nicht geöffnet werden." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                "Demo_VBA_PQ_Formel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
End Class