Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Tools
Public Class ThisAddIn
    Private _lastCategory As DemoCategory = DemoCategory.Regex
    Private _pane As CustomTaskPane
    Private _paneControl As DemoTaskPaneControl
    Private WithEvents demoWs As Excel.Worksheet
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        _paneControl = New DemoTaskPaneControl()
        _pane = Me.CustomTaskPanes.Add(_paneControl, "Demo Beispiele")
        _pane.Width = 620
        _pane.Visible = False
        LoadLastCategory()
    End Sub
    Private Sub Application_WorkbookOpen(wb As Excel.Workbook)
        CheckForDemoSheet(wb)
    End Sub
    Private Sub Application_NewWorkbook(wb As Excel.Workbook)
        CheckForDemoSheet(wb)
    End Sub
    Private Sub CheckForDemoSheet(wb As Excel.Workbook)
        Try
            For Each sheet As Excel.Worksheet In wb.Worksheets
                If sheet.Name.Equals("Demo_Formel_3", StringComparison.OrdinalIgnoreCase) Then
                    AttachDemoChangeHandler(sheet)
                    Exit For
                End If
            Next
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Fehler bei Blattprüfung: " & ex.Message)
        End Try
    End Sub
    Public Sub AttachDemoChangeHandler(sheet As Excel.Worksheet)
        demoWs = sheet
    End Sub
    Private Sub demoWs_Change(ByVal Target As Excel.Range) Handles demoWs.Change
        Try
            If Target.Address(False, False) = "B1" OrElse Target.Address(False, False) = "B2" Then
                DemoRunner.FormatDemoSheet(demoWs)
                'System.Windows.Forms.MessageBox.Show(
                '    $"Zelle {Target.Address} geändert auf: {Target.Value}",
                '    "Demo_Formel_3 Change"
                ')
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Fehler im Change-Handler: " & ex.Message)
        End Try
    End Sub
    Friend Sub ShowExamples(category As DemoCategory)
        _lastCategory = category

        My.Settings.LastCategory = category.ToString()
        My.Settings.Save()

        _paneControl.SetCategory(category)
        _pane.Visible = True
    End Sub

    Friend Sub ShowLastExamples()
        ShowExamples(_lastCategory)
    End Sub

    Private Sub LoadLastCategory()
        Dim value = My.Settings.LastCategory

        If [Enum].TryParse(value, _lastCategory) = False Then
            _lastCategory = DemoCategory.Regex
        End If
    End Sub
    Friend Sub ShowPaneOnly()
        If _pane IsNot Nothing Then
            _pane.Visible = True
        End If
    End Sub
    Friend Sub HideExamples()
        If _pane IsNot Nothing Then _pane.Visible = False
    End Sub
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        demoWs = Nothing
    End Sub
End Class