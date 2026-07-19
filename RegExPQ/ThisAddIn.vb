Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Collections.Generic
Imports Microsoft.Office.Tools
Public Class ThisAddIn
    Private _lastCategory As DemoCategory = DemoCategory.Regex
    Private ReadOnly _panes As New Dictionary(Of Integer, CustomTaskPane)
    Private ReadOnly _paneControls As New Dictionary(Of Integer, DemoTaskPaneControl)
    Private WithEvents demoWs As Excel.Worksheet
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        LoadLastCategory()
        If Me.Application.ActiveWindow IsNot Nothing Then
            CreatePane(Me.Application.ActiveWindow)
        End If
    End Sub
    Private Sub CreatePane(window As Excel.Window)
        Dim hwnd As Integer = window.Hwnd
        If _panes.ContainsKey(hwnd) Then Return
        Dim control As New DemoTaskPaneControl()
        Dim pane = Me.CustomTaskPanes.Add(
        control,
        "Demo Beispiele", window)
        pane.Width = 620
        pane.Visible = False
        _panes.Add(hwnd, pane)
        _paneControls.Add(hwnd, control)
    End Sub
    Private Sub Application_WindowActivate(
    wb As Excel.Workbook,
    Wn As Excel.Window) Handles Application.WindowActivate
        CreatePane(Wn)
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
            If Target.Cells.CountLarge = 1 Then
                If Target.Address(False, False) = "B1" OrElse Target.Address(False, False) = "B2" Then
                    DemoRunner.FormatDemoSheet(demoWs)
                End If
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Fehler im Change-Handler: " & ex.Message)
        End Try
    End Sub
    Private Sub Application_WorkbookBeforeClose(
    wb As Excel.Workbook,
    ByRef Cancel As Boolean) Handles Application.WorkbookBeforeClose
        If demoWs IsNot Nothing Then
            Try
                If demoWs.Parent Is wb Then
                    demoWs = Nothing
                End If
            Catch
                demoWs = Nothing
            End Try
        End If
    End Sub
    Friend Sub ShowExamples(category As DemoCategory)
        _lastCategory = category
        My.Settings.LastCategory = category.ToString()
        My.Settings.Save()
        Dim wnd = Application.ActiveWindow
        If wnd Is Nothing Then Return
        Dim hwnd As Integer = wnd.Hwnd
        CreatePane(wnd)
        _paneControls(hwnd).SetCategory(category)
        _panes(hwnd).Visible = True
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
        Dim wnd = Application.ActiveWindow
        If wnd Is Nothing Then Return
        Dim hwnd As Integer = wnd.Hwnd
        CreatePane(wnd)
        _panes(hwnd).Visible = True
    End Sub
    Friend Sub HideExamples()
        Dim wnd = Application.ActiveWindow
        If wnd Is Nothing Then Return
        Dim hwnd As Integer = wnd.Hwnd
        If _panes.ContainsKey(hwnd) Then
            _panes(hwnd).Visible = False
        End If
    End Sub
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        demoWs = Nothing
        _panes.Clear()
        _paneControls.Clear()
    End Sub
End Class