Imports Microsoft.Office.Tools.Ribbon
Public Class Ribbon1
    Private Sub btnRegex_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRegex.Click
        Globals.ThisAddIn.ShowExamples(DemoCategory.Regex)
    End Sub
    Private Sub btnPowerQuery_Click(sender As Object, e As RibbonControlEventArgs) Handles btnPowerQuery.Click
        Globals.ThisAddIn.ShowExamples(DemoCategory.PowerQuery)
    End Sub
    Private Sub btnVba_Click(sender As Object, e As RibbonControlEventArgs) Handles btnVba.Click
        Globals.ThisAddIn.ShowExamples(DemoCategory.Vba)
    End Sub
    Private Sub btnFormulas_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFormulas.Click
        Globals.ThisAddIn.ShowExamples(DemoCategory.Formula)
    End Sub
    Private Sub btnDeleteDemos_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDeleteDemos.Click
        Dim runner As New DemoRunner(Globals.ThisAddIn.Application)
        runner.DeleteDemoSheets()
    End Sub
    Private Sub btnDeleteDemo1_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDeleteDemo1.Click
        Dim runner As New DemoRunner(Globals.ThisAddIn.Application)
        runner.DeleteDemoSheetEins()
    End Sub
    Private Sub btnHidePane_Click(sender As Object, e As RibbonControlEventArgs) Handles btnHidePane.Click
        Globals.ThisAddIn.HideExamples()
    End Sub
    Private Sub btnShowPane_Click(sender As Object, e As RibbonControlEventArgs) Handles btnShowPane.Click
        Globals.ThisAddIn.ShowLastExamples()
    End Sub
End Class