Imports System.Linq
Imports System.Windows.Forms

Public Class DemoTaskPaneControl
    Inherits UserControl

    Private _items As List(Of DemoDefinition)

    Private ReadOnly lblTitle As New Label()
    Private ReadOnly txtSearch As New TextBox()
    Private ReadOnly lstExamples As New ListBox()
    Private ReadOnly txtDescription As New TextBox()
    Private ReadOnly txtCode As New TextBox()
    Private ReadOnly btnCreate As New Button()
    Private ReadOnly btnCopy As New Button()
    Private ReadOnly btnPowerQueryInfo As New Button()

    Public Sub New()
        Me.Padding = New Padding(10)

        lblTitle.Dock = DockStyle.Top
        lblTitle.Height = 30
        lblTitle.Font = New Drawing.Font("Segoe UI", 12, Drawing.FontStyle.Bold)
        lblTitle.Text = "Demo Beispiele"

        txtSearch.Dock = DockStyle.Top
        txtSearch.Height = 26
        txtSearch.Font = New Drawing.Font("Segoe UI", 16, Drawing.FontStyle.Bold)

        lstExamples.Dock = DockStyle.Top
        lstExamples.Height = 120
        lstExamples.DisplayMember = "Title"
        lstExamples.Font = New Drawing.Font("Segoe UI", 12)

        txtDescription.Dock = DockStyle.Top
        txtDescription.Height = 240
        txtDescription.Multiline = True
        txtDescription.ReadOnly = True
        txtDescription.ScrollBars = ScrollBars.Vertical
        txtDescription.Font = New Drawing.Font("Segoe UI", 12)

        txtCode.Dock = DockStyle.Fill
        txtCode.Multiline = True
        txtCode.ReadOnly = True
        txtCode.WordWrap = False
        txtCode.ScrollBars = ScrollBars.Both
        txtCode.Font = New Drawing.Font("Consolas", 12)

        Dim panel As New FlowLayoutPanel With {
            .Dock = DockStyle.Bottom,
            .Height = 42
        }

        btnCreate.Text = "Demo erzeugen"
        btnCreate.Width = 120

        btnCopy.Text = "Code kopieren"
        btnCopy.Width = 120

        btnPowerQueryInfo.Text = "PQ M-Code Info"
        btnPowerQueryInfo.Width = 120

        panel.Controls.Add(btnCreate)
        panel.Controls.Add(btnCopy)
        panel.Controls.Add(btnPowerQueryInfo)

        Controls.Add(txtCode)
        Controls.Add(panel)
        Controls.Add(txtDescription)
        Controls.Add(lstExamples)
        Controls.Add(txtSearch)
        Controls.Add(lblTitle)

        AddHandler btnPowerQueryInfo.Click, AddressOf ShowPowerQueryInfo
        AddHandler txtSearch.TextChanged, AddressOf FilterList
        AddHandler lstExamples.SelectedIndexChanged, AddressOf ShowSelected
        AddHandler btnCreate.Click, AddressOf CreateSelectedDemo
        AddHandler btnCopy.Click, AddressOf CopyCode
    End Sub

    Public Sub SetCategory(category As DemoCategory)
        lblTitle.Text = GetCategoryTitle(category)
        txtSearch.Text = ""

        _items = DemoCatalog.GetByCategory(category)
        BindList(_items)
    End Sub

    Private Sub FilterList(sender As Object, e As EventArgs)
        If _items Is Nothing Then Return

        Dim term = txtSearch.Text.Trim().ToLowerInvariant()

        If term = "" Then
            BindList(_items)
            Return
        End If

        Dim filtered = _items.Where(Function(x)
                                        Dim tags = If(x.Tags Is Nothing, "", String.Join(" ", x.Tags))
                                        Dim text = (x.Title & " " & x.Description & " " & tags).ToLowerInvariant()
                                        Return text.Contains(term)
                                    End Function).ToList()

        BindList(filtered)
    End Sub

    Private Sub BindList(items As List(Of DemoDefinition))
        lstExamples.DataSource = Nothing
        lstExamples.DataSource = items
        lstExamples.DisplayMember = "Title"

        If items.Count > 0 Then
            lstExamples.SelectedIndex = 0
        Else
            txtDescription.Text = ""
            txtCode.Text = ""
        End If
    End Sub

    Private Sub ShowSelected(sender As Object, e As EventArgs)
        Dim demo = TryCast(lstExamples.SelectedItem, DemoDefinition)
        If demo Is Nothing Then Return

        txtDescription.Text = demo.Description
        txtCode.Text = demo.CodeText

        btnPowerQueryInfo.Visible = (demo.Category = DemoCategory.PowerQuery)
        btnCopy.Visible = (demo.Category <> DemoCategory.Formula AndAlso demo.Category <> DemoCategory.Regex)
    End Sub

    Private Sub ShowPowerQueryInfo(sender As Object, e As EventArgs)
        txtDescription.Text = PowerQueryInfoText()
    End Sub

    Private Sub CreateSelectedDemo(sender As Object, e As EventArgs)
        Dim demo = TryCast(lstExamples.SelectedItem, DemoDefinition)
        If demo Is Nothing Then Return
        Dim runner As New DemoRunner(Globals.ThisAddIn.Application)
        runner.CreateDemo(demo)
    End Sub

    Private Sub CopyCode(sender As Object, e As EventArgs)
        If txtCode.Text.Trim() <> "" Then
            Clipboard.SetText(txtCode.Text, TextDataFormat.Text)
        End If
    End Sub

    Private Function GetCategoryTitle(category As DemoCategory) As String
        Select Case category
            Case DemoCategory.Regex
                Return "REGEXEXTRAHIEREN / REGEXERSETZEN / REGEXTESTEN"
            Case DemoCategory.PowerQuery
                Return "Power Query M-Code Beispiele"
            Case DemoCategory.Vba
                Return "VBA Beispiele"
            Case DemoCategory.Formula
                Return "Formel Beispiele"
            Case Else
                Return "Demo Beispiele"
        End Select
    End Function

End Class