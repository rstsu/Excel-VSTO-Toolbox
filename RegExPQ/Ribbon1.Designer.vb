Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Erforderlich für die Unterstützung des Windows.Forms-Klassenkompositions-Designers
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Dieser Aufruf ist für den Komponenten-Designer erforderlich.
        InitializeComponent()

    End Sub

    'Die Komponente überschreibt den Löschvorgang zum Bereinigen der Komponentenliste.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Komponenten-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Komponenten-Designer erforderlich.
    'Das Bearbeiten ist mit dem Komponenten-Designer möglich.
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Demo = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.btnPowerQuery = Me.Factory.CreateRibbonButton
        Me.btnRegex = Me.Factory.CreateRibbonButton
        Me.btnVba = Me.Factory.CreateRibbonButton
        Me.btnFormulas = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btnDeleteDemos = Me.Factory.CreateRibbonButton
        Me.btnDeleteDemo1 = Me.Factory.CreateRibbonButton
        Me.btnHidePane = Me.Factory.CreateRibbonButton
        Me.btnShowPane = Me.Factory.CreateRibbonButton
        Me.Demo.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Demo
        '
        Me.Demo.Groups.Add(Me.Group1)
        Me.Demo.Groups.Add(Me.Group2)
        Me.Demo.Label = "Demo Beispiele..."
        Me.Demo.Name = "Demo"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btnPowerQuery)
        Me.Group1.Items.Add(Me.btnRegex)
        Me.Group1.Items.Add(Me.btnVba)
        Me.Group1.Items.Add(Me.btnFormulas)
        Me.Group1.Label = "Demos..."
        Me.Group1.Name = "Group1"
        '
        'btnPowerQuery
        '
        Me.btnPowerQuery.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnPowerQuery.Label = "Power Query"
        Me.btnPowerQuery.Name = "btnPowerQuery"
        Me.btnPowerQuery.OfficeImageId = "GetExternalDataFromOtherSources"
        Me.btnPowerQuery.ScreenTip = "Power Query Demo Beispiele"
        Me.btnPowerQuery.ShowImage = True
        Me.btnPowerQuery.SuperTip = "Katalog öffnen um Beispiele zu sehen..."
        '
        'btnRegex
        '
        Me.btnRegex.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnRegex.Label = "Regex"
        Me.btnRegex.Name = "btnRegex"
        Me.btnRegex.OfficeImageId = "FindDialog"
        Me.btnRegex.ScreenTip = "RegEx Demo Beispiele"
        Me.btnRegex.ShowImage = True
        Me.btnRegex.SuperTip = "Katalog öffnen um Beispiele zu sehen..."
        '
        'btnVba
        '
        Me.btnVba.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnVba.Label = "VBA"
        Me.btnVba.Name = "btnVba"
        Me.btnVba.OfficeImageId = "VisualBasic"
        Me.btnVba.ScreenTip = "VBA Demo Beispiele"
        Me.btnVba.ShowImage = True
        Me.btnVba.SuperTip = "Katalog öffnen um Beispiele zu sehen..."
        '
        'btnFormulas
        '
        Me.btnFormulas.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnFormulas.Label = "Formeln"
        Me.btnFormulas.Name = "btnFormulas"
        Me.btnFormulas.OfficeImageId = "FunctionWizard"
        Me.btnFormulas.ScreenTip = "Formel Demo Beispiele"
        Me.btnFormulas.ShowImage = True
        Me.btnFormulas.SuperTip = "Katalog öffnen um Beispiele zu sehen..."
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnDeleteDemos)
        Me.Group2.Items.Add(Me.btnDeleteDemo1)
        Me.Group2.Items.Add(Me.btnHidePane)
        Me.Group2.Items.Add(Me.btnShowPane)
        Me.Group2.Label = "Tools..."
        Me.Group2.Name = "Group2"
        '
        'btnDeleteDemos
        '
        Me.btnDeleteDemos.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnDeleteDemos.Label = "ALLE Demo-Tabs löschen"
        Me.btnDeleteDemos.Name = "btnDeleteDemos"
        Me.btnDeleteDemos.OfficeImageId = "Delete"
        Me.btnDeleteDemos.ScreenTip = "Tool - Demos löschen"
        Me.btnDeleteDemos.ShowImage = True
        Me.btnDeleteDemos.SuperTip = "Alle erstellten Demo Tabellenblätter (und nur diese) werden entfernt"
        '
        'btnDeleteDemo1
        '
        Me.btnDeleteDemo1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnDeleteDemo1.Label = "EIN Demo-Tab löschen"
        Me.btnDeleteDemo1.Name = "btnDeleteDemo1"
        Me.btnDeleteDemo1.OfficeImageId = "Export"
        Me.btnDeleteDemo1.ScreenTip = "Tool - Demo löschen"
        Me.btnDeleteDemo1.ShowImage = True
        Me.btnDeleteDemo1.SuperTip = "Das aktuelle erstellte Demo Tabellenblatt (und nur dieses) wird entfernt"
        '
        'btnHidePane
        '
        Me.btnHidePane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnHidePane.Label = "Katalog schließen..."
        Me.btnHidePane.Name = "btnHidePane"
        Me.btnHidePane.OfficeImageId = "FileClose"
        Me.btnHidePane.ScreenTip = "Tool - Katalog (Taskpane) schließen"
        Me.btnHidePane.ShowImage = True
        Me.btnHidePane.SuperTip = "Der Katalog mit den Beispielen wird geschlossen"
        '
        'btnShowPane
        '
        Me.btnShowPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnShowPane.Label = "Katalog öffnen..."
        Me.btnShowPane.Name = "btnShowPane"
        Me.btnShowPane.OfficeImageId = "FileOpen"
        Me.btnShowPane.ScreenTip = "Tool - Katalog (Taskpane) öffnen"
        Me.btnShowPane.ShowImage = True
        Me.btnShowPane.SuperTip = "Der Katalog mit den Beispielen wird rechts als TaskPane geöffnet"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Demo)
        Me.Demo.ResumeLayout(False)
        Me.Demo.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Demo As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnPowerQuery As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRegex As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnVba As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFormulas As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnDeleteDemos As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnHidePane As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnShowPane As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDeleteDemo1 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
