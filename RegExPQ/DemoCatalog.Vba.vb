Imports System.Xml.Linq

Partial Public Module DemoCatalog
    Private Function GetVbaExamples() As List(Of DemoDefinition)
        Return New List(Of DemoDefinition) From {
            New DemoDefinition With {
                .Id = "vba_001",
                .Category = DemoCategory.Vba,
                .Title = "Alle ENVIRON Variablen ausgeben",
                .Tags = {"vba", "environ", "variablen", "ausgeben", "text"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Alle ENVIRON Variablen mit Inhalt werden in einer neuen Datei aufgelistet.
Wie z. B. USERNAME, USERPROFILE, TEMP...
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
Option Explicit
Public Sub Main()
    Dim wkbBook As Workbook
    Dim strTMP() As String
    Dim lngTMP As Long
    On Error GoTo Fin
    lngTMP = 1
    Set wkbBook = Workbooks.Add(1)
    With wkbBook.Worksheets(lngTMP)
        Do
            strTMP = Split(Environ(lngTMP), "=")
            If Join(strTMP) <> "" Then
                .Cells(lngTMP, 1).Value = strTMP(0)
                .Cells(lngTMP, 2).Value = strTMP(1)
                lngTMP = lngTMP + 1
            End If
        Loop Until Join(strTMP) = ""
        .Columns("A:B").AutoFit
    End With
Fin:
    Set wkbBook = Nothing
    If Err.Number <> 0 Then MsgBox "Fehler: " & Err.Number & " " & Err.Description
End Sub
        ]]>
    </code>
        )
            }, New DemoDefinition With {
                .Id = "vba_002",
                .Category = DemoCategory.Vba,
                .Title = "Alle Filter aller Tabellenblätter zurücksetzen",
                .Tags = {"vba", "filter", "tabellenblätter", "zurücksetzen", "tabellenblatt"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Alle gesetzten Filter werden beim beenden der Datei zurückgesetzt.
Das passiert auf allen Tabellenblättern.

!!!!!!!!WICHTIG!!!!!!!!
Der Code gehört unter "DieseArbeitsmappe" - NICHT in ein Modul.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
Option Explicit
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim wksSheet As Worksheet
    For Each wksSheet In ThisWorkbook.Worksheets
        With wksSheet
            ' Falls die Tabellenblätter mit einem Paaswort geschützt sind
            '.Unprotect Password:="DEINPASSWORT"
            If .AutoFilterMode Then
                If .FilterMode Then
                    .ShowAllData
                End If
            End If
            ' Passwortschutz des Tabellenblattes wieder setzen
            ' UserInterfaceOnly auf True bedeutet - VBA kann Änderungen vornehmen, ohne dass der Blattschutz entfernt werden muss
            '.Protect Password:="DEINPASSWORT", UserInterfaceOnly:=True
        End With
    Next wksSheet
    ThisWorkbook.Save
End Sub
        ]]>
    </code>
        )
            }, New DemoDefinition With {
                .Id = "vba_003",
                .Category = DemoCategory.Vba,
                .Title = "UserForm zur Lauzeit erstellen",
                .Tags = {"vba", "userform", "laufzeit", "erstellen", "ausführen"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Es wird zur Laufzeit eine UserForm mit "OptionButtons" und "CommansButton" erstellt und gleich angezeigt.
Dort kann eine Auswahl getroffen werden.
Dann schließt sich die UserForm (Klick auf CommandButton "OK")und wird auch wieder entfernt, also keine UserForm vorhanden.

!!!!!!!!WICHTIG!!!!!!!!
Datei - Optionen - Trust Center - Einstellungen für das Trust Center... - Makroeinstellungen - "Zugriff auf das VBA-Projektmodell vertrauen" - dieser Haken MUSS gesetzt sein, damit der Code funktioniert.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
Option Explicit
Public Sub Main()
  Dim strC As String
  strC = strC & vbLf & "Private WithEvents Opt1 as MSForms.OptionButton"
  strC = strC & vbLf & "Private WithEvents Opt2 as MSForms.OptionButton"
  strC = strC & vbLf & "Private WithEvents Opt3 as MSForms.OptionButton"
  strC = strC & vbLf & "Private WithEvents Cmd1 as MSForms.CommandButton"
  strC = strC & vbLf
  strC = strC & vbLf & "Private Sub UserForm_Initialize"
  strC = strC & vbLf & "  With Me"
  strC = strC & vbLf & "    .Caption = ""Abfrage"""
  strC = strC & vbLf & "    .Width = 190"
  strC = strC & vbLf & "    With .Controls.Add(""Forms.Label.1"")"
  strC = strC & vbLf & "      .Top = 10"
  strC = strC & vbLf & "      .Left = 20"
  strC = strC & vbLf & "      .Width = 200"
  strC = strC & vbLf & "      .Font.Size = 10"
  strC = strC & vbLf & "      .Caption = ""Bitte wählen Sie aus den Optionen"""
  strC = strC & vbLf & "    End With"
  strC = strC & vbLf & "    Set Opt1 = .Controls.Add(""Forms.OptionButton.1"")"
  strC = strC & vbLf & "    With  Opt1"
  strC = strC & vbLf & "     .Top = 40"
  strC = strC & vbLf & "     .Left = 40"
  strC = strC & vbLf & "     .Caption = ""Drucken"""
  strC = strC & vbLf & "    End With"
  strC = strC & vbLf & "    Set Opt2 = .Controls.Add(""Forms.OptionButton.1"")"
  strC = strC & vbLf & "    With Opt2"
  strC = strC & vbLf & "      .Top = 60"
  strC = strC & vbLf & "      .Left = 40"
  strC = strC & vbLf & "      .Caption = ""Speichern"""
  strC = strC & vbLf & "    End With"
  strC = strC & vbLf & "    Set Opt3 = .Controls.Add(""Forms.OptionButton.1"")"
  strC = strC & vbLf & "    With Opt3"
  strC = strC & vbLf & "      .Top = 80"
  strC = strC & vbLf & "      .Left = 40"
  strC = strC & vbLf & "      .Caption = ""Drucken und Speichern"""
  strC = strC & vbLf & "    End With"
  strC = strC & vbLf & "    Set Cmd1 = .Controls.Add(""Forms.CommandButton.1"")"
  strC = strC & vbLf & "    With Cmd1"
  strC = strC & vbLf & "      .Top = 120"
  strC = strC & vbLf & "      .Left = 35"
  strC = strC & vbLf & "      .Width = 115"
  strC = strC & vbLf & "      .Caption = ""OK"""
  strC = strC & vbLf & "    End With"
  strC = strC & vbLf & "  End With"
  strC = strC & vbLf & "End Sub"
  strC = strC & vbLf
  strC = strC & vbLf & "Private Sub Cmd1_Click"
  strC = strC & vbLf & "  If Opt1 Then"
  strC = strC & vbLf & "    Call Drucken"
  strC = strC & vbLf & "    Unload Me"
  strC = strC & vbLf & "  ElseIf Opt2 Then"
  strC = strC & vbLf & "    Call Speichern"
  strC = strC & vbLf & "    Unload Me"
  strC = strC & vbLf & "  ElseIf Opt3 Then"
  strC = strC & vbLf & "    Call Drucken"
  strC = strC & vbLf & "    Call Speichern"
  strC = strC & vbLf & "    Unload Me"
  strC = strC & vbLf & "  Else"
  strC = strC & vbLf & "    MsgBox ""Bitte wählen Sie eine Option"", vbExclamation, ""Nochmal"""
  strC = strC & vbLf & "  End If"
  strC = strC & vbLf & "End Sub"
  With Application.VBE.ActiveVBProject
    With .VBComponents.Add(3)
      .CodeModule.InsertLines .CodeModule.CountOfLines + 1, strC
      VBA.UserForms.Add(.Name).Show
      .Collection.Remove .CodeModule.Parent
    End With
  End With
End Sub
Public Sub Drucken()
'Code zum drucken
  MsgBox "Drucken"
End Sub
Public Sub Speichern()
'Code zum speichern
  MsgBox "Speichern"
End Sub
        ]]>
    </code>
        )
            }
        }
    End Function

End Module

