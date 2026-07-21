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

!!!!!!!!WICHTIG!!!!!!!!
Den Code über den Button "Code kopieren" in die Zwischenablage übernehmen und dann im VBA-Editor ein Modul einfügen und dort den Code hineinkopieren.
!!!!!!!!WICHTIG!!!!!!!!
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
                .Title = "UserForm zur Laufzeit erstellen",
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
            }, New DemoDefinition With {
                .Id = "vba_004",
                .Category = DemoCategory.Vba,
                .Title = "Kreuztabelle aus Liste erste 3 Buchstaben",
                .Tags = {"vba", "text", "kreuztabelle", "intelligente tabelle", "liste"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Aus einer Liste (A2:A15) wird eine Kreuztabelle erstellt.
Grundlage sind die ersten 3 gleichen Buchstaben.
Es ist auch in Formeln und Power Query gelöst. Mit der gleichen Bezeichnung.
Die Aufgabe wurde mit einfachen Werten und einer "intelligenten Tabelle" gelöst.

!!!!!!!!WICHTIG!!!!!!!!
Den Code über den Button "Code kopieren" in die Zwischenablage übernehmen und dann im VBA-Editor ein Modul einfügen und dort den Code hineinkopieren.
Erst dann funktionieren die beiden Buttons im Tabellenblatt.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
Option Explicit
' Excel-VSTO-Toolbox
' VBA-Demo
' https://github.com/rstsu/Excel-VSTO-Toolbox
Public Sub Main_1()
    Dim strPrefix As String
    Dim vaArrZ() As Variant
    Dim varArrQ As Variant
    Dim lngMaxCol As Long
    Dim lngCount As Long
    Dim lngGroup As Long
    Dim strTMP As String
    Dim lngRow As Long
    Dim lngCol As Long
    With ThisWorkbook.Worksheets("Demo_VBA_4")
        .Range("G1:M" & .Rows.Count).Clear
        varArrQ = Range("C2", Cells(Rows.Count, "C").End(xlUp)).Value2
        lngGroup = 1
        lngCol = 1
        strTMP = Left$(varArrQ(1, 1), 3)
        For lngCount = 2 To UBound(varArrQ, 1)
            strPrefix = Left$(varArrQ(lngCount, 1), 3)
            If strPrefix = strTMP Then
                lngCol = lngCol + 1
            Else
                If lngCol > lngMaxCol Then lngMaxCol = lngCol
                lngGroup = lngGroup + 1
                lngCol = 1
                strTMP = strPrefix
            End If
        Next lngCount
        If lngCol > lngMaxCol Then lngMaxCol = lngCol
        ReDim vaArrZ(1 To lngGroup + 1, 1 To lngMaxCol)
        For lngCount = 1 To lngMaxCol
            vaArrZ(1, lngCount) = "Pos" & lngCount
        Next
        lngRow = 2
        lngCol = 1
        strTMP = Left$(varArrQ(1, 1), 3)
        vaArrZ(lngRow, lngCol) = varArrQ(1, 1)
        For lngCount = 2 To UBound(varArrQ, 1)
            strPrefix = Left$(varArrQ(lngCount, 1), 3)
            If strPrefix = strTMP Then
                lngCol = lngCol + 1
            Else
                lngRow = lngRow + 1
                lngCol = 1
                strTMP = strPrefix
            End If
            vaArrZ(lngRow, lngCol) = varArrQ(lngCount, 1)
        Next lngCount
        .Range("G1").Resize(UBound(vaArrZ, 1), UBound(vaArrZ, 2)).Value = vaArrZ
    End With
End Sub
' Excel-VSTO-Toolbox
' VBA-Demo
' https://github.com/rstsu/Excel-VSTO-Toolbox
Public Sub Main_2()
    Dim objList As ListObject
    Dim strPrefix As String
    Dim vaArrZ() As Variant
    Dim varArrQ As Variant
    Dim lngMaxCol As Long
    Dim lngCount As Long
    Dim lngGroup As Long
    Dim strTMP As String
    Dim lngRow As Long
    Dim lngCol As Long
    With ThisWorkbook.Worksheets("Demo_VBA_4")
        Set objList = .ListObjects("Demo_VBA_4")
        .Range("G1:M" & .Rows.Count).Clear
        varArrQ = objList.ListColumns(1).DataBodyRange.Value2
        lngGroup = 1
        lngCol = 1
        strTMP = Left$(varArrQ(1, 1), 3)
        For lngCount = 2 To UBound(varArrQ, 1)
            strPrefix = Left$(varArrQ(lngCount, 1), 3)
            If strPrefix = strTMP Then
                lngCol = lngCol + 1
            Else
                If lngCol > lngMaxCol Then lngMaxCol = lngCol
                lngGroup = lngGroup + 1
                lngCol = 1
                strTMP = strPrefix
            End If
        Next lngCount
        If lngCol > lngMaxCol Then lngMaxCol = lngCol
        ReDim vaArrZ(1 To lngGroup + 1, 1 To lngMaxCol)
        For lngCount = 1 To lngMaxCol
            vaArrZ(1, lngCount) = "Pos" & lngCount
        Next
        lngRow = 2
        lngCol = 1
        strTMP = Left$(varArrQ(1, 1), 3)
        vaArrZ(lngRow, lngCol) = varArrQ(1, 1)
        For lngCount = 2 To UBound(varArrQ, 1)
            strPrefix = Left$(varArrQ(lngCount, 1), 3)
            If strPrefix = strTMP Then
                lngCol = lngCol + 1
            Else
                lngRow = lngRow + 1
                lngCol = 1
                strTMP = strPrefix
            End If
            vaArrZ(lngRow, lngCol) = varArrQ(lngCount, 1)
        Next
        .Range("G1").Resize(UBound(vaArrZ, 1), UBound(vaArrZ, 2)).Value = vaArrZ
    End With
End Sub
        ]]>
    </code>
        )
            }, New DemoDefinition With {
                .Id = "vba_005",
                .Category = DemoCategory.Vba,
                .Title = "Links in klickbare Hyperlinks umwandeln",
                .Tags = {"vba", "link", "hyperlink", "umwandeln", "liste"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Eine Liste (A2:A7) mit Links wird in klickbare Hyperlinks umgewandelt (Main_1).
Die Texte werden in Spalte C als klickbare Hyperlinks ausgegeben (Main_2).
Die Hyperlinks (nicht der Text) in Spalte A werden entfernt. Und die Links in Spalte C komplett (Main_3).

!!!!!!!!WICHTIG!!!!!!!!
Den Code über den Button "Code kopieren" in die Zwischenablage übernehmen und dann im VBA-Editor ein Modul einfügen und dort den Code hineinkopieren.
Erst dann funktionieren die drei Buttons im Tabellenblatt.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
Option Explicit
' Excel-VSTO-Toolbox
' VBA-Demo
' https://github.com/rstsu/Excel-VSTO-Toolbox

' Die Links in Spalte A werden in klickbare Hyperlinks umgewandelt
' Es wird das aktive Tabellenblatt genommen
Public Sub Main_1_1()
    Dim rngCell As Range
    For Each rngCell In Range(Range("A2"), Cells(Rows.Count, 1).End(xlUp))
        rngCell.Parent.Hyperlinks.Add Anchor:=rngCell, Address:=rngCell.Value, TextToDisplay:=rngCell.Value
    Next rngCell
End Sub
' Oder trage die Formmel per VBA ein - hier in Spalte C, Links stehen in Spalte A
' Es wird das aktive Tabellenblatt genutzt
Public Sub Main_2_1()
    Range("C2:C" & Cells(Rows.Count, 1).End(xlUp).Row).Formula2 = "=HYPERLINK(A2,A2)"
    Columns("C").Autofit
End Sub
' Hyperlinks in Spalte A werden entfernt
' Auch hier das aktive Tabellenblatt
Public Sub Main_3_1()
    Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row).Hyperlinks.Delete
    Range("C2:C" & Cells(Rows.Count, 1).End(xlUp).Row).ClearContents
End Sub
        ]]>
    </code>
        )
            }
        }
    End Function

End Module

