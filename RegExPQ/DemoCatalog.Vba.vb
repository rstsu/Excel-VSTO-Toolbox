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
' Excel-VSTO-Toolbox
' VBA-Demo
' https://github.com/rstsu/Excel-VSTO-Toolbox
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
' Excel-VSTO-Toolbox
' VBA-Demo
' https://github.com/rstsu/Excel-VSTO-Toolbox
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
' Excel-VSTO-Toolbox
' VBA-Demo
' https://github.com/rstsu/Excel-VSTO-Toolbox
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
            }, New DemoDefinition With {
                .Id = "vba_006",
                .Category = DemoCategory.Vba,
                .Title = "Access Datenbank erstellen...",
                .Tags = {"vba", "access", "adox", "ace", "oledb"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Eine Access-Datenbank komplett per VBA erzeugen – inklusive Tabelle, Autowert-Primärschlüssel und Index. Die Datenbank wird am Ende geöffnet.
Die Datei wird im lokalenTemp-Ordner erstellt.

!!!!!!!!WICHTIG!!!!!!!!
Den Code über den Button "Code kopieren" in die Zwischenablage übernehmen und dann im VBA-Editor ein Modul einfügen und dort den Code hineinkopieren.
Erst dann funktioniert der Button im Tabellenblatt.
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
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
Private Const SW_MAXIMIZE = 3
Public Sub Main_Access_1()
    Dim strDescription As String
    Dim varProvider As Variant
    Dim strFileName As String
    Dim strProvider As String
    Dim catCatalog As Object
    Dim objTable As Object
    Dim objIndex As Object
    Dim objConn As Object
    Dim lngError As Long
    strFileName = Environ$("TEMP") & "\VBA_Demo_6.accdb"
    On Error GoTo Fin
    If Len(Dir$(strFileName)) > 0 Then Kill strFileName
    Set catCatalog = CreateObject("ADOX.Catalog")
    For Each varProvider In Array("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0")
        Err.Clear
        On Error Resume Next
        catCatalog.Create "Provider=" & CStr(varProvider) & ";Data Source=" & strFileName & ";"
        If Err.Number = 0 Then
            strProvider = CStr(varProvider)
            On Error GoTo Fin
            Exit For
        End If
        lngError = Err.Number
        strDescription = Err.Description
        On Error GoTo Fin
    Next varProvider
    If Len(strProvider) = 0 Then
        Err.Raise vbObjectError + 1000, _
            "CreateDataBase_1", _
            "Es wurde kein geeigneter ACE-OLEDB-Provider gefunden." & _
            vbCrLf & vbCrLf & _
            "Letzter Fehler: " & lngError & _
            vbCrLf & strDescription
    End If
    Set objConn = CreateObject("ADODB.Connection")
    With objConn
        .CursorLocation = 3       'adUseClient
        .Provider = strProvider
        .Properties("Data Source") = strFileName
        .Open
    End With
    Set objTable = CreateObject("ADOX.Table")
    With objTable
        .Name = "Lieferanten"
        .ParentCatalog = catCatalog
        .Columns.Append "Primaer", 3   'adInteger
        With .Columns("Primaer")
            .Properties("Description") = "Schluessel"
            .Properties("Autoincrement") = True
        End With
        .Columns.Append "Name", 202, 60   'adVarWChar
        With .Columns("Name")
            .Properties("Description") = "Nachname"
            .Properties("Jet OLEDB:Allow Zero Length") = True
            .Properties("Nullable") = True
        End With
    End With
    catCatalog.Tables.Append objTable
    Set objIndex = CreateObject("ADOX.Index")
    With objIndex
        .Name = "PrimaryKey"
        .Columns.Append "Primaer"
        .PrimaryKey = True
        .Unique = True
    End With
    objTable.Indexes.Append objIndex
    MsgBox "Datenbank wurde erstellt:" & vbCrLf & _
        strFileName & vbCrLf & vbCrLf & _
        "Provider: " & strProvider, vbInformation
    If objConn.State = 1 Then objConn.Close
    If ShellExecute(Application.hwnd, "Open", strFileName, vbNullString, vbNullString, SW_MAXIMIZE) <= 32 Then
        MsgBox "Die Datenbank wurde erstellt, konnte aber nicht geöffnet werden.", vbExclamation
    End If
Fin:
    If Err.Number <> 0 Then MsgBox "Fehler: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    If Not objConn Is Nothing Then
        If objConn.State = 1 Then objConn.Close
    End If
    Set objIndex = Nothing
    Set objTable = Nothing
    Set catCatalog = Nothing
    Set objConn = Nothing
End Sub
        ]]>
    </code>
        )
            }, New DemoDefinition With {
                .Id = "vba_007",
                .Category = DemoCategory.Vba,
                .Title = "Spalte nach Vorgabe aufteilen - mit UDF und Formeln",
                .Tags = {"vba", "udf", "aufteilen", "spalte", "formeln"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Spalte A (mit Überschrift) und Spalte C (ohne Überschrift) werden nach Vorgaben aufgeteilt.
Mit einer UDF (Tabellenblatt UDF) und mit Formeln (Tabellenblatt Formeln).
"Alte" UDF mit "neuen" Formeln ersetzt. ;-)

=fncSplit("A")          3er-Gruppen pro Zeile, keine Überschrift
=fncSplit("A";4)        4er-Gruppen pro Zeile, keine Überschrift
=fncSplit("A";;0)       3 Zeilen pro Spalte, Überschrift
=fncSplit("A"; 3; 1)    3er-Gruppen pro Zeile, mit Überschrift
=fncSplit("C";;0;0)     3 Zeilen pro Spalte, keine Überschrift
=fncSplit("C";4;0;0)    4 Zeilen pro Spalte, keine Überschrift

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv
in folgenden Ordner entpackt:

%TEMP%\Excel-VSTO-Toolbox\VBA_07

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
Option Explicit
'Excel -VSTO - Toolbox
'Power Query - Demo
'Ralf Stolzenburg (Case)
'https://github.com/rstsu/Excel-VSTO-Toolbox

'=fncSplit("A")         3er-Gruppen pro Zeile, keine Überschrift
'=fncSplit("A";4)       4er-Gruppen pro Zeile, keine Überschrift
'=fncSplit("A";;0)      3 Zeilen pro Spalte, Überschrift
'=fncSplit("A"; 3; 1)   3er-Gruppen pro Zeile, mit Überschrift (erste Zeile ignoriert)
'=fncSplit("C";;0;0)    3 Zeilen pro Spalte, keine Überschrift
'=fncSplit("C";4;0;0)   4 Zeilen pro Spalte, keine Überschrift
Public Function fncSplit(strColumn As String, Optional lngCount As Long = 3, Optional lngArt As Long = 1, Optional lngHead As Long = 1) As Variant
    Dim varResult() As Variant
    Dim lngColumn As Long
    Dim rngRange As Range
    Dim varArr As Variant
    Dim lngRCount As Long
    Dim lngCCount As Long
    Dim lngStart As Long
    Dim lngRow As Long
    Dim lngRC As Long
    Set rngRange = Range(strColumn & "1", Cells(Rows.Count, strColumn).End(xlUp))
    lngStart = 1 + lngHead
    varArr = rngRange.Resize(rngRange.Rows.Count - lngHead, 1).Offset(lngHead, 0).Value
    lngRC = UBound(varArr, 1)
    Do While lngRC Mod lngCount <> 0
        lngRC = lngRC + 1
    Loop
    If lngArt = 1 Then
        lngRCount = lngRC \ lngCount
        lngCCount = lngCount
        ReDim varResult(1 To lngRCount, 1 To lngCCount)
        For lngRow = 1 To lngRCount
            For lngColumn = 1 To lngCCount
                If (lngRow - 1) * lngCount + lngColumn <= UBound(varArr, 1) Then
                    varResult(lngRow, lngColumn) = varArr((lngRow - 1) * lngCount + lngColumn, 1)
                Else
                    varResult(lngRow, lngColumn) = vbNullString
                End If
            Next lngColumn
        Next lngRow
    Else
        lngRCount = lngCount
        lngCCount = lngRC \ lngCount
        ReDim varResult(1 To lngRCount, 1 To lngCCount)
        For lngColumn = 1 To lngCCount
            For lngRow = 1 To lngRCount
                If (lngColumn - 1) * lngCount + lngRow <= UBound(varArr, 1) Then
                    varResult(lngRow, lngColumn) = varArr((lngColumn - 1) * lngCount + lngRow, 1)
                Else
                    varResult(lngRow, lngColumn) = vbNullString
                End If
            Next lngRow
        Next lngColumn
    End If
    fncSplit = varResult
End Function
        ]]>
    </code>
        )
            }, New DemoDefinition With {
                .Id = "vba_008",
                .Category = DemoCategory.Vba,
                .Title = "Access Datenbank erstellen - die Zweite...",
                .Tags = {"vba", "access", "adox", "ace", "oledb"},
                .Description = TextBlock(
    <text>
        <![CDATA[
Es wird eine Accessdatenbank im lokalen Temp-Ordner mit Namen "VBA_Demo_8.accdb" aus Excel erstellt.
Access_Datenbank_erstellen_VBA_Demo_8_accdb.xlsb

Dann werden folgende Tabellen angelegt:
1. Kunden
2. Artikel
3. Bestellungen
4. Bestellpositionen

In den entsprechenden Tabellen werden Daten zu Kunden, Artikel, Bestellungen und  Bestellpositionen erstellt.
Nachdem noch Beziehungen eingerichtet sind, wird die Datei geöffnet.

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv
in folgenden Ordner entpackt:

%TEMP%\Excel-VSTO-Toolbox\VBA_08

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
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
' Ralf Stolzenburg (Case)
' https://github.com/rstsu/Excel-VSTO-Toolbox
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
Private Const SW_MAXIMIZE = 3
Public Sub Main_Access_2()
    Dim strDescription As String
    Dim varProvider As Variant
    Dim strFileName As String
    Dim strProvider As String
    Dim catCatalog As Object
    Dim objTable As Object
    Dim objIndex As Object
    Dim objConn As Object
    Dim lngError As Long
    strFileName = Environ$("TEMP") & "\VBA_Demo_8.accdb"
    On Error GoTo Fin
    If Len(Dir$(strFileName)) > 0 Then Kill strFileName
    Set catCatalog = CreateObject("ADOX.Catalog")
    For Each varProvider In Array("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0")
        Err.Clear
        On Error Resume Next
        catCatalog.Create "Provider=" & CStr(varProvider) & ";Data Source=" & strFileName & ";"
        If Err.Number = 0 Then
            strProvider = CStr(varProvider)
            On Error GoTo Fin
            Exit For
        End If
        lngError = Err.Number
        strDescription = Err.Description
        On Error GoTo Fin
    Next varProvider
    If Len(strProvider) = 0 Then
        Err.Raise vbObjectError + 1000, _
            "CreateDataBase_1", _
            "Es wurde kein geeigneter ACE-OLEDB-Provider gefunden." & _
            vbCrLf & vbCrLf & _
            "Letzter Fehler: " & lngError & _
            vbCrLf & strDescription
    End If
    Set objConn = CreateObject("ADODB.Connection")
    With objConn
        .CursorLocation = 3
        .Open "Provider=" & strProvider & ";Data Source=" & strFileName & ";"
        .Execute _
            "CREATE TABLE Kunden (" & _
            "KundenID AUTOINCREMENT CONSTRAINT PK_Kunden PRIMARY KEY, " & _
            "Firma VARCHAR(100), " & _
            "Ansprechpartner VARCHAR(100), " & _
            "Ort VARCHAR(50), " & _
            "Land VARCHAR(50))"
        .Execute _
            "CREATE TABLE Artikel (" & _
            "ArtikelID AUTOINCREMENT CONSTRAINT PK_Artikel PRIMARY KEY, " & _
            "Bezeichnung VARCHAR(100), " & _
            "Preis CURRENCY, " & _
            "Bestand INTEGER)"
        .Execute _
            "CREATE TABLE Bestellungen (" & _
            "BestellID AUTOINCREMENT CONSTRAINT PK_Bestellungen PRIMARY KEY, " & _
            "KundenID LONG, " & _
            "Bestelldatum DATETIME)"
        .Execute _
            "CREATE TABLE Bestellpositionen (" & _
            "PositionsID AUTOINCREMENT CONSTRAINT PK_Bestellpositionen PRIMARY KEY, " & _
            "BestellID LONG, " & _
            "ArtikelID LONG, " & _
            "Menge INTEGER, " & _
            "Einzelpreis CURRENCY)"
        .Execute _
            "ALTER TABLE Bestellungen " & _
            "ADD CONSTRAINT FK_Bestellungen_Kunden " & _
            "FOREIGN KEY (KundenID) REFERENCES Kunden (KundenID)"
        .Execute _
            "ALTER TABLE Bestellpositionen " & _
            "ADD CONSTRAINT FK_Positionen_Bestellungen " & _
            "FOREIGN KEY (BestellID) REFERENCES Bestellungen (BestellID)"
        .Execute _
            "ALTER TABLE Bestellpositionen " & _
            "ADD CONSTRAINT FK_Positionen_Artikel " & _
            "FOREIGN KEY (ArtikelID) REFERENCES Artikel (ArtikelID)"
        .Execute "INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Nordstern GmbH', 'Anna Berger', 'Hamburg', 'Deutschland')"
        .Execute "INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Rheinblick AG', 'Michael Weber', 'Köln', 'Deutschland')"
        .Execute "INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Alpenhandel KG', 'Sabine Keller', 'München', 'Deutschland')"
        .Execute "INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Spree Büroservice', 'Thomas Richter', 'Berlin', 'Deutschland')"
        .Execute "INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('MainTech GmbH', 'Julia Hartmann', 'Frankfurt', 'Deutschland')"
        .Execute "INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Elbe Logistik', 'Martin Schulze', 'Dresden', 'Deutschland')"
        .Execute "INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Westfalen Bürobedarf', 'Petra König', 'Dortmund', 'Deutschland')"
        .Execute "INSERT INTO Kunden (Firma, Ansprechpartner, Ort, Land) VALUES ('Donau Consulting', 'Daniel Fischer', 'Ulm', 'Deutschland')"
        .Execute "INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Notebookständer', 29.9, 35)"
        .Execute "INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('USB-C Hub', 39.5, 28)"
        .Execute "INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Funkmaus', 24.9, 42)"
        .Execute "INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Tastatur', 49.9, 21)"
        .Execute "INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Webcam', 59.0, 17)"
        .Execute "INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Headset', 69.9, 14)"
        .Execute "INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Monitorarm', 84.5, 9)"
        .Execute "INSERT INTO Artikel (Bezeichnung, Preis, Bestand) VALUES ('Laptop-Tasche', 44.0, 31)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (1, #01/12/2026#)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (2, #01/18/2026#)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (1, #02/03/2026#)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (4, #02/15/2026#)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (5, #03/02/2026#)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (3, #03/21/2026#)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (6, #04/09/2026#)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (2, #04/28/2026#)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (7, #05/10/2026#)"
        .Execute "INSERT INTO Bestellungen (KundenID, Bestelldatum) VALUES (8, #05/23/2026#)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (1, 1, 2, 29.9)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (1, 3, 2, 24.9)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (2, 2, 3, 39.5)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (2, 4, 1, 49.9)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (3, 5, 2, 59)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (3, 6, 1, 69.9)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (4, 3, 4, 24.9)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (4, 8, 2, 44)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (5, 7, 2, 84.5)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (5, 2, 2, 39.5)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (6, 4, 3, 49.9)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (7, 6, 2, 69.9)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (8, 1, 5, 29.9)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (8, 5, 2, 59)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (9, 8, 3, 44)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (10, 2, 1, 39.5)"
        .Execute "INSERT INTO Bestellpositionen (BestellID, ArtikelID, Menge, Einzelpreis) VALUES (10, 3, 2, 24.9)"
        MsgBox "Datenbank wurde erstellt:" & vbCrLf & _
            strFileName & vbCrLf & vbCrLf & _
            "Provider: " & strProvider, vbInformation
        If .State = 1 Then .Close
    End With
    If ShellExecute(Application.hwnd, "Open", strFileName, vbNullString, vbNullString, SW_MAXIMIZE) <= 32 Then
        MsgBox "Die Datenbank wurde erstellt, konnte aber nicht geöffnet werden.", vbExclamation
    End If
Fin:
    If Err.Number <> 0 Then MsgBox "Fehler: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    If Not objConn Is Nothing Then
        If objConn.State = 1 Then objConn.Close
    End If
    Set objIndex = Nothing
    Set objTable = Nothing
    Set catCatalog = Nothing
    Set objConn = Nothing
End Sub
        ]]>
    </code>
        )
            }, New DemoDefinition With {
                .Id = "vba_009",
                .Category = DemoCategory.Vba,
                .Title = "Klassenprogrammierung - UserForm - TextBox - Text markieren",
                .Tags = {"vba", "userform", "klassenprogrammierung", "textbox", "markieren", "farbe", "doppelklick"},
                .Description = TextBlock(
    <text>
        <![CDATA[
UserForm - TextBox - Klassenprogrammierung.
Klassenprogrammierung_UserForm_TextBox_Text_markieren_Farbe_wechseln_und_zurueck.xlsb

In einer UserForm werden alle TextBoxen in einem Array gesammelt und der Klasse zugewiesen.
Damit wird bei jeder TextBox beim aktivieren der gesamte Text markiert.
Bei einem Doppleklick in die TextBox wird Schrift- und Hintergrundfarbe geändert. Nächster Doppelklick alles wieder zurück.

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv
in folgenden Ordner entpackt:

%TEMP%\Excel-VSTO-Toolbox\VBA_09

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
DieseArbeitsmappe:
Option Explicit
Private Sub Workbook_Open()
    UserForm1.Show
End Sub

UserForm1:
Option Explicit
' Excel-VSTO-Toolbox
' VBA-Demo
' Ralf Stolzenburg (Case)
' https://github.com/rstsu/Excel-VSTO-Toolbox
' Klasse initialisieren. Nur TextBoxen über Array sammeln.
' Text in TextBox1 markieren. Bei Doppelklick Farbe wechseln und zurück.
Private mobjTextBoxClass() As clsTextBox
Private Sub UserForm_Initialize()
    Dim objControl As Control
    Dim lngIndex As Long
    For Each objControl In Controls
        If TypeOf objControl Is MSForms.TextBox Then
            ReDim Preserve mobjTextBoxClass(lngIndex)
            Set mobjTextBoxClass(lngIndex) = New clsTextBox
            Set mobjTextBoxClass(lngIndex).prpTextBox = objControl
            lngIndex = lngIndex + 1
        End If
    Next objControl
    With TextBox1
        .SelStart = 0
        .SelLength = .TextLength
    End With
End Sub
Private Sub UserForm_Terminate()
    Dim lngIndex As Long
    If CBool(Not Not mobjTextBoxClass) Then
        For lngIndex = LBound(mobjTextBoxClass) To UBound(mobjTextBoxClass)
            Set mobjTextBoxClass(lngIndex) = Nothing
        Next lngIndex
    End If
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub

Modul1:
Option Explicit
Public Sub Main()
    UserForm1.Show
End Sub

Klassenmodul (Name = clsTextBox):
Option Explicit
' Excel-VSTO-Toolbox
' VBA-Demo
' Ralf Stolzenburg (Case)
' https://github.com/rstsu/Excel-VSTO-Toolbox
' Klasse initialisieren. Nur TextBoxen über Array sammeln.
' Text in TextBox1 markieren. Bei Doppelklick Farbe wechseln und zurück.
' https://learn.microsoft.com/de-de/dotnet/visual-basic/language-reference/modifiers/withevents
Private WithEvents mobjTextBox As MSForms.TextBox
' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/property-get-statement
Private Property Get TextBox() As MSForms.TextBox
    Set TextBox = mobjTextBox
End Property
' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/property-set-statement
Friend Property Set prpTextBox(objTextBox As MSForms.TextBox)
    Set mobjTextBox = objTextBox
End Property
' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/terminate-event-visual-basic-for-applications
Private Sub Class_Terminate()
    Set mobjTextBox = Nothing
End Sub
' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/mousedown-mouseup-events
Private Sub mobjTextBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With mobjTextBox
        .SelStart = 0
        .SelLength = .TextLength
    End With
End Sub
' https://learn.microsoft.com/de-de/office/vba/language/reference/user-interface-help/dblclick-event
Private Sub mobjTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With TextBox
        .BackColor = IIf(.BackColor = &H80000005, &HC0FFC0, &H80000005)
        .ForeColor = IIf(.ForeColor = -2147483640, &HFF&, -2147483640)
    End With
End Sub
        ]]>
    </code>
        )
            }, New DemoDefinition With {
                .Id = "vba_0010",
                .Category = DemoCategory.Vba,
                .Title = "Datum umwandeln",
                .Tags = {"formel", "datum", "utc", "mez", "mesz", "vba", "power query"},
                .Description = TextBlock(
    <text>
        <![CDATA[
March 6 2020 8:00:12 AM
February 28 2020 11:17:26 AM
January 24 2020 3:22:44 PM
December 30 2023 1:38 PM
November 22 2024 7:29 AM
8/31/2025 8:15:57 AM
4/24/2022 9:31:16 AM
9/24/2026 9:31:16 PM
2026-11-22T14:30:00Z
2026-08-12T14:30:00Z

Diese Daten werden umgewandelt. Auch UTC nach MEZ/MESZ.
VBA_Power_Query_Formel_Datum_umwandeln.xlsb

Beim Klick auf "Demo erzeugen" wird das mitgelieferte ZIP-Archiv in folgenden Ordner entpackt:
%TEMP%\Excel-VSTO-Toolbox\Demo_VBA_PQ_Formel

Ein bereits vorhandener Demo-Ordner wird vorher gelöscht.
Anschließend wird die enthaltene Excel-Arbeitsmappe geöffnet.

!!!!!!!!WICHTIG!!!!!!!!
Falls eine Datei aus dem Demo-Ordner noch geöffnet ist, kann der
vorhandene Ordner nicht gelöscht und das Beispiel nicht erneut
bereitgestellt werden.
!!!!!!!!WICHTIG!!!!!!!!
        ]]>
    </text>
        ),
.CodeText = TextBlock(
    <code>
        <![CDATA[
Option Explicit
' Excel -VSTO - Toolbox
' Power Query - Demo
' Ralf Stolzenburg (Case)
' https://github.com/rstsu/Excel-VSTO-Toolbox
Private Declare PtrSafe Function VarDateFromStr Lib "oleaut32.dll" _
    (ByVal strIn As LongPtr, _
    ByVal lcid As Long, _
    ByVal dwFlags As Long, _
    ByRef pDateOut As Date) As Long
' https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-vardatefromstr
Function fncDate(ByVal strTMP As String, Optional ByVal lcid As Long = 1033) As Date
    Call VarDateFromStr(StrPtr(strTMP), lcid, &H80000000, fncDate)
End Function
Public Sub Main()
    Dim lngLastRow As Long
    Dim dateTMP As Date
    On Error GoTo Fin
    With Tabelle1
        For lngLastRow = 2 To IIf(Len(.Cells(.Rows.Count, 1)), .Rows.Count, .Cells(.Rows.Count, 1).End(xlUp).Row)
            If .Cells(lngLastRow, 1).Value <> "" Then
                dateTMP = fncDate(.Cells(lngLastRow, 1).Text)
                If dateTMP = "00:00:00" Then dateTMP = fncUTCtoDE(.Cells(lngLastRow, 1).Text)
                .Cells(lngLastRow, 4).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                .Cells(lngLastRow, 4).Value = dateTMP
                .Cells(lngLastRow, 6).NumberFormat = "dd/mm/yyyy"
                .Cells(lngLastRow, 6).Value = DateValue(dateTMP)
                .Cells(lngLastRow, 7).NumberFormat = "hh:mm:ss"
                .Cells(lngLastRow, 7).Value = TimeValue(dateTMP)
            End If
        Next lngLastRow
    End With
Fin:
    If Err.Number <> 0 Then MsgBox "Fehler: " & Err.Number & " " & Err.Description
End Sub
Public Function fncUTCtoDE(utcString As String) As Date
    Dim lngOffset As Long
    Dim dateStart As Date
    Dim dateEnd As Date
    Dim dateUTC As Date
    dateUTC = DateSerial(CInt(Mid$(utcString, 1, 4)), CInt(Mid$(utcString, 6, 2)), CInt(Mid$(utcString, 9, 2))) + _
        TimeSerial(CInt(Mid$(utcString, 12, 2)), CInt(Mid$(utcString, 15, 2)), CInt(Mid$(utcString, 18, 2)))
    dateStart = DateSerial(CInt(Mid$(utcString, 1, 4)), 3, 31 - Weekday(DateSerial(CInt(Mid$(utcString, 1, 4)), 3, 31), vbMonday) + 1)
    dateEnd = DateSerial(CInt(Mid$(utcString, 1, 4)), 10, 31 - Weekday(DateSerial(CInt(Mid$(utcString, 1, 4)), 10, 31), vbMonday) + 1)
    lngOffset = 1
    If dateUTC >= dateStart + TimeSerial(1, 0, 0) And dateUTC < dateEnd + TimeSerial(1, 0, 0) Then
        lngOffset = lngOffset + 1
    End If
    fncUTCtoDE = dateUTC + lngOffset / 24
End Function
        ]]>
    </code>
        )
            }
        }
    End Function
End Module

