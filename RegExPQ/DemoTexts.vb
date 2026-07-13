Public Module DemoTexts
    Public Function PowerQueryInfoText() As String
        Return TextBlock(
        <text>
            <![CDATA[
Um den M-Code zu testen:
Button "Demo erzeugen" klicken.
Der Button "Code kopieren" kopiert den M-Code in die Zwischenablage.
Dann im Ribbon Daten - "Daten abrufen und transformieren - Daten abrufen - Aus anderen Quellen - Leere Abfrage" anklicken.
Im PQ-Editor auf "Erweiterter Editor" klicken und dort den gesamten Text markieren und den Inhalt der Zwischenablage einfügen. Nun Fertig klicken.
Dann Schließen & laden - Schließen und laden in... anklicken.
Im folgenden Fenster entweder "Neues Tabellenblatt" lassen, oder auf "Bestehendes Arbeitsblatt:" klicken und Zielzelle auswählen/angeben.
OK klicken.
                ]]>
        </text>
        )
    End Function
End Module