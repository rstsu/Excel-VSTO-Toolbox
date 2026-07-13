'Imports Excel = Microsoft.Office.Interop.Excel
'Imports System.Windows.Forms

Partial Public Class DemoRunner

    Private Sub CreateFormulaDemo_1()
        Dim ws = CreateFreshSheet("Demo_Formel_1")

        ws.Range("A1").Value = "Werte"
        ws.Range("A2").Value = "Excel"
        ws.Range("A3").Value = "Power Query"
        ws.Range("A4").Value = "Excel"
        ws.Range("A5").Value = "VBA"
        ws.Range("A6").Value = "Formeln"
        ws.Range("A7").Value = "Power Query"

        ws.Range("C1").Value = "Eindeutig sortiert"
        ws.Range("C2").Formula2 = "=SORT(UNIQUE(A2:.A999))"
        ws.Range("D1").Formula2 = "=VSTACK(A1,SORT(UNIQUE(A2:.A999)))"
        FormatSheet(ws)
    End Sub
    Private Sub CreateFormulaDemo_2()
        Dim ws = CreateFreshSheet("Demo_Formel_2")

        Dim data(,) As Object = {
            {"Code", "Text"},
            {"3.40.20", "Text 1"},
            {"3.10.20", "Text 2"},
            {"2.11.73", "Text 3"},
            {"1.10.11", "Text 4"},
            {"3.50.20", "Text 5"},
            {"1.11.72", "Text 6"},
            {"1.11.74", "Text 7"},
            {"12.52.63", "Text 8"},
            {"2.11.74", "Text 4"},
            {"2.10.11", "Text 10"}
        }

        ws.Range("A1:B11").Value = data

        ws.Range("D1").Value = "Code"
        ws.Range("E1").Value = "Text"
        ws.Range("D2").Formula2 = "=SORTBY(A2:.B999,--TEXTBEFORE(A2:.A999,"".""),1,--MID(TEXTAFTER(A2:.A999,"".""),1,2),1)"
        ws.Range("G1").Formula2 = "=VSTACK(A1:B1,SORTBY(A2:.B999,--TEXTBEFORE(A2:.A999,"".""),1,--MID(TEXTAFTER(A2:.A999,"".""),1,2),1))"
        ws.Range("J1").Value = "Code"
        ws.Range("K1").Value = "Text"
        ws.Range("J2").Formula2 = "=LET(w,A2:.A999,x,A2:.B999,y,--TEXTBEFORE(w,"".""),z,--MID(TEXTAFTER(w,"".""),1,2),SORTBY(x,y,1,z,1))"
        ws.Range("M1").Formula2 = "=LET(w,A2:.A999,x,A2:.B999,y,--TEXTBEFORE(w,"".""),z,--MID(TEXTAFTER(w,"".""),1,2),VSTACK(A1:B1,SORTBY(x,y,1,z,1)))"
        FormatSheet(ws)
    End Sub
    Private Sub CreateFormulaDemo_3()
        Dim ws = CreateFreshSheet("Demo_Formel_3")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Jahr:"
            ws.Range("A2").Value = "Monat:"
            ws.Range("B1").Formula2 = "=YEAR(TODAY())"
            ws.Range("B2").Formula2 = "=MONTH(TODAY())"
            ws.Range("A2:B2").Font.Bold = True
            ws.Range("I1").Value = "Datum"
            ws.Range("J1").Value = "Feiertag"
            ws.Range("I2").Formula2 = "=LET(Jahr,$B$1,Ostersonntag,ROUND((DAY(MINUTE(Jahr/38)/2+55)&"".4.""&Jahr)/7,0)*7-6,Daten,VSTACK(DATE(Jahr,{1;5;10;12;12},{1;1;3;25;26}),Ostersonntag+{-2;1;39;50}),Namen,{""Neujahr"";""Tag der Arbeit"";""Tag der Deutschen Einheit"";""1. Weihnachtstag"";""2. Weihnachtstag"";""Karfreitag"";""Ostermontag"";""Christi Himmelfahrt"";""Pfingstmontag""},SORTBY(HSTACK(Daten,Namen),Daten))"
            ws.Range("A4").Formula2 = "=LET(Jahr,$B$1,Monat,$B$2,Erster,DATE(Jahr,Monat,1),Start,Erster-WEEKDAY(Erster,2)+1,SEQUENCE(6,7,Start,1))"
            ws.Range("A12").Formula2 = "=LET(Jahr,$B$1,Monat,$B$2,Erster,DATE(Jahr,Monat,1),Start,Erster-WEEKDAY(Erster,2)+1,VSTACK(TEXT(SEQUENCE(1,7,2,1),""TTT""),SEQUENCE(6,7,Start,1)))"
            ws.Range("A21").Formula2 = "=LET(Jahr,$B$1,Monat,$B$2,Erster,DATE(Jahr,Monat,1),Start,Erster-WEEKDAY(Erster,2)+1,VSTACK(HSTACK(TEXT(Erster,""MMMM JJJJ""),"""","""","""","""","""",""""),TEXT(SEQUENCE(1,7,2,1),""TTT""),SEQUENCE(6,7,Start,1)))"
            ws.Range("I11").Formula2 = "=Today()"
            ws.Range("I12").Formula2 = "=Today()+4"
            ws.Range("J11").Value = "Sehr wichtiger Tag"
            ws.Range("J12").Value = "Internationaler Schwarzbier-Tag"
            SetupYearMonthValidation(ws)
            FormatDemoSheet(ws)
            Globals.ThisAddIn.AttachDemoChangeHandler(ws)
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
    End Sub
End Class