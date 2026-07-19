Partial Public Class DemoRunner
    Private Sub CreateFormulaDemo_1()
        Dim ws = CreateFreshSheet("Demo_Formel_1")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
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
            ws.Range("F1").Value = "Excel-VSTO-Toolbox"
            ws.Range("F2").Value = "Formel-Demo"
            ws.Range("F3").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("F1:F3").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateFormulaDemo_2()
        Dim ws = CreateFreshSheet("Demo_Formel_2")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
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
            ws.Range("P1").Value = "Excel-VSTO-Toolbox"
            ws.Range("P2").Value = "Formel-Demo"
            ws.Range("P3").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("P1:P3").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
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
            ws.Range("L1").Value = "Excel-VSTO-Toolbox"
            ws.Range("L2").Value = "Formel-Demo"
            ws.Range("L3").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("L1:L3").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
            SetupYearMonthValidation(ws)
            FormatDemoSheet(ws)
            Globals.ThisAddIn.AttachDemoChangeHandler(ws)
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
    End Sub
    Private Sub CreateFormulaDemo_4()
        Dim ws = CreateFreshSheet("Demo_Formel_4")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Wert"
            ws.Range("A2").Value = "FRD001"
            ws.Range("A3").Value = "FRD003"
            ws.Range("A4").Value = "FRD005"
            ws.Range("A5").Value = "FRD007"
            ws.Range("A6").Value = "FRD008"
            ws.Range("A7").Value = "FRD009"
            ws.Range("A8").Value = "FRD010"
            ws.Range("A9").Value = "JUG003"
            ws.Range("A10").Value = "JUG005"
            ws.Range("A11").Value = "JUG006"
            ws.Range("A12").Value = "KIM001"
            ws.Range("A13").Value = "STD005"
            ws.Range("A14").Value = "STD006"
            ws.Range("A15").Value = "STD007"
            Dim lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, ws.Range("A1:A15"),, Excel.XlYesNoGuess.xlYes)
            lo.Name = "Demo_Formel_4"
            lo.TableStyle = "TableStyleMedium2"
            ws.Range("C1").Value = "Wert"
            ws.Range("C2").Value = "FRD001"
            ws.Range("C3").Value = "FRD003"
            ws.Range("C4").Value = "FRD005"
            ws.Range("C5").Value = "FRD007"
            ws.Range("C6").Value = "FRD008"
            ws.Range("C7").Value = "FRD009"
            ws.Range("C8").Value = "FRD010"
            ws.Range("C9").Value = "JUG003"
            ws.Range("C10").Value = "JUG005"
            ws.Range("C11").Value = "JUG006"
            ws.Range("C12").Value = "KIM001"
            ws.Range("C13").Value = "STD005"
            ws.Range("C14").Value = "STD006"
            ws.Range("C15").Value = "STD007"
            ws.Range("E1").Formula2 = "=LET(t,Demo_Formel_4[Wert],u,LEFT(t,3),v,UNIQUE(u),w,MAX(MAP(v,LAMBDA(x,SUM(--(u=x))))),VSTACK(""Pos""&SEQUENCE(,w),MAKEARRAY(ROWS(v),w,LAMBDA(y,z,IFERROR(INDEX(FILTER(t,u=INDEX(v,y)),z),"""")))))"
            ws.Range("M1").Formula2 = "=LET(m,Demo_Formel_4[Wert],k,UNIQUE(LEFT(m,3)),n,MAX(MAP(k,LAMBDA(x,SUM(--(LEFT(m,3)=x))))),VSTACK(""Pos""&SEQUENCE(,n),MAKEARRAY(ROWS(k),n,LAMBDA(z,p,IFERROR(INDEX(FILTER(m,LEFT(m,3)=INDEX(k,z)),p),"""")))))"
            ws.Range("U1").Formula2 = "=LET(t,$A$2:.$A$999,u,LEFT(t,3),v,UNIQUE(u),w,MAX(MAP(v,LAMBDA(x,SUM(--(u=x))))),VSTACK(""Pos""&SEQUENCE(,w),MAKEARRAY(ROWS(v),w,LAMBDA(y,z,IFERROR(INDEX(FILTER(t,u=INDEX(v,y)),z),"""")))))"
            ws.Range("B18").Value = "Excel-VSTO-Toolbox"
            ws.Range("B19").Value = "Formel-Demo"
            ws.Range("B20").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("B18:B20").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
    Private Sub CreateFormulaDemo_5()
        Dim ws = CreateFreshSheet("Demo_Formel_5")
        Dim app As Excel.Application = ws.Application
        Try
            app.ScreenUpdating = False
            app.EnableEvents = False
            ws.Range("A1").Value = "Ferien BW"
            ws.Range("B1").Value = "Von"
            ws.Range("C1").Value = "Bis"
            ws.Range("A2").Value = "Ostern"
            ws.Range("B2").Value = "25.03.2027"
            ws.Range("C2").Value = "03.04.2027"
            ws.Range("A3").Value = "Pfingsten"
            ws.Range("B3").Value = "18.05.2027"
            ws.Range("C3").Value = "29.05.2027"
            ws.Range("A4").Value = "Sommer"
            ws.Range("B4").Value = "29.07.2027"
            ws.Range("C4").Value = "11.09.2027"
            ws.Range("A5").Value = "Herbst"
            ws.Range("B5").Value = "02.11.2027"
            ws.Range("C5").Value = "06.11.2027"
            ws.Range("A6").Value = "Weihnachten"
            ws.Range("B6").Value = "23.12.2027"
            ws.Range("C6").Value = "08.01.2028"
            Dim lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, ws.Range("A1:C6"),, Excel.XlYesNoGuess.xlYes)
            lo.Name = "FBW"
            lo.TableStyle = "TableStyleMedium2"
            ws.Range("E1").Value = "Datum"
            ws.Range("F1").Value = "Tag"
            ws.Range("E2").Formula2 = "=LET(v,FBW[Von],b,FBW[Bis],n,b-v+1,m,MAX(n),d,TOCOL(IF(SEQUENCE(m,,0)<=TRANSPOSE(n-1),TRANSPOSE(v)+SEQUENCE(m,,0),NA()),3,TRUE),HSTACK(d,TEXT(d,""TTTT"")))"
            ws.Range("H1").Formula2 = "=VSTACK({""Datum"",""Tag""},LET(v,FBW[Von],b,FBW[Bis],n,b-v+1,m,MAX(n),d,TOCOL(IF(SEQUENCE(m,,0)<=TRANSPOSE(n-1),TRANSPOSE(v)+SEQUENCE(m,,0),NA()),3,TRUE),HSTACK(d,TEXT(d,""TTTT""))))"
            ws.Range("E2:E90").NumberFormat = "m/d/yyyy"
            ws.Range("H2:H90").NumberFormat = "m/d/yyyy"
            ws.Range("K1").Value = "Excel-VSTO-Toolbox"
            ws.Range("K2").Value = "Formel-Demo"
            ws.Range("K3").Value = "https://github.com/rstsu/Excel-VSTO-Toolbox"
            With ws.Range("K1:K3").Font
                .ColorIndex = 16
                .Size = 8
                .Italic = True
            End With
        Finally
            app.ScreenUpdating = True
            app.EnableEvents = True
        End Try
        FormatSheet(ws)
    End Sub
End Class