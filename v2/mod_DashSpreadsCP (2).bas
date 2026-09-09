Attribute VB_Name = "mod_DashSpreadsCP"
Option Explicit

' ============================================================
' mod_DashSpreadsCP
' Dashboard de spreads de colocacion - instrumentos de corto plazo
' Hoja origen: DB-Historica CP (detecta columnas por encabezado)
' Macro a ejecutar: CrearDashboardSpreadsCP
' ============================================================

Private Const SH_DASH As String = "Dashboard CP"
Private Const N_MESES As Long = 180
Private Const N_PTOS  As Long = 150

Sub CrearDashboardSpreadsCP()
    Dim wb As Workbook, db As Worksheet, ws As Worksheet
    Dim hdr As Long, lastR As Long, i As Long
    Dim cFec As Long, cEmi As Long, cSec As Long, cRat As Long, cSpr As Long, cPlz As Long
    Dim usaPb As Boolean, factor As Double
    Dim rats As Variant, cols As Variant, lnk As String

    Set wb = ThisWorkbook
    Set db = HojaDB(wb)
    If db Is Nothing Then MsgBox "No encontre la hoja DB-Historica CP.", vbExclamation: Exit Sub
    hdr = FilaEncabezado(db)
    If hdr = 0 Then MsgBox "No encontre la fila de encabezados (columna 'Emisor').", vbExclamation: Exit Sub

    cFec = BuscarCol(db, hdr, "emision")
    cEmi = BuscarCol(db, hdr, "emisor")
    cSec = BuscarCol(db, hdr, "sector")
    cRat = BuscarCol(db, hdr, "rating")
    cPlz = BuscarCol(db, hdr, "os vt")
    cSpr = BuscarCol(db, hdr, "spread (pb)")
    If cSpr = 0 Then cSpr = BuscarCol(db, hdr, "spread pb")
    usaPb = (cSpr > 0)
    If cSpr = 0 Then cSpr = BuscarCol(db, hdr, "spread")
    factor = IIf(usaPb, 1, 10000)
    If cFec * cEmi * cSec * cRat * cSpr = 0 Then
        MsgBox "Faltan columnas. Emision=" & cFec & " Emisor=" & cEmi & " Sector=" & cSec & _
               " Rating=" & cRat & " Spread=" & cSpr, vbExclamation: Exit Sub
    End If
    If cPlz = 0 Then cPlz = cRat

    lastR = db.Cells(db.Rows.Count, cEmi).End(xlUp).Row
    DefNombre wb, "dbFecha", db.Name, ColL(cFec), hdr + 1, lastR
    DefNombre wb, "dbEmisor", db.Name, ColL(cEmi), hdr + 1, lastR
    DefNombre wb, "dbSector", db.Name, ColL(cSec), hdr + 1, lastR
    DefNombre wb, "dbRating", db.Name, ColL(cRat), hdr + 1, lastR
    DefNombre wb, "dbSpread", db.Name, ColL(cSpr), hdr + 1, lastR
    DefNombre wb, "dbPlazo", db.Name, ColL(cPlz), hdr + 1, lastR

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(SH_DASH).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = wb.Worksheets.Add(After:=db)
    ws.Name = SH_DASH
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 8
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Rows(1).RowHeight = 22
    ws.Rows(6).RowHeight = 20

    With ws.Range("A1:J1")
        .Merge
        .Value = "  Spreads de colocacion - instrumentos de corto plazo"
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 11
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A2").Value = "  Promedio mensual del spread por rating - PEN - DB Historica CP"
    ws.Range("A2").Font.Color = RGB(110, 110, 110)

    Dim et As Variant
    et = Array("DESDE", "HASTA", "SECTOR", "EMISOR 1", "EMISOR 2", "PLAZO MIN", "PLAZO MAX")
    For i = 0 To 6
        With ws.Cells(3, 2 + i)
            .Value = et(i)
            .Font.Size = 7
            .Font.Color = RGB(140, 140, 140)
        End With
        ws.Columns(2 + i).ColumnWidth = IIf(i >= 3 And i <= 4, 24, 13)
    Next i

    ws.Range("B4").Formula = "=EDATE(TODAY(),-12)"
    ws.Range("C4").Formula = "=TODAY()"
    ws.Range("B4:C4").NumberFormat = "dd/mm/yyyy"
    ws.Range("D4").Value = "(todos)"
    ws.Range("E4").Value = "(ninguno)"
    ws.Range("F4").Value = "(ninguno)"
    ws.Range("G4").Value = 0
    ws.Range("H4").Value = 100
    With ws.Range("B4:H4")
        .Interior.Color = RGB(255, 252, 232)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 212, 181)
        .HorizontalAlignment = xlLeft
    End With
    ws.Range("A6").Value = "  RATINGS"
    ws.Range("A6").Font.Size = 7
    ws.Range("A6").Font.Color = RGB(140, 140, 140)
    ws.Columns(1).ColumnWidth = 12

    ws.Range("BP2").Value = "(todos)"
    ws.Range("BP3").Formula2 = "=SORT(UNIQUE(FILTER(dbSector,dbSector<>"""")))"
    ws.Range("BR2").Value = "(ninguno)"
    ws.Range("BR3").Formula2 = "=SORT(UNIQUE(FILTER(dbEmisor,(dbEmisor<>"""")*IF($D$4=""(todos)"",1,dbSector=$D$4))))"
    wb.Names.Add Name:="lstSectores", RefersTo:="='" & SH_DASH & "'!$BP$2:INDEX('" & SH_DASH & "'!$BP$2:$BP$300,COUNTA('" & SH_DASH & "'!$BP$2:$BP$300))"
    wb.Names.Add Name:="lstEmisores", RefersTo:="='" & SH_DASH & "'!$BR$2:INDEX('" & SH_DASH & "'!$BR$2:$BR$900,COUNTA('" & SH_DASH & "'!$BR$2:$BR$900))"
    PonerLista ws.Range("D4"), "=lstSectores"
    PonerLista ws.Range("E4"), "=lstEmisores"
    PonerLista ws.Range("F4"), "=lstEmisores"

    rats = Array("CP-1+", "CP-1", "CP-1-", "CP-2+", "CP-2")
    cols = Array(RGB(31, 59, 87), RGB(46, 127, 140), RGB(176, 135, 59), RGB(122, 92, 142), RGB(140, 140, 140))

    For i = 0 To 4
        lnk = "$" & ColL(62 + i) & "$1"
        With ws.CheckBoxes.Add(ws.Cells(6, 2).Left + i * 78, ws.Cells(6, 2).Top - 1, 74, 18)
            .Caption = rats(i)
            .LinkedCell = "'" & SH_DASH & "'!" & lnk
            .Value = IIf(i <= 2, xlOn, xlOff)
            On Error Resume Next
            .Characters.Font.Name = "Arial"
            .Characters.Font.Size = 8
            On Error GoTo 0
            .Name = "chk" & i
        End With
    Next i

    ws.Range("AH1").Value = factor
    ws.Range("AG1").Formula = "=""Spreads corto plazo - ""&TEXT($B$4,""mmm-aa"")&"" a ""&TEXT($C$4,""mmm-aa"")&IF($D$4=""(todos)"","""","" - ""&$D$4)"

    For i = 0 To 4
        ws.Cells(2, 35 + i).Value = rats(i)
        ws.Cells(1, 35 + i).Formula = "=IF($" & ColL(62 + i) & "$1,1,0)"
    Next i

    ws.Range("AH3").Formula = "=DATE(YEAR($B$4),MONTH($B$4),1)"
    ws.Range("AH4").Formula = "=IF(AH3=0,0,IF(EDATE(AH3,1)>$C$4,0,EDATE(AH3,1)))"
    ws.Range("AH4").AutoFill Destination:=ws.Range("AH4:AH" & (2 + N_MESES))

    ws.Range("AI3").Formula = "=IF($AH3=0,0,IF(AI$1=0,0,IFERROR($AH$1*AVERAGEIFS(dbSpread," & _
        "dbFecha,"">=""&$AH3,dbFecha,""<""&EDATE($AH3,1),dbRating,AI$2," & _
        "dbSector,IF($D$4=""(todos)"",""<>"",$D$4),dbPlazo,"">=""&$G$4,dbPlazo,""<=""&$H$4),0)))"
    ws.Range("AI3").Copy ws.Range("AI3:AM" & (2 + N_MESES))

    ws.Range("AO3").Formula = "=IF($AH3=0,0,IF(AI$1=0,0,COUNTIFS(dbRating,AI$2," & _
        "dbFecha,"">=""&$AH3,dbFecha,""<""&EDATE($AH3,1)," & _
        "dbSector,IF($D$4=""(todos)"",""<>"",$D$4),dbPlazo,"">=""&$G$4,dbPlazo,""<=""&$H$4)))"
    ws.Range("AO3").Copy ws.Range("AO3:AS" & (2 + N_MESES))

    Dim x As String, y As String, cn As String
    For i = 0 To 4
        x = ColL(47 + i * 2): y = ColL(48 + i * 2): cn = ColL(41 + i)
        ws.Range(x & "3").Formula2 = "=IFERROR(INDEX(FILTER($AH$3:$AH$" & (2 + N_MESES) & _
            ",$" & cn & "$3:$" & cn & "$" & (2 + N_MESES) & ">0),ROWS($" & x & "$3:$" & x & "3)),NA())"
        ws.Range(y & "3").Formula2 = "=IFERROR(INDEX(FILTER($" & ColL(35 + i) & "$3:$" & ColL(35 + i) & "$" & (2 + N_MESES) & _
            ",$" & cn & "$3:$" & cn & "$" & (2 + N_MESES) & ">0),ROWS($" & y & "$3:$" & y & "3)),NA())"
        ws.Range(x & "3:" & y & "3").Copy ws.Range(x & "3:" & y & (2 + N_MESES))
    Next i

    Puntos ws, 58, "$E$4"
    Puntos ws, 60, "$F$4"
    ws.Columns("AG:BS").Hidden = True
    Application.Calculate

    Dim ch As Chart, sr As Series
    Set ch = ws.Shapes.AddChart2(-1, xlXYScatterLinesNoMarkers, ws.Cells(8, 1).Left + 5, ws.Cells(8, 1).Top, 920, 430).Chart
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    For i = 0 To 4
        Set sr = ch.SeriesCollection.NewSeries
        sr.Name = "='" & SH_DASH & "'!" & ws.Cells(2, 35 + i).Address
        sr.XValues = "='" & SH_DASH & "'!" & ws.Range(ws.Cells(3, 47 + i * 2), ws.Cells(2 + N_MESES, 47 + i * 2)).Address
        sr.Values = "='" & SH_DASH & "'!" & ws.Range(ws.Cells(3, 48 + i * 2), ws.Cells(2 + N_MESES, 48 + i * 2)).Address
        sr.ChartType = xlXYScatterLinesNoMarkers
        sr.Format.Line.ForeColor.RGB = cols(i)
        sr.Format.Line.Weight = 2.25
        sr.Smooth = False
    Next i

    Set sr = ch.SeriesCollection.NewSeries
    sr.Name = "='" & SH_DASH & "'!$E$4"
    sr.XValues = "='" & SH_DASH & "'!$BF$3:$BF$" & (2 + N_PTOS)
    sr.Values = "='" & SH_DASH & "'!$BG$3:$BG$" & (2 + N_PTOS)
    sr.ChartType = xlXYScatter
    sr.MarkerStyle = xlMarkerStyleCircle
    sr.MarkerSize = 8
    sr.MarkerBackgroundColor = RGB(212, 12, 12)
    sr.MarkerForegroundColor = RGB(255, 255, 255)

    Set sr = ch.SeriesCollection.NewSeries
    sr.Name = "='" & SH_DASH & "'!$F$4"
    sr.XValues = "='" & SH_DASH & "'!$BH$3:$BH$" & (2 + N_PTOS)
    sr.Values = "='" & SH_DASH & "'!$BI$3:$BI$" & (2 + N_PTOS)
    sr.ChartType = xlXYScatter
    sr.MarkerStyle = xlMarkerStyleSquare
    sr.MarkerSize = 8
    sr.MarkerBackgroundColor = RGB(60, 60, 60)
    sr.MarkerForegroundColor = RGB(255, 255, 255)

    ch.PlotVisibleOnly = False
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.ChartArea.Font.Name = "Arial"
    ch.HasTitle = True
    ch.ChartTitle.Formula = "='" & SH_DASH & "'!$AG$1"
    ch.ChartTitle.Font.Size = 11
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Left = 10
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Legend.Font.Size = 8
    ch.Legend.Format.Line.Visible = msoFalse
    With ch.Axes(xlCategory)
        .TickLabels.NumberFormat = "mmm-yy"
        .TickLabels.Font.Size = 8
        .TickLabels.Font.Color = RGB(90, 90, 90)
        .HasMajorGridlines = False
        .MajorTickMark = xlNone
        .Format.Line.ForeColor.RGB = RGB(214, 214, 214)
    End With
    With ch.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Spread (pb)"
        .AxisTitle.Font.Size = 8
        .AxisTitle.Font.Color = RGB(110, 110, 110)
        .TickLabels.Font.Size = 8
        .TickLabels.Font.Color = RGB(90, 90, 90)
        .MajorTickMark = xlNone
        .Format.Line.Visible = msoFalse
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(234, 234, 234)
        .MajorGridlines.Format.Line.Weight = 0.75
    End With

    ws.Range("A32").Value = "Fuente: DB Historica CP. Cada punto de la linea es el promedio de las emisiones del mes para ese rating; los meses sin emisiones no interrumpen la linea."
    ws.Range("A32").Font.Size = 7
    ws.Range("A32").Font.Color = RGB(140, 140, 140)

    ws.Activate
    ws.Range("B4").Select
    MsgBox "Listo." & vbCrLf & "Hoja: " & db.Name & " (encabezados fila " & hdr & ", datos " & hdr + 1 & "-" & lastR & ")" & vbCrLf & _
           "Emision=" & ColL(cFec) & "  Emisor=" & ColL(cEmi) & "  Sector=" & ColL(cSec) & "  Rating=" & ColL(cRat) & _
           "  Spread=" & ColL(cSpr) & "  Plazo=" & ColL(cPlz) & vbCrLf & _
           "Spread leido en " & IIf(usaPb, "pb (factor 1)", "% (factor 10000)") & " - REVISAR si no cuadra.", vbInformation
End Sub

Private Sub Puntos(ws As Worksheet, cX As Long, celEmi As String)
    Dim x As String, y As String, cond As String
    x = ColL(cX): y = ColL(cX + 1)
    cond = "(dbEmisor=" & celEmi & ")*(dbFecha>=$B$4)*(dbFecha<=$C$4)*(dbPlazo>=$G$4)*(dbPlazo<=$H$4)"
    ws.Range(x & "3").Formula2 = "=IFERROR(INDEX(FILTER(dbFecha," & cond & "),ROWS($" & x & "$3:$" & x & "3)),NA())"
    ws.Range(y & "3").Formula2 = "=IFERROR($AH$1*INDEX(FILTER(dbSpread," & cond & "),ROWS($" & y & "$3:$" & y & "3)),NA())"
    ws.Range(x & "3:" & y & "3").Copy ws.Range(x & "3:" & y & (2 + N_PTOS))
    ws.Range(x & "3:" & x & (2 + N_PTOS)).NumberFormat = "dd/mm/yyyy"
End Sub

Private Sub PonerLista(c As Range, lst As String)
    With c.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=lst
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Sub DefNombre(wb As Workbook, nm As String, hoja As String, col As String, f0 As Long, f1 As Long)
    On Error Resume Next
    wb.Names(nm).Delete
    On Error GoTo 0
    wb.Names.Add Name:=nm, RefersTo:="='" & hoja & "'!$" & col & "$" & f0 & ":$" & col & "$" & f1
End Sub

Private Function HojaDB(wb As Workbook) As Worksheet
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If InStr(1, sh.Name, "DB-Hist", vbTextCompare) > 0 Then
            Set HojaDB = sh
            Exit Function
        End If
    Next sh
End Function

Private Function FilaEncabezado(db As Worksheet) As Long
    Dim r As Long, c As Long
    For r = 1 To 15
        For c = 1 To 60
            If LCase$(Trim$(CStr(db.Cells(r, c).Value))) = "emisor" Then
                FilaEncabezado = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function BuscarCol(db As Worksheet, hdr As Long, txt As String) As Long
    Dim c As Long, s As String
    For c = 1 To 60
        s = LCase$(Trim$(CStr(db.Cells(hdr, c).Value)))
        s = Replace(s, ChrW(225), "a")
        s = Replace(s, ChrW(233), "e")
        s = Replace(s, ChrW(237), "i")
        s = Replace(s, ChrW(243), "o")
        s = Replace(s, ChrW(250), "u")
        s = Replace(s, Chr(10), " ")
        If Len(s) > 0 And InStr(1, s, txt, vbTextCompare) > 0 Then
            BuscarCol = c
            Exit Function
        End If
    Next c
End Function

Private Function ColL(i As Long) As String
    ColL = Split(Cells(1, i).Address(True, False), "$")(0)
End Function
