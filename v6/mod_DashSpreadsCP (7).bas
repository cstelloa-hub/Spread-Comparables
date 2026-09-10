Attribute VB_Name = "mod_DashSpreadsCP"
Option Explicit

' ============================================================
' mod_DashSpreadsCP
' Dashboard de colocaciones de corto plazo - DB Historica CP
' Tres graficos apilados (spread, tasa final, tasa base) con
' filtros compartidos arriba y casillas por rating.
'
' Macros:
'   CrearDashboardSpreadsCP  -> arma la hoja Dashboard CP
'   DiagnosticoDashCP        -> revisa la base y reporta columnas
' ============================================================

Private Const SH_DASH As String = "Dashboard CP"
Private Const N_MESES As Long = 180
Private Const N_PTOS  As Long = 150

' Columnas de la hoja de datos para las tasas (cambiar aqui si se mueven)
' True: el panel de spread se muestra en puntos basicos. False: en porcentaje.
Private Const SPREAD_EN_PB As Boolean = True

Private Const COL_TASA_FINAL As String = "Q"
Private Const COL_TASA_BASE  As String = "R"

Sub CrearDashboardSpreadsCP()
    Dim wb As Workbook, db As Worksheet, ws As Worksheet
    Dim hdr As Long, lastR As Long, i As Long, m As Long, b As Long
    Dim cFec As Long, cEmi As Long, cSec As Long, cRat As Long, cSpr As Long, cPlz As Long
    Dim cTF As Long, cTB As Long, cDiv As Long
    Dim usaPb As Boolean, usaPlz As Boolean
    Dim factor As Double, facTF As Double, facTB As Double
    Dim rats As Variant, cols As Variant, mets As Variant, facs As Variant
    Dim lnk As String, sPI As String, sPF As String
    Dim x As String, y As String, cn As String, pr As String, nm As String, fc As String

    Set wb = ThisWorkbook
    Set db = HojaDB(wb)
    If db Is Nothing Then MsgBox "No encontre la hoja DB-Historica CP.", vbExclamation: Exit Sub
    hdr = FilaEncabezado(db)
    If hdr = 0 Then MsgBox "No encontre la fila de encabezados (columna 'Emisor').", vbExclamation: Exit Sub

    cFec = BuscarCol(db, hdr, "emision")
    cEmi = BuscarCol(db, hdr, "emisor")
    cSec = BuscarCol(db, hdr, "sector")
    cRat = BuscarCol(db, hdr, "rating")
    cDiv = BuscarCol(db, hdr, "divisa")
    If cDiv = 0 Then cDiv = BuscarCol(db, hdr, "moneda")
    cPlz = BuscarCol(db, hdr, "anos")
    If cPlz = 0 Then cPlz = BuscarCol(db, hdr, "os vt")
    If cPlz = 0 Then cPlz = BuscarCol(db, hdr, "plazo")
    If cPlz = 0 Then cPlz = BuscarCol(db, hdr, "duracion")
    cSpr = BuscarCol(db, hdr, "spread (pb)")
    If cSpr = 0 Then cSpr = BuscarCol(db, hdr, "spread pb")
    usaPb = (cSpr > 0)
    If cSpr = 0 Then cSpr = BuscarCol(db, hdr, "spread")
    usaPlz = (cPlz > 0)
    If cFec * cEmi * cSec * cRat * cSpr = 0 Then
        MsgBox "Faltan columnas. Emision=" & cFec & " Emisor=" & cEmi & " Sector=" & cSec & _
               " Rating=" & cRat & " Spread=" & cSpr, vbExclamation: Exit Sub
    End If

    cTF = db.Range(COL_TASA_FINAL & "1").Column
    cTB = db.Range(COL_TASA_BASE & "1").Column
    lastR = db.Cells(db.Rows.Count, cEmi).End(xlUp).Row
    facTF = FactorTasa(db, hdr, lastR, cTF)
    facTB = FactorTasa(db, hdr, lastR, cTB)
    factor = FactorSpread(db, hdr, lastR, cSpr)

    DefNombre wb, "dbFecha", db.Name, ColL(cFec), hdr + 1, lastR
    DefNombre wb, "dbEmisor", db.Name, ColL(cEmi), hdr + 1, lastR
    DefNombre wb, "dbSector", db.Name, ColL(cSec), hdr + 1, lastR
    DefNombre wb, "dbRating", db.Name, ColL(cRat), hdr + 1, lastR
    DefNombre wb, "dbSpread", db.Name, ColL(cSpr), hdr + 1, lastR
    DefNombre wb, "dbTFinal", db.Name, ColL(cTF), hdr + 1, lastR
    DefNombre wb, "dbTBase", db.Name, ColL(cTB), hdr + 1, lastR
    If usaPlz Then DefNombre wb, "dbPlazo", db.Name, ColL(cPlz), hdr + 1, lastR
    If cDiv > 0 Then DefNombre wb, "dbDivisa", db.Name, ColL(cDiv), hdr + 1, lastR

    If usaPlz Then
        sPI = ",dbPlazo,"">=""&$G$4,dbPlazo,""<=""&$H$4"
        sPF = "*(dbPlazo>=$G$4)*(dbPlazo<=$H$4)"
    End If
    If cDiv > 0 Then
        sPI = sPI & ",dbDivisa,IF($I$4=""(todas)"",""<>"",$I$4)"
        sPF = sPF & "*IF($I$4=""(todas)"",1,dbDivisa=$I$4)"
    End If

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
        .Value = "  Colocaciones de corto plazo - spread, tasa final y tasa base"
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 11
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A2").Value = "  Promedio mensual por rating - PEN - DB Historica CP"
    ws.Range("A2").Font.Color = RGB(110, 110, 110)

    Dim et As Variant
    et = Array("DESDE", "HASTA", "SECTOR", "EMISOR 1", "EMISOR 2", "PLAZO MIN", "PLAZO MAX", "MONEDA")
    For i = 0 To 7
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
    ws.Range("I4").Value = "PEN"
    With ws.Range("B4:I4")
        .Interior.Color = RGB(255, 252, 232)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 212, 181)
        .HorizontalAlignment = xlLeft
    End With
    If Not usaPlz Then
        ws.Range("G4:H4").Value = "sin filtro"
        ws.Range("G4:H4").Font.Color = RGB(150, 150, 150)
    End If
    ws.Range("A6").Value = "  RATINGS"
    ws.Range("A6").Font.Size = 7
    ws.Range("A6").Font.Color = RGB(140, 140, 140)
    ws.Columns(1).ColumnWidth = 12

    ws.Cells(2, 110).Value = "(todos)"
    ws.Cells(3, 110).Formula2 = "=SORT(UNIQUE(FILTER(dbSector,dbSector<>"""")))"
    ws.Cells(2, 112).Value = "(ninguno)"
    ws.Cells(3, 112).Formula2 = "=SORT(UNIQUE(FILTER(dbEmisor,(dbEmisor<>"""")*IF($D$4=""(todos)"",1,dbSector=$D$4))))"
    wb.Names.Add Name:="lstSectores", RefersTo:="='" & SH_DASH & "'!$" & ColL(110) & "$2:INDEX('" & SH_DASH & "'!$" & ColL(110) & "$2:$" & ColL(110) & "$300,COUNTA('" & SH_DASH & "'!$" & ColL(110) & "$2:$" & ColL(110) & "$300))"
    wb.Names.Add Name:="lstEmisores", RefersTo:="='" & SH_DASH & "'!$" & ColL(112) & "$2:INDEX('" & SH_DASH & "'!$" & ColL(112) & "$2:$" & ColL(112) & "$900,COUNTA('" & SH_DASH & "'!$" & ColL(112) & "$2:$" & ColL(112) & "$900))"
    ws.Cells(2, 120).Value = "(todas)"
    ws.Cells(3, 120).Formula2 = "=SORT(UNIQUE(FILTER(dbDivisa,dbDivisa<>"""")))"
    wb.Names.Add Name:="lstMonedas", RefersTo:="='" & SH_DASH & "'!$" & ColL(120) & "$2:INDEX('" & SH_DASH & "'!$" & ColL(120) & "$2:$" & ColL(120) & "$50,COUNTA('" & SH_DASH & "'!$" & ColL(120) & "$2:$" & ColL(120) & "$50))"
    If cDiv > 0 Then PonerLista ws.Range("I4"), "=lstMonedas"
    PonerLista ws.Range("D4"), "=lstSectores"
    PonerLista ws.Range("E4"), "=lstEmisores"
    PonerLista ws.Range("F4"), "=lstEmisores"

    rats = Array("CP-1+", "CP-1", "CP-1-", "CP-2+", "CP-2")
    cols = Array(RGB(31, 59, 87), RGB(46, 127, 140), RGB(176, 135, 59), RGB(122, 92, 142), RGB(140, 140, 140))
    mets = Array("dbSpread", "dbTFinal", "dbTBase")

    For i = 0 To 4
        lnk = "$" & ColL(114 + i) & "$1"
        With ws.CheckBoxes.Add(ws.Cells(6, 2).Left + i * 78, ws.Cells(6, 2).Top - 1, 74, 18)
            .Caption = rats(i)
            .LinkedCell = "'" & SH_DASH & "'!" & lnk
            .Value = IIf(i <= 2, xlOn, xlOff)
            .Name = "chk" & i
        End With
    Next i

    ws.Range("AF1").Value = factor
    ws.Range("AF2").Value = facTF
    ws.Range("AF3").Value = facTB
    ws.Range("AE1").Formula = "=""Spread de colocacion (" & IIf(SPREAD_EN_PB, "pb", "%") & ") - ""&TEXT($B$4,""mmm-aa"")&"" a ""&TEXT($C$4,""mmm-aa"")&IF($D$4=""(todos)"","""","" - ""&$D$4)"
    ws.Range("AE2").Formula = "=""Tasa final - ""&TEXT($B$4,""mmm-aa"")&"" a ""&TEXT($C$4,""mmm-aa"")&IF($D$4=""(todos)"","""","" - ""&$D$4)"
    ws.Range("AE3").Formula = "=""Tasa base (curva de mercado, no depende del rating) - ""&TEXT($B$4,""mmm-aa"")&"" a ""&TEXT($C$4,""mmm-aa"")&IF($D$4=""(todos)"","""","" - ""&$D$4)"

    ws.Range("AH3").Formula = "=DATE(YEAR($B$4),MONTH($B$4),1)"
    ws.Range("AH4").Formula = "=IF(AH3=0,0,IF(EDATE(AH3,1)>$C$4,0,EDATE(AH3,1)))"
    ws.Range("AH4").AutoFill Destination:=ws.Range("AH4:AH" & (2 + N_MESES))

    For m = 0 To 2
        b = 36 + m * 24
        nm = CStr(mets(m))
        fc = "$AF$" & (m + 1)
        If m < 2 Then
            For i = 0 To 4
                ws.Cells(2, b + i).Value = rats(i)
                ws.Cells(1, b + i).Formula = "=IF($" & ColL(114 + i) & "$1,1,0)"
            Next i

            ws.Cells(3, b).Formula = "=IF($AH3=0,0,IF(" & ColL(b) & "$1=0,0,IFERROR(" & fc & "*AVERAGEIFS(" & nm & _
                ",dbFecha,"">=""&$AH3,dbFecha,""<""&EDATE($AH3,1),dbRating," & ColL(b) & "$2," & _
                "dbSector,IF($D$4=""(todos)"",""<>"",$D$4)" & sPI & "),0)))"
            ws.Cells(3, b).Copy ws.Range(ws.Cells(3, b), ws.Cells(2 + N_MESES, b + 4))

            ws.Cells(3, b + 5).Formula = "=IF($AH3=0,0,IF(" & ColL(b) & "$1=0,0,COUNTIFS(dbRating," & ColL(b) & "$2," & _
                "dbFecha,"">=""&$AH3,dbFecha,""<""&EDATE($AH3,1)," & _
                "dbSector,IF($D$4=""(todos)"",""<>"",$D$4)" & sPI & ")))"
            ws.Cells(3, b + 5).Copy ws.Range(ws.Cells(3, b + 5), ws.Cells(2 + N_MESES, b + 9))

            For i = 0 To 4
                x = ColL(b + 10 + i * 2): y = ColL(b + 11 + i * 2)
                cn = ColL(b + 5 + i): pr = ColL(b + i)
                ws.Range(x & "3").Formula2 = "=IFERROR(INDEX(FILTER($AH$3:$AH$" & (2 + N_MESES) & _
                    ",$" & cn & "$3:$" & cn & "$" & (2 + N_MESES) & ">0),ROWS($" & x & "$3:$" & x & "3)),NA())"
                ws.Range(y & "3").Formula2 = "=IFERROR(INDEX(FILTER($" & pr & "$3:$" & pr & "$" & (2 + N_MESES) & _
                    ",$" & cn & "$3:$" & cn & "$" & (2 + N_MESES) & ">0),ROWS($" & y & "$3:$" & y & "3)),NA())"
                ws.Range(x & "3:" & y & "3").Copy ws.Range(x & "3:" & y & (2 + N_MESES))
            Next i
        Else
            ws.Cells(2, b).Value = "Mercado"
            ws.Cells(3, b).Formula = "=IF($AH3=0,0,IFERROR(" & fc & "*AVERAGEIFS(" & nm & _
                ",dbFecha,"">=""&$AH3,dbFecha,""<""&EDATE($AH3,1)," & _
                "dbSector,IF($D$4=""(todos)"",""<>"",$D$4)" & sPI & "),0))"
            ws.Cells(3, b).Copy ws.Range(ws.Cells(3, b), ws.Cells(2 + N_MESES, b))

            ws.Cells(3, b + 5).Formula = "=IF($AH3=0,0,COUNTIFS(dbFecha,"">=""&$AH3," & _
                "dbFecha,""<""&EDATE($AH3,1),dbSector,IF($D$4=""(todos)"",""<>"",$D$4)" & sPI & "))"
            ws.Cells(3, b + 5).Copy ws.Range(ws.Cells(3, b + 5), ws.Cells(2 + N_MESES, b + 5))

            x = ColL(b + 10): y = ColL(b + 11): cn = ColL(b + 5): pr = ColL(b)
            ws.Range(x & "3").Formula2 = "=IFERROR(INDEX(FILTER($AH$3:$AH$" & (2 + N_MESES) & _
                ",$" & cn & "$3:$" & cn & "$" & (2 + N_MESES) & ">0),ROWS($" & x & "$3:$" & x & "3)),NA())"
            ws.Range(y & "3").Formula2 = "=IFERROR(INDEX(FILTER($" & pr & "$3:$" & pr & "$" & (2 + N_MESES) & _
                ",$" & cn & "$3:$" & cn & "$" & (2 + N_MESES) & ">0),ROWS($" & y & "$3:$" & y & "3)),NA())"
            ws.Range(x & "3:" & y & "3").Copy ws.Range(x & "3:" & y & (2 + N_MESES))
        End If

        Puntos ws, b + 20, "$E$4", sPF, nm, fc
        Puntos ws, b + 22, "$F$4", sPF, nm, fc
    Next m

    ws.Range(ws.Columns(31), ws.Columns(124)).EntireColumn.Hidden = True
    Application.Calculate

    Dim ch As Chart, sr As Series, nSer As Long, r0 As Long
    For m = 0 To 2
        b = 36 + m * 24
        r0 = 8 + m * 34
        Set ch = ws.Shapes.AddChart2(-1, xlXYScatterLinesNoMarkers, ws.Cells(r0, 1).Left + 5, _
                 ws.Cells(r0, 1).Top, ws.Cells(r0, 12).Left - ws.Cells(r0, 1).Left - 12, _
                 ws.Cells(r0 + 32, 1).Top - ws.Cells(r0, 1).Top).Chart
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
        Loop

        nSer = IIf(m < 2, 5, 1)
        For i = 0 To nSer - 1
            Set sr = ch.SeriesCollection.NewSeries
            sr.Name = "='" & SH_DASH & "'!" & ws.Cells(2, b + i).Address
            sr.XValues = "='" & SH_DASH & "'!" & ws.Range(ws.Cells(3, b + 10 + i * 2), ws.Cells(2 + N_MESES, b + 10 + i * 2)).Address
            sr.Values = "='" & SH_DASH & "'!" & ws.Range(ws.Cells(3, b + 11 + i * 2), ws.Cells(2 + N_MESES, b + 11 + i * 2)).Address
            sr.ChartType = xlXYScatterLinesNoMarkers
            sr.Format.Line.ForeColor.RGB = IIf(m < 2, cols(i), RGB(31, 59, 87))
            sr.Format.Line.Weight = 2.25
            sr.Smooth = False
        Next i

        Set sr = ch.SeriesCollection.NewSeries
        sr.Name = "='" & SH_DASH & "'!$E$4"
        sr.XValues = "='" & SH_DASH & "'!" & ws.Range(ws.Cells(3, b + 20), ws.Cells(2 + N_PTOS, b + 20)).Address
        sr.Values = "='" & SH_DASH & "'!" & ws.Range(ws.Cells(3, b + 21), ws.Cells(2 + N_PTOS, b + 21)).Address
        sr.ChartType = xlXYScatter
        sr.MarkerStyle = xlMarkerStyleCircle
        sr.MarkerSize = 8
        sr.MarkerBackgroundColor = RGB(212, 12, 12)
        sr.MarkerForegroundColor = RGB(255, 255, 255)

        Set sr = ch.SeriesCollection.NewSeries
        sr.Name = "='" & SH_DASH & "'!$F$4"
        sr.XValues = "='" & SH_DASH & "'!" & ws.Range(ws.Cells(3, b + 22), ws.Cells(2 + N_PTOS, b + 22)).Address
        sr.Values = "='" & SH_DASH & "'!" & ws.Range(ws.Cells(3, b + 23), ws.Cells(2 + N_PTOS, b + 23)).Address
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
        ch.ChartTitle.Formula = "='" & SH_DASH & "'!$AE$" & (m + 1)
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
            .AxisTitle.Text = IIf(m = 0, IIf(SPREAD_EN_PB, "Spread (pb)", "Spread (%)"), "Tasa (%)")
            .AxisTitle.Font.Size = 8
            .AxisTitle.Font.Color = RGB(110, 110, 110)
            .TickLabels.Font.Size = 8
            .TickLabels.Font.Color = RGB(90, 90, 90)
            .TickLabels.NumberFormat = IIf(m = 0, IIf(SPREAD_EN_PB, "0", "0.00"), "0.00")
            .MajorTickMark = xlNone
            .Format.Line.Visible = msoFalse
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(234, 234, 234)
            .MajorGridlines.Format.Line.Weight = 0.75
        End With

        Resumen ws, r0, m, CStr(mets(m)), "$AF$" & (m + 1), sPI, sPF, rats
    Next m

    ws.Range("A112").Value = "Fuente: DB Historica CP. Fechas, sector y plazo aplican a los tres graficos; el filtro de rating solo a spread y tasa final. La tasa base va como linea unica de mercado porque no depende del rating del emisor."
    ws.Range("A112").Font.Size = 7
    ws.Range("A112").Font.Color = RGB(140, 140, 140)

    ws.Activate
    ws.Range("B4").Select
    MsgBox "Listo." & vbCrLf & "Hoja: " & db.Name & " (encabezados fila " & hdr & ", datos " & hdr + 1 & "-" & lastR & ")" & vbCrLf & vbCrLf & _
           "Emision=" & ColL(cFec) & "  Emisor=" & ColL(cEmi) & "  Sector=" & ColL(cSec) & "  Rating=" & ColL(cRat) & vbCrLf & _
           "Spread=" & ColL(cSpr) & " [" & db.Cells(hdr, cSpr).Value & "]  max leido " & _
           Format(MaxCol(db, hdr, lastR, cSpr), "0.0000") & "  factor " & factor & _
           "  -> mostrado en " & IIf(SPREAD_EN_PB, "pb", "%") & vbCrLf & _
           "Tasa final=" & COL_TASA_FINAL & " [" & db.Cells(hdr, cTF).Value & "]  factor " & facTF & vbCrLf & _
           "Tasa base=" & COL_TASA_BASE & " [" & db.Cells(hdr, cTB).Value & "]  factor " & facTB & vbCrLf & _
           "Plazo=" & ColL(cPlz) & IIf(usaPlz, "", "  (no detectada, filtro desactivado)") & vbCrLf & _
           "Divisa=" & ColL(cDiv) & IIf(cDiv > 0, "  (filtro en I4, por defecto PEN)", "  (no detectada, filtro desactivado)") & vbCrLf & vbCrLf & _
           "REVISAR que los encabezados entre corchetes sean los correctos.", vbInformation
End Sub

Private Sub Resumen(ByVal ws As Worksheet, ByVal r0 As Long, ByVal m As Long, ByVal nm As String, _
                    ByVal fc As String, ByVal sPI As String, ByVal sPF As String, ByVal rats As Variant)
    Dim c As Long, r As Long, i As Long, nFil As Long
    Dim sec As String, fech As String, cond As String, flag As String, cel As String
    Dim fmt As String, tit As String

    c = 12                                   ' columna L
    sec = "dbSector,IF($D$4=""(todos)"",""<>"",$D$4)"
    fech = "dbFecha,"">=""&$B$4,dbFecha,""<=""&$C$4"
    fmt = IIf(m = 0, IIf(SPREAD_EN_PB, "0.0", "0.000"), "0.00")
    tit = IIf(m = 0, IIf(SPREAD_EN_PB, "RESUMEN - SPREAD (pb)", "RESUMEN - SPREAD (%)"), IIf(m = 1, "RESUMEN - TASA FINAL (%)", "RESUMEN - TASA BASE (%)"))

    r = r0 + 1
    With ws.Cells(r, c)
        .Value = tit
        .Font.Size = 7
        .Font.Bold = True
        .Font.Color = RGB(140, 140, 140)
    End With

    r = r + 1
    ws.Cells(r, c).Value = ""
    ws.Cells(r, c + 1).Value = "N"
    ws.Cells(r, c + 2).Value = "Ultimo"
    ws.Cells(r, c + 3).Value = "Prom"
    ws.Cells(r, c + 4).Value = "Min"
    ws.Cells(r, c + 5).Value = "Max"
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + 6))
        .Font.Size = 7
        .Font.Color = RGB(140, 140, 140)
        .HorizontalAlignment = xlRight
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(214, 214, 214)
    End With
    ws.Cells(r, c).HorizontalAlignment = xlLeft

    nFil = IIf(m < 2, 5, 1)
    For i = 0 To nFil - 1
        r = r + 1
        If m < 2 Then
            ws.Cells(r, c).Value = rats(i)
            flag = "$" & ColL(114 + i) & "$1"
            cel = "$" & ColL(c) & "$" & r
            cond = "(dbRating=" & cel & ")*(dbFecha>=$B$4)*(dbFecha<=$C$4)*IF($D$4=""(todos)"",1,dbSector=$D$4)" & sPF
            ws.Cells(r, c + 1).Formula = "=IF(" & flag & "=FALSE,"""",COUNTIFS(dbRating," & cel & "," & fech & "," & sec & sPI & "))"
            ws.Cells(r, c + 2).Formula = "=IF(" & flag & "=FALSE,"""",IFERROR(" & fc & "*XLOOKUP(MAX(FILTER(dbFecha," & cond & ")),FILTER(dbFecha," & cond & "),FILTER(" & nm & "," & cond & ")),""""))"
            ws.Cells(r, c + 3).Formula = "=IF(" & flag & "=FALSE,"""",IFERROR(" & fc & "*AVERAGEIFS(" & nm & ",dbRating," & cel & "," & fech & "," & sec & sPI & "),""""))"
            ws.Cells(r, c + 4).Formula = "=IF(" & flag & "=FALSE,"""",IFERROR(" & fc & "*MINIFS(" & nm & ",dbRating," & cel & "," & fech & "," & sec & sPI & "),""""))"
            ws.Cells(r, c + 5).Formula = "=IF(" & flag & "=FALSE,"""",IFERROR(" & fc & "*MAXIFS(" & nm & ",dbRating," & cel & "," & fech & "," & sec & sPI & "),""""))"
        Else
            ws.Cells(r, c).Value = "Mercado"
            cond = "(dbFecha>=$B$4)*(dbFecha<=$C$4)*IF($D$4=""(todos)"",1,dbSector=$D$4)" & sPF
            ws.Cells(r, c + 1).Formula = "=COUNTIFS(" & fech & "," & sec & sPI & ")"
            ws.Cells(r, c + 2).Formula = "=IFERROR(" & fc & "*XLOOKUP(MAX(FILTER(dbFecha," & cond & ")),FILTER(dbFecha," & cond & "),FILTER(" & nm & "," & cond & ")),"""")"
            ws.Cells(r, c + 3).Formula = "=IFERROR(" & fc & "*AVERAGEIFS(" & nm & "," & fech & "," & sec & sPI & "),"""")"
            ws.Cells(r, c + 4).Formula = "=IFERROR(" & fc & "*MINIFS(" & nm & "," & fech & "," & sec & sPI & "),"""")"
            ws.Cells(r, c + 5).Formula = "=IFERROR(" & fc & "*MAXIFS(" & nm & "," & fech & "," & sec & sPI & "),"""")"
        End If
        ws.Range(ws.Cells(r, c + 1), ws.Cells(r, c + 5)).NumberFormat = fmt
        ws.Cells(r, c + 1).NumberFormat = "0"
    Next i

    r = r + 1
    For i = 0 To 1
        r = r + 1
        cel = IIf(i = 0, "$E$4", "$F$4")
        cond = "(dbEmisor=" & cel & ")*(dbFecha>=$B$4)*(dbFecha<=$C$4)" & sPF
        ws.Cells(r, c).Formula = "=" & cel
        ws.Cells(r, c).Font.Bold = True
        ws.Cells(r, c + 1).Formula = "=IFERROR(COUNTIFS(dbEmisor," & cel & "," & fech & sPI & "),"""")"
        ws.Cells(r, c + 2).Formula = "=IFERROR(" & fc & "*XLOOKUP(MAX(FILTER(dbFecha," & cond & ")),FILTER(dbFecha," & cond & "),FILTER(" & nm & "," & cond & ")),"""")"
        ws.Cells(r, c + 3).Formula = "=IFERROR(" & fc & "*AVERAGEIFS(" & nm & ",dbEmisor," & cel & "," & fech & sPI & "),"""")"
        ws.Cells(r, c + 4).Formula = "=IFERROR(" & fc & "*MINIFS(" & nm & ",dbEmisor," & cel & "," & fech & sPI & "),"""")"
        ws.Cells(r, c + 5).Formula = "=IFERROR(" & fc & "*MAXIFS(" & nm & ",dbEmisor," & cel & "," & fech & sPI & "),"""")"
        ws.Cells(r, c + 6).Formula = "=IFERROR(" & ColL(c + 2) & r & "-" & fc & "*AVERAGEIFS(" & nm & "," & fech & "," & sec & sPI & "),"""")"
        ws.Range(ws.Cells(r, c + 1), ws.Cells(r, c + 6)).NumberFormat = fmt
        ws.Cells(r, c + 1).NumberFormat = "0"
        ws.Cells(r, c + 6).NumberFormat = IIf(m = 0, IIf(SPREAD_EN_PB, "+0.0;-0.0", "+0.000;-0.000"), "+0.00;-0.00")
        ws.Cells(r, c + 6).Font.Bold = True
    Next i
    ws.Cells(r - 1, c + 6).Value = ""
    ws.Cells(r0 + 2, c + 6).Value = "vs merc."

    ws.Range(ws.Cells(r0 + 1, c), ws.Cells(r, c + 6)).Font.Size = 8
    ws.Columns(c).ColumnWidth = 22
    For i = 1 To 6
        ws.Columns(c + i).ColumnWidth = 9
    Next i
End Sub

Private Function MaxCol(db As Worksheet, hdr As Long, lastR As Long, c As Long) As Double
    Dim r As Long, tope As Long, v As Variant
    tope = hdr + 300
    If lastR < tope Then tope = lastR
    For r = hdr + 1 To tope
        v = db.Cells(r, c).Value
        If IsNumeric(v) And v <> "" Then
            If CDbl(v) > MaxCol Then MaxCol = CDbl(v)
        End If
    Next r
End Function

Private Function FactorTasa(db As Worksheet, hdr As Long, lastR As Long, c As Long) As Double
    ' deja la tasa expresada en % (4.55)
    FactorTasa = IIf(MaxCol(db, hdr, lastR, c) > 1, 1, 100)
End Function

Private Function FactorSpread(db As Worksheet, hdr As Long, lastR As Long, c As Long) As Double
    ' detecta la escala por los valores, no por el encabezado
    Dim mx As Double, esc As Double
    mx = MaxCol(db, hdr, lastR, c)
    If mx <= 0.5 Then
        esc = 100        ' viene como fraccion decimal (0.0045) -> 0.45 %
    ElseIf mx <= 20 Then
        esc = 1          ' ya viene en porcentaje (0.45)
    Else
        esc = 0.01       ' ya viene en puntos basicos (45) -> 0.45 %
    End If
    FactorSpread = esc * IIf(SPREAD_EN_PB, 100, 1)
End Function

Private Sub Puntos(ByVal ws As Worksheet, ByVal cX As Long, ByVal celEmi As String, _
                   ByVal sPF As String, ByVal nm As String, ByVal fc As String)
    Dim x As String, y As String, cond As String
    x = ColL(cX): y = ColL(cX + 1)
    cond = "(dbEmisor=" & celEmi & ")*(dbFecha>=$B$4)*(dbFecha<=$C$4)" & sPF
    ws.Range(x & "3").Formula2 = "=IFERROR(INDEX(FILTER(dbFecha," & cond & "),ROWS($" & x & "$3:$" & x & "3)),NA())"
    ws.Range(y & "3").Formula2 = "=IFERROR(" & fc & "*INDEX(FILTER(" & nm & "," & cond & "),ROWS($" & y & "$3:$" & y & "3)),NA())"
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
        s = Replace(s, ChrW(241), "n")
        s = Replace(s, Chr(10), " ")
        s = Replace(s, Chr(13), " ")
        s = Replace(s, Chr(160), " ")
        Do While InStr(s, "  ") > 0
            s = Replace(s, "  ", " ")
        Loop
        If Len(s) > 0 And InStr(1, s, txt, vbTextCompare) > 0 Then
            BuscarCol = c
            Exit Function
        End If
    Next c
End Function

Private Function ColL(i As Long) As String
    If i < 1 Then
        ColL = "n/d"
        Exit Function
    End If
    ColL = Split(Cells(1, i).Address(True, False), "$")(0)
End Function

' ============================================================
' DIAGNOSTICO
' ============================================================

Sub DiagnosticoDashCP()
    Dim wb As Workbook, db As Worksheet
    Dim hdr As Long, lastR As Long, r As Long, i As Long
    Dim cFec As Long, cEmi As Long, cSec As Long, cRat As Long, cSpr As Long, cPlz As Long
    Dim msg As String, v As Variant
    Dim fMin As Double, fMax As Double
    Dim nFec As Long, nPlz As Long, nSpr As Long, nVent As Long
    Dim d1 As Double, d2 As Double
    Dim dic As Object

    Set wb = ThisWorkbook
    Set db = HojaDB(wb)
    If db Is Nothing Then MsgBox "No encontre la hoja DB-Historica CP.", vbExclamation: Exit Sub
    hdr = FilaEncabezado(db)
    If hdr = 0 Then MsgBox "No encontre la fila de encabezados.", vbExclamation: Exit Sub

    cFec = BuscarCol(db, hdr, "emision")
    cEmi = BuscarCol(db, hdr, "emisor")
    cSec = BuscarCol(db, hdr, "sector")
    cRat = BuscarCol(db, hdr, "rating")
    cPlz = BuscarCol(db, hdr, "anos")
    If cPlz = 0 Then cPlz = BuscarCol(db, hdr, "os vt")
    If cPlz = 0 Then cPlz = BuscarCol(db, hdr, "plazo")
    If cPlz = 0 Then cPlz = BuscarCol(db, hdr, "duracion")
    cSpr = BuscarCol(db, hdr, "spread (pb)")
    If cSpr = 0 Then cSpr = BuscarCol(db, hdr, "spread pb")
    If cSpr = 0 Then cSpr = BuscarCol(db, hdr, "spread")

    lastR = db.Cells(db.Rows.Count, cEmi).End(xlUp).Row
    msg = "HOJA: " & db.Name & vbCrLf & _
          "Encabezados en fila " & hdr & ", datos " & hdr + 1 & " a " & lastR & vbCrLf & vbCrLf
    msg = msg & "COLUMNAS DETECTADAS" & vbCrLf
    msg = msg & "  Emision = " & DHead(db, hdr, cFec) & vbCrLf
    msg = msg & "  Emisor  = " & DHead(db, hdr, cEmi) & vbCrLf
    msg = msg & "  Sector  = " & DHead(db, hdr, cSec) & vbCrLf
    msg = msg & "  Rating  = " & DHead(db, hdr, cRat) & vbCrLf
    msg = msg & "  Plazo   = " & DHead(db, hdr, cPlz) & vbCrLf
    msg = msg & "  Spread  = " & DHead(db, hdr, cSpr) & vbCrLf & vbCrLf
    msg = msg & "TODOS LOS ENCABEZADOS DE LA FILA " & hdr & vbCrLf
    For i = 1 To 40
        If Len(Trim$(CStr(db.Cells(hdr, i).Value))) > 0 Then
            msg = msg & "  " & ColL(i) & " = [" & db.Cells(hdr, i).Value & "]" & vbCrLf
        End If
    Next i
    msg = msg & vbCrLf

    d1 = CDbl(DateSerial(Year(Date) - 1, Month(Date), Day(Date)))
    d2 = CDbl(Date)
    fMin = 9999999: fMax = 0
    Set dic = CreateObject("Scripting.Dictionary")
    For r = hdr + 1 To lastR
        v = db.Cells(r, cFec).Value
        If IsDate(v) Then
            nFec = nFec + 1
            If CDbl(v) < fMin Then fMin = CDbl(v)
            If CDbl(v) > fMax Then fMax = CDbl(v)
            If CDbl(v) >= d1 And CDbl(v) <= d2 Then nVent = nVent + 1
        End If
        If cPlz > 0 Then
            If IsNumeric(db.Cells(r, cPlz).Value) And db.Cells(r, cPlz).Value <> "" Then nPlz = nPlz + 1
        End If
        If IsNumeric(db.Cells(r, cSpr).Value) And db.Cells(r, cSpr).Value <> "" Then nSpr = nSpr + 1
        v = CStr(db.Cells(r, cRat).Value)
        If Len(v) > 0 Then
            If Not dic.Exists(v) Then dic.Add v, 0
            dic(v) = dic(v) + 1
        End If
    Next r

    msg = msg & "CALIDAD DE DATOS (" & lastR - hdr & " filas)" & vbCrLf
    msg = msg & "  Fechas validas    : " & nFec & vbCrLf
    If nFec > 0 Then msg = msg & "  Rango de emision  : " & Format(fMin, "dd/mm/yyyy") & "  a  " & Format(fMax, "dd/mm/yyyy") & vbCrLf
    msg = msg & "  En ultimos 12 mes : " & nVent & "   <-- si es 0, el grafico sale vacio" & vbCrLf
    If cPlz > 0 Then
        msg = msg & "  Plazo numerico    : " & nPlz & " de " & lastR - hdr & vbCrLf
    Else
        msg = msg & "  Plazo             : columna NO detectada" & vbCrLf
    End If
    msg = msg & "  Spread numerico   : " & nSpr & " de " & lastR - hdr & vbCrLf
    If nSpr > 0 Then msg = msg & "  Ejemplo de spread : " & db.Cells(hdr + 1, cSpr).Value & "   (si es 0.0045 el factor debe ser 10000)" & vbCrLf
    msg = msg & vbCrLf & "RATINGS ENCONTRADOS (entre corchetes, para ver espacios)" & vbCrLf
    For i = 0 To dic.Count - 1
        msg = msg & "  [" & dic.Keys()(i) & "] = " & dic.Items()(i) & vbCrLf
        If i >= 14 Then msg = msg & "  ...": Exit For
    Next i

    MsgBox msg, vbInformation, "Diagnostico DB-Historica CP"
End Sub

Private Function DHead(db As Worksheet, hdr As Long, c As Long) As String
    If c < 1 Then
        DHead = "NO ENCONTRADA"
        Exit Function
    End If
    DHead = ColL(c) & "   [" & db.Cells(hdr, c).Value & "]"
End Function
