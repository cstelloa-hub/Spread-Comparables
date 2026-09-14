Attribute VB_Name = "Módulo1"
Option Explicit

' ============================================================
' mod_DashSpreadsCP
' Dashboard de colocaciones de corto plazo
' ============================================================

Private Const SH_DASH As String = "Dashboard CP"
Private Const N_MESES As Long = 180
Private Const N_PTOS As Long = 150

Private Const SPREAD_EN_PB As Boolean = True

' En tu base actual:
' Q = Tasa
' R = Tasa base
Private Const COL_TASA_FINAL As String = "Q"
Private Const COL_TASA_BASE As String = "R"


' ============================================================
' CREAR DASHBOARD
' ============================================================

Sub CrearDashboardSpreadsCP()

    Dim wb As Workbook, db As Worksheet, ws As Worksheet
    Dim hdr As Long, lastR As Long, i As Long, m As Long, b As Long
    Dim cFec As Long, cEmi As Long, cSec As Long, cRat As Long
    Dim cSpr As Long, cPlz As Long, cTF As Long, cTB As Long, cDiv As Long
    Dim cMon As Long   ' NUEVO: columna del monto colocado
    Dim usaPb As Boolean, usaPlz As Boolean
    Dim factor As Double, facTF As Double, facTB As Double
    Dim rats As Variant, cols As Variant, mets As Variant
    Dim lnk As String, sPI As String, sPF As String
    Dim x As String, y As String, cn As String, pr As String
    Dim nm As String, fc As String
    Dim et As Variant
    Dim ch As Chart, sr As Series
    Dim nSer As Long, r0 As Long

    Set wb = ThisWorkbook
    Set db = HojaDB(wb)

    If db Is Nothing Then
        MsgBox "No encontré una hoja cuyo nombre contenga 'DB-Hist'.", vbExclamation
        Exit Sub
    End If

    hdr = FilaEncabezado(db)

    If hdr = 0 Then
        MsgBox "No encontré la fila de encabezados.", vbExclamation
        Exit Sub
    End If

    cFec = BuscarCol(db, hdr, "emision")
    cEmi = BuscarCol(db, hdr, "emisor")
    cSec = BuscarCol(db, hdr, "sector")
    cRat = BuscarCol(db, hdr, "rating")
    cDiv = BuscarCol(db, hdr, "divisa")

    If cDiv = 0 Then
        cDiv = BuscarCol(db, hdr, "moneda")
    End If

    cPlz = BuscarCol(db, hdr, "anos")

    If cPlz = 0 Then
        cPlz = BuscarCol(db, hdr, "plazo")
    End If

    If cPlz = 0 Then
        cPlz = BuscarCol(db, hdr, "duracion")
    End If

    ' ========================================================
    ' CORRECCIÓN PRINCIPAL:
    ' La base contiene:
    ' S = Spread %
    ' T = Spread (pbs)
    ' Se debe utilizar T.
    ' ========================================================

    cSpr = BuscarCol(db, hdr, "spread (pbs)")

    If cSpr = 0 Then
        cSpr = BuscarCol(db, hdr, "spread (pb)")
    End If

    If cSpr = 0 Then
        cSpr = BuscarCol(db, hdr, "spread %")
    End If

    ' NUEVO: columna del monto colocado, para el promedio ponderado
    cMon = BuscarCol(db, hdr, "asignado")
    If cMon = 0 Then cMon = BuscarCol(db, hdr, "monto")

    usaPb = False

    If cSpr > 0 Then
        If InStr(1, LCase$(CStr(db.Cells(hdr, cSpr).Value)), "pb", vbTextCompare) > 0 Then
            usaPb = True
        End If
    End If

    If cFec = 0 Or cEmi = 0 Or cSec = 0 Or cRat = 0 Or cSpr = 0 Then

        MsgBox "Faltan columnas necesarias." & vbCrLf & _
               "Emisión = " & cFec & vbCrLf & _
               "Emisor = " & cEmi & vbCrLf & _
               "Sector = " & cSec & vbCrLf & _
               "Rating = " & cRat & vbCrLf & _
               "Spread = " & cSpr, vbExclamation

        Exit Sub
    End If

    cTF = db.Range(COL_TASA_FINAL & "1").Column
    cTB = db.Range(COL_TASA_BASE & "1").Column

    lastR = db.Cells(db.Rows.Count, cEmi).End(xlUp).Row

    facTF = FactorTasa(db, hdr, lastR, cTF)
    facTB = FactorTasa(db, hdr, lastR, cTB)
    factor = FactorSpread(db, hdr, lastR, cSpr)

    ' ========================================================
    ' NOMBRES DEFINIDOS DE LA BASE
    ' ========================================================

    DefNombre wb, "dbFecha", db.Name, ColL(cFec), hdr + 1, lastR
    DefNombre wb, "dbEmisor", db.Name, ColL(cEmi), hdr + 1, lastR
    DefNombre wb, "dbSector", db.Name, ColL(cSec), hdr + 1, lastR
    DefNombre wb, "dbRating", db.Name, ColL(cRat), hdr + 1, lastR
    DefNombre wb, "dbSpread", db.Name, ColL(cSpr), hdr + 1, lastR
    DefNombre wb, "dbTFinal", db.Name, ColL(cTF), hdr + 1, lastR
    DefNombre wb, "dbTBase", db.Name, ColL(cTB), hdr + 1, lastR

    ' NUEVO: rango con nombre para el monto colocado
    If cMon > 0 Then
        DefNombre wb, "dbMonto", db.Name, ColL(cMon), hdr + 1, lastR
    Else
        BorrarNombre wb, "dbMonto"
    End If

    If cPlz > 0 Then
        DefNombre wb, "dbPlazo", db.Name, ColL(cPlz), hdr + 1, lastR
        usaPlz = True
    Else
        usaPlz = False
    End If

    If cDiv > 0 Then
        DefNombre wb, "dbDivisa", db.Name, ColL(cDiv), hdr + 1, lastR
    End If

    If usaPlz Then
        sPI = ",dbPlazo,"">=""&$G$4,dbPlazo,""<=""&$H$4"
        sPF = "*(dbPlazo>=$G$4)*(dbPlazo<=$H$4)"
    End If

    If cDiv > 0 Then
        sPI = sPI & ",dbDivisa,IF($I$4=""(todas)"",""<>"",$I$4)"
        sPF = sPF & "*IF($I$4=""(todas)"",1,dbDivisa=$I$4)"
    End If

    ' ========================================================
    ' ELIMINAR DASHBOARD ANTERIOR
    ' ========================================================

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

    ' ========================================================
    ' TÍTULO
    ' ========================================================

    With ws.Range("A1:J1")
        .Merge
        .Value = "  Colocaciones de corto plazo - spread, tasa final y tasa base"
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 11
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A2").Value = "  Promedio mensual por rating - PEN - DB Histórica CP"
    ws.Range("A2").Font.Color = RGB(110, 110, 110)

    ' ========================================================
    ' FILTROS
    ' ========================================================

    et = Array("DESDE", "HASTA", "SECTOR", "EMISOR 1", _
               "EMISOR 2", "PLAZO MIN", "PLAZO MAX", "MONEDA")

    For i = 0 To 7

        With ws.Cells(3, 2 + i)
            .Value = et(i)
            .Font.Size = 7
            .Font.Color = RGB(140, 140, 140)
        End With

        ws.Columns(2 + i).ColumnWidth = IIf(i >= 3 And i <= 4, 24, 13)

    Next i

    ' Se mantienen los últimos 12 meses como filtro inicial
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

    ' ========================================================
    ' LISTAS AUXILIARES
    ' ========================================================

    ws.Cells(2, 110).Value = "(todos)"
    ws.Cells(3, 110).Formula2 = "=SORT(UNIQUE(FILTER(dbSector,dbSector<>"""")))"

    ws.Cells(2, 112).Value = "(ninguno)"
    ws.Cells(3, 112).Formula2 = _
        "=SORT(UNIQUE(FILTER(dbEmisor,(dbEmisor<>"""")*" & _
        "IF($D$4=""(todos)"",1,dbSector=$D$4))))"

    ws.Cells(2, 120).Value = "(todas)"
    ws.Cells(3, 120).Formula2 = "=SORT(UNIQUE(FILTER(dbDivisa,dbDivisa<>"""")))"

    ' Eliminar nombres anteriores para evitar errores al ejecutar nuevamente
    BorrarNombre wb, "lstSectores"
    BorrarNombre wb, "lstEmisores"
    BorrarNombre wb, "lstMonedas"

    wb.Names.Add Name:="lstSectores", _
        RefersTo:="='" & SH_DASH & "'!$" & ColL(110) & "$2:INDEX('" & _
        SH_DASH & "'!$" & ColL(110) & "$2:$" & ColL(110) & "$300,COUNTA('" & _
        SH_DASH & "'!$" & ColL(110) & "$2:$" & ColL(110) & "$300))"

    wb.Names.Add Name:="lstEmisores", _
        RefersTo:="='" & SH_DASH & "'!$" & ColL(112) & "$2:INDEX('" & _
        SH_DASH & "'!$" & ColL(112) & "$2:$" & ColL(112) & "$900,COUNTA('" & _
        SH_DASH & "'!$" & ColL(112) & "$2:$" & ColL(112) & "$900))"

    wb.Names.Add Name:="lstMonedas", _
        RefersTo:="='" & SH_DASH & "'!$" & ColL(120) & "$2:INDEX('" & _
        SH_DASH & "'!$" & ColL(120) & "$2:$" & ColL(120) & "$50,COUNTA('" & _
        SH_DASH & "'!$" & ColL(120) & "$2:$" & ColL(120) & "$50))"

    If cDiv > 0 Then
        PonerLista ws.Range("I4"), "=lstMonedas"
    End If

    PonerLista ws.Range("D4"), "=lstSectores"
    PonerLista ws.Range("E4"), "=lstEmisores"
    PonerLista ws.Range("F4"), "=lstEmisores"

    ' ========================================================
    ' RATINGS
    ' ========================================================

    rats = Array("CP-1+", "CP-1", "CP-1-", "CP-2+", "CP-2")

    cols = Array( _
        RGB(31, 59, 87), _
        RGB(46, 127, 140), _
        RGB(176, 135, 59), _
        RGB(122, 92, 142), _
        RGB(140, 140, 140))

    mets = Array("dbSpread", "dbTFinal", "dbTBase")

    For i = 0 To 4

        lnk = "$" & ColL(114 + i) & "$1"

        With ws.CheckBoxes.Add( _
            ws.Cells(6, 2).Left + i * 78, _
            ws.Cells(6, 2).Top - 1, _
            74, 18)

            .Caption = rats(i)
            .LinkedCell = "'" & SH_DASH & "'!" & lnk
            .Value = IIf(i <= 2, xlOn, xlOff)
            .Name = "chk" & i

        End With

    Next i

    ' ========================================================
    ' FACTORES DE ESCALA
    ' ========================================================

    ws.Range("AF1").Value = factor
    ws.Range("AF2").Value = facTF
    ws.Range("AF3").Value = facTB

    ws.Range("AE1").Formula = _
        "=""Spread de colocación (" & IIf(SPREAD_EN_PB, "pb", "%") & _
        ") - ""&TEXT($B$4,""mmm-aa"")&"" a ""&TEXT($C$4,""mmm-aa"")&" & _
        "IF($D$4=""(todos)"","""","" - ""&$D$4)"

    ws.Range("AE2").Formula = _
        "=""Tasa final - ""&TEXT($B$4,""mmm-aa"")&"" a ""&TEXT($C$4,""mmm-aa"")&" & _
        "IF($D$4=""(todos)"","""","" - ""&$D$4)"

    ws.Range("AE3").Formula = _
        "=""Tasa base (curva de mercado, no depende del rating) - ""&" & _
        "TEXT($B$4,""mmm-aa"")&"" a ""&TEXT($C$4,""mmm-aa"")&" & _
        "IF($D$4=""(todos)"","""","" - ""&$D$4)"

    ' ========================================================
    ' FECHAS MENSUALES
    ' ========================================================

    ws.Range("AH3").Formula = "=DATE(YEAR($B$4),MONTH($B$4),1)"
    ws.Range("AH4").Formula = _
        "=IF(AH3=0,0,IF(EDATE(AH3,1)>$C$4,0,EDATE(AH3,1)))"

    ws.Range("AH4").AutoFill Destination:= _
        ws.Range("AH4:AH" & (2 + N_MESES))

    ' ========================================================
    ' CÁLCULOS
    ' ========================================================

    For m = 0 To 2

        b = 36 + m * 24
        nm = CStr(mets(m))
        fc = "$AF$" & (m + 1)

        If m < 2 Then

            For i = 0 To 4

                ws.Cells(2, b + i).Value = rats(i)
                ws.Cells(1, b + i).Formula = _
                    "=IF($" & ColL(114 + i) & "$1,1,0)"

            Next i

            ws.Cells(3, b).Formula = _
                "=IF($AH3=0,0,IF(" & ColL(b) & "$1=0,0," & _
                "IFERROR(" & fc & "*AVERAGEIFS(" & nm & _
                ",dbFecha,"">=""&$AH3,dbFecha,""<""&EDATE($AH3,1)," & _
                "dbRating," & ColL(b) & "$2," & _
                "dbSector,IF($D$4=""(todos)"",""<>"",$D$4)" & _
                sPI & "),0)))"

            ws.Cells(3, b).Copy _
                ws.Range(ws.Cells(3, b), ws.Cells(2 + N_MESES, b + 4))

            ws.Cells(3, b + 5).Formula = _
                "=IF($AH3=0,0,IF(" & ColL(b) & "$1=0,0," & _
                "COUNTIFS(dbRating," & ColL(b) & "$2," & _
                "dbFecha,"">=""&$AH3,dbFecha,""<""&EDATE($AH3,1)," & _
                "dbSector,IF($D$4=""(todos)"",""<>"",$D$4)" & _
                sPI & ")))"

            ws.Cells(3, b + 5).Copy _
                ws.Range(ws.Cells(3, b + 5), ws.Cells(2 + N_MESES, b + 9))

            For i = 0 To 4

                x = ColL(b + 10 + i * 2)
                y = ColL(b + 11 + i * 2)
                cn = ColL(b + 5 + i)
                pr = ColL(b + i)

                ws.Range(x & "3").Formula2 = _
                    "=IFERROR(INDEX(FILTER($AH$3:$AH$" & _
                    (2 + N_MESES) & "," & _
                    "$" & cn & "$3:$" & cn & "$" & _
                    (2 + N_MESES) & ">0),ROWS($" & _
                    x & "$3:" & x & "3)),NA())"

                ws.Range(y & "3").Formula2 = _
                    "=IFERROR(INDEX(FILTER($" & pr & "$3:$" & _
                    pr & "$" & (2 + N_MESES) & "," & _
                    "$" & cn & "$3:$" & cn & "$" & _
                    (2 + N_MESES) & ">0),ROWS($" & _
                    y & "$3:" & y & "3)),NA())"

                ws.Range(x & "3:" & y & "3").Copy _
                    ws.Range(x & "3:" & y & (2 + N_MESES))

            Next i

        Else

            ws.Cells(2, b).Value = "Mercado"

            ws.Cells(3, b).Formula = _
                "=IF($AH3=0,0,IFERROR(" & fc & _
                "*AVERAGEIFS(" & nm & _
                ",dbFecha,"">=""&$AH3,dbFecha,""<""&EDATE($AH3,1)," & _
                "dbSector,IF($D$4=""(todos)"",""<>"",$D$4)" & _
                sPI & "),0))"

            ws.Cells(3, b).Copy _
                ws.Range(ws.Cells(3, b), ws.Cells(2 + N_MESES, b))

            ws.Cells(3, b + 5).Formula = _
                "=IF($AH3=0,0,COUNTIFS(" & _
                "dbFecha,"">=""&$AH3,dbFecha,""<""&EDATE($AH3,1)," & _
                "dbSector,IF($D$4=""(todos)"",""<>"",$D$4)" & _
                sPI & "))"

            ws.Cells(3, b + 5).Copy _
                ws.Range(ws.Cells(3, b + 5), ws.Cells(2 + N_MESES, b + 5))

            x = ColL(b + 10)
            y = ColL(b + 11)
            cn = ColL(b + 5)
            pr = ColL(b)

            ws.Range(x & "3").Formula2 = _
                "=IFERROR(INDEX(FILTER($AH$3:$AH$" & _
                (2 + N_MESES) & "," & _
                "$" & cn & "$3:$" & cn & "$" & _
                (2 + N_MESES) & ">0),ROWS($" & _
                x & "$3:" & x & "3)),NA())"

            ws.Range(y & "3").Formula2 = _
                "=IFERROR(INDEX(FILTER($" & pr & "$3:$" & _
                pr & "$" & (2 + N_MESES) & "," & _
                "$" & cn & "$3:$" & cn & "$" & _
                (2 + N_MESES) & ">0),ROWS($" & _
                y & "$3:" & y & "3)),NA())"

            ws.Range(x & "3:" & y & "3").Copy _
                ws.Range(x & "3:" & y & (2 + N_MESES))

        End If

        Puntos ws, b + 20, "$E$4", sPF, nm, fc
        Puntos ws, b + 22, "$F$4", sPF, nm, fc

    Next m

    ws.Range(ws.Columns(31), ws.Columns(124)).EntireColumn.Hidden = True

    Application.Calculate

    ' ========================================================
    ' GRÁFICOS
    ' ========================================================

    For m = 0 To 2

        b = 36 + m * 24
        r0 = 8 + m * 34

        Set ch = ws.Shapes.AddChart2( _
            -1, _
            xlXYScatterLinesNoMarkers, _
            ws.Cells(r0, 1).Left + 5, _
            ws.Cells(r0, 1).Top, _
            ws.Cells(r0, 12).Left - ws.Cells(r0, 1).Left - 12, _
            ws.Cells(r0 + 32, 1).Top - ws.Cells(r0, 1).Top).Chart

        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
        Loop

        nSer = IIf(m < 2, 5, 1)

        For i = 0 To nSer - 1

            Set sr = ch.SeriesCollection.NewSeries

            sr.Name = "='" & SH_DASH & "'!" & _
                      ws.Cells(2, b + i).Address

            sr.XValues = "='" & SH_DASH & "'!" & _
                         ws.Range(ws.Cells(3, b + 10 + i * 2), _
                         ws.Cells(2 + N_MESES, b + 10 + i * 2)).Address

            sr.Values = "='" & SH_DASH & "'!" & _
                        ws.Range(ws.Cells(3, b + 11 + i * 2), _
                        ws.Cells(2 + N_MESES, b + 11 + i * 2)).Address

            sr.ChartType = xlXYScatterLinesNoMarkers
            sr.Format.Line.ForeColor.RGB = _
                IIf(m < 2, cols(i), RGB(31, 59, 87))
            sr.Format.Line.Weight = 2.25
            sr.Smooth = False

        Next i

        Set sr = ch.SeriesCollection.NewSeries

        sr.Name = "='" & SH_DASH & "'!$E$4"

        sr.XValues = "='" & SH_DASH & "'!" & _
                     ws.Range(ws.Cells(3, b + 20), _
                     ws.Cells(2 + N_PTOS, b + 20)).Address

        sr.Values = "='" & SH_DASH & "'!" & _
                    ws.Range(ws.Cells(3, b + 21), _
                    ws.Cells(2 + N_PTOS, b + 21)).Address

        sr.ChartType = xlXYScatter
        sr.MarkerStyle = xlMarkerStyleCircle
        sr.MarkerSize = 8
        sr.MarkerBackgroundColor = RGB(212, 12, 12)
        sr.MarkerForegroundColor = RGB(255, 255, 255)

        Set sr = ch.SeriesCollection.NewSeries

        sr.Name = "='" & SH_DASH & "'!$F$4"

        sr.XValues = "='" & SH_DASH & "'!" & _
                     ws.Range(ws.Cells(3, b + 22), _
                     ws.Cells(2 + N_PTOS, b + 22)).Address

        sr.Values = "='" & SH_DASH & "'!" & _
                    ws.Range(ws.Cells(3, b + 23), _
                    ws.Cells(2 + N_PTOS, b + 23)).Address

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

            .AxisTitle.Text = IIf( _
                m = 0, _
                IIf(SPREAD_EN_PB, "Spread (pb)", "Spread (%)"), _
                "Tasa (%)")

            .AxisTitle.Font.Size = 8
            .AxisTitle.Font.Color = RGB(110, 110, 110)
            .TickLabels.Font.Size = 8
            .TickLabels.Font.Color = RGB(90, 90, 90)

            .TickLabels.NumberFormat = IIf( _
                m = 0, _
                IIf(SPREAD_EN_PB, "0", "0.00"), _
                "0.00")

            .MajorTickMark = xlNone
            .Format.Line.Visible = msoFalse
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(234, 234, 234)
            .MajorGridlines.Format.Line.Weight = 0.75

        End With

        Resumen ws, r0, m, CStr(mets(m)), _
                "$AF$" & (m + 1), sPI, sPF, rats

    Next m

    ws.Range("A112").Value = _
        "Fuente: DB Histórica CP. Las fechas, sector y plazo aplican a los tres gráficos. " & _
        "El filtro de rating aplica al spread y a la tasa final. " & _
        "La tasa base se presenta como una única línea de mercado."

    ws.Range("A112").Font.Size = 7
    ws.Range("A112").Font.Color = RGB(140, 140, 140)

    ws.Activate
    ws.Range("B4").Select

    MsgBox "Dashboard creado correctamente." & vbCrLf & vbCrLf & _
           "Hoja: " & db.Name & vbCrLf & _
           "Encabezados: fila " & hdr & vbCrLf & _
           "Datos: " & hdr + 1 & " a " & lastR & vbCrLf & vbCrLf & _
           "Emisión: " & ColL(cFec) & vbCrLf & _
           "Emisor: " & ColL(cEmi) & vbCrLf & _
           "Sector: " & ColL(cSec) & vbCrLf & _
           "Rating: " & ColL(cRat) & vbCrLf & _
           "Spread utilizado: " & ColL(cSpr) & _
           " [" & db.Cells(hdr, cSpr).Value & "]" & vbCrLf & _
           "Factor del spread: " & factor & vbCrLf & _
           "Tasa final: " & ColL(cTF) & _
           " [" & db.Cells(hdr, cTF).Value & "]" & vbCrLf & _
           "Tasa base: " & ColL(cTB) & _
           " [" & db.Cells(hdr, cTB).Value & "]", vbInformation

End Sub


' ============================================================
' TABLA DE RESUMEN
' ============================================================

Private Sub Resumen( _
    ByVal ws As Worksheet, _
    ByVal r0 As Long, _
    ByVal m As Long, _
    ByVal nm As String, _
    ByVal fc As String, _
    ByVal sPI As String, _
    ByVal sPF As String, _
    ByVal rats As Variant)

    Dim c As Long, r As Long, i As Long, nFil As Long
    Dim sec As String, fech As String, cond As String
    Dim flag As String, cel As String
    Dim fmt As String, tit As String

    c = 12

    sec = "dbSector,IF($D$4=""(todos)"",""<>"",$D$4)"
    fech = "dbFecha,"">=""&$B$4,dbFecha,""<=""&$C$4"

    fmt = IIf( _
        m = 0, _
        IIf(SPREAD_EN_PB, "0.0", "0.000"), _
        "0.00")

    tit = IIf( _
        m = 0, _
        IIf(SPREAD_EN_PB, _
        "RESUMEN - SPREAD (pb)", _
        "RESUMEN - SPREAD (%)"), _
        IIf(m = 1, _
        "RESUMEN - TASA FINAL (%)", _
        "RESUMEN - TASA BASE (%)"))

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
    ws.Cells(r, c + 2).Value = "Último"
    ws.Cells(r, c + 3).Value = "Prom"
    ws.Cells(r, c + 4).Value = "Min"
    ws.Cells(r, c + 5).Value = "Max"
    ws.Cells(r, c + 6).Value = "vs merc."
    ws.Cells(r, c + 7).Value = "Prom pond."   ' NUEVO

    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + 7))
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

            cond = _
                "(dbRating=" & cel & ")*" & _
                "(dbFecha>=$B$4)*" & _
                "(dbFecha<=$C$4)*" & _
                "IF($D$4=""(todos)"",1,dbSector=$D$4)" & sPF

            ws.Cells(r, c + 1).Formula = _
                "=IF(" & flag & "=FALSE,"""",COUNTIFS(" & _
                "dbRating," & cel & "," & fech & "," & sec & sPI & "))"

            ws.Cells(r, c + 2).Formula = _
                "=IF(" & flag & "=FALSE,"""",IFERROR(" & fc & _
                "*XLOOKUP(MAX(FILTER(dbFecha," & cond & "))," & _
                "FILTER(dbFecha," & cond & ")," & _
                "FILTER(" & nm & "," & cond & ")),""""))"

            ws.Cells(r, c + 3).Formula = _
                "=IF(" & flag & "=FALSE,"""",IFERROR(" & fc & _
                "*AVERAGEIFS(" & nm & ",dbRating," & cel & _
                "," & fech & "," & sec & sPI & "),""""))"

            ws.Cells(r, c + 4).Formula = _
                "=IF(" & flag & "=FALSE,"""",IFERROR(" & fc & _
                "*MINIFS(" & nm & ",dbRating," & cel & _
                "," & fech & "," & sec & sPI & "),""""))"

            ws.Cells(r, c + 5).Formula = _
                "=IF(" & flag & "=FALSE,"""",IFERROR(" & fc & _
                "*MAXIFS(" & nm & ",dbRating," & cel & _
                "," & fech & "," & sec & sPI & "),""""))"

            ' NUEVO: promedio ponderado por monto colocado
            If TieneNombre(ws.Parent, "dbMonto") Then
                ws.Cells(r, c + 7).Formula = _
                    "=IF(" & flag & "=FALSE,"""",IFERROR(" & fc & _
                    "*SUMPRODUCT((" & cond & ")*dbMonto*" & nm & ")/" & _
                    "SUMPRODUCT((" & cond & ")*dbMonto),""""))"
                ws.Cells(r, c + 7).NumberFormat = fmt
            End If

        Else

            ws.Cells(r, c).Value = "Mercado"

            cond = _
                "(dbFecha>=$B$4)*" & _
                "(dbFecha<=$C$4)*" & _
                "IF($D$4=""(todos)"",1,dbSector=$D$4)" & sPF

            ws.Cells(r, c + 1).Formula = _
                "=COUNTIFS(" & fech & "," & sec & sPI & ")"

            ws.Cells(r, c + 2).Formula = _
                "=IFERROR(" & fc & _
                "*XLOOKUP(MAX(FILTER(dbFecha," & cond & "))," & _
                "FILTER(dbFecha," & cond & ")," & _
                "FILTER(" & nm & "," & cond & ")),"""")"

            ws.Cells(r, c + 3).Formula = _
                "=IFERROR(" & fc & _
                "*AVERAGEIFS(" & nm & "," & fech & _
                "," & sec & sPI & "),"""")"

            ws.Cells(r, c + 4).Formula = _
                "=IFERROR(" & fc & _
                "*MINIFS(" & nm & "," & fech & _
                "," & sec & sPI & "),"""")"

            ws.Cells(r, c + 5).Formula = _
                "=IFERROR(" & fc & _
                "*MAXIFS(" & nm & "," & fech & _
                "," & sec & sPI & "),"""")"

        End If

        ws.Range(ws.Cells(r, c + 1), ws.Cells(r, c + 5)).NumberFormat = fmt
        ws.Cells(r, c + 1).NumberFormat = "0"

    Next i

    ' ========================================================
    ' EMISORES SELECCIONADOS
    ' ========================================================

    r = r + 1

    For i = 0 To 1

        r = r + 1

        cel = IIf(i = 0, "$E$4", "$F$4")

        cond = _
            "(dbEmisor=" & cel & ")*" & _
            "(dbFecha>=$B$4)*" & _
            "(dbFecha<=$C$4)" & sPF

        ws.Cells(r, c).Formula = "=" & cel
        ws.Cells(r, c).Font.Bold = True

        ws.Cells(r, c + 1).Formula = _
            "=IFERROR(COUNTIFS(dbEmisor," & cel & _
            "," & fech & sPI & "),"""")"

        ws.Cells(r, c + 2).Formula = _
            "=IFERROR(" & fc & _
            "*XLOOKUP(MAX(FILTER(dbFecha," & cond & "))," & _
            "FILTER(dbFecha," & cond & ")," & _
            "FILTER(" & nm & "," & cond & ")),"""")"

        ws.Cells(r, c + 3).Formula = _
            "=IFERROR(" & fc & _
            "*AVERAGEIFS(" & nm & ",dbEmisor," & cel & _
            "," & fech & sPI & "),"""")"

        ws.Cells(r, c + 4).Formula = _
            "=IFERROR(" & fc & _
            "*MINIFS(" & nm & ",dbEmisor," & cel & _
            "," & fech & sPI & "),"""")"

        ws.Cells(r, c + 5).Formula = _
            "=IFERROR(" & fc & _
            "*MAXIFS(" & nm & ",dbEmisor," & cel & _
            "," & fech & sPI & "),"""")"

        ws.Cells(r, c + 6).Formula = _
            "=IFERROR(" & ColL(c + 2) & r & _
            "-" & fc & "*AVERAGEIFS(" & nm & _
            "," & fech & "," & sec & sPI & "),"""")"

        ws.Range(ws.Cells(r, c + 1), ws.Cells(r, c + 6)).NumberFormat = fmt
        ws.Cells(r, c + 1).NumberFormat = "0"

        ws.Cells(r, c + 6).NumberFormat = _
            IIf(m = 0, _
            IIf(SPREAD_EN_PB, "+0.0;-0.0", "+0.000;-0.000"), _
            "+0.00;-0.00")

        ws.Cells(r, c + 6).Font.Bold = True

    Next i

    ws.Range(ws.Cells(r0 + 1, c), ws.Cells(r, c + 7)).Font.Size = 8

    ws.Columns(c).ColumnWidth = 22

    For i = 1 To 6
        ws.Columns(c + i).ColumnWidth = 9
    Next i

End Sub


' ============================================================
' PUNTOS DE EMISORES
' ============================================================

Private Sub Puntos( _
    ByVal ws As Worksheet, _
    ByVal cX As Long, _
    ByVal celEmi As String, _
    ByVal sPF As String, _
    ByVal nm As String, _
    ByVal fc As String)

    Dim x As String, y As String, cond As String

    x = ColL(cX)
    y = ColL(cX + 1)

    cond = _
        "(dbEmisor=" & celEmi & ")*" & _
        "(dbFecha>=$B$4)*" & _
        "(dbFecha<=$C$4)" & sPF

    ws.Range(x & "3").Formula2 = _
        "=IFERROR(INDEX(FILTER(dbFecha," & cond & ")," & _
        "ROWS($" & x & "$3:" & x & "3)),NA())"

    ws.Range(y & "3").Formula2 = _
        "=IFERROR(" & fc & _
        "*INDEX(FILTER(" & nm & "," & cond & ")," & _
        "ROWS($" & y & "$3:" & y & "3)),NA())"

    ws.Range(x & "3:" & y & "3").Copy _
        ws.Range(x & "3:" & y & (2 + N_PTOS))

    ws.Range(x & "3:" & x & (2 + N_PTOS)).NumberFormat = "dd/mm/yyyy"

End Sub


' ============================================================
' DIAGNÓSTICO
' ============================================================

Sub DiagnosticoDashCP()

    Dim wb As Workbook, db As Worksheet
    Dim hdr As Long, lastR As Long, r As Long, i As Long
    Dim cFec As Long, cEmi As Long, cSec As Long
    Dim cRat As Long, cSpr As Long, cPlz As Long
    Dim msg As String, v As Variant
    Dim fMin As Double, fMax As Double
    Dim nFec As Long, nPlz As Long, nSpr As Long, nVent As Long
    Dim d1 As Double, d2 As Double
    Dim dic As Object

    Set wb = ThisWorkbook
    Set db = HojaDB(wb)

    If db Is Nothing Then
        MsgBox "No encontré una hoja cuyo nombre contenga 'DB-Hist'.", vbExclamation
        Exit Sub
    End If

    hdr = FilaEncabezado(db)

    If hdr = 0 Then
        MsgBox "No encontré la fila de encabezados.", vbExclamation
        Exit Sub
    End If

    cFec = BuscarCol(db, hdr, "emision")
    cEmi = BuscarCol(db, hdr, "emisor")
    cSec = BuscarCol(db, hdr, "sector")
    cRat = BuscarCol(db, hdr, "rating")

    cPlz = BuscarCol(db, hdr, "anos")

    If cPlz = 0 Then
        cPlz = BuscarCol(db, hdr, "plazo")
    End If

    If cPlz = 0 Then
        cPlz = BuscarCol(db, hdr, "duracion")
    End If

    ' Corrección: buscar primero Spread (pbs)
    cSpr = BuscarCol(db, hdr, "spread (pbs)")

    If cSpr = 0 Then
        cSpr = BuscarCol(db, hdr, "spread (pb)")
    End If

    If cSpr = 0 Then
        cSpr = BuscarCol(db, hdr, "spread %")
    End If

    lastR = db.Cells(db.Rows.Count, cEmi).End(xlUp).Row

    msg = "HOJA: " & db.Name & vbCrLf & _
          "Encabezados en fila " & hdr & _
          ", datos " & hdr + 1 & " a " & lastR & vbCrLf & vbCrLf

    msg = msg & "COLUMNAS DETECTADAS" & vbCrLf
    msg = msg & "  Emisión = " & DHead(db, hdr, cFec) & vbCrLf
    msg = msg & "  Emisor  = " & DHead(db, hdr, cEmi) & vbCrLf
    msg = msg & "  Sector  = " & DHead(db, hdr, cSec) & vbCrLf
    msg = msg & "  Rating  = " & DHead(db, hdr, cRat) & vbCrLf
    msg = msg & "  Plazo   = " & DHead(db, hdr, cPlz) & vbCrLf
    msg = msg & "  Spread  = " & DHead(db, hdr, cSpr) & vbCrLf & vbCrLf

    msg = msg & "TODOS LOS ENCABEZADOS DE LA FILA " & hdr & vbCrLf

    For i = 1 To 40

        If Len(Trim$(CStr(db.Cells(hdr, i).Value))) > 0 Then
            msg = msg & "  " & ColL(i) & _
                  " = [" & db.Cells(hdr, i).Value & "]" & vbCrLf
        End If

    Next i

    msg = msg & vbCrLf

    d1 = CDbl(DateSerial(Year(Date) - 1, Month(Date), Day(Date)))
    d2 = CDbl(Date)

    fMin = 9999999
    fMax = 0

    Set dic = CreateObject("Scripting.Dictionary")

    For r = hdr + 1 To lastR

        v = db.Cells(r, cFec).Value

        If IsDate(v) Then

            nFec = nFec + 1

            If CDbl(v) < fMin Then fMin = CDbl(v)
            If CDbl(v) > fMax Then fMax = CDbl(v)

            If CDbl(v) >= d1 And CDbl(v) <= d2 Then
                nVent = nVent + 1
            End If

        End If

        If cPlz > 0 Then

            If IsNumeric(db.Cells(r, cPlz).Value) _
               And db.Cells(r, cPlz).Value <> "" Then

                nPlz = nPlz + 1

            End If

        End If

        If IsNumeric(db.Cells(r, cSpr).Value) _
           And db.Cells(r, cSpr).Value <> "" Then

            nSpr = nSpr + 1

        End If

        v = CStr(db.Cells(r, cRat).Value)

        If Len(v) > 0 Then

            If Not dic.Exists(v) Then
                dic.Add v, 0
            End If

            dic(v) = dic(v) + 1

        End If

    Next r

    msg = msg & "CALIDAD DE DATOS (" & lastR - hdr & " filas)" & vbCrLf
    msg = msg & "  Fechas válidas    : " & nFec & vbCrLf

    If nFec > 0 Then

        msg = msg & "  Rango de emisión  : " & _
              Format(fMin, "dd/mm/yyyy") & " a " & _
              Format(fMax, "dd/mm/yyyy") & vbCrLf

    End If

    msg = msg & "  En últimos 12 meses: " & nVent & _
          "  <-- si es 0, el gráfico sale vacío" & vbCrLf

    If cPlz > 0 Then

        msg = msg & "  Plazo numérico    : " & nPlz & _
              " de " & lastR - hdr & vbCrLf

    Else

        msg = msg & "  Plazo             : columna no detectada" & vbCrLf

    End If

    msg = msg & "  Spread numérico   : " & nSpr & _
          " de " & lastR - hdr & vbCrLf

    If nSpr > 0 Then

        msg = msg & "  Ejemplo de spread : " & _
              db.Cells(hdr + 1, cSpr).Value & vbCrLf

    End If

    msg = msg & vbCrLf & _
          "RATINGS ENCONTRADOS" & vbCrLf

    For i = 0 To dic.Count - 1

        msg = msg & "  [" & dic.Keys()(i) & "] = " & _
              dic.Items()(i) & vbCrLf

        If i >= 14 Then
            msg = msg & "  ..."
            Exit For
        End If

    Next i

    MsgBox msg, vbInformation, "Diagnóstico DB-Histórica CP"

End Sub


' ============================================================
' FUNCIONES AUXILIARES
' ============================================================

Private Function MaxCol( _
    db As Worksheet, _
    hdr As Long, _
    lastR As Long, _
    c As Long) As Double

    Dim r As Long, tope As Long, v As Variant

    tope = hdr + 300

    If lastR < tope Then
        tope = lastR
    End If

    For r = hdr + 1 To tope

        v = db.Cells(r, c).Value

        If IsNumeric(v) And v <> "" Then

            If CDbl(v) > MaxCol Then
                MaxCol = CDbl(v)
            End If

        End If

    Next r

End Function


Private Function FactorTasa( _
    db As Worksheet, _
    hdr As Long, _
    lastR As Long, _
    c As Long) As Double

    ' Deja la tasa expresada en porcentaje.
    ' 0.0455 se convierte en 4.55.
    If MaxCol(db, hdr, lastR, c) > 1 Then
        FactorTasa = 1
    Else
        FactorTasa = 100
    End If

End Function


Private Function FactorSpread( _
    db As Worksheet, _
    hdr As Long, _
    lastR As Long, _
    c As Long) As Double

    Dim mx As Double, esc As Double

    mx = MaxCol(db, hdr, lastR, c)

    If mx <= 0.5 Then

        ' Ejemplo: 0.065 = 6.5%
        esc = 100

    ElseIf mx <= 20 Then

        ' Ejemplo: 6.5 = 6.5%
        esc = 1

    Else

        ' Ejemplo: 300 = 300 pb
        esc = 0.01

    End If

    If SPREAD_EN_PB Then
        FactorSpread = esc * 100
    Else
        FactorSpread = esc
    End If

End Function


Private Sub PonerLista(c As Range, lst As String)

    With c.Validation

        .Delete

        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=lst

        .IgnoreBlank = True
        .InCellDropdown = True

    End With

End Sub


Private Sub DefNombre( _
    wb As Workbook, _
    nm As String, _
    hoja As String, _
    col As String, _
    f0 As Long, _
    f1 As Long)

    On Error Resume Next
    wb.Names(nm).Delete
    On Error GoTo 0

    wb.Names.Add Name:=nm, _
        RefersTo:="='" & hoja & "'!$" & _
        col & "$" & f0 & ":$" & col & "$" & f1

End Sub


Private Sub BorrarNombre(wb As Workbook, nm As String)

    On Error Resume Next
    wb.Names(nm).Delete
    On Error GoTo 0

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


Private Function BuscarCol( _
    db As Worksheet, _
    hdr As Long, _
    txt As String) As Long

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

        If Len(s) > 0 And _
           InStr(1, s, txt, vbTextCompare) > 0 Then

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


Private Function DHead( _
    db As Worksheet, _
    hdr As Long, _
    c As Long) As String

    If c < 1 Then

        DHead = "NO ENCONTRADA"
        Exit Function

    End If

    DHead = ColL(c) & "   [" & db.Cells(hdr, c).Value & "]"

End Function


'================================================================
' NUEVO: verifica si existe un rango con nombre en el libro
'================================================================
Private Function TieneNombre(wb As Workbook, nm As String) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = wb.Names(nm)
    On Error GoTo 0
    TieneNombre = Not n Is Nothing
End Function
