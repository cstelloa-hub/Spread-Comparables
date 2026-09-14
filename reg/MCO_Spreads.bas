Attribute VB_Name = "modMCO"
'==============================================================================
' REGRESION MCO - DETERMINANTES DEL SPREAD DE COLOCACION
' Corre directamente sobre la hoja "DB-Historica CP (2)" de tu archivo.
'
' COMO INSTALARLA
'   1. Abre TU archivo (el que contiene la hoja de la base de datos)
'   2. Alt + F11  (abre el editor de VBA)
'   3. Archivo > Importar archivo... > elige MCO_Spreads.bas
'      (o: Insertar > Modulo, y pega todo este texto)
'   4. Cierra el editor, Alt + F8, elige CorrerMCO y dale Ejecutar
'   5. Guarda como .xlsm para no perder la macro
'
'   El nombre del archivo no importa: la macro trabaja siempre sobre el libro
'   donde esta instalada. Lo unico que necesita es encontrar la hoja de datos.
'
' QUE HACE
'   Crea tres hojas nuevas (las borra y rehace cada vez que la corres):
'     MCO_Datos      -> las emisiones que entran al modelo, con sus variables
'     MCO_Resultados -> coeficientes, errores estandar, t y p-values
'     MCO_Residuos   -> que spread pago cada emision vs. el que le tocaba,
'                       y el ranking de emisores por prima
'
' NO TOCA TU HOJA DE DATOS. Solo la lee.
'==============================================================================

Option Explicit

' ---- parametros que puedes cambiar ----------------------------------------
Private Const ANIO_DESDE   As Long = 2022      ' desde donde hay tasa base
Private Const MONEDA       As String = "PEN"
Private Const MIN_COLOC    As Long = 4         ' minimo de colocaciones para el ranking
Private Const HOJA_DATOS   As String = "MCO_Datos"
Private Const HOJA_RES     As String = "MCO_Resultados"
Private Const HOJA_RESID   As String = "MCO_Residuos"

' ---- indices de columna, se llenan solos -----------------------------------
Private cFecha As Long, cEmisor As Long, cSector As Long, cInstr As Long
Private cTipo As Long, cRating As Long, cDur As Long, cDivisa As Long
Private cTasaBase As Long, cSpread As Long, cAsignado As Long, cBTC As Long


'==============================================================================
' RUTINA PRINCIPAL
'==============================================================================
Public Sub CorrerMCO()

    Dim wsDB As Worksheet
    Dim ultima As Long, i As Long, j As Long, k As Long
    Dim nObs As Long, nX As Long, nAnios As Long
    Dim anioBase As Long
    Dim anios() As Long
    Dim Y() As Double, X() As Double
    Dim emisores() As String, fechas() As Double
    Dim res As Variant
    Dim t0 As Single: t0 = Timer

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' -------- 1. ubicar la hoja de datos --------
    Set wsDB = BuscarHojaDatos()
    If wsDB Is Nothing Then
        MsgBox "No encontre la hoja de datos." & vbCrLf & vbCrLf & _
               "Deberia llamarse ""DB-Historica CP (2)"" o contener ""DB"" y ""CP"" en el nombre.", _
               vbExclamation, "MCO"
        GoTo Salir
    End If

    ' -------- 2. ubicar las columnas por encabezado --------
    If Not UbicarColumnas(wsDB) Then GoTo Salir

    ultima = wsDB.Cells(wsDB.Rows.Count, cFecha).End(xlUp).Row

    ' -------- 3. primera pasada: contar filas validas y detectar anios --------
    ReDim anios(1 To 50)
    nAnios = 0
    nObs = 0
    For i = 2 To ultima
        If FilaValida(wsDB, i) Then
            nObs = nObs + 1
            k = Year(wsDB.Cells(i, cFecha).Value)
            If Not EstaEnLista(k, anios, nAnios) Then
                nAnios = nAnios + 1
                anios(nAnios) = k
            End If
        End If
    Next i

    If nObs < 30 Then
        MsgBox "Solo encontre " & nObs & " colocaciones validas. Muy pocas para el modelo." & vbCrLf & vbCrLf & _
               "Revisa que las columnas Spread (pbs), Tasa base, Duracion, Asignado y BTC tengan numeros.", _
               vbExclamation, "MCO"
        GoTo Salir
    End If

    Call OrdenarAscendente(anios, nAnios)
    anioBase = anios(1)

    ' 10 variables fijas + una dummy por cada anio salvo el base
    nX = 10 + (nAnios - 1)

    ' -------- 4. segunda pasada: armar la matriz --------
    ReDim Y(1 To nObs)
    ReDim X(1 To nObs, 1 To nX)
    ReDim emisores(1 To nObs)
    ReDim fechas(1 To nObs)

    Dim f As Long: f = 0
    Dim rt As String, anio As Long
    For i = 2 To ultima
        If FilaValida(wsDB, i) Then
            f = f + 1
            Y(f) = CDbl(wsDB.Cells(i, cSpread).Value)
            emisores(f) = Trim(CStr(wsDB.Cells(i, cEmisor).Value))
            fechas(f) = CDbl(wsDB.Cells(i, cFecha).Value)
            rt = Trim(UCase(CStr(wsDB.Cells(i, cRating).Value)))
            anio = Year(wsDB.Cells(i, cFecha).Value)

            X(f, 1) = IIf(rt = "CP-1", 1, 0)                              ' vs CP-1+
            X(f, 2) = IIf(rt = "CP-1-", 1, 0)
            X(f, 3) = IIf(rt = "CP-2" Or rt = "CP-2+", 1, 0)
            X(f, 4) = CDbl(wsDB.Cells(i, cTasaBase).Value) * 100          ' tasa base en %
            X(f, 5) = CDbl(wsDB.Cells(i, cDur).Value)                     ' duracion en anios
            X(f, 6) = Log(CDbl(wsDB.Cells(i, cAsignado).Value))           ' Log() en VBA es natural
            X(f, 7) = CDbl(wsDB.Cells(i, cBTC).Value)
            X(f, 8) = IIf(Trim(CStr(wsDB.Cells(i, cSector).Value)) <> "Financiero", 1, 0)
            ' cubre "Privada" y "Privado", que aparecen escritos de las dos formas
            X(f, 9) = IIf(InStr(1, CStr(wsDB.Cells(i, cTipo).Value), "Privad", vbTextCompare) > 0, 1, 0)
            X(f, 10) = IIf(Trim(CStr(wsDB.Cells(i, cInstr).Value)) <> "Certificados de Depósito", 1, 0)

            For j = 2 To nAnios
                X(f, 10 + j - 1) = IIf(anio = anios(j), 1, 0)
            Next j
        End If
    Next i

    ' -------- 5. escribir la hoja de datos --------
    Dim wsD As Worksheet
    Set wsD = RehacerHoja(HOJA_DATOS)
    Call EscribirDatos(wsD, Y, X, emisores, fechas, nObs, nX, nAnios, anios, anioBase)

    ' -------- 6. estimar --------
    Dim rgY As Range, rgX As Range
    Set rgY = wsD.Range(wsD.Cells(2, 3), wsD.Cells(nObs + 1, 3))
    Set rgX = wsD.Range(wsD.Cells(2, 4), wsD.Cells(nObs + 1, 3 + nX))
    res = Application.WorksheetFunction.LinEst(rgY, rgX, True, True)

    ' -------- 7. resultados y residuos --------
    Call EscribirResultados(res, nX, nObs, nAnios, anios, anioBase)
    Call EscribirResiduos(res, Y, X, emisores, fechas, nObs, nX)

    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(HOJA_RES).Activate

    MsgBox "Listo en " & Format(Timer - t0, "0.0") & " segundos." & vbCrLf & vbCrLf & _
           "Colocaciones usadas: " & nObs & vbCrLf & _
           "Variables explicativas: " & nX & vbCrLf & _
           "R cuadrado: " & Format(res(3, 1), "0.000") & vbCrLf & vbCrLf & _
           "Revisa las hojas " & HOJA_DATOS & ", " & HOJA_RES & " y " & HOJA_RESID & ".", _
           vbInformation, "MCO"
Salir:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Se detuvo con un error." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Lo mas comun es que alguna celda de Spread, Tasa base, Duracion, Asignado o BTC " & _
           "tenga texto en vez de numero.", vbExclamation, "MCO"
End Sub


'==============================================================================
' FILTRO: que filas entran al modelo
'==============================================================================
Private Function FilaValida(ws As Worksheet, i As Long) As Boolean

    Dim rt As String
    FilaValida = False

    If Not IsDate(ws.Cells(i, cFecha).Value) Then Exit Function
    If Year(ws.Cells(i, cFecha).Value) < ANIO_DESDE Then Exit Function
    If Trim(UCase(CStr(ws.Cells(i, cDivisa).Value))) <> UCase(MONEDA) Then Exit Function

    rt = Trim(UCase(CStr(ws.Cells(i, cRating).Value)))
    Select Case rt
        Case "CP-1+", "CP-1", "CP-1-", "CP-2+", "CP-2"
        Case Else: Exit Function
    End Select

    If Not EsNumPositivo(ws.Cells(i, cSpread).Value) Then Exit Function
    If Not EsNumPositivo(ws.Cells(i, cTasaBase).Value) Then Exit Function
    If Not EsNumPositivo(ws.Cells(i, cDur).Value) Then Exit Function
    If Not EsNumPositivo(ws.Cells(i, cAsignado).Value) Then Exit Function
    If Not EsNumPositivo(ws.Cells(i, cBTC).Value) Then Exit Function

    FilaValida = True
End Function

Private Function EsNumPositivo(v As Variant) As Boolean
    EsNumPositivo = False
    If IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CDbl(v) <= 0 Then Exit Function
    EsNumPositivo = True
End Function


'==============================================================================
' UBICAR HOJA Y COLUMNAS
'==============================================================================
Private Function BuscarHojaDatos() As Worksheet

    Dim ws As Worksheet
    On Error Resume Next
    Set BuscarHojaDatos = ThisWorkbook.Worksheets("DB-Histórica CP (2)")
    On Error GoTo 0
    If Not BuscarHojaDatos Is Nothing Then Exit Function

    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "DB", vbTextCompare) > 0 And _
           InStr(1, ws.Name, "CP", vbTextCompare) > 0 Then
            Set BuscarHojaDatos = ws
            Exit Function
        End If
    Next ws
End Function


Private Function UbicarColumnas(ws As Worksheet) As Boolean

    UbicarColumnas = False

    cFecha = Col(ws, "Emisión"):        If cFecha = 0 Then cFecha = Col(ws, "Emision")
    cEmisor = Col(ws, "Emisor")
    cSector = Col(ws, "Sector")
    cInstr = Col(ws, "Instrumento")
    cTipo = Col(ws, "Tipo de emisión"): If cTipo = 0 Then cTipo = Col(ws, "Tipo de emision")
    cRating = Col(ws, "RATING")
    cDur = Col(ws, "Duración"):         If cDur = 0 Then cDur = Col(ws, "Duracion")
    cDivisa = Col(ws, "Divisa")
    cTasaBase = Col(ws, "Tasa base")
    cSpread = Col(ws, "Spread (pbs)"):  If cSpread = 0 Then cSpread = Col(ws, "Spread (pb")
    cAsignado = Col(ws, "Asignado")
    cBTC = Col(ws, "BTC")

    Dim faltan As String
    If cFecha = 0 Then faltan = faltan & vbCrLf & " - Emisión"
    If cEmisor = 0 Then faltan = faltan & vbCrLf & " - Emisor"
    If cSector = 0 Then faltan = faltan & vbCrLf & " - Sector"
    If cInstr = 0 Then faltan = faltan & vbCrLf & " - Instrumento"
    If cTipo = 0 Then faltan = faltan & vbCrLf & " - Tipo de emisión"
    If cRating = 0 Then faltan = faltan & vbCrLf & " - RATING"
    If cDur = 0 Then faltan = faltan & vbCrLf & " - Duración"
    If cDivisa = 0 Then faltan = faltan & vbCrLf & " - Divisa"
    If cTasaBase = 0 Then faltan = faltan & vbCrLf & " - Tasa base"
    If cSpread = 0 Then faltan = faltan & vbCrLf & " - Spread (pbs)"
    If cAsignado = 0 Then faltan = faltan & vbCrLf & " - Asignado"
    If cBTC = 0 Then faltan = faltan & vbCrLf & " - BTC"

    If Len(faltan) > 0 Then
        MsgBox "No encontre estas columnas en la fila 1 de """ & ws.Name & """:" & faltan, _
               vbExclamation, "MCO"
        Exit Function
    End If

    UbicarColumnas = True
End Function


' Busca un encabezado en la fila 1. Ignora mayusculas y espacios sobrantes,
' por eso funciona con "RATING " y con "Demanda  MM".
Private Function Col(ws As Worksheet, encabezado As String) As Long

    Dim c As Long, ultimaCol As Long, txt As String
    ultimaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If ultimaCol < 1 Then ultimaCol = 100

    For c = 1 To ultimaCol
        txt = Trim(CStr(ws.Cells(1, c).Value))
        If StrComp(txt, Trim(encabezado), vbTextCompare) = 0 Then
            Col = c: Exit Function
        End If
    Next c

    ' segundo intento: coincidencia parcial
    For c = 1 To ultimaCol
        txt = Trim(CStr(ws.Cells(1, c).Value))
        If Len(txt) > 0 Then
            If InStr(1, txt, Trim(encabezado), vbTextCompare) = 1 Then
                Col = c: Exit Function
            End If
        End If
    Next c

    Col = 0
End Function


'==============================================================================
' ESCRITURA DE HOJAS
'==============================================================================
Private Function RehacerHoja(nombre As String) As Worksheet

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nombre).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set RehacerHoja = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    RehacerHoja.Name = nombre
    RehacerHoja.Cells.Font.Name = "Arial"
    RehacerHoja.Cells.Font.Size = 10
End Function


Private Sub EscribirDatos(ws As Worksheet, Y() As Double, X() As Double, _
                          emisores() As String, fechas() As Double, _
                          nObs As Long, nX As Long, nAnios As Long, _
                          anios() As Long, anioBase As Long)

    Dim nombres() As String
    ReDim nombres(1 To nX)
    nombres(1) = "d_CP-1"
    nombres(2) = "d_CP-1-"
    nombres(3) = "d_CP-2"
    nombres(4) = "Tasa base %"
    nombres(5) = "Duración"
    nombres(6) = "ln(Monto)"
    nombres(7) = "BTC"
    nombres(8) = "d_No financiero"
    nombres(9) = "d_Oferta privada"
    nombres(10) = "d_No es CD"

    Dim j As Long
    For j = 2 To nAnios
        nombres(10 + j - 1) = "d_" & anios(j)
    Next j

    ws.Cells(1, 1).Value = "Fecha"
    ws.Cells(1, 2).Value = "Emisor"
    ws.Cells(1, 3).Value = "Spread (pb)  [Y]"
    For j = 1 To nX
        ws.Cells(1, 3 + j).Value = nombres(j) & "  [X" & j & "]"
    Next j

    Dim i As Long
    Dim datos() As Variant
    ReDim datos(1 To nObs, 1 To 3 + nX)
    For i = 1 To nObs
        datos(i, 1) = fechas(i)
        datos(i, 2) = emisores(i)
        datos(i, 3) = Y(i)
        For j = 1 To nX
            datos(i, 3 + j) = X(i, j)
        Next j
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(nObs + 1, 3 + nX)).Value = datos

    ws.Range(ws.Cells(2, 1), ws.Cells(nObs + 1, 1)).NumberFormat = "dd/mm/yyyy"
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 3 + nX))
        .Font.Bold = True
        .Interior.Color = RGB(31, 31, 31)
        .Font.Color = RGB(255, 255, 255)
        .WrapText = True
    End With
    ws.Rows(1).RowHeight = 30
    ws.Columns("A:B").ColumnWidth = 20
    ws.Range(ws.Cells(1, 3), ws.Cells(1, 3 + nX)).EntireColumn.ColumnWidth = 12
    ws.Range("C2").Select
    ActiveWindow.FreezePanes = False
    ws.Range("C2").Select
    ActiveWindow.FreezePanes = True

    ws.Cells(nObs + 3, 1).Value = "Categoría base: CP-1+, sector financiero, oferta pública, " & _
        "certificado de depósito, año " & anioBase & ". Todos los coeficientes se leen contra ese perfil."
    ws.Cells(nObs + 3, 1).Font.Italic = True
End Sub


Private Sub EscribirResultados(res As Variant, nX As Long, nObs As Long, _
                               nAnios As Long, anios() As Long, anioBase As Long)

    Dim ws As Worksheet
    Set ws = RehacerHoja(HOJA_RES)

    Dim nombres() As String
    ReDim nombres(1 To nX)
    nombres(1) = "Rating CP-1 (vs CP-1+)"
    nombres(2) = "Rating CP-1-"
    nombres(3) = "Rating CP-2"
    nombres(4) = "Tasa base (por cada +1 pp)"
    nombres(5) = "Duración (por cada +1 año)"
    nombres(6) = "ln(Monto asignado)"
    nombres(7) = "Bid-to-cover (por cada +1x)"
    nombres(8) = "Sector no financiero"
    nombres(9) = "Oferta privada"
    nombres(10) = "Instrumento distinto de CD"
    Dim j As Long
    For j = 2 To nAnios
        nombres(10 + j - 1) = "Año " & anios(j) & " (vs " & anioBase & ")"
    Next j

    Dim r2 As Double, gl As Double, fStat As Double
    r2 = res(3, 1): gl = res(4, 2): fStat = res(4, 1)

    ws.Range("A1").Value = "Regresión MCO - determinantes del spread de colocación"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Variable dependiente: spread en puntos básicos sobre la curva del mismo plazo."
    ws.Range("A2").Font.Italic = True

    ws.Range("A4").Value = "Variable"
    ws.Range("B4").Value = "Coeficiente (pb)"
    ws.Range("C4").Value = "Error estándar"
    ws.Range("D4").Value = "Estadístico t"
    ws.Range("E4").Value = "p-value"
    ws.Range("F4").Value = "Significancia"
    With ws.Range("A4:F4")
        .Font.Bold = True
        .Interior.Color = RGB(31, 31, 31)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' intercepto: es la ULTIMA columna que devuelve LinEst
    Dim fila As Long: fila = 5
    ws.Cells(fila, 1).Value = "Intercepto"
    Call FilaCoef(ws, fila, res(1, nX + 1), res(2, nX + 1), gl)

    ' el resto viene en orden INVERSO al de las columnas X
    Dim i As Long
    For i = 1 To nX
        fila = fila + 1
        ws.Cells(fila, 1).Value = nombres(i)
        Call FilaCoef(ws, fila, res(1, nX + 1 - i), res(2, nX + 1 - i), gl)
    Next i

    fila = fila + 2
    ws.Cells(fila, 1).Value = "BONDAD DE AJUSTE"
    ws.Cells(fila, 1).Font.Bold = True
    ws.Cells(fila, 1).Font.Color = RGB(236, 17, 26)
    ws.Cells(fila + 1, 1).Value = "R cuadrado":            ws.Cells(fila + 1, 2).Value = r2
    ws.Cells(fila + 2, 1).Value = "R cuadrado ajustado":   ws.Cells(fila + 2, 2).Value = 1 - (1 - r2) * (nObs - 1) / gl
    ws.Cells(fila + 3, 1).Value = "Error estándar (pb)":   ws.Cells(fila + 3, 2).Value = res(3, 2)
    ws.Cells(fila + 4, 1).Value = "Estadístico F":         ws.Cells(fila + 4, 2).Value = fStat
    ws.Cells(fila + 5, 1).Value = "Grados de libertad":    ws.Cells(fila + 5, 2).Value = gl
    ws.Cells(fila + 6, 1).Value = "N (observaciones)":     ws.Cells(fila + 6, 2).Value = nObs
    ws.Range(ws.Cells(fila + 1, 1), ws.Cells(fila + 6, 1)).Font.Bold = True
    ws.Range(ws.Cells(fila + 1, 2), ws.Cells(fila + 4, 2)).NumberFormat = "0.000"

    fila = fila + 8
    ws.Cells(fila, 1).Value = "*** p<0.01   ** p<0.05   * p<0.10   n.s. no significativo"
    ws.Cells(fila + 1, 1).Value = "Errores estándar clásicos. La presentación usa errores clusterizados " & _
        "por emisor, que Excel no calcula de forma nativa. Los coeficientes son los mismos."
    ws.Cells(fila + 2, 1).Value = "El coeficiente negativo de Duración refleja composición de la muestra " & _
        "(los plazos cortos son papeles comerciales de emisores con peor rating), no una curva invertida."
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila + 2, 1)).Font.Italic = True
    ws.Range(ws.Cells(fila, 1), ws.Cells(fila + 2, 1)).Font.Size = 9

    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B:F").ColumnWidth = 17
End Sub


Private Sub FilaCoef(ws As Worksheet, fila As Long, coef As Double, ee As Double, gl As Double)

    Dim t As Double, p As Double, sig As String

    ws.Cells(fila, 2).Value = coef
    ws.Cells(fila, 3).Value = ee

    If ee > 0 Then
        t = coef / ee
        ws.Cells(fila, 4).Value = t
        p = 1
        On Error Resume Next
        p = Application.WorksheetFunction.TDist(Abs(t), gl, 2)
        On Error GoTo 0
        ws.Cells(fila, 5).Value = p
        If p < 0.01 Then
            sig = "***"
        ElseIf p < 0.05 Then
            sig = "**"
        ElseIf p < 0.1 Then
            sig = "*"
        Else
            sig = "n.s."
        End If
        ws.Cells(fila, 6).Value = sig
        If sig <> "n.s." Then ws.Cells(fila, 6).Font.Color = RGB(236, 17, 26)
    End If

    ws.Cells(fila, 1).Font.Bold = True
    ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 4)).NumberFormat = "0.00"
    ws.Cells(fila, 5).NumberFormat = "0.000"
    ws.Cells(fila, 6).HorizontalAlignment = xlCenter
    With ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 6)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(216, 216, 216)
    End With
End Sub


'==============================================================================
' RESIDUOS: que pago cada emision vs. lo que le tocaba, y ranking por emisor
'==============================================================================
Private Sub EscribirResiduos(res As Variant, Y() As Double, X() As Double, _
                             emisores() As String, fechas() As Double, _
                             nObs As Long, nX As Long)

    Dim ws As Worksheet
    Set ws = RehacerHoja(HOJA_RESID)

    Dim i As Long, j As Long
    Dim pred As Double
    Dim dSum As Object, dN As Object
    Set dSum = CreateObject("Scripting.Dictionary")
    Set dN = CreateObject("Scripting.Dictionary")

    ws.Range("A1").Value = "¿Qué emisiones pagaron más de lo que les tocaba?"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Prima = spread que pagó menos el que el modelo estima por su rating, plazo, tamaño y año."
    ws.Range("A2").Font.Italic = True

    ws.Range("A4").Value = "Fecha"
    ws.Range("B4").Value = "Emisor"
    ws.Range("C4").Value = "Spread real (pb)"
    ws.Range("D4").Value = "Spread estimado (pb)"
    ws.Range("E4").Value = "Prima (pb)"
    With ws.Range("A4:E4")
        .Font.Bold = True
        .Interior.Color = RGB(31, 31, 31)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With

    Dim salida() As Variant
    ReDim salida(1 To nObs, 1 To 5)
    Dim e As String, extra As Double

    For i = 1 To nObs
        pred = res(1, nX + 1)                       ' intercepto
        For j = 1 To nX
            pred = pred + res(1, nX + 1 - j) * X(i, j)
        Next j
        extra = Y(i) - pred
        e = emisores(i)

        salida(i, 1) = fechas(i)
        salida(i, 2) = e
        salida(i, 3) = Y(i)
        salida(i, 4) = pred
        salida(i, 5) = extra

        If dSum.Exists(e) Then
            dSum(e) = dSum(e) + extra
            dN(e) = dN(e) + 1
        Else
            dSum.Add e, extra
            dN.Add e, 1
        End If
    Next i

    ws.Range(ws.Cells(5, 1), ws.Cells(nObs + 4, 5)).Value = salida
    ws.Range(ws.Cells(5, 1), ws.Cells(nObs + 4, 1)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(5, 3), ws.Cells(nObs + 4, 5)).NumberFormat = "0.0"

    ' ---- ranking por emisor ----
    Dim fila As Long: fila = 4
    ws.Cells(fila, 7).Value = "RANKING POR EMISOR"
    ws.Cells(fila, 7).Font.Bold = True
    ws.Cells(fila, 7).Font.Color = RGB(236, 17, 26)
    ws.Cells(fila + 1, 7).Value = "Emisor"
    ws.Cells(fila + 1, 8).Value = "N"
    ws.Cells(fila + 1, 9).Value = "Prima promedio (pb)"
    With ws.Range(ws.Cells(fila + 1, 7), ws.Cells(fila + 1, 9))
        .Font.Bold = True
        .Interior.Color = RGB(31, 31, 31)
        .Font.Color = RGB(255, 255, 255)
        .WrapText = True
    End With

    Dim claves As Variant: claves = dSum.Keys
    Dim n As Long: n = dSum.Count
    Dim prom() As Double
    ReDim prom(0 To n - 1)
    For i = 0 To n - 1
        prom(i) = dSum(claves(i)) / dN(claves(i))
    Next i

    ' ordenar de mayor a menor prima (burbuja: son pocos emisores)
    Dim tmpD As Double, tmpS As Variant, a As Long, b As Long
    For a = 0 To n - 2
        For b = 0 To n - 2 - a
            If prom(b) < prom(b + 1) Then
                tmpD = prom(b): prom(b) = prom(b + 1): prom(b + 1) = tmpD
                tmpS = claves(b): claves(b) = claves(b + 1): claves(b + 1) = tmpS
            End If
        Next b
    Next a

    Dim fr As Long: fr = fila + 2
    For i = 0 To n - 1
        If dN(claves(i)) >= MIN_COLOC Then
            ws.Cells(fr, 7).Value = claves(i)
            ws.Cells(fr, 8).Value = dN(claves(i))
            ws.Cells(fr, 9).Value = prom(i)
            ws.Cells(fr, 9).NumberFormat = "0.0"
            If prom(i) > 0 Then
                ws.Cells(fr, 9).Font.Color = RGB(0, 120, 60)
            Else
                ws.Cells(fr, 9).Font.Color = RGB(236, 17, 26)
            End If
            fr = fr + 1
        End If
    Next i

    ws.Cells(fr + 1, 7).Value = "Solo emisores con " & MIN_COLOC & " o más colocaciones. " & _
        "Verde = paga por encima del modelo."
    ws.Cells(fr + 1, 7).Font.Italic = True
    ws.Cells(fr + 1, 7).Font.Size = 9

    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C:E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 3
    ws.Columns("G").ColumnWidth = 24
    ws.Columns("H").ColumnWidth = 6
    ws.Columns("I").ColumnWidth = 16
    ws.Rows(5).RowHeight = 28
End Sub


'==============================================================================
' UTILIDADES
'==============================================================================
Private Function EstaEnLista(v As Long, lista() As Long, n As Long) As Boolean
    Dim i As Long
    For i = 1 To n
        If lista(i) = v Then EstaEnLista = True: Exit Function
    Next i
    EstaEnLista = False
End Function

Private Sub OrdenarAscendente(lista() As Long, n As Long)
    Dim i As Long, j As Long, t As Long
    For i = 1 To n - 1
        For j = 1 To n - i
            If lista(j) > lista(j + 1) Then
                t = lista(j): lista(j) = lista(j + 1): lista(j + 1) = t
            End If
        Next j
    Next i
End Sub
