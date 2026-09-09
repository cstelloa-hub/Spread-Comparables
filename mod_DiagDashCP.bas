Attribute VB_Name = "mod_DiagDashCP"
Option Explicit

' Diagnostico para el dashboard de spreads CP.
' Ejecutar: DiagnosticoDashCP

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
    Set db = DHojaDB(wb)
    If db Is Nothing Then MsgBox "No encontre la hoja DB-Historica CP.", vbExclamation: Exit Sub
    hdr = DFilaEnc(db)
    If hdr = 0 Then MsgBox "No encontre la fila de encabezados.", vbExclamation: Exit Sub

    cFec = DCol(db, hdr, "emision")
    cEmi = DCol(db, hdr, "emisor")
    cSec = DCol(db, hdr, "sector")
    cRat = DCol(db, hdr, "rating")
    cPlz = DCol(db, hdr, "os vt")
    cSpr = DCol(db, hdr, "spread (pb)")
    If cSpr = 0 Then cSpr = DCol(db, hdr, "spread pb")
    If cSpr = 0 Then cSpr = DCol(db, hdr, "spread")

    lastR = db.Cells(db.Rows.Count, cEmi).End(xlUp).Row
    msg = "HOJA: " & db.Name & vbCrLf & _
          "Encabezados en fila " & hdr & ", datos " & hdr + 1 & " a " & lastR & vbCrLf & vbCrLf
    msg = msg & "COLUMNAS DETECTADAS" & vbCrLf
    msg = msg & "  Emision = " & DL(cFec) & "   [" & db.Cells(hdr, cFec).Value & "]" & vbCrLf
    msg = msg & "  Emisor  = " & DL(cEmi) & "   [" & db.Cells(hdr, cEmi).Value & "]" & vbCrLf
    msg = msg & "  Sector  = " & DL(cSec) & "   [" & db.Cells(hdr, cSec).Value & "]" & vbCrLf
    msg = msg & "  Rating  = " & DL(cRat) & "   [" & db.Cells(hdr, cRat).Value & "]" & vbCrLf
    msg = msg & "  Plazo   = " & DL(cPlz) & "   [" & db.Cells(hdr, cPlz).Value & "]" & vbCrLf
    msg = msg & "  Spread  = " & DL(cSpr) & "   [" & db.Cells(hdr, cSpr).Value & "]" & vbCrLf & vbCrLf

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
        If IsNumeric(db.Cells(r, cPlz).Value) And db.Cells(r, cPlz).Value <> "" Then nPlz = nPlz + 1
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
    msg = msg & "  Plazo numerico    : " & nPlz & " de " & lastR - hdr & "   <-- si es 0, el filtro de plazo mata todo" & vbCrLf
    msg = msg & "  Spread numerico   : " & nSpr & " de " & lastR - hdr & vbCrLf
    If nSpr > 0 Then msg = msg & "  Ejemplo de spread : " & db.Cells(hdr + 1, cSpr).Value & "   (si es 0.0045 el factor debe ser 10000)" & vbCrLf
    msg = msg & vbCrLf & "RATINGS ENCONTRADOS (entre corchetes, para ver espacios)" & vbCrLf
    For i = 0 To dic.Count - 1
        msg = msg & "  [" & dic.Keys()(i) & "] = " & dic.Items()(i) & vbCrLf
        If i >= 14 Then msg = msg & "  ...": Exit For
    Next i

    MsgBox msg, vbInformation, "Diagnostico DB-Historica CP"
End Sub

Private Function DHojaDB(wb As Workbook) As Worksheet
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If InStr(1, sh.Name, "DB-Hist", vbTextCompare) > 0 Then
            Set DHojaDB = sh
            Exit Function
        End If
    Next sh
End Function

Private Function DFilaEnc(db As Worksheet) As Long
    Dim r As Long, c As Long
    For r = 1 To 15
        For c = 1 To 60
            If LCase$(Trim$(CStr(db.Cells(r, c).Value))) = "emisor" Then
                DFilaEnc = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function DCol(db As Worksheet, hdr As Long, txt As String) As Long
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
            DCol = c
            Exit Function
        End If
    Next c
End Function

Private Function DL(i As Long) As String
    If i = 0 Then DL = "NO ENCONTRADA": Exit Function
    DL = Split(Cells(1, i).Address(True, False), "$")(0)
End Function
