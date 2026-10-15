Attribute VB_Name = "CubicadorAutomatico"
'==============================================================================
' CUBICADOR AUTOMATICO
'
' Flujo diario:
'   1. CubicarExportSAP          -> lee el export de SAP, toma solo las entregas
'                                    que NO estan en CRUCE, calcula cubicaje,
'                                    valor y zona, y arma el resumen por zona y
'                                    ciudad con el vehiculo sugerido.
'   2. GenerarCorreoTransporte   -> deja un borrador de Outlook con el resumen
'                                    (no lo envia).
'   3. MarcarRemisionesCubicadas -> agrega las remisiones procesadas a CRUCE
'                                    para que no se vuelvan a cubicar.
'
' Hojas que usa del cubicador (no se modifican, salvo CRUCE en el paso 3):
'   MATERIAL         A = IP, C = volumen teorico m3/und
'   VALO MATERIAL    C = IP, D = valor sin IVA (desde la fila 4)
'   llanta_zona      J = ciudad, K = zona
'   CRUCE            A = remisiones ya cubicadas
'   CUBICAJE BODEGA  B3:C6 = tipo de vehiculo y m3 maximo
'
' Hojas que crea/reemplaza: CUBICAR HOY, RESUMEN TRANSPORTE, ALERTAS
'==============================================================================
Option Explicit

' --- Configuracion ----------------------------------------------------------
Private Const CORREO_TRANSPORTADORA As String = ""   ' ej. "programacion@transportadora.com"
Private Const CORREO_COPIA As String = ""
Private Const TOPE_SATURACION As Double = 1#        ' 1 = 100 %. Bajelo (ej. 0.9) si quiere margen.
Private Const CENTRO_CUBICADOR As String = "CDLI"

Private Const HOJA_DETALLE As String = "CUBICAR HOY"
Private Const HOJA_RESUMEN As String = "RESUMEN TRANSPORTE"
Private Const HOJA_ALERTAS As String = "ALERTAS"
Private Const SIN_ZONA As String = "SIN ZONA"

' Vehiculos (se leen de CUBICAJE BODEGA!B3:C6)
Private vehNombre() As String
Private vehCap() As Double

'==============================================================================
' 1. CUBICAR EXPORT DE SAP
'==============================================================================
Public Sub CubicarExportSAP()
    Dim archivo As Variant, wbSAP As Workbook, datos As Variant
    Dim dMat As Object, dVal As Object, dZona As Object, dCruce As Object
    Dim alertas As Collection

    archivo = Application.GetOpenFilename("Excel (*.xls*),*.xls*", , "Seleccione el export de SAP del dia")
    If VarType(archivo) = vbBoolean Then Exit Sub

    Dim corte As String
    corte = InputBox("Texto de la Fecha Corte:", "Fecha de corte", "1 Corte " & MesEnEspanol(Date) & " " & Format(Date, "dd"))
    If Len(Trim(corte)) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Fallo

    LeerVehiculos
    Set dMat = Diccionario(ThisWorkbook.Worksheets("MATERIAL"), 1, 3, 2)
    Set dVal = Diccionario(ThisWorkbook.Worksheets("VALO MATERIAL"), 3, 4, 4)
    Set dZona = DiccionarioCiudades(ThisWorkbook.Worksheets("llanta_zona"), 10, 11, 2)
    Set dCruce = Diccionario(ThisWorkbook.Worksheets("CRUCE"), 1, 1, 2)
    Set alertas = New Collection

    Set wbSAP = Workbooks.Open(archivo, ReadOnly:=True)
    datos = wbSAP.Worksheets(1).UsedRange.Value
    wbSAP.Close SaveChanges:=False

    ' Columnas del export (se busca por nombre; si un nombre se repite se toma la primera)
    Dim cFecha&, cEntrega&, cCliente&, cCiudad&, cMaterial&, cDesc&, cCant&, cPeso&, cCodSAP&, cDir&
    cFecha = Col(datos, "Fec.Entrega"):        cEntrega = Col(datos, "Entrega")
    cCliente = Col(datos, "Nombre Recibid"):   cCiudad = Col(datos, "Ciudad Recibidor")
    cMaterial = Col(datos, "Material"):        cDesc = Col(datos, "Descripc.Material")
    cCant = Col(datos, "Cantidad Entrega"):    cPeso = Col(datos, "Peso Bruto")
    cCodSAP = Col(datos, "Cold.Recibidor"):    cDir = Col(datos, "Dir.Entrega")

    ' --- Detalle -------------------------------------------------------------
    Const NC As Long = 17
    Dim out() As Variant, n As Long, i As Long
    ReDim out(1 To UBound(datos, 1), 1 To NC)
    Dim yaCruce As Long, nuevas As Object: Set nuevas = CreateObject("Scripting.Dictionary")
    Dim yaContada As Object: Set yaContada = CreateObject("Scripting.Dictionary")

    For i = 2 To UBound(datos, 1)
        Dim rem_ As String, cant As Double, mat As String, ciudad As String, zona As String
        rem_ = Clave(datos(i, cEntrega))
        cant = Num(datos(i, cCant))
        If Len(rem_) > 0 And cant > 0 Then
            If dCruce.Exists(rem_) Then
                If Not yaContada.Exists(rem_) Then yaContada.Add rem_, True: yaCruce = yaCruce + 1
            Else
                n = n + 1
                nuevas(rem_) = True
                mat = Clave(datos(i, cMaterial))
                ciudad = Trim(CStr(datos(i, cCiudad)))

                If dZona.Exists(NormCiudad(ciudad)) Then
                    zona = dZona(NormCiudad(ciudad))
                Else
                    zona = SIN_ZONA
                    Alerta alertas, "Ciudad sin zona en llanta_zona", ciudad, rem_
                End If
                If Not dMat.Exists(mat) Then Alerta alertas, "Material sin volumen en MATERIAL (cubicaje = 0)", mat & " " & datos(i, cDesc), rem_
                If Not dVal.Exists(mat) Then Alerta alertas, "Material sin valor en VALO MATERIAL", mat & " " & datos(i, cDesc), rem_

                out(n, 1) = datos(i, cFecha)
                out(n, 2) = corte
                out(n, 3) = CENTRO_CUBICADOR
                out(n, 4) = CDbl(rem_)
                out(n, 5) = datos(i, cCliente)
                out(n, 6) = ciudad
                out(n, 7) = zona
                out(n, 8) = Val(mat)
                out(n, 9) = datos(i, cDesc)
                out(n, 10) = cant
                out(n, 11) = Num(datos(i, cPeso))
                out(n, 12) = datos(i, cCodSAP)
                out(n, 13) = datos(i, cDir)
                out(n, 14) = IIf(dMat.Exists(mat), Num(dMat(mat)), 0)
                out(n, 15) = out(n, 14) * cant
                out(n, 16) = IIf(Left(rem_, 2) = "29", "PIRELLI TYRE", "PROMETEON")
                out(n, 17) = IIf(dVal.Exists(mat), Num(dVal(mat)), 0) * cant
            End If
        End If
    Next i

    If n = 0 Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "No hay remisiones nuevas: todas las entregas del export ya estan en CRUCE (" & yaCruce & ").", vbInformation
        Exit Sub
    End If

    Dim wsD As Worksheet: Set wsD = HojaNueva(HOJA_DETALLE)
    wsD.Range("A1").Resize(1, NC).Value = Array("Fecha Remision", "Fecha Corte", "CENTRO", "Remision", "Cliente", _
        "Ciudad", "Zona", "Codigo", "Descripcion", "Cantidad", "Peso (kg)", "Codigo SAP", "Direccion", _
        "Cubicaje/und", "Cubicaje Total", "EMPRESA", "Valor total sin IVA")
    wsD.Range("A2").Resize(n, NC).Value = out
    wsD.Columns("A").NumberFormat = "dd/mm/yyyy"
    wsD.Columns("D").NumberFormat = "0"
    wsD.Columns("N:O").NumberFormat = "0.000"
    wsD.Columns("Q").NumberFormat = "#,##0"
    FormatoTabla wsD, n + 1, NC

    ' --- Resumen ---------------------------------------------------------------
    ArmarResumen out, n

    ' --- Alertas ---------------------------------------------------------------
    Dim wsA As Worksheet: Set wsA = HojaNueva(HOJA_ALERTAS)
    wsA.Range("A1:C1").Value = Array("Alerta", "Dato", "Remision")
    If alertas.Count = 0 Then
        wsA.Range("A2").Value = "Sin alertas"
    Else
        For i = 1 To alertas.Count
            wsA.Cells(i + 1, 1).Resize(1, 3).Value = alertas(i)
        Next i
    End If
    FormatoTabla wsA, alertas.Count + 1, 3

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(HOJA_RESUMEN).Activate

    MsgBox "Remisiones nuevas: " & nuevas.Count & " (" & n & " lineas)" & vbCrLf & _
           "Ya cubicadas antes (en CRUCE): " & yaCruce & vbCrLf & _
           "Alertas: " & alertas.Count & IIf(alertas.Count > 0, "  -> revise la hoja " & HOJA_ALERTAS, ""), _
           IIf(alertas.Count > 0, vbExclamation, vbInformation), "Cubicaje listo"
    Exit Sub

Fallo:
    Dim msgError As String: msgError = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wbSAP Is Nothing Then wbSAP.Close SaveChanges:=False
    MsgBox "Error: " & msgError, vbCritical, "Cubicador"
End Sub

' Resumen por zona y por ciudad: m3, kilos, unidades, valor, saturacion y vehiculo
Private Sub ArmarResumen(out As Variant, n As Long)
    Dim grpZona As Object, grpCiudad As Object, i As Long
    Set grpZona = CreateObject("Scripting.Dictionary")
    Set grpCiudad = CreateObject("Scripting.Dictionary")

    For i = 1 To n
        Acumular grpZona, CStr(out(i, 7)), "", out, i
        Acumular grpCiudad, CStr(out(i, 7)) & "|" & NormCiudad(CStr(out(i, 6))), NormCiudad(CStr(out(i, 6))), out, i
    Next i

    Dim ws As Worksheet: Set ws = HojaNueva(HOJA_RESUMEN)
    Dim enc As Variant, k As Long, nv As Long
    nv = UBound(vehNombre)
    ReDim enc(1 To 7 + nv + 1)
    enc(1) = "Zona": enc(2) = "Ciudad": enc(3) = "Remisiones": enc(4) = "Unidades"
    enc(5) = "Kilos": enc(6) = "Cubicaje m3": enc(7) = "Valor sin IVA"
    For k = 1 To nv: enc(7 + k) = "% " & vehNombre(k) & " (" & vehCap(k) & " m3)": Next k
    enc(8 + nv) = "Vehiculo sugerido"

    Dim fila As Long: fila = 1
    ws.Cells(fila, 1).Value = "POR ZONA": ws.Cells(fila, 1).Font.Bold = True
    fila = fila + 1
    fila = EscribirGrupos(ws, fila, enc, grpZona, False)

    fila = fila + 2
    ws.Cells(fila, 1).Value = "POR CIUDAD": ws.Cells(fila, 1).Font.Bold = True
    fila = fila + 1
    fila = EscribirGrupos(ws, fila, enc, grpCiudad, True)

    ws.Columns.AutoFit
End Sub

Private Sub Acumular(d As Object, clave As String, ciudad As String, out As Variant, i As Long)
    Dim g As Variant
    If Not d.Exists(clave) Then
        ' zona, ciudad, remisiones(dict), unidades, kilos, m3, valor
        d.Add clave, Array(CStr(out(i, 7)), ciudad, CreateObject("Scripting.Dictionary"), 0#, 0#, 0#, 0#)
    End If
    g = d(clave)
    g(2).Item(CStr(out(i, 4))) = True
    g(3) = g(3) + out(i, 10)
    g(4) = g(4) + out(i, 11)
    g(5) = g(5) + out(i, 15)
    g(6) = g(6) + out(i, 17)
    d(clave) = g
End Sub

Private Function EscribirGrupos(ws As Worksheet, fila As Long, enc As Variant, d As Object, conCiudad As Boolean) As Long
    Dim nv As Long, k As Long, c As Long, key As Variant, g As Variant, ini As Long
    Dim claves As Variant
    nv = UBound(vehNombre)
    ini = fila
    ws.Cells(fila, 1).Resize(1, UBound(enc)).Value = enc
    claves = Ordenadas(d)
    For Each key In claves
        g = d(key)
        fila = fila + 1
        ws.Cells(fila, 1).Value = g(0)
        ws.Cells(fila, 2).Value = IIf(conCiudad, g(1), "(todas)")
        ws.Cells(fila, 3).Value = g(2).Count
        ws.Cells(fila, 4).Value = g(3)
        ws.Cells(fila, 5).Value = g(4)
        ws.Cells(fila, 6).Value = g(5)
        ws.Cells(fila, 7).Value = g(6)
        For k = 1 To nv
            ws.Cells(fila, 7 + k).Value = g(5) / vehCap(k)
        Next k
        ws.Cells(fila, 8 + nv).Value = VehiculoSugerido(CDbl(g(5)))
    Next key
    ws.Range(ws.Cells(ini + 1, 5), ws.Cells(fila, 5)).NumberFormat = "#,##0.0"
    ws.Range(ws.Cells(ini + 1, 6), ws.Cells(fila, 6)).NumberFormat = "0.00"
    ws.Range(ws.Cells(ini + 1, 7), ws.Cells(fila, 7)).NumberFormat = "#,##0"
    With ws.Range(ws.Cells(ini + 1, 8), ws.Cells(fila, 7 + nv))
        .NumberFormat = "0%"
        .FormatConditions.Delete
        .FormatConditions.Add(xlCellValue, xlGreater, TOPE_SATURACION).Interior.Color = RGB(255, 199, 206)
    End With
    With ws.Range(ws.Cells(ini, 1), ws.Cells(ini, UBound(enc)))
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0)
    End With
    ws.Range(ws.Cells(ini, 1), ws.Cells(fila, UBound(enc))).Borders.LineStyle = xlContinuous
    EscribirGrupos = fila
End Function

' Regla: el vehiculo mas pequeno cuya saturacion sea <= TOPE_SATURACION.
' Si no cabe en el mas grande, se indica cuantos del mas grande se necesitan.
Private Function VehiculoSugerido(m3 As Double) As String
    Dim k As Long, nv As Long, mayor As Long
    nv = UBound(vehNombre)
    If m3 <= 0 Then VehiculoSugerido = "Sin cubicaje (revisar ALERTAS)": Exit Function
    mayor = 1
    For k = 1 To nv
        If vehCap(k) > vehCap(mayor) Then mayor = k
    Next k
    Dim mejor As Long: mejor = 0
    For k = 1 To nv
        If m3 / vehCap(k) <= TOPE_SATURACION Then
            If mejor = 0 Then
                mejor = k
            ElseIf vehCap(k) < vehCap(mejor) Then
                mejor = k
            End If
        End If
    Next k
    If mejor > 0 Then
        VehiculoSugerido = vehNombre(mejor) & " (" & Format(m3 / vehCap(mejor), "0%") & ")"
    Else
        Dim cuantos As Long
        cuantos = -Int(-m3 / (vehCap(mayor) * TOPE_SATURACION))
        VehiculoSugerido = cuantos & " x " & vehNombre(mayor) & " - revisar"
    End If
End Function

'==============================================================================
' 2. BORRADOR DE CORREO A LA TRANSPORTADORA (Outlook, no se envia)
'==============================================================================
Public Sub GenerarCorreoTransporte()
    Dim ws As Worksheet, ol As Object, mail As Object, rng As Range, ultima As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_RESUMEN)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Primero ejecute CubicarExportSAP.", vbExclamation: Exit Sub

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set rng = ws.Range("A1", ws.Cells(ultima, ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column))

    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        rng.Copy
        MsgBox "No se pudo abrir Outlook. El resumen quedo copiado: peguelo en el correo.", vbExclamation
        Exit Sub
    End If

    Set mail = ol.CreateItem(0)
    With mail
        .To = CORREO_TRANSPORTADORA
        .CC = CORREO_COPIA
        .Subject = "Solicitud de transporte - " & Format(Date, "dd/mm/yyyy")
        .HTMLBody = "<p>Buenos dias,</p><p>Solicitamos programar el transporte para las siguientes remisiones:</p>" & _
                    RangoAHtml(rng) & "<p>El detalle por remision va en la hoja " & HOJA_DETALLE & ".</p><p>Gracias.</p>"
        .Display   ' Solo muestra el borrador. Revisar y enviar a mano.
    End With
End Sub

'==============================================================================
' 3. MARCAR REMISIONES COMO CUBICADAS (agrega a CRUCE)
'==============================================================================
Public Sub MarcarRemisionesCubicadas()
    Dim wsD As Worksheet, wsC As Worksheet, dCruce As Object, rems As Object
    Dim i As Long, ultima As Long, fila As Long, r As String, k As Variant

    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(HOJA_DETALLE)
    On Error GoTo 0
    If wsD Is Nothing Then MsgBox "No existe la hoja " & HOJA_DETALLE & ".", vbExclamation: Exit Sub

    Set wsC = ThisWorkbook.Worksheets("CRUCE")
    Set dCruce = Diccionario(wsC, 1, 1, 2)
    Set rems = CreateObject("Scripting.Dictionary")

    ultima = wsD.Cells(wsD.Rows.Count, 4).End(xlUp).Row
    For i = 2 To ultima
        r = Clave(wsD.Cells(i, 4).Value)
        If Len(r) > 0 And Not dCruce.Exists(r) Then rems(r) = wsD.Cells(i, 2).Value
    Next i
    If rems.Count = 0 Then MsgBox "Todas las remisiones ya estan en CRUCE.", vbInformation: Exit Sub

    If MsgBox("Se agregaran " & rems.Count & " remisiones a CRUCE como ya cubicadas." & vbCrLf & _
              "Hagalo solo despues de enviar la solicitud de transporte. Continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    fila = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row + 1
    For Each k In rems.Keys
        wsC.Cells(fila, 1).Value = CDbl(k)
        wsC.Cells(fila, 2).Value = rems(k) & " - cubicada " & Format(Now, "dd/mm/yyyy hh:nn")
        fila = fila + 1
    Next k
    MsgBox rems.Count & " remisiones agregadas a CRUCE.", vbInformation
End Sub

'==============================================================================
' Utilidades
'==============================================================================
Private Sub LeerVehiculos()
    Dim ws As Worksheet, i As Long, n As Long
    Set ws = ThisWorkbook.Worksheets("CUBICAJE BODEGA")
    ReDim vehNombre(1 To 4): ReDim vehCap(1 To 4)
    For i = 3 To 6
        If Len(Trim(CStr(ws.Cells(i, 2).Value))) > 0 And Num(ws.Cells(i, 3).Value) > 0 Then
            n = n + 1
            vehNombre(n) = Trim(CStr(ws.Cells(i, 2).Value))
            vehCap(n) = Num(ws.Cells(i, 3).Value)
        End If
    Next i
    If n = 0 Then Err.Raise vbObjectError + 1, , "No se encontraron vehiculos en CUBICAJE BODEGA!B3:C6"
    ReDim Preserve vehNombre(1 To n): ReDim Preserve vehCap(1 To n)
End Sub

Private Function Diccionario(ws As Worksheet, colClave As Long, colValor As Long, filaIni As Long) As Object
    Dim d As Object, v As Variant, i As Long, ultima As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    ultima = ws.Cells(ws.Rows.Count, colClave).End(xlUp).Row
    If ultima >= filaIni Then
        v = ws.Range(ws.Cells(filaIni, 1), ws.Cells(ultima, Application.Max(colClave, colValor))).Value
        For i = 1 To UBound(v, 1)
            k = Clave(v(i, colClave))
            If Len(k) > 0 And Not d.Exists(k) Then d.Add k, v(i, colValor)
        Next i
    End If
    Set Diccionario = d
End Function

Private Function DiccionarioCiudades(ws As Worksheet, colCiudad As Long, colZona As Long, filaIni As Long) As Object
    Dim d As Object, v As Variant, i As Long, ultima As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    ultima = ws.Cells(ws.Rows.Count, colCiudad).End(xlUp).Row
    v = ws.Range(ws.Cells(filaIni, colCiudad), ws.Cells(ultima, colZona)).Value
    For i = 1 To UBound(v, 1)
        k = NormCiudad(CStr(v(i, 1)))
        If Len(k) > 0 And Not d.Exists(k) Then d.Add k, Trim(CStr(v(i, colZona - colCiudad + 1)))
    Next i
    Set DiccionarioCiudades = d
End Function

' Busca la columna por nombre exacto en la fila 1 (primera coincidencia)
Private Function Col(datos As Variant, nombre As String) As Long
    Dim j As Long
    For j = 1 To UBound(datos, 2)
        If Trim(CStr(datos(1, j))) = nombre Then Col = j: Exit Function
    Next j
    Err.Raise vbObjectError + 2, , "El export no tiene la columna '" & nombre & "'"
End Function

' Normaliza codigos: 2915059484, "2915059484" y 2915059484.0 dan la misma clave
Private Function Clave(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Dim s As String: s = Trim(CStr(v))
    If Len(s) = 0 Then Exit Function
    If IsNumeric(s) Then Clave = Format(CDbl(s), "0") Else Clave = UCase(s)
End Function

Private Function Num(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then Num = CDbl(v)
End Function

' "BOGOTA D.C." y "BOGOTA, D.C." quedan iguales; tambien quita tildes
Private Function NormCiudad(s As String) As String
    Dim r As String, a As Variant, b As Variant, i As Long
    r = UCase(Trim(s))
    a = Array("Á", "É", "Í", "Ó", "Ú", "Ü", ",", ".")
    b = Array("A", "E", "I", "O", "U", "U", " ", " ")
    For i = 0 To UBound(a): r = Replace(r, a(i), b(i)): Next i
    Do While InStr(r, "  ") > 0: r = Replace(r, "  ", " "): Loop
    NormCiudad = Trim(r)
End Function

Private Sub Alerta(c As Collection, tipo As String, dato As String, rem_ As String)
    On Error Resume Next   ' evita duplicados de la misma alerta
    c.Add Array(tipo, dato, CDbl(rem_)), tipo & "|" & dato & "|" & rem_
End Sub

Private Function HojaNueva(nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
    Else
        ws.Cells.Clear
        ws.Cells.FormatConditions.Delete
    End If
    Set HojaNueva = ws
End Function

Private Sub FormatoTabla(ws As Worksheet, filas As Long, cols As Long)
    With ws.Range("A1").Resize(1, cols)
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0)
    End With
    ws.Range("A1").Resize(filas, cols).Borders.LineStyle = xlContinuous
    ws.Range("A1").Resize(filas, cols).AutoFilter
    ws.Columns.AutoFit
End Sub

Private Function Ordenadas(d As Object) As Variant
    Dim a As Variant, i As Long, j As Long, t As Variant
    a = d.Keys
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If a(j) < a(i) Then t = a(i): a(i) = a(j): a(j) = t
        Next j
    Next i
    Ordenadas = a
End Function

Private Function MesEnEspanol(f As Date) As String
    MesEnEspanol = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", _
        "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")(Month(f) - 1)
End Function

Private Function RangoAHtml(rng As Range) As String
    Dim h As String, r As Long, c As Long, txt As String
    h = "<table border='1' cellspacing='0' cellpadding='4' style='border-collapse:collapse;font-family:Calibri;font-size:11pt'>"
    For r = 1 To rng.Rows.Count
        If Application.WorksheetFunction.CountA(rng.Rows(r)) = 0 Then
            h = h & "<tr><td colspan='" & rng.Columns.Count & "'>&nbsp;</td></tr>"
        Else
            h = h & "<tr>"
            For c = 1 To rng.Columns.Count
                txt = rng.Cells(r, c).Text
                If rng.Cells(r, c).Font.Bold Then
                    h = h & "<td style='background:#FFFF00;font-weight:bold'>" & txt & "</td>"
                Else
                    h = h & "<td>" & txt & "</td>"
                End If
            Next c
            h = h & "</tr>"
        End If
    Next r
    RangoAHtml = h & "</table>"
End Function
