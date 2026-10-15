Attribute VB_Name = "CubicadorAutomatico"
'==============================================================================
' CUBICADOR AUTOMATICO
'
' Flujo diario:
'   1. CubicarExportSAP    -> lee el export de SAP, toma las entregas que NO
'                             estan en CRUCE, calcula cubicaje, valor y zona.
'                             Crea CUBICAR HOY, RESUMEN TRANSPORTE y ALERTAS.
'   2. EnviarCorteBodega   -> borrador del correo a la bodega ("PRIMER PIRELLI
'                             OCTUBRE 14") con las remisiones nuevas del corte y
'                             sus archivos TXT adjuntos.
'   3. SolicitarVehiculo   -> filtre en CUBICAR HOY las remisiones que van en
'                             el vehiculo y ejecutelo: arma el borrador
'                             "Solicitud Vehiculo No X Doc. Transporte ..." y
'                             ofrece marcarlas en CRUCE. Repetir por vehiculo.
'   MarcarRemisionesCubicadas -> marca en CRUCE las filas visibles (manual).
'
' Ningun correo se envia solo: todos quedan como borrador para revisar.
'
' Hojas que usa del cubicador:
'   MATERIAL         A = IP, C = volumen teorico m3/und
'   VALO MATERIAL    C = IP, D = valor sin IVA (desde la fila 4)
'   llanta_zona      J = ciudad, K = zona
'   CRUCE            A = remisiones ya cubicadas (se agregan en el paso 3)
'   CUBICAJE BODEGA  A3:C6 = etiqueta (ej. "42 m3"), tipo de vehiculo, m3 maximo
'
' Hojas que crea: CUBICAR HOY, RESUMEN TRANSPORTE, ALERTAS, HISTORIAL BODEGA
'==============================================================================
Option Explicit

' --- Configuracion: destinatarios ------------------------------------------
' Separe con punto y coma. Puede usar correos o nombres como aparecen en Outlook.
Private Const PARA_TRANSPORTE As String = ""
Private Const CC_TRANSPORTE As String = ""
Private Const PARA_BODEGA As String = ""
Private Const CC_BODEGA As String = ""

' --- Configuracion: otros ----------------------------------------------------
Private Const CARPETA_TXT As String = ""            ' carpeta donde SAP deja los TXT de remision; vacio = se pregunta
Private Const TOPE_SATURACION As Double = 1#        ' solo para el vehiculo sugerido del resumen
Private Const CENTRO_CUBICADOR As String = "CDLI"

Private Const HOJA_DETALLE As String = "CUBICAR HOY"
Private Const HOJA_RESUMEN As String = "RESUMEN TRANSPORTE"
Private Const HOJA_ALERTAS As String = "ALERTAS"
Private Const HOJA_HIST As String = "HISTORIAL BODEGA"
Private Const SIN_ZONA As String = "SIN ZONA"

' Columnas de CUBICAR HOY
Private Const C_FECHA As Long = 1, C_CORTE As Long = 2, C_CENTRO As Long = 3, C_REM As Long = 4
Private Const C_CLIENTE As Long = 5, C_CIUDAD As Long = 6, C_ZONA As Long = 7, C_COD As Long = 8
Private Const C_DESC As Long = 9, C_CANT As Long = 10, C_PESO As Long = 11, C_CODSAP As Long = 12
Private Const C_DIR As Long = 13, C_CUBUND As Long = 14, C_CUB As Long = 15, C_EMP As Long = 16
Private Const C_VALOR As Long = 17, C_NUEVO As Long = 18

' Vehiculos (se leen de CUBICAJE BODEGA!A3:C6)
Private vehNombre() As String
Private vehCap() As Double
Private vehEtiqueta() As String

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
    Dim dHist As Object: Set dHist = Diccionario(HojaHistorial(), 1, 2, 2)
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
    Const NC As Long = 18
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
                out(n, 2) = IIf(dHist.Exists(rem_), dHist(rem_), corte)
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
                out(n, 18) = IIf(dHist.Exists(rem_), "NO", "SI")
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
        "Cubicaje/und", "Cubicaje Total", "EMPRESA", "Valor total sin IVA", "Nuevo en corte")
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
' 2. CORREO A LA BODEGA: remisiones nuevas del corte + TXT adjuntos
'==============================================================================
Public Sub EnviarCorteBodega()
    Dim ws As Worksheet, filas As Collection, i As Long, ultima As Long
    Set ws = HojaDetalle(): If ws Is Nothing Then Exit Sub

    ' Remisiones nuevas (no enviadas antes a bodega), agrupadas por empresa
    Dim porEmpresa As Object: Set porEmpresa = CreateObject("Scripting.Dictionary")
    Dim corte As String, emp As String
    ultima = ws.Cells(ws.Rows.Count, C_REM).End(xlUp).Row
    For i = 2 To ultima
        If ws.Cells(i, C_NUEVO).Value = "SI" Then
            emp = CStr(ws.Cells(i, C_EMP).Value)
            If Not porEmpresa.Exists(emp) Then porEmpresa.Add emp, New Collection
            porEmpresa(emp).Add i
            corte = CStr(ws.Cells(i, C_CORTE).Value)
        End If
    Next i
    If porEmpresa.Count = 0 Then MsgBox "No hay remisiones nuevas para enviar a bodega.", vbInformation: Exit Sub

    ' Archivos TXT de remision
    Dim carpeta As String, txt As Object, sinTxt As String
    carpeta = CarpetaTxt()
    If Len(carpeta) > 0 Then Set txt = ArchivosTxtPorRemision(carpeta) Else Set txt = CreateObject("Scripting.Dictionary")

    Dim k As Variant, adj As Collection, vistas As Object, r As String, total As Long
    For Each k In porEmpresa.Keys
        Set filas = porEmpresa(k)
        Set adj = New Collection
        Set vistas = CreateObject("Scripting.Dictionary")
        For i = 1 To filas.Count
            r = Clave(ws.Cells(filas(i), C_REM).Value)
            If Not vistas.Exists(r) Then
                vistas.Add r, True
                If txt.Exists(r) Then adj.Add txt(r) Else sinTxt = sinTxt & vbCrLf & r
            End If
        Next i
        total = total + vistas.Count

        CrearCorreo PARA_BODEGA, CC_BODEGA, AsuntoCorte(corte, CStr(k)), _
            "<p>Buen d&iacute;a</p><p>Adjunto se encuentra el " & TextoCorte(corte) & ":</p>" & _
            TablaHtml(ws, filas, False), adj
    Next k

    If Len(sinTxt) > 0 Then MsgBox "Remisiones sin archivo TXT en la carpeta (adjuntar a mano):" & sinTxt, vbExclamation

    If MsgBox("Se crearon los borradores para bodega (" & total & " remisiones)." & vbCrLf & _
              "Despues de enviarlos, registrarlas como enviadas a bodega?" & vbCrLf & _
              "(Si responde No, volveran a salir en el proximo corte.)", vbYesNo + vbQuestion) = vbYes Then
        RegistrarEnviadasBodega ws, porEmpresa
    End If
End Sub

'==============================================================================
' 3. SOLICITUD DE VEHICULO A LA TRANSPORTADORA (filas visibles de CUBICAR HOY)
'==============================================================================
Public Sub SolicitarVehiculo()
    Dim ws As Worksheet, filas As Collection, i As Long, ultima As Long
    Set ws = HojaDetalle(): If ws Is Nothing Then Exit Sub
    LeerVehiculos

    Set filas = New Collection
    ultima = ws.Cells(ws.Rows.Count, C_REM).End(xlUp).Row
    Dim unid As Double, kg As Double, m3 As Double, destinos As String, ciudad As String
    Dim rems As Object: Set rems = CreateObject("Scripting.Dictionary")
    For i = 2 To ultima
        If Not ws.Rows(i).Hidden And Len(Clave(ws.Cells(i, C_REM).Value)) > 0 Then
            filas.Add i
            rems(Clave(ws.Cells(i, C_REM).Value)) = True
            unid = unid + Num(ws.Cells(i, C_CANT).Value)
            kg = kg + Num(ws.Cells(i, C_PESO).Value)
            m3 = m3 + Num(ws.Cells(i, C_CUB).Value)
            ciudad = UCase(Trim(CStr(ws.Cells(i, C_CIUDAD).Value)))
            If InStr(1, " / " & destinos & " / ", " / " & ciudad & " / ") = 0 Then
                destinos = destinos & IIf(Len(destinos) > 0, " / ", "") & ciudad
            End If
        End If
    Next i
    If filas.Count = 0 Then MsgBox "No hay filas visibles en " & HOJA_DETALLE & ".", vbExclamation: Exit Sub

    If filas.Count = ultima - 1 Then
        If MsgBox("No hay filtro: se incluirian TODAS las " & rems.Count & " remisiones en un solo vehiculo." & vbCrLf & _
                  "Filtre primero las remisiones del vehiculo. Continuar de todas formas?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    ' Elegir vehiculo mostrando la saturacion de cada uno
    Dim lista As String, k As Long, opc As String
    For k = 1 To UBound(vehNombre)
        lista = lista & k & " = " & vehEtiqueta(k) & "  ->  " & Format(m3 / vehCap(k), "0%") & vbCrLf
    Next k
    opc = InputBox("Remisiones: " & rems.Count & "   Unidades: " & Format(unid, "#,##0") & _
                   "   Peso: " & Format(kg, "#,##0") & " kg   Cubicaje: " & Format(m3, "0.00") & " m3" & vbCrLf & vbCrLf & _
                   "Saturacion por vehiculo:" & vbCrLf & lista & vbCrLf & "Numero del tipo de vehiculo:", "Tipo de vehiculo", _
                   CStr(IndiceSugerido(m3)))
    If Not IsNumeric(opc) Then Exit Sub
    k = CLng(opc): If k < 1 Or k > UBound(vehNombre) Then Exit Sub

    Dim numVeh As String, docTransp As String, notaDespacho As String, cliente As String
    numVeh = InputBox("Numero del vehiculo del dia:", "Vehiculo No", "1"): If Len(numVeh) = 0 Then Exit Sub
    docTransp = InputBox("Doc. Transporte:", "Doc. Transporte", Clave(ws.Cells(filas(1), C_REM).Value)): If Len(docTransp) = 0 Then Exit Sub
    notaDespacho = InputBox("Nota inicial del correo:", "Nota", "Vehículo para despacho mañana.")
    cliente = InputBox("Texto adicional para la nota (opcional, ej. '1 CLIENTE'):", "Nota", "")

    Dim cuerpo As String
    cuerpo = "<p>Buen d&iacute;a</p>" & _
        "<p><b>NOTA: " & Html(notaDespacho) & "</b></p>" & _
        "<table style='font-family:Calibri;font-size:11pt'>" & _
        FilaDato("Tipo Veh&iacute;culo", Html(vehEtiqueta(k))) & _
        FilaDato("Destino", Html(destinos)) & _
        FilaDato("Doc. Transporte", Html(docTransp)) & _
        FilaDato("Observaciones:", "Unidades cargadas " & Format(unid, "#,##0")) & _
        FilaDato("Nota:", "El veh&iacute;culo se encuentra " & Format(m3 / vehCap(k), "0%") & " de su capacidad " & Html(cliente)) & _
        FilaDato("", "Despachar seg&uacute;n TXT al menos que tenga instrucci&oacute;n PESO " & Format(kg, "0")) & _
        "</table><br>" & TablaHtml(ws, filas, True)

    CrearCorreo PARA_TRANSPORTE, CC_TRANSPORTE, _
        "Solicitud Vehículo No " & numVeh & " Doc. Transporte " & docTransp & " " & Format(Date, "dd/mm/yyyy"), cuerpo, Nothing

    If MsgBox("Borrador creado." & vbCrLf & "Despues de enviarlo, marcar estas " & rems.Count & _
              " remisiones en CRUCE como cubicadas?", vbYesNo + vbQuestion) = vbYes Then
        MarcarEnCruce rems, "Vehiculo " & numVeh & " - Doc " & docTransp
        On Error Resume Next
        If ws.FilterMode Then ws.ShowAllData
        On Error GoTo 0
    End If
End Sub

'==============================================================================
' Marcar manualmente en CRUCE las filas visibles de CUBICAR HOY
'==============================================================================
Public Sub MarcarRemisionesCubicadas()
    Dim ws As Worksheet, i As Long, ultima As Long, rems As Object, r As String
    Set ws = HojaDetalle(): If ws Is Nothing Then Exit Sub
    Set rems = CreateObject("Scripting.Dictionary")
    ultima = ws.Cells(ws.Rows.Count, C_REM).End(xlUp).Row
    For i = 2 To ultima
        r = Clave(ws.Cells(i, C_REM).Value)
        If Not ws.Rows(i).Hidden And Len(r) > 0 Then rems(r) = True
    Next i
    If rems.Count = 0 Then Exit Sub
    If MsgBox("Marcar en CRUCE las " & rems.Count & " remisiones visibles?", vbYesNo + vbQuestion) = vbYes Then
        MarcarEnCruce rems, "Marcada manual"
    End If
End Sub

Private Sub MarcarEnCruce(rems As Object, obs As String)
    Dim wsC As Worksheet, dCruce As Object, fila As Long, k As Variant, n As Long
    Set wsC = ThisWorkbook.Worksheets("CRUCE")
    Set dCruce = Diccionario(wsC, 1, 1, 2)
    fila = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row + 1
    For Each k In rems.Keys
        If Not dCruce.Exists(CStr(k)) Then
            wsC.Cells(fila, 1).Value = CDbl(k)
            wsC.Cells(fila, 2).Value = obs & " - " & Format(Now, "dd/mm/yyyy hh:nn")
            fila = fila + 1: n = n + 1
        End If
    Next k
    MsgBox n & " remisiones agregadas a CRUCE.", vbInformation
End Sub

Private Sub RegistrarEnviadasBodega(ws As Worksheet, porEmpresa As Object)
    Dim h As Worksheet, fila As Long, k As Variant, i As Long, r As String, vistas As Object
    Set h = HojaHistorial()
    Set vistas = CreateObject("Scripting.Dictionary")
    fila = h.Cells(h.Rows.Count, 1).End(xlUp).Row + 1
    For Each k In porEmpresa.Keys
        For i = 1 To porEmpresa(k).Count
            r = Clave(ws.Cells(porEmpresa(k)(i), C_REM).Value)
            If Not vistas.Exists(r) Then
                vistas.Add r, True
                h.Cells(fila, 1).Value = CDbl(r)
                h.Cells(fila, 2).Value = ws.Cells(porEmpresa(k)(i), C_CORTE).Value
                h.Cells(fila, 3).Value = Now
                fila = fila + 1
            End If
            ws.Cells(porEmpresa(k)(i), C_NUEVO).Value = "NO"
        Next i
    Next k
End Sub

'==============================================================================
' Utilidades de correo
'==============================================================================
Private Sub CrearCorreo(para As String, cc As String, asunto As String, cuerpoHtml As String, adjuntos As Collection)
    Dim ol As Object, mail As Object, i As Long
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then MsgBox "No se pudo abrir Outlook.", vbCritical: Exit Sub

    Set mail = ol.CreateItem(0)
    mail.Display                       ' primero se muestra para conservar la firma de Outlook
    mail.To = para
    mail.CC = cc
    mail.Subject = asunto
    mail.HTMLBody = "<div style='font-family:Calibri;font-size:11pt'>" & cuerpoHtml & "</div>" & mail.HTMLBody
    If Not adjuntos Is Nothing Then
        For i = 1 To adjuntos.Count
            mail.Attachments.Add adjuntos(i)
        Next i
    End If
End Sub

' Tabla del correo con las mismas columnas que se envian hoy
Private Function TablaHtml(ws As Worksheet, filas As Collection, conValor As Boolean) As String
    Dim cols As Variant, titulos As Variant, h As String, i As Long, j As Long, v As Variant, txt As String
    If conValor Then
        cols = Array(C_FECHA, C_CENTRO, C_REM, C_CLIENTE, C_CIUDAD, C_COD, C_DESC, C_CANT, C_PESO, C_CODSAP, C_DIR, C_EMP, C_VALOR)
        titulos = Array("Fecha Remisi&oacute;n", "CENTRO", "Remisi&oacute;n", "Cliente", "Ciudad", "C&oacute;digo", "Descripci&oacute;n", _
                        "Cantidad", "Peso", "Codigo SAP", "Direcci&oacute;n", "EMPRESA", "VALOR TOTAL SIN IVA")
    Else
        cols = Array(C_FECHA, C_CENTRO, C_REM, C_CLIENTE, C_CIUDAD, C_COD, C_DESC, C_CANT, C_PESO, C_CODSAP, C_DIR, C_EMP)
        titulos = Array("Fecha Remisi&oacute;n", "CENTRO", "Remisi&oacute;n", "Cliente", "Ciudad", "C&oacute;digo", "Descripci&oacute;n", _
                        "Cantidad", "Peso", "Codigo SAP", "Direcci&oacute;n", "EMPRESA")
    End If
    h = "<table border='1' cellspacing='0' cellpadding='3' style='border-collapse:collapse;font-family:Calibri;font-size:10pt'><tr>"
    For j = 0 To UBound(titulos)
        h = h & "<th style='background:#FFFF00'>" & titulos(j) & "</th>"
    Next j
    h = h & "</tr>"
    For i = 1 To filas.Count
        h = h & "<tr>"
        For j = 0 To UBound(cols)
            v = ws.Cells(filas(i), cols(j)).Value
            Select Case cols(j)
                Case C_FECHA: If IsDate(v) Then txt = Format(v, "d/mm/yyyy") Else txt = CStr(v)
                Case C_REM, C_COD, C_CODSAP: txt = Clave(v)
                Case C_CANT: txt = Format(Num(v), "#,##0")
                Case C_PESO: txt = Format(Num(v), "#,##0.000")
                Case C_VALOR: txt = "$ " & Format(Num(v), "#,##0.00")
                Case Else: txt = Html(CStr(v))
            End Select
            h = h & "<td" & IIf(cols(j) = C_CANT Or cols(j) = C_PESO Or cols(j) = C_VALOR, " align='right'", "") & ">" & txt & "</td>"
        Next j
        h = h & "</tr>"
    Next i
    TablaHtml = h & "</table>"
End Function

Private Function FilaDato(etiqueta As String, valor As String) As String
    FilaDato = "<tr><td style='padding-right:40px'><b>" & etiqueta & "</b></td><td>" & valor & "</td></tr>"
End Function

Private Function Html(s As String) As String
    Html = Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function

' "1 Corte Octubre 14" + "PIRELLI TYRE" -> "PRIMER PIRELLI OCTUBRE 14"
Private Function AsuntoCorte(corte As String, empresa As String) As String
    Dim p As Variant, marca As String
    marca = UCase(Split(Trim(empresa) & " ", " ")(0))
    p = Split(Application.WorksheetFunction.Trim(corte), " ")
    If UBound(p) >= 3 Then
        AsuntoCorte = Ordinal(CStr(p(0))) & " " & marca & " " & UCase(p(2)) & " " & p(3)
    Else
        AsuntoCorte = UCase(corte) & " " & marca
    End If
End Function

' "1 Corte Octubre 14" -> "primer corte del 14 de octubre"
Private Function TextoCorte(corte As String) As String
    Dim p As Variant
    p = Split(Application.WorksheetFunction.Trim(corte), " ")
    If UBound(p) >= 3 Then
        TextoCorte = LCase(Ordinal(CStr(p(0)))) & " corte del " & p(3) & " de " & LCase(p(2))
    Else
        TextoCorte = "corte " & corte
    End If
End Function

Private Function Ordinal(n As String) As String
    Select Case Trim(n)
        Case "1": Ordinal = "PRIMER"
        Case "2": Ordinal = "SEGUNDO"
        Case "3": Ordinal = "TERCER"
        Case "4": Ordinal = "CUARTO"
        Case "5": Ordinal = "QUINTO"
        Case Else: Ordinal = n
    End Select
End Function

' Lee los TXT de la carpeta y los indexa por el numero que aparece tras "Remision:"
' (si hay varios para la misma remision se toma el mas reciente)
Private Function ArchivosTxtPorRemision(carpeta As String) As Object
    Dim d As Object, fechas As Object, re As Object, f As String, ruta As String, s As String, ff As Integer, m As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set fechas = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "Remisi.n:\s*(\d{8,12})"
    re.IgnoreCase = True
    If Right(carpeta, 1) <> "\" Then carpeta = carpeta & "\"
    f = Dir(carpeta & "*.txt")
    Do While Len(f) > 0
        ruta = carpeta & f
        ff = FreeFile
        Open ruta For Binary Access Read As #ff
        s = Space$(Application.Min(LOF(ff), 4000))
        Get #ff, , s
        Close #ff
        If re.Test(s) Then
            Set m = re.Execute(s)(0)
            If Not d.Exists(m.SubMatches(0)) Then
                d.Add m.SubMatches(0), ruta: fechas.Add m.SubMatches(0), FileDateTime(ruta)
            ElseIf FileDateTime(ruta) > fechas(m.SubMatches(0)) Then
                d(m.SubMatches(0)) = ruta: fechas(m.SubMatches(0)) = FileDateTime(ruta)
            End If
        End If
        f = Dir()
    Loop
    Set ArchivosTxtPorRemision = d
End Function

' Carpeta de los TXT: la constante, o la ultima usada (HISTORIAL BODEGA!F1), o se pregunta
Private Function CarpetaTxt() As String
    Dim h As Worksheet, c As String
    Set h = HojaHistorial()
    c = CARPETA_TXT
    If Len(c) = 0 Then c = CStr(h.Range("F1").Value)
    If Len(c) > 0 Then
        If Len(Dir(c, vbDirectory)) = 0 Then c = ""
    End If
    If Len(c) = 0 Then
        With Application.FileDialog(4)   ' msoFileDialogFolderPicker
            .Title = "Carpeta donde estan los TXT de las remisiones"
            If .Show = -1 Then c = .SelectedItems(1)
        End With
    End If
    If Len(c) > 0 Then h.Range("E1").Value = "Carpeta TXT": h.Range("F1").Value = c
    CarpetaTxt = c
End Function

Private Function HojaHistorial() As Worksheet
    Dim h As Worksheet
    On Error Resume Next
    Set h = ThisWorkbook.Worksheets(HOJA_HIST)
    On Error GoTo 0
    If h Is Nothing Then
        Set h = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        h.Name = HOJA_HIST
        h.Range("A1:C1").Value = Array("Remision", "Fecha Corte", "Enviada a bodega")
        h.Range("A1:C1").Font.Bold = True
        ' Primera vez: las remisiones que ya estan en CUBICAJE BODEGA ya se reportaron a bodega
        Dim cb As Worksheet, v As Variant, i As Long, fila As Long, ultima As Long
        Set cb = ThisWorkbook.Worksheets("CUBICAJE BODEGA")
        ultima = cb.Cells(cb.Rows.Count, 5).End(xlUp).Row
        fila = 2
        If ultima >= 10 Then
            v = cb.Range("A10:E" & ultima).Value
            For i = 1 To UBound(v, 1)
                If Len(Clave(v(i, 5))) > 0 Then
                    h.Cells(fila, 1).Value = CDbl(Clave(v(i, 5)))
                    h.Cells(fila, 2).Value = v(i, 3)
                    h.Cells(fila, 3).Value = "Ya en cubicador"
                    fila = fila + 1
                End If
            Next i
        End If
    End If
    Set HojaHistorial = h
End Function

Private Function HojaDetalle() As Worksheet
    On Error Resume Next
    Set HojaDetalle = ThisWorkbook.Worksheets(HOJA_DETALLE)
    On Error GoTo 0
    If HojaDetalle Is Nothing Then MsgBox "Primero ejecute CubicarExportSAP.", vbExclamation
End Function

Private Function IndiceSugerido(m3 As Double) As Long
    Dim k As Long, mejor As Long
    For k = 1 To UBound(vehNombre)
        If m3 / vehCap(k) <= TOPE_SATURACION Then
            If mejor = 0 Then
                mejor = k
            ElseIf vehCap(k) < vehCap(mejor) Then
                mejor = k
            End If
        End If
    Next k
    If mejor = 0 Then
        For k = 1 To UBound(vehNombre)
            If mejor = 0 Then
                mejor = k
            ElseIf vehCap(k) > vehCap(mejor) Then
                mejor = k
            End If
        Next k
    End If
    IndiceSugerido = mejor
End Function

'==============================================================================
' Utilidades
'==============================================================================
Private Sub LeerVehiculos()
    Dim ws As Worksheet, i As Long, n As Long, etq As String
    Set ws = ThisWorkbook.Worksheets("CUBICAJE BODEGA")
    ReDim vehNombre(1 To 4): ReDim vehCap(1 To 4): ReDim vehEtiqueta(1 To 4)
    For i = 3 To 6
        If Len(Trim(CStr(ws.Cells(i, 2).Value))) > 0 And Num(ws.Cells(i, 3).Value) > 0 Then
            n = n + 1
            vehNombre(n) = Trim(CStr(ws.Cells(i, 2).Value))
            vehCap(n) = Num(ws.Cells(i, 3).Value)
            etq = Trim(Replace(LCase(CStr(ws.Cells(i, 1).Value)), "m3", ""))   ' "42 m3" -> "42"
            vehEtiqueta(n) = vehNombre(n) & IIf(Len(etq) > 0, " " & etq, "")
        End If
    Next i
    If n = 0 Then Err.Raise vbObjectError + 1, , "No se encontraron vehiculos en CUBICAJE BODEGA!B3:C6"
    ReDim Preserve vehNombre(1 To n): ReDim Preserve vehCap(1 To n): ReDim Preserve vehEtiqueta(1 To n)
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
