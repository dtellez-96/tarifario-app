Attribute VB_Name = "Module1"
Option Explicit

Sub Cotizar()

    '========================================================
    ' DECLARACIÓN DE VARIABLES
    '========================================================

    Dim wsC As Worksheet
    Dim wsTar As Worksheet
    Dim tblTar As ListObject
    Dim filaTar As ListRow

    Dim ultimaFila As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim origen As String
    Dim destino As String

    Dim cantidad20 As Double
    Dim cantidad40ST As Double
    Dim cantidad40HC As Double
    Dim cantidadContenedores As Double
    Dim cantidadBL As Double
    Dim valorFOB As Double

    Dim agente As String
    Dim naviera As String

    Dim flete20 As Double
    Dim flete40ST As Double
    Dim flete40HC As Double

    Dim fleteTotal As Double
    Dim origenBL As Double
    Dim gastosOrigen As Double
    Dim recargosMaritimos As Double
    Dim gastosDestino As Double
    Dim costoTotal As Double
    Dim costoMedio As Double

    Dim transito As Variant
    Dim diasLibres As Variant

    Dim agentesUnicos As Object
    Dim informacionAgentes As Object

    Dim datosAgente As Variant
    Dim listaResultados() As Variant
    Dim temporal As Variant

    Dim cantidadAgentes As Long
    Dim cantidadMostrar As Long
    Dim colSalida As Long

    Dim clave As Variant

    On Error GoTo ErrorCotizar

    Set wsC = ThisWorkbook.Worksheets("Cotizador")

    Set wsTar = ThisWorkbook.Worksheets("TARIFARIOS_AGENTES")
    Set tblTar = wsTar.ListObjects("TARIFARIOS_AGENTES")

    Set agentesUnicos = CreateObject("Scripting.Dictionary")
    Set informacionAgentes = CreateObject("Scripting.Dictionary")


    '========================================================
    ' LEER PARÁMETROS DEL COTIZADOR
    '========================================================

    origen = UCase(Trim(wsC.Range("B6").Value))
    destino = UCase(Trim(wsC.Range("B7").Value))

    cantidad20 = ValorNumerico(wsC.Range("B8").Value)
    cantidad40ST = ValorNumerico(wsC.Range("B9").Value)
    cantidad40HC = ValorNumerico(wsC.Range("B10").Value)

    cantidadContenedores = _
        cantidad20 + cantidad40ST + cantidad40HC

    cantidadBL = ValorNumerico(wsC.Range("B11").Value)
    valorFOB = ValorNumerico(wsC.Range("B12").Value)


    '========================================================
    ' VALIDACIONES
    '========================================================

    If origen = "" Then

        MsgBox "Seleccione el Puerto de Origen.", _
               vbExclamation, "Falta origen"

        wsC.Range("B6").Select
        Exit Sub

    End If


    If destino = "" Then

        MsgBox "Seleccione el Puerto de Destino.", _
               vbExclamation, "Falta destino"

        wsC.Range("B7").Select
        Exit Sub

    End If


    If cantidadContenedores <= 0 Then

        MsgBox "Ingrese por lo menos un contenedor." _
               & vbCrLf & vbCrLf _
               & "Puede ingresar cantidad en:" _
               & vbCrLf _
               & "20 ST" _
               & vbCrLf _
               & "40 ST" _
               & vbCrLf _
               & "40 HC", _
               vbExclamation, "Falta cantidad"

        Exit Sub

    End If


    If cantidadBL <= 0 Then
        cantidadBL = 1
        wsC.Range("B11").Value = 1
    End If


    '========================================================
    ' PREPARAR EXCEL
    '========================================================

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Buscando y comparando tarifas..."

    Application.CalculateFull


    '========================================================
    ' LIMPIAR COMPARATIVO ANTERIOR
    '========================================================

    wsC.Range("B15:F31").ClearContents

    'Si no hay opción, quedará este texto
    wsC.Range("B15:F15").Value = "SIN OPCIÓN"

    'Limpiar colores previos
    wsC.Range("B15:F31").Interior.Pattern = xlNone

    'Restaurar encabezado azul
    With wsC.Range("B15:F15")

        .Interior.Color = RGB(91, 155, 213)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True

    End With


    'Restaurar fila de costo total
    With wsC.Range("B25:F25")

        .Interior.Color = RGB(255, 242, 204)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True

    End With


    '========================================================
' BUSCAR TODAS LAS OPCIONES PARA LA RUTA
' DIRECTAMENTE DESDE POWER QUERY
'========================================================

For Each filaTar In tblTar.ListRows

    'Solo considerar la ruta seleccionada
    'y excluir tarifas vencidas

    If UCase(Trim(filaTar.Range.Cells(1, _
    tblTar.ListColumns("POL").Index).Value)) = origen _
And UCase(Trim(filaTar.Range.Cells(1, _
    tblTar.ListColumns("POD").Index).Value)) = destino Then

        '--------------------------------------------
        ' AGENTE
        '--------------------------------------------

        agente = UCase(Trim( _
            filaTar.Range.Cells(1, _
            tblTar.ListColumns("AGENTE").Index).Value))


        '--------------------------------------------
        ' NAVIERA
        '--------------------------------------------

        naviera = UCase(Trim( _
            filaTar.Range.Cells(1, _
            tblTar.ListColumns("NAVIERA").Index).Value))


        If agente <> "" Then


            '--------------------------------------------
            ' TARIFAS UNITARIAS
            '--------------------------------------------

            flete20 = ValorNumerico( _
                filaTar.Range.Cells(1, _
                tblTar.ListColumns("FLETE_20ST").Index).Value)


            flete40ST = ValorNumerico( _
                filaTar.Range.Cells(1, _
                tblTar.ListColumns("FLETE_40ST").Index).Value)


            flete40HC = ValorNumerico( _
                filaTar.Range.Cells(1, _
                tblTar.ListColumns("FLETE_40HC").Index).Value)


                '--------------------------------------------
                ' FLETE MARÍTIMO TOTAL
                '--------------------------------------------

                    fleteTotal = _
    (flete20 * cantidad20) + _
    (flete40ST * cantidad40ST) + _
    (flete40HC * cantidad40HC)

                
                End If


                '--------------------------------------------
                ' ORIGEN / BL
                '--------------------------------------------

               origenBL = ValorNumerico( _
    filaTar.Range.Cells(1, _
    tblTar.ListColumns("DOCUMENTACION_BL").Index).Value) _
    * cantidadBL
    
'--------------------------------------------
' TRÁNSITO ESTIMADO
'--------------------------------------------

transito = filaTar.Range.Cells(1, _
    tblTar.ListColumns("TT").Index).Value


'--------------------------------------------
' DÍAS LIBRES
'--------------------------------------------

diasLibres = filaTar.Range.Cells(1, _
    tblTar.ListColumns("DIAS_LIBRES").Index).Value

'--------------------------------------------
' GASTOS EN ORIGEN
' MANUAL - NO MODIFICAR DESDE VBA
'--------------------------------------------

'--------------------------------------------
' LLENADO DEL CONTENEDOR
' MANUAL - NO MODIFICAR DESDE VBA
'--------------------------------------------
                '--------------------------------------------
                ' GASTOS EN DESTINO
                '--------------------------------------------

                gastosDestino = GastoDestino( _
                    agente, _
                    destino, _
                    naviera, _
                    cantidadContenedores, _
                    cantidadBL)


                '--------------------------------------------
                ' COSTO TOTAL
                '--------------------------------------------

                costoTotal = _
                    fleteTotal _
                    + origenBL _
                    + gastosOrigen _
                    + recargosMaritimos _
                    + gastosDestino


                If cantidadContenedores > 0 Then

                    costoMedio = _
                        costoTotal / cantidadContenedores

                Else

                    costoMedio = 0

                End If


                transito = filaTar.Range.Cells(1, _
                tblTar.ListColumns("TT").Index).Value

               diasLibres = filaTar.Range.Cells(1, _
                tblTar.ListColumns("DIAS_LIBRES").Index).Value

                '--------------------------------------------
                ' GUARDAR MEJOR OPCIÓN DE CADA AGENTE
                '--------------------------------------------

                datosAgente = Array( _
                    agente, _
                    naviera, _
                    flete20, _
                    flete40ST, _
                    flete40HC, _
                    origenBL, _
                    gastosOrigen, _
                    fleteTotal, _
                    recargosMaritimos, _
                    gastosDestino, _
                    costoTotal, _
                    costoMedio, _
                    transito, _
                    diasLibres, _
                    0)


                If Not agentesUnicos.Exists(agente) Then

                    agentesUnicos.Add agente, costoTotal
                    informacionAgentes.Add agente, datosAgente

                Else

                    'Solo guardar la opción más económica
                    'del mismo agente
                    If costoTotal < agentesUnicos(agente) Then

                        agentesUnicos(agente) = costoTotal
                        informacionAgentes(agente) = datosAgente

                    End If

                End If


            End If


   Next filaTar


    '========================================================
    ' VALIDAR RESULTADOS
    '========================================================

    cantidadAgentes = agentesUnicos.Count


    If cantidadAgentes = 0 Then

        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.StatusBar = False

        MsgBox "No se encontraron tarifas para:" _
               & vbCrLf & vbCrLf _
               & "Origen: " & origen _
               & vbCrLf _
               & "Destino: " & destino _
               & vbCrLf & vbCrLf _
               & "Revise que la ruta esté cargada en la hoja Tarifas.", _
               vbInformation, "Ruta sin tarifa"

        Exit Sub

    End If


    '========================================================
    ' PASAR DICCIONARIO A UN ARREGLO
    '========================================================

    ReDim listaResultados(1 To cantidadAgentes, 1 To 15)

    i = 0


    For Each clave In informacionAgentes.Keys

        i = i + 1

        datosAgente = informacionAgentes(clave)

        For j = 0 To 14

            listaResultados(i, j + 1) = _
                datosAgente(j)

        Next j

    Next clave


    '========================================================
    ' ORDENAR AGENTES DE MENOR A MAYOR COSTO
    '
    ' COLUMNA 11 DEL ARREGLO = COSTO TOTAL
    '========================================================

    For i = 1 To cantidadAgentes - 1

        For j = i + 1 To cantidadAgentes

            If ValorNumerico(listaResultados(j, 11)) _
            < ValorNumerico(listaResultados(i, 11)) Then

                For k = 1 To 15

                    temporal = listaResultados(i, k)

                    listaResultados(i, k) = _
                        listaResultados(j, k)

                    listaResultados(j, k) = temporal

                Next k

            End If

        Next j

    Next i


    '========================================================
    ' DEFINIR CUÁNTOS AGENTES MOSTRAR
    '========================================================

    If cantidadAgentes > 5 Then

        cantidadMostrar = 5

    Else

        cantidadMostrar = cantidadAgentes

    End If


    '========================================================
    ' LLENAR TABLA HORIZONTAL
    '========================================================

    For i = 1 To cantidadMostrar

        'B, C, D, E y F
        colSalida = 1 + i


        'AGENTE / FORWARDER
        wsC.Cells(15, colSalida).Value = _
            listaResultados(i, 1)


        'NAVIERA
        wsC.Cells(16, colSalida).Value = _
            listaResultados(i, 2)


'=================================================
' MOSTRAR ÚNICAMENTE LOS CONTENEDORES SOLICITADOS
'=================================================

'20 ST
If cantidad20 > 0 Then

    wsC.Cells(17, colSalida).Value = _
        listaResultados(i, 3) * cantidad20

    wsC.Cells(17, colSalida).NumberFormat = "$#,##0.00"

Else

    wsC.Cells(17, colSalida).Value = "-"

End If


'40 ST
If cantidad40ST > 0 Then

    wsC.Cells(18, colSalida).Value = _
        listaResultados(i, 4) * cantidad40ST

    wsC.Cells(18, colSalida).NumberFormat = "$#,##0.00"

Else

    wsC.Cells(18, colSalida).Value = "-"

End If


'40 HC
If cantidad40HC > 0 Then

    wsC.Cells(19, colSalida).Value = _
        listaResultados(i, 5) * cantidad40HC

    wsC.Cells(19, colSalida).NumberFormat = "$#,##0.00"

Else

    wsC.Cells(19, colSalida).Value = "-"

End If
        'ORIGEN / BL
        wsC.Cells(20, colSalida).Value = _
            listaResultados(i, 6)


        'GASTOS EN ORIGEN
        wsC.Cells(21, colSalida).Value = _
            listaResultados(i, 7)


        'FLETE MARÍTIMO TOTAL
        wsC.Cells(22, colSalida).Value = _
            listaResultados(i, 8)


        'RECARGOS MARÍTIMOS
        wsC.Cells(23, colSalida).Value = _
            listaResultados(i, 9)


        'GASTOS EN DESTINO
        wsC.Cells(24, colSalida).Value = _
            listaResultados(i, 10)


        'COSTO TOTAL DINÁMICO
wsC.Cells(25, colSalida).Formula = _
    "=SUM(" & _
    wsC.Cells(20, colSalida).Address(False, False) & ":" & _
    wsC.Cells(24, colSalida).Address(False, False) & ")"

        'COSTO MEDIO POR CONTENEDOR
        wsC.Cells(26, colSalida).Value = _
            listaResultados(i, 12)


        'TRÁNSITO ESTIMADO
        wsC.Cells(27, colSalida).Value = _
            listaResultados(i, 13)


        'DÍAS LIBRES
        wsC.Cells(28, colSalida).Value = _
            listaResultados(i, 14)


        'ESTADO DE LA TARIFA
        wsC.Cells(29, colSalida).Value = "VALIDAR"


        'PENDIENTES / VALIDACIÓN
        wsC.Cells(30, colSalida).Value = ""


        'RANKING POR PRECIO
        wsC.Cells(31, colSalida).Value = i

    Next i


    '========================================================
    ' FORMATOS
    '========================================================

    wsC.Range("B17:F26").NumberFormat = _
        "$#,##0.00;$#,##0.00;-"

    wsC.Range("B31:F31").NumberFormat = "0"


    'Resaltar mejor opción
    With wsC.Range( _
        wsC.Cells(15, 2), _
        wsC.Cells(31, 2))

        .Interior.Color = RGB(198, 239, 206)

    End With


    With wsC.Range("B15:B31")

        .Font.Color = RGB(0, 97, 0)

    End With


    wsC.Range("B15").Font.Bold = True
    wsC.Range("B25").Font.Bold = True


    '========================================================
    ' ACTUALIZAR RESUMEN SUPERIOR
    '========================================================

    wsC.Range("E7").Value = listaResultados(1, 1)
    wsC.Range("E8").Value = listaResultados(1, 2)
    wsC.Range("E9").Value = listaResultados(1, 11)

    wsC.Range("E9").NumberFormat = _
        "$#,##0.00;$#,##0.00;-"


    If valorFOB > 0 Then

        wsC.Range("E11").Value = _
            listaResultados(1, 11) / valorFOB

    Else

        wsC.Range("E11").Value = 0

    End If

    wsC.Range("E11").NumberFormat = "0.00%"


    wsC.Range("E12").Value = _
        listaResultados(1, 13) _
        & " / " _
        & listaResultados(1, 14)


    '========================================================
    ' FINALIZAR
    '========================================================

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    MsgBox cantidadAgentes & " agente(s) encontrados." _
           & vbCrLf & vbCrLf _
           & "Mejor agente: " _
           & listaResultados(1, 1) _
           & vbCrLf _
           & "Naviera: " _
           & listaResultados(1, 2) _
           & vbCrLf _
           & "Costo total: USD " _
           & Format(listaResultados(1, 11), "#,##0.00"), _
           vbInformation, "Cotización FCL"

    Exit Sub


'============================================================
' CONTROL DE ERRORES
'============================================================

ErrorCotizar:

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Error: " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, "Error en Cotizar"


End Sub

Private Function ValorNumerico(ByVal valor As Variant) As Double

    If IsError(valor) Then

        ValorNumerico = 0

    ElseIf Trim(CStr(valor)) = "" Then

        ValorNumerico = 0

    ElseIf Trim(CStr(valor)) = "-" Then

        ValorNumerico = 0

    ElseIf IsNumeric(valor) Then

        ValorNumerico = CDbl(valor)

    Else

        ValorNumerico = 0

    End If

End Function

Function GastoDestino( _
    ByVal agente As String, _
    ByVal destino As String, _
    ByVal naviera As String, _
    ByVal cantidadContenedores As Double, _
    ByVal cantidadBL As Double) As Double


    Dim ws As Worksheet
    Dim tabla As ListObject
    Dim fila As ListRow

    Dim agenteCargo As String
    Dim destinoCargo As String
    Dim navieraCargo As String
    Dim estadoCargo As String
    Dim baseCargo As String

    Dim importe As Double
    Dim total As Double
    Dim navieraCoincide As Boolean


    On Error GoTo SinCargos


    Set ws = _
        ThisWorkbook.Worksheets("Cargos_Destino")

    Set tabla = _
        ws.ListObjects("tblCargosDestino")


    total = 0


    For Each fila In tabla.ListRows

        agenteCargo = UCase(Trim( _
            fila.Range.Cells(1, _
            tabla.ListColumns("Agente").Index).Value))

        destinoCargo = UCase(Trim( _
            fila.Range.Cells(1, _
            tabla.ListColumns("Puerto_Destino").Index).Value))

        navieraCargo = UCase(Trim( _
            fila.Range.Cells(1, _
            tabla.ListColumns("Naviera_Aplica").Index).Value))

        estadoCargo = UCase(Trim( _
            fila.Range.Cells(1, _
            tabla.ListColumns("Estado").Index).Value))

        baseCargo = UCase(Trim( _
            fila.Range.Cells(1, _
            tabla.ListColumns("Base").Index).Value))

        importe = ValorNumerico( _
            fila.Range.Cells(1, _
            tabla.ListColumns("Importe").Index).Value)


        navieraCoincide = _
            navieraCargo = "TODAS" _
            Or navieraCargo = "" _
            Or navieraCargo = UCase(Trim(naviera))


        If agenteCargo = UCase(Trim(agente)) _
        And destinoCargo = UCase(Trim(destino)) _
        And estadoCargo = "VIGENTE" _
        And navieraCoincide Then


            Select Case baseCargo

                Case "CONTENEDOR"

                    total = total _
                        + importe * cantidadContenedores


                Case "BL"

                    total = total _
                        + importe * cantidadBL


                Case "FIJO", "EMBARQUE", "FIJO EMBARQUE"

                    total = total + importe


                Case "PORCENTAJE", "PORCENTAJE_FLETE"

                    'El porcentaje no se incluye aquí
                    'hasta definir sobre qué base se calcula.

            End Select

        End If

    Next fila


    GastoDestino = total
    Exit Function


SinCargos:

    GastoDestino = 0

End Function

Sub GuardarHistorico()

    Dim wsC As Worksheet
    Dim wsH As Worksheet
    Dim tbl As ListObject
    Dim nuevaFila As ListRow

    Dim fechaCotizacion As Date
    Dim IDcotizacion As String
    Dim consecutivo As Long

    Dim origen As String
    Dim destino As String

    Dim cantidad20 As Double
    Dim cantidad40ST As Double
    Dim cantidad40HC As Double
    Dim cantidadBL As Double

    Dim valorFOB As Double

    Dim mejorAgente As String
    Dim mejorNaviera As String

    Dim costoTotal As Double
    Dim impactoFOB As Double

    On Error GoTo ErrorHistorico


    Set wsC = ThisWorkbook.Worksheets("Cotizador")
    Set wsH = ThisWorkbook.Worksheets("Historico")
    Set tbl = wsH.ListObjects("tblHistorico")


    '========================================
    ' LEER DATOS
    '========================================

    fechaCotizacion = wsC.Range("B5").Value

    origen = wsC.Range("B6").Value
    destino = wsC.Range("B7").Value

    cantidad20 = ValorNumerico(wsC.Range("B8").Value)
    cantidad40ST = ValorNumerico(wsC.Range("B9").Value)
    cantidad40HC = ValorNumerico(wsC.Range("B10").Value)

    cantidadBL = ValorNumerico(wsC.Range("B11").Value)

    valorFOB = ValorNumerico(wsC.Range("B12").Value)


    '========================================
    ' TOMAR MEJOR OPCIÓN
    '========================================

    mejorAgente = wsC.Range("B15").Value
    mejorNaviera = wsC.Range("B16").Value

    costoTotal = ValorNumerico(wsC.Range("B25").Value)


    '========================================
    ' VALIDAR
    '========================================

    If mejorAgente = "" _
    Or mejorAgente = "SIN OPCIÓN" _
    Or costoTotal <= 0 Then

        MsgBox "Primero debe generar una cotización válida.", _
               vbExclamation, "Sin cotización"

        Exit Sub

    End If


    '========================================
    ' GENERAR ID
    '========================================

    consecutivo = tbl.ListRows.Count + 1

    IDcotizacion = _
        "COT-" _
        & Format(fechaCotizacion, "yyyymmdd") _
        & "-" _
        & Format(consecutivo, "000")


    '========================================
    ' IMPACTO SOBRE FOB
    '========================================

    If valorFOB > 0 Then

        impactoFOB = costoTotal / valorFOB

    Else

        impactoFOB = 0

    End If


    '========================================
    ' CREAR NUEVA FILA
    '========================================

    If tbl.ListRows.Count = 1 _
    And WorksheetFunction.CountA(tbl.ListRows(1).Range) = 0 Then

        Set nuevaFila = tbl.ListRows(1)

    Else

        Set nuevaFila = tbl.ListRows.Add

    End If


    '========================================
    ' GUARDAR HISTÓRICO
    '========================================

    With nuevaFila.Range

        'ID
        .Cells(1, 1).Value = IDcotizacion

        'Fecha
        .Cells(1, 2).Value = fechaCotizacion

        'Origen
        .Cells(1, 3).Value = origen

        'Destino
        .Cells(1, 4).Value = destino

        '20 ST
        .Cells(1, 5).Value = cantidad20

        '40 ST
        .Cells(1, 6).Value = cantidad40ST

        '40 HC
        .Cells(1, 7).Value = cantidad40HC

        'FOB
        .Cells(1, 8).Value = valorFOB

        'Agente
        .Cells(1, 9).Value = mejorAgente

        'Naviera
        .Cells(1, 10).Value = mejorNaviera

        'Costo Total
        .Cells(1, 11).Value = costoTotal

        'Impacto FOB
        .Cells(1, 12).Value = impactoFOB

    End With


    '========================================
    ' FORMATOS
    '========================================

    nuevaFila.Range.Cells(1, 2).NumberFormat = "dd-mmm-yyyy"

    nuevaFila.Range.Cells(1, 8).NumberFormat = "$#,##0.00"

    nuevaFila.Range.Cells(1, 11).NumberFormat = "$#,##0.00"

    nuevaFila.Range.Cells(1, 12).NumberFormat = "0.00%"


    '========================================
    ' CONFIRMACIÓN
    '========================================

    MsgBox "Cotización guardada correctamente." _
           & vbCrLf & vbCrLf _
           & "ID: " & IDcotizacion _
           & vbCrLf _
           & "Ruta: " & origen & " - " & destino _
           & vbCrLf _
           & "Agente: " & mejorAgente _
           & vbCrLf _
           & "Costo Total: USD " _
           & Format(costoTotal, "#,##0.00"), _
           vbInformation, "Histórico actualizado"

    Exit Sub


ErrorHistorico:

    MsgBox "No fue posible guardar la cotización." _
           & vbCrLf & vbCrLf _
           & "Error: " & Err.Number _
           & vbCrLf _
           & Err.Description, _
           vbCritical, "Error al guardar histórico"

End Sub

Sub Limpiar()

    Dim wsC As Worksheet

    On Error GoTo ErrorLimpiar

    Set wsC = ThisWorkbook.Worksheets("Cotizador")

    Application.ScreenUpdating = False


    '========================================
    ' DATOS DE ENTRADA
    '========================================

    'Fecha actual
    wsC.Range("B5").Value = Date

    'Origen
    wsC.Range("B6").ClearContents

    'Destino
    wsC.Range("B7").ClearContents

    'Cantidades
    wsC.Range("B8").Value = 0
    wsC.Range("B9").Value = 0
    wsC.Range("B10").Value = 0

    'Número de BL
    wsC.Range("B11").Value = 1

    'FOB
    wsC.Range("B12").ClearContents


    '========================================
    ' LIMPIAR TABLA COMPARATIVA
    '========================================

    wsC.Range("B15:F31").ClearContents

    'Restaurar encabezados
    wsC.Range("B15:F15").Value = "SIN OPCIÓN"


    '========================================
    ' RESTAURAR FORMATO
    '========================================

    With wsC.Range("B15:F15")

        .Interior.Color = RGB(91, 155, 213)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True

    End With


    With wsC.Range("B16:F31")

        .Interior.Pattern = xlNone
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False

    End With


    With wsC.Range("B25:F25")

        .Interior.Color = RGB(255, 242, 204)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True

    End With


    '========================================
    ' LIMPIAR RESUMEN SUPERIOR
    ' IMPORTANTE:
    ' no usar E7:E12 porque hay celdas combinadas
    '========================================

    wsC.Range("E7").MergeArea.ClearContents
    wsC.Range("E8").MergeArea.ClearContents
    wsC.Range("E9").MergeArea.ClearContents
    wsC.Range("E10").MergeArea.ClearContents
    wsC.Range("E11").MergeArea.ClearContents
    wsC.Range("E12").MergeArea.ClearContents


    '========================================
    ' FINALIZAR
    '========================================

    Application.CalculateFull

    wsC.Range("B6").Select

    Application.ScreenUpdating = True


    MsgBox "Cotizador limpio correctamente." _
           & vbCrLf & vbCrLf _
           & "Puede iniciar una nueva cotización.", _
           vbInformation, "Nueva cotización"

    Exit Sub


ErrorLimpiar:

    Application.ScreenUpdating = True

    MsgBox "No fue posible limpiar el cotizador." _
           & vbCrLf & vbCrLf _
           & "Error: " & Err.Number _
           & vbCrLf _
           & Err.Description, _
           vbCritical, "Error"

End Sub

