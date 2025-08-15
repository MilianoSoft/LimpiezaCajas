Attribute VB_Name = "Módulo1"
Option Explicit

' ===============================
' MÓDULO DE ASIGNACIÓN DE LIMPIEZA DE CAJAS
' ===============================
' Uso:
' - Guardar el archivo como .xlsm
' - Importar este módulo (Archivo > Importar archivo)
' - Insertar un botón en "Asignación Actual" y asignarle la macro: GenerarAsignacion
' - Configuración en hoja "Configuracion":
'   A: Empleados AM, C: Empleados PM
'   E: Cajas AM, F: Cajas PM
'   B2: Semana actual, B3-B4: Horas AM, D3-D4: Horas PM
' ===============================

Public Sub GenerarAsignacion()
    Dim wsCfg As Worksheet, wsHist As Worksheet, wsAsig As Worksheet, wsHoras As Worksheet
    Dim semana As Long
    Dim amInicio As Date, amFin As Date, pmInicio As Date, pmFin As Date
    
    Set wsCfg = ThisWorkbook.Worksheets("Configuracion")
    Set wsHist = ThisWorkbook.Worksheets("RegistroHistorico")
    Set wsAsig = ThisWorkbook.Worksheets("AsignacionActual")
    Set wsHoras = ThisWorkbook.Worksheets("HorasClientes")
    
    ' Semana actual
    On Error Resume Next
    semana = CLng(wsCfg.Range("B2").Value)
    On Error GoTo 0
    If semana <= 0 Then
        MsgBox "No se pudo leer la semana actual en Configuracion!B2.", vbExclamation
        Exit Sub
    End If
    
    amInicio = TimeValue(NzStr(wsCfg.Range("B3").Value, "08:00"))
    amFin = TimeValue(NzStr(wsCfg.Range("B4").Value, "16:00"))
    pmInicio = TimeValue(NzStr(wsCfg.Range("D3").Value, "16:00"))
    pmFin = TimeValue(NzStr(wsCfg.Range("D4").Value, "00:00"))
    
    Dim empleadosAM As Collection, empleadosPM As Collection
    Dim cajasAM As Collection, cajasPM As Collection
    
    Set empleadosAM = LeerColumna(wsCfg, "A", 7)
    Set empleadosPM = LeerColumna(wsCfg, "C", 7)
    Set cajasAM = LeerColumna(wsCfg, "E", 7)
    Set cajasPM = LeerColumna(wsCfg, "F", 7)
    
    ' Validar que existan cajas y empleados
    If cajasAM.Count = 0 Or cajasPM.Count = 0 Then
        MsgBox "Define las cajas en Configuracion (columnas E y F).", vbExclamation
        Exit Sub
    End If
    If empleadosAM.Count = 0 Or empleadosPM.Count = 0 Then
        MsgBox "Define las listas de empleados AM/PM en Configuracion.", vbExclamation
        Exit Sub
    End If
    
    ' Leer horas "Baja" por turno
    Dim horasBajasAM As Collection, horasBajasPM As Collection
    Set horasBajasAM = HorasFiltradas(wsHoras, "Baja", amInicio, amFin)
    Set horasBajasPM = HorasFiltradas(wsHoras, "Baja", pmInicio, pmFin)
    
    ' Calcular asignaciones con rotación
    Dim asignAM As Collection, asignPM As Collection
    Set asignAM = AsignarTurno("AM", semana, empleadosAM, cajasAM, horasBajasAM, wsHist)
    Set asignPM = AsignarTurno("PM", semana, empleadosPM, cajasPM, horasBajasPM, wsHist)
    
    ' Pintar asignación
    PintarAsignacion wsAsig, asignAM, asignPM, cajasAM.Count
    
    ' Escribir historial
    EscribirHistorico wsHist, asignAM, semana, "AM"
    EscribirHistorico wsHist, asignPM, semana, "PM"
    
    MsgBox "Asignación generada para la semana " & semana & ".", vbInformation
End Sub

' -------------------------------
' Funciones auxiliares
' -------------------------------

Private Function NzStr(ByVal v As Variant, ByVal defVal As String) As String
    If IsError(v) Then NzStr = defVal: Exit Function
    If Len(Trim$(CStr(v))) = 0 Then NzStr = defVal Else NzStr = CStr(v)
End Function

Private Function LeerColumna(ws As Worksheet, ByVal col As String, ByVal startRow As Long) As Collection
    Dim c As Collection: Set c = New Collection
    Dim r As Long: r = startRow
    Do While Len(Trim$(ws.Range(col & r).Value)) > 0
        c.Add CStr(ws.Range(col & r).Value)
        r = r + 1
    Loop
    Set LeerColumna = c
End Function

Private Function HorasFiltradas(ws As Worksheet, ByVal clasif As String, ByVal hIni As Date, ByVal hFin As Date) As Collection
    Dim c As Collection: Set c = New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        Dim h As String, cl As String
        h = NzStr(ws.Cells(i, 1).Value, "")
        cl = NzStr(ws.Cells(i, 3).Value, "")
        If LCase$(Trim$(cl)) = LCase$(Trim$(clasif)) And Len(h) > 0 Then
            Dim th As Date
            On Error Resume Next
            th = TimeValue(h)
            On Error GoTo 0
            If th <> 0 Then
                If EnRangoHora(th, hIni, hFin) Then c.Add Format$(th, "hh:nn")
            End If
        End If
    Next i
    Set HorasFiltradas = c
End Function

Private Function EnRangoHora(ByVal h As Date, ByVal hIni As Date, ByVal hFin As Date) As Boolean
    If hIni <= hFin Then
        EnRangoHora = (h >= hIni And h <= hFin)
    Else
        EnRangoHora = (h >= hIni Or h <= hFin)
    End If
End Function

Private Function AsignarTurno(ByVal turno As String, ByVal semana As Long, _
                              empleados As Collection, cajas As Collection, _
                              horas As Collection, wsHist As Worksheet) As Collection
    Dim result As Collection: Set result = New Collection
    Dim conteos As Object: Set conteos = CreateObject("Scripting.Dictionary")
    Dim lastWeek As Object: Set lastWeek = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To empleados.Count
        conteos(empleados(i)) = 0
        lastWeek(empleados(i)) = 0
    Next i
    
    ' Leer histórico
    Dim lastRow As Long
    lastRow = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If NzStr(wsHist.Cells(i, 2).Value, "") = turno Then
            Dim emp As String: emp = NzStr(wsHist.Cells(i, 4).Value, "")
            If Len(emp) > 0 Then
                conteos(emp) = conteos(emp) + 1
                Dim wk As Long: wk = 0
                On Error Resume Next
                wk = CLng(wsHist.Cells(i, 1).Value)
                On Error GoTo 0
                If wk > 0 Then lastWeek(emp) = Application.WorksheetFunction.Max(lastWeek(emp), wk)
            End If
        End If
    Next i
    
    ' Ordenar empleados por conteo asc y última semana asignada
    Dim arr() As Variant
    ReDim arr(1 To empleados.Count, 1 To 3)
    For i = 1 To empleados.Count
        arr(i, 1) = empleados(i)
        arr(i, 2) = CLng(conteos(empleados(i)))
        arr(i, 3) = CLng(lastWeek(empleados(i)))
    Next i
    
    ' Burbuja simple
    Dim j As Long
    For i = 1 To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If (arr(i, 2) > arr(j, 2)) Or (arr(i, 2) = arr(j, 2) And arr(i, 3) > arr(j, 3)) Then
                Dim t1, t2, t3
                t1 = arr(i, 1): t2 = arr(i, 2): t3 = arr(i, 3)
                arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2): arr(i, 3) = arr(j, 3)
                arr(j, 1) = t1: arr(j, 2) = t2: arr(j, 3) = t3
            End If
        Next j
    Next i
    
    ' Asignar empleados a cajas
    Dim empIdx As Long: empIdx = 1
    For i = 1 To cajas.Count
        Dim triple(1 To 3) As Variant
        triple(1) = cajas(i)
        triple(2) = arr(empIdx, 1)
        If horas.Count >= i Then
            triple(3) = horas(i)
        Else
            triple(3) = "" ' asignación manual
        End If
        result.Add triple
        empIdx = empIdx + 1
        If empIdx > UBound(arr, 1) Then empIdx = 1
    Next i
    
    ' Aviso si no hay horas suficientes
    If horas.Count < cajas.Count Then
        MsgBox "Turno " & turno & ": no hay horas 'Baja' suficientes (" & horas.Count & "/" & cajas.Count & ")." & vbCrLf & _
               "Se dejaron horas en blanco para asignación manual.", vbExclamation
    End If
    
    Set AsignarTurno = result
End Function

Private Sub PintarAsignacion(ws As Worksheet, asignAM As Collection, asignPM As Collection, amRows As Long)
    Dim i As Long, startPM As Long
    
    ' --- Limpiar rango AM ---
    For i = 6 To 6 + amRows - 1
        ws.Cells(i, 1).Value = ""
        ws.Cells(i, 2).Value = ""
        ws.Cells(i, 3).Value = ""
    Next i
    
    ' --- Escribir datos AM ---
    For i = 1 To asignAM.Count
        With ws.Range(ws.Cells(5 + i, 1), ws.Cells(5 + i, 3))
            .Value = Array(asignAM(i)(1), asignAM(i)(2), asignAM(i)(3))
            .Font.Name = "Arial"
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(255, 255, 255)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
    Next i
    
    ' --- Calcular inicio dinámico de PM ---
    startPM = 5 + asignAM.Count + 2
    
    ' --- Escribir encabezado PM ---
    ws.Cells(startPM, 1).Value = "Turno PM"
    With ws.Range(ws.Cells(startPM, 1), ws.Cells(startPM, 3))
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 255)
    End With
    
    ws.Cells(startPM + 1, 1).Value = "Caja"
    ws.Cells(startPM + 1, 2).Value = "Empleado"
    ws.Cells(startPM + 1, 3).Value = "Hora Asignada"
    With ws.Range(ws.Cells(startPM + 1, 1), ws.Cells(startPM + 1, 3))
        .Font.Bold = True
        .Interior.Color = RGB(180, 180, 230)
    End With
    
    ' --- Limpiar rango PM ---
    For i = startPM + 2 To startPM + 2 + asignPM.Count + 5
        ws.Cells(i, 1).Value = ""
        ws.Cells(i, 2).Value = ""
        ws.Cells(i, 3).Value = ""
    Next i
    
    ' --- Escribir datos PM ---
    For i = 1 To asignPM.Count
        With ws.Range(ws.Cells(startPM + 1 + i, 1), ws.Cells(startPM + 1 + i, 3))
            .Value = Array(asignPM(i)(1), asignPM(i)(2), asignPM(i)(3))
            .Font.Name = "Arial"
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(255, 255, 255)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
    Next i
End Sub

Private Sub EscribirHistorico(ws As Worksheet, asign As Collection, semana As Long, turno As String)
    Dim i As Long, nextRow As Long
    For i = 1 To asign.Count
        nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(nextRow, 1).Value = semana
        ws.Cells(nextRow, 2).Value = turno
        ws.Cells(nextRow, 3).Value = asign(i)(1) ' Caja
        ws.Cells(nextRow, 4).Value = asign(i)(2) ' Empleado
        ws.Cells(nextRow, 5).Value = asign(i)(3) ' Hora
        ws.Cells(nextRow, 6).Value = Date
    Next i
End Sub

