Attribute VB_Name = "ModuloUI"
Option Explicit

' ==============================================================================
' Modulo de Interfaz de Usuario y Animacion Visual
' Resaltado de buses, celdas de RAM, actualizacion de registros y log
' ==============================================================================

Private Const HOJA_SIMULADOR As String = "Simulador"

' Colores estandar de segmentacion de memoria
Private Const COLOR_CODIGO As Long = 15918553    ' Azul claro (#D9E1F2)
Private Const COLOR_DATOS As Long = 14872546     ' Verde claro (#E2EFDA)
Private Const COLOR_RESALTADO As Long = 65535    ' Amarillo brillante (vbYellow)
Private Const COLOR_ACTIVO_FLAG As Long = 13561798 ' Verde pastel brillante
Private Const COLOR_INACTIVO_FLAG As Long = 15921906 ' Gris claro

' Posicion base de la matriz de memoria 16x16
Private Const FILA_BASE_RAM As Integer = 7
Private Const COL_BASE_RAM As Integer = 3        ' Columna C

' Referencia a la celda de memoria previamente resaltada
Private UltimaFilaResaltada As Integer
Private UltimaColResaltada As Integer
Private UltimaAddrResaltada As Integer

' Actualizar el valor de una celda individual de la RAM en la hoja
Public Sub ActualizarCeldaMemoria(ByVal Address As Integer, ByVal Value As Byte)
    Dim ws As Worksheet
    Dim r As Integer, c As Integer
    Dim h As String
    
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    r = FILA_BASE_RAM + (Address \ 16)
    c = COL_BASE_RAM + (Address Mod 16)
    
    h = Hex(Value)
    If Len(h) = 1 Then h = "0" & h
    ws.Cells(r, c).Value = h
End Sub

' Refrescar toda la matriz de memoria 16x16
Public Sub RefrescarTodaLaMemoria()
    Dim ws As Worksheet
    Dim i As Integer
    Dim r As Integer, c As Integer
    Dim h As String
    
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    Application.ScreenUpdating = False
    
    For i = 0 To 255
        r = FILA_BASE_RAM + (i \ 16)
        c = COL_BASE_RAM + (i Mod 16)
        h = Hex(ModuloMemoria.RAM(i))
        If Len(h) = 1 Then h = "0" & h
        ws.Cells(r, c).Value = h
        
        ' Restaurar color de segmentacion
        If i < 128 Then
            ws.Cells(r, c).Interior.Color = COLOR_CODIGO
        Else
            ws.Cells(r, c).Interior.Color = COLOR_DATOS
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub

' Resaltar la celda de RAM en ejecucion
Public Sub ResaltarCeldaRAM(ByVal Address As Integer, ByVal Activo As Boolean)
    Dim ws As Worksheet
    Dim r As Integer, c As Integer
    
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    r = FILA_BASE_RAM + (Address \ 16)
    c = COL_BASE_RAM + (Address Mod 16)
    
    If Activo Then
        ws.Cells(r, c).Interior.Color = COLOR_RESALTADO
        UltimaFilaResaltada = r
        UltimaColResaltada = c
        UltimaAddrResaltada = Address
    Else
        If Address < 128 Then
            ws.Cells(r, c).Interior.Color = COLOR_CODIGO
        Else
            ws.Cells(r, c).Interior.Color = COLOR_DATOS
        End If
    End If
End Sub

' Restaurar cualquier celda o componente resaltado a su color base
Public Sub RestaurarResaltados()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    
    ' Restaurar memoria
    If UltimaFilaResaltada >= FILA_BASE_RAM And UltimaFilaResaltada <= (FILA_BASE_RAM + 15) Then
        If UltimaAddrResaltada < 128 Then
            ws.Cells(UltimaFilaResaltada, UltimaColResaltada).Interior.Color = COLOR_CODIGO
        Else
            ws.Cells(UltimaFilaResaltada, UltimaColResaltada).Interior.Color = COLOR_DATOS
        End If
    End If
    
    ' Restaurar registros a blanco
    ws.Range("U7:W12").Interior.Color = RGB(255, 255, 255)
End Sub

' Resaltar un registro en el panel
Public Sub ResaltarComponente(ByVal Nombre As String, ByVal Activo As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    Dim filaReg As Integer
    
    Select Case UCase(Trim(Nombre))
        Case "PC": filaReg = 7
        Case "IR": filaReg = 8
        Case "MAR": filaReg = 9
        Case "MDR": filaReg = 10
        Case "AX": filaReg = 11
        Case "BX": filaReg = 12
        Case Else: Exit Sub
    End Select
    
    If Activo Then
        ws.Range(ws.Cells(filaReg, 21), ws.Cells(filaReg, 23)).Interior.Color = COLOR_RESALTADO
    Else
        ws.Range(ws.Cells(filaReg, 21), ws.Cells(filaReg, 23)).Interior.Color = RGB(255, 255, 255)
    End If
End Sub

' Actualizar todos los valores del panel de registros y banderas
Public Sub ActualizarPanelRegistros()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    
    ' Registros: Col U (Hex), Col V (Dec), Col W (Bin o Detalle)
    ws.Range("U7").Value = ModuloCPU.ByteToHex(ModuloCPU.PC)
    ws.Range("V7").Value = ModuloCPU.PC
    ws.Range("W7").Value = ModuloCPU.ByteToBin(ModuloCPU.PC)
    
    ws.Range("U8").Value = ModuloCPU.ByteToHex(ModuloCPU.IR)
    ws.Range("V8").Value = ModuloCPU.IR
    ws.Range("W8").Value = ModuloCPU.MnemonicActual
    
    ws.Range("U9").Value = ModuloCPU.ByteToHex(ModuloCPU.MAR)
    ws.Range("V9").Value = ModuloCPU.MAR
    ws.Range("W9").Value = ModuloCPU.ByteToBin(ModuloCPU.MAR)
    
    ws.Range("U10").Value = ModuloCPU.ByteToHex(ModuloCPU.MDR)
    ws.Range("V10").Value = ModuloCPU.MDR
    ws.Range("W10").Value = ModuloCPU.ByteToBin(ModuloCPU.MDR)
    
    ws.Range("U11").Value = ModuloCPU.ByteToHex(ModuloCPU.RegAX)
    ws.Range("V11").Value = ModuloCPU.RegAX
    ws.Range("W11").Value = ModuloCPU.ByteToBin(ModuloCPU.RegAX)
    
    ws.Range("U12").Value = ModuloCPU.ByteToHex(ModuloCPU.RegBX)
    ws.Range("V12").Value = ModuloCPU.RegBX
    ws.Range("W12").Value = ModuloCPU.ByteToBin(ModuloCPU.RegBX)
    
    ' Banderas de la ALU (ZF, CF, SF)
    ' ZF
    ws.Range("U16").Value = ModuloCPU.FlagZF
    If ModuloCPU.FlagZF = 1 Then
        ws.Range("U16").Interior.Color = COLOR_ACTIVO_FLAG
    Else
        ws.Range("U16").Interior.Color = COLOR_INACTIVO_FLAG
    End If
    
    ' CF
    ws.Range("V16").Value = ModuloCPU.FlagCF
    If ModuloCPU.FlagCF = 1 Then
        ws.Range("V16").Interior.Color = RGB(255, 199, 206) ' Rojo suave activo
    Else
        ws.Range("V16").Interior.Color = COLOR_INACTIVO_FLAG
    End If
    
    ' SF
    ws.Range("W16").Value = ModuloCPU.FlagSF
    If ModuloCPU.FlagSF = 1 Then
        ws.Range("W16").Interior.Color = RGB(255, 235, 156) ' Naranja suave activo
    Else
        ws.Range("W16").Interior.Color = COLOR_INACTIVO_FLAG
    End If
    
    ' Ultima operacion ALU
    ws.Range("U17").Value = ModuloALU.UltimaOperacionALU
End Sub

' Actualizar los 4 indicadores visuales de fase del ciclo
Public Sub ActualizarIndicadorFase(ByVal Fase As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    
    ' Restaurar todas las fases a color neutro
    ws.Range("C26:E26").Interior.Color = RGB(230, 230, 230)
    ws.Range("G26:I26").Interior.Color = RGB(230, 230, 230)
    ws.Range("K26:M26").Interior.Color = RGB(230, 230, 230)
    ws.Range("O26:Q26").Interior.Color = RGB(230, 230, 230)
    
    Select Case Fase
        Case 1 ' Fetch
            ws.Range("C26:E26").Interior.Color = COLOR_RESALTADO
        Case 2 ' Decode
            ws.Range("G26:I26").Interior.Color = COLOR_RESALTADO
        Case 3 ' Execute
            ws.Range("K26:M26").Interior.Color = COLOR_RESALTADO
        Case 4 ' Store
            ws.Range("O26:Q26").Interior.Color = COLOR_RESALTADO
    End Select
End Sub

' Agregar una entrada al log de micro-operaciones
Public Sub AgregarLog(ByVal Fase As String, ByVal Detalle As String)
    Dim ws As Worksheet
    Dim i As Integer
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    
    ' Desplazar filas del log hacia abajo (conserva las ultimas 12 lineas)
    For i = 37 To 26 Step -1
        ws.Range("T" & i & ":X" & i).Value = ws.Range("T" & (i - 1) & ":X" & (i - 1)).Value
    Next i
    
    ' Insertar nueva linea en la fila superior (fila 25)
    ws.Range("T25").Value = Format(ModuloCPU.ContadorPasos, "000")
    ws.Range("U25").Value = Fase
    ws.Range("V25").Value = ModuloCPU.ByteToHex(ModuloCPU.PC)
    ws.Range("W25").Value = ModuloCPU.ByteToHex(ModuloCPU.MAR)
    ws.Range("X25").Value = Detalle
End Sub

' Limpiar todo el historial del log
Public Sub LimpiarLog()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    ws.Range("T25:X37").ClearContents
End Sub

' Obtener el retardo configurado en la hoja (milisegundos)
Public Function ObtenerRetardoMs() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    Dim v As Variant
    v = ws.Range("V18").Value
    If IsNumeric(v) And v > 0 Then
        ObtenerRetardoMs = CLng(v)
    Else
        ObtenerRetardoMs = 250
    End If
End Function

' Inicializacion y maquetacion grafica completa de la hoja de calculo
Public Sub InicializarDisenoVisual()
    Dim ws As Worksheet
    Dim wsISA As Worksheet
    Dim shp As Shape
    Dim i As Integer, j As Integer, r As Integer
    Dim colL As String
    
    ' Obtener o crear hoja Simulador
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULADOR)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = HOJA_SIMULADOR
    End If
    On Error GoTo 0
    
    ws.Activate
    Application.ScreenUpdating = False
    
    ' Limpiar formas y contenido previo
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    ws.Cells.Clear
    
    ' Ancho de columnas
    ws.Range("A:A").ColumnWidth = 3
    ws.Range("B:B").ColumnWidth = 7
    For i = 3 To 18
        ws.Columns(i).ColumnWidth = 5.5
    Next i
    ws.Range("S:S").ColumnWidth = 3
    ws.Range("T:T").ColumnWidth = 20
    ws.Range("U:U").ColumnWidth = 10
    ws.Range("V:V").ColumnWidth = 10
    ws.Range("W:W").ColumnWidth = 14
    ws.Range("X:X").ColumnWidth = 32
    
    ' Banner Principal
    With ws.Range("B2:X3")
        .Merge
        .Value = "SIMULADOR DE CPU VON NEUMANN (8 BITS) & MEMORIA PRINCIPAL"
        .Font.Name = "Segoe UI"
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("B4:X4")
        .Merge
        .Value = "Arquitectura de Computadoras (SIS-131) | Ciclo Completo: Fetch -> Decode -> Execute -> Store"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Italic = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Matriz RAM Header
    With ws.Range("B5:R5")
        .Merge
        .Value = "MEMORIA RAM DE 8 BITS (256 BYTES: 00h a FFh)"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(38, 38, 38)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Encabezados de Columna (+0 a +F)
    ws.Range("B6").Value = "Base"
    Dim colsHex As Variant
    colsHex = Array("+0", "+1", "+2", "+3", "+4", "+5", "+6", "+7", "+8", "+9", "+A", "+B", "+C", "+D", "+E", "+F")
    For i = 0 To 15
        ws.Cells(6, 3 + i).Value = colsHex(i)
    Next i
    With ws.Range("B6:R6")
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Filas (00h a F0h)
    Dim rowsHex As Variant
    rowsHex = Array("00h", "10h", "20h", "30h", "40h", "50h", "60h", "70h", "80h", "90h", "A0h", "B0h", "C0h", "D0h", "E0h", "F0h")
    For i = 0 To 15
        ws.Cells(7 + i, 2).Value = rowsHex(i)
    Next i
    With ws.Range("B7:B22")
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Area de Datos RAM (Segmento Codigo 00h-7Fh y Datos 80h-FFh)
    With ws.Range("C7:R14")
        .Value = "00"
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_CODIGO
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(190, 190, 190)
    End With
    With ws.Range("C15:R22")
        .Value = "00"
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_DATOS
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(190, 190, 190)
    End With
    
    ' Leyenda
    ws.Range("B23:F23").Merge: ws.Range("B23:F23").Value = " Segmento Codigo (00h - 7Fh)": ws.Range("B23:F23").Interior.Color = COLOR_CODIGO
    ws.Range("G23:L23").Merge: ws.Range("G23:L23").Value = " Segmento Datos (80h - FFh)": ws.Range("G23:L23").Interior.Color = COLOR_DATOS
    ws.Range("M23:R23").Merge: ws.Range("M23:R23").Value = " Celda Activa (MAR / Bus)": ws.Range("M23:R23").Interior.Color = RGB(255, 242, 204)
    With ws.Range("B23:R23")
        .Font.Size = 9
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' Panel de Registros CPU
    With ws.Range("T5:X5")
        .Merge
        .Value = "REGISTROS DE LA CPU (8 BITS)"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(38, 38, 38)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("T6").Value = "Registro"
    ws.Range("U6").Value = "Hex"
    ws.Range("V6").Value = "Dec"
    ws.Range("W6:X6").Merge: ws.Range("W6:X6").Value = "Binario / Detalle"
    With ws.Range("T6:X6")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
    End With
    
    Dim regNames As Variant
    regNames = Array("PC (Program Counter)", "IR (Instruction Reg)", "MAR (Memory Address)", "MDR (Memory Data)", "AX (Acumulador)", "BX (Proposito Gral)")
    For i = 0 To 5
        r = 7 + i
        ws.Cells(r, 20).Value = regNames(i)
        ws.Cells(r, 20).Font.Bold = True
        ws.Cells(r, 21).Value = "00h"
        ws.Cells(r, 21).Font.Name = "Consolas"
        ws.Cells(r, 22).Value = 0
        ws.Range(ws.Cells(r, 23), ws.Cells(r, 24)).Merge
        ws.Cells(r, 23).Value = IIf(i = 1, "NOP", "00000000")
        ws.Cells(r, 23).Font.Name = "Consolas"
    Next i
    With ws.Range("T7:X12")
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    ws.Range("U7:X12").HorizontalAlignment = xlCenter
    
    ' Panel ALU y Flags
    With ws.Range("T14:X14")
        .Merge
        .Value = "UNIDAD ARITMETICO-LOGICA (ALU) & BANDERAS"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(38, 38, 38)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("T15").Value = "Ultima Op."
    ws.Range("U15").Value = "ZF (Cero)"
    ws.Range("V15").Value = "CF (Acarreo)"
    ws.Range("W15:X15").Merge: ws.Range("W15:X15").Value = "SF (Signo Ca2)"
    With ws.Range("T15:X15")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
    End With
    
    ws.Range("T16").Value = "NINGUNA"
    ws.Range("U16").Value = 0
    ws.Range("V16").Value = 0
    ws.Range("W16:X16").Merge: ws.Range("W16:X16").Value = 0
    With ws.Range("T16:X16")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    ws.Range("U16:X16").Interior.Color = COLOR_INACTIVO_FLAG
    
    ' Retardo
    ws.Range("T18:U18").Merge: ws.Range("T18:U18").Value = "Retardo de Reloj (ms):"
    ws.Range("T18:U18").Font.Bold = True
    ws.Range("V18").Value = 250
    ws.Range("V18").Font.Bold = True
    ws.Range("V18").HorizontalAlignment = xlCenter
    ws.Range("V18").Interior.Color = RGB(255, 255, 204)
    ws.Range("W18:X18").Merge: ws.Range("W18:X18").Value = "(Velocidad en modo RUN)"
    ws.Range("W18:X18").Font.Italic = True
    
    ' Indicador de Fases
    With ws.Range("B25:R25")
        .Merge
        .Value = "FASE ACTUAL DEL CICLO DE INSTRUCCION"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlCenter
    End With
    
    Dim faseRanges As Variant, faseTexts As Variant
    faseRanges = Array("C26:E26", "G26:I26", "K26:M26", "O26:Q26")
    faseTexts = Array("1. FETCH", "2. DECODE", "3. EXECUTE", "4. STORE")
    For i = 0 To 3
        With ws.Range(faseRanges(i))
            .Merge
            .Value = faseTexts(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = IIf(i = 0, COLOR_RESALTADO, RGB(235, 235, 235))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
        End With
    Next i
    
    ' Panel de Log
    With ws.Range("T20:X20")
        .Merge
        .Value = "LOG CRONOLOGICO DE MICRO-OPERACIONES"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(38, 38, 38)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("T21").Value = "Paso"
    ws.Range("U21").Value = "Fase"
    ws.Range("V21").Value = "PC"
    ws.Range("W21").Value = "MAR"
    ws.Range("X21").Value = "Detalle de Micro-operacion"
    With ws.Range("T21:X21")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
    End With
    
    For i = 22 To 37
        With ws.Range("T" & i & ":X" & i)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(230, 230, 230)
        End With
        ws.Range("T" & i & ":W" & i).HorizontalAlignment = xlCenter
        ws.Range("T" & i & ":W" & i).Font.Name = "Consolas"
        ws.Range("X" & i).Font.Name = "Segoe UI"
        ws.Range("X" & i).Font.Size = 9
    Next i
    
    ' Botones de Control
    Dim b1 As Shape, b2 As Shape, b3 As Shape, b4 As Shape, b5 As Shape
    Set b1 = ws.Shapes.AddShape(5, 40, 480, 120, 32)
    b1.TextFrame.Characters.Text = "PASO A PASO"
    b1.Fill.ForeColor.RGB = RGB(46, 117, 182)
    b1.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    b1.TextFrame.Characters.Font.Bold = True
    b1.OnAction = "btn_Step"
    
    Set b2 = ws.Shapes.AddShape(5, 170, 480, 110, 32)
    b2.TextFrame.Characters.Text = "EJECUTAR"
    b2.Fill.ForeColor.RGB = RGB(56, 142, 60)
    b2.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    b2.TextFrame.Characters.Font.Bold = True
    b2.OnAction = "btn_Run"
    
    Set b3 = ws.Shapes.AddShape(5, 290, 480, 100, 32)
    b3.TextFrame.Characters.Text = "PAUSAR"
    b3.Fill.ForeColor.RGB = RGB(217, 83, 79)
    b3.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    b3.TextFrame.Characters.Font.Bold = True
    b3.OnAction = "btn_Pause"
    
    Set b4 = ws.Shapes.AddShape(5, 400, 480, 100, 32)
    b4.TextFrame.Characters.Text = "REINICIAR"
    b4.Fill.ForeColor.RGB = RGB(108, 117, 125)
    b4.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    b4.TextFrame.Characters.Font.Bold = True
    b4.OnAction = "btn_Reset"
    
    Set b5 = ws.Shapes.AddShape(5, 510, 480, 150, 32)
    b5.TextFrame.Characters.Text = "CARGAR DEMO"
    b5.Fill.ForeColor.RGB = RGB(112, 48, 160)
    b5.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    b5.TextFrame.Characters.Font.Bold = True
    b5.OnAction = "btn_LoadDemo"
    
    ' Hoja Repertorio ISA
    On Error Resume Next
    Set wsISA = ThisWorkbook.Sheets("Repertorio_ISA")
    If wsISA Is Nothing Then
        Set wsISA = ThisWorkbook.Sheets.Add(After:=ws)
        wsISA.Name = "Repertorio_ISA"
    End If
    On Error GoTo 0
    
    wsISA.Cells.Clear
    With wsISA.Range("B2:F2")
        .Merge
        .Value = "REPERTORIO FORMAL DE INSTRUCCIONES (ISA 8-BITS)"
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    
    Dim isaH As Variant
    isaH = Array("Mnemonic", "Sintaxis", "Opcode (Hex)", "Bytes", "Descripcion y Efecto")
    For i = 0 To 4
        wsISA.Cells(4, 2 + i).Value = isaH(i)
    Next i
    With wsISA.Range("B4:F4")
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    Dim filaISA As Integer
    filaISA = 5
    Call SetFilaISA(wsISA, filaISA, "MOV", "MOV AX, imm", "01h", "2", "AX <- Inmediato de 8 bits")
    Call SetFilaISA(wsISA, filaISA, "MOV", "MOV BX, imm", "02h", "2", "BX <- Inmediato de 8 bits")
    Call SetFilaISA(wsISA, filaISA, "MOV", "MOV AX, BX", "03h", "1", "AX <- BX")
    Call SetFilaISA(wsISA, filaISA, "MOV", "MOV BX, AX", "04h", "1", "BX <- AX")
    Call SetFilaISA(wsISA, filaISA, "LOAD", "LOAD AX, [dir]", "05h", "2", "AX <- RAM[dir]")
    Call SetFilaISA(wsISA, filaISA, "LOAD", "LOAD BX, [dir]", "06h", "2", "BX <- RAM[dir]")
    Call SetFilaISA(wsISA, filaISA, "STORE", "STORE [dir], AX", "07h", "2", "RAM[dir] <- AX")
    Call SetFilaISA(wsISA, filaISA, "STORE", "STORE [dir], BX", "08h", "2", "RAM[dir] <- BX")
    Call SetFilaISA(wsISA, filaISA, "ADD", "ADD AX, imm", "10h", "2", "AX <- AX + imm (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "ADD", "ADD AX, BX", "11h", "1", "AX <- AX + BX (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "SUB", "SUB AX, imm", "12h", "2", "AX <- AX - imm (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "SUB", "SUB AX, BX", "13h", "1", "AX <- AX - BX (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "INC", "INC AX", "14h", "1", "AX <- AX + 1 (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "INC", "INC BX", "15h", "1", "BX <- BX + 1 (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "DEC", "DEC AX", "16h", "1", "AX <- AX - 1 (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "DEC", "DEC BX", "17h", "1", "BX <- BX - 1 (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "CMP", "CMP AX, imm", "18h", "2", "Compara AX con imm sin alterar AX (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "CMP", "CMP AX, BX", "19h", "1", "Compara AX con BX sin alterar AX (Flags ZF, CF, SF)")
    Call SetFilaISA(wsISA, filaISA, "AND", "AND AX, imm", "20h", "2", "AX <- AX AND imm (Flags ZF, SF)")
    Call SetFilaISA(wsISA, filaISA, "AND", "AND AX, BX", "21h", "1", "AX <- AX AND BX (Flags ZF, SF)")
    Call SetFilaISA(wsISA, filaISA, "OR", "OR AX, imm", "22h", "2", "AX <- AX OR imm (Flags ZF, SF)")
    Call SetFilaISA(wsISA, filaISA, "OR", "OR AX, BX", "23h", "1", "AX <- AX OR BX (Flags ZF, SF)")
    Call SetFilaISA(wsISA, filaISA, "XOR", "XOR AX, imm", "24h", "2", "AX <- AX XOR imm (Flags ZF, SF)")
    Call SetFilaISA(wsISA, filaISA, "XOR", "XOR AX, BX", "25h", "1", "AX <- AX XOR BX (Flags ZF, SF)")
    Call SetFilaISA(wsISA, filaISA, "NOT", "NOT AX", "26h", "1", "AX <- NOT AX (Flags ZF, SF)")
    Call SetFilaISA(wsISA, filaISA, "JMP", "JMP dir", "30h", "2", "Salto incondicional: PC <- dir")
    Call SetFilaISA(wsISA, filaISA, "JZ", "JZ dir", "31h", "2", "Salto si Zero: Si ZF=1 entonces PC <- dir")
    Call SetFilaISA(wsISA, filaISA, "JNZ", "JNZ dir", "32h", "2", "Salto si Not Zero: Si ZF=0 entonces PC <- dir")
    Call SetFilaISA(wsISA, filaISA, "HLT", "HLT", "FFh", "1", "Detiene el ciclo de reloj y ejecucion")
    
    wsISA.Columns(2).ColumnWidth = 14
    wsISA.Columns(3).ColumnWidth = 20
    wsISA.Columns(4).ColumnWidth = 16
    wsISA.Columns(5).ColumnWidth = 10
    wsISA.Columns(6).ColumnWidth = 45
    
    ws.Activate
    Application.ScreenUpdating = True
    
    Call ModuloCPU.ResetCPU
End Sub

Private Sub SetFilaISA(ByVal ws As Worksheet, ByRef fila As Integer, ByVal mnem As String, ByVal sint As String, ByVal opc As String, ByVal b As String, ByVal desc As String)
    ws.Cells(fila, 2).Value = mnem
    ws.Cells(fila, 2).HorizontalAlignment = xlCenter
    ws.Cells(fila, 2).Font.Name = "Consolas"
    
    ws.Cells(fila, 3).Value = sint
    
    ws.Cells(fila, 4).Value = opc
    ws.Cells(fila, 4).HorizontalAlignment = xlCenter
    ws.Cells(fila, 4).Font.Name = "Consolas"
    
    ws.Cells(fila, 5).Value = b
    ws.Cells(fila, 5).HorizontalAlignment = xlCenter
    ws.Cells(fila, 5).Font.Name = "Consolas"
    
    ws.Cells(fila, 6).Value = desc
    
    ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 6)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 6)).Borders.Color = RGB(210, 210, 210)
    
    fila = fila + 1
End Sub
