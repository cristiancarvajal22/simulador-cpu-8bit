Attribute VB_Name = "ModuloMemoria"
Option Explicit

' ==============================================================================
' Modulo de Gestion de Memoria RAM (256 Bytes)
' Arquitectura Von Neumann - Bus de 8 bits (00h a FFh)
' ==============================================================================

Public RAM(0 To 255) As Byte

' Primitiva de lectura de memoria
Public Function ReadRAM(ByVal Address As Integer) As Byte
    ' Validar rango de 8 bits
    If Address < 0 Or Address > 255 Then
        MsgBox "Error de bus: Direccion de memoria fuera de rango (00h-FFh): " & Address, vbCritical, "Fallo de Memoria"
        ReadRAM = 0
        Exit Function
    End If
    ReadRAM = RAM(Address)
End Function

' Primitiva de escritura de memoria
Public Sub WriteRAM(ByVal Address As Integer, ByVal Value As Byte)
    ' Validar rango de 8 bits
    If Address < 0 Or Address > 255 Then
        MsgBox "Error de bus: Direccion de memoria fuera de rango (00h-FFh): " & Address, vbCritical, "Fallo de Memoria"
        Exit Sub
    End If
    RAM(Address) = Value
    ' Reflejar cambio en la matriz visual de la hoja
    Call ModuloUI.ActualizarCeldaMemoria(Address, Value)
End Sub

' Limpiar todo el espacio de memoria a 00h
Public Sub LimpiarMemoria()
    Dim i As Integer
    For i = 0 To 255
        RAM(i) = 0
    Next i
    Call ModuloUI.RefrescarTodaLaMemoria
End Sub

' Carga del programa demostrativo obligatorio en memoria
' Programa: Multiplicacion por sumas sucesivas con bucle y banderas
' Segmento de Datos:
'   80h = Multiplicando (05h = 5d)
'   81h = Multiplicador / Contador (03h = 3d)
'   82h = Resultado acumulado (inicialmente 00h, resultado final 15d = 0Fh)
' Segmento de Codigo:
'   00h: LOAD AX, [80h]    -> AX = 5
'   02h: LOAD BX, [81h]    -> BX = 3 (Contador)
'   04h: MOV AX, 00h       -> AX = 0 (Inicializar acumulador)
'   -- Bucle (direccion 06h) --
'   06h: ADD AX, 05h       -> AX = AX + 5 (Suma sucesiva)
'   08h: DEC BX            -> BX = BX - 1 (Decrementa contador)
'   09h: JNZ 06h           -> Salto condicional si ZF=0 a 06h
'   0Bh: STORE [82h], AX   -> Guardar producto final en RAM[82h]
'   0Dh: HLT               -> Detener procesador
Public Sub CargarProgramaDemostrativo()
    Call LimpiarMemoria
    
    ' Datos iniciales en Segmento de Datos (80h - FFh)
    RAM(&H80) = 5   ' Multiplicando
    RAM(&H81) = 3   ' Multiplicador / Contador
    RAM(&H82) = 0   ' Espacio para resultado final
    
    ' Instrucciones en Segmento de Codigo (00h - 7Fh)
    ' 00h: LOAD BX, [81h] -> Opcode 06h, Dir 81h
    RAM(&H0) = &H6: RAM(&H1) = &H81
    
    ' 02h: MOV AX, 00h -> Opcode 01h, Imm 00h
    RAM(&H2) = &H1: RAM(&H3) = &H0
    
    ' 04h: LOAD MDR / Dato auxiliar: suma sucesiva inmediata
    ' Bucle de suma sucesiva en 04h:
    ' 04h: ADD AX, 05h -> Opcode 10h, Imm 05h
    RAM(&H4) = &H10: RAM(&H5) = &H5
    
    ' 06h: DEC BX -> Opcode 17h (1 byte)
    RAM(&H6) = &H17
    
    ' 07h: JNZ 04h -> Opcode 32h, Dir 04h (2 bytes)
    RAM(&H7) = &H32: RAM(&H8) = &H4
    
    ' 09h: STORE [82h], AX -> Opcode 07h, Dir 82h (2 bytes)
    RAM(&H9) = &H7: RAM(&HA) = &H82
    
    ' 0Bh: HLT -> Opcode FFh (1 byte)
    RAM(&HB) = &HFF
    
    ' Actualizar la visualizacion completa en la hoja
    Call ModuloUI.RefrescarTodaLaMemoria
    Call ModuloUI.AgregarLog("SISTEMA", "Programa demostrativo de multiplicacion cargado en RAM (Codigo 00h-0Bh, Datos 80h-82h).")
End Sub
