Attribute VB_Name = "ModuloCPU"
Option Explicit

' ==============================================================================
' Modulo de CPU y Registros de 8 Bits (Arquitectura Von Neumann / x86)
' ==============================================================================

' Registros del Procesador (8 bits c/u)
Public PC As Byte          ' Program Counter (Puntero a proxima instruccion)
Public IR As Byte          ' Instruction Register (Opcode y operacion activa)
Public MAR As Byte         ' Memory Address Register (Bus de direcciones)
Public MDR As Byte         ' Memory Data Register / MBR (Bus de datos)
Public RegAX As Byte       ' Acumulador (Proposito general y ALU)
Public RegBX As Byte       ' Registro Base / Auxiliar

' Registro de Estado / Flags (1 bit c/u)
Public FlagZF As Byte      ' Zero Flag (1 = resultado nulo)
Public FlagCF As Byte      ' Carry Flag (1 = acarreo o desborde sin signo)
Public FlagSF As Byte      ' Sign Flag (1 = bit 7 activo / negativo en Ca2)

' Estado de Control del Ciclo de Reloj
' 1 = FETCH, 2 = DECODE, 3 = EXECUTE, 4 = STORE
Public FaseCiclo As Integer
Public EnEjecucion As Boolean
Public CpuDetenida As Boolean
Public ContadorPasos As Long

' Variables internas para decodificacion y micro-operaciones
Public MnemonicActual As String
Public ModoOperando As String
Public OperandoDato As Byte
Public DireccionEfectiva As Byte
Public UltimoResultadoALU As Byte
Public DestinoEsRAM As Boolean
Public DestinoEsAX As Boolean
Public DestinoEsBX As Boolean

' Reiniciar registros y estado a condiciones iniciales
Public Sub ResetCPU()
    PC = 0
    IR = 0
    MAR = 0
    MDR = 0
    RegAX = 0
    RegBX = 0
    FlagZF = 0
    FlagCF = 0
    FlagSF = 0
    
    FaseCiclo = 1   ' Comienza siempre en FETCH
    EnEjecucion = False
    CpuDetenida = False
    ContadorPasos = 0
    MnemonicActual = "NOP"
    ModoOperando = ""
    OperandoDato = 0
    DireccionEfectiva = 0
    UltimoResultadoALU = 0
    DestinoEsRAM = False
    DestinoEsAX = False
    DestinoEsBX = False
    
    Call ModuloUI.RestaurarResaltados
    Call ModuloUI.ActualizarPanelRegistros
    Call ModuloUI.ActualizarIndicadorFase(1)
    Call ModuloUI.LimpiarLog
    Call ModuloUI.AgregarLog("RESET", "CPU y registros reiniciados a cero. Ciclo listo en FETCH.")
End Sub

' Funciones auxiliares de formato numerico para visualizacion
Public Function ByteToHex(ByVal val As Byte) As String
    Dim h As String
    h = Hex(val)
    If Len(h) = 1 Then h = "0" & h
    ByteToHex = h & "h"
End Function

Public Function ByteToBin(ByVal val As Byte) As String
    Dim b As String
    Dim i As Integer
    b = ""
    For i = 7 To 0 Step -1
        If (val And (2 ^ i)) <> 0 Then
            b = b & "1"
        Else
            b = b & "0"
        End If
    Next i
    ByteToBin = b
End Function
