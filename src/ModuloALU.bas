Attribute VB_Name = "ModuloALU"
Option Explicit

' ==============================================================================
' Modulo de Unidad Aritmetico-Logica (ALU)
' Procesamiento de operaciones de 8 bits y actualizacion de Flags (ZF, CF, SF)
' ==============================================================================

Public UltimaOperacionALU As String

' Ejecucion centralizada de operaciones de la ALU
Public Function CalcularALU(ByVal Operacion As String, ByVal OperandoA As Byte, ByVal OperandoB As Byte) As Byte
    Dim ResultadoInt As Long
    Dim ResByte As Byte
    
    UltimaOperacionALU = Operacion
    
    Select Case UCase(Trim(Operacion))
        Case "ADD"
            ResultadoInt = CLng(OperandoA) + CLng(OperandoB)
            ' Carry Flag si sobrepasa 255 (desborde sin signo)
            If ResultadoInt > 255 Then
                FlagCF = 1
            Else
                FlagCF = 0
            End If
            ResByte = CByte(ResultadoInt And &HFF)
            
        Case "SUB"
            ResultadoInt = CLng(OperandoA) - CLng(OperandoB)
            ' Carry Flag (Borrow) si es menor a cero
            If ResultadoInt < 0 Then
                FlagCF = 1
            Else
                FlagCF = 0
            End If
            ResByte = CByte(ResultadoInt And &HFF)
            
        Case "INC"
            ResultadoInt = CLng(OperandoA) + 1
            If ResultadoInt > 255 Then
                FlagCF = 1
            Else
                FlagCF = 0
            End If
            ResByte = CByte(ResultadoInt And &HFF)
            
        Case "DEC"
            ResultadoInt = CLng(OperandoA) - 1
            If ResultadoInt < 0 Then
                FlagCF = 1
            Else
                FlagCF = 0
            End If
            ResByte = CByte(ResultadoInt And &HFF)
            
        Case "CMP"
            ' Compara OperandoA con OperandoB sin modificar acumulador
            ResultadoInt = CLng(OperandoA) - CLng(OperandoB)
            If ResultadoInt < 0 Then
                FlagCF = 1
            Else
                FlagCF = 0
            End If
            ResByte = CByte(ResultadoInt And &HFF)
            ' Actualizar banderas de condicion
            If (ResByte = 0) And (ResultadoInt = 0) Then FlagZF = 1 Else FlagZF = 0
            If (ResByte And &H80) <> 0 Then FlagSF = 1 Else FlagSF = 0
            CalcularALU = OperandoA  ' Retorna el valor original de A
            Exit Function
            
        Case "AND"
            ResByte = OperandoA And OperandoB
            FlagCF = 0
            
        Case "OR"
            ResByte = OperandoA Or OperandoB
            FlagCF = 0
            
        Case "XOR"
            ResByte = OperandoA Xor OperandoB
            FlagCF = 0
            
        Case "NOT"
            ResByte = (Not OperandoA) And &HFF
            FlagCF = 0
            
        Case Else
            ResByte = OperandoA
    End Select
    
    ' Actualizacion estandar de banderas ZF y SF para operaciones con resultado
    ' Zero Flag (ZF): 1 si el resultado final de 8 bits es cero
    If ResByte = 0 Then
        FlagZF = 1
    Else
        FlagZF = 0
    End If
    
    ' Sign Flag (SF): 1 si el bit 7 (MSB) esta activo (negativo en complemento a 2)
    If (ResByte And &H80) <> 0 Then
        FlagSF = 1
    Else
        FlagSF = 0
    End If
    
    CalcularALU = ResByte
End Function
