Attribute VB_Name = "ModuloCiclo"
Option Explicit

' ==============================================================================
' Modulo de Control del Ciclo de Instruccion y Reloj de la CPU
' Fases formales: Fetch -> Decode -> Execute -> Store
' ==============================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Ejecutar un unico paso del ciclo de reloj (Avanza una fase)
Public Sub PasoCiclo()
    If CpuDetenida Then
        Call ModuloUI.AgregarLog("ESTADO", "La CPU esta detenida (HLT). Presione RESET o CARGAR PROGRAMA para reiniciar.")
        Exit Sub
    End If
    
    ContadorPasos = ContadorPasos + 1
    Call ModuloUI.RestaurarResaltados
    
    Select Case FaseCiclo
        Case 1
            Call FaseFetch
            FaseCiclo = 2
            Call ModuloUI.ActualizarIndicadorFase(2)
            
        Case 2
            Call FaseDecode
            If CpuDetenida Then
                FaseCiclo = 1
                Call ModuloUI.ActualizarIndicadorFase(1)
            Else
                FaseCiclo = 3
                Call ModuloUI.ActualizarIndicadorFase(3)
            End If
            
        Case 3
            Call FaseExecute
            FaseCiclo = 4
            Call ModuloUI.ActualizarIndicadorFase(4)
            
        Case 4
            Call FaseStore
            FaseCiclo = 1
            Call ModuloUI.ActualizarIndicadorFase(1)
    End Select
    
    Call ModuloUI.ActualizarPanelRegistros
End Sub

' ------------------------------------------------------------------------------
' 1. FASE FETCH: Busqueda de la instruccion en memoria
' MAR <- PC; MDR <- RAM[MAR]; IR <- MDR; PC <- PC + 1
' ------------------------------------------------------------------------------
Private Sub FaseFetch()
    MAR = PC
    Call ModuloUI.ResaltarComponente("PC", True)
    Call ModuloUI.ResaltarComponente("MAR", True)
    Call ModuloUI.ResaltarCeldaRAM(MAR, True)
    
    MDR = ModuloMemoria.ReadRAM(MAR)
    IR = MDR
    
    Call ModuloUI.ResaltarComponente("MDR", True)
    Call ModuloUI.ResaltarComponente("IR", True)
    
    ' Incrementar Program Counter a la siguiente posicion
    PC = CByte((CInt(PC) + 1) And &HFF)
    
    Call ModuloUI.AgregarLog("FETCH", "MAR=" & ByteToHex(MAR) & " | MDR=" & ByteToHex(MDR) & " -> IR=" & ByteToHex(IR) & " | PC=" & ByteToHex(PC))
End Sub

' ------------------------------------------------------------------------------
' 2. FASE DECODE: Decodificacion de Opcode y obtencion de operandos
' ------------------------------------------------------------------------------
Private Sub FaseDecode()
    DestinoEsRAM = False
    DestinoEsAX = False
    DestinoEsBX = False
    ModoOperando = ""
    OperandoDato = 0
    DireccionEfectiva = 0
    
    Select Case IR
        Case &HFF ' HLT
            MnemonicActual = "HLT"
            CpuDetenida = True
            EnEjecucion = False
            Call ModuloUI.AgregarLog("DECODE", "Instruccion HLT detectada. Reloj y procesador detenidos.")
            Exit Sub
            
        Case &H1 ' MOV AX, imm (2 bytes)
            MnemonicActual = "MOV AX, imm"
            MAR = PC
            OperandoDato = ModuloMemoria.ReadRAM(MAR)
            MDR = OperandoDato
            PC = CByte((CInt(PC) + 1) And &HFF)
            DestinoEsAX = True
            
        Case &H2 ' MOV BX, imm (2 bytes)
            MnemonicActual = "MOV BX, imm"
            MAR = PC
            OperandoDato = ModuloMemoria.ReadRAM(MAR)
            MDR = OperandoDato
            PC = CByte((CInt(PC) + 1) And &HFF)
            DestinoEsBX = True
            
        Case &H3 ' MOV AX, BX (1 byte)
            MnemonicActual = "MOV AX, BX"
            OperandoDato = RegBX
            DestinoEsAX = True
            
        Case &H4 ' MOV BX, AX (1 byte)
            MnemonicActual = "MOV BX, AX"
            OperandoDato = RegAX
            DestinoEsBX = True
            
        Case &H5 ' LOAD AX, [dir] (2 bytes)
            MnemonicActual = "LOAD AX, [dir]"
            MAR = PC
            DireccionEfectiva = ModuloMemoria.ReadRAM(MAR)
            MDR = DireccionEfectiva
            PC = CByte((CInt(PC) + 1) And &HFF)
            DestinoEsAX = True
            
        Case &H6 ' LOAD BX, [dir] (2 bytes)
            MnemonicActual = "LOAD BX, [dir]"
            MAR = PC
            DireccionEfectiva = ModuloMemoria.ReadRAM(MAR)
            MDR = DireccionEfectiva
            PC = CByte((CInt(PC) + 1) And &HFF)
            DestinoEsBX = True
            
        Case &H7 ' STORE [dir], AX (2 bytes)
            MnemonicActual = "STORE [dir], AX"
            MAR = PC
            DireccionEfectiva = ModuloMemoria.ReadRAM(MAR)
            MDR = DireccionEfectiva
            PC = CByte((CInt(PC) + 1) And &HFF)
            OperandoDato = RegAX
            DestinoEsRAM = True
            
        Case &H8 ' STORE [dir], BX (2 bytes)
            MnemonicActual = "STORE [dir], BX"
            MAR = PC
            DireccionEfectiva = ModuloMemoria.ReadRAM(MAR)
            MDR = DireccionEfectiva
            PC = CByte((CInt(PC) + 1) And &HFF)
            OperandoDato = RegBX
            DestinoEsRAM = True
            
        Case &H10 ' ADD AX, imm (2 bytes)
            MnemonicActual = "ADD AX, imm"
            MAR = PC
            OperandoDato = ModuloMemoria.ReadRAM(MAR)
            MDR = OperandoDato
            PC = CByte((CInt(PC) + 1) And &HFF)
            DestinoEsAX = True
            
        Case &H11 ' ADD AX, BX (1 byte)
            MnemonicActual = "ADD AX, BX"
            OperandoDato = RegBX
            DestinoEsAX = True
            
        Case &H12 ' SUB AX, imm (2 bytes)
            MnemonicActual = "SUB AX, imm"
            MAR = PC
            OperandoDato = ModuloMemoria.ReadRAM(MAR)
            MDR = OperandoDato
            PC = CByte((CInt(PC) + 1) And &HFF)
            DestinoEsAX = True
            
        Case &H13 ' SUB AX, BX (1 byte)
            MnemonicActual = "SUB AX, BX"
            OperandoDato = RegBX
            DestinoEsAX = True
            
        Case &H14 ' INC AX (1 byte)
            MnemonicActual = "INC AX"
            DestinoEsAX = True
            
        Case &H15 ' INC BX (1 byte)
            MnemonicActual = "INC BX"
            DestinoEsBX = True
            
        Case &H16 ' DEC AX (1 byte)
            MnemonicActual = "DEC AX"
            DestinoEsAX = True
            
        Case &H17 ' DEC BX (1 byte)
            MnemonicActual = "DEC BX"
            DestinoEsBX = True
            
        Case &H18 ' CMP AX, imm (2 bytes)
            MnemonicActual = "CMP AX, imm"
            MAR = PC
            OperandoDato = ModuloMemoria.ReadRAM(MAR)
            MDR = OperandoDato
            PC = CByte((CInt(PC) + 1) And &HFF)
            
        Case &H19 ' CMP AX, BX (1 byte)
            MnemonicActual = "CMP AX, BX"
            OperandoDato = RegBX
            
        Case &H20 ' AND AX, imm
            MnemonicActual = "AND AX, imm"
            MAR = PC
            OperandoDato = ModuloMemoria.ReadRAM(MAR)
            MDR = OperandoDato
            PC = CByte((CInt(PC) + 1) And &HFF)
            DestinoEsAX = True
            
        Case &H21 ' AND AX, BX
            MnemonicActual = "AND AX, BX"
            OperandoDato = RegBX
            DestinoEsAX = True
            
        Case &H22 ' OR AX, imm
            MnemonicActual = "OR AX, imm"
            MAR = PC
            OperandoDato = ModuloMemoria.ReadRAM(MAR)
            MDR = OperandoDato
            PC = CByte((CInt(PC) + 1) And &HFF)
            DestinoEsAX = True
            
        Case &H23 ' OR AX, BX
            MnemonicActual = "OR AX, BX"
            OperandoDato = RegBX
            DestinoEsAX = True
            
        Case &H24 ' XOR AX, imm
            MnemonicActual = "XOR AX, imm"
            MAR = PC
            OperandoDato = ModuloMemoria.ReadRAM(MAR)
            MDR = OperandoDato
            PC = CByte((CInt(PC) + 1) And &HFF)
            DestinoEsAX = True
            
        Case &H25 ' XOR AX, BX
            MnemonicActual = "XOR AX, BX"
            OperandoDato = RegBX
            DestinoEsAX = True
            
        Case &H26 ' NOT AX
            MnemonicActual = "NOT AX"
            DestinoEsAX = True
            
        Case &H30 ' JMP dir (2 bytes)
            MnemonicActual = "JMP dir"
            MAR = PC
            DireccionEfectiva = ModuloMemoria.ReadRAM(MAR)
            MDR = DireccionEfectiva
            PC = CByte((CInt(PC) + 1) And &HFF)
            
        Case &H31 ' JZ dir (2 bytes)
            MnemonicActual = "JZ dir"
            MAR = PC
            DireccionEfectiva = ModuloMemoria.ReadRAM(MAR)
            MDR = DireccionEfectiva
            PC = CByte((CInt(PC) + 1) And &HFF)
            
        Case &H32 ' JNZ dir (2 bytes)
            MnemonicActual = "JNZ dir"
            MAR = PC
            DireccionEfectiva = ModuloMemoria.ReadRAM(MAR)
            MDR = DireccionEfectiva
            PC = CByte((CInt(PC) + 1) And &HFF)
            
        Case Else
            MnemonicActual = "UNKNOWN (" & ByteToHex(IR) & ")"
            Call ModuloUI.AgregarLog("DECODE", "Opcode no reconocido: " & ByteToHex(IR))
            Exit Sub
    End Select
    
    Call ModuloUI.AgregarLog("DECODE", "Instruccion: " & MnemonicActual & " | Operando/Dir=" & ByteToHex(OperandoDato))
End Sub

' ------------------------------------------------------------------------------
' 3. FASE EXECUTE: Cómputo en ALU o evaluación de bifurcación
' ------------------------------------------------------------------------------
Private Sub FaseExecute()
    Call ModuloUI.ResaltarComponente("ALU", True)
    
    Select Case IR
        Case &H1, &H2, &H3, &H4 ' MOV
            UltimoResultadoALU = OperandoDato
            
        Case &H5, &H6 ' LOAD
            MAR = DireccionEfectiva
            UltimoResultadoALU = ModuloMemoria.ReadRAM(MAR)
            MDR = UltimoResultadoALU
            
        Case &H7, &H8 ' STORE
            MAR = DireccionEfectiva
            MDR = OperandoDato
            UltimoResultadoALU = OperandoDato
            
        Case &H10, &H11 ' ADD
            UltimoResultadoALU = ModuloALU.CalcularALU("ADD", RegAX, OperandoDato)
            
        Case &H12, &H13 ' SUB
            UltimoResultadoALU = ModuloALU.CalcularALU("SUB", RegAX, OperandoDato)
            
        Case &H14 ' INC AX
            UltimoResultadoALU = ModuloALU.CalcularALU("INC", RegAX, 0)
            
        Case &H15 ' INC BX
            UltimoResultadoALU = ModuloALU.CalcularALU("INC", RegBX, 0)
            
        Case &H16 ' DEC AX
            UltimoResultadoALU = ModuloALU.CalcularALU("DEC", RegAX, 0)
            
        Case &H17 ' DEC BX
            UltimoResultadoALU = ModuloALU.CalcularALU("DEC", RegBX, 0)
            
        Case &H18, &H19 ' CMP
            UltimoResultadoALU = ModuloALU.CalcularALU("CMP", RegAX, OperandoDato)
            
        Case &H20, &H21 ' AND
            UltimoResultadoALU = ModuloALU.CalcularALU("AND", RegAX, OperandoDato)
            
        Case &H22, &H23 ' OR
            UltimoResultadoALU = ModuloALU.CalcularALU("OR", RegAX, OperandoDato)
            
        Case &H24, &H25 ' XOR
            UltimoResultadoALU = ModuloALU.CalcularALU("XOR", RegAX, OperandoDato)
            
        Case &H26 ' NOT
            UltimoResultadoALU = ModuloALU.CalcularALU("NOT", RegAX, 0)
            
        Case &H30 ' JMP incondicional
            PC = DireccionEfectiva
            Call ModuloUI.AgregarLog("EXECUTE", "Bifurcacion JMP ejecutada hacia dir " & ByteToHex(PC))
            Exit Sub
            
        Case &H31 ' JZ (Salto si ZF = 1)
            If FlagZF = 1 Then
                PC = DireccionEfectiva
                Call ModuloUI.AgregarLog("EXECUTE", "Bifurcacion JZ tomada (ZF=1) hacia dir " & ByteToHex(PC))
            Else
                Call ModuloUI.AgregarLog("EXECUTE", "Bifurcacion JZ ignorada (ZF=0). Flujo continua en " & ByteToHex(PC))
            End If
            Exit Sub
            
        Case &H32 ' JNZ (Salto si ZF = 0)
            If FlagZF = 0 Then
                PC = DireccionEfectiva
                Call ModuloUI.AgregarLog("EXECUTE", "Bifurcacion JNZ tomada (ZF=0) hacia dir " & ByteToHex(PC))
            Else
                Call ModuloUI.AgregarLog("EXECUTE", "Bifurcacion JNZ ignorada (ZF=1). Flujo continua en " & ByteToHex(PC))
            End If
            Exit Sub
    End Select
    
    Call ModuloUI.AgregarLog("EXECUTE", "ALU Res=" & ByteToHex(UltimoResultadoALU) & " | ZF=" & FlagZF & " CF=" & FlagCF & " SF=" & FlagSF)
End Sub

' ------------------------------------------------------------------------------
' 4. FASE STORE / WRITE-BACK: Almacenamiento en Registro o RAM
' ------------------------------------------------------------------------------
Private Sub FaseStore()
    If DestinoEsAX Then
        RegAX = UltimoResultadoALU
        Call ModuloUI.ResaltarComponente("AX", True)
        Call ModuloUI.AgregarLog("STORE", "Resultado escrito en registro AX: " & ByteToHex(RegAX))
    ElseIf DestinoEsBX Then
        RegBX = UltimoResultadoALU
        Call ModuloUI.ResaltarComponente("BX", True)
        Call ModuloUI.AgregarLog("STORE", "Resultado escrito en registro BX: " & ByteToHex(RegBX))
    ElseIf DestinoEsRAM Then
        ModuloMemoria.WriteRAM MAR, MDR
        Call ModuloUI.ResaltarCeldaRAM(MAR, True)
        Call ModuloUI.AgregarLog("STORE", "Dato escrito en RAM[" & ByteToHex(MAR) & "] = " & ByteToHex(MDR))
    Else
        Call ModuloUI.AgregarLog("STORE", "Sin escritura en registro/RAM (Operacion completada).")
    End If
End Sub

' ------------------------------------------------------------------------------
' Control de Ejecucion Continua (RUN / PAUSE)
' ------------------------------------------------------------------------------
Public Sub EjecutarContinuo()
    If CpuDetenida Then
        Call ModuloUI.AgregarLog("ESTADO", "La CPU esta detenida. Reinicie antes de ejecutar.")
        Exit Sub
    End If
    
    EnEjecucion = True
    Call ModuloUI.AgregarLog("CONTROL", "Iniciando ejecucion continua...")
    
    Dim retardoMs As Long
    retardoMs = ModuloUI.ObtenerRetardoMs()
    If retardoMs < 20 Then retardoMs = 20
    
    Do While EnEjecucion And Not CpuDetenida
        Call PasoCiclo
        DoEvents
        Sleep retardoMs
    Loop
    
    If CpuDetenida Then
        Call ModuloUI.AgregarLog("CONTROL", "Ejecucion terminada exitosamente por instruccion HLT.")
    Else
        Call ModuloUI.AgregarLog("CONTROL", "Ejecucion pausada por el usuario.")
    End If
End Sub

Public Sub PausarEjecucion()
    EnEjecucion = False
    Call ModuloUI.AgregarLog("CONTROL", "Pausa solicitada. Esperando comando Step o Run.")
End Sub
