Attribute VB_Name = "ModuloControladores"
Option Explicit

' ==============================================================================
' Controladores de Botones de la Interfaz (Eventos de Usuario)
' ==============================================================================

' Boton: Paso a Paso (STEP)
Public Sub btn_Step()
    Call ModuloCiclo.PasoCiclo
End Sub

' Boton: Ejecucion Continua (RUN)
Public Sub btn_Run()
    Call ModuloCiclo.EjecutarContinuo
End Sub

' Boton: Pausar Ejecucion (PAUSE)
Public Sub btn_Pause()
    Call ModuloCiclo.PausarEjecucion
End Sub

' Boton: Reiniciar CPU (RESET)
Public Sub btn_Reset()
    Call ModuloCPU.ResetCPU
End Sub

' Boton: Cargar Programa Demostrativo (LOAD PROGRAM)
Public Sub btn_LoadDemo()
    Call ModuloCPU.ResetCPU
    Call ModuloMemoria.CargarProgramaDemostrativo
End Sub
