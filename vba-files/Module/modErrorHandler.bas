Attribute VB_Name = "modErrorHandler"
Option Explicit

' Códigos de error personalizados del sistema
Public Const ERR_CLASE_COHERENCIA As Long = vbObjectError + 514

Public Sub ManejarErrorCritico(ByVal numeroError As Long, ByVal fuente As String, ByVal descripcion As String, Optional ByVal filaExcel As Long = 0)
    ' Invalida la caché global para asegurar que no queden datos corruptos en memoria
    Call modEmp_Rol.invalidarCache
    
    Dim mensaje As String
    
    Select Case numeroError
        Case ERR_CLASE_COHERENCIA
            mensaje = "PROCESO ABORTADO POR CORRUPCIÓN DE DATOS" & vbCrLf & vbCrLf & _
                      "Origen: " & fuente & vbCrLf & _
                      "Detalle: " & descripcion
                      
            If filaExcel > 0 Then
                mensaje = mensaje & vbCrLf & "Fila afectada en Excel: " & filaExcel
            End If
            
        Case CaseElse ' Para errores nativos de VBA (Error 9, 13, etc.)
            mensaje = "ERROR INESPERADO EN EL SISTEMA" & vbCrLf & vbCrLf & _
                      "Código de Error: " & numeroError & vbCrLf & _
                      "Origen: " & fuente & vbCrLf & _
                      "Descripción: " & descripcion
    End Select
    
    ' Muestra el panel con el error crítico
    MsgBox mensaje, vbCritical, "Error de Integridad del Sistema"
    
    ' Detiene por completo la ejecución de VBA (Fail-Fast)
    End
End Sub


