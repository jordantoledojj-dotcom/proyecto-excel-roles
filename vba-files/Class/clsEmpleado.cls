VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 1. ATRIBUTOS PRIVADOS (Llevan validación interna)
Private p_rfc_corto As String
Private p_RFC As String

' 2. ATRIBUTOS PÚBLICOS (No necesitan validación, VBA crea su Get/Set automático)
Public Nombre As String
Public aPaterno As String
Public aMaterno As String
Public nombramiento As String
Public puesto As String
Public subadm As String
Public fNacimiento As Date
Public Departamento As String
Public jefeDirecto As String
Public estatus As String
Public fIngreso As Date
Public fInicioCargoActual As Date
Public extension As String


' =================================================================
' ENCAPSULAMIENTO PARA LOS ATRIBUTOS PRIVADOS
' =================================================================
' --- PROPIEDAD RFC ---
Public Property Let rfc(ByVal vNewValue As String)
    p_RFC = UCase(Trim(vNewValue))
    
    ' Si el usuario escribe primero el RFC largo y ya había algo en rfc_corto,
    ' podemos forzar la validación de inmediato.
    If p_rfc_corto <> "" Then
        ValidarCoherenciaRfc
    End If
End Property

Public Property Get rfc() As String
    rfc = p_RFC
End Property


' --- PROPIEDAD RFC_CORTO algosss---
Public Property Let RFC_corto(ByVal vNewValue As String)
    p_rfc_corto = UCase(Trim(vNewValue))
    
    ' Si el RFC largo ya existe, validamos la coherencia
    If p_RFC <> "" Then
        ValidarCoherenciaRfc
    End If
End Property

Public Property Get RFC_corto() As String
    RFC_corto = p_rfc_corto
End Property


' --- MÉTODO PRIVADO DE VALIDACIÓN
Private Sub ValidarCoherenciaRfc()
    Dim rfcCortoCalculado As String
    
    rfcCortoCalculado = CalcularRfcCorto(p_RFC)
    
    ' Si hay algún error, se límpia la variable y se manda mensaje de error
    If p_rfc_corto <> rfcCortoCalculado Then
        Dim errorMsg As String
        errorMsg = "Error de Coherencia: El RFC corto '" & p_rfc_corto & _
                   "' no coincide con el algoritmo esperado para el RFC '" & p_RFC & _
                   "' (Se esperaba: '" & rfcCortoCalculado & "')."
                   
        p_rfc_corto = "" ' Reseteamos el valor por seguridad
        
        ' Lanzamos la excepción al Controlador
        Err.Raise vbObjectError + 514, "clsEmpleado", errorMsg
    End If
End Sub




