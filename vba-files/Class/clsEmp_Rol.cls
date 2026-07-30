VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmp_Rol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pkRFC_corto As String
Private m_pkClaveRolTecnico As String
Private m_fechaAlta As Date
Private m_fechaBaja As Date
Private m_solicitante As String
Private m_notas As String

Public Property Get pkRFC_corto() As String
    pkRFC_corto = m_pkRFC_corto
End Property

Public Property Let pkRFC_corto(ByVal vNewValue As String)
    m_pkRFC_corto = vNewValue
End Property

Public Property Get pkClaveRolTecnico() As String
    pkClaveRolTecnico = m_pkClaveRolTecnico
End Property

Public Property Let pkClaveRolTecnico(ByVal vNewValue As String)
    m_pkClaveRolTecnico = vNewValue
End Property

Public Property Get fechaAlta() As Date
    fechaAlta = m_fechaAlta
End Property

Public Property Let fechaAlta(ByVal vNewValue As Variant)
    m_fechaAlta = 0
    If IsDate(vNewValue) Then m_fechaAlta = vNewValue
End Property

Public Property Get fechaBaja() As Date
    fechaBaja = m_fechaBaja
End Property

Public Property Let fechaBaja(ByVal vNewValue As Variant)
    m_fechaBaja = 0
    If IsDate(vNewValue) Then m_fechaBaja = vNewValue
End Property

Public Property Get solicitante() As String
    solicitante = m_solicitante
End Property

Public Property Let solicitante(ByVal vNewValue As String)
    m_solicitante = vNewValue
End Property

Public Property Get notas() As String
    notas = m_notas
End Property

Public Property Let notas(ByVal vNewValue As String)
    m_notas = vNewValue
End Property

'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////Propiedades calculadas////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Property Get esRolActivo() As Boolean
    ' Un rol es activo si no tiene registrada una fecha de baja (es decir, vale 0)
    esRolActivo = (m_fechaBaja = 0)
End Property

Public Property Get pkID() As String
    ' validamos los valores necesarios para el ID
    hayDatosParaID
    ' Devolvemos la concatenación de ambas llaves compuestas
    pkID = m_pkRFC_corto & "-" & m_pkClaveRolTecnico & "-" & Format(m_fechaAlta, "yyyymmdd")
End Property


'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////Propiedades privadas de validación//////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub hayDatosParaID()
    
    If m_pkRFC_corto = "" Or m_pkClaveRolTecnico = "" Or m_fechaAlta = 0 Then
        Dim errorMsg As String
        errorMsg = "Error de Coherencia: No se puede generar pkID en clsEmp_Rol porque faltan elementos clave " & _
                   "(pkRFC_corto, pkClaveRolTecnico o fechaAlta están vacíos)."
        
        ' Lanzamos la excepción al Controlador
        Err.Raise vbObjectError + 514, "clsEmp_Rol", errorMsg
    End If
End Sub

