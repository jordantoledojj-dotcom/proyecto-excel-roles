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
Private m_fechaAlta As String
Private m_fechaBaja As String
Private m_solicitante As String
Private m_notas As String

Public Property Get pkRFC_corto()
    pkRFC_corto = m_pkRFC_corto
End Property

Public Property Let pkRFC_corto(ByVal vNewValue As String)
    m_pkRFC_corto = vNewValue
End Property

Public Property Get pkClaveRolTecnico()
    pkClaveRolTecnico = m_pkClaveRolTecnico
End Property

Public Property Let pkClaveRolTecnico(ByVal vNewValue As String)
    m_pkClaveRolTecnico = vNewValue
End Property

Public Property Get fechaAlta()
    fechaAlta = m_fechaAlta
End Property

Public Property Let fechaAlta(ByVal vNewValue As String)
    m_fechaAlta = vNewValue
End Property

Public Property Get fechaBaja()
    fechaBaja = m_fechaBaja
End Property

Public Property Let fechaBaja(ByVal vNewValue As String)
    m_fechaBaja = vNewValue
End Property


Public Property Get solicitante()
    solicitante = m_solicitante
End Property

Public Property Let solicitante(ByVal vNewValue As String)
    m_solicitante = vNewValue
End Property

Public Property Get notas()
    notas = m_notas
End Property

Public Property Let notas(ByVal vNewValue As String)
    m_notas = vNewValue
End Property
