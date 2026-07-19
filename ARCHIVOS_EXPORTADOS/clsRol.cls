VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pkClaveRol As String
Private m_rolFuncional As String
Private m_rolTecnico As String
Private m_definicionAlcance As String
Private m_fkClavePlataforma As String

Public Property Get pkClaveRol() As String
    pkClaveRol = m_pkClaveRol
End Property

Public Property Let pkClaveRol(ByVal vNewValue As String)
    m_pkClaveRol = vNewValue
End Property

Public Property Get rolFuncional() As String
    rolFuncional = m_rolFuncional
End Property

Public Property Let rolFuncional(ByVal vNewValue As String)
    m_rolFuncional = vNewValue
End Property

Public Property Get rolTecnico() As String
    rolTecnico = m_rolTecnico
End Property

Public Property Let rolTecnico(ByVal vNewValue As String)
    m_rolTecnico = vNewValue
End Property

Public Property Get definicionAlcance() As String
    definicionAlcance = m_definicionAlcance
End Property

Public Property Let definicionAlcance(ByVal vNewValue As String)
    m_definicionAlcance = vNewValue
End Property

Public Property Get fkClavePlataforma() As String
    fkClavePlataforma = m_fkClavePlataforma
End Property

Public Property Let fkClavePlataforma(ByVal vNewValue As String)
    m_fkClavePlataforma = vNewValue
End Property
