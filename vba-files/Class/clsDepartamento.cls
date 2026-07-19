VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDepartamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_nombre As String
Private m_subadministracion As String

Public Property Get Nombre() As String
    Nombre = m_nombre
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    m_nombre = vNewValue
End Property

Public Property Get Subadministracion() As String
    Subadministracion = m_subadministracion
End Property

Public Property Let Subadministracion(ByVal vNewValue As String)
    m_subadministracion = vNewValue
End Property
