VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSubadministracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Atributos privados
Private m_nombre As String
Private m_subadministrador As String ' Almacena el RFC_corto del jefe de la subadministración

' Propiedad Nombre (Actúa como el ID/Llave de la entidad)
Public Property Get Nombre() As String
    Nombre = m_nombre
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    m_nombre = vNewValue
End Property

' Propiedad Subadministrador (RFC_corto del administrador a cargo)
Public Property Get subadministrador() As String
    subadministrador = m_subadministrador
End Property

Public Property Let subadministrador(ByVal vNewValue As String)
    m_subadministrador = vNewValue
End Property
