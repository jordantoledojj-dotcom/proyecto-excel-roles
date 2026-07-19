VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlataforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private m_ID_clavaPlataforma As String
Private m_nombrePlataforma As String
Private m_descripcion As String

Public Property Get ID_clavaPlataforma() As String
    ID_clavaPlataforma = m_ID_clavaPlataforma
End Property

Public Property Let ID_clavaPlataforma(value As String)
    m_ID_clavaPlataforma = value
End Property

Public Property Get nombrePlataforma() As String
    nombrePlataforma = m_nombrePlataforma
End Property

Public Property Let nombrePlataforma(value As String)
    m_nombrePlataforma = value
End Property

Public Property Get descripcion() As String
    descripcion = m_descripcion
End Property

Public Property Let descripcion(value As String)
    m_descripcion = value
End Property
