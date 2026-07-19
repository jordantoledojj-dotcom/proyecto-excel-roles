VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInicio 
   Caption         =   "Inicio"
   ClientHeight    =   7176
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10116
   OleObjectBlob   =   "frmInicio.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnAbrirEmpleado_Click()
    Me.Hide
    frmEmpleado.Show
End Sub

