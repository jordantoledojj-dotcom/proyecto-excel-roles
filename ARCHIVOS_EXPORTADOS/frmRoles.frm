VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRoles 
   Caption         =   "Control de Roles"
   ClientHeight    =   8550.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   14760
   OleObjectBlob   =   "frmRoles.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbRfc_Change()
    Dim empl As Worksheet
    Dim valor As Range
        Set empl = Sheets("Empleados")
        Set valor = empl.Columns("A").Find(cmbRfc.value, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            txtNombre.value = empl.Cells(valor.Row, 2).value & " " & _
                      empl.Cells(valor.Row, 3).value & " " & _
                      empl.Cells(valor.Row, 4).value
            cmbSubadministracion.value = empl.Cells(valor.Row, 9).value
        End If
End Sub
Private Sub CommandButton1_Click()
        Me.Hide
        frmAsig_rol.Show
End Sub
Private Sub UserForm_Initialize()
    Me.cmbRfc.SetFocus
        Dim empl As Worksheet
        Dim tbl As ListObject
        Dim celda As Range
            Set empl = ThisWorkbook.Sheets("Empleados")
            Set tbl = empl.ListObjects("tblEmpleado")
            cmbRfc.Clear
        For Each celda In tbl.ListColumns("RFC Corto").DataBodyRange
                If celda.value <> "" Then
                    cmbRfc.AddItem celda.value
                End If
        Next celda
        
        'clave
        Dim rol As Worksheet
        Dim clave As ListObject
        Dim tec As Range
            Set rol = ThisWorkbook.Sheets("Roles")
            Set clave = rol.ListObjects("tblRoles")
            cmbClave.Clear
        For Each tec In clave.ListColumns("CLAVE DEL SERVICIO").DataBodyRange
                If tec.value <> "" Then
                    cmbClave.AddItem tec.value
                End If
        Next tec
End Sub
