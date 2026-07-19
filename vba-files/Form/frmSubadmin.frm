VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSubadmin 
   Caption         =   "UserForm1"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9216.001
   OleObjectBlob   =   "frmSubadmin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSubadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton3_Click()
    Me.Hide
    frmInicio.Show
End Sub
Private Sub cmbGuardar_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nuevaFila As ListRow

    Set ws = ThisWorkbook.Sheets("Subadministraciones")
    Set tbl = ws.ListObjects("tSubs")
    Set nuevaFila = tbl.ListRows.Add
        With nuevaFila.Range
            .Cells(tbl.ListColumns("Subadministrador").Index).value = cmbID.value
            .Cells(tbl.ListColumns("Nombre").Index).value = cmbNombre.value
            .Cells(tbl.ListColumns("Nombramiento actual").Index).value = cmbNombram.value
            .Cells(tbl.ListColumns("Subadministración").Index).value = cmbSubadministracion.value
        End With
    MsgBox "Registro guardado correctamente.", vbInformation
    
    cmbID.value = ""
    cmbNombre.value = ""
    cmbNombram.value = ""
    cmbSubadministracion.value = ""

    cmbSubadministracion.SetFocus
    End Sub

Private Sub UserForm_Initialize()
    Image1.BackColor = RGB(231, 210, 149)
    'AplicarColor Me
    Me.cmbSubadministracion.SetFocus
    'cmb Subadministracion
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow

    Set ws = ThisWorkbook.Sheets("Subadministraciones")
    Set tbl = ws.ListObjects("tSubs")

    cmbSubadministracion.Clear

    For Each fila In tbl.ListRows
        cmbSubadministracion.AddItem fila.Range(tbl.ListColumns("Nombre").Index)
    Next fila
    
    'cmb ID Subadministrador
     Dim ds As Worksheet
    Dim dato As ListObject
    Dim columna As ListRow

    Set ds = ThisWorkbook.Sheets("Subadministraciones")
    Set dato = ds.ListObjects("tSubs")

    cmbID.Clear

    For Each columna In dato.ListRows
        cmbID.AddItem columna.Range(dato.ListColumns("Subadministrador").Index).value
    Next columna
End Sub
Private Sub cmbSubadministracion_Change()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow

    Set ws = ThisWorkbook.Sheets("Subadministraciones")
    Set tbl = ws.ListObjects("tSubs")
    For Each fila In tbl.ListRows
        If fila.Range(tbl.ListColumns("Nombre").Index).value = cmbSubadministracion.value Then
            cmbID.value = fila.Range(tbl.ListColumns("Subadministrador").Index).value

            cmbNombre.value = _
                fila.Range(tbl.ListColumns("Nombre").Index + 2).value & " " & _
                fila.Range(tbl.ListColumns("A. Paterno").Index).value & " " & _
                fila.Range(tbl.ListColumns("A. Materno").Index).value

            cmbNombram.value = _
                fila.Range(tbl.ListColumns("Nombramiento actual").Index).value
            Exit For
        End If
    Next fila
End Sub

