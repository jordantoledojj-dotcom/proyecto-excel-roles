VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAsig_rol 
   Caption         =   "Roles-Personal"
   ClientHeight    =   9795.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12540
   OleObjectBlob   =   "frmAsig_rol.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmAsig_rol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnF_Solicitud_Click()
  Dim fechaTmp As Date
    fechaTmp = CalendarForm.GetDate(, , "01/01/1950", Date, 80)
    txtF_Solicitud.value = IIf(fechaTmp = 0, "", Format(fechaTmp, "dd/mm/yyyy"))
End Sub
Private Sub btnInicio_Click()
    Me.Hide
    frmInicio.Show
End Sub
Private Sub cmbRfc_Change()
    cargarDatosEmpleado
End Sub


Private Sub UserForm_Initialize()
    AplicarColor
    CargarDatosRFC_corto
End Sub

Private Sub CargarDatosRFC_corto()
    Dim listaRFC As Variant
    
    cmbRfc.Clear
    listaRFC = modAsignarRolController.obtenerListaRFC_corto
    If IsEmpty(listaRFC) Then Exit Sub
    Me.cmbRfc.List = listaRFC
End Sub

Private Sub AplicarColor()
    Dim con As Control
    For Each con In Me.Controls
        Select Case TypeName(con)
            Case "Frame"
                con.BackColor = RGB(30, 91, 79)
            Case "Label"
                con.BackColor = RGB(30, 91, 79)
            Case "TextBox"
                con.BackColor = RGB(255, 255, 255)
            Case "ComboBox"
                con.BackColor = RGB(255, 255, 255)
            Case "OptionButton"
                con.BackColor = RGB(30, 91, 79)
        End Select
    Next con
    
    Me.imgTop.BackColor = modConfig.COLOR_SAM
    Me.imgBottom.BackColor = modConfig.COLOR_SAM
    'tabla empleados
    
    Me.BackColor = modConfig.COLOR_VERDE_SAM
    Me.btnInicio.BackColor = modConfig.COLOR_VERDE_OSCURO
    Me.btnCancelar.BackColor = modConfig.COLOR_VERDE_OSCURO
    Me.btnGuardar.BackColor = modConfig.COLOR_VINO
End Sub

Private Sub cargarDatosEmpleado()
    Dim objEmpleado As clsEmpleado
    
    Set objEmpleado = modAsignarRolController.obtenerEmpleadoByRFC_Corto(cmbRfc.value())
    mostrarDatosEmpleado objEmpleado
End Sub

Public Sub mostrarDatosEmpleado(ByRef objEmpleado As clsEmpleado)
    Me.txtSubadministracion.value = objEmpleado.subadm
    Me.txtNombre.value = objEmpleado.Nombre & " " & objEmpleado.aPaterno & " " & objEmpleado.aMaterno
End Sub




