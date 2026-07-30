VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmpleado 
   Caption         =   "Formulario de capturacion de informacion"
   ClientHeight    =   11265
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11400
   OleObjectBlob   =   "frmEmpleado.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' ===========================================================================================================
' RECORDATORIO: Los formularios deben cumplir con las siguientes características:
' DIBUJAR LA PANTALLA: Mostrar los controles y sus respectivos datos al usuario
' CACHAR EVENTOS: Manipular clics, teclas, etc y pasarlos inmediatamente al controlador
' SOLO ESO, LO DEMÁS (CONEXIONES, VALIDACIONES DE REGLAS DE NEGOCIO, ETC son responsabilidad de alguien más
' ===========================================================================================================

Option Explicit

Private eventoFrm As Integer

' ===========================================================================
' Funciones que el controlador invoca para obtener o mostrar datos del form
' ===========================================================================

' El formulario tiene una propiedad pública para pasar los datos limpios en un objeto
Public Function getEmployeeFromForm() As clsEmpleado
    Dim emp As New clsEmpleado
    
    ' Mapeo: De los controles de la pantalla al Objeto
    emp.RFC_corto = Me.txtRFC_corto.value
    emp.rfc = Me.txtRFC.value
    emp.Nombre = Me.txtNombre.value
    emp.aPaterno = Me.txtApellido_paterno.value
    emp.aMaterno = Me.txtApellido_materno.value
    emp.nombramiento = Me.txtNombramiento_act.value
    emp.puesto = Me.txtPuesto.value
    emp.estatus = Me.txtEstatus.value
    emp.extension = Me.txtNo_extension.value
    
    emp.subadm = Me.cmbSubadministracion.value
    emp.Departamento = Me.cmbDepartamento.value
    emp.jefeDirecto = Me.cmbJefe_directo.value
    
    If IsDate(Me.txtFecha_ingreso.value) Then emp.fIngreso = CDate(Me.txtFecha_ingreso.value)
    If IsDate(Me.txtFecha_inicio_cargo.value) Then emp.fInicioCargoActual = CDate(Me.txtFecha_inicio_cargo.value)
    If IsDate(Me.txtFecha_nacimiento.value) Then emp.fNacimiento = CDate(Me.txtFecha_nacimiento.value)
        
    Set getEmployeeFromForm = emp
End Function

' Método inverso: Pintar los datos de un objeto recuperado en la pantalla
Public Sub showEmployeeOnForm(ByRef emp As clsEmpleado)
    If emp Is Nothing Then Exit Sub
    
    Me.txtRFC_corto.value = emp.RFC_corto
    Me.txtRFC.value = emp.rfc
    Me.txtNombre.value = emp.Nombre
    Me.txtApellido_paterno.value = emp.aPaterno
    Me.txtApellido_materno.value = emp.aMaterno
    Me.txtNombramiento_act = emp.nombramiento
    Me.txtPuesto.value = emp.puesto
    Me.txtEstatus.value = emp.estatus
    Me.txtNo_extension.value = emp.extension
    
    Me.cmbSubadministracion.value = emp.subadm
    Me.cmbDepartamento.value = emp.Departamento
    Me.cmbJefe_directo.value = emp.jefeDirecto
    
    Me.txtFecha_ingreso.value = Format(emp.fIngreso, "dd/mm/yyyy")
    Me.txtFecha_inicio_cargo.value = Format(emp.fInicioCargoActual, "dd/mm/yyyy")
    Me.txtFecha_nacimiento.value = Format(emp.fNacimiento, "dd/mm/yyyy")
End Sub

Public Sub LlenarListaEmpleados(ByRef matrizDatosEmpleados As Variant)
    Dim emp As clsEmpleado
    
    ' 1. Limpiar la lista por si ya tenía datos
    Me.lstEmpleados.Clear
    
    ' Si la colección está vacía, terminamos temprano
    If isEmpty(matrizDatosEmpleados) Then Exit Sub
    
    ' 2. Configurar las propiedades visuales del ListBox
    With Me.lstEmpleados
        .ColumnCount = 5
        .ColumnWidths = "70 pt; 90 pt; 160 pt;120 pt;90 pt" ' Ancho de cada columna en puntos
        .ColumnHeads = False
        .List = matrizDatosEmpleados
    End With
    
End Sub

' =================================================================
' EVENTOS DE LA INTERFAZ (Gritan al Controlador)
' =================================================================


Private Sub txtBusqueda_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Detectamos si la tecla presionada fue Enter
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    ' Evitamos el "beep" molesto que hace Excel por defecto al pulsar Enter en un TextBox de una sola línea
    KeyCode = 0
    cmdBuscar_Click
End Sub

Private Sub cmdBuscar_Click()
    Dim strBuscar As String
    strBuscar = Trim(Me.txtBusqueda.value)
    Call modEmpleadoController.EjecutarBusquedaGlobal(Me, strBuscar)
End Sub

Private Sub lstEmpleados_Click()
    Dim strRFC_corto As String
    If lstEmpleados.ListIndex = -1 Then
        MsgBox "Selecciona un registro"
        Exit Sub
    End If
    strRFC_corto = lstEmpleados.List(lstEmpleados.ListIndex, 0)
    Call modEmpleadoController.EjecutarBuscar(Me, strRFC_corto)
End Sub

Private Sub cmbSubadministracion_Change()
    cargarComboBoxDepartamento
End Sub

Private Sub inicializaLista()
    Call modEmpleadoController.CargarCatalogoEmpleados(Me)
End Sub

Private Sub inicializaComboBoxes()
    cargarComboBoxSubadministracion
    cargarComboBoxJefeDirecto
End Sub

Private Sub cargarComboBoxDepartamento()
    Me.cmbDepartamento.Clear
    Dim listaNombresDeptos As Variant
    listaNombresDeptos = modEmpleadoController.cargarCatalogoDepartamentoXSubs(Me.cmbSubadministracion.value)
    If isEmpty(listaNombresDeptos) Then Exit Sub
    cmbDepartamento.List = listaNombresDeptos
End Sub

Private Sub cargarComboBoxSubadministracion()
    Me.cmbSubadministracion.Clear
    Dim listaNombresSub As Variant
    listaNombresSub = modEmpleadoController.CargarCatalogoSubadministraciones
    If isEmpty(listaNombresSub) Then Exit Sub
    Me.cmbSubadministracion.List = listaNombresSub
End Sub

Private Sub cargarComboBoxJefeDirecto()
    Me.cmbJefe_directo.Clear
    Dim matrizDatosJefes As Variant
    matrizDatosJefes = modEmpleadoController.ObtenerMatrizJefes
    If isEmpty(matrizDatosJefes) Then Exit Sub
    With Me.cmbJefe_directo
        .ColumnCount = 2
        .BoundColumn = 1
        .TextColumn = 2
        .ColumnWidths = "55pt;150pt"
        .List = matrizDatosJefes
    End With
End Sub


' ===========================================================================================================
' Funciones fundamentales AGREGAR, MODIFICAR, ELIMINAR. No llaman al controlador
' ===========================================================================================================

Private Sub btnAgregar_Click()
    eventoFrm = modConfig.EVENTO_NUEVO_REGISTRO
    visibilidadBotoneraSuperior False
    visibilidadBotoneraInferior True
    habilitarParteSuperior False
    habilitarFrameEmpleado True
    limpiarFormularioEmpleado
    ' Establecer el focus es lo último que se debe hacer
    establecerFocustxtRFC
End Sub

Private Sub btnEditar_Click()
    If Me.lstEmpleados.ListIndex = -1 Then
        MsgBox "Selecciona un registro para Editar", vbExclamation, "Sin registro seleccionado"
        Exit Sub
    End If
    eventoFrm = modConfig.EVENTO_ACTUALIZAR_REGISTRO
    
    visibilidadBotoneraSuperior False
    visibilidadBotoneraInferior True
    habilitarParteSuperior False
    habilitarFrameEmpleado True
    With txtRFC
        .Locked = True
        .ForeColor = modConfig.COLOR_FUENTE_DESHABILITADA
        .BackColor = modConfig.COLOR_DESHABILITADO
    End With
    With txtRFC_corto
        .Locked = True
        .ForeColor = modConfig.COLOR_FUENTE_DESHABILITADA
        .BackColor = modConfig.COLOR_DESHABILITADO
    End With
    
End Sub

Private Sub btnEliminar_Click()
    If Me.lstEmpleados.ListIndex = -1 Then
        MsgBox "Selecciona un registro para Editar", vbExclamation, "Sin registro seleccionado"
        Exit Sub
    End If
    eventoFrm = modConfig.EVENTO_ELIMINAR_REGISTRO
    
    visibilidadBotoneraSuperior False
    visibilidadBotoneraInferior True
    habilitarParteSuperior False
    habilitarFrameEmpleado False

End Sub

' ===========================================================================================================
' Funciones privadas del formulario
' ===========================================================================================================


Private Sub btnCancelar_Click()
    inicializarFormularioEmpleado
End Sub

' Evento de Autorelleno: Cuando el usuario termina de escribir el RFC largo
Private Sub txtRFC_AfterUpdate()
    ' Usamos nuestra función utilitaria para rellenar el RFC corto automáticamente
    If Me.txtRFC.value <> "" Then
        Me.txtRFC_corto.value = modUtilerias.CalcularRfcCorto(Me.txtRFC.Text)
    End If
End Sub

Private Sub btnConfirmar_Click()
    ' Un botón para las funciones AGREGAR, ACTUALIZAR y ELIMINAR
    Select Case eventoFrm
        Case modConfig.EVENTO_ACTUALIZAR_REGISTRO
            Call modEmpleadoController.EjecutarModificarEmpleado(Me)
        Case modConfig.EVENTO_NUEVO_REGISTRO
            Call modEmpleadoController.ejecutarGuardarEmpleado(Me)
        Case modConfig.EVENTO_ELIMINAR_REGISTRO
            Call modEmpleadoController.EjecutarEliminarEmpleado(Me, txtRFC_corto)
        Case Else
    End Select
End Sub

Private Sub btnInicio_Click()
    Me.Hide
    frmInicio.Show
End Sub

Private Sub establecerFocustxtRFC()
    Me.txtRFC.SetFocus
End Sub

Private Sub limpiarFormularioEmpleado()
    Dim ctrl As Control
    
    ' Recorre absolutamente todos los objetos que están dentro del formulario
    For Each ctrl In Me.Controls
        ' Si es un cuadro de texto (TextBox) o cuadro combinado (ComboBox), lo limpia
        If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Then
            ctrl.value = ""
        End If
    Next ctrl
End Sub

Private Sub habilitarParteSuperior(valorBoleano As Boolean)
    lstEmpleados.Enabled = valorBoleano
    txtBusqueda.Enabled = valorBoleano
    btnBuscar.Enabled = valorBoleano
    lstEmpleados.BackColor = IIf(valorBoleano, modConfig.COLOR_HABILITADO, modConfig.COLOR_DESHABILITADO)
    txtBusqueda.BackColor = IIf(valorBoleano, modConfig.COLOR_HABILITADO, modConfig.COLOR_DESHABILITADO)
    btnBuscar.BackColor = IIf(valorBoleano, modConfig.COLOR_VINO, modConfig.COLOR_DESHABILITADO)
End Sub

Private Sub habilitarFrameEmpleado(valorBoleano As Boolean)
    Dim ctrl As Control
    
    ' Recorre absolutamente todos los objetos que están dentro del formulario
    For Each ctrl In Me.frameEmpleado.Controls
        ' Si es un cuadro de texto (TextBox) o cuadro combinado (ComboBox), lo habilita o deshabilita, según el valor booleano
        If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Or TypeOf ctrl Is MSForms.CommandButton Then
            ctrl.Locked = Not valorBoleano
            ctrl.ForeColor = IIf(valorBoleano, modConfig.COLOR_FUENTE_HABILITADA, modConfig.COLOR_FUENTE_DESHABILITADA)
            ctrl.BackColor = IIf(valorBoleano, modConfig.COLOR_HABILITADO, modConfig.COLOR_DESHABILITADO)
        End If
    Next ctrl
End Sub

Private Sub visibilidadBotoneraInferior(valorBoleano As Boolean)
    Me.btnConfirmar.Visible = valorBoleano
End Sub

Private Sub visibilidadBotoneraSuperior(valorBoleano As Boolean)
    Me.btnEditar.Visible = valorBoleano
    Me.btnAgregar.Visible = valorBoleano
    Me.btnEliminar.Visible = valorBoleano
End Sub

Private Sub colorearFormulario()
    Me.BackColor = modConfig.COLOR_FONDO_FORM
    frameEmpleado.BackColor = modConfig.COLOR_FONDO_FORM
    frameEmpleado.ForeColor = modConfig.COLOR_ARENA
    lbTitulo.BackColor = modConfig.COLOR_FONDO_FORM
    
    imgLogoSAT.BackColor = modConfig.COLOR_VINO
    lbTitulo.ForeColor = modConfig.COLOR_BLANCO
    btnInicio.BackColor = modConfig.COLOR_VERDE_OSCURO
    
    btnBuscar.BackColor = modConfig.COLOR_VERDE_OSCURO
    
    btnEliminar.BackColor = modConfig.COLOR_VINO
    btnAgregar.BackColor = modConfig.COLOR_VERDE_OSCURO
    btnEditar.BackColor = modConfig.COLOR_ARENA
    
    btnConfirmar.BackColor = modConfig.COLOR_VINO
    btnCancelar.BackColor = modConfig.COLOR_ARENA
End Sub

Private Sub establecerFocusBuscar()
    Me.txtBusqueda.SetFocus
End Sub

Private Sub btnFIngreso_Click()
    Dim fechaTmp As Date
    fechaTmp = CalendarForm.GetDate(, , "01/01/1950", Date, 80)
    txtFecha_ingreso.value = IIf(fechaTmp = 0, "", Format(fechaTmp, "dd/mm/yyyy"))
End Sub

Private Sub btnFinicio_Click()
    Dim fechaTmp As Date
    fechaTmp = CalendarForm.GetDate(, , "01/01/1950", Date, 80)
    txtFecha_inicio_cargo.value = IIf(fechaTmp = 0, "", Format(fechaTmp, "dd/mm/yyyy"))
End Sub

Private Sub btnFNacimiento_Click()
    Dim fechaTmp As Date
    fechaTmp = CalendarForm.GetDate(, , "01/01/1950", Date, 80)
    txtFecha_nacimiento = IIf(fechaTmp = 0, "", Format(fechaTmp, "dd/mm/yyyy"))
End Sub

Private Sub btnBuscar_Click()
    cmdBuscar_Click
End Sub

Public Sub inicializarFormularioEmpleado()
    eventoFrm = modConfig.EVENTO_SIN_ACCION
    limpiarFormularioEmpleado
    habilitarParteSuperior True
    habilitarFrameEmpleado False
    inicializaLista
    inicializaComboBoxes
    visibilidadBotoneraSuperior True
    visibilidadBotoneraInferior False
    establecerFocusBuscar
End Sub

Private Sub UserForm_Initialize()
    colorearFormulario
    inicializarFormularioEmpleado
End Sub

Private Sub UserForm_Activate()
    inicializarFormularioEmpleado
End Sub












' Private Sub btnPruebas_Click()
'     Me.txtRFC_corto.value = "VATJ8712"
'     Me.txtRFC.value = "VATJ870102XXX"
'     Me.txtNombre.value = "Fulano"
'     Me.txtApellido_paterno.value = "Mengano"
'     Me.txtApellido_materno.value = "Pancho"
'     Me.txtPuesto.value = "Jefe"
'     Me.txtEstatus.value = "activo"
'     Me.txtNo_extension.value = "6378545"
'     Me.txtNombramiento_act = "algo"
'
'     Me.cmbSubadministracion.value = "Administración"
'     Me.cmbDepartamento.value = "Administración"
'     Me.cmbJefe_directo.value = "LOBK735K"
'
'     Me.txtFecha_ingreso.value = Format("20/10/2024", "dd/mm/yyyy")
'     Me.txtFecha_inicio_cargo.value = Format("20/10/2025", "dd/mm/yyyy")
'     Me.txtFecha_nacimiento.value = Format("20/10/2026", "dd/mm/yyyy")
' End Sub
