Attribute VB_Name = "modEmpleadoController"
Option Explicit

Public Sub ejecutarGuardarEmpleado(ByRef vista As frmEmpleado)
    Dim nuevoEmpleado As clsEmpleado
    Dim mensajeError As String
    Dim exito As Boolean
    
    ' 1. El controlador le pide a la vista el objeto armado
    ' (Si el objeto detecta inconsistencias en el RFC al armarse, saltará el error aquí)
    On Error GoTo CatchError
    Set nuevoEmpleado = vista.getEmployeeFromForm()
    On Error GoTo 0
    
    ' 2. Validación de controlador (Campos vacíos obligatorios)
    If nuevoEmpleado.rfc = "" Or nuevoEmpleado.Nombre = "" Then
        MsgBox "Por favor, llene los campos obligatorios (RFC y Nombre).", vbExclamation, "Validación"
        Exit Sub
    End If
    
    ' 3. El controlador da la orden al DAO de persistir el objeto en Excel
    exito = modEmpleadoDAO.Insertar(nuevoEmpleado, mensajeError)
    
    ' 4. El controlador decide qué responderle al usuario en la pantalla
    If exito Then
        MsgBox "¡Empleado registrado con éxito!", vbInformation, "Éxito"
        vista.inicializarFormularioEmpleado
    Else
        MsgBox "Se generó un problema al intentar guardar: " & mensajeError, vbCritical, "Error de Persistencia"
    End If
    
    Exit Sub

CatchError:
    ' Captura los errores de lógica del objeto (como el RFC corto que no coincide)
    MsgBox Err.Description, vbCritical, "Error de Negocio"
End Sub


' =================================================================
' Muestra el catálogo de todos los empleados en el formato que utilizará la vista (matriz
' Puede parecer muy rebuscado pero es la forma de respetar el MVC
' =================================================================

Public Sub CargarCatalogoEmpleados(ByRef vista As frmEmpleado)
    Dim coleccionEmpleados As Collection
    Dim matrizDatosEmpleados As Variant
    
    ' 1. El controlador va al DAO por los datos
    Set coleccionEmpleados = modEmpleadoDAO.getAllEmployees()
    
    matrizDatosEmpleados = ObtenerMatrizEmpleados(coleccionEmpleados)
    
    ' 2. El controlador le envía los datos directamente a la Vista para que se pinte
    Call vista.LlenarListaEmpleados(matrizDatosEmpleados)
End Sub

' =================================================================
' BUSQUEDA: Orquesta la consulta de un empleado por su RFC corto
' =================================================================


Public Sub EjecutarBuscar(ByRef vista As frmEmpleado, ByVal strRFC_corto As String)
    Dim empleadoEncontrado As clsEmpleado
    
    strRFC_corto = UCase(Trim(strRFC_corto))
    
    ' Validación rápida de entrada en el controlador
    If strRFC_corto = "" Then
        MsgBox "Por favor, ingrese un identificador válido para buscar.", vbExclamation, "Búsqueda"
        Exit Sub
    End If
    
    ' 1. El controlador invoca al DAO para ir a la tabla de Excel
    Set empleadoEncontrado = modEmpleadoDAO.ObtenerPorId(strRFC_corto)
    
    ' 2. El controlador evalúa el resultado y actúa sobre la Vista
    If Not empleadoEncontrado Is Nothing Then
        ' Si lo encuentra, le ordena a la vista pintarlo
        Call vista.showEmployeeOnForm(empleadoEncontrado)
    Else
        MsgBox "No se encontró ningún empleado con el identificador: " & strRFC_corto, _
               vbInformation, "Búsqueda Sin Resultados"
    End If
End Sub

' =================================================================
' BUSQUEDA GLOBAL: Filtra la lista de empleados por cualquier término
' =================================================================
Public Sub EjecutarBusquedaGlobal(ByRef vista As frmEmpleado, ByVal termino As String)
    Dim empleadosFiltrados As Collection
    
    termino = Trim(termino)
    
    ' Si el usuario borra el cuadro de búsqueda, entendemos que quiere ver a todos otra vez
    If termino = "" Then
        Call CargarCatalogoEmpleados(vista)
        Exit Sub
    End If
    
    ' El controlador va al DAO por los datos filtrados
    Set empleadosFiltrados = modEmpleadoDAO.obtenerEmpleadosPorBusqueda(termino)
    
    ' Transformamos la colección de empleados filtrados en una matriz lista para enviar al formulario
    Dim matrizEmpleadosFiltrados As Variant
    matrizEmpleadosFiltrados = ObtenerMatrizEmpleados(empleadosFiltrados)
    
    ' Le mandamos la colección resultante a la vista para refrescar el ListBox/ListView
    Call vista.LlenarListaEmpleados(matrizEmpleadosFiltrados)
    
    ' Avisar si no hubo coincidencias sin romper el flujo
    ' Podría ser antes del punto anterior, pero la función .LlenarListaEmpleados(variant) se encarga de pintar la lista
    ' cuando esta no tenga elementos, así que no pasa nada
    If empleadosFiltrados.Count = 0 Then
        MsgBox "No se encontraron coincidencias para: '" & termino & "'", vbInformation, "Búsqueda Sin Coincidencias"
        Call vista.inicializarFormularioEmpleado
    End If
End Sub

' =================================================================
' MODIFICAR: Orquesta la actualización de un empleado
' =================================================================
Public Sub EjecutarModificarEmpleado(ByRef vista As frmEmpleado)
    Dim empleadoModificado As clsEmpleado
    Dim mensajeError As String
    Dim exito As Boolean
    
    On Error GoTo CatchError
    ' 1. Le pedimos a la vista el objeto con los datos modificados del formulario
    Set empleadoModificado = vista.getEmployeeFromForm()
    On Error GoTo 0
    
    ' 2. Validación rápida a nivel controlador
    If empleadoModificado.RFC_corto = "" Then
        MsgBox "No se puede modificar un registro sin un RFC Corto (ID) válido.", vbExclamation, "Validación"
        Exit Sub
    End If
    
    If empleadoModificado.rfc = "" Or empleadoModificado.Nombre = "" Then
        MsgBox "Los campos RFC y Nombre son obligatorios.", vbExclamation, "Validación"
        Exit Sub
    End If
    
    ' 3. Ordenamos al DAO que actualice en Excel
    exito = modEmpleadoDAO.Actualizar(empleadoModificado, mensajeError)
    
    ' 4. Respondemos a la interfaz de usuario
    If exito Then
        MsgBox "¡Datos del empleado actualizados con éxito!", vbInformation, "Éxito"
        
        ' Refrescamos la lista general para ver los cambios reflejados de inmediato
        vista.inicializarFormularioEmpleado
    Else
        MsgBox "No se pudo actualizar: " & mensajeError, vbCritical, "Error de Persistencia"
    End If
    
    Exit Sub

CatchError:
    MsgBox "Error de lógica de negocio al modificar: " & Err.Description, vbCritical, "Error"
End Sub

' =================================================================
' ELIMINAR: Orquesta la baja de un empleado previa confirmación
' =================================================================
Public Sub EjecutarEliminarEmpleado(ByRef vista As frmEmpleado, ByVal strRFC_corto_eliminar As String)
    Dim mensajeError As String
    Dim exito As Boolean
    Dim respuesta As VbMsgBoxResult
    
    strRFC_corto_eliminar = UCase(Trim(strRFC_corto_eliminar))
    
    ' 1. Validación de seguridad en la entrada
    If strRFC_corto_eliminar = "" Then
        MsgBox "Seleccione o busque un empleado antes de intentar eliminarlo.", vbExclamation, "Búsqueda"
        Exit Sub
    End If
    
    ' 2. Regla de negocio crítica: Doble confirmación por seguridad del usuario
    respuesta = MsgBox("¿Está completamente seguro de que desea ELIMINAR permanentemente al empleado con identificador '" _
        & strRFC_corto_eliminar & "'?" & vbCrLf & "Esta acción borrará el registro de la base de datos de Excel y no se puede deshacer.", _
        vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Baja Permanente")

                       
    If respuesta = vbNo Then Exit Sub ' Cancelar operación
    
    ' 3. El controlador da la orden de ejecución al DAO
    exito = modEmpleadoDAO.Eliminar(strRFC_corto_eliminar, mensajeError)
    
    ' 4. Decisión de respuesta en pantalla
    If exito Then
        MsgBox "El empleado ha sido removido del sistema con éxito.", vbInformation, "Registro Eliminado"
        
        ' Limpiamos los campos del formulario y refrescamos la lista
        vista.inicializarFormularioEmpleado
    Else
        MsgBox "No se pudo completar la baja: " & mensajeError, vbCritical, "Error de Persistencia"
    End If
End Sub

' Usado para recuperar una matriz de 1 columna (lista) para
' establecer los valores del comboBox
' Puede parecer muy elaborado, pero es para respetar el MVC
Public Function CargarCatalogoSubadministraciones() As Variant
    Dim listaSubs As Collection
    Set listaSubs = modSubadministracionDAO.ObtenerListaNombresSubs()
    
    If listaSubs Is Nothing Then
        CargarCatalogoSubadministraciones = Empty
        Exit Function
    End If
    
    Dim listaNombresSub() As String
    ReDim listaNombresSub(0 To listaSubs.Count - 1, 0 To 0)
    
    Dim subObj As clsSubadministracion
    Dim i As Long
    i = 0
    
    For Each subObj In listaSubs
        listaNombresSub(i, 0) = subObj.Nombre
        i = i + 1
    Next subObj
    CargarCatalogoSubadministraciones = listaNombresSub
End Function

' Usado para recuperar una matriz de 2 columnas exclusivamente para
' llenar los datos del cmbBox_JefeDirecto
Public Function ObtenerMatrizJefes() As Variant
    Dim empleados As Collection
    Set empleados = modEmpleadoDAO.getAllEmployees()
    
    If empleados Is Nothing Then
        ObtenerMatrizJefes = Empty
        Exit Function
    End If
    
    If empleados.Count = 0 Then
        ObtenerMatrizJefes = Empty
        Exit Function
    End If
    
    Dim matriz() As String
    ReDim matriz(0 To empleados.Count - 1, 0 To 1)
    
    Dim emp As clsEmpleado
    Dim i As Long
    i = 0
    
    For Each emp In empleados
        matriz(i, 0) = emp.RFC_corto
        matriz(i, 1) = emp.Nombre & " " & emp.aPaterno & " " & emp.aMaterno
        i = i + 1
    Next emp
    
    ObtenerMatrizJefes = matriz
End Function


Public Function cargarCatalogoDepartamentoXSubs(ByRef subSeleccionada As String) As Variant

    If subSeleccionada = "" Then
        cargarCatalogoDepartamentoXSubs = Empty
        Exit Function
    End If
    
    ' 1. El controlador pide los departamentos filtrados al DAO correspondiente
    Dim deptosXSub As Collection
    Set deptosXSub = modDepartamentoDAO.obtenerDeptosXSub(subSeleccionada)

    
    If deptosXSub.Count = 0 Then
        cargarCatalogoDepartamentoXSubs = Empty
        Exit Function
    End If
    
    
    Dim listaNombresDeptos() As Variant
    ReDim listaNombresDeptos(0 To deptosXSub.Count - 1, 0 To 0)
    Dim i As Long
    i = 0
    Dim deptoObj As clsDepartamento
    
    For Each deptoObj In deptosXSub
        listaNombresDeptos(i, 0) = deptoObj.Nombre
        i = i + 1
    Next deptoObj
    
    cargarCatalogoDepartamentoXSubs = listaNombresDeptos
End Function


' =================================================================
' Funciones privadas de controlador
' =================================================================

Public Function ObtenerMatrizEmpleados(ByRef coleccionEmpleados As Collection) As Variant
    If coleccionEmpleados Is Nothing Then
        ObtenerMatrizEmpleados = Empty
        Exit Function
    End If
    
    If coleccionEmpleados.Count = 0 Then
        ObtenerMatrizEmpleados = Empty
        Exit Function
    End If
    
    ' Dimensionamos la matriz para albergar 5 columnas, con formato exacto para la vista
    Dim matriz() As String
    ReDim matriz(0 To coleccionEmpleados.Count - 1, 0 To 4)
    
    Dim emp As clsEmpleado
    Dim i As Long
    i = 0
    
    For Each emp In coleccionEmpleados
        matriz(i, 0) = emp.RFC_corto
        matriz(i, 1) = emp.rfc
        matriz(i, 2) = emp.Nombre & " " & emp.aPaterno & " " & emp.aMaterno
        matriz(i, 3) = emp.puesto
        matriz(i, 4) = emp.subadm
        i = i + 1
    Next emp
    
    ObtenerMatrizEmpleados = matriz
End Function
