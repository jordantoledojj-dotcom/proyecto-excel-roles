Attribute VB_Name = "modEmpleadoDAO"
Option Explicit

' Variable de estado privada al módulo para control de caché
Private tablaEmpleados As ListObject

' =================================================================
' CREATE: Inserta un nuevo empleado
' =================================================================
Public Function Insertar(ByRef empleado As clsEmpleado, ByRef ErrMensaje As String) As Boolean
    Dim nuevaFila As ListRow
    
    ' 1. Intentar conectar con la tabla de Excel
    If ObtenertablaEmpleados() Is Nothing Then
        ErrMensaje = "Error crítico: No se encontró la tabla '" & modConfig.TABLA_EMPLEADOS & "' en el libro."
        Insertar = False
        Exit Function
    End If
    
    ' 2. Validación de Integridad: Verificar que el rfc_corto (ID) NO exista ya
    If ExisteId(empleado.RFC_corto) Then
        ErrMensaje = "Error de integridad: El RFC corto '" & empleado.RFC_corto & "' ya se encuentra registrado."
        Insertar = False
        Exit Function
    End If
    
    ' 3. Si pasa la validación, añadimos una nueva fila a la tabla
    Set nuevaFila = ObtenertablaEmpleados().ListRows.Add
    
    ' 4. Mapeo del Objeto hacia las columnas de la Tabla (por el nombre de la cabecera)
    Call mapearObjetoAFila(empleado, nuevaFila)
    Call LimpiarCacheTablaEmpleados
    
    Insertar = True ' Operación exitosa
    
End Function

' =================================================================
' UPDATE: Edita un empleado existente
' =================================================================
Public Function Actualizar(ByRef empleado As clsEmpleado, ByRef ErrMensaje As String) As Boolean
    Dim celdaRFC_corto As Range
    
    Set celdaRFC_corto = ObtenertablaEmpleados().ListColumns(modConfig.COL_EMP_RFC_CORTO).DataBodyRange.Find( _
                    What:=empleado.RFC_corto, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    
    If celdaRFC_corto Is Nothing Then
        ErrMensaje = "El empleado con RFC Corto '" & empleado.RFC_corto & "' no existe."
        Actualizar = False
        Exit Function
    End If
    
    Call mapearObjetoAFila(empleado, ObtenertablaEmpleados().ListRows(celdaRFC_corto.row - ObtenertablaEmpleados().HeaderRowRange.row)) 'nótese la resta
    ' La resta dentro del argumento es para determinar en qué fila relativa de la tabla está.
    
    Call LimpiarCacheTablaEmpleados
    
    Actualizar = True
End Function

' =================================================================
' READ: Obtiene un objeto clsEmpleado mediante su rfc_corto
' =================================================================
Public Function ObtenerPorId(ByVal RFC_corto As String) As clsEmpleado
    Dim celdaRFC_corto As Range
    Dim filaIndex As Long
    Dim empleadoEncontrado As clsEmpleado
    
    If ObtenertablaEmpleados() Is Nothing Then
        Set ObtenerPorId = Nothing
        Exit Function
    End If
    
    If ObtenertablaEmpleados().ListRows.Count = 0 Then
        Set ObtenerPorId = Nothing
        Exit Function
    End If
    
    ' Buscamos el RFC_corto únicamente en la columna de "rfc_corto"
    Set celdaRFC_corto = ObtenertablaEmpleados().ListColumns(modConfig.COL_EMP_RFC_CORTO).DataBodyRange.Find( _
                    What:=RFC_corto, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    ' Si se encuentra el RFC_corto en la tabla
    If Not celdaRFC_corto Is Nothing Then
        ' Determinamos en qué fila relativa de la tabla está
        filaIndex = celdaRFC_corto.row - ObtenertablaEmpleados().HeaderRowRange.row
        
        ' Instanciamos el objeto y realizamos el mapeo inverso (De Tabla a Objeto)
        Set empleadoEncontrado = New clsEmpleado
        Call mapearFilaAObjeto(ObtenertablaEmpleados.ListRows(filaIndex), empleadoEncontrado)
        Set ObtenerPorId = empleadoEncontrado
    Else
        Set ObtenerPorId = Nothing ' No existe (null)
    End If
    
    Call LimpiarCacheTablaEmpleados
End Function


' =================================================================
' READ, CONSULTA GENERAL: Devuelve todos los empleados en una Colección
' =================================================================
Public Function getAllEmployees() As Collection
    Dim fila As ListRow
    Dim lista As New Collection
    Dim emp As clsEmpleado
    
    'Si la tabla no existe o no tiene registros, nos salimos
    If ObtenertablaEmpleados() Is Nothing Then
        Set getAllEmployees = Nothing
        Exit Function
    End If
    
    If ObtenertablaEmpleados().ListRows.Count = 0 Then
        Set getAllEmployees = Nothing
        Exit Function
    End If
        
    For Each fila In ObtenertablaEmpleados().ListRows
        Set emp = New clsEmpleado
        
        ' Mapeo de la tabla de Excel al Objeto
        Call mapearFilaAObjeto(fila, emp)
        
        ' Agregamos el objeto a la colección utilizando el rfc_corto como Key
        lista.Add emp, emp.RFC_corto
    Next fila
    
    Set getAllEmployees = lista
    Call LimpiarCacheTablaEmpleados
End Function

' =================================================================
' READ, CONSULTA FILTRADA: Busca un término en cualquier celda de la tabla
' =================================================================
Public Function obtenerEmpleadosPorBusqueda(ByVal strBuscar As String) As Collection
    Dim listaResultados As New Collection
    Dim celdaEncontrada As Range
    Dim primeraDireccion As String
    Dim filaIndex As Long
    Dim emp As clsEmpleado
    Dim llavesRegistradas As New Collection ' Para evitar duplicados si el término aparece dos veces en la misma fila
    
    ' --- : Si la cadena es nula/vacía, devolvemos TODO ---
    If strBuscar = "" Then
        Set obtenerEmpleadosPorBusqueda = getAllEmployees()
        Exit Function
    End If
    
    strBuscar = Trim(strBuscar)
    
    If ObtenertablaEmpleados() Is Nothing Then
        Set obtenerEmpleadosPorBusqueda = listaResultados
        Exit Function
    End If
    
    If ObtenertablaEmpleados().ListRows.Count = 0 Then
        Set obtenerEmpleadosPorBusqueda = listaResultados
        Exit Function
    End If
    
    ' Buscamos la primera coincidencia en cualquier parte del cuerpo de la tabla
    ' Lookat:=xlPart permite coincidencias parciales (ej: "Sán" encuentra "Sánchez")
    Set celdaEncontrada = ObtenertablaEmpleados().DataBodyRange.Find( _
                            What:=strBuscar, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                            
    If Not celdaEncontrada Is Nothing Then
        primeraDireccion = celdaEncontrada.Address
        
        Do
            ' Determinamos la fila relativa en la tabla
            filaIndex = celdaEncontrada.row - ObtenertablaEmpleados().HeaderRowRange.row
            
            ' Obtenemos el ID único (RFC Corto) de esa fila para validar duplicados
            Dim idCorto As String
            idCorto = ObtenertablaEmpleados().ListRows(filaIndex).Range(ObtenertablaEmpleados().ListColumns(modConfig.COL_EMP_RFC_CORTO).Index).value
            
            ' Verificamos si ya procesamos esta fila en este ciclo de búsqueda
            On Error Resume Next
            llavesRegistradas.Add idCorto, idCorto
            Dim esDuplicado As Boolean
            esDuplicado = (Err.Number <> 0)
            On Error GoTo 0
            
            If Not esDuplicado Then
                ' Instanciamos y mapeamos el empleado (Reutiliza la lógica de mapeo de tu getAllEmployees)
                Set emp = New clsEmpleado
                Call mapearFilaAObjeto(ObtenertablaEmpleados().ListRows(filaIndex), emp)
                
                listaResultados.Add emp, emp.RFC_corto
            End If
            
            ' Buscamos la siguiente coincidencia
            Set celdaEncontrada = ObtenertablaEmpleados().DataBodyRange.FindNext(celdaEncontrada)
            
        Loop While Not celdaEncontrada Is Nothing And celdaEncontrada.Address <> primeraDireccion
    End If
    
    Set obtenerEmpleadosPorBusqueda = listaResultados
    Call LimpiarCacheTablaEmpleados
End Function

' =================================================================
' DELETE: Elimina un empleado de la tabla de Excel
' =================================================================
Public Function Eliminar(ByVal RFC_corto As String, ByRef ErrMensaje As String) As Boolean
    Dim celdaId As Range
    Dim filaIndex As Long
    Dim tabla As ListObject
    
    Set tabla = ObtenertablaEmpleados()
    If tabla Is Nothing Then
        ErrMensaje = "Error crítico: No se encontró la tabla de empleados."
        Eliminar = False
        Exit Function
    End If
    
    ' Localizamos el registro a eliminar
    Set celdaId = tabla.ListColumns(modConfig.COL_EMP_RFC_CORTO).DataBodyRange.Find( _
                    What:=RFC_corto, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    
    If celdaId Is Nothing Then
        ErrMensaje = "No se puede eliminar: El empleado con RFC Corto '" & RFC_corto & "' no existe."
        Eliminar = False
        Exit Function
    End If
    
    filaIndex = celdaId.row - tabla.HeaderRowRange.row
    
    ' Eliminamos físicamente la fila del ListObject
    tabla.ListRows(filaIndex).Delete
    
    Eliminar = True
End Function

' =================================================================
' METODOS AUXILIARES (PRIVADOS DEL DAO)
' =================================================================

' Función interna para verificar si un ID ya existe (Devuelve True/False)
Private Function ExisteId(ByVal ID As String) As Boolean
    Dim busqueda As Range
    
    If ObtenertablaEmpleados().ListRows.Count = 0 Then
        ExisteId = False
        Exit Function
    End If
    
    Set busqueda = ObtenertablaEmpleados.ListColumns(modConfig.COL_EMP_RFC_CORTO).DataBodyRange.Find( _
                    What:=ID, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    
    ExisteId = (Not busqueda Is Nothing)
End Function

' Función para localizar el objeto ListObject recorriendo las hojas del libro
Private Function ObtenertablaEmpleados() As ListObject
    ' Simulación caché: Si ya tiene una referencia, la devolvemos de inmediato sin hacer una
    ' nueva búsqueda
    If Not tablaEmpleados Is Nothing Then
        Set ObtenertablaEmpleados = tablaEmpleados
        Exit Function
    End If
    
    Dim hoja As Worksheet
    Dim objTabla As ListObject
    For Each hoja In ThisWorkbook.Worksheets
        On Error Resume Next
        Set objTabla = hoja.ListObjects(modConfig.TABLA_EMPLEADOS)
        On Error GoTo 0
        
        If Not objTabla Is Nothing Then
            Set tablaEmpleados = objTabla ' <-- Guardar en la caché
            Set ObtenertablaEmpleados = tablaEmpleados
            Exit Function
        End If
    Next hoja
    Set ObtenertablaEmpleados = Nothing
End Function


' =================================================================
' MAPEO: De Objeto (Memoria) hacia Fila (Excel)
' =================================================================
Private Sub mapearObjetoAFila(ByRef empleado As clsEmpleado, ByRef fila As ListRow)
    Dim tblEmplt As ListObject
    Set tblEmplt = ObtenertablaEmpleados() ' Usa la caché, costo de rendimiento = 0
    
    With fila
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_RFC_CORTO).Index).value = empleado.RFC_corto
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_NOMBRE).Index).value = empleado.Nombre
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_PATERNO).Index).value = empleado.aPaterno
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_MATERNO).Index).value = empleado.aMaterno
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_RFC).Index).value = empleado.rfc
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_F_NACIMIENTO).Index).value = empleado.fNacimiento
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_NOMBRAMIENTO).Index).value = empleado.nombramiento
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_PUESTO).Index).value = empleado.puesto
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_SUBADMIN).Index).value = empleado.subadm
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_DEPARTAMENTO).Index).value = empleado.Departamento
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_JEFE_DIRECTO).Index).value = empleado.jefeDirecto
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_ESTATUS).Index).value = empleado.estatus
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_F_INGRESO).Index).value = empleado.fIngreso
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_F_INICIO_CARGO).Index).value = empleado.fInicioCargoActual
        .Range(tblEmplt.ListColumns(modConfig.COL_EMP_EXTENSION).Index).value = empleado.extension
    End With
End Sub

' =================================================================
' MAPEO: De Fila (Excel) hacia Objeto (Memoria)
' =================================================================
Private Sub mapearFilaAObjeto(ByRef fila As ListRow, ByRef empleado As clsEmpleado)
    Dim tblEmpl As ListObject
    Set tblEmpl = ObtenertablaEmpleados()
    
    With fila
        empleado.RFC_corto = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_RFC_CORTO).Index).value
        empleado.Nombre = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_NOMBRE).Index).value
        empleado.aPaterno = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_PATERNO).Index).value
        empleado.aMaterno = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_MATERNO).Index).value
        empleado.rfc = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_RFC).Index).value
        empleado.nombramiento = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_NOMBRAMIENTO).Index).value
        empleado.puesto = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_PUESTO).Index).value
        empleado.subadm = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_SUBADMIN).Index).value
        empleado.Departamento = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_DEPARTAMENTO).Index).value
        empleado.jefeDirecto = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_JEFE_DIRECTO).Index).value
        empleado.estatus = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_ESTATUS).Index).value
        empleado.fIngreso = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_F_INGRESO).Index).value
        empleado.fNacimiento = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_F_NACIMIENTO).Index).value
        empleado.fInicioCargoActual = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_F_INICIO_CARGO).Index).value
        empleado.extension = .Range(tblEmpl.ListColumns(modConfig.COL_EMP_EXTENSION).Index).value
    End With
End Sub

Private Sub LimpiarCacheTablaEmpleados()
    Set tablaEmpleados = Nothing
End Sub
