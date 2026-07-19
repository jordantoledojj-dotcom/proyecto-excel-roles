Attribute VB_Name = "modAsignarRolController"
Public Function obtenerListaRFC_corto() As Variant
    Dim coleccionEmpleados As Collection
    Dim listaRFC_Empleados As Variant
    
    Set coleccionEmpleados = modEmpleadoDAO.getAllEmployees()
    
    If coleccionEmpleados Is Nothing Then
        obtenerListaRFC_corto = Nothing
        Exit Function
    End If
        
    If coleccionEmpleados.Count = 0 Then
        obtenerListaRFC_corto = Nothing
        Exit Function
    End If
    
    listaRFC_Empleados = ObtenerListaEmpleados(coleccionEmpleados)
    
    obtenerListaRFC_corto = listaRFC_Empleados
End Function

Public Function obtenerEmpleadoByRFC_Corto(ByVal strRFC_corto As String) As clsEmpleado
    Dim empleado As clsEmpleado
    
    Set empleado = modEmpleadoDAO.ObtenerPorId(strRFC_corto)
    
    If empleado Is Nothing Then
        ' Corregido: Se agrega Set para asignar el objeto Nothing
        Set obtenerEmpleadoByRFC_Corto = Nothing
        Exit Function
    End If
    
    ' Corregido: Se agrega Set para retornar la instancia del objeto
    Set obtenerEmpleadoByRFC_Corto = empleado
End Function

Private Function ObtenerListaEmpleados(ByRef coleccionEmp As Collection) As Variant
    Dim empObj As clsEmpleado
    Dim lista() As Variant
    ReDim lista(0 To coleccionEmp.Count - 1, 0 To 0)
    Dim i As Long
    i = 0
    
    For Each empObj In coleccionEmp
        lista(i, 0) = empObj.RFC_corto
        i = i + 1
    Next empObj
    
    ObtenerListaEmpleados = lista
End Function
