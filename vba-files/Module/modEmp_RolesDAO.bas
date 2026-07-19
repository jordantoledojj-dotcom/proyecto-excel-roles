Attribute VB_Name = "modEmp_RolesDAO"
Option Explicit

Private tablaEmpRoles As ListObject

Private Function getRolesEmpleado(ByRef RFC_corto As String) As Collection
    Dim rolesEmpleado As New Collection
    Dim tablaEmpRoles As ListObject
    Dim fila As ListRow
    Dim empRolObj As clsEmp_Rol

    Set tablaEmpRoles = obtenerTablaEmpRoles()
    If tablaEmpRoles Is Nothing Then
        Set getRolesEmpleado = Nothing
        Exit Function
    End If

    For Each fila In tablaEmpRoles.ListRows
        If fila.Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_PK_RFC_CORTO).Index).value = RFC_corto Then
            Set empRolObj = new clsEmp_Rol
            Call mapearFilaAObjeto(fila, empRolObj)
            rolesEmpleado.Add fila
        End If
    Next fila

    Set getRolesEmpleado = rolesEmpleado
End Function

Private Function obtenerTablaEmpRoles() As ListObject
    If Not tablaEmpRoles Is Nothing Then
        Set obtenerTablaEmpRoles = tablaEmpRoles
        Exit Function
    End If

    Dim hoja As Worksheet
    Dim objTabla As ListObject
    For Each hoja In ThisWorkbook.Worksheets
        On Error Resume Next
        Set objTabla = hoja.ListObjects(modConfig.TABLA_EMPLEADOS_ROLES)
        On Error GoTo 0
        If Not objTabla Is Nothing Then
            Set tablaEmpRoles = objTabla
            Set obtenerTablaEmpRoles = tablaEmpRoles
            Exit Function
        End If
    Next hoja
    Set obtenerTablaEmpRoles = Nothing
End Function

' =================================================================
' MAPEO: De Objeto (Memoria) hacia Fila (Excel)
' =================================================================
Private Sub MaeparObjetofila(ByRef empRol As clsEmp_Rol, ByRef fila As ListRow)
    Dim tablaEmpRoles As ListObject
    Set tablaEmpRoles = obtenerTablaEmpRoles()
    
    With fila
        .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_PK_RFC_CORTO).Index).value = empRol.pkRFC_corto
        .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_PK_CLAVE_ROL).Index).value = empRol.pkClaveRolTecnico
        .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_FECHA_ALTA).Index).value = empRol.fechaAlta
        .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_FECHA_BAJA).Index).value = empRol.fechaBaja
        .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_SOLICITANTE).Index).value = empRol.solicitante
        .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_NOTAS).Index).value = empRol.notas
    End With
End Sub

' =================================================================
' MAPEO: De Fila (Excel) hacia Objeto (Memoria)
' =================================================================

Private Sub mapearFilaAObjeto(ByRef fila As ListRow, ByRef empRol As clsEmp_Rol)
    Dim tablaEmpRoles As ListObject
    Set tablaEmpRoles = obtenerTablaEmpRoles()
    
    With fila
        empRol.pkRFC_corto = .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_PK_RFC_CORTO).Index).value
        empRol.pkClaveRolTecnico = .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_PK_CLAVE_ROL).Index).value
        empRol.fechaAlta = .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_FECHA_ALTA).Index).value
        empRol.fechaBaja = .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_FECHA_BAJA).Index).value
        empRol.solicitante = .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_SOLICITANTE).Index).value
        empRol.notas = .Range(tablaEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_NOTAS).Index).value
    End With
End Sub
