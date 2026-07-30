Attribute VB_Name = "modEmp_RolesDAO"
Option Explicit

Private dictEmpRolesCache As Scripting.Dictionary
Private tablaEmpRoles As ListObject

Public Sub invalidarCache()
    Set dictEmpRolesCache = Nothing
    Set tablaEmpRoles = Nothing
End Sub

Private Function getRolesActivosEmpleado(ByRef RFC_corto As String) As Scripting.Dictionary
    Call initDictionaryEmpRolesCache

    Set getRolesActivosEmpleado = Nothing
    If dictEmpRolesCache Is Nothing Then Exit Function

    Dim rolesActivosEmpleado As New Scripting.Dictionary
    Dim key As Variant, empRolObj As clsEmp_Rol

    For Each key In dictEmpRolesCache
        Set empRolObj = dictEmpRolesCache(key)
        ' If empRolObj.pkRFC_corto = RFC_corto And empRolObj.esRolActivo Then
        ' La instrucción anterior y la siguiente hacen lo mismo, pero supuestamente es más rápida
        ' la de abajo a nivel procesador. Quién sabe...
        If StrComp(empRolObj.pkRFC_corto, RFC_corto, vbTextCompare) = 0 And empRolObj.esRolActivo Then
            rolesActivosEmpleado.Add empRolObj.pkID, empRolObj
        End If
    Next key
    Set getRolesActivosEmpleado = rolesActivosEmpleado
End Function

' Esto se usará algún momento?
' Porque en realidad, parece ser que solo se usa getRolesActivosEmpleados y los inactivos,
' solo como histórico se mostrarían
Private Function getRolesEmpleado(ByRef RFC_corto As String) As Scripting.Dictionary
    Call initDictionaryEmpRolesCache

    Set getRolesEmpleado = Nothing
    If dictEmpRolesCache Is Nothing Then Exit Function

    Dim rolesEmpleado As New Scripting.Dictionary
    Dim key As Variant, empRolObj As clsEmp_Rol

    For Each key In dictEmpRolesCache
        Set empRolObj = dictEmpRolesCache(key)
        If StrCmp(empRolObj.pkRFC_corto, RFC_corto, vbTextCompare) = 0 Then
            rolesEmpleado.Add empRolObj.pkID, empRolObj
        End If
    Next key

    Set getRolesEmpleado = rolesEmpleado
End Function

' =================================================================
' MAPEO: De Fila (Excel) hacia Objeto (Memoria)
' =================================================================

Private Sub mapearFilaAObjeto(ByRef row As ListRow, ByRef empRolObj As clsEmp_Rol, ByRef idxs() As Long)
    With row
        empRolObj.pkRFC_corto = .Range(idxs(1)).value
        empRolObj.pkClaveRolTecnico = .Range(idxs(2)).value
        empRolObj.fechaAlta = .Range(idxs(3)).value
        empRolObj.fechaBaja = .Range(idxs(4)).value
        empRolObj.solicitante = .Range(idxs(5)).value
        empRolObj.notas = .Range(idxs(6)).value
    End With
End Sub

' =================================================================
' MAPEO: De Objeto (Memoria) hacia Fila (Excel)
' =================================================================
Private Sub mapearObjetoAFila(ByRef empRolObj As clsEmp_Rol, ByRef row As ListRow, ByRef idxs() As Long)
    With row
        .Range(idxs(1)).value = empRolObj.pkRFC_corto
        .Range(idxs(2)).value = empRolObj.pkClaveRolTecnico
        .Range(idxs(3)).value = empRolObj.fechaAlta
        .Range(idxs(4)).value = empRolObj.fechaBaja
        .Range(idxs(5)).value = empRolObj.solicitante
        .Range(idxs(6)).value = empRolObj.notas
    End With
End Sub

' =================================================================
' INFRAESTRUCTURA DE CACHÉ DICTIONARY Y TABLA EN EXCEL
' =================================================================
Private Sub actualizarCache()
    Dim tblEmpRoles As ListObject
    Set tblEmpRoles = obtenerTablaEmpRoles
    If tblEmpRoles Is Nothing Then
        MsgBox "Error Crítico: No se encontró la tabla de Empleados-Roles '" & modConfig.TABLA_EMPLEADOS_ROLES & "'.", vbCritical, "Error modEmprolesDAO.bas"
        End
    End If

    Dim row As ListRow, empRolObj As clsEmp_Rol

    Set dictEmpRolesCache = New Scripting.Dictionary
    dictEmpRolesCache.CompareMode = TextCompare
    Dim idxs() As Long
    idxs = getIndexes(tblEmpRoles)

    For Each row In tblEmpRoles.ListRows
        Set empRolObj = New clsEmp_Rol
        Call mapearFilaAObjeto(row, empRolObj, idxs)
        If dictEmpRolesCache.Exists(empRolObj.pkID) Then
            MsgBox "Error de Integridad: Registro duplicado exacto detectado en la TABLA EMP-ROL, fila " & row.Index & _
                   " (" & empRolObj.pkID & ")." & vbCrLf & "El programa se detendrá.", vbCritical, "Datos Corruptos"
            Call invalidarCache
            End
        Else
            dictEmpRolesCache.Add empRolObj.pkID, empRolObj
        End If
    Next row
End Sub

Private Function getIndexes(ByRef tblEmpRoles As ListObject) As Long()
    Dim idxs() As Long
    ReDim idxs(1 To modConfig.TABLA_EMPLEADO_ROLES_COLUMN_COUNT)

    On Error GoTo ErrColumnas
    idxs(1) = tblEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_PK_RFC_CORTO).Index
    idxs(2) = tblEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_PK_CLAVE_ROL).Index
    idxs(3) = tblEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_FECHA_ALTA).Index
    idxs(4) = tblEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_FECHA_BAJA).Index
    idxs(5) = tblEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_SOLICITANTE).Index
    idxs(6) = tblEmpRoles.ListColumns(modConfig.COL_EMP_ROLES_NOTAS).Index
    getIndexes = idxs
    Exit Function
ErrColumnas:
    MsgBox "Error Crítico: Estructura de columnas incorrecta en la tabla de Emplados Roles." & vbCrLf & _
           "La macro se detendrá inmediatamente.", vbCritical, "Error de Configuración, modEmp_RolDAO"
    Call invalidarCache
    End
End Function

Private Sub initDictionaryEmpRolesCache()
    If Not dictEmpRolesCache Is Nothing Then Exit Sub
    Call actualizarCache
End Sub

Private Function obtenerTablaEmpRoles() As ListObject
    If Not tablaEmpRoles Is Nothing Then
        Set obtenerTablaEmpRoles = tablaEmpRoles
        Exit Function
    End If

    Set obtenerTablaEmpRoles = Nothing
    On Error Resume Next
    Set tablaEmpRoles = Application.Evaluate(modConfig.TABLA_EMPLEADOS_ROLES)
    Set obtenerTablaEmpRoles = tablaEmpRoles
    On Error GoTo 0
End Function

