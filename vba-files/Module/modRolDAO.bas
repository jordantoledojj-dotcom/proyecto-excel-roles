Attribute VB_Name = "modRolDAO"
Option Explicit

Private tablaRoles As ListObject

Public Function getRolByClave(ByVal claveRol As String) As clsRol
    Dim rolObj As clsRol
    Dim tablaRoles As ListObject
    Dim fila As ListRow
    
    Set tablaRoles = obtenerTablaRoles()
    If tablaRoles Is Nothing Then
        Set getRolByClave = Nothing
        Exit Function
    End If
    
    For Each fila In tablaRoles.ListRows
        If fila.Range(tablaRoles.ListColumns(modConfig.COL_ROLES_ID_CLAVE_ROL_TECNICO).Index).value = claveRol Then
            Set rolObj = New clsRol
            Call mapearFilaAObjeto(fila, rolObj)
            Set getRolByClave = rolObj
            Exit Function
        End If
    Next fila
    
    Set getRolByClave = Nothing
End Function

Public Function getRolesByPlataforma(ByVal clavePlataforma As String) As Collection
    Dim listaRoles As New Collection
    Dim tablaRoles As ListObject
    Dim fila As ListRow
    Dim rolObj As clsRol
    
    Set tablaRoles = obtenerTablaRoles()
    If tablaRoles Is Nothing Then
        Set getRolesByPlataforma = listaRoles
        Exit Function
    End If
    
    For Each fila In tablaRoles.ListRows
        If fila.Range(tablaRoles.ListColumns(modConfig.COL_ROLES_FK_CLAVE_PLATAFORMA).Index).value = clavePlataforma Then
            Set rolObj = New clsRol
            Call mapearFilaAObjeto(fila, rolObj)
            listaRoles.Add rolObj
        End If
    Next fila
    
    Set getRolesByPlataforma = listaRoles
End Function

Private Function obtenerTablaRoles() As ListObject
    If Not tablaRoles Is Nothing Then
        Set obtenerTablaRoles = tablaRoles
        Exit Function
    End If
    
    Dim hoja As Worksheet
    Dim objTabla As ListObject
    For Each hoja In ThisWorkbook.Worksheets
        On Error Resume Next
        Set objTabla = hoja.ListObjects(modConfig.TABLA_ROLES)
        On Error GoTo 0
        
        If Not objTabla Is Nothing Then
            Set tablaRoles = objTabla
            Set obtenerTablaRoles = tablaRoles
            Exit Function
        End If
    Next hoja
    
    Set obtenerTablaRoles = Nothing
End Function

' =================================================================
' MAPEO: De Fila (Excel) hacia Objeto (Memoria)
' =================================================================

Private Sub mapearFilaAObjeto(ByVal fila As ListRow, ByRef rolObj As clsRol)
    Dim tblRoles As ListObject
    Set tblRoles = obtenerTablaRoles()

    With fila
        rolObj.pkClaveRol = .Range(tblRoles.ListColumns(modConfig.COL_ROLES_ID_CLAVE_ROL_TECNICO).Index).value
        rolObj.rolFuncional = .Range(tblRoles.ListColumns(modConfig.COL_ROLES_ROL_FUNCIONAL).Index).value
        rolObj.rolTecnico = .Range(tblRoles.ListColumns(modConfig.COL_ROLES_ROL_TECNICO).Index).value
        rolObj.definicionAlcance = .Range(tblRoles.ListColumns(modConfig.COL_ROLES_DEFINICION_ALCANCE).Index).value
        rolObj.fkClavePlataforma = .Range(tblRoles.ListColumns(modConfig.COL_ROLES_FK_CLAVE_PLATAFORMA).Index).value
    End With
End Sub
