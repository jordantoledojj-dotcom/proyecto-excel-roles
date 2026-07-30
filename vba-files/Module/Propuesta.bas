Attribute VB_Name = "Propuesta"
Option Explicit

' Caches en memoria a nivel de módulo con tipos fuertes de Scripting Runtime
Private tblRolesCache As ListObject
Private dictRolesCache As Scripting.Dictionary

' =================================================================
' MÉTODOS DE LECTURA (Consultas directas a la Caché en Memoria)
' =================================================================

Public Function getRolByClave(ByVal claveRol As String) As clsRol
    Call asegurarCacheCargada
    
    If dictRolesCache Is Nothing Then
        Set getRolByClave = Nothing
        Exit Function
    End If
    
    ' El autocompletado de VBA ahora te sugerirá .Exists
    If dictRolesCache.Exists(claveRol) Then
        Set getRolByClave = dictRolesCache(claveRol)
    Else
        Set getRolByClave = Nothing
    End If
End Function

Public Function getRolesByPlataforma(ByVal clavePlataforma As String) As Collection
    Dim listaRoles As New Collection
    Dim vElemento As Variant
    Dim rolObj As clsRol
    
    Call asegurarCacheCargada
    If dictRolesCache Is Nothing Then
        Set getRolesByPlataforma = listaRoles
        Exit Function
    End If
    
    ' Iteración tipada y veloz en memoria
    For Each vElemento In dictRolesCache.Items
        Set rolObj = vElemento
        If rolObj.fkClavePlataforma = clavePlataforma Then
            listaRoles.Add rolObj
        End If
    Next vElemento
    
    Set getRolesByPlataforma = listaRoles
End Function

Public Function getRoles() As Collection
    Dim listaRoles As New Collection
    Dim vElemento As Variant
    Dim rolObj As clsRol
    
    Call asegurarCacheCargada
    If dictRolesCache Is Nothing Then
        Set getRoles = listaRoles
        Exit Function
    End If
    
    For Each vElemento In dictRolesCache.Items
        Set rolObj = vElemento
        listaRoles.Add rolObj, rolObj.pkClaveRol
    Next vElemento
    
    Set getRoles = listaRoles
End Function

' =================================================================
' INFRAESTRUCTURA DE CACHÉ Y MAPEO (Privados)
' =================================================================

Private Sub asegurarCacheCargada()
    If Not dictRolesCache Is Nothing Then Exit Sub
    
    ' Instanciación directa gracias a la referencia explícita
    Set dictRolesCache = New Scripting.Dictionary
    dictRolesCache.CompareMode = TextCompare ' Ahora podemos usar la constante nativa en vez de un número
    
    Dim tbl As ListObject
    Set tbl = obtenerTablaRoles()
    If tbl Is Nothing Then Exit Sub
    
    Dim fila As ListRow
    Dim rolObj As clsRol
    
    For Each fila In tbl.ListRows
        Set rolObj = New clsRol
        Call mapearFilaAObjeto(fila, rolObj)
        
        If Not dictRolesCache.Exists(rolObj.pkClaveRol) Then
            dictRolesCache.Add rolObj.pkClaveRol, rolObj
        End If
    Next fila
End Sub

Private Function obtenerTablaRoles() As ListObject
    If Not tblRolesCache Is Nothing Then
        Set obtenerTablaRoles = tblRolesCache
        Exit Function
    End If
    
    Dim hoja As Worksheet
    Dim objTabla As ListObject
    
    For Each hoja In ThisWorkbook.Worksheets
        On Error Resume Next
        Set objTabla = hoja.ListObjects(modConfig.TABLA_ROLES)
        On Error GoTo 0
        
        If Not objTabla Is Nothing Then
            Set tblRolesCache = objTabla
            Set obtenerTablaRoles = tblRolesCache
            Exit Function
        End If
    Next hoja
    
    Set obtenerTablaRoles = Nothing
End Function

Private Sub mapearFilaAObjeto(ByVal fila As ListRow, ByRef rolObj As clsRol)
    Dim tbl As ListObject
    Set tbl = obtenerTablaRoles()

    With fila
        rolObj.pkClaveRol = .Range(tbl.ListColumns(modConfig.COL_ROLES_ID_CLAVE_ROL_TECNICO).Index).value
        rolObj.rolFuncional = .Range(tbl.ListColumns(modConfig.COL_ROLES_ROL_FUNCIONAL).Index).value
        rolObj.rolTecnico = .Range(tbl.ListColumns(modConfig.COL_ROLES_ROL_TECNICO).Index).value
        rolObj.definicionAlcance = .Range(tbl.ListColumns(modConfig.COL_ROLES_DEFINICION_ALCANCE).Index).value
        rolObj.fkClavePlataforma = .Range(tbl.ListColumns(modConfig.COL_ROLES_FK_CLAVE_PLATAFORMA).Index).value
    End With
End Sub

Private Sub mapearObjetoAFila(ByVal rolObj As clsRol, ByRef fila As ListRow)
    Dim tbl As ListObject
    Set tbl = obtenerTablaRoles()
    
    With fila
        .Range(tbl.ListColumns(modConfig.COL_ROLES_ID_CLAVE_ROL_TECNICO).Index).value = rolObj.pkClaveRol
        .Range(tbl.ListColumns(modConfig.COL_ROLES_ROL_FUNCIONAL).Index).value = rolObj.rolFuncional
        .Range(tbl.ListColumns(modConfig.COL_ROLES_ROL_TECNICO).Index).value = rolObj.rolTecnico
        .Range(tbl.ListColumns(modConfig.COL_ROLES_DEFINICION_ALCANCE).Index).value = rolObj.definicionAlcance
        .Range(tbl.ListColumns(modConfig.COL_ROLES_FK_CLAVE_PLATAFORMA).Index).value = rolObj.fkClavePlataforma
    End With
End Sub

Public Sub invalidarCache()
    Set tblRolesCache = Nothing
    Set dictRolesCache = Nothing
End Sub

