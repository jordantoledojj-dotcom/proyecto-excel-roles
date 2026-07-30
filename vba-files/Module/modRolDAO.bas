Attribute VB_Name = "modRolDAO"
Option Explicit

' Dictionary caché con la información mapeada de la tabla en excel al objeto en memoria
' Este sí se usa directamente, por lo que se debe usar con cuidado
Private dictRolesCache As Scripting.Dictionary

Public Function getRolByClave(ByVal claveRol As String) As clsRol
    Set getRolByClave = Nothing
    Call initDictionaryRolesCache
    If dictRolesCache Is Nothing Then Exit Function
    If dictRolesCache.Exists(claveRol) Then
        Set getRolByClave = dictRolesCache(claveRol)
    End If
End Function

Public Function getRolesByPlataforma(ByVal clavePlataforma As String) As Scripting.Dictionary
    Call initDictionaryRolesCache
    Set getRolesByPlataforma = Nothing

    If dictRolesCache Is Nothing Then Exit Function
        
    Dim listaRoles As New Scripting.Dictionary
    Dim key As Variant, rolObj As clsRol

    ' Recuerda: estás iterando por "llaves" (Key). Si quieres, puedes iterar por "Elementos" (Items)
    ' Pero en ese caso, debes usar For Each key In dictRolesCache.Items
    For Each key In dictRolesCache
        Set rolObj = dictRolesCache(key)
        If rolObj.fkClavePlataforma = clavePlataforma Then
            listaRoles.Add rolObj.pkClaveRol, rolObj
        End If
    Next key
    
    Set getRolesByPlataforma = listaRoles
End Function

' =================================================================
' ADVERTENCIA: Estás retornado una referencia directa del cache.
' Es responsabilidad del programador no modificar esta cache.
' Se puede resolver (solo un poco) con la instrucción:
' Set getRoles = clonarDictionary(dictRolesCache)
' Pero afecta el rendimiento que se busca mejorar con la cache y al final, los objetos del
' diccionario no se clonan, solo el dictionary en sí, así que termina siendo un poco igual
' =================================================================
Public Function getRoles() As Scripting.Dictionary
    Call initDictionaryRolesCache
    Set getRoles = dictRolesCache
End Function

' =================================================================
' MAPEO: De Fila (Excel) hacia Objeto (Memoria)
' =================================================================

Private Sub mapearFilaAObjeto(ByVal fila As ListRow, ByRef rolObj As clsRol, ByRef idxs() As Long)
    With fila
        rolObj.pkClaveRol = .Range(idxs(1)).value
        rolObj.rolFuncional = .Range(idxs(2)).value
        rolObj.rolTecnico = .Range(idxs(3)).value
        rolObj.definicionAlcance = .Range(idxs(4)).value
        rolObj.fkClavePlataforma = .Range(idxs(5)).value
    End With
End Sub

' =================================================================
' INFRAESTRUCTURA DE CACHÉ DICTIONARY Y TABLA EN EXCEL
' =================================================================
Private Sub actualizarCache()
    Dim tblRoles As ListObject
    Set tblRoles = obtenerTablaRoles
    If tblRoles Is Nothing Then
        MsgBox "Error Crítico: No se encontró la tabla de roles '" & modConfig.TABLA_ROLES & "'.", vbCritical, "Error DAO"
        End
    End If

    Dim row As ListRow
    Dim rolObj As clsRol

    ' Tambíen funciona la inicialización directa. Pero por si las moscas...
    ' Set dictRolesCache = new Dictionary
    Set dictRolesCache = New Scripting.Dictionary
    ' Por si acaso hacemos que el dictionary no sea KeySensitive
    dictRolesCache.CompareMode = TextCompare
    Dim idxs() As Long

    ' Las llamadas a ListColumns().Index aparentemente consumen mucha memoria. Por eso se hace una sola vez la llamada fuera
    ' del ciclo For Each.
    idxs = getIndexes(tblRoles)

    For Each row In tblRoles.ListRows
        Set rolObj = New clsRol
        Call mapearFilaAObjeto(row, rolObj, idxs)
        If Not dictRolesCache.Exists(rolObj.pkClaveRol) Then
            dictRolesCache.Add rolObj.pkClaveRol, rolObj
        Else
            MsgBox "Error de Integridad: Registro duplicado exacto detectado en la TABLA ROL, fila " & row.Index & _
                   " (" & rolObj.pkClaveRol & ")." & vbCrLf & "El programa se detendrá.", vbCritical, "Datos Corruptos"
            Call invalidarCache
            End
        End If
    Next row
End Sub

Private Function getIndexes(tblRoles As ListObject) As Long()
    Dim idxs() As Long
    ReDim idxs(1 To modConfig.TABLA_ROLES_COLUMN_COUNT)
    
    On Error GoTo ErrColumnas
    idxs(1) = tblRoles.ListColumns(modConfig.COL_ROLES_ID_CLAVE_ROL_TECNICO).Index
    idxs(2) = tblRoles.ListColumns(modConfig.COL_ROLES_ROL_FUNCIONAL).Index
    idxs(3) = tblRoles.ListColumns(modConfig.COL_ROLES_ROL_TECNICO).Index
    idxs(4) = tblRoles.ListColumns(modConfig.COL_ROLES_DEFINICION_ALCANCE).Index
    idxs(5) = tblRoles.ListColumns(modConfig.COL_ROLES_FK_CLAVE_PLATAFORMA).Index
    getIndexes = idxs
    Exit Function
ErrColumnas:
    MsgBox "Error Crítico: Estructura de columnas incorrecta en la tabla ROLES." & vbCrLf & _
           "La macro se detendrá inmediatamente.", vbCritical, "Error de Configuración, modRolDAO"
    Call invalidarCache
    End
End Function

Private Sub initDictionaryRolesCache()
    If Not dictRolesCache Is Nothing Then Exit Sub
    Call actualizarCache
End Sub

Private Function obtenerTablaRoles() As ListObject
    Set obtenerTablaRoles = Nothing

    On Error Resume Next
    Set obtenerTablaRoles = Application.Evaluate(modConfig.TABLA_ROLES)
    On Error GoTo 0
End Function

' Creo que nunca se usa en este DAO, pero por si acaso...
Public Sub invalidarCache()
    Set dictRolesCache = Nothing
End Sub




