Attribute VB_Name = "modSubadministracionDAO"
Option Explicit

Private cacheTablaSubs As ListObject

' Devuelve una colección de objetos de tipo 'clsSubadministracion'
Public Function ObtenerListaNombresSubs() As Collection
    Dim listaNombres As New Collection
    Dim tablaSubs As ListObject
    Dim filaSubs As ListRow
    Dim subObj As clsSubadministracion
    
    Set tablaSubs = ObtenerTablaSubs()
    If tablaSubs Is Nothing Then
        Set ObtenerListaNombresSubs = listaNombres
        Exit Function
    End If
    
    For Each filaSubs In tablaSubs.ListRows
        Set subObj = New clsSubadministracion
        
        Call mapearFilaAObjeto(filaSubs, subObj)
        
        ' Agregamos el POJO a la colección usando el Nombre como llave única
        listaNombres.Add subObj, subObj.Nombre
    Next filaSubs
    
    LimpiarCacheTablaSubs ' Limpieza de caché del ListObject
    Set ObtenerListaNombresSubs = listaNombres
End Function

Private Function ObtenerTablaSubs() As ListObject
    If Not cacheTablaSubs Is Nothing Then
        Set ObtenerTablaSubs = cacheTablaSubs
        Exit Function
    End If
    
    Dim hoja As Worksheet
    Dim objTabla As ListObject
    For Each hoja In ThisWorkbook.Worksheets
        On Error Resume Next
        Set objTabla = hoja.ListObjects(modConfig.TABLA_SUBS)
        On Error GoTo 0
        If Not objTabla Is Nothing Then
            Set cacheTablaSubs = objTabla
            Set ObtenerTablaSubs = cacheTablaSubs
            Exit Function
        End If
    Next hoja
    Set ObtenerTablaSubs = Nothing
End Function

' =================================================================
' MAPEO: De Objeto (Memoria) hacia Fila (Excel)
' =================================================================
Private Sub mapearObjetoAFila(ByRef Subadministracion As clsSubadministracion, ByRef fila As ListRow)
    Dim tblSubs As ListObject
    Set tblSubs = ObtenerTablaSubs() ' Usa la caché, costo de rendimiento = 0
    
    With fila
        .Range(tblSubs.ListColumns(modConfig.COL_SUBS_NOMBRE).Index).value = Subadministracion.Nombre
        .Range(tblSubs.ListColumns(modConfig.COL_SUBS_JEFE).Index).value = Subadministracion.subadministrador
    End With
End Sub

' =================================================================
' MAPEO: De Fila (Excel) hacia Objeto (Memoria)
' =================================================================
Private Sub mapearFilaAObjeto(ByRef fila As ListRow, ByRef Subadministracion As clsSubadministracion)
    Dim tblSubs As ListObject
    Set tblSubs = ObtenerTablaSubs()
    
    With fila
        Subadministracion.Nombre = .Range(tblSubs.ListColumns(modConfig.COL_SUBS_NOMBRE).Index).value
        Subadministracion.subadministrador = .Range(tblSubs.ListColumns(modConfig.COL_SUBS_JEFE).Index).value
    End With
End Sub

Private Sub LimpiarCacheTablaSubs()
    Set cacheTablaSubs = Nothing
End Sub


