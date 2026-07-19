Attribute VB_Name = "modPlataformaDAO"

Option Explicit

Private cacheTablaPlataformas As ListObject

Public Function obtenerPlataformas() As Collection
    Dim listaPlataformas As New Collection
    Dim tablaPlataformas As ListObject
    Dim fila As ListRow
    Dim plataformaObj As clsPlataforma
    
    Set tablaPlataformas = ObtenerTablaPlataformas()
    If tablaPlataformas Is Nothing Then
        Set obtenerPlataformas = listaPlataformas
        Exit Function
    End If
    
    For Each fila In tablaPlataformas.ListRows
        Set plataformaObj = New clsPlataforma
        Call mapearFilaAObjeto(fila, plataformaObj)
        listaPlataformas.Add plataformaObj
    Next fila
    
    'LimpiarCacheTablaPlataformas
    Set obtenerPlataformas = listaPlataformas
End Function

Public Function getPlataformaByID_Clave(ByVal clavePlataforma As String) As clsPlataforma
    Dim tablaPlataformas As ListObject
    Dim fila As ListRow
    Dim plataformaBuscada As clsPlataforma
    
    Set tablaPlataformas = ObtenerTablaPlataformas()
    If tablaPlataformas Is Nothing Then
        Set getPlataformaByID_Clave = Nothing
        Exit Function
    End If
    
    For Each fila In tablaPlataformas.ListRows
        If fila.Range(tablaPlataformas.ListColumns(modConfig.COL_PLATAFORMAS_ID_CLAVE_PLATAFORMA).Index).value = clavePlataforma Then
            Set plataformaBuscada = New clsPlataforma
            Call mapearFilaAObjeto(fila, plataformaBuscada)
            Set getPlataformaByID_Clave = plataformaBuscada
            Exit Function
        End If
    Next fila
    
    ' LimpiarCacheTablaPlataformas
    Set getPlataformaByID_Clave = Nothing
End Function

Private Function ObtenerTablaPlataformas() As ListObject
    If Not cacheTablaPlataformas Is Nothing Then
        Set ObtenerTablaPlataformas = cacheTablaPlataformas
        Exit Function
    End If
    
    Dim hoja As Worksheet
    Dim objTabla As ListObject
    For Each hoja In ThisWorkbook.Worksheets
        On Error Resume Next
        Set objTabla = hoja.ListObjects(modConfig.TABLA_PLATAFORMAS)
        On Error GoTo 0
        If Not objTabla Is Nothing Then
            Set cacheTablaPlataformas = objTabla
            Set ObtenerTablaPlataformas = cacheTablaPlataformas
            Exit Function
        End If
    Next hoja
    Set ObtenerTablaPlataformas = Nothing
End Function

' =================================================================
' MAPEO: De Fila (Excel) hacia Objeto (Memoria)
' =================================================================

Private Sub mapearFilaAObjeto(ByRef fila As ListRow, ByRef Plataforma As clsPlataforma)
    Dim tablaPlataformas As ListObject
    Set tablaPlataformas = ObtenerTablaPlataformas()
    
    With fila
        Plataforma.ID_clavaPlataforma = .Range(tablaPlataformas.ListColumns(modConfig.COL_PLATAFORMAS_ID_CLAVE_PLATAFORMA).Index).value
        Plataforma.nombrePlataforma = .Range(tablaPlataformas.ListColumns(modConfig.COL_PLATAFORMAS_NOMBRE).Index).value
        Plataforma.descripcion = .Range(tablaPlataformas.ListColumns(modConfig.COL_PLATAFORMAS_DESCRIPCION).Index).value
    End With
End Sub
