Attribute VB_Name = "modDepartamentoDAO"
Option Explicit

Private cacheTablaDepartamentos As ListObject

' Devuelve una colección de POJOs clsDepartamento filtrados por Subadministración
Public Function obtenerDeptosXSub(ByVal nombreSub As String) As Collection
    Dim listaDeptos As New Collection
    Dim tablaDeptos As ListObject
    Dim fila As ListRow
    Dim deptoObj As clsDepartamento
    
    ' Si no viene ninguna subadministración, regresamos la colección vacía
    If Trim(nombreSub) = "" Then
        Set obtenerDeptosXSub = listaDeptos
        Exit Function
    End If
    
    Set tablaDeptos = ObtenerTablaDeptos()
    If tablaDeptos Is Nothing Then
        Set obtenerDeptosXSub = listaDeptos
        Exit Function
    End If
    
    For Each fila In tablaDeptos.ListRows
        If fila.Range(tablaDeptos.ListColumns(modConfig.COL_DEPTOS_SUBADMIN).Index).value = nombreSub Then
            Set deptoObj = New clsDepartamento
            Call MapearFilaAObjeto(fila, deptoObj)
            listaDeptos.Add deptoObj
        End If
    Next fila
    
    LimpiarCacheTablaDepartamentos
    Set obtenerDeptosXSub = listaDeptos
End Function

Private Function ObtenerTablaDeptos() As ListObject
    If Not cacheTablaDepartamentos Is Nothing Then
        Set ObtenerTablaDeptos = cacheTablaDepartamentos
        Exit Function
    End If
    
    Dim hoja As Worksheet
    Dim objTabla As ListObject
    For Each hoja In ThisWorkbook.Worksheets
        On Error Resume Next
        Set objTabla = hoja.ListObjects(modConfig.TABLA_DEPTOS)
        On Error GoTo 0
        If Not objTabla Is Nothing Then
            Set cacheTablaDepartamentos = objTabla
            Set ObtenerTablaDeptos = cacheTablaDepartamentos
            Exit Function
        End If
    Next hoja
    Set ObtenerTablaDeptos = Nothing
End Function

' =================================================================
' MAPEO: De Fila (Excel) hacia Objeto (Memoria)
' =================================================================
Private Sub MapearFilaAObjeto(ByRef fila As ListRow, ByRef Departamento As clsDepartamento)
    Dim tablaDeptos As ListObject
    Set tablaDeptos = ObtenerTablaDeptos()
    
    With fila
        Departamento.Nombre = .Range(tablaDeptos.ListColumns(modConfig.COL_DEPTOS_NOMBRE).Index).value
        Departamento.Subadministracion = .Range(tablaDeptos.ListColumns(modConfig.COL_DEPTOS_SUBADMIN).Index).value
    End With
End Sub

Private Sub LimpiarCacheTablaDepartamentos()
    Set cacheTablaDepartamentos = Nothing
End Sub
