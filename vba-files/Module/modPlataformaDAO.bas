Attribute VB_Name = "modPlataformaDAO"

Option Explicit

Private dictPlataformaCache As Scripting.Dictionary

Public Function getPlataformas() As Scripting.Dictionary
    Call initDictionaryPlataformasCache
    Set getPlataformas = dictPlataformaCache
End Function

Public Function getPlataformaByID_Clave(ByVal clavePlataforma As String) As clsPlataforma
    Call initDictionaryPlataformasCache
    Dim plataformabuscada As clsPlataforma
    Set plataformabuscada = Nothing
    If dictPlataformaCache Is Nothing Then Exit Function
    If dictPlataformaCache.Exists(clavePlataforma) Then
        Set getPlataformaByID_Clave = dictPlataformaCache(clavePlataforma)
    End If
    getPlataformaByID_Clave = plataformabuscada
End Function

' =================================================================
' MAPEO: De Fila (Excel) hacia Objeto (Memoria)
' =================================================================

Private Sub mapearFilaAObjeto(ByRef fila As ListRow, ByRef Plataforma As clsPlataforma, ByRef idxs() As Long)
    With fila
        Plataforma.ID_clavaPlataforma = .Range(idxs(1)).value
        Plataforma.nombrePlataforma = .Range(idxs(2)).value
        Plataforma.descripcion = .Range(idxs(3)).value
    End With
End Sub

' =================================================================
' INFRAESTRUCTURA DE CACHÉ DICTIONARY Y TABLA EN EXCEL
' =================================================================

Private Sub initDictionaryPlataformasCache()
    If Not dictPlataformaCache Is Nothing Then Exit Sub
    Call actualizarCache
End Sub

Private Sub actualizarCache()
    Dim tblPlataformas As ListObject
    Set tblPlataformas = ObtenerTablaPlataformas
    If tblPlataformas Is Nothing Then
        MsgBox "Error Crítico: No se encontró la tabla PLATAFORMAS '" & modConfig.TABLA_ROLES & "'.", vbCritical, "Error PlataformasDAO"
        End
    End If

    Dim row As ListRow, plataformaObj As clsPlataforma, idxs() As Long

    Set dictPlataformaCache = New Scripting.Dictionary
    dictPlataformaCache.CompareMode = TextCompare
    idxs = getIndexes(tblPlataformas)

    For Each row In tblPlataformas.ListRows
        Set plataformaObj = New clsPlataforma
        Call mapearFilaAObjeto(row, plataformaObj, idxs)
        If dictPlataformaCache.Exists(plataformaObj.ID_clavaPlataforma) Then
            MsgBox "Error de Integridad: Registro duplicado exacto detectado en la TABLA PLATAFORMAS, fila " & row.Index & _
                   " (" & plataformaObj.ID_clavaPlataforma & ")." & vbCrLf & "El programa se detendrá.", vbCritical, "Datos Corruptos"
            Call invalidarCache
            End
        Else
            dictPlataformaCache.Add plataformaObj.ID_clavaPlataforma, plataformaObj
        End If
    Next row
End Sub

Private Function getIndexes(ByRef tblPlataformas As ListObject) As Long()
    Dim idxs() As Long
    ReDim idxs(1 To modConfig.TABLA_PLATAFORMAS_COLUMN_COUNT)

    On Error GoTo ErrColumnas
    idxs(1) = tblPlataformas.ListColumns(modConfig.COL_PLATAFORMAS_ID_CLAVE_PLATAFORMA).Index
    idxs(2) = tblPlataformas.ListColumns(modConfig.COL_PLATAFORMAS_NOMBRE).Index
    idxs(3) = tblPlataformas.ListColumns(modConfig.COL_PLATAFORMAS_DESCRIPCION).Index
    getIndexes = idxs
    Exit Function
ErrColumnas:
    MsgBox "Error Crítico: Estructura de columnas incorrecta en la tabla de Plataformas." & vbCrLf & _
           "La macro se detendrá inmediatamente.", vbCritical, "Error de Configuración, modPlataformasDAO"
    Call invalidarCache
    End
End Function

Private Function ObtenerTablaPlataformas() As ListObject
    Set ObtenerTablaPlataformas = Nothing
    On Error Resume Next
    Set ObtenerTablaPlataformas = Application.Evaluate(modConfig.TABLA_PLATAFORMAS)
    On Error GoTo 0
End Function

Public Sub invalidCache()
    Set dictPlataformaCache = Nothing
End Sub
