Attribute VB_Name = "modDevTools"
' --- DENTRO DE UN MÓDULO TEMPORAL: modDevTools ---
' (Solo se usa en desarrollo, no se lleva al trabajo)
Option Explicit

Public Sub ActualizarCodigoDesdeVSC()
    Dim rutaProyecto As String
    Dim componente As Object
    Dim fso As Object
    
    ' 1. La ruta de la carpeta donde trabajas en VS Code
    rutaProyecto = "E:\Roles\Roles_MVC\" ' <--- Ajusta esta ruta
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 2. Limpiar los módulos viejos para evitar duplicados (excepto este módulo)
    For Each componente In ThisWorkbook.VBProject.VBComponents
        If componente.Name <> "modDevTools" And componente.Type <> 3 Then ' 3 = UserForm (esos mejor manejarlos con cuidado)
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Remove componente
            On Error GoTo 0
        End If
    Next componente
    
    ' 3. Inyectar los archivos frescos de VS Code
    With ThisWorkbook.VBProject.VBComponents
        .Import rutaProyecto & "modConfig.bas"
        .Import rutaProyecto & "modUtilerias.bas"
        .Import rutaProyecto & "clsEmpleado.cls"
        .Import rutaProyecto & "modEmpleadoDAO.bas"
        .Import rutaProyecto & "modEmpleadoController.bas"
    End With
    
    MsgBox "¡Código de VS Code sincronizado en Excel!", vbInformation, "DevTools"
End Sub

Sub ExportarCodigoParaIA()
    Dim vbaProyecto As Object
    Dim componente As Object
    Dim fso As Object
    Dim archivoTexto As Object
    Dim rutaArchivo As String
    Dim lineaCodigo As String
    Dim i As Long
    
    ' Definir la ruta donde se guardará el archivo TXT (en la misma carpeta que este libro)
    If ThisWorkbook.Path = "" Then
        MsgBox "Por favor, guarda tu libro de Excel antes de ejecutar esta macro.", vbExclamation, "Libro no guardado"
        Exit Sub
    End If
    
    rutaArchivo = ThisWorkbook.Path & "\contexto_proyecto_vba.txt"
    
    ' Crear objetos para manejo de archivos
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set archivoTexto = fso.CreateTextFile(rutaArchivo, True, True) ' Unicode activado
    
    ' Escribir encabezado general del archivo
    archivoTexto.WriteLine "=================================================="
    archivoTexto.WriteLine " ESTRUCTURA Y CÓDIGO COMPLETO DEL PROYECTO VBA"
    archivoTexto.WriteLine " Generado el: " & Now
    archivoTexto.WriteLine "=================================================="
    archivoTexto.WriteLine ""
    
    ' Acceder al proyecto VBA del libro activo
    Set vbaProyecto = ThisWorkbook.VBProject
    
    ' Recorrer cada componente del proyecto
    For Each componente In vbaProyecto.VBComponents
        
        ' Filtrar solo Módulos Estándar (1) y Módulos de Clase (2)
        ' Si también quieres formularios (UserForms), añade: Or componente.Type = 3
        If componente.Type = 1 Or componente.Type = 2 Then
            
            ' Encabezado llamativo para la IA
            archivoTexto.WriteLine "=================================================="
            If componente.Type = 1 Then
                archivoTexto.WriteLine "[MÓDULO ESTÁNDAR]: " & componente.Name
            Else
                archivoTexto.WriteLine "[MÓDULO DE CLASE]: " & componente.Name
            End If
            archivoTexto.WriteLine "=================================================="
            
            ' Validar si el módulo tiene líneas de código
            If componente.CodeModule.CountOfLines > 0 Then
                ' Leer línea por línea e imprimirla en el TXT
                For i = 1 To componente.CodeModule.CountOfLines
                    lineaCodigo = componente.CodeModule.Lines(i, 1)
                    archivoTexto.WriteLine lineaCodigo
                Next i
            Else
                archivoTexto.WriteLine "' (Este módulo está vacío)"
            End If
            
            ' Espaciado entre módulos
            archivoTexto.WriteLine ""
            archivoTexto.WriteLine ""
        End If
    Next componente
    
    ' Cerrar archivo y limpiar memoria
    archivoTexto.Close
    Set archivoTexto = Nothing
    Set fso = Nothing
    
    ' Notificar al usuario
    MsgBox "¡Listo! El código se ha exportado con éxito a:" & vbCrLf & rutaArchivo, vbInformation, "Exportación Completada"
End Sub


Sub ExportarTodoElProyectoVBA()
    Dim componente As Object
    Dim rutaExportar As String
    Dim fso As Object
    Dim extension As String
    
    ' Verificar si el libro actual está guardado
    If ActiveWorkbook.Path = "" Then
        MsgBox "Por favor, guarda tu archivo de Excel antes de ejecutar esta macro.", vbCritical, "Archivo no guardado"
        Exit Sub
    End If
    
    ' Definir la ruta de la carpeta (en la misma ubicación del archivo de Excel)
    rutaExportar = ActiveWorkbook.Path & "\ARCHIVOS_EXPORTADOS\"
    
    ' Crear la carpeta si no existe usando FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rutaExportar) Then
        fso.CreateFolder (rutaExportar)
    End If
    
    ' Recorrer todos los componentes del proyecto de VBA
    For Each componente In ActiveWorkbook.VBProject.VBComponents
        
        ' Determinar la extensión según el tipo de componente
        Select Case componente.Type
            Case 1 ' vbext_ct_StdModule (Módulo Estándar)
                extension = ".bas"
            Case 2 ' vbext_ct_ClassModule (Módulo de Clase)
                extension = ".cls"
            Case 3 ' vbext_ct_MSForm (Formulario / UserForm)
                extension = ".frm"
            Case Else
                ' Hojas y ThisWorkbook (Tipo 100), se suelen ignorar porque no se pueden reimportar limpiamente
                extension = ""
        End Select
        
        ' Si el componente es exportable, lo exportamos
        If extension <> "" Then
            componente.Export rutaExportar & componente.Name & extension
        End If
        
    Next componente
    
    ' Limpieza de objetos
    Set fso = Nothing
    
    MsgBox "Todos los módulos, clases y formularios se han exportado con éxito en: " & vbCrLf & rutaExportar, vbInformation, "Exportación Completada"
End Sub

