Attribute VB_Name = "modUtilerias"
' --- DENTRO DE UN MÓDULO ESTÁNDAR (Ej: modUtilerias) ---
Option Explicit

' =================================================================
' FUNCIÓN: CalcularRfcCorto
' OBJETIVO: Recibe un RFC de 13 caracteres y devuelve su equivalente de 8
' =================================================================
Public Function CalcularRfcCorto(ByVal rfcCompleto As String) As String
    Dim rfcLimpio As String
    Dim letras As String
    Dim anio As String
    Dim mesTexto As String, diaTexto As String
    Dim mesNum As Integer, diaNum As Integer
    Dim mesCorto As String, diaCorto As String
    
    ' Lista de caracteres de sustitución (Posición 1=1, Posición 9=9, Posición 10=A, etc.)
    ' El espacio al inicio es para que el índice coincida exactamente con el número de día/mes
    Dim mapaCaracteres As String
    mapaCaracteres = " 123456789ABCDEFGHIJKLMNOPQRSTUV"
    
    ' 1. Limpieza básica
    rfcLimpio = UCase(Trim(rfcCompleto))
    
    ' Validación básica de longitud (Debe ser un RFC de persona física de 13 caracteres)
    If Len(rfcLimpio) < 13 Then
        CalcularRfcCorto = ""
        Exit Function
    End If
    
    ' 2. Extraer los componentes del RFC completo
    letras = Left(rfcLimpio, 4)           ' Primeros 4 caracteres (Letras)
    anio = Mid(rfcLimpio, 5, 2)             ' Siguientes 2 (Año: posiciones 5 y 6)
    mesTexto = Mid(rfcLimpio, 7, 2)         ' Siguientes 2 (Mes: posiciones 7 y 8)
    diaTexto = Mid(rfcLimpio, 9, 2)         ' Siguientes 2 (Día: posiciones 9 y 10)
    
    ' 3. Convertir mes y día de texto a número para poder indexar
    mesNum = CInt(mesTexto)
    diaNum = CInt(diaTexto)
    
    ' 4. Obtener el equivalente usando el mapa de caracteres
    ' Si mesNum es 9, Mid(..., 9, 1) da "9". Si es 12, Mid(..., 12, 1) da "C"
    mesCorto = Mid(mapaCaracteres, mesNum + 1, 1)
    diaCorto = Mid(mapaCaracteres, diaNum + 1, 1)
    
    ' 5. Armar y retornar el RFC Corto (4 letras + 2 año + 1 mes + 1 día = 8 caracteres)
    CalcularRfcCorto = letras & anio & mesCorto & diaCorto
End Function

' =================================================================
' FUNCIÓN INVERSA: CalcularRfcLargo
' OBJETIVO: Recibe un RFC Corto de 8 caracteres y devuelve los primeros
'           13 caracteres del RFC original (Letras + Fecha AAMMDD)
' =================================================================
Public Function CalcularRfcLargo(ByVal rfcCorto As String) As String
    Dim rfcCortoLimpio As String
    Dim letras As String
    Dim anio As String
    Dim mesCorto As String, diaCorto As String
    Dim mesNum As Integer, diaNum As Integer
    Dim mesTexto As String, diaTexto As String
    Dim mapaCaracteres As String
    
    ' El mismo mapa de caracteres que usamos para comprimir
    mapaCaracteres = " 123456789ABCDEFGHIJKLMNOPQRSTUV"
    
    ' 1. Limpieza y validación de longitud estricta
    rfcCortoLimpio = UCase(Trim(rfcCorto))
    If Len(rfcCortoLimpio) <> 8 Then
        CalcularRfcLargo = ""
        Exit Function
    End If
    
    ' 2. Desestructurar el RFC Corto
    letras = Left(rfcCortoLimpio, 4)       ' Las primeras 4 letras
    anio = Mid(rfcCortoLimpio, 5, 2)         ' Los 2 dígitos del año
    mesCorto = Mid(rfcCortoLimpio, 7, 1)     ' El carácter del mes
    diaCorto = Mid(rfcCortoLimpio, 8, 1)     ' El carácter del día
    
    ' 3. Encontrar la posición matemática (Invertir el algoritmo)
    ' InStr nos dice en qué posición de la cadena está el carácter.
    ' Restamos 1 debido al espacio inicial del mapa.
    mesNum = InStr(1, mapaCaracteres, mesCorto) - 1
    diaNum = InStr(1, mapaCaracteres, diaCorto) - 1
    
    ' Validación por si viene un carácter corrupto que no está en el mapa
    If mesNum <= 0 Or diaNum <= 0 Then
        CalcularRfcLargo = ""
        Exit Function
    End If
    
    ' 4. Formatear a dos dígitos agregando el cero a la izquierda si es necesario
    mesTexto = Format(mesNum, "00")
    diaTexto = Format(diaNum, "00")
    
    ' 5. Armar el RFC de 13 caracteres
    CalcularRfcLargo = letras & anio & mesTexto & diaTexto & "XXX"
End Function
