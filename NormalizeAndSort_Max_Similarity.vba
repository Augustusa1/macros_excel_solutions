Function NormalizeAndSort(inputString As String) As String
    Dim words() As String
    Dim sortedString As String
    Dim i As Long

    ' Normalizar: convertir a minúsculas, eliminar espacios adicionales y caracteres especiales
    inputString = LCase(Trim(inputString))
    inputString = Replace(inputString, ".", "")
    inputString = Replace(inputString, ",", "")
    inputString = Replace(inputString, "-", "")
    inputString = Replace(inputString, "  ", " ")

    ' Dividir en palabras
    words = Split(inputString, " ")

    ' Ordenar las palabras
    For i = LBound(words) To UBound(words) - 1
        Dim j As Long
        For j = i + 1 To UBound(words)
            If words(i) > words(j) Then
                Dim temp As String
                temp = words(i)
                words(i) = words(j)
                words(j) = temp
            End If
        Next j
    Next i

    ' Reunir palabras ordenadas
    sortedString = Join(words, " ")
    NormalizeAndSort = sortedString
End Function

Function Levenshtein(str1 As String, str2 As String) As Integer
    Dim i As Integer, j As Integer
    Dim d() As Integer
    Dim n As Integer, m As Integer
    Dim cost As Integer

    n = Len(str1)
    m = Len(str2)

    ReDim d(0 To n, 0 To m)

    For i = 0 To n
        d(i, 0) = i
    Next i

    For j = 0 To m
        d(0, j) = j
    Next j

    For i = 1 To n
        For j = 1 To m
            If Mid(str1, i, 1) = Mid(str2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            d(i, j) = WorksheetFunction.Min(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i

    Levenshtein = d(n, m)
End Function

Function SimilarityPercent(str1 As String, str2 As String) As Double
    Dim levDist As Integer
    levDist = Levenshtein(str1, str2)
    SimilarityPercent = 1 - (levDist / Application.Max(Len(str1), Len(str2)))
End Function

Sub MatchBySimilarity()
    Dim wsOficial As Worksheet, wsNoOficial As Worksheet
    Dim lastRowOficial As Long, lastRowNoOficial As Long
    Dim i As Long, j As Long
    Dim maxSimilarity As Double, currentSimilarity As Double
    Dim bestMatch As String
    Dim similarityThreshold As Double
    
    ' Configurar hojas
    Set wsOficial = ThisWorkbook.Sheets("Hoja1") ' Nombres oficiales
    Set wsNoOficial = ThisWorkbook.Sheets("Hoja2") ' Nombres no oficiales
    
    ' Configurar umbral de similitud
    similarityThreshold = 0.5 ' Ajustar según necesidad (porcentaje mínimo de similitud)

    ' Últimas filas
    lastRowOficial = wsOficial.Cells(wsOficial.Rows.Count, 1).End(xlUp).Row
    lastRowNoOficial = wsNoOficial.Cells(wsNoOficial.Rows.Count, 1).End(xlUp).Row
    
    ' Iterar sobre nombres no oficiales
    For i = 2 To lastRowNoOficial
        maxSimilarity = 0
        bestMatch = "Sin coincidencias"
        
        ' Comparar con cada nombre oficial
        For j = 2 To lastRowOficial
            ' Normalizar y ordenar palabras antes de comparar
            currentSimilarity = SimilarityPercent(NormalizeAndSort(wsNoOficial.Cells(i, 1).Value), NormalizeAndSort(wsOficial.Cells(j, 1).Value))
            If currentSimilarity > maxSimilarity Then
                maxSimilarity = currentSimilarity
                bestMatch = wsOficial.Cells(j, 1).Value
            End If
        Next j
        
        ' Asignar el mejor match si supera el umbral
        If maxSimilarity >= similarityThreshold Then
            wsNoOficial.Cells(i, 2).Value = bestMatch
            wsNoOficial.Cells(i, 3).Value = maxSimilarity ' Agregar columna con porcentaje de similitud
        Else
            wsNoOficial.Cells(i, 2).Value = "Sin coincidencias"
            wsNoOficial.Cells(i, 3).Value = maxSimilarity
        End If
    Next i
    
    MsgBox "Reemplazo completado", vbInformation
End Sub

