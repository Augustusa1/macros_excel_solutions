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

Sub MatchBySimilarity()
    Dim wsOficial As Worksheet, wsNoOficial As Worksheet
    Dim lastRowOficial As Long, lastRowNoOficial As Long
    Dim i As Long, j As Long
    Dim minDistance As Integer, currentDistance As Integer
    Dim bestMatch As String
    
    ' Configurar hojas
    Set wsOficial = ThisWorkbook.Sheets("Hoja1") ' Nombres oficiales
    Set wsNoOficial = ThisWorkbook.Sheets("Hoja2") ' Nombres no oficiales
    
    ' Últimas filas
    lastRowOficial = wsOficial.Cells(wsOficial.Rows.Count, 1).End(xlUp).Row
    lastRowNoOficial = wsNoOficial.Cells(wsNoOficial.Rows.Count, 1).End(xlUp).Row
    
    ' Iterar sobre nombres no oficiales
    For i = 2 To lastRowNoOficial
        minDistance = 9999
        bestMatch = ""
        
        ' Comparar con cada nombre oficial
        For j = 2 To lastRowOficial
            currentDistance = Levenshtein(wsNoOficial.Cells(i, 1).Value, wsOficial.Cells(j, 1).Value)
            If currentDistance < minDistance Then
                minDistance = currentDistance
                bestMatch = wsOficial.Cells(j, 1).Value
            End If
        Next j
        
        ' Asignar el mejor match
        wsNoOficial.Cells(i, 2).Value = bestMatch ' Columna B para el resultado
    Next i
    
    MsgBox "Reemplazo completado", vbInformation
End Sub

