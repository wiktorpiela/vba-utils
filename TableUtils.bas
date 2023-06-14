Option Explicit
Option Base 0

Function getColumnIdx(tableObj, columnHeader)
' This function returns the index of a named column from a specified table

On Error GoTo ErrorHandler
    getColumnIdx = tableObj.ListColumns(columnHeader).Index
    Exit Function
    
ErrorHandler:
    MsgBox ("The column named " & columnHeader & " could not be found in the table")

End Function

Function getRowIdx(tableObj, colLookupName, rowValue)

Dim rowIdx, i As Integer
Dim colValues As Variant

    colValues = tableObj.ListColumns(colLookupName).DataBodyRange
    If IsArray(colValues) Then
        For i = 1 To UBound(colValues)
            If (colValues(i, 1) = rowValue) Then
                getRowIdx = i
                Exit For
            End If
        Next i
    Else
        getRowIdx = 1
    End If
    
End Function

Function getRowIdx3Columns(tableObj, colName As String, colType As String, colProdAcc As String, rowValueName As String, rowValueType As String, rowValueProdAcc As String) As Integer

Dim rowIdx, i, redim_counter, temp_int As Integer
Dim colValues, rowIdxsName, rowIdxsType, rowIdxsProdAcc  As Variant
Dim temp_name As String

'get idexes with proper structure name
redim_counter = -1
colValues = tableObj.ListColumns(colName).DataBodyRange
If IsArray(colValues) Then
    For i = 1 To UBound(colValues)
        If (colValues(i, 1) = rowValueName) Then
            redim_counter = redim_counter + 1
            If redim_counter = 0 Then
                ReDim rowIdxsName(redim_counter)
                rowIdxsName(redim_counter) = i
            Else
                ReDim Preserve rowIdxsName(redim_counter)
                rowIdxsName(redim_counter) = i
            End If
        End If
    Next i
End If
    
'get idexes with proper structure type for current structure name
redim_counter = -1
colValues = tableObj.ListColumns(colType).DataBodyRange
If IsArray(colValues) Then
    For i = 0 To UBound(rowIdxsName)
        If colValues(rowIdxsName(i), 1) = rowValueType Then
            redim_counter = redim_counter + 1
            If redim_counter = 0 Then
                ReDim rowIdxsType(redim_counter)
                rowIdxsType(redim_counter) = rowIdxsName(i)
            Else
                ReDim Preserve rowIdxsType(redim_counter)
                rowIdxsType(redim_counter) = rowIdxsName(i)
            End If
        End If
    Next i
End If
    
'wczesc wspolna obu tabel
redim_counter = -1
For i = 0 To UBound(rowIdxsType)
    temp_int = rowIdxsType(i)
    If Not IsError(Application.Match(temp_int, rowIdxsName, 0)) Then
        redim_counter = redim_counter + 1
        If redim_counter = 0 Then
            ReDim rowIdxsProdAcc(redim_counter)
            rowIdxsProdAcc(redim_counter) = temp_int
        Else
            ReDim Preserve rowIdxsProdAcc(redim_counter)
            rowIdxsProdAcc(redim_counter) = temp_int
        End If
    End If
Next i
    
colValues = tableObj.ListColumns(colProdAcc).DataBodyRange
If IsArray(colValues) Then
    For i = 0 To UBound(rowIdxsProdAcc)
        If (colValues(rowIdxsProdAcc(i), 1) = rowValueProdAcc) Then
            getRowIdx3Columns = rowIdxsProdAcc(i)
            Exit For
        End If
    Next
End If
    
End Function

Function getRowIdx2Columns(tableObj, colName1 As String, colName2 As String, rowValue1 As String, rowValue2 As String) As Integer

Dim rowIdx, i, redim_counter, temp_int As Integer
Dim colValues, rowIdxsName, rowIdxsType, rowIdxsProdAcc  As Variant
Dim temp_name As String

'get idexes with proper accumulation name
redim_counter = -1
colValues = tableObj.ListColumns(colName1).DataBodyRange
If IsArray(colValues) Then
    For i = 1 To UBound(colValues)
        If (colValues(i, 1) = rowValue1) Then
            redim_counter = redim_counter + 1
            If redim_counter = 0 Then
                ReDim rowIdxsName(redim_counter)
                rowIdxsName(redim_counter) = i
            Else
                ReDim Preserve rowIdxsName(redim_counter)
                rowIdxsName(redim_counter) = i
            End If
        End If
    Next i
End If
    
'get idexes with proper products type for current accumulation name
redim_counter = -1
colValues = tableObj.ListColumns(colName2).DataBodyRange
If IsArray(colValues) Then
    For i = 0 To UBound(rowIdxsName)
        If colValues(rowIdxsName(i), 1) = rowValue2 Then
            redim_counter = redim_counter + 1
            If redim_counter = 0 Then
                ReDim rowIdxsType(redim_counter)
                rowIdxsType(redim_counter) = rowIdxsName(i)
            Else
                ReDim Preserve rowIdxsType(redim_counter)
                rowIdxsType(redim_counter) = rowIdxsName(i)
            End If
        End If
    Next i
End If
    
'wczesc wspolna obu tabel
redim_counter = -1
For i = 0 To UBound(rowIdxsType)
    temp_int = rowIdxsType(i)
    If Not IsError(Application.Match(temp_int, rowIdxsName, 0)) Then
        redim_counter = redim_counter + 1
        If redim_counter = 0 Then
            ReDim rowIdxsProdAcc(redim_counter)
            rowIdxsProdAcc(redim_counter) = temp_int
        Else
            ReDim Preserve rowIdxsProdAcc(redim_counter)
            rowIdxsProdAcc(redim_counter) = temp_int
        End If
    End If
Next i
    
getRowIdx2Columns = rowIdxsProdAcc(0)
    
End Function

Function getArrRowIdx(arr, colIdx, rowValue)

Dim rowIdx, i, n As Integer

n = UBound(arr) - LBound(arr)
For i = 0 To n
    If (arr(i, colIdx) = rowValue) Then
        getArrRowIdx = i
        Exit For
    End If
Next i
    
End Function

Function getValueFromTable(tableObj, colName, colLookupName, rowValue)
' This function will vlookup a value from a the specified table. The arguments:
'    - tableObj - the entire table to look through
'    - colName - the name of the column from which you return a value
'    - colLookupName - the name of the column you want to search for a value in.
'    - rowValue - the value you are looking for within colLookupName

' Note it works in the same way as a vlookup for non-unique values i.e. it will return the first value that it comes across.

Dim colIdx, rowIdx, i As Integer
Dim colValues As Variant

colIdx = tableObj.ListColumns(colName).Index
colValues = tableObj.ListColumns(colLookupName).DataBodyRange
If IsArray(colValues) Then
    For i = 1 To UBound(colValues)
        If (colValues(i, 1) = rowValue) Then
            rowIdx = i
            getValueFromTable = tableObj.DataBodyRange(rowIdx, colIdx).Value
            Exit For
        End If
    Next i
Else
    getValueFromTable = tableObj.ListColumns(colIdx).DataBodyRange.Value
End If
    
End Function

Function GetArrayFromFilteredRange(rng As Range) As Variant
'--------------------------
Dim i As Long
Dim j As Long
Dim Row As Range
Dim arr As Variant
'--------------------------
'If 0 results in Filter just exit
If Not rng.SpecialCells(xlCellTypeVisible).count > 0 Then Exit Function
i = 1
ReDim arr(1 To rng.Columns.count, 1 To rng.Columns(1).SpecialCells(xlCellTypeVisible).count)

For Each Row In rng.Rows
    If Not Row.Hidden Then
        For j = LBound(arr, 1) To UBound(arr, 1)
            arr(j, i) = Row.Cells(j)
        Next j
        i = i + 1
    End If
Next Row
GetArrayFromFilteredRange = arr
End Function

