Option Explicit
Option Base 0

Sub SaveGlobalFile(FileName As String, SourceTab As String)

Dim FSO, TS As Object
Dim line As String
Dim lRow, lCol As Integer
Dim sht As Worksheet
Dim CurrentStringArray() As Variant
Dim i, j As Integer

Set FSO = CreateObject("Scripting.FileSystemObject")
Set TS = FSO.CreateTextFile(FileName)

If SourceTab = "Params" Then
    Set sht = EditTables
ElseIf SourceTab = "Tables" Then
    Set sht = WSTables
End If

lRow = sht.Cells(Rows.count, 1).End(xlUp).Row

For i = 5 To lRow
    lCol = sht.Cells(i, Columns.count).End(xlToLeft).column
    ReDim Preserve CurrentStringArray(lCol)
    
    For j = 1 To lCol
        CurrentStringArray(j - 1) = sht.Cells(i, j)
    Next j
    
    line = Join(CurrentStringArray, ",")
    line = Left(line, Len(line) - 1) & vbNewLine
    TS.Write line
    
Next i

TS.Close

MsgBox "The table: " & GetFileNameFromPath(FileName) & " has been saved sucessfully!"

End Sub

Function folderPath(Optional Level As Integer) As String

Dim arrNames As Variant
Dim i As Integer

    arrNames = Split(ThisWorkbook.path, "\")
    
    For i = 0 To UBound(arrNames, 1) - Level
        If i = UBound(arrNames, 1) - Level Then
            folderPath = folderPath & arrNames(i)
        Else
            folderPath = folderPath & arrNames(i) & "\"
        End If
    Next i

End Function

Sub ReadGlobalFile(SheetName As String, file_name As String)

Dim ws As Worksheet
Dim line As String
Dim FSO As Object
Dim TS As Object
Dim str As String
Dim Row, Col, i, ArrayLen As Integer
Dim StringArray As Variant

Set ws = ActiveWorkbook.Sheets(SheetName)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set TS = FSO.OpenTextFile(file_name)

'initial cell location
Row = 5
Col = 1

Do While Not TS.AtEndOfStream
    line = TS.ReadLine
    
    If InStr(line, ",") Then
        StringArray = Split(line, ",")
        ArrayLen = UBound(StringArray) - LBound(StringArray) + 1
        
        For i = 1 To ArrayLen
            ws.Cells(Row, Col) = StringArray(i - 1)
            Col = Col + 1
        Next i
    Else
        ws.Cells(Row, Col) = line
    End If
    
    Row = Row + 1
    Col = 1
    
Loop

TS.Close
Set TS = Nothing
Set FSO = Nothing

tablePath = file_name

End Sub
