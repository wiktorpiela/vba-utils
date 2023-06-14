Option Explicit
Option Base 0

Sub ReadTextFileTable(FilePath as String, InitialRowIndex as Integer, InitialColIndex as Integer)

Dim line As String
Dim FSO As Object
Dim TS As Object
Dim str As String
Dim sht As Worksheet
Dim StringArray As Variant
Dim Row, Col As Integer
Dim ArrayLen As Integer
Dim i As Integer

Set FSO = CreateObject("Scripting.FileSystemObject")
Set TS = FSO.OpenTextFile(FilePath)
Set sht = textFile

'initial cell location
Row = InitialRowIndex
Col = InitialColIndex

Do While Not TS.AtEndOfStream

    line = TS.ReadLine
    
    If InStr(line, ",") Then
        StringArray = Split(line, ",")
        ArrayLen = UBound(StringArray) - LBound(StringArray) + 1
        
        For i = 1 To ArrayLen
            sht.Cells(Row, Col) = StringArray(i - 1)
            Col = Col + 1
        Next i
    Else
        sht.Cells(Row, Col) = line
    End If
    
    Row = Row + 1
    Col = InitialColIndex
    
Loop

TS.Close
Set TS = Nothing
Set FSO = Nothing

End Sub