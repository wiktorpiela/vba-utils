Public Sub saveArrayAsCSV(MyArray As Variant, sFileName As String, Optional sDelimiter As String = ",", Optional removeQuotes = 0)

Dim n As Long 'counter
Dim m As Long 'counter
Dim sCSV As String 'csv string to print

    On Error GoTo ErrHandler_SaveAsCSV
      'save the file
      Open sFileName For Output As #1
      For n = LBound(MyArray, 1) To UBound(MyArray, 1)
        sCSV = ""
        For m = LBound(MyArray, 1) To UBound(MyArray, 2)
            If removeQuotes = 1 Then
                sCSV = sCSV & MyArray(n, m) & sDelimiter
            Else
                If TypeName(MyArray(n, m)) = "String" Then
                    sCSV = sCSV & """" & MyArray(n, m) & """" & sDelimiter
                ElseIf TypeName(MyArray(n, m)) = "Date" Then
                    sCSV = sCSV & """" & Format(MyArray(n, m), "YYYY-MM-DD") & """" & sDelimiter
                Else
                    sCSV = sCSV & MyArray(n, m) & sDelimiter
                End If
            End If
          'sCSV = sCSV & """" & Format(MyArray(n, m)) & """" & sDelimiter
        Next m
        sCSV = Left(sCSV, Len(sCSV) - 1) 'remove last Delimiter
        Print #1, sCSV
      Next n
      Close #1
ErrHandler_SaveAsCSV:
      Close #1
End Sub

Sub Export_Table_To_CSV(Wb As Workbook, tableName As String, FilePath As String)

    Dim wbNew As Workbook
    Dim tblExport As ListObject
    
    Set tblExport = getTableInWorkbook(Wb, tableName)

    'If file already exists, delete it
    If Dir$(FilePath) <> "" Then
        Kill FilePath
    End If
    
    Set wbNew = Workbooks.Add
    
    With wbNew
        tblExport.Range.Copy
        .Sheets("Sheet1").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        .SaveAs FileName:=FilePath, FileFormat:=xlCSV, CreateBackup:=False
        .Close
    End With

End Sub

Sub saveTableToCSV(Wb As Workbook, tableName As String, csvFilePath As String)
'----------------------
Dim tbl As ListObject
Dim fNum As Integer
Dim tblArr
Dim rowArr
Dim csvVal
Dim colArr
'------------------------

    Set tbl = getTableInWorkbook(Wb, tableName)
   
    'tblArr = tbl.DataBodyRange.Value
    tblArr = tbl.Range.Value

    fNum = FreeFile()
    Open csvFilePath For Output As #fNum
    For i = 1 To UBound(tblArr)
        rowArr = Application.Index(tblArr, i, 0)
        csvVal = VBA.Join(rowArr, ",")
        Print #1, csvVal
    Next i
    Close #fNum
    Set tblArr = Nothing
    Set rowArr = Nothing
    Set csvVal = Nothing
End Sub


Function readCsvIntoArray(FileName)

Dim fNum As Integer
Dim wholeFile As String
Dim lines As Variant
Dim oneLine As Variant
Dim numRows As Long
Dim numCols As Long
Dim theArray() As String
Dim r As Long
Dim c As Long

    ' Load the file.
    fNum = FreeFile
    Open FileName For Input As fNum
    wholeFile = Input$(LOF(fNum), #fNum)
    Close fNum

    ' Break the file into lines.
    lines = Split(wholeFile, vbCrLf)

    ' Dimension the array.
    numRows = UBound(lines)
    oneLine = Split(lines(0), ",")
    numCols = UBound(oneLine)
    ReDim theArray(numRows - 1, numCols)

    ' Copy the data into the array.
    For r = 0 To numRows
        If Len(lines(r)) > 0 Then
            oneLine = Split(lines(r), ",")
            For c = 0 To numCols
                theArray(r, c) = oneLine(c)
            Next c
        End If
    Next r
readCsvIntoArray = theArray
End Function
