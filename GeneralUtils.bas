Option Explicit
Option Base 0
Option Private Module

Function isWorkbookOpen(ByVal strWorkbookName As String) As Boolean
Dim Wb As Workbook

    On Error Resume Next
    Set Wb = Workbooks(strWorkbookName)
    If err Then
        isWorkbookOpen = False
    Else
        isWorkbookOpen = True
    End If
    
End Function

Function insertRowIntoData(baseData, rowNumber, dataToInsert)
    
    Dim newData As Variant
    Dim r, c As Integer

    ReDim newData(LBound(baseData, 1) To UBound(baseData, 1) + 1, LBound(baseData, 2) To UBound(baseData, 2))
    
    For r = LBound(newData) To UBound(newData) Step 1
        For c = LBound(newData, 2) To UBound(newData, 2) Step 1
            If r < rowNumber Then
                newData(r, c) = baseData(r, c)
            ElseIf r = rowNumber Then
                newData(r, c) = dataToInsert(c)
            Else
                newData(r, c) = baseData(r - 1, c)
            End If
        Next c
    Next r
    
    insertRowIntoData = newData

End Function

Function getTableInWorkbook(Wb, tblName)

Dim ws As Worksheet
Dim lo As ListObject

For Each ws In Wb.Sheets
    For Each lo In ws.ListObjects
        If lo.Name = tblName Then
            Set getTableInWorkbook = ws.ListObjects(lo.Name)
            Exit Function
        End If
    Next lo
Next ws
Set getTableInWorkbook = Nothing

End Function

Function getValueFromSpecificColumnOfTable(table, columnName, rowIdx)
Dim colIdx As Integer
    colIdx = getColumnIdx(table, columnName)
    getValueFromSpecificColumnOfTable = table.DataBodyRange(rowIdx, colIdx).Value
End Function

Public Function HeaderExists(tableName As ListObject, headerName As String) As Boolean
'TRUE value if column (headerName) exists in specified table (tableName)
'''''''''''''''''' Variable declaration ''''''''''''''''''''''''
    Dim hdr As ListColumn
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo NotExist
        Set hdr = tableName.ListColumns(headerName)
        HeaderExists = True
    Exit Function

'Error Handler
NotExist:
    HeaderExists = False

End Function

Function AddNewVersion(ByVal strTableName As String, ByRef arrData As Variant)
    Dim tblVersionNumber As ListObject
    Dim NewRow As ListRow

    Set tblVersionNumber = Range(strTableName).ListObject
    Set NewRow = tblVersionNumber.ListRows.Add(AlwaysInsert:=True)

    If TypeName(arrData) = "Range" Then
        NewRow.Range = arrData.Value
    Else
        NewRow.Range = arrData
    End If
End Function

Function LastRow(Sh As Worksheet, Col As String)

    On Error Resume Next
    
    LastRow = Sh.Columns(Col).Find(What:="*", after:=Sh.Range("CreateMainFolder"), _
        LookAt:=xlPart, LookIn:=xlValues, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False).Row
                            
    On Error GoTo 0
    
End Function

Function ParentFolder(path As String) As String

    Static objFSO As Object
   
    If objFSO Is Nothing Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    End If

    ParentFolder = objFSO.GetParentFolderName(path)
    
End Function

Sub RefreshAllQueries()

    Dim con As WorkbookConnection
    Dim Cname As String
    
    For Each con In ActiveWorkbook.Connections
        If Left(con.Name, 8) = "Query - " Then
        Cname = con.Name
            With ActiveWorkbook.Connections(Cname).OLEDBConnection
                '.BackgroundQuery = False
                .Refresh
            End With
        End If
    Next

End Sub

Function selectNonEmptyRows(arr As Variant, colIdx As Integer) As Variant

Dim i, k, n As Integer
Dim newArr() As Variant
'Dim finArr() As Variant
Dim aha As String

n = UBound(arr) - LBound(arr)

k = -1
For i = 0 To n
    If arr(i, colIdx) <> "" Then
        k = k + 1
        ReDim Preserve newArr(colIdx, k)
        newArr(0, k) = arr(i, 0)
        newArr(1, k) = arr(i, 1)
    End If
Next i

'ReDim finArr(UBound(newArr, 2) - LBound(newArr, 2), UBound(newArr, 1) - LBound(newArr, 1))
'finArr = Application.WorksheetFunction.Transpose(newArr)
'aha = finArr(1, 0)

If k = -1 Then
    ReDim Preserve newArr(0, 0)
    newArr(0, 0) = "No data!"
End If

selectNonEmptyRows = newArr()

End Function

Function groupByArray(arr As Variant, column As Integer) As Variant

Dim i, j, k, l, m, n, o, p As Integer
Dim newArr() As Variant
Dim a As String

k = 0
n = UBound(arr, 1) - LBound(arr, 1)
m = UBound(arr, 2) - LBound(arr, 2)

ReDim Preserve newArr(n, 0)

For j = 1 To n + 1
    newArr(j - 1, 0) = arr(j - 1, 0)
Next j

For i = 1 To m
    p = -1
    o = UBound(newArr, 2) - LBound(newArr, 2)
    For l = 0 To o
        If newArr(0, l) = arr(0, i) Then
            p = l
            Exit For
        End If
    Next l
    
    If p = -1 Then
        k = k + 1
        ReDim Preserve newArr(n, k)
        newArr(0, k) = arr(0, i)
        newArr(1, k) = arr(1, i)
    Else
        newArr(1, p) = newArr(1, p) & vbNewLine & arr(1, i)
    End If
Next i


groupByArray = newArr()

End Function

Function TableSheetName(Wb1 As Workbook, table As String) As String
Dim sheet As Worksheet
Dim TableTemp As ListObject

    TableSheetName = vbNullString
    
    For Each sheet In Wb1.Worksheets
        On Error Resume Next
        Set TableTemp = sheet.ListObjects(table)
        If err.Number = 0 Then
            TableSheetName = sheet.Name
            Exit Function
        Else
            err.Clear
        End If
    Next sheet
End Function

' ------------------------------------------------------
' ApplicationFreeze()

Sub ApplicationFreeze()
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
End Sub

' ------------------------------------------------------
' ApplicationUnfreeze()

Sub ApplicationUnfreeze()
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = ""
    End With
End Sub

Function LastColumn(Sh As Worksheet, Row As Integer)

    On Error Resume Next
    
    'LastRow = SH.Columns(COL).Find(What:="*", After:=SH.Range("CreateMainFolder"), _
        LookAt:=xlPart, LookIn:=xlValues, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False).row
                            
    LastColumn = Sh.Rows(Row).Find(What:="*", after:=Sh.Cells(Row, 1), _
        LookAt:=xlPart, LookIn:=xlValues, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False).column
        
    On Error GoTo 0
    
End Function

Function LastRow2(Sh As Worksheet, Col As Integer)

    On Error Resume Next
                            
    LastRow2 = Sh.Columns(Col).Find(What:="*", after:=Sh.Cells(1, Col), _
        LookAt:=xlPart, LookIn:=xlValues, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False).Row
        
    On Error GoTo 0
    
End Function

Function CreateSubtableFromWorkbook(strToFind As String, Wb As Workbook, tblName As String, vlookupcol As String)

Dim tblIn As ListObject
Dim arrOut() As Variant
Dim arrFin() As Variant
Dim x, y As String
Dim i, j, n, m, k As Integer

Set tblIn = getTableInWorkbook(Wb, tblName)

n = tblIn.Range.Rows.count - 1
m = tblIn.Range.Columns.count - 1

ReDim Preserve arrOut(m, 0)

k = -1
For i = 1 To n
    x = getValueFromSpecificColumnOfTable(tblIn, vlookupcol, i)
        If (x = strToFind) Then
            k = k + 1
            ReDim Preserve arrOut(m, k)
            For j = 0 To m
                arrOut(j, k) = tblIn.DataBodyRange.Cells(i, j + 1).Value
            Next j
        End If
Next i

m = UBound(arrOut, 1) - LBound(arrOut, 1)
n = UBound(arrOut, 2) - LBound(arrOut, 2)

ReDim Preserve arrFin(n, m)

For i = 0 To n
    For j = 0 To m
        arrFin(i, j) = arrOut(j, i)
    Next j
Next i

CreateSubtableFromWorkbook = arrFin

End Function

Function CreateSubtable(strToFind As String, tblIn As ListObject, vlookupcol As String)

Dim arrOut() As Variant
Dim arrFin() As Variant
Dim x, y As String
Dim i, j, n, m, k As Integer


n = tblIn.Range.Rows.count - 1
m = tblIn.Range.Columns.count - 1

ReDim Preserve arrOut(m, 0)

k = -1
For i = 1 To n
    x = getValueFromSpecificColumnOfTable(tblIn, vlookupcol, i)
       If (x = strToFind) Then
            k = k + 1
            ReDim Preserve arrOut(m, k)
            For j = 0 To m
                arrOut(j, k) = tblIn.DataBodyRange.Cells(i, j + 1).Value
            Next j
        End If
Next i

m = UBound(arrOut, 1) - LBound(arrOut, 1)
n = UBound(arrOut, 2) - LBound(arrOut, 2)

ReDim Preserve arrFin(n, m)

For i = 0 To n
    For j = 0 To m
        arrFin(i, j) = arrOut(j, i)
    Next j
Next i

CreateSubtable = arrFin

End Function

Function CreateSubArray(strToFind As String, ArrIn As Variant, vlookupcol As Integer)

Dim tblIn As ListObject
Dim arrOut() As Variant
Dim arrFin() As Variant
Dim x, y As String
Dim i, j, n, m, k As Integer


n = UBound(ArrIn, 1) - LBound(ArrIn, 1)
m = UBound(ArrIn, 2) - LBound(ArrIn, 2)

ReDim Preserve arrOut(m, 0)

k = -1
For i = 0 To n
    x = ArrIn(i, vlookupcol)
        If (x = strToFind) Then
            k = k + 1
            ReDim Preserve arrOut(m, k)
            For j = 0 To m
                arrOut(j, k) = ArrIn(i, j)
            Next j
        End If
Next i

m = UBound(arrOut, 1) - LBound(arrOut, 1)
n = UBound(arrOut, 2) - LBound(arrOut, 2)

ReDim Preserve arrFin(n, m)

For i = 0 To n
    For j = 0 To m
        arrFin(i, j) = arrOut(j, i)
    Next j
Next i

CreateSubArray = arrFin

End Function


Function MergeArrays(Array1 As Variant, Array2 As Variant) As Variant

Dim LB1, UB1, LB2, UB2, i, z As Integer
Dim Array3() As Variant

LB1 = LBound(Array1)
UB1 = UBound(Array1)
LB2 = LBound(Array2)
UB2 = UBound(Array2)

z = -1

    For i = LB1 To UB1
        z = z + 1
        ReDim Preserve Array3(z)
        Array3(z) = Array1(i)
    Next i
    
    For i = LB2 To UB2
        z = z + 1
        ReDim Preserve Array3(z)
        Array3(z) = Array2(i)
    Next i
    
MergeArrays = Array3

End Function

Function RemoveDupesColl(MyArray As Variant) As Variant
'DESCRIPTION: Removes duplicates from your array using the collection method.
'NOTES: (1) This function returns unique elements in your array, but
' it converts your array elements to strings.
'SOURCE: https://wellsr.com
'-----------------------------------------------------------------------
    Dim i As Long
    Dim n As Integer
    Dim arrColl As New Collection
    Dim arrDummy() As Variant
    Dim arrDummy1() As Variant
    Dim item As Variant
    ReDim arrDummy1(LBound(MyArray) To UBound(MyArray))
    
    For i = LBound(MyArray) To UBound(MyArray) 'convert to string
        arrDummy1(i) = CStr(MyArray(i))
    Next i
    On Error Resume Next
    For Each item In arrDummy1
       arrColl.Add item, item
    Next item
    err.Clear
    ReDim arrDummy(LBound(MyArray) To arrColl.count + LBound(MyArray) - 1)
    i = LBound(MyArray)
    For Each item In arrColl
       arrDummy(i) = item
       i = i + 1
    Next item
    RemoveDupesColl = arrDummy
End Function


Function checkDataProvided(Col As Variant, arr As Variant)

Dim i, j, m, n As Integer
Dim idx As Integer

n = UBound(arr) - LBound(arr)
m = UBound(Col) - LBound(Col)

idx = 0
For i = 0 To n
    For j = 0 To m
        If arr(i, Col(j)) = "" Then
            idx = 1
            Exit For
        End If
    Next j
    If idx = 1 Then
        Exit For
    End If
Next i

checkDataProvided = idx

End Function

Function selectColumnFromArray(arr As Variant, colIdx As Integer)

Dim i, n As Integer
Dim arrFin() As Variant

n = UBound(arr) - LBound(arr) + 1
ReDim arrFin(n - 1)

For i = 1 To n
    arrFin(i - 1) = arr(i - 1, colIdx)
Next i

selectColumnFromArray = arrFin

End Function

Function CheckIfAnyDataProvided(arr As Variant)

Dim i, j, m As Integer
Dim n As Long
Dim idx As Integer

n = UBound(arr, 1) - LBound(arr, 1) + 1
m = UBound(arr, 2) - LBound(arr, 2) + 1

idx = 0
For i = 2 To n
    For j = 1 To m
        If arr(i, j) <> "" Then
            idx = 1
            GoTo skipLoop
        End If
    Next j
Next i

skipLoop:
CheckIfAnyDataProvided = idx
End Function


Public Function selectFolder()

Dim sFolder As String

    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    selectFolder = ""
    If sFolder <> "" Then ' if a file was chosen
        selectFolder = sFolder
    End If
    
End Function

Function isInitialised(ByRef a() As Variant) As Boolean
    On Error Resume Next
    isInitialised = IsNumeric(UBound(a))
    On Error GoTo 0
End Function

Function cutString(str As String, num As Integer)

Dim finStr As String

If IsEmpty(str) Or str = "" Then
    finStr = str
ElseIf Len(str) > num Then
    finStr = Left(str, num)
    finStr = finStr & "..."
Else
    finStr = str
End If

cutString = finStr

End Function

Function checkIfAnyCellIsEmptyInTableCol(table As ListObject, colName As String) As Boolean

Dim i, n As Integer
Dim temp_cell_val As String

n = table.DataBodyRange.Rows.count
checkIfAnyCellIsEmptyInTableCol = False

For i = 1 To n
    temp_cell_val = getValueFromSpecificColumnOfTable(table, colName, i)
    If temp_cell_val = "" Then
        GoTo Breaker
    End If
Next i
    
Exit Function

Breaker:
    checkIfAnyCellIsEmptyInTableCol = True
    Exit Function

End Function

Function IsAllElementsTheSame(arr As Variant) As Boolean
    Dim buf As Variant
    buf = arr(0)
    Dim i As Integer
    i = 0
    Do While (buf = arr(i) And i < UBound(arr))
    i = i + 1
    Loop
    IsAllElementsTheSame = False
    If i = UBound(arr) And buf = arr(UBound(arr)) Then
        IsAllElementsTheSame = True
    End If
End Function

Function SortArrayAtoZ(MyArray As Variant)

Dim i As Long
Dim j As Long
Dim Temp

'Sort the Array A-Z
For i = LBound(MyArray) To UBound(MyArray) - 1
    For j = i + 1 To UBound(MyArray)
        If UCase(MyArray(i)) > UCase(MyArray(j)) Then
            Temp = MyArray(j)
            MyArray(j) = MyArray(i)
            MyArray(i) = Temp
        End If
    Next j
Next i

SortArrayAtoZ = MyArray

End Function

Function ArrayIsEmpty(arr() As Variant) As Boolean

Dim k As Integer

    On Error GoTo SkipAll
    k = UBound(arr)
    
    ArrayIsEmpty = False
    
    Exit Function
    
SkipAll:
ArrayIsEmpty = True
    
End Function

Function IsInArray(StringToBeFound As String, MyArray As Variant) As Boolean

IsInArray = UBound(Filter(MyArray, StringToBeFound)) > -1

End Function
