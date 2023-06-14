Option Explicit
Option Base 0

Public Function CheckFolderExists(folderPath As String) As Boolean
' Check whether a specified folder exists
    Dim strFolderName As String
    Dim strFolderExists As String
    
    On Error GoTo err

    strFolderName = folderPath
    strFolderExists = Dir(strFolderName, vbDirectory)

err:
    If strFolderExists = "" Then
        CheckFolderExists = False
    Else
        CheckFolderExists = True
    End If

End Function

Function GetFileNameFromPath(ByVal strPath As String) As String

    GetFileNameFromPath = ""
    If Len(strPath) > 0 Then
        GetFileNameFromPath = Right(strPath, Len(strPath) - InStrRev(strPath, "\"))
    End If


End Function
Function CheckFileExists(FileName As String) As Boolean
' Check whether a specified file exists
    Dim strFileName As String
    Dim strFileExists As String
    
    On Error GoTo err
    strFileName = FileName
    strFileExists = Dir(strFileName)
    
err:
   If strFileExists = "" Or strFileName = vbNullString Then
        CheckFileExists = False
    Else
        CheckFileExists = True
    End If

End Function

Function FindFiles(stirngToBeFound As String, filesArray As Variant) As Variant
'function finds all files with given string in folder

    Dim i As Integer
    Dim fileInArray As Variant
    Dim necessaryFiles() As Variant
    Dim filesFound As Integer
    
    filesFound = 0
    
    For Each fileInArray In filesArray

        If InStr(fileInArray, stirngToBeFound) <> 0 Then
            ReDim Preserve necessaryFiles(filesFound)
            necessaryFiles(filesFound) = fileInArray
            filesFound = filesFound + 1
        End If
    Next
    
    If filesFound = 0 Then
        ReDim necessaryFiles(0)
        necessaryFiles(0) = ""
        FindFiles = necessaryFiles
    Else
        FindFiles = necessaryFiles
    End If
    
End Function

Function FindFile(StringToBeFound As String, arr As Variant) As String

Dim el As Integer
Dim FileName As String
Dim arr_el As Variant

On Error GoTo Handler

el = 0
For Each arr_el In arr
    el = InStr(arr_el, StringToBeFound)
    If el <> 0 Then
      FileName = arr_el
      GoTo skipLoop
    End If
Next

skipLoop:
FindFile = FileName

Handler:
    Exit Function

End Function

Function FindFileByExtension(StringToBeFound As String, arr As Variant) As String

Dim el As String
Dim FileName As String
Dim arr_el As Variant

On Error GoTo Handler

el = 0
For Each arr_el In arr
    el = Right(arr_el, Len(StringToBeFound))
    If el = StringToBeFound Then
      FileName = arr_el
      GoTo skipLoop
    End If
Next

skipLoop:
FindFileByExtension = FileName

Handler:
    Exit Function

End Function


Function FindExactFileName(StringToBeFound As String, arr As Variant) As String

Dim el As Integer
Dim FileName As String
Dim arr_el As Variant

On Error GoTo Handler

el = 0
For Each arr_el In arr
    If arr_el = StringToBeFound Then
        FileName = arr_el
        GoTo skipLoop
    End If
Next

skipLoop:
FindExactFileName = FileName

Handler:
    Exit Function

End Function
Sub createFolder(folderPath As String)
    
    On Error GoTo ErrorHandler
    MkDir (folderPath)
    Exit Sub
    
ErrorHandler:
    MsgBox ("The folder " & folderPath & " could not be created. The macro will now terminate")
    End
End Sub
Function isFolderEmpty(FilePath As String, Optional fileExtension As Variant) As Boolean

    If IsMissing(fileExtension) Then
        fileExtension = ""
    End If
    
    'OutputFilePath = ws.Cells(1, "B").Value
    If Right(FilePath, 1) <> "\" Then
        FilePath = FilePath & "\"
    End If
    
    If Dir(FilePath & "*" & fileExtension) <> vbNullString Then
        isFolderEmpty = False
        Exit Function
    Else
        isFolderEmpty = True
        Exit Function
    End If

End Function

Function listFoldersInDirectory(dirPath)
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    Dim i As Integer
    Dim subFolder As Variant
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(dirPath)
    If objFolder.SubFolders.count <> 0 Then
        ReDim subFolder(1 To objFolder.SubFolders.count)
        i = 1
        'loops through each folder in the directory and stores their name
        For Each objSubFolder In objFolder.SubFolders
            subFolder(i) = objSubFolder.Name
            i = i + 1
        Next objSubFolder
        listFoldersInDirectory = subFolder
    Else
        listFoldersInDirectory = Empty
    End If
        
End Function


Function listFilesInDirectory(dirPath)
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim i As Integer
    Dim fileNames As Variant
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(dirPath)
    
    On Error GoTo ErrorHandler 'a folder is empty
    ReDim fileNames(1 To objFolder.Files.count)
    i = 1
    'loops through each file in the directory and prints their names and path
    For Each objFile In objFolder.Files
        fileNames(i) = objFile.Name
        i = i + 1
    Next objFile
        
    listFilesInDirectory = fileNames
    Exit Function

ErrorHandler:
    MsgBox "The " & dirPath & " is empty!", vbExclamation
    Exit Function
End Function

Function IsOpenWorkbook(strWorkbookName As String, Optional bFullname As Boolean = False) As Boolean
'Purpose: check if there is already a file open with the same name

Dim Wb As Workbook
Dim strName As String
    IsOpenWorkbook = False
    For Each Wb In Workbooks
        If bFullname = False Then
            strName = Wb.Name
        Else
            strName = Wb.FullName
        End If
        If (StrComp(strName, strWorkbookName, vbTextCompare) = 0) Then
            IsOpenWorkbook = True
            Exit Function
        End If
    Next
End Function

Function FileInUse(sFileName As String) As Boolean
    On Error Resume Next
    Open sFileName For Binary Access Read Lock Read As #1
    Close #1
    FileInUse = IIf(err.Number > 0, True, False)
    On Error GoTo 0
End Function

Function PathToFile(pth As String) As String

Dim arrNames As Variant
Dim i As Integer

    arrNames = Split(pth, "\")
    
    For i = 0 To UBound(arrNames, 1) - 1
        If i = UBound(arrNames, 1) - 1 Then
            PathToFile = PathToFile & arrNames(i)
        Else
            PathToFile = PathToFile & arrNames(i) & "\"
        End If
    Next i

End Function
