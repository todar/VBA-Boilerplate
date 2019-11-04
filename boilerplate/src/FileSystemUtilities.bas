Attribute VB_Name = "FileSystemUtilities"
'/**
' * Utility Library for using FileSystem.
' *
' * @ref {Microsoft Scripting Runtime}
' */
Option Explicit

'/** Opens files with their default application */
Public Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" ( _
ByVal hWnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

'/**
' * Build entire folder path. The Standard CreateFolder only works for one level.
' * @ref {Microsoft Scripting Runtime}
' * @param {String} fullPath - The path that needs to get created.
' * @returns {Boolean} True if no errors occured and path was created.
' */
Public Function BuildOutFilePath(ByVal fullPath As String) As Boolean
    On Error GoTo Catch
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    ' Clean path make sure to only have '\' in the path.
    Dim absolutePath As String
    absolutePath = fso.GetAbsolutePathName(fullPath)

    ' Split the folder into each folder name.
    Dim folderNames() As String
    folderNames = split(absolutePath, "\")
    
    ' Loop each folder and make folder path if it doesn't already exist.
    Dim index As Integer
    For index = LBound(folderNames, 1) To UBound(folderNames, 1) - 1
        ' This builds the path in steps.
        Dim currentPath As String
        currentPath = currentPath & folderNames(index) & "\"
        If Not fso.FolderExists(currentPath) Then
            fso.CreateFolder currentPath
        End If
    Next index
    
    ' Lastly, if a file was included, create it if it doesn't exist.
    If Len(fso.GetExtensionName(absolutePath)) > 0 Then
        If Not fso.FileExists(absolutePath) Then
            fso.CreateTextFile fullPath
        End If
    End If
    
    BuildOutFilePath = True
    Exit Function
Catch:
    ' Any errors will return false.
End Function

'/**
' * Attempts to Create a text file and write to it to see if user has write access.
' *
' * @example: HasWriteAccessToFolder("C:\Program Files") ~> True || False
' */
Public Function HasWriteAccessToFolder(ByVal folderPath As String) As Boolean
    ' Make sure folder exists, this function returns false if it does not
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Exit Function
    End If

    ' Get unique temp filepath, don't want to overwrite something that already exists
    Do
        Dim count As Integer
        Dim FILEPATH As String
        
        FILEPATH = fso.BuildPath(folderPath, "TestWriteAccess" & count & ".tmp")
        count = count + 1
    Loop Until Not fso.FileExists(FILEPATH)
    
    ' Attempt to create the tmp file, error returns false
    On Error GoTo Catch
    fso.CreateTextFile(FILEPATH).Write ("Test Folder Access")
    Kill FILEPATH
    
    ' No error, able to write to file; return true!
    HasWriteAccessToFolder = True
Catch:
End Function

'/**
' * Read any type of text file.
' * @ref {Microsoft Scripting Runtime}
' * @param {String} filePath - Path to the file to read.
' */
Public Function ReadTextFile(ByVal FILEPATH As String) As String
    On Error GoTo Catch
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject

    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FILEPATH, ForReading, False)
    ReadTextFile = ts.ReadAll
    Exit Function
Catch:
    ' If error in getting file, return empty string.
End Function

'/**
' * Write to any type of text file.
' * @ref {Microsoft Scripting Runtime}
' * @param {String} filePath - Path to the file to write to.
' */
Public Function WriteToTextFile(ByVal FILEPATH As String, ByVal value As String) As Boolean
    On Error GoTo Catch
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    If BuildOutFilePath(FILEPATH) = True Then
        Set ts = fso.OpenTextFile(FILEPATH, ForWriting, True)
        ts.Write value
        WriteToTextFile = True
    End If
    
    Set fso = Nothing
    Set ts = Nothing
    Exit Function
Catch:
    ' Errors will return false
End Function

'/**
' * Appends to any type of text file.
' * @ref {Microsoft Scripting Runtime}
' * @param {String} filePath - Path to the file to write to.
' */
Public Function AppendToTextFile(ByVal FILEPATH As String, ByVal message As String) As Boolean
    On Error GoTo Catch
    Dim fso As New FileSystemObject
    If Not fso.FileExists(FILEPATH) Then
        BuildOutFilePath FILEPATH
    End If
    
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FILEPATH, ForAppending, True)
    ts.WriteLine message
    
    AppendToTextFile = True
    Exit Function
Catch:
    ' Errors will return false.
End Function

'/**
' * Checks to see if file exists, then opens it if it does
' */
Public Function OpenAnyFile(ByVal FILEPATH As String) As Boolean
    ' Will only open files that exist
    Dim fso As New Scripting.FileSystemObject
    If fso.FileExists(FILEPATH) Then
        OpenAnyFile = True
        
        ' API function for opening files with their default program.
        Call ShellExecute(0, "Open", FILEPATH & vbNullString, _
        vbNullString, vbNullString, 1)
    End If
End Function

'/**
' * Checks to see if folder exists, then opens windows explorer to that path
' */
Public Function OpenFileExplorer(ByVal folderPath As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    If fso.FolderExists(folderPath) Then
        OpenFileExplorer = True
        Call Shell("explorer.exe " & Chr(34) & folderPath & Chr(34), vbNormalFocus)
    End If
End Function
