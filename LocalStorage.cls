VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LocalStorageFilePath = "C:\rtmacros\Data\LocalStorage.json"


Public Function GetItem(Key As Variant) As String
    
    If CreateFilePath(LocalStorageFilePath) = False Then
        Exit Function
    End If
    
    Dim Data As Scripting.Dictionary
    Set Data = JsonParse(ReadTextFile(LocalStorageFilePath))
    
    If Data.Exists(Key) Then
        GetItem = JsonStringify(Data(Key))
    End If

End Function

Public Sub SetItem(Key As Variant, Value As Variant)
    
    If CreateFilePath(LocalStorageFilePath) = False Then
        Exit Sub
    End If
    
    Dim Data As Variant
    Set Data = JsonParse(ReadTextFile(LocalStorageFilePath))
    
    Data(Key) = Value
    WriteToTextFile LocalStorageFilePath, JsonStringify(Data, 2)

End Sub

Public Sub RemoveItem(Key As Variant)

    If CreateFilePath(LocalStorageFilePath) = False Then
        Exit Sub
    End If
    
    Dim Data As Scripting.Dictionary
    Set Data = JsonParse(ReadTextFile(LocalStorageFilePath))
    
    If Data.Exists(Key) Then
        Data.Remove (Key)
    End If

    WriteToTextFile LocalStorageFilePath, JsonStringify(Data, 2)

End Sub

Public Sub Clear()
    
    If CreateFilePath(LocalStorageFilePath) = False Then
        Exit Sub
    End If
    
    Dim Data As Scripting.Dictionary
    Set Data = JsonParse(ReadTextFile(LocalStorageFilePath))

    Data.RemoveAll

    WriteToTextFile LocalStorageFilePath, JsonStringify(Data)
    
End Sub

Public Sub Display()
    
    If CreateFilePath(LocalStorageFilePath) = False Then
        Exit Sub
    End If
    Debug.Print ReadTextFile(LocalStorageFilePath)

End Sub



'******************************************************************************************
' PRIVATE METHODS
'******************************************************************************************
Private Function ReadTextFile(FilePath As String) As String
    
    Dim FSO As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = FSO.OpenTextFile(FilePath, ForReading, False)
    On Error Resume Next
    ReadTextFile = ts.ReadAll
    
    Set FSO = Nothing
    Set ts = Nothing
    
End Function

Private Function WriteToTextFile(FilePath As String, Value As String)
    
    Dim FSO As New FileSystemObject
    Dim ts As TextStream
    
    If CreateFilePath(FilePath) = True Then
        Set ts = FSO.OpenTextFile(FilePath, ForWriting, True)
        ts.Write Value
    End If
    
    Set FSO = Nothing
    Set ts = Nothing
    
End Function

Private Function CreateFilePath(FullPath As String) As Boolean

    Dim FSO As New FileSystemObject
    Dim I As Integer
    Dim sPath() As String
    Dim CurPath As String
    
    On Error GoTo CATCH
    sPath = Split(FullPath, "\")
    
    'CREATES EACH FOLDER PATH IF NEEDED
    For I = LBound(sPath, 1) To UBound(sPath, 1) - 1
        CurPath = CurPath & sPath(I) & "\"
        If Not FSO.FolderExists(CurPath) Then
            FSO.createFolder CurPath
        End If
    Next I
    
    'CREATES FILE IF NEEDED
    If Not FSO.FileExists(FullPath) Then
        FSO.CreateTextFile FullPath
    End If
    
    CreateFilePath = True
    Exit Function
CATCH:
    'RETURNS FALSE
End Function

