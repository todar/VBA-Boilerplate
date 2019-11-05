VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**/
' * Store values on users machine.
' *
' * @ref {Library} Microsoft Scripting Runtime
' * @ref {Class Module} Console - Used to log to immediate window and log files.
' * @ref {Class Module} JSON - Used to store key value pairs.
' * @ref {Modlue} FileSystemUtilities
' * @ref {Function} FileSystemUtilities.BuildOutFilePath
' * @ref {Function} FileSystemUtilities.ReadTextFile
' */
Option Explicit

'**/ The location of the saved file. */
Private Property Get localStorageFilePath() As String
    localStorageFilePath = "C:\Data\LocalStorage.json"
End Property

'**/
' * Gets an item from a key.
' * @param {String} key The key of the value.
' * @param {String} [fallback] A default value in case the key doesn't exist.
' */
Public Function GetItem(ByVal Key As String, ByVal fallback As Variant) As Variant
    On Error GoTo Catch
    If BuildOutFilePath(localStorageFilePath) = False Then
        GetItem = fallback
        Exit Function
    End If
    
    Dim data As Scripting.Dictionary
    Set data = JSON.Parse(ReadTextFile(localStorageFilePath))
    
    If data.Exists(Key) Then
        GetItem = data(Key)
    Else
        GetItem = fallback
    End If
    Exit Function
Catch:
    GetItem = fallback
    Console.Error "LocalStorage.GetItem", Err.description
End Function

' Set a key value pair to the saved text file.
Public Sub SetItem(ByVal Key As String, ByVal value As Variant)
    On Error GoTo Catch
    If BuildOutFilePath(localStorageFilePath) = False Then
        Exit Sub
    End If
    
    ' Read text file. If null then set to an empty string dictionary object
    Dim jsonValue As String
    jsonValue = ReadTextFile(localStorageFilePath)
    jsonValue = IIf(jsonValue = vbNullString, "{}", jsonValue)
    
    Dim data As Variant
    Set data = JSON.Parse(jsonValue)
    
    data(Key) = value
    WriteToTextFile localStorageFilePath, JSON.Stringify(data, 2)
    
    Exit Sub
Catch:
    Console.Error "LocalStorage.SetItem", Err.description
End Sub

' Remove an item based on the key
Public Sub RemoveItem(ByVal Key As Variant)
    On Error GoTo Catch
    If BuildOutFilePath(localStorageFilePath) = False Then
        Exit Sub
    End If
    
    Dim data As Scripting.Dictionary
    Set data = JSON.Parse(ReadTextFile(localStorageFilePath))
    
    If data.Exists(Key) Then
        data.Remove (Key)
    End If

    WriteToTextFile localStorageFilePath, JSON.Stringify(data, 2)
    
    Exit Sub
Catch:
    Console.Error "LocalStorage.RemoveItem", Err.description
End Sub

' Clear file contents to an empty dictionary string
Public Sub Clear()
    On Error GoTo Catch
    If BuildOutFilePath(localStorageFilePath) = False Then
        Exit Sub
    End If
    
    WriteToTextFile localStorageFilePath, "{}"
    Exit Sub
Catch:
    Console.Error "LocalStorage.Clear", Err.description
End Sub

' Easy way of reading the contents in the log
Public Sub DisplayToImmediateWindow()
    On Error GoTo Catch
    If BuildOutFilePath(localStorageFilePath) = False Then
        Exit Sub
    End If
    
    Debug.Print ReadTextFile(localStorageFilePath)
    Exit Sub
Catch:
    Console.Error "LocalStorage.DisplayToImmediateWindow", Err.description
End Sub



