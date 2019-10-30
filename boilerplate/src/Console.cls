VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**
' * Various types of logs to immediate window and folder paths.
' * Saves to local and shared. Filename will be Year and Monthname
' *
' * <tr>
' *   <td><i class="material-icons" title="info|warning|error">info|warning|flag</i></td>
' *   <td>10/28/2019</td>
' *   <td>9:03:31 AM</td>
' *   <td>rtoda02</td>
' *   <td>Console.log</td>
' *   <td class="message">This is just a test message to see what it would look like.</td>
' * </tr>
' *
' * @ref {Modlue} FileSystemUtilities
' * @ref {Function} FileSystemUtilities.BuildOutFilePath
' * @ref {Function} FileSystemUtilities.ReadTextFile
' */
Option Explicit

Private Type state
    fso As Scripting.FileSystemObject
    folderPaths() As String
End Type

Private Const HTMLKEY As String = "<tbody>"
Private Const TEMPLATEPATH As String = "templates\console.html"

Private this As state

'/**
' * These are the folder paths that log files will be saved to. Add as needed.
' */
Private Sub setFolderPaths()
    Dim foldersPaths(0 To 1) As String
    foldersPaths(0) = "C:\test\SHARE\"
    foldersPaths(1) = "C:\test\LOCAL\"
    this.folderPaths = foldersPaths
End Sub

'/**
' * Create FileSystemObject on start.
' */
Private Sub Class_Initialize()
    On Error GoTo Catch
    Set this.fso = New Scripting.FileSystemObject
    setFolderPaths
Catch:
End Sub

'/** Standard message */
Public Sub Log(ByVal source As String, ByVal message As String)
    logMessage "[+]", "<i class=""material-icons"" title=""info"">info</i>", message, source
End Sub

'/** Error message */
Public Sub Error(ByVal source As String, ByVal message As String)
    logMessage "[X]", "<i class=""material-icons"" title=""error"">flag</i>", message, source
End Sub

'/** Warning message */
Public Sub Warn(ByVal source As String, ByVal message As String)
    logMessage "[!]", "<i class=""material-icons"" title=""warning"">warning</i>", message, source
End Sub

'/** Main source for logging message. Called from all log methods. */
Private Sub logMessage(ByVal debugLevel As String, ByVal HTMLlevel As String, ByRef message As String, ByVal source As String)
    ' Log message to immediate window.
    Debug.Print debugLevel & " [" & Time & "] [" & source & "] " & message
    
    Dim htmlMessage As String
    htmlMessage = "<tr>" & _
                     "<td>" & HTMLlevel & "</td>" & _
                     "<td>" & Date & "</td>" & _
                     "<td>" & Time & "</td>" & _
                     "<td>" & Environ("Username") & "</td>" & _
                     "<td>" & source & "</td>" & _
                     "<td class=""message"">" & message & "</td>" & _
                  "</tr>"
    
    ' Log to each folder path
    Dim index As Long
    For index = LBound(this.folderPaths) To UBound(this.folderPaths)
        injectHtmlToFile this.fso.BuildPath(this.folderPaths(index), fileName), htmlMessage
    Next index
End Sub

Public Property Get TemplateFilePath() As String
    TemplateFilePath = this.fso.GetAbsolutePathName(this.fso.BuildPath(ThisWorkbook.Path, TEMPLATEPATH))
End Property

Private Property Get fileName() As String
    fileName = "Log-Report-" & Year(Date) & "-" & MonthName(Month(Date)) & ".html"
End Property

Private Sub injectHtmlToFile(ByVal filePath As String, ByVal htmlMessage As String)
    ' Need to create copy of log file if it doesn't exist.
    If Not this.fso.FileExists(filePath) Then
        ' Template file must exist to work.
        If Not this.fso.FileExists(TemplateFilePath) Then
            Debug.Print "[X] [injectHtmlToFile] Not able to log to file. Notify developer the template file does not exist."
            Exit Sub
        End If
        ' Must be able to build file path to work.
        If BuildOutFilePath(filePath) = False Then
            Debug.Print "[X] [injectHtmlToFile] Not able to log to file. Unable to build filePath: " & filePath
            Exit Sub
        End If
        
        this.fso.CopyFile TemplateFilePath, filePath, True
    End If
    
    ' Only will get to this section if the filePath exists. Ok to inject to HTMLKEY.
    Dim html As String
    html = ReadTextFile(filePath)
    html = Replace(html, HTMLKEY, HTMLKEY & vbNewLine & htmlMessage)
    
    ' Write back to log the appended html.
    WriteToTextFile filePath, html
End Sub
