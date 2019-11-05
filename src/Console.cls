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

'**/
' * {string} HTMLKEY This is where the record will be append below.
' * {string} TEMPLATEPATH is where the template is located. Required to work!
' * {Type} this The local state of this class.
' */
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

' Standard message
Public Sub Log(ByVal source As String, ByVal message As String)
    logMessage "[+]", "<i class=""material-icons"" title=""info"">info</i>", message, source
End Sub

' Error message
Public Sub Error(ByVal source As String, ByVal message As String)
    logMessage "[X]", "<i class=""material-icons"" title=""error"">flag</i>", message, source
End Sub

' Warning message
Public Sub Warn(ByVal source As String, ByVal message As String)
    logMessage "[!]", "<i class=""material-icons"" title=""warning"">warning</i>", message, source
End Sub

' Main source for logging message. Called from all log methods.
' This
Private Sub logMessage(ByVal debugLevel As String, ByVal HTMLlevel As String, ByRef message As String, ByVal source As String)
    ' Log message to immediate window.
    Debug.Print debugLevel & " [" & Time & "] [" & source & "] " & message
    
    ' Get html message template. This will be appended to the HTML file.
    Dim htmlMessage As String
    htmlMessage = "<tr>" & _
                     "<td>" & HTMLlevel & "</td>" & _
                     "<td>" & Date & "</td>" & _
                     "<td>" & Time & "</td>" & _
                     "<td>" & Environ("Username") & "</td>" & _
                     "<td>" & source & "</td>" & _
                     "<td class=""message"">" & message & "</td>" & _
                  "</tr>"
    
    ' Append each file added the htmlMessage to the top of the list.
    Dim index As Long
    For index = LBound(this.folderPaths) To UBound(this.folderPaths)
        injectHtmlToFile this.fso.BuildPath(this.folderPaths(index), fileName), htmlMessage
    Next index
End Sub

' This is where the template if located.
' If this file is missing and the user does not have a
' copy on their system then logs will not be recorded.
Public Property Get TemplateFilePath() As String
    TemplateFilePath = this.fso.GetAbsolutePathName(this.fso.BuildPath(ThisWorkbook.Path, TEMPLATEPATH))
End Property

' Create a new name by using the current year and month.
' @example Log-Report-2019-November.html
Private Property Get fileName() As String
    fileName = "Log-Report-" & Year(Date) & "-" & MonthName(Month(Date)) & ".html"
End Property

' This appends the htmlMessage to the specified file.
' The message is inserted right below the starting body tag `<tbody>`.
' This ensures the latest logs are at the top.
Private Sub injectHtmlToFile(ByVal FILEPATH As String, ByVal htmlMessage As String)
    ' Need to create copy of log file if it doesn't exist.
    If Not this.fso.FileExists(FILEPATH) Then
        ' Template file must exist to work.
        ' If it doesn't then log to immediate window
        ' and exit sub.
        ' Note, this will be failing silently to any user
        ' including the developers. But this also ensure no
        ' unwanted errors occur.
        If Not this.fso.FileExists(TemplateFilePath) Then
            Debug.Print "[X] [injectHtmlToFile] Not able to log to file. Notify developer the template file does not exist."
            Exit Sub
        End If
        ' Must be able to build file path to work.
        If BuildOutFilePath(FILEPATH) = False Then
            Debug.Print "[X] [injectHtmlToFile] Not able to log to file. Unable to build filePath: " & FILEPATH
            Exit Sub
        End If
        
        this.fso.CopyFile TemplateFilePath, FILEPATH, True
    End If
    
    ' Only will get to this section if the filePath exists. Ok to inject to HTMLKEY.
    Dim html As String
    html = ReadTextFile(FILEPATH)
    html = Replace(html, HTMLKEY, HTMLKEY & vbNewLine & htmlMessage)
    
    ' Write back to log the appended html.
    WriteToTextFile FILEPATH, html
End Sub
