VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuickAccessToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private xmlDoc

'Custom Class to be able to interact with Excels Quick Access Toolbar.
'Allows an easy way to append new icons to call macros.
'If file doesn't exist, or SharedControls node is not setup, has the option to copy a file over.
'
'@Author Robert Todar
Private Sub Class_Initialize()
  
  'CREATE XML OBJECT AND LOAD
  Set xmlDoc = CreateObject("Microsoft.XMLDOM")
  xmlDoc.Async = "False"

  'LOAD XML FROM FILE
  xmlDoc.Load (FilePath)

End Sub

'--------------------------------------------------------------------------
' PUBLIC PROPERTIES
'--------------------------------------------------------------------------
Public Property Get FileExists() As Boolean
  If CreateObject("Scripting.FileSystemObject").FileExists(FilePath) Then
      FileExists = True
  End If
End Property

Public Property Get FilePath() As String
  FilePath = "C:\Users\" & Environ("UserName") & "\AppData\Local\Microsoft\Office\Excel.officeUI"
End Property

Public Property Get RootNode() As Object
  Set RootNode = xmlDoc.SelectSingleNode("/mso:customUI/mso:ribbon/mso:qat/mso:sharedControls")
End Property

Public Property Get IsSetupToAppend() As Boolean
  If FileExists And Not RootNode Is Nothing Then
    IsSetupToAppend = True
  End If
End Property

'--------------------------------------------------------------------------
' PUBLIC FUNCTIONS
'--------------------------------------------------------------------------
Public Sub Append(Id, Label, Icon, Action)

  'EXIT IF NODE ALREADY EXISTS
  If Not xmlDoc.SelectSingleNode("/mso:customUI/mso:ribbon/mso:qat/mso:sharedControls/mso:button[@idQ='" & Id & "']") Is Nothing Then
    Exit Sub
  End If

  'Create a new node.
  Dim el As Object
  Set el = xmlDoc.CreateElement("mso:button")
  el.SetAttribute "idQ", Id
  el.SetAttribute "visible", "true"
  el.SetAttribute "label", Label
  el.SetAttribute "imageMso", Icon
  el.SetAttribute "onAction", Action

  'APPEND TO ROOT NODE
  RootNode.AppendChild el

  'SAVE UPDATES TO XML
  xmlDoc.Save FilePath

End Sub

'Get the UI File from a copy elsewhere
Public Function SetupFromCopy(SourceFilePath)
  
  Dim Fso As Object
  Set Fso = CreateObject("Scripting.FileSystemObject")
  
  If Fso.FileExists(SourceFilePath) Then
    Fso.CopyFile SourceFilePath, FilePath, True
  End If
  Set Fso = Nothing

  'RELOAD XML FROM FILE, DUE TO IT BEING A NEW FILE
  xmlDoc.Load (FilePath)

End Function



