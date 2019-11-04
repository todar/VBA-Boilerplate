Attribute VB_Name = "SecurityUtilities"
Option Explicit

'/**
' * Encode URI component
' */
Public Function EncodeURIComponent(ByVal value As String) As String
    Dim JS As New ScriptControl
    JS.Language = "JScript"
    EncodeURIComponent = JS.Run("encodeURIComponent", value)
    ' Optionally 2013 && >
    ' WorksheetFunction.EncodeURL(value)
End Function

'/**
' * Decode URI component
' */
Public Function DecodeURIComponent(ByVal value As String) As String
    Dim JS As New ScriptControl
    JS.Language = "JScript"
    DecodeURIComponent = JS.Run("decodeURIComponent", value)
End Function
