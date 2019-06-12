Attribute VB_Name = "InternetFunctions"
Option Explicit

'To Make early binding use: Microsoft Internet Controls

' Examples using tools
Private Sub TestingJavascript()
    
    Dim IE As InternetExplorer
    Set IE = NewInternetExplorer("www.google.com")
    If Not IE Is Nothing Then
        InjectJavascript IE, "document.querySelectorAll('.gb_d')[0].click();"
    End If
    
    Dim Doc As mshtml.HTMLDocument
    
End Sub

' A simple factory for creating an Internet Explorer Object
Public Function NewInternetExplorer(ByVal URL As String, Optional ByVal Visible As Boolean = True) As InternetExplorer
    
    ' @author: Robert Todar <robert@roberttodar.com>
    ' @example: Set IE = NewInternetExplorer("www.google.com")
    ' @ref: Microsoft Internet Controls
    
    Set NewInternetExplorer = CreateObject("InternetExplorer.Application")
    With NewInternetExplorer
        .Visible = Visible
        .Navigate URL
    End With
    
    WaitForInternetExplorer NewInternetExplorer
   
End Function

' Returns Instance of IE, Looks to find first matching URL
Public Function GetOpenInternetExplorer(ByVal URL As String) As InternetExplorer
    
    ' @author: Robert Todar <robert@roberttodar.com>
    ' @example: Set IE = GetOpenInternetExplorer("www.google.com")
    ' @ref: Microsoft Internet Controls
    
    Dim Window As Object
    For Each Window In CreateObject("Shell.Application").Windows
        If InStr(Window.LocationURL, URL) > 0 Then
            Set GetOpenInternetExplorer = Window
            Exit Function
        End If
    Next Window
    
End Function

' Waits for Internet Explorer to have readystate 4 and not busy
Public Sub WaitForInternetExplorer(ByRef IE As InternetExplorer)

    ' @ref: Microsoft Internet Controls
    While IE.readyState <> 4 Or IE.busy: DoEvents: Wend
    
End Sub

' Simple way to execute scripts in IE.
Public Sub InjectJavascript(ByVal IE As InternetExplorer, ByVal Code As String)
    
    ' @ref: Microsoft Internet Controls
    IE.Document.parentWindow.execScript Code:=Code
    
End Sub



''Late Binding Examples below

'' Examples using tools
'Private Sub TestingJavascript()
'
'    Dim IE As Object
'    Set IE = NewInternetExplorer("www.google.com")
'    If Not IE Is Nothing Then
'        InjectJavascript IE, "document.querySelectorAll('.gb_d')[0].click();"
'    End If
'
'End Sub
'
'' A simple factory for creating an Internet Explorer Object
'Public Function NewInternetExplorer(ByVal URL As String, Optional ByVal Visible As Boolean = True) As Object
'
'    ' @author: Robert Todar <robert@roberttodar.com>
'    ' @example: Set IE = NewInternetExplorer("www.google.com")
'
'    Set NewInternetExplorer = CreateObject("InternetExplorer.Application")
'    With NewInternetExplorer
'        .Visible = Visible
'        .Navigate URL
'    End With
'
'    WaitForInternetExplore NewInternetExplorer
'
'End Function
'
'' Returns Instance of IE, Looks to find first matching URL
'Public Function GetOpenInternetExplorer(ByVal URL As String) As Object
'
'    ' @author: Robert Todar <robert@roberttodar.com>
'    ' @example: Set IE = GetOpenInternetExplorer("www.google.com")
'
'    Dim Window As Object
'    For Each Window In CreateObject("Shell.Application").Windows
'        If InStr(Window.LocationURL, URL) > 0 Then
'            Set GetOpenInternetExplorer = Window
'            Exit Function
'        End If
'    Next Window
'
'End Function
'
'' Waits for Internet Explorer to have readystate 4 and not busy
'Public Sub WaitForInternetExplore(ByRef IE As Object)
'    While IE.readyState <> 4 Or IE.busy: DoEvents: Wend
'End Sub
'
'' Simple way to execute scripts in IE.
'Public Sub InjectJavascript(ByVal IE As Object, ByVal Code As String)
'    IE.Document.parentWindow.execScript Code:=Code
'End Sub
