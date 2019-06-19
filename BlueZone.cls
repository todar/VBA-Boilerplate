VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlueZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'DEPENDENCIES
' - CreateObject("bluezone.System")

'NOTES:
' - THIS CLASS IS USED TO AUTOMATE BLUEZONE AND EXTRA.
' - sendkeys: http://docs.attachmate.com/extra/x-treme/apis/com/5250functionkeys_des.htm
' - Index: http://docs.attachmate.com/extra/x-treme/apis/com/index.htm#sendkeysmethod_con.htm

Private Type TBlueZone
    Sessions As Object
    System As Object
    Session As Object
    Screen As Object
End Type

Private this As TBlueZone

Private Sub Class_Initialize()
    'This class relys 100% on the object "bluezone.System"
    'If it can't find it must raise an error!
    CreateBlueZoneObject
End Sub

Private Sub CreateBlueZoneObject()
    
    On Error GoTo Catch
    Set this.System = CreateObject("bluezone.System")
    Exit Sub
Catch:
    Err.Raise 429, "Bluezone Class", "BlueZone.System Library not found on current system."
    
End Sub

Public Property Get ActiveSession() As String
    ActiveSession = this.Session
End Property

Public Property Get Session() As Object
    Set Session = this.Session
End Property

Public Property Get Screen() As Object
    Set Screen = this.Screen
End Property

Public Function TryConnectToActiveSession(Optional PromptForLogin As Boolean _
                                        , Optional ByRef outConnectionError As String) As Boolean
    
    Set this.Sessions = this.System.Sessions
    Set this.Session = this.System.ActiveSession
    
    'CHECK TO SEE IF SESSION WAS CREATED TO BLUEZONE
    If this.Session Is Nothing Then
        outConnectionError = "BlueZone is not open."
        Exit Function
    End If
    
    Set this.Screen = this.Session.Screen
    
    'Optional check to see if user is logged into session
    If PromptForLogin Then
        If CheckForLogin = False Then
            outConnectionError = "User is not logged into BlueZone."
            Exit Function
        End If
    End If
    
    TryConnectToActiveSession = True
    
End Function

'CHECKS TO MAKE SURE USER IS LOGED INTO BLUEZONE
Private Function CheckForLogin() As Boolean
    
CheckAgain:
    'CHECK TO SEE IF CURRENTLY IN THE LOGIN SCREEN
    If Me.GetString(16, 19, 13) = "New Password:" Then
        
        'PROMPT FOR USER TO LOGIN
        If MsgBox("Please Login to BlueZone and then click OK to continue", vbOKCancel) = vbCancel Then
            Exit Function
        Else
            'OK PRESSED, CHECK TO SEE THAT THEY DID LOGIN
            GoTo CheckAgain
        End If
    Else
        CheckForLogin = True
    End If
    
End Function

Public Sub MaximizeScreen()
    this.Session.WindowState = 2
End Sub

Public Sub MinimizeScreen()
    this.Session.WindowState = 0
End Sub

Public Sub RestoreDownScreen()
    this.Session.WindowState = 1
End Sub

Public Sub ResetErrors()
    this.Screen.Sendkeys ("<RESET>")
End Sub

'SENDTEXT TO CURRENT SESSION AT CURRENT CURSOR POSITION
Public Sub Sendkeys(ByVal Value As String)
    this.Screen.Sendkeys Value
End Sub

'SETS CURSOR TO A SPECIFIC LOCATION IN BLUEZONE, GOOD FOR USER INTERACTION
Public Sub SetCursor(ByVal Row As Integer, ByVal Column As Integer)
    this.Screen.MoveTo Row, Column
End Sub

'PUTS STRING TO A SPECIFIC LOCATION IN BLUEZONE
Public Function PutString(ByVal Value As String, ByVal Row As Integer, ByVal Column As Integer) As Boolean
    this.Screen.PutString Value, Row, Column
End Function

'RETURNS A STRING FROM THE CURRENT SCREEN
Public Function GetString(ByVal Row As Integer, ByVal Column As Integer, ByVal Length As Integer)
     GetString = this.Screen.GetString(Row, Column, Length)
End Function

'RETURNS THE FULL PAGE STRING
Public Function GetFullScreenText() As String
    GetFullScreenText = Trim(this.Screen.Area(1, 1, this.Screen.Rows, this.Screen.Cols))
End Function

'SELECTS TEXT ON THE CURRENT SCREEN
Public Sub SelectText(ByVal Row As Integer, ByVal Column As Integer, ByVal Length As Integer)
    this.Screen.Select Row, Column, Row, Column + Length
End Sub

'CHECK TO SEE IF CURRENT PAGE CONTAINS A STRING VALUE.
Public Function PageContains(Value) As Boolean
    
    If GetFullScreenText Like "*" & Value & "*" Then
        PageContains = True
    End If
    
End Function

'WAITS FOR STRING TO APPEAR ON SCREEN, WILL TIME OUT ONCE IT HITS MAXSECONDS
Public Function WaitForStringOnPage(ByVal Value As String, Optional maxSeconds As Integer = 3) As Boolean
    
    Do
        If PageContains(Value) Then
            WaitForStringOnPage = True
            Exit Function
        End If
        
        If maxTimeLimitReached(maxSeconds) Then
            Exit Function
        End If
        
    Loop

End Function

'USED FOR GETTING BLANK SCREEN TO HAVE A CLEAR LINE OF NAVIGATION
Public Function ClearScreen() As Boolean
    
    'SENDKEYS FOR CLEARING OUT SCREEN
    Me.Sendkeys ("<PF3><PF3><PF3><CLEAR><CLEAR><CLEAR><CLEAR><CLEAR>")
    
    'RUN LOOP CHECKING FOR A BLANK SCREEN
    Do
        'CHECK FOR BLANK SCREEN
        If Me.GetFullScreenText = "" Then
            ClearScreen = True
            Exit Function
        End If
        
        'MAX TIME OF RUNNING LOOP IS THREE SECONDS
        If maxTimeLimitReached(3) Then
            Exit Function
        End If
    Loop
    
End Function


'USED TO HAVE A TIME LIMIT ON HOW LONG TO LOOK FOR A TEXT IN EXTRA (RAN IN LOOPS)
Private Function maxTimeLimitReached(maxSeconds As Integer) As Boolean
    
    Static StartTime As Single
    
    If StartTime = 0 Then
        StartTime = Timer
    End If
    
    If Timer >= StartTime + maxSeconds Then
        StartTime = 0
        maxTimeLimitReached = True
    End If
    
End Function


'WAITS FOR PAGE TO NO LONGER BE BLANK. WAS NEEDED FOR CLEARING OUT TO ENTER PACS.
'Private Function waitForPageToLoad(Optional maxSeconds As Integer = 3) As Boolean
'
'    Do
'        If GetScreenText <> "" Then
'            waitForPageToLoad = True
'            Exit Function
'        End If
'
'        If maxTimeLimitReached(maxSeconds) Then
'            Exit Function
'        End If
'    Loop
'
'End Function




