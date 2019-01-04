VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "blueZoneObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'USED FOR GETTING MILLISECONDS
Private Declare Function GetSystemTimeAdjustment Lib "kernel32" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Long) As Long

'BLUEZONE\EXTRA OBJECTS
Private pSessions As Object
Private pSystem As Object
Private pSession As Object
Public pScreen As Object

'DEPENDENCIES
' - SOME OF THE TOOLS USE SCRIPTING.DICTIONARY. REFERENCE: MICROSOFT SCRIPTING RUNTIME

'Public Functions - General Tools:
' - init
' - Goto_17Set2
' - pageContains

'Public Functions - Specific Tools:

'Private Helper Funtions:
' - clearOutToABlankScreen
' - maxTimeLimitReached

'NOTES:
' - THIS CLASS IS USED TO AUTOMATE BLUEZONE AND EXTRA.
' - sendkeys: http://docs.attachmate.com/extra/x-treme/apis/com/5250functionkeys_des.htm
' - Index: http://docs.attachmate.com/extra/x-treme/apis/com/index.htm#sendkeysmethod_con.htm

'TODO:
' - N\A.

'******************************************************************************************
' EVENTS
'******************************************************************************************

Private Sub Class_Initialize()
    
    'SET CURRENT SESSION TO OBJECTS
    connectToActiveSession
   
End Sub

Private Sub Class_Terminate()
    
    'BEST PRACTICE OF CLEARING OUT MEMORY
    Set pSessions = Nothing
    Set pSystem = Nothing
    Set pSession = Nothing
    Set pScreen = Nothing
    
End Sub


'******************************************************************************************
' PUBLIC METHODS - GENERAL TOOLS
'******************************************************************************************

' - TO HELP SHORTEN THE WAY OF USING THE SCREENS SENDKEYS
Public Sub Sendkeys(Value As String)

    pScreen.Sendkeys Value
    
End Sub

' - SETS CURSOR TO A SPECIFIC LOCATION IN BLUEZONE, GOOD FOR USER INTERACTION
Public Sub setCursor(Row As Integer, col As Integer)

    pScreen.MoveTo Row, col
    
End Sub

' - PUTS STRING TO A SPECIFIC LOCATION IN BLUEZONE
Public Function putString(ByVal Value As String, ByVal Row As Integer, ByVal col As Integer) As Boolean

    pScreen.putString Value, Row, col
    
End Function

' - RETURNS THE FULL PAGE STRING
Public Function getScreenText() As String

    getScreenText = Trim(pScreen.Area(1, 1, pScreen.Rows, pScreen.Cols))
    
End Function

' - RETURNS A PAGE STRING
Public Function getText(Row As Integer, col As Integer, length As Integer) As String

    getText = pScreen.GetString(Row, col, length)
    
End Function

' - RETURNS THE FULL PAGE STRING
Public Sub selectText(Row As Integer, col As Integer, length As Integer)

    pScreen.Select Row, col, Row, col + length
    
End Sub

' - CHECK TO SEE IF CURRENT PAGE CONTAINS A STRING VALUE.
Public Function pageContains(Value) As Boolean
    
    If getScreenText Like "*" & Value & "*" Then
        pageContains = True
    End If
    
End Function

' - WAITS FOR STRING TO APPEAR ON SCREEN, WILL TIME OUT ONCE IT HITS MAXSECONDS
Public Function waitForStringOnPage(Value As String, Optional maxSeconds As Integer = 3) As Boolean
    
    Do
        If pageContains(Value) Then
            waitForStringOnPage = True
            Exit Function
        End If
        
        If maxTimeLimitReached(maxSeconds) Then
            waitForStringOnPage = False
            Exit Function
        End If
        
    Loop

End Function

' - WAITS FOR PAGE TO NO LONGER BE BLANK. WAS NEEDED FOR CLEARING OUT TO ENTER PACS.
Public Function waitForPageToLoad(Optional maxSeconds As Integer = 3) As Boolean
    
    Do
        If getScreenText <> "" Then
            Exit Do
        End If
        
        If maxTimeLimitReached(maxSeconds) Then
            Exit Function
        End If
    Loop
    
    waitForPageToLoad = True
    
End Function

Public Function connectToActiveSession(Optional errorMessage As Boolean = False) As Boolean
    
    Dim ReturnValue As New Scripting.Dictionary
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ATTEMPT TO FIND SESSION OF EXTRA
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Set pSystem = CreateObject("EXTRA.System")
'    setSessionToActive
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SEE IF A SESSION TO EXTRA WAS FOUND, IF NOT TRY TO GET BLUEZONE SESSION
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If pSession Is Nothing Then
    
        'SET SESSION TO BLUEZONE
        Set pSystem = CreateObject("bluezone.System")
        setSessionToActive
        
        'CHECK TO SEE IF SESSION WAS CREATED TO BLUEZONE
        If pSession Is Nothing Then
            connectToActiveSession = False
            If errorMessage = True Then
                MsgBox "Error connecting to an instance of BlueZone", vbCritical
            End If
            Exit Function
        End If

    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GET SCREEN OBJECT, EXPAND, AND CHECK TO MAKE SURE USER IS LOGED IN
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'CREATE SCREEN OBJECT AND OPEN
    Set pScreen = pSession.screen
    
    
    'DON'T MAXIMIZE FOR HEATHER, HURTS HER EYES
    If Environ("Username") <> "hball06" And Environ("Username") <> "kdern01" Then
        pSession.WindowState = 2
    End If
    
    'CLEAR ANY HANG UPS
    pScreen.Sendkeys ("<RESET>")
    
    'CHECK TO SEE THAT THE USER IS LOGED IN
    checkForLogin
    
    connectToActiveSession = True
    
End Function

'******************************************************************************************
' PUBLIC METHODS - SPECIFIC TOOLS
'******************************************************************************************

' - USED FOR GOING TO THE 17/2 SCREEN IN PACS
Public Function Goto_17Set2(Div As String, vendorNumber As String, logNumber As String) As Boolean
    
    'CLEAR ANY HANG UPS
    pScreen.Sendkeys ("<RESET>")
    
    'CLEAR OUT OF ANY CURRENT SCREEN
    If clearOutToABlankScreen = False Then
        
        'RETURN ERROR MESSAGE
        Goto_17Set2 = False
        Exit Function
        
    End If
   
    'FROM BLANK SCREEN, NAVIGATE TO PACS
    pScreen.Sendkeys ("PACS<ENTER>")
    
    'ONCE IN PACS, NAVIGATE TO PROPER DIVISION
    If waitForStringOnPage("*PACS -  Logo*Division:*") = False Then
        Goto_17Set2 = False
        Exit Function
    End If
    Sleep 50
    pScreen.putString Div, 3, 18
    pScreen.Sendkeys "<Enter>"
    
    'ONCE IN DIVISION, SWITCH PF SET 1 TO PF SET 2, THEN ENTER THE 17/2
    waitForStringOnPage "*PF*SET:*1*"
    pScreen.Sendkeys ("<PF1>")
    waitForStringOnPage "*PF Set: 2*"
    pScreen.Sendkeys ("<PF17>")
    
    'ONCE IN THE 17/2 ENTER THE PROPER VENDOR NUMBER AND LOG NUMBER
    waitForStringOnPage "*PF17/2*"
    pScreen.putString vendorNumber, 3, 73
    pScreen.putString logNumber, 3, 90
    pScreen.Sendkeys ("<Enter>")
    
    Goto_17Set2 = True
    
End Function

' - USED FOR BILLING ON THE 15/2 SCREEN FOR CASE AUDIT.
Public Function bill_15Set2_pos(CIC As String, poList As Scripting.Dictionary, billLetter As String) As String
    
    Dim Index As Integer
    Dim lastPoPage As Boolean
    Dim poInputedToPage As Boolean
    
    'CLEAR ANY HANG UPS
    pScreen.Sendkeys ("<RESET>")
    
    'IF ON 17/2 THEN SWITCH TO 15/2
    If pageContains("*PF17/2*") Then
        pScreen.Sendkeys ("<PF15>")
        Sleep 100
    End If
    
    'DOUBLE CHECK THAT CURRENT PAGE IS 15/2
    If waitForStringOnPage("*PF15/2*", 5) = False Then
        bill_15Set2_pos = "[ERROR]: Was not able to navigate to the 15 set 2 screen."
        Exit Function
    End If
    
    'NAVIGATE TO CURRENT CIC
    If Trim(pScreen.Area(3, 24, 3, 38)) <> CIC Then
        pScreen.putString "               ", 3, 24
        pScreen.putString CIC, 3, 24
        pScreen.Sendkeys ("<ENTER>")
        Sleep 60
    End If
    
    'DOUBLE CHECK THAT CURRENT CIC IS CORRECT
    If waitForStringOnPage("*" & CIC & "*", 5) = False Then
        bill_15Set2_pos = "Cic not found"
        Exit Function
    End If
    
    'MAKE SURE TO START PO LIST AT THE TOP
    pScreen.putString "T", 8, 16
    pScreen.Sendkeys ("<ENTER>")
    Sleep 60
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' OUTER LOOP GOES THROUGH ALL THE POS PASSED IN THE PARAMETER poList
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Do
        
        'CHECK TO SEE IF AT THE BOTTOM OF THE LIST OF POS
        If pageContains("*NO MORE RECEIVINGS FOR THIS ITEM*") Then
            lastPoPage = True
        End If
        
        'SECOND CHECK TO SEE IF AT THE BOTTOM OF THE LIST.
        If pageContains("*SCROLL F=FWD INVALID SINCE NO FOLLOWING PAGES*") Then
            lastPoPage = True
        End If
        
        'IS USED TO CHECK TO SEE IF ADDITIONAL ACTIONS ARE NEEDED. RESET FOR NEW PO LIST.
        poInputedToPage = False
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' INNER LOOP GOES THROUGH CURRENT LIST OF POS ON THE SCREEN
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For Index = 17 To 26
        
            'CHECK TO SEE IF PO IS TO BE BILLED
            If poList.exists(pScreen.GetString(Index, 15, 6)) Then
                
                'ADD BILLING LETTER
                pScreen.putString billLetter, Index, 7
                
                'ADD AMOUNT REC IF IT IS NOT 0 AND BILLLETTER IS NOT AN UNDO BILLING(" ").
                If poList(pScreen.GetString(Index, 15, 6)) <> 0 And billLetter <> " " Then
                    pScreen.putString poList(pScreen.GetString(Index, 15, 6)), Index, 111
                Else
                
                    'CLEAR AMOUNT REC
                    pScreen.putString "          ", Index, 111
                End If
                
                'MARK PAGE AS HAVING AN INPUT, TO MAKE SURE PAGE IS UPDATED AFTER LOOP
                poInputedToPage = True
                
            End If
        Next Index
        
        
        'CHECK IF ANY POS NEED TO BE SUBMITED
        If poInputedToPage = True Then
            pScreen.Sendkeys ("<ENTER>")
            Sleep 60
        End If
        
        'CHECK FOR ERROR IN SUBMITING THE PO BILLINGS
        If pageContains("BILLING NOT ALLOWED. EXTENDED BILLING AMOUNT IS ZERO") Then
            bill_15Set2_pos = "[ERROR]: BILLING NOT ALLOWED. EXTENDED BILLING AMOUNT IS ZERO. Clear any errors in Pacs, try again making sure you use the correct billing letter."
            Exit Function
        End If
        
        'GO TO THE NEXT SET OF POS FOR CURRENT CIC IF NOT ON THE LAST PAGE OF POS
        If lastPoPage = False Then
            pScreen.putString "F", 8, 16
            pScreen.Sendkeys ("<ENTER>")
            Sleep 60
        End If
        
    Loop While lastPoPage = False
    
    'MARK AS COMPLETE
    bill_15Set2_pos = ""
    
    
End Function


'******************************************************************************************
' PRIVATE HELPER FUNCTIONS
'******************************************************************************************

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


'USED FOR GETTING BLANK SCREEN TO HAVE A CLEAR LINE OF NAVIGATION
Public Function clearOutToABlankScreen() As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'SENDKEYS FOR CLEARING OUT SCREEN
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    pScreen.Sendkeys ("<PF3><PF3><Clear><Clear><CLEAR>")
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'RUN LOOP CHECKING FOR A BLANK SCREEN
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Do
        'CHECK FOR BLANK SCREEN
        If getScreenText = "" Then
            Exit Do
        End If
        
        'MAX TIME OF RUNNING LOOP IS THREE SECONDS
        If maxTimeLimitReached(3) Then
            clearOutToABlankScreen = False
            Exit Function
        End If
    Loop
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SUCSESS! SCREEN IS CLEARED AND BLANK.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    clearOutToABlankScreen = True

End Function

'CHECKS TO MAKE SURE USER IS LOGED INTO EXTRA
Private Function checkForLogin() As Boolean
    
    'CHECK TO SEE IF CURRENTLY IN THE LOGIN SCREEN
    If pScreen.GetString(16, 19, 13) = "New Password:" Then
        
        'PROMPT FOR USER TO LOGIN
        If MsgBox("Please login to Extra and then click OK to continue", vbOKCancel) = vbCancel Then
            
            'CANCEL PRESSED
            checkForLogin = False
            Exit Function
            
        Else
            'OK PRESSED, CHECK TO SEE THAT THEY DID LOGIN
            checkForLogin
            
        End If
    
    End If
    
End Function

Private Sub setSessionToActive()
    Set pSessions = pSystem.Sessions
    Set pSession = pSystem.ActiveSession
End Sub

Private Sub createNewSession()

    pSessions.Open "Safeway"
    setSessionToActive
    
End Sub


