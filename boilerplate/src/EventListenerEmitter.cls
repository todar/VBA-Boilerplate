VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListenerEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' ARRAY OF ALL THE DIFFERENT EVENT ListenerS FOR EVERY USERFORM CONTROL +FORM ITSELF
Private EventList() As New EventListenerItem

' ALL CURRENT POSSIBLE EVENTS THAT CAN BE EMITTED. NOTE, EMITTEDEVENT IS SENT FOR ALL!
Public Event EmittedEvent(ByRef Control As Object, ByVal EventName As String, ByRef EventParameters As Scripting.Dictionary)

Public Event Click(ByRef Control As Object)
Public Event DblClick(ByRef Control As Object, ByRef Cancel As MSForms.ReturnBoolean)
Public Event KeyUp(ByRef Control As Object, ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
Public Event KeyDown(ByRef Control As Object, ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
Public Event MouseOver(ByRef Control As Object)
Public Event MouseOut(ByRef Control As Object)
Public Event MouseMove(ByRef Control As Object, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
Public Event Focus(ByRef Control As Object)
Public Event Blur(ByRef Control As Object)
Public Event Change(ByRef Control As Object)

Public Event LabelMouseOver(ByRef lbl As MSForms.Label)
Public Event LabelMouseOut(ByRef lbl As MSForms.Label)
Public Event LabelClick(ByRef lbl As MSForms.Label)
Public Event LabelDblClick(ByRef lbl As MSForms.Label, ByRef Cancel As MSForms.ReturnBoolean)

' CALLED BY EVENTLISTENERCOLLECTION CLASS - MAIN ENTRYWAY OF EMITTING ALL EVENTS
Public Sub EmitEvent(ByRef Control As Object, ByVal EventName As String, ByRef EventParameters As Scripting.Dictionary)
    ' EVENT RAISED FOR ALL EVENTS. THIS IS A WAY FOR THE USER TO COLLECT FROM A SINGLE LOCATION.
    RaiseEvent EmittedEvent(Control, EventName, EventParameters)
    
    ' SPECIFIC EVENTS PER OBJECT-TYPE
    Select Case EventName
        Case "Click"
            RaiseEvent Click(Control)
        
        Case "DblClick"
            RaiseEvent DblClick(Control, EventParameters("Cancel"))
        
        Case "KeyUp"
            RaiseEvent KeyUp(Control, EventParameters("KeyCode"), EventParameters("Shift"))
            
        Case "KeyDown"
            RaiseEvent KeyDown(Control, EventParameters("KeyCode"), EventParameters("Shift"))
        
        Case "MouseOver"
            RaiseEvent MouseOver(Control)
            
        Case "MouseOut"
            RaiseEvent MouseOut(Control)
            
        Case "Focus"
            RaiseEvent Focus(Control)
            
        Case "Blur"
            RaiseEvent Blur(Control)
            
        Case "MouseMove"
            RaiseEvent MouseMove(Control, EventParameters("Shift"), EventParameters("X"), EventParameters("Y"))
    End Select
    
    Select Case TypeName(Control)
        Case "Label"
            EmitLabelEvent Control, EventName, EventParameters
    End Select
End Sub

Private Sub EmitLabelEvent(ByRef lbl As MSForms.Label, ByVal EventName As String, ByRef EventParameters As Scripting.Dictionary)
    Select Case EventName
        Case "Click"
            RaiseEvent LabelClick(lbl)
            
        Case "DblClick"
            RaiseEvent LabelDblClick(lbl, EventParameters("Cancel"))
            
        Case "MouseOver"
            RaiseEvent LabelMouseOver(lbl)
            
        Case "MouseOut"
            RaiseEvent LabelMouseOut(lbl)

        Case "MouseMove"
            'RaiseEvent MouseMove(Control, EventParameters("Shift"), EventParameters("X"), EventParameters("Y"))
    End Select
End Sub
    

' MUST CALL THIS IF YOU WANT TO programmatically SET CONTROL! OTHERWISE, EVENT'S WILL BE OFF!
Public Sub SetFocusToControl(ByRef Control As Object)
    ' If the user was to set focus through VBA then this code will fall apart considering
    ' it is unaware of that event occurring.
    If Not Control Is Nothing Then
        Control.setFocus
        EmitEvent Control, "Focus", Nothing
    End If
End Sub

'ADD EVENT ListenerS ON SPECIFIC CONTROLS - ALSO CALLED BY AddEventListenerAll
Public Sub AddEventListener(ByRef Control As Object)
    ' Events are stored in a private EventListenerItem array
    If IsArrayEmpty(EventList) Then
        ReDim EventList(0 To 0)
    Else
        ReDim Preserve EventList(0 To UBound(EventList) + 1)
    End If
    
    ' CALL AddEventListener IN EventListenerItem. THIS IS KEPT IN
    EventList(UBound(EventList)).AddEventListener Control, Me
End Sub

'ADD EVENT Listener TO ALL CONTROLS INCLUDING THE FORM
Public Sub AddEventListenerAll(ByRef form As Object)
    AddEventListener form
    
    Dim Ctrl As MSForms.Control
    For Each Ctrl In form.controls
        AddEventListener Ctrl
    Next Ctrl
End Sub

Private Function IsArrayEmpty(arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CPEARSON: http://www.cpearson.com/excel/VBAArrays.htm
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Err.Clear
    On Error Resume Next
    If IsArray(arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    Dim ub As Long
    ub = UBound(arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array. On these occasions, LBound is 0 and
        ' UBound is -1. To accommodate the weird behavior, test to see if LB > UB.
        ' If so, the array is not allocated.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Dim LB As Long
        LB = LBound(arr)
        If LB > ub Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If
End Function


