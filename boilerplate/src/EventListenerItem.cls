VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListenerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * Wrapper around various MSForms controls (and Userform) to listen to events.
' * This adds greater control in creating more events, namely: Blur, Focus,
' * MouseOver, MouseOut, ect.
' * This works hand in hand with `EventListenerEmitter`. As events happen here
' * This raises events for all other controls to see what is going on.
' *
' * @ref {Libary} Microsoft Scripting Runtime.
' *
' * TODO
' *  - ADD EVENTS FOR ALL USERFORM CONTROLS
' *  - ADD MORE EVENTS THAN THE LIST BELOW
' *  - DOUBLE CHECK THAT EMITTS ARE ALL CURRENT
' *
' * CURRENT EVENTS (WILL ADD MORE AS I HAVE NEED)
' *  - MouseOver
' *  - MouseOut
' *  - MouseMove
' *  - Click
' *  - DblClick
' *  - Focus
' *  - Blur
' *  - KeyUp
' *  - KeyDown
' */
Option Explicit
Option Compare Text

' Set from addeventlistener - needed to emit event back to it.
Private WithEvents emitter As EventListenerEmitter
Attribute emitter.VB_VarHelpID = -1

' Controls that have the events.
Public WithEvents UserForm As MSForms.UserForm
Attribute UserForm.VB_VarHelpID = -1
Public WithEvents textbox As MSForms.textbox
Attribute textbox.VB_VarHelpID = -1
Public WithEvents label As MSForms.label
Attribute label.VB_VarHelpID = -1
Public WithEvents CommandButton As MSForms.CommandButton
Attribute CommandButton.VB_VarHelpID = -1
Public WithEvents ComboBox As MSForms.ComboBox
Attribute ComboBox.VB_VarHelpID = -1
Public WithEvents Frame As MSForms.Frame
Attribute Frame.VB_VarHelpID = -1

' Private state within this class.
Private Type state
    controlReference As Object
    isHoveredControl As Boolean
    isFocusedControl As Boolean
End Type

Private this As state

'/**
' * The only public method. called from eventlistener class module.
' * That class will have a copy of this class for every control.
' */
Public Sub AddEventListener(ByRef ControlOrForm As Object, ByRef EmitterSource As EventListenerEmitter)
    ' Capture the emitter class. will use this to emit events from each control
    Set emitter = EmitterSource
    
    ' Used to store a reference to what this class control is referenced to.
    ' Essentially, this class only holds reference to one control type.
    ' All other controls have no references to them.
    Set this.controlReference = ControlOrForm
    
    ' Set control (or form) based on it's type
    Select Case TypeName(ControlOrForm)
        Case "CommandButton"
            Set CommandButton = ControlOrForm
        
        Case "ComboBox"
            Set ComboBox = ControlOrForm
            
        Case "Frame"
            Set Frame = ControlOrForm
        
        Case "Label"
            Set label = ControlOrForm
        
        Case "TextBox"
            Set textbox = ControlOrForm
            
        Case Else
            ' Check to see if it is a userform.
            If TypeOf ControlOrForm Is MSForms.UserForm Then
                Set UserForm = ControlOrForm
            End If
    End Select
End Sub

' Called on mousemove event, this is a way of creating a mouseover and mouseout event
Private Sub CheckIfHoveredControl()
    If Not this.isHoveredControl Then
        this.isHoveredControl = True
        ' Emmit the event letting other controls do their own logic.
        emitter.EmitEvent this.controlReference, MouseOver, ToDictionary()
    End If
End Sub

' Called on mousemove event, this is a way of creating a mouseover and mouseout event
Private Sub CheckIfFocusedControl()
    If Not this.isFocusedControl Then
        If TypeName(this.controlReference) = "Frame" Then
            ' Emmit the event letting other controls do their own logic.
             emitter.SetFocusToControl this.controlReference.ActiveControl
        Else
            ' Emmit the event letting other controls do their own logic.
            this.isFocusedControl = True
            emitter.EmitEvent this.controlReference, Focus, ToDictionary()
        End If
    End If
End Sub

' A simple Dictionary Factory.
Private Function ToDictionary(ParamArray keyValuePairs() As Variant) As Scripting.Dictionary
    ' Get length of array to check to see if there are valid parameters.
    Dim ArrayLenght As Long
    ArrayLenght = UBound(keyValuePairs) - LBound(keyValuePairs) + 1
    
    ' Check to see that key/value pairs passed in (an even number).
    If ArrayLenght Mod 2 <> 0 Then
        Err.Raise 5, TypeName(Me), "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
    End If
    
    ' Add key values to the return Dictionary.
    Set ToDictionary = New Scripting.Dictionary
    Dim index As Long
    For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        ToDictionary.Add keyValuePairs(index), keyValuePairs(index + 1)
    Next index
End Function

'/**
' * This listens to events from other controls. This is the entry point for that
' * Open communication between the various controls.
' * This is what allows for our custom events.
' */
Private Sub emitter_EmittedEvent(ByRef Control As Object, ByVal EventType As EmittedEvent, ByRef EventParameters As Scripting.Dictionary)
    ' Create our custom events.
    ' Note, this also helps with performance by not emitting events if they are already
    ' active. Why do a mouseover event if the mouse is already over? Yay!
    Select Case EventType
        Case MouseOver
            If this.controlReference.Name <> Control.Name And this.isHoveredControl Then
                this.isHoveredControl = False
                emitter.EmitEvent this.controlReference, MouseOut, ToDictionary()
            End If
            
        Case Focus
            If this.controlReference.Name <> Control.Name And this.isFocusedControl Then
                this.isFocusedControl = False
                emitter.EmitEvent this.controlReference, Blur, ToDictionary()
            ElseIf this.controlReference.Name = Control.Name And this.isFocusedControl = False Then
                this.isFocusedControl = True
            End If
    End Select
End Sub

'/**
' * Userform Events
' */
Private Sub UserForm_Click()
    emitter.EmitEvent this.controlReference, Click, ToDictionary()
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    emitter.EmitEvent this.controlReference, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   CheckIfHoveredControl
   emitter.EmitEvent this.controlReference, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    emitter.EmitEvent this.controlReference, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    emitter.EmitEvent this.controlReference, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub


'/**
' * CommandButton Events
' */
Private Sub CommandButton_Click()
    emitter.EmitEvent this.controlReference, Click, ToDictionary()
End Sub

Private Sub CommandButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    emitter.EmitEvent this.controlReference, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub CommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    emitter.EmitEvent this.controlReference, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub CommandButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl
    emitter.EmitEvent this.controlReference, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub CommandButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    emitter.EmitEvent this.controlReference, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub CommandButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    emitter.EmitEvent this.controlReference, KeyUp, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub CommandButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    emitter.EmitEvent this.controlReference, KeyDown, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub


'/**
' * Label Events
' */
Private Sub Label_Click()
    emitter.EmitEvent this.controlReference, Click, ToDictionary()
End Sub

Private Sub Label_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    emitter.EmitEvent this.controlReference, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    emitter.EmitEvent this.controlReference, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Label_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    emitter.EmitEvent this.controlReference, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Label_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    emitter.EmitEvent this.controlReference, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub


'/**
' * Frame Events
' */
Private Sub Frame_Click()
    emitter.EmitEvent this.controlReference, Click, ToDictionary()
End Sub

Private Sub Frame_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    emitter.EmitEvent this.controlReference, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    emitter.EmitEvent this.controlReference, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Frame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Frame doesn't take focus but active control in frame does
    CheckIfFocusedControl
    emitter.EmitEvent this.controlReference, MouseDown, ToDictionary()
End Sub

Private Sub Frame_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    emitter.EmitEvent this.controlReference, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub


'/**
' * Textbox Events
' */
Private Sub Textbox_Click()
    emitter.EmitEvent this.controlReference, Click, ToDictionary()
End Sub

Private Sub Textbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    emitter.EmitEvent this.controlReference, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub Textbox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    emitter.EmitEvent this.controlReference, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Textbox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl
    emitter.EmitEvent this.controlReference, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Textbox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    emitter.EmitEvent this.controlReference, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Textbox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    emitter.EmitEvent this.controlReference, KeyUp, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub Textbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    emitter.EmitEvent this.controlReference, KeyDown, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub


'/**
' * Combobox Events
' */
Private Sub ComboBox_Click()
    emitter.EmitEvent this.controlReference, Click, ToDictionary()
End Sub

Private Sub ComboBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    emitter.EmitEvent this.controlReference, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub ComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    emitter.EmitEvent this.controlReference, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub
