VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListenerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'@dependancy Reference to Microsoft Scripting Runtime (For Dictionarys)

'SET FROM AddEventListener - NEEDED TO EMIT EVENT BACK TO IT.
Private WithEvents pEmitter As EventListenerEmitter
Attribute pEmitter.VB_VarHelpID = -1

'CONTROLS THAT HAVE THE EVENTS
Private WithEvents form As MSForms.UserForm
Attribute form.VB_VarHelpID = -1
Private WithEvents Txt As MSForms.Textbox
Attribute Txt.VB_VarHelpID = -1
Private WithEvents lbl As MSForms.Label
Attribute lbl.VB_VarHelpID = -1
Private WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Private WithEvents Cmb As MSForms.ComboBox
Attribute Cmb.VB_VarHelpID = -1
Private WithEvents Frm As MSForms.Frame
Attribute Frm.VB_VarHelpID = -1

'PROPERTIES OF CONTROL
Private pControl As Object
Private IsHoveredControl As Boolean
Private IsFocusedControl As Boolean

'TODO
' - ADD EVENTS FOR ALL USERFORM CONTROLS
' - ADD MORE EVENTS THAN THE LIST BELOW
' - DOUBLE CHECK THAT EMITTS ARE ALL CURRENT

'CURRENT EVENTS (WILL ADD MORE AS I HAVE NEED)
' - MouseOver
' - MouseOut
' - MouseMove
' - Click
' - DblClick
' - Focus
' - Blur
' - KeyUp
' - KeyDown

'***********************************************************************************
' PUBLIC FUNCTIONS
'***********************************************************************************

'ONLY PUBLIC METHOD. CALLED FROM EVENTListener CLASS MODULE
Public Sub AddEventListener(ByRef ControlOrForm As Object, ByRef emitter As EventListenerEmitter)
    
    'CAPTURE THE EMITTER CLASS. WILL USE THIS TO EMIT EVENTS FROM EACH CONTROL
    Set pEmitter = emitter
    
    'USED TO COMPARE CHECK IF IT IS THE CONTROL TRIGGERING THE EVENT
    Set pControl = ControlOrForm
    
    'SET CONTROL(OR FORM) BASED ON IT'S TYPE
    Select Case TypeName(ControlOrForm)
        Case "CommandButton"
            Set Btn = ControlOrForm
        
        Case "ComboBox"
            Set Cmb = ControlOrForm
            
        Case "Frame"
            Set Frm = ControlOrForm
        
        Case "Label"
            Set lbl = ControlOrForm
        
        Case "TextBox"
            Set Txt = ControlOrForm
            
        Case Else
 
            If IsUserform(ControlOrForm) Then
                Set form = ControlOrForm
            End If

    End Select
    
End Sub

'***********************************************************************************
' PRIVATE HELPER FUNCTIONS
'***********************************************************************************

'CALLED ON MOUSEMOVE EVENT, THIS IS A WAY OF CREATING A MOUSEOVER AND MOUSEOUT EVENT
Private Sub CheckIfHoveredControl()
    
    If Not IsHoveredControl Then
        IsHoveredControl = True
        pEmitter.EmitEvent pControl, "MouseOver", ToDictionary()
    End If
    
End Sub

'CALLED ON MOUSEMOVE EVENT, THIS IS A WAY OF CREATING A MOUSEOVER AND MOUSEOUT EVENT
Private Sub CheckIfFocusedControl()
    
    If Not IsFocusedControl Then
    
        If TypeName(pControl) = "Frame" Then
             pEmitter.SetFocusToControl pControl.ActiveControl
        Else
            IsFocusedControl = True
            pEmitter.EmitEvent pControl, "Focus", ToDictionary()
        End If
        
    End If
    
End Sub

'CHECK TO SEE IF OBJ IS A USERFORM
Private Function IsUserform(ByRef Obj As Object) As Boolean

    If TypeOf Obj Is MSForms.UserForm Then
      IsUserform = True
    End If

End Function

'A simple Dictionary Factory.
Private Function ToDictionary(ParamArray keyValuePairs() As Variant) As Scripting.Dictionary
    
    '@author: Robert Todar <robert@roberttodar.com>
    '
    
    'Get length of array to check to see if there are valid parameters.
    Dim ArrayLenght As Long
    ArrayLenght = UBound(keyValuePairs) - LBound(keyValuePairs) + 1
    
    'Check to see that key/value pairs passed in (an even number).
    If ArrayLenght Mod 2 <> 0 Then
        Err.Raise 5, TypeName(Me), "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
    End If
    
    'Add key values to the return Dictionary.
    Set ToDictionary = New Scripting.Dictionary
    Dim index As Long
    For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        ToDictionary.Add keyValuePairs(index), keyValuePairs(index + 1)
    Next index
    
End Function


'***********************************************************************************
' EVENTS
'***********************************************************************************

'----------------------------------------------------------------------------
' ONCE AN EVENT HAS EMMITED, EACH EVENTListenerITEM WILL LISTEN FOR THAT EVENT
'----------------------------------------------------------------------------
Private Sub pEmitter_EmittedEvent(ByRef Control As Object, ByVal EventName As String, ByRef EventParameters As Scripting.Dictionary)
    
    'CREATE A MOUSEOVER MOUSEOUT EVENTS
    Select Case EventName
    
        Case "MouseOver"
            If pControl.Name <> Control.Name And IsHoveredControl Then
                IsHoveredControl = False
                pEmitter.EmitEvent pControl, "MouseOut", ToDictionary()
            End If
            
        Case "Focus"
            If pControl.Name <> Control.Name And IsFocusedControl Then
                IsFocusedControl = False
                pEmitter.EmitEvent pControl, "Blur", ToDictionary()
            ElseIf pControl.Name = Control.Name And IsFocusedControl = False Then
                IsFocusedControl = True
            End If
    
    End Select

End Sub


'------------------------------------------------------------------------
' USERFORM
'------------------------------------------------------------------------
Private Sub Form_Click()
    pEmitter.EmitEvent pControl, "Click", ToDictionary()
End Sub

Private Sub Form_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", ToDictionary("Cancel", Cancel)
End Sub

Private Sub Form_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   CheckIfHoveredControl
   pEmitter.EmitEvent pControl, "MouseMove", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Form_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseDown", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Form_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseUp", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub


'------------------------------------------------------------------------
' COMMAND BUTTON
'------------------------------------------------------------------------
Private Sub Btn_Click()
    pEmitter.EmitEvent pControl, "Click", ToDictionary()
End Sub

Private Sub Btn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", ToDictionary("Cancel", Cancel)
End Sub

Private Sub Btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Btn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, "MouseUp", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Btn_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseDown", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Btn_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, "KeyUp", ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub Btn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    pEmitter.EmitEvent pControl, "KeyDown", ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub


'------------------------------------------------------------------------
' LABEL
'------------------------------------------------------------------------
Private Sub Lbl_Click()
    pEmitter.EmitEvent pControl, "Click", ToDictionary()
End Sub

Private Sub Lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", ToDictionary("Cancel", Cancel)
End Sub

Private Sub Lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseDown", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub lbl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseUp", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub


'------------------------------------------------------------------------
' Frame
'------------------------------------------------------------------------
Private Sub Frm_Click()
    pEmitter.EmitEvent pControl, "Click", ToDictionary()
End Sub

Private Sub Frm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", ToDictionary("Cancel", Cancel)
End Sub

Private Sub Frm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Frm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl 'FRAME DOESN'T TAKE FOCUS BUT ACTIVE CONTROL IN FRAME DOES
    pEmitter.EmitEvent pControl, "MouseDown", ToDictionary()
End Sub

Private Sub Frm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseUp", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub


'------------------------------------------------------------------------
' Textbox
'------------------------------------------------------------------------
Private Sub Txt_Click()
    pEmitter.EmitEvent pControl, "Click", ToDictionary()
End Sub

Private Sub Txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", ToDictionary("Cancel", Cancel)
End Sub

Private Sub Txt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Txt_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, "MouseUp", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Txt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, "MouseDown", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Txt_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, "KeyUp", ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    pEmitter.EmitEvent pControl, "KeyDown", ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub


'------------------------------------------------------------------------
' Combobox
'------------------------------------------------------------------------
Private Sub Cmb_Click()
    pEmitter.EmitEvent pControl, "Click", ToDictionary()
End Sub

Private Sub Cmb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, "DblClick", ToDictionary("Cancel", Cancel)
End Sub

Private Sub Cmb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, "MouseMove", ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub




