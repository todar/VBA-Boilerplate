Attribute VB_Name = "TimingUtilities"
Option Explicit

' Sleep API to pause execution of code.
#If VBA7 And Win64 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
