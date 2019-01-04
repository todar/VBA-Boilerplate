Attribute VB_Name = "MicroTimerFunctions"
Option Explicit
Option Compare Text
Option Private Module

'Sleep FUNCTIONLITY
#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'USED FOR MICRO TIMER
#If VBA7 Then
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias _
    "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias _
    "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" Alias _
    "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" Alias _
    "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If



Public Function MicroTimer(Optional StartTime As Boolean = False) As Double
   
    ' uses Windows API calls to the high resolution timer
    Static dTime As Double
    
    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    
    MicroTimer = 0

    'get frequency
    If cyFrequency = 0 Then getFrequency cyFrequency
    
    'get ticks
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    
    'calc seconds
    If cyFrequency Then MicroTimer = cyTicks2 / cyFrequency
    
    If StartTime = True Then
        dTime = MicroTimer
        MicroTimer = 0
    Else
        MicroTimer = (MicroTimer - dTime) * 1000  'CONVERT TO MILSECS
    End If
    
End Function


Private Sub TestMicroTimer()

    MicroTimer True
    Sleep 1000
    Debug.Print MicroTimer
    
End Sub
