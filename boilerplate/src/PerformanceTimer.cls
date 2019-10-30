VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * Stopwatch that is accurate to the millisecond. This is a good tool for tracking
' * code performance or used with animations.
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Option Explicit
Option Compare Text

'/**
' * API's that are needed to get accurate time measurements.
' */
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

'/**
' * State is a {Type} containing all the private variables for this class.
' */
Private Type state
    Frequency As Currency
    StartTime As Double
    SplitTime As Double
    Laps As Collection
End Type

'/**
' * property {state} this - The private properties stored in one place.
' */
Private this As state

'/**
' * Laps are stored in a collection. Must create new collection.
' */
Private Sub Class_Initialize()
    Set this.Laps = New Collection
End Sub

'/**
' * @property {Double} Milliseconds - Returns current milliseconds.
' */
Public Property Get Milliseconds() As Double
    Milliseconds = Seconds * 1000
End Property

'/**
' * @property {Double} Seconds - Returns current Seconds.
' */
Public Property Get Seconds() As Double
    If this.Frequency = 0 Then
        Seconds = 0
    Else
        Seconds = ((Tick / this.Frequency) - this.StartTime)
    End If
End Property

'/**
' * This is where the user starts the timer. Must be called, otherwise it stays at 0.
' */
Public Sub StartTimer()
    getFrequency this.Frequency
    this.StartTime = Tick / this.Frequency
End Sub

'Public Sub PauseTimer()
'End Sub

'/**
' * Set's everything back to 0. Timer must be started again.
' */
Public Sub ResetTimer()
    this.SplitTime = 0
    this.StartTime = 0
    this.Frequency = 0
End Sub

'/**
' * Records the current lap.
' */
Public Sub Lap()
    Dim split As New Collection
    split.Add Round(Me.Seconds, 2), "Time"
    
    If this.Laps.count = 0 Then
        split.Add split("Time"), "Lenght"
    Else
        split.Add split("Time") - this.SplitTime, "Lenght"
    End If
    
    this.SplitTime = split("Time")
    this.Laps.Add split
    
    Debug.Print "Split: " & this.Laps.count & " Time: " & split("Time") & " Length: " & split("Lenght")
End Sub

Public Function GetLapTime(ByVal index As Long) As Double
    GetLapTime = this.Laps(index)("Time")
End Function

Public Function GetLapLenght(ByVal index As Long) As Double
    GetLapLenght = this.Laps(index)("Lenght")
End Function

Public Function LapCount() As Long
    LapCount = this.Laps.count
End Function

'/**
' * Needed to get accuarate time.
' */
Private Property Get Tick() As Currency
    Dim Ticks1 As Currency
    getTickCount Ticks1
    
    Dim Ticks2 As Currency
    getTickCount Ticks2
    
    If Ticks2 < Ticks1 Then
        Tick = Ticks1
    Else
        Tick = Ticks2
    End If
End Property



