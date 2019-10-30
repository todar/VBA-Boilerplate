VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type state
    description As String
    testMatchers As Collection
    perfTimer As PerformanceTimer
End Type

Private this As state

Private Sub Class_Initialize()
    Set this.testMatchers = New Collection
    Set this.perfTimer = New PerformanceTimer
    this.perfTimer.StartTimer
End Sub

Public Property Get description() As String
    description = this.description
End Property

Public Property Let description(ByVal value As String)
    this.description = value
End Property

Public Property Get FailureMessages() As Collection
    Set FailureMessages = New Collection
    Dim test As TestMatcher
    For Each test In this.testMatchers
        If test.Passed = False Then
            FailureMessages.Add test.message
        End If
    Next test
End Property

Public Property Get Passed() As Boolean
    Dim test As TestMatcher
    For Each test In this.testMatchers
        If test.Passed = False Then
            Exit Property
        End If
    Next test
    Passed = True
End Property

Public Property Get TotalCount() As Long
    TotalCount = this.testMatchers.count
End Property

Public Property Get PassedCount() As Long
    Dim test As TestMatcher
    For Each test In this.testMatchers
        If test.Passed = True Then
            PassedCount = PassedCount + 1
        End If
    Next test
End Property

Public Property Get FailedCount() As Long
    Dim test As TestMatcher
    For Each test In this.testMatchers
        If test.Passed = False Then
            FailedCount = FailedCount + 1
        End If
    Next test
End Property

Public Property Get Runtime() As Double
    Dim test As TestMatcher
    For Each test In this.testMatchers
        Runtime = Runtime + test.Runtime
    Next test
End Property

Public Property Get Expect(Optional ByVal received As Variant) As TestMatcher
    Set Expect = New TestMatcher
    Expect.Runtime = Math.Round(this.perfTimer.Milliseconds, 3)
    this.perfTimer.ResetTimer
    Expect.received = received
    this.testMatchers.Add Expect
End Property




