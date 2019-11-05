VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/**
' * This is a collection of related tests.
' *
' *
' */
Option Explicit

Private Type state
    description As String
    testCases As Collection
End Type

Private this As state

Private Sub Class_Initialize()
    Set this.testCases = New Collection
End Sub

Public Property Get Self() As TestSuite
    Set Self = Me
End Property

Public Property Get description() As String
    description = this.description
End Property

Public Property Let description(ByVal value As String)
    this.description = value
End Property

'/** Adds a new Test Case, returns that object. */
Public Property Get test(ByVal description As String) As TestCase
    Set test = New TestCase
    test.description = description
    this.testCases.Add test
End Property

Public Property Get Passed() As Boolean
    Dim spec As TestCase
    For Each spec In this.testCases
        If spec.Passed = False Then
            Exit Property
        End If
    Next spec
    Passed = True
End Property

Public Property Get Runtime() As Double
    Dim spec As TestCase
    For Each spec In this.testCases
        Runtime = Runtime + spec.Runtime
    Next spec
End Property

Public Property Get TotalCount() As Long
    Dim spec As TestCase
    For Each spec In this.testCases
        TotalCount = TotalCount + spec.TotalCount
    Next spec
End Property

Public Property Get PassedTestsCount() As Long
    Dim spec As TestCase
    For Each spec In this.testCases
        PassedTestsCount = PassedTestsCount + spec.PassedCount
    Next spec
End Property

Public Property Get PassedCaseCount() As Long
    Dim spec As TestCase
    For Each spec In this.testCases
        If spec.Passed Then
            PassedCaseCount = PassedCaseCount + 1
        End If
    Next spec
End Property

Public Property Get FailedCaseCount() As Long
    Dim spec As TestCase
    For Each spec In this.testCases
        If spec.Passed = False Then
            FailedCaseCount = FailedCaseCount + 1
        End If
    Next spec
End Property

Public Property Get FailedCount() As Long
    Dim spec As TestCase
    For Each spec In this.testCases
        FailedCount = FailedCount + spec.FailedCount
    Next spec
End Property


Private Sub Class_Terminate()
    Debug.Print ""
    Debug.Print "--------------------------------------"
    Debug.Print "# " & this.description
    Debug.Print "--------------------------------------"
    
    Dim spec As TestCase
    For Each spec In this.testCases
        Debug.Print IIf(spec.Passed, "+ ", "x ") & spec.description & " (" & spec.Runtime & "ms)"
        If spec.Passed = False Then
            Dim message As Variant
            For Each message In spec.FailureMessages
                Debug.Print "  * " & message
            Next message
        End If
    Next spec
    
    Debug.Print ""
    Debug.Print "Test Cases: " & _
                IIf(FailedCaseCount > 0, FailedCaseCount & " failed, ", "") & _
                IIf(PassedCaseCount > 0, PassedCaseCount & " passed, ", "") & _
                this.testCases.count & " total"
    
    Debug.Print "Tests:      " & _
                IIf(FailedCount > 0, FailedCount & " failed, ", "") & _
                IIf(PassedTestsCount > 0, PassedTestsCount & " passed, ", "") & _
                TotalCount & " total"
    Debug.Print "Time:       " & Me.Runtime & "ms"
    Debug.Print "Status:     " & IIf(Me.Passed, "PASS ", "FAIL ")
End Sub


