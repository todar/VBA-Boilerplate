VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type state
    received As Variant
    Passed As Boolean
    message As String
    Runtime As Double
End Type

Private this As state

Private Sub Class_Initialize()
    
End Sub

' /** Gets status if test Passed for this single test */
Public Property Let Passed(ByVal value As Boolean)
    this.Passed = value
End Property

Public Property Get Passed() As Boolean
    Passed = this.Passed
End Property

' /** Gets status if test Message for this single test */
Public Property Let message(ByVal value As String)
    this.message = value
End Property

Public Property Get message() As String
    message = this.message
End Property

' /** The runtime for this single test */
Public Property Let Runtime(ByVal value As Double)
    this.Runtime = value
End Property

Public Property Get Runtime() As Double
    Runtime = this.Runtime
End Property

' /** The received value. This is what will be tested on. */
Public Property Let received(ByVal value As Variant)
    If IsObject(value) Then
        Set this.received = value
    Else
        this.received = value
    End If
End Property

Public Property Get received() As Variant
    If IsObject(this.received) Then
        Set received = this.received
    Else
        received = this.received
    End If
End Property


' /** Checks for equality */
Public Sub ToEqual(B As Variant, Optional message As String = _
    "Expected ${2}, Received ${1}")

    Check IsDeepEqual(this.received, B), message, this.received, B
End Sub

' /** Checks for inequality */
Public Sub ToNotEqual(B As Variant, Optional message As String = _
    "Expected ${1} to not equal ${2}")

    Check Not IsDeepEqual(this.received, B), message, this.received, B
End Sub

' /** Checks for greater than only */
Public Sub ToBeGreaterThan(B As Variant, Optional message As String = _
    "Expected ${1} to be greater than ${2}")

    Check this.received > B, message, this.received, B
End Sub

' /** Checks for Greater Than Or Equal */
Public Sub ToBeGreaterThanOrEqual(B As Variant, Optional message As String = _
    "Expected ${1} to be greater than or equal to ${2}")

    Check this.received >= B, message, this.received, B
End Sub

Public Sub ToBeLessThan(B As Variant, Optional message As String = _
    "Expected ${1} to be less than ${2}")

    Check this.received < B, message, this.received, B
End Sub

Public Sub ToBeLessThanOrEqual(B As Variant, Optional message As String = _
    "Expected ${1} to be less than or equal to ${2}")

    Check this.received <= B, message, this.received, B
End Sub

''
' Check if a value is "truthy"
'
' From https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/if-then-else-statement
'
' Must evaluate to True or False, or to a data type that is implicitly convertible to Boolean.
' If the expression is a Nullable Boolean variable that evaluates to Nothing, the condition is treated as if the expression is False.
''
Public Sub ToBeTruthy(Optional message As String = _
    "Expected ${1} to be truthy")

    Check this.received, message, this.received
End Sub

''
' Check if a value is not "truthy" (See .IsOk)
''
Public Sub ToBeFalsy(Optional message As String = _
    "Expected ${1} to not be ok")

    Check Not CBool(this.received), message, this.received
End Sub

''
' Check if a value is "undefined": Nothing, Empty, Null, or Missing
'
' @param {Variant} Value
' @param {String} [Message]
''
Public Sub ToBeUndefined(Optional message As String = _
    "Expected ${1} to be undefined")

    Check IsNothing(this.received) Or VBA.IsEmpty(this.received) Or VBA.IsNull(this.received) Or VBA.IsMissing(this.received), message, this.received
End Sub

''
' Check if a value is not "undefined": Nothing, Empty, Null, or Missing
'
' @param {Variant} Value
' @param {String} [Message]
''
Public Sub ToBeDefined(Optional message As String = _
    "Expected ${1} to be defined")

    Check Not IsNothing(this.received) And Not VBA.IsEmpty(this.received) And Not VBA.IsNull(this.received) And Not VBA.IsMissing(this.received), message, this.received
End Sub

''
' Check if a value is included in an arbitrarily nested Array or Collection
'
' @param {Array|Collection} Values
' @param {Variant} Value
' @param {String} [Message]
''
Public Sub ToInclude(value As Variant, Optional message As String = _
    "Expected ${2} to be included in ${1}")

    If IsCollection(this.received) Then
        Check CollectionIncludes(this.received, value), message, this.received, value
    ElseIf IsArray(this.received) Then
        Check ArrayIncludes(this.received, value), message, this.received, value
    Else
        Debug.Print FormatMessage(message, this.received, value) & " (Incompatible type for Values)"
    End If
End Sub

''
' Check if a value is not included in an arbitrarily nested Array or Collection
'
' @param {Array|Collection} Values
' @param {Variant} Value
' @param {String} [Message]
''
Public Sub ToNotInclude(value As Variant, Optional message As String = _
    "Expected ${2} not to be included in ${1}")

    If IsCollection(this.received) Then
        Check Not CollectionIncludes(this.received, value), message, this.received, value
    ElseIf IsArray(this.received) Then
        Check Not ArrayIncludes(this.received, value), message, this.received, value
    Else
        Debug.Print FormatMessage(message, this.received, value) & " (Incompatible type for received)"
    End If
End Sub

''
' Check if two values are approximately equal, up to the given amount of significant figures
'
' @example
' ```vb
' .IsApproximate 1.001, 1.002, 3
'
' ' Equivalent to .IsEqual 1.00e+0, 1.00e+0
' ```
' @param {Variant} B
' @param {String} [Message]
''
Public Sub ToBeApproximate(B As Variant, SignificantFigures As Integer, Optional message As String = _
    "Expected ${1} to be approximately equal to ${2} (with ${3} significant figures of precision)")

    If SignificantFigures < 1 Or SignificantFigures > 15 Then
        Check False, "IsApproximate can only compare from 1 to 15 significant figures"
    Else
        Check IsApproximatelyEqual(this.received, B, SignificantFigures), message, this.received, B, SignificantFigures
    End If
End Sub

''
' Check if two values are approximately equal, up to the given amount of significant figures
'
' @example
' ```vb
' .NotApproximate 1.001, 1.009, 3
'
' ' Equivalent to .IsEqual 1.00e+0, 1.01e+0
' ```
' @param {Variant} A
' @param {Variant} B
' @param {String} [Message]
''
Public Sub ToNotBeApproximate(B As Variant, SignificantFigures As Integer, Optional message As String = _
    "Expected ${1} to not be approximately equal to ${2} (with ${3} significant figures of precision)")

    If SignificantFigures < 1 Or SignificantFigures > 15 Then
        Check False, "To Not be Approximate can only compare from 1 to 15 significant figures"
    Else
        Check Not IsApproximatelyEqual(this.received, B, SignificantFigures), message, this.received, B, SignificantFigures
    End If
End Sub


' ============================================= '
' Private Functions
' ============================================= '
Private Sub Check(Assertion As Variant, message As String, ParamArray values() As Variant)
    If Assertion Then
        this.Passed = True
    Else
        this.Passed = False
        this.message = FormatMessage(message, values)
    End If
End Sub

Private Function IsDeepEqual(A As Variant, B As Variant) As Boolean
    Dim AType As VbVarType
    Dim BType As VbVarType

    AType = VBA.VarType(A)
    BType = VBA.VarType(B)

    If VBA.IsError(A) Or VBA.IsError(B) Then
        IsDeepEqual = False

    ElseIf VBA.IsArray(A) And VBA.IsArray(B) Then
        IsDeepEqual = IsArrayEqual(A, B)

    ElseIf AType = VBA.vbObject Or BType = VBA.vbObject Then
        If AType <> BType Or VBA.TypeName(A) <> VBA.TypeName(B) Then
            IsDeepEqual = False
        ElseIf VBA.TypeName(A) = "Collection" Then
            IsDeepEqual = IsCollectionEqual(A, B)
        ElseIf VBA.TypeName(A) = "Dictionary" Then
            IsDeepEqual = IsDictionaryEqual(A, B)
        Else
            IsDeepEqual = A Is B
        End If

    ElseIf VBA.VarType(A) = VBA.vbDouble Or VBA.VarType(B) = VBA.vbDouble Then
        ' It is inherently difficult/almost impossible to check equality of Double
        ' http://support.microsoft.com/kb/78113
        '
        ' -> Compare up to 15 significant figures
        IsDeepEqual = IsApproximatelyEqual(A, B, 15)

    Else
        IsDeepEqual = A = B
    End If
End Function

Private Function IsArrayEqual(A As Variant, B As Variant) As Boolean
    If UBound(A) <> UBound(B) Then
        IsArrayEqual = False
        Exit Function
    End If

    Dim i As Long
    For i = LBound(A) To UBound(A)
        If Not IsDeepEqual(A(i), B(i)) Then
            IsArrayEqual = False
            Exit Function
        End If
    Next i

    IsArrayEqual = True
End Function

Private Function IsCollectionEqual(A As Variant, B As Variant) As Boolean
    If A.count <> B.count Then
        IsCollectionEqual = False
        Exit Function
    End If

    Dim i As Long
    For i = 1 To A.count
        If Not IsDeepEqual(A(i), B(i)) Then
            IsCollectionEqual = False
            Exit Function
        End If
    Next i

    IsCollectionEqual = True
End Function

Private Function IsDictionaryEqual(A As Variant, B As Variant) As Boolean
    If UBound(A.keys) <> UBound(B.keys) Then
        IsDictionaryEqual = False
        Exit Function
    End If

    Dim AKeys As Variant
    Dim BKeys As Variant
    Dim i As Long

    AKeys = A.keys
    BKeys = B.keys

    For i = LBound(AKeys) To UBound(AKeys)
        If AKeys(i) <> BKeys(i) Or A.item(AKeys(i)) <> B.item(BKeys(i)) Then
            IsDictionaryEqual = False
            Exit Function
        End If
    Next i

    IsDictionaryEqual = True
End Function

Private Function IsCollection(value As Variant) As Boolean
    IsCollection = VBA.VarType(value) = VBA.vbObject And VBA.TypeName(value) = "Collection"
End Function

Private Function IsNothing(value As Variant) As Boolean
    If VBA.IsObject(value) Then
        IsNothing = value Is Nothing
    Else
        IsNothing = False
    End If
End Function

Private Function ArrayIncludes(values As Variant, value As Variant) As Boolean
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If VBA.IsArray(values(i)) Then
            If ArrayIncludes(values(i), value) Then
                ArrayIncludes = True
                Exit Function
            End If
        ElseIf IsCollection(values(i)) Then
            If CollectionIncludes(values(i), value) Then
                ArrayIncludes = True
                Exit Function
            End If
        ElseIf IsDeepEqual(values(i), value) Then
            ArrayIncludes = True
            Exit Function
        End If
    Next i

    ArrayIncludes = False
End Function

Private Function CollectionIncludes(values As Variant, value As Variant) As Boolean
    Dim item As Variant
    For Each item In values
        If VBA.IsArray(item) Then
            If ArrayIncludes(item, value) Then
                CollectionIncludes = True
                Exit Function
            End If
        ElseIf IsCollection(item) Then
            If CollectionIncludes(item, value) Then
                CollectionIncludes = True
                Exit Function
            End If
        ElseIf IsDeepEqual(item, value) Then
            CollectionIncludes = True
            Exit Function
        End If
    Next item

    CollectionIncludes = False
End Function

Private Function IsApproximatelyEqual(A As Variant, B As Variant, SignificantFigures As Integer) As Boolean
    If SignificantFigures < 1 Or SignificantFigures > 15 Or VBA.IsError(A) Or VBA.IsError(B) Then
        IsApproximatelyEqual = False
        Exit Function
    End If

    Dim AValue As String
    Dim BValue As String

    AValue = VBA.Format$(A, VBA.Left$("0.00000000000000", SignificantFigures + 1) & IIf(A > 1, "e+0", "e-0"))
    BValue = VBA.Format$(B, VBA.Left$("0.00000000000000", SignificantFigures + 1) & IIf(B > 1, "e+0", "e-0"))

    IsApproximatelyEqual = AValue = BValue
End Function

Private Function FormatMessage(message As String, ParamArray values() As Variant) As String
    Dim value As Variant
    Dim index As Long

    FormatMessage = message
    For Each value In IIf(VBA.IsArray(values(0)), values(0), values)
        index = index + 1
        FormatMessage = VBA.Replace(FormatMessage, "${" & index & "}", PrettyPrint(value))
    Next value
End Function

Private Function PrettyPrint(value As Variant, Optional Indentation As Long = 0) As String
    If VBA.IsMissing(value) Then
        PrettyPrint = "[Missing]"
        Exit Function
    End If

    Dim i As Long
    Dim Indented As String
    Indented = VBA.String$(Indentation + 1, "  ")

    Select Case VBA.VarType(value)
    Case VBA.vbObject
        ' Nothing
        If value Is Nothing Then
            PrettyPrint = "[Nothing]"

        ' Collection
        ElseIf VBA.TypeName(value) = "Collection" Then
            PrettyPrint = "[Collection [" & vbNewLine

            For i = 1 To value.count
                PrettyPrint = PrettyPrint & Indent(Indentation + 1) & _
                    PrettyPrint(value(i), Indentation + 1) & _
                    IIf(i <> value.count, ",", "") & vbNewLine
            Next i

            PrettyPrint = PrettyPrint & Indent(Indentation) & "]"

        ' Dictionary
        ElseIf VBA.TypeName(value) = "Dictionary" Then
            PrettyPrint = "[Dictionary {" & vbNewLine

            For i = LBound(value.keys) To UBound(value.keys)
                PrettyPrint = PrettyPrint & Indent(Indentation + 1) & _
                    value.keys(i) & ": " & _
                    PrettyPrint(value.item(value.keys(i)), Indentation + 1) & _
                    IIf(i <> value.count, ",", "") & vbNewLine
            Next i

            PrettyPrint = PrettyPrint & Indent(Indentation) & "}]"

        ' Object
        Else
            PrettyPrint = "[" & VBA.TypeName(value) & "]"
        End If

    ' Array
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        PrettyPrint = "[" & vbNewLine

        For i = LBound(value) To UBound(value)
            PrettyPrint = PrettyPrint & Indent(Indentation + 1) & _
                PrettyPrint(value(i), Indentation + 1) & _
                IIf(i <> UBound(value), ",", "") & vbNewLine
        Next i

        PrettyPrint = PrettyPrint & Indent(Indentation) & "]"

    ' Empty
    Case VBA.vbEmpty
        PrettyPrint = "[Empty]"

    ' Null
    Case VBA.vbNull
        PrettyPrint = "[Null]"

    ' String
    Case VBA.vbString
        PrettyPrint = """" & value & """"

    ' Everything else
    Case Else
        PrettyPrint = CStr(value)
    End Select
End Function

Private Function Indent(Optional Indentation As Long)
    Indent = VBA.String$(Indentation, "  ")
End Function




