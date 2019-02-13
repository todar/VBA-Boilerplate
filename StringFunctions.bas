Attribute VB_Name = "stringFunctions"
Option Explicit
Option Private Module
Option Compare Text

'THIS RETURNS A PERCENTAGE OF HOW SIMILAR TWO STRINGS ARE USING THE Levenshtein FORMULA
Public Function StringSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    
    '@AUTHOR: ROBERT TODAR
    '@EXAMPLE: StringSimilarity("Test", "Tester") ->  66.6666666666667
    
    Dim l1 As Integer
    Dim l2 As Integer
    
    l1 = Len(s1)
    l2 = Len(s2)
    
    Dim d() As Integer
    ReDim d(l1, l2)
    Dim I As Integer
    For I = 0 To l1
        d(I, 0) = I
    Next
    
    Dim j As Integer
    For j = 0 To l2
        d(0, j) = j
    Next
    
    Dim min1 As Integer
    Dim min2 As Integer
    For I = 1 To l1
        For j = 1 To l2
            If Mid(s1, I, 1) = Mid(s2, j, 1) Then
                d(I, j) = d(I - 1, j - 1)
            Else
                min1 = d(I - 1, j) + 1
                min2 = d(I, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                min2 = d(I - 1, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                d(I, j) = min1
            End If
        Next
    Next
    
    Dim Levenshtein As Double
    Levenshtein = d(l1, l2)
    
    StringSimilarity = (1 - (Levenshtein / Application.Max(Len(s1), Len(s2)))) * 100
    
End Function


'METHOD THAT ALLOWS A STRING TO BE REPLACED WITH VARIABLES AND SPECIAL CHARACTERS

'@author ROBERT TODAR
'@required REFERENCE TO MICROSOFT SCRIPTING RUNTIME (SCRIPTING.DICTIONARY)
'@example StringInterpolation("${0}\n\t${1}", "First Line", "Tab and Second Line")
Public Function StringInterpolation(ByRef Source As String, ParamArray Args() As Variant) As String
    
    'USE REGULAR EXPRESSION REPLACE SPECIAL CHARATERS (NEWLINE, TAB)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .Pattern = "(^|[^\\])\\n"
        Source = .Replace(Source, "$1" & vbNewLine)
        Source = .Replace(Source, "$1" & vbNewLine)
        .Pattern = "(^|[^\\])\\t"
        Source = regEx.Replace(Source, "$1" & vbTab)
        Source = regEx.Replace(Source, "$1" & vbTab)
    End With
    
    'REPLACE ${#} WITH VALUES STORED IN VARIABLE.
    Dim Index As Integer
    Select Case True
    
        Case TypeName(Args(0)) = "Dictionary":
            
            Dim Dict As Scripting.Dictionary
            Set Dict = Args(0)
            For Index = 0 To Dict.Count - 1
                Source = Replace(Source, "${" & Dict.Keys(Index) & "}", Dict.items(Index), , , vbTextCompare)
            Next Index
            
        Case TypeName(Args(0)) = "Collection":
            Dim Col As Collection
            Set Col = Args(0)
            For Index = 1 To Col.Count
                Source = Replace(Source, "${" & Index - 1 & "}", Col(Index), , , vbTextCompare)
            Next Index
            
        Case Else:
        
            Dim Arr As Variant
            If IsArray(Args(0)) Then
                Arr = Args(0)
            Else
                Arr = Args
            End If
            
            For Index = LBound(Arr, 1) To UBound(Arr, 1)
                Source = Replace(Source, "${" & Index & "}", Arr(Index), , , vbTextCompare)
            Next Index
            
    End Select
    
    StringInterpolation = Source

End Function


Public Function Truncate(ByRef Source As String, MaxLength As Integer) As String
    
    If Len(Source) <= MaxLength Then
        Truncate = Source
        Exit Function
    End If
    
    Const Extention As String = "..."
    Source = Left(Source, MaxLength - Len(Extention)) & Extention
    Truncate = Source
    
End Function

' Find Value between two words in a string
Public Function StringBetween(ByVal Main As String, ByVal Between1 As String, Optional ByVal Between2 As String) As String

    Dim I As Integer
    Dim i2 As Integer
    
    I = InStr(Main, Between1)
    I = I + Len(Between1)
    
    If Between2 = "" Then
        i2 = Len(Main) + 1
    Else
        i2 = InStr(I, Main, Between2)
    End If
    
    StringBetween = Trim(Mid(Main, I, i2 - I))
    
End Function


Public Function StringProperLength(ByVal Value As String, Length As Integer, FillValue As String, Optional AfterString As Boolean = True) As String

    If Len(Value) >= Length Then
        Value = Left(Value, Length)
    Else

        'INSURE INFINITE LOOP DOESN'T OCCUR DUE TO AN EMPTY STRING
        If FillValue = "" Then FillValue = " "

        'LOOP AND ADD EXTRA VALUE
        Do While Len(Value) < Length
            If AfterString = True Then
                Value = Value & FillValue
            ElseIf AfterString = False Then
                Value = FillValue & Value
            End If
        Loop
    End If

    StringProperLength = Value

End Function
