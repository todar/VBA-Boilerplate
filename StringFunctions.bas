Attribute VB_Name = "stringFunctions"
Option Explicit
Option Compare Text
Option Private Module

'@Author: Robert Todar <robert@roberttodar.com>
'@Licence: MIT

'DEPENDENCIES
' - REFERENCE TO SCRIPTING RUNTIME FOR Scripting.Dictionary

'PUBLIC FUNCTIONS
' - StringSimilarity
' - LevenshteinDistance
' - StringInterpolation (also under alias Inject)
' - Truncate
' - StringBetween
' - StringProperLength

'NOTES:
'TODO:

'EXAMPLES OF ALL THE FUNCTIONS
Private Sub StringFunctionExamples()
    
    '@AUTHOR: ROBERT TODAR
    
    StringSimilarity "Test", "Tester"        '->  66.6666666666667
    LevenshteinDistance "Test", "Tester"     '->  2
    StringInterpolation "${0}\n\t${1}", "First", "Tab and Second" '-> First
                                                                  '->   Tab and Second
                                                                  
    Truncate "This is a long sentence", 10                '-> "This is..."
    StringBetween "Robert Paul Todar", "Robert", "Todar"  '-> "Paul"
    StringProperLength "1001", 6, "0", True               '-> "100100"
    
    'Inject() is a copy of StringInterpolation, this alias is easier to remember (shorter too!)
    'Here is an example using a dictionary!
    Dim Person As New Scripting.Dictionary
    Person("Name") = "Robert"
    Person("Age") = 30
    
    'REMEMBER, DICTIONARY KEYS ARE CASE SENSITIVE!
    Debug.Print Inject("Hello,\nMy name is ${Name} and I am ${Age}!", Person)
        '-> Hello,
        '-> My name is Robert and I am 30!
End Sub

'******************************************************************************************
' PUBLIC FUNCTIONS
'******************************************************************************************

'THIS RETURNS A PERCENTAGE OF HOW SIMILAR TWO STRINGS ARE USING THE Levenshtein FORMULA
Public Function StringSimilarity(ByVal FirstString As String, ByVal SecondString As String) As Double
    
    '@AUTHOR: ROBERT TODAR
    '@EXAMPLE: StringSimilarity("Test", "Tester") ->  66.6666666666667
    
    'LEVENSHTEIN IS THE DISTANCE BETWEEN TWO SEQUENCES
    Dim Levenshtein As Double
    Levenshtein = LevenshteinDistance(FirstString, SecondString)
    
    'CONVERT LEVENSHTEIN INTO A PERCENTAGE(0 TO 100)
    StringSimilarity = (1 - (Levenshtein / Application.Max(Len(FirstString), Len(SecondString)))) * 100
    
End Function

'LEVENSHTEIN IS THE DISTANCE BETWEEN TWO SEQUENCES OF WORDS
Public Function LevenshteinDistance(ByVal FirstString As String, ByVal SecondString As String) As Double
    
    '@AUTHOR: ROBERT TODAR
    '@REF: https://www.cuelogic.com/blog/the-levenshtein-algorithm
    '@EXAMPLE: LevenshteinDistance("Test", "Tester") ->  2
    
    Dim FirstLength As Integer
    FirstLength = Len(FirstString)

    Dim SecondLength As Integer
    SecondLength = Len(SecondString)
    
    'PREPARE DISTANCE ARRAY MATRIX WITH THE PROPER INDEXES
    Dim Distance() As Integer
    ReDim Distance(FirstLength, SecondLength)
    
    Dim Index As Integer
    For Index = 0 To FirstLength
        Distance(Index, 0) = Index
    Next
    
    Dim InnerIndex As Integer
    For InnerIndex = 0 To SecondLength
        Distance(0, InnerIndex) = InnerIndex
    Next
    
    'OUTER LOOP IS FOR THE FIRST STRING
    For Index = 1 To FirstLength

        'INNER LOOP IS FOR THE SECOND STRING
        For InnerIndex = 1 To SecondLength

            'CHARACTER MATCHES EXACTLY
            If Mid(FirstString, Index, 1) = Mid(SecondString, InnerIndex, 1) Then
                Distance(Index, InnerIndex) = Distance(Index - 1, InnerIndex - 1)
            
            'CHARACTER IS OFF, OFFSET THE MATRIX BY THE APPROPRIATE NUMBER
            Else
                Dim Min1 As Integer
                Min1 = Distance(Index - 1, InnerIndex) + 1

                Dim Min2 As Integer
                Min2 = Distance(Index, InnerIndex - 1) + 1

                If Min2 < Min1 Then
                    Min1 = Min2
                End If
                Min2 = Distance(Index - 1, InnerIndex - 1) + 1
    
                If Min2 < Min1 Then
                    Min1 = Min2
                End If
                Distance(Index, InnerIndex) = Min1

            End If
        Next
    Next
    
    'LEVENSHTEIN IS THE LAST INDEX OF THE ARRAY
    LevenshteinDistance = Distance(FirstLength, SecondLength)
    
End Function

'METHOD THAT ALLOWS A STRING TO BE REPLACED WITH VARIABLES AND SPECIAL CHARACTERS
Public Function StringInterpolation(ByRef Source As String, ParamArray Args() As Variant) As String
    
    '@AUTHOR: ROBERT TODAR
    '@REQUIRED: REFERENCE TO MICROSOFT SCRIPTING RUNTIME (SCRIPTING.DICTIONARY)
    '@EXAMPLE: StringInterpolation("${0}\n\t${1}", "First Line", "Tab and Second Line")
    
    'USE REGULAR EXPRESSION REPLACE SPECIAL CHARATERS (NEWLINE, TAB)
    'NOTE THE REPLACE IS RAN TWICE SINCE IT'S POSSIBLE FOR BACK TO BACK PATTERNS.
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .Global = True
        .Pattern = "((?:^|[^\\])(?:\\{2})*)(?:\\n)+"
        Source = .Replace(Source, "$1" & vbNewLine)
        Source = .Replace(Source, "$1" & vbNewLine)
        .Pattern = "((?:^|[^\\])(?:\\{2})*)(?:\\t)+"
        Source = RegEx.Replace(Source, "$1" & vbTab)
        Source = RegEx.Replace(Source, "$1" & vbTab)
    End With
    
    'REPLACE ${#} WITH VALUES STORED IN VARIABLE.
    Dim Index As Integer
    Select Case True
    
        Case IsMissing(Args)
    
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

'METHOD THAT ALLOWS A STRING TO BE REPLACED WITH VARIABLES AND SPECIAL CHARACTERS
'SHORTENED NAME TO StringInterpolation... FOR EASE OF USE.
Public Function Inject(ByRef Source As String, ParamArray Args() As Variant) As String
    
    '@AUTHOR: ROBERT TODAR
    '@REQUIRED: REFERENCE TO MICROSOFT SCRIPTING RUNTIME (SCRIPTING.DICTIONARY)
    '@EXAMPLE: Inject("${0}\n\t${1}", "First Line", "Tab and Second Line")
    
    'USE REGULAR EXPRESSION REPLACE SPECIAL CHARATERS (NEWLINE, TAB)
    'NOTE THE REPLACE IS RAN TWICE SINCE IT'S POSSIBLE FOR BACK TO BACK PATTERNS.
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .Global = True
        .Pattern = "((?:^|[^\\])(?:\\{2})*)(?:\\n)+"
        Source = .Replace(Source, "$1" & vbNewLine)
        Source = .Replace(Source, "$1" & vbNewLine)
        .Pattern = "((?:^|[^\\])(?:\\{2})*)(?:\\t)+"
        Source = RegEx.Replace(Source, "$1" & vbTab)
        Source = RegEx.Replace(Source, "$1" & vbTab)
    End With
    
    'REPLACE ${#} WITH VALUES STORED IN VARIABLE.
    Dim Index As Integer
    Select Case True
    
        Case IsMissing(Args)
    
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
    
    Inject = Source

End Function

'CREATE A MAX LENGHT OF STRING AND RETURN IT WITH EXTENSION
Public Function Truncate(ByRef Source As String, MaxLength As Integer) As String
    
    'AUTHOR: ROBERT TODAR
    'EXAMPLE: Truncate("This is a long sentence", 10)  -> "This is..."
    
    If Len(Source) <= MaxLength Then
        Truncate = Source
        Exit Function
    End If
    
    Const Extention As String = "..."
    Source = Left(Source, MaxLength - Len(Extention)) & Extention
    Truncate = Source
    
End Function

'FIND A STRING BETWEEN TWO WORDS
Public Function StringBetween(ByVal Main As String, ByVal Between1 As String, Optional ByVal Between2 As String) As String
    
    'AUTHOR: ROBERT TODAR
    'EXAMPLE: StringBetween("Robert Paul Todar", "Robert", "Todar")  -> "Paul"
    
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

'RETURNS A STRING WITH THE PROPER PADDING ON EACH SIDE.
Public Function StringProperLength(ByVal Value As String, ByVal Length As Integer, ByVal FillValue As String, Optional AfterString As Boolean = True) As String
    
    'AUTHOR: ROBERT TODAR
    'EXAMPLE: StringProperLength("1001", 6, "0", True) -> "100100"
    
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
