Attribute VB_Name = "StringFunctions"
'/**
' * This is a utility library for string functions.
' *
' * @author Robert Todar <robert@robertodar.com> <https://github.com/todar>
' * @licence MIT
' * @ref {Microsoft Scripting Runtime} Scripting.Dictionary
' * @ref {Microsoft VBScript Regular Expressions 5.5} [RegExp, Match]
' */
Option Explicit
Option Compare Text

'/**
' * Tests
' * @author Robert Todar <robert@robertodar.com>
' */
Private Sub testsForStringFunctions()
    Debug.Print StringSimilarity("Test", "Tester")                     '->  66.6666666666667
    Debug.Print LevenshteinDistance("Test", "Tester")                  '->  2
    Debug.Print Truncate("This is a long sentence", 10)                '-> "This is..."
    Debug.Print StringBetween("Robert Paul Todar", "Robert", "Todar")  '-> "Paul"
    Debug.Print StringPadding("1001", 6, "0", True)                    '-> "100100"
    Debug.Print Inject("Hello,\nMy name is {Name} and I am {Age}!", "Robert", 31)
        '-> Hello,
        '-> My name is Robert and I am 30!
End Sub

'/**
' * This returns a percentage of how similar two strings are using the levenshtein formula.
' *
' * @author Robert Todar <robert@robertodar.com>
' * @example StringSimilarity("Test", "Tester") ->  66.6666666666667
' */
Public Function StringSimilarity(ByVal firstString As String, ByVal secondString As String) As Double
    ' Levenshtein is the distance between two sequences
    Dim levenshtein As Double
    levenshtein = LevenshteinDistance(firstString, secondString)
    
    ' Convert levenshtein into a percentage(0 to 100)
    StringSimilarity = (1 - (levenshtein / Application.Max(Len(firstString), Len(secondString)))) * 100
End Function

'/**
' * Levenshtein is the distance between two sequences of words.
' *
' * @author Robert Todar <robert@robertodar.com>
' * @see <https://www.cuelogic.com/blog/the-levenshtein-algorithm>
' * @example LevenshteinDistance("Test", "Tester") ->  2
' */
Public Function LevenshteinDistance(ByVal firstString As String, ByVal secondString As String) As Double
    Dim firstLength As Integer
    firstLength = Len(firstString)

    Dim secondLength As Integer
    secondLength = Len(secondString)
    
    ' Prepare distance array matrix with the proper indexes
    Dim distance() As Integer
    ReDim distance(firstLength, secondLength)
    
    Dim index As Integer
    For index = 0 To firstLength
        distance(index, 0) = index
    Next
    
    Dim innerIndex As Integer
    For innerIndex = 0 To secondLength
        distance(0, innerIndex) = innerIndex
    Next
    
    ' Outer loop is for the first string
    For index = 1 To firstLength

        ' Inner loop is for the second string
        For innerIndex = 1 To secondLength

            ' Character matches exactly
            If Mid(firstString, index, 1) = Mid(secondString, innerIndex, 1) Then
                distance(index, innerIndex) = distance(index - 1, innerIndex - 1)
            
            ' Character is off, offset the matrix by the appropriate number.
            Else
                Dim min1 As Integer
                min1 = distance(index - 1, innerIndex) + 1

                Dim min2 As Integer
                min2 = distance(index, innerIndex - 1) + 1

                If min2 < min1 Then
                    min1 = min2
                End If
                min2 = distance(index - 1, innerIndex - 1) + 1
    
                If min2 < min1 Then
                    min1 = min2
                End If
                distance(index, innerIndex) = min1

            End If
        Next
    Next
    
    ' Levenshtein is the last index of the array.
    LevenshteinDistance = distance(firstLength, secondLength)
End Function

'/**
' * Returns a new cloned string that replaced special {keys} with its associated pair value.
' * Keys can be anything since it goes off of the index, so variables must be in proper order!
' * Can't have whitespace in the key.
' * Also Replaces "\t" with VbTab and "\n" with VbNewLine
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @ref {Microsoft Scripting Runtime} Scripting.Dictionary
' * @ref {Microsoft VBScript Regular Expressions 5.5} [RegExp, Match]
' * @example Inject("Hello, {name}!\nJS Object = {name: {name}, age: {age}}\n", "Robert", 31)
' */
Public Function Inject(ByVal source As String, ParamArray values() As Variant) As String
    ' Want to get a copy and not mutate original
    Inject = source

    Dim regEx As RegExp
    Set regEx = New RegExp ' Late Binding would be: CreateObject("vbscript.regexp")
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True

        ' This section is only when user passes in variables
        If Not IsMissing(values) Then

            ' Looking for pattern like: {key}
            ' First capture group is the full pattern: {key}
            ' Second capture group is just the name:    key
            .Pattern = "(?:^|[^\\])(\{([A-zÀ-ÿ0-9\s]*)\})"

            ' Used to make sure there are even number of uniqueKeys and values.
            Dim keys As New Scripting.Dictionary

            Dim keyMatch As Match
            For Each keyMatch In .Execute(Inject)

                ' Extract key name
                Dim key As Variant
                key = keyMatch.SubMatches(1)

                ' Only want to increment on unique keys.
                If Not keys.Exists(key) Then

                    If (keys.Count) > UBound(values) Then
                        Err.Raise 9, "Inject", "Inject expects an equal amount of keys to values. Keys found: " & Join(keys.keys, ", ") & ", " & key
                    End If

                    ' Replace {key} with the pairing value.
                    Inject = Replace(Inject, keyMatch.SubMatches(0), values(keys.Count))

                    ' Add key to make sure it isn't looped again.
                    keys.Add key, vbNullString

               End If
            Next
        End If

        ' Replace extra special characters. Must allow code above to run first!
        .Pattern = "(^|[^\\])\{"
        Inject = .Replace(Inject, "$1" & "{")

        .Pattern = "(^|[^\\])\\t"
        Inject = .Replace(Inject, "$1" & vbTab)

        .Pattern = "(^|[^\\])\\n"
        Inject = .Replace(Inject, "$1" & vbNewLine)

        .Pattern = "(^|[^\\])\\"
        Inject = .Replace(Inject, "$1" & "")
    End With
End Function

'/**
' * Create a max lenght of string and return it with extension.
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @example Truncate("This is a long sentence", 10)  -> "This is..."
' */
Public Function Truncate(ByRef source As String, maxLength As Integer) As String
    If Len(source) <= maxLength Then
        Truncate = source
        Exit Function
    End If
    
    Const extention As String = "..."
    source = Left(source, maxLength - Len(extention)) & extention
    Truncate = source
End Function

'/**
' * Find string between two words.
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @example StringBetween("Robert Paul Todar", "Robert", "Todar")  -> "Paul"
' */
Public Function StringBetween(ByVal main As String, ByVal between1 As String, Optional ByVal between2 As String) As String
    Dim startIndex As Integer
    startIndex = InStr(main, between1) + Len(between1)
    
    Dim endIndex As Integer
    endIndex = IIf(between2 = vbNullString, Len(main) + 1, InStr(startIndex, main, between2))
    
    StringBetween = Trim(Mid(main, startIndex, endIndex - startIndex))
End Function

'/**
' * Returns a string with the proper padding on each side.
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @example StringPadding("1001", 6, "0", True) -> "100100"
' */
Public Function StringPadding(ByVal value As String, ByVal length As Integer, ByVal fillValue As String, Optional afterString As Boolean = True) As String
    If Len(value) >= length Then
        value = Left(value, length)
    Else
        ' Insure infinite loop doesn't occur due to an empty string.
        If fillValue = vbNullString Then fillValue = " "

        ' Add extra value
        Do While Len(value) < length
            value = IIf(afterString, value & fillValue, fillValue & value)
        Loop
    End If
    StringPadding = value
End Function


'/**
' * Tests for ToString Function.
' */
Private Sub testToStringFunction()
    ' Test values
    Debug.Print ToString("String")
    Debug.Print ToString(31)
    
    ' Test Array
    Debug.Print ToString(Array(1, 2, 3, 4))
    
    ' Test collections
    Dim col As New Collection
    col.Add "item", "key"
    Debug.Print ToString(col)
    
    ' Test Dictionary
    Dim dic As New Scripting.Dictionary
    dic.Add "Name", "Robert"
    dic.Add "Age", 31
    Debug.Print ToString(dic)
    
    ' Test objects
    Dim fso As New FileSystemObject
    Debug.Print ToString(fso)
End Sub

'/**
' * Created this and helper functions to easily read different containers.
' *
' * @author Robert Todar <robert@roberttodar.com>
' */
Public Function ToString(ByVal source As Variant) As String
    Const delimiter As String = ", "
    
    Select Case True
        Case TypeName(source) = "Dictionary"
            ToString = toStingDictionary(source, delimiter)
        
        Case TypeName(source) = "Collection"
            ToString = toStingCollection(source, delimiter)
        
        Case IsArrayDimension(source, 1)
            ToString = toStingSingleDimArray(source, delimiter)
        
        Case IsArrayDimension(source, 2)
            ToString = toStingTwoDimArray(source, delimiter)
        
        Case IsObject(source)
            ToString = TypeName(source) & " {}"
            
        Case TypeName(source) = "String"
            ToString = """" & Replace(Replace(source, vbNewLine, "\n"), vbTab, "\t") & """"
        
        Case IsNull(source)
            ToString = ""
        
        Case Else 'IsNumeric(source), TypeName(source) = "Boolean"
            ToString = source
            
    End Select
End Function

'/**
' * Helper function to add lines for ToString()
' */
Private Function AddLineIfNeeded(ByVal source As Variant) As String
        If TypeName(source) = "Dictionary" _
            Or TypeName(source) = "Collection" _
            Or IsArrayDimension(source, 1) _
            Or IsArrayDimension(source, 2) Then
            
            AddLineIfNeeded = vbNewLine & "  "
        End If
End Function

'/**
' * Dictionary as a string
' */
Private Function toStingDictionary(ByVal source As Scripting.Dictionary, ByVal delimiter As String) As String
    toStingDictionary = "{"
    
    Dim key As Variant
    For Each key In source.keys
        toStingDictionary = toStingDictionary & AddLineIfNeeded(source(key)) & """" & key & """" & ": " & ToString(source(key)) & delimiter
    Next key
    toStingDictionary = Left(toStingDictionary, Len(toStingDictionary) - Len(delimiter)) & IIf(InStr(toStingDictionary, vbNewLine), vbNewLine, "") & "}"
End Function

'/**
' * Collection as a string
' */
Private Function toStingCollection(ByVal source As Collection, ByVal delimiter As String) As String
    toStingCollection = "{"
    
    Dim item As Variant
    For Each item In source
        toStingCollection = toStingCollection & AddLineIfNeeded(item) & ToString(item) & delimiter
    Next item
    toStingCollection = Left(toStingCollection, Len(toStingCollection) - Len(delimiter)) & IIf(InStr(toStingCollection, vbNewLine), vbNewLine, "") & "}"
End Function

'/**
' * Single Array as a string
' */
Private Function toStingSingleDimArray(ByVal source As Variant, ByVal delimiter As String) As String
    toStingSingleDimArray = "["
    
    Dim index As Long
    For index = LBound(source) To UBound(source)
        toStingSingleDimArray = toStingSingleDimArray & AddLineIfNeeded(source(index)) & ToString(source(index)) & IIf(index < UBound(source), delimiter, "")
    Next index
    toStingSingleDimArray = toStingSingleDimArray & IIf(InStr(toStingSingleDimArray, vbNewLine), vbNewLine, "") & "]"
End Function

'/**
' * Two Dim Array as a string
' */
Private Function toStingTwoDimArray(ByVal source As Variant, ByVal delimiter As String) As String
    toStingTwoDimArray = "[" & vbNewLine
    Dim rowIndex As Long
    For rowIndex = LBound(source) To UBound(source)
        toStingTwoDimArray = toStingTwoDimArray & "  ["
        
        ' Add row elements to the string
        Dim colIndex As Long
        For colIndex = LBound(source, 2) To UBound(source, 2)
            toStingTwoDimArray = toStingTwoDimArray & ToString(source(rowIndex, colIndex)) & IIf(colIndex < UBound(source, 2), delimiter, "")
        Next colIndex
        
        toStingTwoDimArray = toStingTwoDimArray & "]" & IIf(rowIndex < UBound(source), "," & vbNewLine, "")
    Next rowIndex
    toStingTwoDimArray = toStingTwoDimArray & vbNewLine & "]"
End Function

'/**
' * Helper function to see if array is two dim or single
' */
Private Function IsArrayDimension(ByVal source As Variant, ByVal dimension As Long) As Boolean
    If IsArray(source) Then
        IsArrayDimension = (dimension = arrayDimensionLength(source))
    End If
End Function

'/**
' * Helper function to get the legnth of the dimension of an array.
' */
Private Function arrayDimensionLength(ByVal sourceArray As Variant) As Long
    ' Run loop until error. Remove one and it gives the array dimension =)
    On Error GoTo Catch
    Do
        Dim iterator As Long
        iterator = iterator + 1
        
        Dim test As Long
        test = UBound(sourceArray, iterator)
    Loop
Catch:
    On Error GoTo 0
    arrayDimensionLength = iterator - 1
End Function

'/**
' * Returns the number of elements in an Array.
' */
Private Function arrayLength(ByRef source As Variant) As Long
    arrayLength = UBound(source) - LBound(source) + 1
End Function



