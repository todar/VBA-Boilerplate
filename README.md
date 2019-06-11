<h1 align="center">VBA Function Library</h1>

---

You found a repository that is created for storing helpful functions for VBA.

## Examples
Example Single Dim Array functions
```VB
Private Sub ArrayFunctionExamples()
    
    Dim A As Variant
    
    'SINGLE DIM FUNCTIONS
    ArrayPush A, "Banana", "Apple", "Carrot" '--> Banana,Apple,Carrot
    ArrayPop A                               '--> Banana,Apple --> returns Carrot
    ArrayUnShift A, "Mango", "Orange"        '--> Mango,Orange,Banana,Apple
    ArrayShift A                             '--> Orange,Banana,Apple
    ArraySplice A, 2, 0, "Coffee"            '--> Orange,Banana,Coffee,Apple
    ArraySplice A, 0, 1, "Mango", "Coffee"   '--> Mango,Coffee,Banana,Coffee,Apple
    ArrayRemoveDuplicates A                  '--> Mango,Coffee,Banana,Apple
    ArraySort A                              '--> Apple,Banana,Coffee,Mango
    ArrayReverse A                           '--> Mango,Coffee,Banana,Apple
    ArrayIndexOf A, "Coffee"                 '--> 1
    ArrayIncludes A, "Banana"                '--> True
    
    Debug.Print ArrayToString(A)
    
End Sub
```

Example String Funtion:
```VB
Private Sub ExamplesOfStringInterpolation()

    'Dictionaries are the best to use, since you can use the keys!!
    Dim Dict As Object
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict("name") = "Robert"
    Dict("age") = 29
    Debug.Print StringInterpolation("Hello, my name is ${name}\nand I am ${age} years old", Dict)
    
    'Collection example
    Dim Col As New Collection
    Col.Add "Robert"
    Col.Add 29
    Debug.Print StringInterpolation("Hello, my name is ${0} and I am ${1} years old", Col)
    
    'Array example
    Dim Arr As Variant
    Arr = Array("Robert", 29)
    Debug.Print StringInterpolation("Hello, my name is ${0} and I am ${1} years old", Arr)
   
    'Passing Variables into the parameters (A cool and fast way of doing it!)
    Debug.Print StringInterpolation("Hello, my name is ${0} and I am ${1} years old", "Robert", 29)
    
End Sub
```
```vb
Public Function StringInterpolation(ByRef Source As String, ParamArray Args() As Variant) As String
    
    '@AUTHOR: ROBERT TODAR
    '@REQUIRED: REFERENCE TO MICROSOFT SCRIPTING RUNTIME (SCRIPTING.DICTIONARY)
    '@EXAMPLE: StringInterpolation("${0}\n\t${1}", "First Line", "Tab and Second Line")
    
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
                Source = Replace(Source, "${" & Dict.Keys(Index) & "}", Dict.Items(Index), , , vbTextCompare)
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

```

>These are the ones I felt are easy for others to use as well, will plan on modifing some of my specific functions and add them once they are more distributable.

Funtions:
* StringInterpolation - Allows putting variables in string. Makes for easy concatination.
* CenterForm - Centers userforms to excel application. Helpful for multiple monitors.
* ArrayPop - Removes last array element in single dim array. Returns popped element.
* ArrayPush - Adds element to the end of a single dim array. Returns new lenght of array.
* ArrayShift - Removes element from the start of a single dim array. Retrurns removed element.
* ArrayUnShift - Adds element to the start of an array, returns array lenght.
* ArrayQuery - Query a 2 dim array using ado and Microsoft.Jet. Saves array as csv to textfile. Really helpful!!
* ArrayDimensionLength
* ArrayExtract
* ArrayFromRecordset - Used with ArrayQuery. Also helpful to get full array from and Ado recorsets.
* ArrayGetColumnNumber
* ArrayIncludes
* ArrayIndexOf
* ArrayRemoveDuplicates
* ArrayReverse
* ArraySort
* ArraySplice
* ArrayToCSVFile
* ArrayToRange
* ArrayToString
* ArrayToTextFile
* ArrayTranspose
* ConvertToArray
* CollectionAddUniue - Adds to a collection if the string value is unique.
* ClipboardSet - Takes a string and puts it into the clipboard.
* ClipboardGet - Gets text from clipboard and sets it to a string value.
* InsertSlicer - Adds slicer to a table.
* FindColumnData - Searchs for heading in top row, sets ranges to the columns data.
* FindGroup - Finds groups of values in a column. (Column must be sorted first).
* Findheading - Searchs for the heading in the first row.
* FreezeTopRow - self explanitory :)
* Pause - Pause on code off of mil secs.
* OpenAnyFile - Open any file with it's default application.
* ProperLength - Ensures lenght of string is as long as you set it to be. Will add fill value at begining or end based on parameters.
* String between - Find the value between two words in a string.
* TextboxWordSelect - Selects text in a textbox in a userform.
* StringSimilarity - Compares two strings and returns a percentage of how similar they are. Uses the Levenshtein formula.
