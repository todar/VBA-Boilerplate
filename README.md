# VBA-Functions
Functions I have created that could be helpful for others to use. 

Example Funtion:

```VB
'=========================================================================
' Example call
'=========================================================================
Sub testingTemplateStrings()

    Dim dict As object
    Dim col As New Collection
    Dim arr As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")

    'Dictionaries are the best to use, since you can use the keys!! 
        dict("name") = "Robert"
        dict("age") = 29
        Debug.Print StringTemplate("Hello, my name is ${name}\nand I am ${age} years old", dict)
    
    'Collection example
        col.Add "Robert"
        col.Add 29
        Debug.Print StringTemplate("Hello, my name is ${0} and I am ${1} years old", col)
    
    'Array example
        arr = Array("Robert", 29)
        Debug.Print StringTemplate("Hello, my name is ${0} and I am ${1} years old", arr)
    
    
    'Passing Variables into the parameters (A cool and fast way of doing it!) 
        Debug.Print StringTemplate("Hello, my name is ${0} and I am ${1} years old", "Robert", 29)
    
End Sub

'=========================================================================
' WAY OF CREATING A STRING SIMILAR TO HOW JS WORKS
'=========================================================================
Function StringTemplate(S As String, ParamArray Args() As Variant) As String

    Dim i As Integer
    Dim lb As Integer
    Dim ub As Integer
    Dim Obj As Dictionary
    Dim Arr As Variant
    Dim Reduce As Integer
    Dim regEx As Object

    'REGULAR EXPRESSION REPLACE SPECIAL CHARATERS
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True

    'newline
    regEx.Pattern = "(^|[^\\])\\n"
    S = regEx.Replace(S, "$1" & vbNewLine)

    'tab
    regEx.Pattern = "(^|[^\\])\\t"
    S = regEx.Replace(S, "$1" & vbTab)


    '============================================================
    ' PASSED IN AS DICTIONARY (PASSED IN AS VARIABLES)
    '============================================================
    If TypeName(Args(0)) = "Dictionary" Then

        Set Obj = Args(0)

        For i = 0 To Obj.Count - 1
            S = Replace(S, "${" & Obj.Keys(i) & "}", Obj.Items(i), , , vbTextCompare)
        Next i
        StringTemplate = S
        Exit Function

    End If

    '============================================================
    ' PASSED IN AS COLLECTION/ARRAY/ParamArray
    '============================================================
    If TypeName(Args(0)) = "Collection" Then
        lb = 1
        ub = Args(0).Count
        Set Arr = Args(0)
        Reduce = 1
    ElseIf IsArray(Args(0)) Then
        lb = LBound(Args(0), 1)
        ub = UBound(Args(0), 1)
        Arr = Args(0)
    Else
        lb = LBound(Args, 1)
        ub = UBound(Args, 1)
        Arr = Args()
    End If

    'LOOP
    For i = lb To ub
        S = Replace(S, "${" & (i - Reduce) & "}", Arr(i), , , vbTextCompare)
    Next i

    StringTemplate = S

End Function

```

>These are the ones I felt are easy for others to use as well, will plan on modifing some of my specific functions and add them once they are more distributable.

Funtions:
* StringTemplate - Allows putting variables in string. Makes for easy concatination.
* CenterForm - Centers userforms to excel application. Helpful for multiple monitors.
* ArrayPop - Removes last array element in single dim array. Returns popped element.
* ArrayPush - Adds element to the end of a single dim array. Returns new lenght of array.
* ArrayShift - Removes element from the start of a single dim array. Retrurns removed element.
* ArrayUnShift - Adds element to the start of an array, returns array lenght.
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
