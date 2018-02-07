# VBA-Functions
Functions I have created that could be helpful for others to use. 

Example Funtion:

```VB
'=========================================================================
' WAY OF CREATING A STRING USING THE WAY C# & JS PROGRAMMING LANGUAGES DO
' EXAMPLE:
'   myString("This {0} a test of\n {1} Function", Array("is", "myString"))
' RESULT:
'   "This is a test of" & vbnewline & "myString Function"
'=========================================================================
Public Function myString(Original As String, Optional Arr As Variant) As String
    
    Dim I As Integer
    
    'SPECIAL CHARACTERS TO MAKE NEWLINES\TABS EASIER
    Original = Replace(Original, "\n", vbNewLine)
    Original = Replace(Original, "\t", vbTab)
    
    'REPLACE WITH ARRAYS
    If IsArray(Arr) Then
        For I = LBound(Arr, 1) To UBound(Arr, 1)
            Original = Replace(Original, "{" & I & "}", Arr(I))
        Next I
     End If
    myString = Original

End Function
```

>These are the ones I felt are easy for others to use as well, will plan on modifing some of my specific functions and add them once they are more distributable.

Funtions:
* myString - Allows putting variables in string. Makes for easy concatination.
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
