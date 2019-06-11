<p align="center">
    <img width="200px" alt="function meme" src="https://i.pinimg.com/736x/2e/e7/b3/2ee7b37349f798c3460e244143bdd0bc--math-puns-math-humor.jpg">
    <h1 align="center">VBA Function Library</h1>
</p>

You've found my VBA Libray GitHub repository, which contains functions to help programming in VBA easier.

This repository is currently under construction, but will be intended to be a place to help make VBA more open source.

## Index

1. [Style Guide](#style-guide)
2. [Array Examples](#arrays)
3. [String Examples](#strings)

## Style Guide

Below is an example of how I entend all my code to be written. 

If you'd like to contribute please try to format similar to this for consistency.

```vb
'A simple Dictionary Factory.
Private Function ToDictionary(ParamArray keyValuePairs() As Variant) As Scripting.Dictionary
    
    '@author: Robert Todar <robert@roberttodar.com>
    '@ref: MicroSoft Scripting Runtime
    '@example: ToDictionary("Name", "Robert", "Age", 30) '--> { "Name": "Robert, "Age": 30 }
    
    'Get length of array to check to see if there are valid parameters.
    Dim ArrayLenght As Long
    ArrayLenght = UBound(keyValuePairs) - LBound(keyValuePairs) + 1
    
    'Check to see that key/value pairs passed in (an even number).
    If ArrayLenght Mod 2 <> 0 Then
        Err.Raise 5, "ToDictionary", "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
    End If
    
    'Add key values to the return Dictionary.
    Set ToDictionary = New Scripting.Dictionary
    Dim Index As Long
    For Index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        ToDictionary.Add keyValuePairs(Index), keyValuePairs(Index + 1)
    Next Index
    
End Function
```

Above the function should be a simple description of what the function does.

Function names and parameters should be **descripitive** and can be long if needed. **Use action word**s.

Just inside the function is where I will put important details. This could be author, library references, notes, Ect. I've styled this to be similar to [JSDoc documentation](https://devdocs.io/jsdoc/). 

Functions should be as small as possible designed to resusable. This means they should be very readable.

Notes should be clear and full sentences. Explain anything that doesn't immediatly make sence from the code.


## Arrays

Example Single Dim Array functions
```VB
Private Sub ArrayFunctionExamples()
    
    'keeping this variable name short for easy reading...
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

## Strings

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

## List of Functions

>These are the ones I felt are easy for others to use as well, will plan on modifing some of my specific functions and add them once they are more distributable.

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
