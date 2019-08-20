<p align="center">
    <img width="350px" alt="function meme" src="https://i.pinimg.com/736x/2e/e7/b3/2ee7b37349f798c3460e244143bdd0bc--math-puns-math-humor.jpg">
    <h1 align="center">VBA Function Library</h1>
</p>

You've found my VBA Libray GitHub repository, which contains functions to help make programming in VBA easier.

> This repository is currently under construction, but will be intended to be a place to help make VBA more open source.


## Table of Contents

  1. [Style Example](#style-example)
  2. [Array Examples](#array-functions)
  3. [String Examples](#string-functions)

## Style Example

  Below is an example function of how I entend all my code to be written. 

  > Please see my [Style Guide](https://github.com/todar/VBA) for how to write clean and maintainable VBA code.

  ```vb
'/**
' * A simple Dictionary Factory.
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @param {...Variant} keyValuePairs - Key must be valid dictionary key, value can be anything.
' * @ref {Scripting.Dictionary} MicroSoft Scripting Runtime
' * @example ToDictionary("Name", "Robert", "Age", 30) '--> { "Name": "Robert, "Age": 30 }
' */
Public Function ToDictionary(ParamArray keyValuePairs() As Variant) As Scripting.Dictionary
    ' Check to see that key/value pairs passed in (an even number).
    If arrayLength(CVar(keyValuePairs)) Mod 2 <> 0 Then
        Err.Raise 5, "ToDictionary", "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
    End If
    
    ' Add key values to the return Dictionary.
    Set ToDictionary = New Scripting.Dictionary
    Dim index As Long
    For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        ToDictionary.Add keyValuePairs(index), keyValuePairs(index + 1)
    Next index
End Function

'/**
' * Helper function that Returns the number of elements in an Array.
' * @param {Array<Variant>} source - Array that you want to return the length of.
' */
Private Function ArrayLength(ByRef source As Variant) As Long
    ArrayLength = UBound(source) - LBound(source) + 1
End Function
  ```

----

## Array Functions

  ```vb
  '/**
  ' * Examples of different array functions.
  ' * @author Robert Todar <robert@roberttodar.com>
  ' */
  Private Sub arrayFunctionExamples()
      ' * Made array single letter for ease of reading. Do not do this is production code!
      Dim A As Variant
      
      ' Single dim functions to manipulate
      ArrayPush A, "Banana", "Apple", "Carrot" '--> Banana,Apple,Carrot
      ArrayPop A                               '--> Banana,Apple --> returns Carrot
      ArrayUnShift A, "Mango", "Orange"        '--> Mango,Orange,Banana,Apple
      ArrayShift A                             '--> Orange,Banana,Apple
      ArraySplice A, 2, 0, "Coffee"            '--> Orange,Banana,Coffee,Apple
      ArraySplice A, 0, 1, "Mango", "Coffee"   '--> Mango,Coffee,Banana,Coffee,Apple
      ArrayRemoveDuplicates A                  '--> Mango,Coffee,Banana,Apple
      ArraySort A                              '--> Apple,Banana,Coffee,Mango
      ArrayReverse A                           '--> Mango,Coffee,Banana,Apple
      
      ' Functions for Array properties
      ArrayLength A                            '--> 4
      ArrayIndexOf A, "Coffee"                 '--> 1
      ArrayIncludes A, "Banana"                '--> True
      ArrayContains A, Array("Test", "Banana") '--> True
      ArrayContainsEmpties A                   '--> False
      ArrayDimensionLength A                   '--> 1 (single dim array)
      IsArrayEmpty A                           '--> False
      
      ' Example where you can flatten a jagged array.
      ' @note You can also spread dictionaries and collections as well.
      A = Array(1, 2, 3, Array(4, 5, 6, Array(7, 8, 9)))
      A = ArraySpread(A)                       '--> 1,2,3,4,5,6,7,8,9
      
      ' Functions dealing with Math operators.
      ArraySum A                               '--> 45
      ArrayAverage A                           '--> 5
      
      ' Filter uses REGEX patterns.
      A = Array("Banana", "Coffee", "Apple", "Carrot", "Canolope")
      A = ArrayFilter(A, "^Ca|^Ap")
      
      ' Array to string works with both single and double DIM arrays.
      Debug.Print ArrayToString(A)
  End Sub
  ```
  **[⬆ back to top](#table-of-contents)**

## String Functions

  ```vb
'/**
' * Examples of different string functions.
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
  ```
  **[⬆ back to top](#table-of-contents)**
