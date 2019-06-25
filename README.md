<p align="center">
    <img width="350px" alt="function meme" src="https://i.pinimg.com/736x/2e/e7/b3/2ee7b37349f798c3460e244143bdd0bc--math-puns-math-humor.jpg">
    <h1 align="center">VBA Function Library</h1>
</p>

You've found my VBA Libray GitHub repository, which contains functions to help make programming in VBA easier.

> Please see my [Style Guide](https://github.com/todar/VBA) for how to write clean and maintainable code using VBA.

> This repository is currently under construction, but will be intended to be a place to help make VBA more open source.


## Table of Contents

  1. [Style Example](#style-example)
  2. [Array Examples](#array-functions)
  3. [String Examples](#string-functions)

## Style Example

  Below is an example function of how I entend all my code to be written. 

  > If you'd like to contribute please try to format similar to this for consistency.

  ```vb
' A simple Dictionary Factory.
Private Function ToDictionary(ParamArray keyValuePairs() As Variant) As Scripting.Dictionary
    
    ' @author: Robert Todar <robert@roberttodar.com>
    ' @ref: MicroSoft Scripting Runtime
    ' @example: ToDictionary("Name", "Robert", "Age", 30) '--> { "Name": "Robert, "Age": 30 }
    
    ' Check to see that key/value pairs passed in (an even number).
    If ArrayLength(CVar(keyValuePairs)) Mod 2 <> 0 Then
        Err.Raise 5, "ToDictionary", "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
    End If
    
    ' Add key values to the return Dictionary.
    Set ToDictionary = New Scripting.Dictionary
    Dim Index As Long
    For Index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        ToDictionary.Add keyValuePairs(Index), keyValuePairs(Index + 1)
    Next Index
    
End Function

' Returns the number of elements in an Array. (Notice how I use abstraction, this should be in its own library)
Private Function ArrayLength(ByRef Source As Variant) As Long
    ArrayLength = UBound(Source) - LBound(Source) + 1
End Function
  ```

----

## Array Functions

  ```vb
  'EXAMPLES OF VARIOUS FUNCTIONS
  Private Sub ArrayFunctionExamples()
      
      Dim A As Variant
      
      'SINGLE DIM FUNCTIONS TO MANIPULATE
      ArrayPush A, "Banana", "Apple", "Carrot" '--> Banana,Apple,Carrot
      ArrayPop A                               '--> Banana,Apple --> returns Carrot
      ArrayUnShift A, "Mango", "Orange"        '--> Mango,Orange,Banana,Apple
      ArrayShift A                             '--> Orange,Banana,Apple
      ArraySplice A, 2, 0, "Coffee"            '--> Orange,Banana,Coffee,Apple
      ArraySplice A, 0, 1, "Mango", "Coffee"   '--> Mango,Coffee,Banana,Coffee,Apple
      ArrayRemoveDuplicates A                  '--> Mango,Coffee,Banana,Apple
      ArraySort A                              '--> Apple,Banana,Coffee,Mango
      ArrayReverse A                           '--> Mango,Coffee,Banana,Apple
      
      'ARRAY PROPERTIES
      ArrayLength A                            '--> 4
      ArrayIndexOf A, "Coffee"                 '--> 1
      ArrayIncludes A, "Banana"                '--> True
      ArrayContains A, Array("Test", "Banana") '--> True
      ArrayContainsEmpties A                   '--> False
      ArrayDimensionLength A                   '--> 1 (single dim array)
      IsArrayEmpty A                           '--> False
      
      'CAN FLATTEN JAGGED ARRAY WITH SPREAD FORMULA
      A = Array(1, 2, 3, Array(4, 5, 6, Array(7, 8, 9))) 'COULD ALSO SPREAD DICTIONAIRES AND COLLECTIONS AS WELL
      A = ArraySpread(A)                       '--> 1,2,3,4,5,6,7,8,9
      
      'MATH EXAMPLES
      ArraySum A                               '--> 45
      ArrayAverage A                           '--> 5
      
      'FILTER USE'S REGEX PATTERN
      A = Array("Banana", "Coffee", "Apple", "Carrot", "Canolope")
      A = ArrayFilter(A, "^Ca|^Ap")
      
      'ARRAY TO STRING WORKS WITH BOTH SINGLE AND DOUBLE DIM ARRAYS!
      Debug.Print ArrayToString(A)
      
  End Sub
  ```
  **[⬆ back to top](#table-of-contents)**

## String Functions

  ```vb
  Private Sub StringFunctionExamples()
      
      StringSimilarity "Test", "Tester"        '->  66.6666666666667
      LevenshteinDistance "Test", "Tester"     '->  2
                                                        
      Truncate "This is a long sentence", 10                '-> "This is..."
      StringBetween "Robert Paul Todar", "Robert", "Todar"  '-> "Paul"
      StringProperLength "1001", 6, "0", True               '-> "100100"
      
      'Inject is a copy of StringInterpolation
      Inject "${0}\n\t${1}", "First", "Tab and Second" '-> First
                                                        '->   Tab and Second
      
      'Here is an example using a dictionary!
      Dim Person As New Scripting.Dictionary
      Person("Name") = "Robert"
      Person("Age") = 30
      
      'REMEMBER, DICTIONARY KEYS ARE CASE SENSITIVE!
      Debug.Print Inject("Hello,\nMy name is ${Name} and I am ${Age}!", Person)
          '-> Hello,
          '-> My name is Robert and I am 30!
  End Sub
  ```
  **[⬆ back to top](#table-of-contents)**
