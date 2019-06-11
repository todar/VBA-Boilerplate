<p align="center">
    <img width="350px" alt="function meme" src="https://i.pinimg.com/736x/2e/e7/b3/2ee7b37349f798c3460e244143bdd0bc--math-puns-math-humor.jpg">
    <h1 align="center">VBA Function Library</h1>
</p>

You've found my VBA Libray GitHub repository, which contains functions as well as a style guide to help make programming in VBA easier.

> This repository is currently under construction, but will be intended to be a place to help make VBA more open source.


## Table of Contents

  1. [Style Guide](#style-guide)
     1. [Naming Conventions](#naming-conventions)
     2. [Comments](#comments)
     3. [Design](#design)
  2. [Array Examples](#array-functions)
  3. [String Examples](#string-functions)

## Style Guide

  Below is an example function of how I entend all my code to be written. 

  > If you'd like to contribute please try to format similar to this for consistency.

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

## Naming Conventions

  <a name="single--letter--names"></a><a name="1.1"></a>
  - [1.1](#single--letter--names) Avoid single letter names. Be descriptive with your naming.
    ```vb
    ' bad
    Function Q ()
      Dim i as Long
      ' ...
    End Function

    ' good
    Function Query ()
      Dim RecordIndex as Long
      ' ...
    End Function
    ```

  <a name="pascal--case"></a><a name="1.2"></a>
  - [1.2](#pascal--case) Use PascalCase for all your naming.
    ```vb
    ' good
    Function GreetUser ()
      ' ...
    End Function
    ```

  <a name="underscore--case"></a><a name="1.3"></a>
  - [1.3](#underscore--case) Do not use underscore case.
    
    > Why? VBA uses underscores for pointing out events and implementation. Adding underscores not only makes it confusing, but can also lead to bugs.
    ```vb
    ' bad
    Dim First_Name as String

    ' good
    Dim FirstName as String
    ```
  **[⬆ back to top](#table-of-contents)**

### Comments

  <a name="description-header-comment"></a><a name="2.1"></a>
  - [2.1](#description-header-comment) Above the function should be a simple description of what the function does.

  <a name="doc--comment"></a><a name="2.2"></a>
  - [2.1](#doc--comment) Just inside the function is where I will put important details. This could be author, library references, notes, Ect. I've styled this to be similar to [JSDoc documentation](https://devdocs.io/jsdoc/). 

  <a name="descriptive--comment"></a><a name="2.1"></a>
  - [2.1](#descriptive--comment) Notes should be clear and full sentences. Explain anything that doesn't immediatly make sence from the code.

  **[⬆ back to top](#table-of-contents)**


### Design

  Functions should be as small as possible designed to resusable. This means they should be very readable.

  Declarations should be made where the variables are needed. Notice `Dim Index as Long` is declared right before the loop. This makes it easier to read, debug, and refactor if need be.

  **[⬆ back to top](#table-of-contents)**

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
