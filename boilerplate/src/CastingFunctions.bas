Attribute VB_Name = "CastingFunctions"
'**/
' * Functions to help structure data in a variety of ways.
' */
Option Explicit

' Jagged array.. Just sticking with a consistent naming convention with other functions
Public Function ToArrayOfArrays(ByRef sourceArray As Variant) As Variant
    Dim temp As Variant
    ReDim temp(LBound(sourceArray) To UBound(sourceArray))
    
    Dim rowIndex As Long
    For rowIndex = LBound(sourceArray) To UBound(sourceArray)
        Dim RowList As Variant
        ReDim RowList(LBound(sourceArray, 2) To UBound(sourceArray, 2))
        
        Dim colIndex As Long
        For colIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            RowList(colIndex) = sourceArray(rowIndex, colIndex)
        Next colIndex
        
        temp(rowIndex) = RowList
    Next rowIndex
    
    ToArrayOfArrays = temp
End Function

' Be carefull with this one. Not good for large amount of records.
Public Function ToArrayOfDictionarys(ByRef sourceArray As Variant) As Variant
    ' EXTRACT FIRST ROW INDEX, EVERY OBJECT SHARES THE SAME HEADINGS
    Dim FirstRow As Long
    FirstRow = LBound(sourceArray)
    
    ' RESIZE ARRAY TO STORE ALL THE ROW OBJECTS
    Dim RowArray As Variant
    ReDim RowArray(FirstRow To UBound(sourceArray) - 1)
    
    Dim rowIndex As Long
    For rowIndex = LBound(sourceArray) + 1 To UBound(sourceArray)
        ' ADD ROW VALUES TO DICTIONARY
        Dim RowObject As Scripting.Dictionary
        Set RowObject = New Scripting.Dictionary
        
        Dim colIndex As Long
        For colIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            RowObject.Add sourceArray(FirstRow, colIndex), sourceArray(rowIndex, colIndex)
        Next colIndex
        
        ' ADD DICTIONARY TO ARRAY
        Set RowArray(rowIndex - 1) = RowObject
        Set RowObject = Nothing
    Next rowIndex
    
    ' CONVERT ARRAY TO JSON STRING
    ToArrayOfDictionarys = RowArray
End Function

' This one is safe. Not as fast as a two dim array.
Public Function ToArrayOfCollections(ByRef sourceArray As Variant) As Variant
    ' EXTRACT FIRST ROW INDEX, EVERY OBJECT SHARES THE SAME HEADINGS
    Dim FirstRow As Long
    FirstRow = LBound(sourceArray)
    
    ' RESIZE ARRAY TO STORE ALL THE ROW OBJECTS (MINUS THE TOP HEADER ROW)
    Dim RowArray As Variant
    ReDim RowArray(FirstRow To UBound(sourceArray) - 1)
    
    Dim rowIndex As Long
    For rowIndex = LBound(sourceArray) + 1 To UBound(sourceArray)
        ' ADD ROW VALUES TO Collection
        Dim RowObject As Collection
        Set RowObject = New Collection
        
        Dim colIndex As Long
        For colIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            RowObject.Add sourceArray(rowIndex, colIndex), sourceArray(FirstRow, colIndex)
        Next colIndex
        
        ' ADD Collection TO ARRAY
        Set RowArray(rowIndex - 1) = RowObject
        Set RowObject = Nothing
    Next rowIndex
    
    ' CONVERT ARRAY TO JSON STRING
    ToArrayOfCollections = RowArray
End Function

' Safe as well. Not as fast as a two dim array. But works nicely with For Each Loops
Public Function ToCollectionOfCollections(ByRef sourceArray As Variant) As Collection
    Set ToCollectionOfCollections = New Collection
    
    ' EXTRACT FIRST ROW INDEX, EVERY OBJECT SHARES THE SAME HEADINGS
    Dim FirstRow As Long
    FirstRow = LBound(sourceArray)
    
    Dim rowIndex As Long
    For rowIndex = LBound(sourceArray) + 1 To UBound(sourceArray)
        ' ADD ROW VALUES TO Collection
        Dim RowObject As Collection
        Set RowObject = New Collection
        
        Dim colIndex As Long
        For colIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            RowObject.Add sourceArray(rowIndex, colIndex), sourceArray(FirstRow, colIndex)
        Next colIndex
        
        ' ADD Collection TO ARRAY
        ToCollectionOfCollections.Add RowObject
        Set RowObject = Nothing
    Next rowIndex
End Function

'/**
' * A simple Dictionary Factory.
' * @author: Robert Todar <robert@roberttodar.com>
' * @ref: MicroSoft Scripting Runtime
' * @example: ToDictionary("Name", "Robert", "Age", 30) '--> { "Name": "Robert, "Age": 30 }
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
' * A simple Collection Factory.
' * @author: Robert Todar <robert@roberttodar.com>
' * @example: ToCollection("Name", "Robert", "Age", 30) '--> { "Name": "Robert, "Age": 30 }
' */
Public Function ToCollection(ParamArray keyValuePairs() As Variant) As Collection
    ' Check to see that key/value pairs passed in (an even number).
    If arrayLength(CVar(keyValuePairs)) Mod 2 <> 0 Then
        Err.Raise 5, "ToCollection", "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
    End If
    
    ' Add key values to the return Dictionary. ()
    Set ToCollection = New Collection
    Dim index As Long
    For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        ' Collections first take value then the key
        ToCollection.Add keyValuePairs(index + 1), keyValuePairs(index)
    Next index
End Function

Private Function IsArrayDimension(ByVal source As Variant, ByVal dimension As Long) As Boolean
    If IsArray(source) Then
        IsArrayDimension = (dimension = ArrayDimensionLength(source))
    End If
End Function

' RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Private Function ArrayDimensionLength(sourceArray As Variant) As Long
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
    ArrayDimensionLength = iterator - 1
End Function

' Returns the number of elements in an Array. (Notice how I use abstraction, this should be in its own library)
Private Function arrayLength(ByRef source As Variant) As Long
  arrayLength = UBound(source) - LBound(source) + 1
End Function
