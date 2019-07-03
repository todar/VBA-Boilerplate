Attribute VB_Name = "CastingFunctions"
Option Explicit

' Testing Data
Private Property Get sampleData() As Variant
    Dim Data(1 To 3, 1 To 3) As Variant
    Data(1, 1) = "Id"
    Data(1, 2) = "Name"
    Data(1, 3) = "Age"
    
    Data(2, 1) = 1
    Data(2, 2) = "Robert"
    Data(2, 3) = 31
    
    Data(3, 1) = 2
    Data(3, 2) = "Mark"
    Data(3, 3) = 35
    
    sampleData = Data
End Property

' Test each function
Private Sub testFunctions()
    
    Dim FSO As New FileSystemObject
    Debug.Print ToString(FSO), vbNewLine
    
    Debug.Print ToString(sampleData), vbNewLine
    
    Dim jagged As Variant
    jagged = ToArrayOfArrays(sampleData)
    Debug.Print ToString(jagged)
    
    Dim jaggedDictionary As Variant
    jaggedDictionary = ToArrayOfDictionarys(sampleData)
    Debug.Print ToString(jaggedDictionary)
    
    Dim jaggedCollection As Variant
    jaggedCollection = ToArrayOfCollections(sampleData)
    Debug.Print ToString(jaggedCollection)
    
    Dim listOfCollections As Collection
    Set listOfCollections = ToCollectionOfCollections(sampleData)
    Debug.Print ToString(listOfCollections)
   
    
    Dim person1 As Collection
    Set person1 = ToCollection("Name", "Robert", "Age", 31)
    Debug.Print ToString(person1)
    
    Dim person2 As Scripting.Dictionary
    Set person2 = ToDictionary("Name", "Mark", "Age", 35)
    Debug.Print ToString(person2)
    
End Sub

' Created this and helper functions to easily read different containers.
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
            ToString = "{Object: " & TypeName(source) & "}"
            
        Case TypeName(source) = "String"
            ToString = """" & source & """"
        
        Case Else 'IsNumeric(source), TypeName(source) = "Boolean"
            ToString = source
            
    End Select
    
End Function

Private Function AddLineIfNeeded(ByVal source As Variant) As String
        
        If TypeName(source) = "Dictionary" _
        Or TypeName(source) = "Collection" _
        Or IsArrayDimension(source, 1) _
        Or IsArrayDimension(source, 2) Then
            
            AddLineIfNeeded = vbNewLine & "  "
        End If
            
End Function

Private Function toStingDictionary(ByVal source As Scripting.Dictionary, ByVal delimiter As String) As String
    toStingDictionary = "{"
    Dim key As Variant
    For Each key In source.Keys
        toStingDictionary = toStingDictionary & AddLineIfNeeded(source(key)) & """" & key & """" & ": " & ToString(source(key)) & delimiter
    Next key
    
    toStingDictionary = Left(toStingDictionary, Len(toStingDictionary) - Len(delimiter)) & IIf(InStr(toStingDictionary, vbNewLine), vbNewLine, "") & "}"
End Function

Private Function toStingCollection(ByVal source As Collection, ByVal delimiter As String) As String
    toStingCollection = "{"
    
    Dim item As Variant
    For Each item In source
        toStingCollection = toStingCollection & AddLineIfNeeded(item) & ToString(item) & delimiter
    Next item
    toStingCollection = Left(toStingCollection, Len(toStingCollection) - Len(delimiter)) & IIf(InStr(toStingCollection, vbNewLine), vbNewLine, "") & "}"
End Function

Private Function toStingSingleDimArray(ByVal source As Variant, ByVal delimiter As String) As String
    
    toStingSingleDimArray = "["
    Dim index As Long
    For index = LBound(source) To UBound(source)
        toStingSingleDimArray = toStingSingleDimArray & AddLineIfNeeded(source(index)) & ToString(source(index)) & IIf(index < UBound(source), delimiter, "")
    Next index
    toStingSingleDimArray = toStingSingleDimArray & IIf(InStr(toStingSingleDimArray, vbNewLine), vbNewLine, "") & "]"
    
End Function

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


' Jagged array.. Just sticking with a consistent naming convention with other functions
Public Function ToArrayOfArrays(ByRef SourceArray As Variant) As Variant
    
    Dim temp As Variant
    ReDim temp(LBound(SourceArray) To UBound(SourceArray))
    
    Dim rowIndex As Long
    For rowIndex = LBound(SourceArray) To UBound(SourceArray)
        
        Dim RowList As Variant
        ReDim RowList(LBound(SourceArray, 2) To UBound(SourceArray, 2))
        
        Dim colIndex As Long
        For colIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
            RowList(colIndex) = SourceArray(rowIndex, colIndex)
        Next colIndex
        
        temp(rowIndex) = RowList
        
    Next rowIndex
    
    ToArrayOfArrays = temp
    
End Function

' Be carefull with this one. Not good for large amount of records.
Public Function ToArrayOfDictionarys(ByRef SourceArray As Variant) As Variant
    
    'EXTRACT FIRST ROW INDEX, EVERY OBJECT SHARES THE SAME HEADINGS
    Dim FirstRow As Long
    FirstRow = LBound(SourceArray)
    
    'RESIZE ARRAY TO STORE ALL THE ROW OBJECTS
    Dim RowArray As Variant
    ReDim RowArray(FirstRow To UBound(SourceArray) - 1)
    
    Dim rowIndex As Long
    For rowIndex = LBound(SourceArray) + 1 To UBound(SourceArray)
        
        'ADD ROW VALUES TO DICTIONARY
        Dim RowObject As Scripting.Dictionary
        Set RowObject = New Scripting.Dictionary
        
        Dim colIndex As Long
        For colIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
            RowObject.Add SourceArray(FirstRow, colIndex), SourceArray(rowIndex, colIndex)
        Next colIndex
        
        'ADD DICTIONARY TO ARRAY
        Set RowArray(rowIndex - 1) = RowObject
        Set RowObject = Nothing
        
    Next rowIndex
    
    'CONVERT ARRAY TO JSON STRING
    ToArrayOfDictionarys = RowArray
    
End Function

' This one is safe. Not as fast as a two dim array.
Public Function ToArrayOfCollections(ByRef SourceArray As Variant) As Variant
    
    'EXTRACT FIRST ROW INDEX, EVERY OBJECT SHARES THE SAME HEADINGS
    Dim FirstRow As Long
    FirstRow = LBound(SourceArray)
    
    'RESIZE ARRAY TO STORE ALL THE ROW OBJECTS (MINUS THE TOP HEADER ROW)
    Dim RowArray As Variant
    ReDim RowArray(FirstRow To UBound(SourceArray) - 1)
    
    Dim rowIndex As Long
    For rowIndex = LBound(SourceArray) + 1 To UBound(SourceArray)
        
        'ADD ROW VALUES TO Collection
        Dim RowObject As Collection
        Set RowObject = New Collection
        
        Dim colIndex As Long
        For colIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
            RowObject.Add SourceArray(rowIndex, colIndex), SourceArray(FirstRow, colIndex)
        Next colIndex
        
        'ADD Collection TO ARRAY
        Set RowArray(rowIndex - 1) = RowObject
        Set RowObject = Nothing
        
    Next rowIndex
    
    'CONVERT ARRAY TO JSON STRING
    ToArrayOfCollections = RowArray
    
End Function

' Safe as well. Not as fast as a two dim array. But works nicely with For Each Loops
Public Function ToCollectionOfCollections(ByRef SourceArray As Variant) As Collection
    
    Set ToCollectionOfCollections = New Collection
    
    'EXTRACT FIRST ROW INDEX, EVERY OBJECT SHARES THE SAME HEADINGS
    Dim FirstRow As Long
    FirstRow = LBound(SourceArray)
    
    Dim rowIndex As Long
    For rowIndex = LBound(SourceArray) + 1 To UBound(SourceArray)
        
        'ADD ROW VALUES TO Collection
        Dim RowObject As Collection
        Set RowObject = New Collection
        
        Dim colIndex As Long
        For colIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
            RowObject.Add SourceArray(rowIndex, colIndex), SourceArray(FirstRow, colIndex)
        Next colIndex
        
        'ADD Collection TO ARRAY
        ToCollectionOfCollections.Add RowObject
        Set RowObject = Nothing
        
    Next rowIndex
    
End Function

' A simple Dictionary Factory.
Public Function ToDictionary(ParamArray keyValuePairs() As Variant) As Scripting.Dictionary
  
  ' @author: Robert Todar <robert@roberttodar.com>
  ' @ref: MicroSoft Scripting Runtime
  ' @example: ToDictionary("Name", "Robert", "Age", 30) '--> { "Name": "Robert, "Age": 30 }
  
  ' Check to see that key/value pairs passed in (an even number).
  If ArrayLength(CVar(keyValuePairs)) Mod 2 <> 0 Then
      Err.Raise 5, "ToDictionary", "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
  End If
  
  ' Add key values to the return Dictionary.
  Set ToDictionary = New Scripting.Dictionary
  Dim index As Long
  For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
      ToDictionary.Add keyValuePairs(index), keyValuePairs(index + 1)
  Next index
  
End Function

' A simple Collection Factory.
Public Function ToCollection(ParamArray keyValuePairs() As Variant) As Collection
  
    ' @author: Robert Todar <robert@roberttodar.com>
    ' @example: ToCollection("Name", "Robert", "Age", 30) '--> { "Name": "Robert, "Age": 30 }
    
    ' Check to see that key/value pairs passed in (an even number).
    If ArrayLength(CVar(keyValuePairs)) Mod 2 <> 0 Then
        Err.Raise 5, "ToCollection", "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
    End If
    
    ' Add key values to the return Dictionary. ()
    Set ToCollection = New Collection
    Dim index As Long
    For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        'Collections first take value then the key
        ToCollection.Add keyValuePairs(index + 1), keyValuePairs(index)
    Next index
  
End Function

Private Function IsArrayDimension(ByVal source As Variant, ByVal dimension As Long) As Boolean
    If IsArray(source) Then
        IsArrayDimension = (dimension = ArrayDimensionLength(source))
    End If
End Function


'RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Private Function ArrayDimensionLength(SourceArray As Variant) As Long
    
    'Run loop until error. Remove one and it gives the array dimension =)
    On Error GoTo Catch
    Do
        Dim iterator As Long
        iterator = iterator + 1
        
        Dim test As Long
        test = UBound(SourceArray, iterator)
    Loop
Catch:
    On Error GoTo 0
    ArrayDimensionLength = iterator - 1

End Function

' Returns the number of elements in an Array. (Notice how I use abstraction, this should be in its own library)
Private Function ArrayLength(ByRef source As Variant) As Long
  ArrayLength = UBound(source) - LBound(source) + 1
End Function
