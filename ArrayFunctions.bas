Attribute VB_Name = "arrayFunctions"
Option Explicit
Option Compare Text
Option Private Module
Option Base 0


'@AUTHOR: ROBERT TODAR

'DEPENDENCIES
' -

'PUBLIC FUNCTIONS
' - arrayGetColumnNumber
' - ArrayQuery
' - ArrayFromRecordset
' - ArrayToString
' -
' -
' -

'PRIVATE METHODS/FUNCTIONS (IN DEVELOPMENT)
' - ArrayToTextFile -NEED TO ADJUST BEFORE MAKING PUBLIC
' - arrayPush
' - isSingleDimension
' - dimensionLength
' - asign
' -
' -

'NOTES:
' - I'VE CREATE AN ARRAY CLASS MODULE THAT DOES MANY OF THESE FUNCTIONS, DECIDED TO ALSO
' - CREATE FUNCTIONS AWAY FROM CLASS MODULE OBJECT TO MAKE THEM WORK WITH ANY ARRAY.

'TODO:
' - ADD MORE FUNCTIONS FROM ARRAYOBJECT CLASS MODULE
' - FINISH PRIVATE METHODS AND MAKE THEM PUBLIC.
' - REMOVE THE NEED TO HAVE A ADODB REFERENCE

'EXAMPLES:
' -

'******************************************************************************************
' TESTING
'******************************************************************************************

'USED TO GET SAMPLE DATA FROM ACTIVESHEET
Private Property Get TestData() As Variant
    TestData = Range("A1").CurrentRegion
End Property

'USED FOR TESTING NEW AND MODIFIED FUNCTIONS
Public Sub TestingArrayFunctions()

    Dim arr As Variant
    Dim sql As String
    
    sql = "SELECT * FROM []"
    
    arr = ArrayQuery(TestData, sql)
    Debug.Print ArrayToString(arr)
    
End Sub


'******************************************************************************************
' PUBLIC FUNCTIONS
'******************************************************************************************

'LOOKS FOR VALUE IN FIRST ROW OF A TWO DIMENSIONAL ARRAY, RETURNS IT'S COL INDEX
Public Function ArrayGetColumnNumber(arr As Variant, HeadingValue As String) As Integer
    
    Dim columnIndex As Integer
    For columnIndex = LBound(arr, 2) To UBound(arr, 2)
        If arr(LBound(arr, 1), columnIndex) = HeadingValue Then
            ArrayGetColumnNumber = columnIndex
            Exit Function
        End If
    Next columnIndex
    
    'RETURN NEGATIVE IF NOT FOUND
    ArrayGetColumnNumber = -1
    
End Function

'CREATES TEMP TEXT FILE AND SAVES ARRAY VALUES IN A CSV FORMAT, THEN QUERIES AND RETURNS ARRAY.
'
'@AUTHOR ROBERT TODAR
'@USES ArrayToTextFile
'@USES ArrayFromRecordset
'@RETURNS 2D ARRAY || EMPTY (IF NO RECORDS)
'@PARAM {ARR} MUST BE A TWO DIMENSIONAL ARRAY, SETUP AS IF IT WERE A TABLE.
'@PARAM {SQL} ADO SQL STATEMENT FOR A TEXT FILE. MUST INCLUDE 'FROM []'
'@PARAM {IncludeHeaders} BOOLEAN TO RETURN HEADERS WITH DATA OR NOT
'@EXAMPLE SQL = "SELECT * FROM [] WHERE [FIRSTNAME] = 'ROBERT'"
Public Function ArrayQuery(arr As Variant, sql As String, Optional IncludeHeaders As Boolean = True) As Variant
    
    'CREATE TEMP FOLDER AND FILE NAMES
    Const fileName As String = "temp.txt"
    Dim filePath As String
    filePath = Environ("temp")
    
    'UPDATE SQL WITH TEMP FILE NAME
    sql = Replace(sql, "FROM []", "FROM [" & fileName & "]")
    
    'SEND ARRAY TO TEMP TEXTFILE IN CSV FORMAT
    ArrayToTextFile arr, filePath & "\" & fileName, ","
    
    'CREATE CONNECTION TO TEMP FILE - CONNECTION IS SET TO COMMA SEPERATED FORMAT
    Dim cnn As Object
    Set cnn = CreateObject("ADODB.Connection")
    cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnn.ConnectionString = "Data Source=" & filePath & ";" & "Extended Properties=""text;HDR=Yes;FMT=Delimited;"""
    cnn.Open
    
    'CREATE RECORDSET AND QUERY ON PASSED IN SQL (QUERIES THE TEMP TEXT FILE)
    Dim rs As Object
    Set rs = CreateObject("ADODB.RecordSet")
    With rs
        .ActiveConnection = cnn
        .Open sql
        
        'GET AN ARRAY FROM THE RECORDSET
         ArrayQuery = ArrayFromRecordset(rs, IncludeHeaders)
        .Close
    End With
    
    'CLOSE CONNECTION AND KILL TEMP FILE
    cnn.Close
    Kill filePath & "\" & fileName
    
End Function

'RETURNS A 2D ARRAY FROM A RECORDSET, OPTIONALLY INCLUDING HEADERS, AND IT TRANSPOSES TO KEEP
'ORIGINAL OPTION BASE. (TRANSPOSE WILL SET IT TO BASE 1 AUTOMATICALLY.)
'
'@AUTHOR ROBERT TODAR
Public Function ArrayFromRecordset(rs As Recordset, Optional IncludeHeaders As Boolean = True) As Variant
    
    '@NOTE: -Int(IncludeHeaders) RETURNS A BOOLEAN TO AN INT (0 OR 1)
    Dim HeadingIncrement As Integer
    HeadingIncrement = -Int(IncludeHeaders)
    
    'CHECK TO MAKE SURE THERE ARE RECORDS TO PULL FROM
    If rs.BOF Or rs.EOF Then
        Exit Function
    End If
    
    'STORE RS DATA
    Dim rsData As Variant
    rsData = rs.GetRows
    
    'REDIM TEMP TO ALLOW FOR HEADINGS AS WELL AS DATA
    Dim Temp As Variant
    ReDim Temp(LBound(rsData, 2) To UBound(rsData, 2) + HeadingIncrement, LBound(rsData, 1) To UBound(rsData, 1))
        
    If IncludeHeaders = True Then
        'GET HEADERS
        Dim headerIndex As Long
        For headerIndex = 0 To rs.fields.Count - 1
            Temp(LBound(Temp, 1), headerIndex) = rs.fields(headerIndex).name
        Next headerIndex
    End If
    
    'GET DATA
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = LBound(Temp, 1) + HeadingIncrement To UBound(Temp, 1)
        
        For colIndex = LBound(Temp, 2) To UBound(Temp, 2)
            Temp(rowIndex, colIndex) = rsData(colIndex, rowIndex - HeadingIncrement)
        Next colIndex
        
    Next rowIndex
    
    'RETURN
    ArrayFromRecordset = Temp
    
End Function

'RETURNS A STRING FROM A 2 DIM ARRAY, SPERATED BY OPTIONAL DELIMITER AND VBNEWLINE FOR EACH ROW
'
'@AUTHOR ROBERT TODAR
Public Function ArrayToString(SourceArray As Variant, Optional Delimiter As String = ",") As String
    
    Dim Temp As String
    
    Select Case ArrayDimensionLength(SourceArray)
        'SINGLE DIMENTIONAL ARRAY
        Case 1
            Temp = Join(SourceArray, Delimiter)
        
        '2 DIMENSIONAL ARRAY
        Case 2
            Dim rowIndex As Long
            Dim colIndex As Long
            
            'LOOP EACH ROW IN MULTI ARRAY
            For rowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                
                'LOOP EACH COLUMN ADDING VALUE TO STRING
                For colIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                    Temp = Temp & SourceArray(rowIndex, colIndex)
                    If colIndex <> UBound(SourceArray, 2) Then Temp = Temp & Delimiter
                Next colIndex
                
                'ADD NEWLINE FOR THE NEXT ROW (MINUS LAST ROW)
                If rowIndex <> UBound(SourceArray, 1) Then Temp = Temp & vbNewLine
        
            Next rowIndex
    End Select
    
    ArrayToString = Temp
    
End Function

'RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    
    Dim i As Integer
    Dim test As Long

    On Error GoTo Catch
    Do
        i = i + 1
        test = UBound(SourceArray, i)
    Loop
    
Catch:
    ArrayDimensionLength = i - 1

End Function

'SENDS AN ARRAY TO A TEXTFILE
Public Sub ArrayToTextFile(arr As Variant, filePath As String, Optional delimeter As String = ",")
    
    Dim fso As Object
    Dim ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 2, True) '2=WRITEABLE
   
    ts.Write ArrayToString(arr, delimeter)
    Set ts = Nothing

End Sub


'******************************************************************************************
' PRIVATE FUNCTIONS - BEING DEVELOPED STILL
'******************************************************************************************

Private Sub TestArrayPush()
    
    Dim arr As Variant
    
    ArrayPush arr, 1, 2, 3, 4, 5
    ArrayPush arr, 6, 7, 8
    
    Debug.Print ArrayToString(arr)

End Sub

' - ADDS A NEW ELEMENT(S) TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
Private Function ArrayPush(SourceArray As Variant, ParamArray Element() As Variant) As Long

    Dim index As Long
    Dim firstEmptyBound As Long
    Dim OptionBase As Integer
    
    OptionBase = 0

    '@TODO: FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. UPDATE TO PUSH TO MULTI DIM ARRAYS?
    If ArrayDimensionLength(SourceArray) > 1 Then
        ArrayPush = -1
        Exit Function
    End If
    
    'REDIM IF EMPTY, OR INCREASE ARRAY IF NOT EMPTY
    If IsArrayEmpty(SourceArray) Then
    
        ReDim SourceArray(OptionBase To UBound(Element, 1) + OptionBase)
        firstEmptyBound = LBound(SourceArray, 1)
        
    Else
        firstEmptyBound = UBound(SourceArray, 1) + 1
        ReDim Preserve SourceArray(UBound(SourceArray, 1) + UBound(Element, 1) + 1)
        
    End If
    
    'LOOP EACH NEW ELEMENT
    For index = LBound(Element, 1) To UBound(Element, 1)
        
        'ADD ELEMENT TO THE END OF THE ARRAY
        asign SourceArray(firstEmptyBound), Element(index)
        
        'INCREMENT TO THE NEXT firstEmptyBound
        firstEmptyBound = firstEmptyBound + 1
        
    Next index
    
    'RETURN NEW ARRAY LENGTH
    ArrayPush = UBound(SourceArray, 1) + 1

End Function

' - QUICK TOOL TO EITHER SET OR LET DEPENDING ON IF ELEMENT IS AN OBJECT
Private Function asign(variable As Variant, value As Variant)

    If IsObject(value) Then
        Set variable = value
    Else
        Let variable = value
    End If
    
End Function


