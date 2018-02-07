
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

'===================================================
' FORM CONTROL - SETS FORM TO CENTER OF APPLICATION
' EXAMPLE: Centerform me
'===================================================
Public Sub CenterForm(Form)

    With Form
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With

End Sub

'============================================================
' REMOVES LAST ELEMENT IN 1 DIM ARRAY, RETURNS POPPED ELEMENT
'============================================================
Public Function ArrayPop(ByRef arr As Variant) As Variant
    
    If IsArrayEmpty(arr) Then
        ReDim arr(0)
        arr(0) = ""
    Else
        ArrayPop = arr(UBound(arr, 1))
        ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1) - 1)
    End If
    
End Function

'=======================================================================
' ADDS ELEMENT TO THE END OF A 1 DIM ARRAY - RETURNS NEW LENGHT OF ARRAY
'=======================================================================
Public Function ArrayPush(ByRef arr As Variant, Value As Variant) As Long
    
    If IsArrayEmpty(arr) Then
        ReDim arr(0)
        arr(0) = Value
    Else
        ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1) + 1)
        arr(UBound(arr, 1)) = Value
    End If
    
    ArrayPush = UBound(arr, 1)
    
End Function

'============================================================
' REMOVES ELEMENT FROM ARRAY - RETURNS REMOVED ELEMENT
'============================================================
Public Function ArrayShift(ByRef arr As Variant, Optional ElementNumber As Long = 0) As Variant
    
    Dim I As Long
    
    If IsArrayEmpty(arr) Then
        ReDim arr(0)
        arr(0) = ""
    Else
        ArrayShift = arr(ElementNumber)
        
        For I = ElementNumber To UBound(arr) - 1
            arr(I) = arr(I + 1)
        Next I
        
        ReDim Preserve arr(UBound(arr, 1) - 1)
        
    End If
    
End Function

'=============================================================
' ADDS ELEMENT TO THE START OF AN ARRAY - RETURNS ARRAY LENGTH
'=============================================================
Public Function ArrayUnShift(ByRef arr As Variant, Value As Variant) As Long
    
    Dim I As Long
    
    If IsArrayEmpty(arr) Then
        ReDim arr(0)
        arr(0) = ""
    Else
    
        ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1) + 1)
        
        For I = UBound(arr, 1) To LBound(arr, 1) Step -1
            
            If I > LBound(arr, 1) Then
                arr(I) = arr(I - 1)
            Else
                 arr(I) = Value
            End If
    
        Next I
  
    End If
    
    ArrayUnShift = (UBound(arr, 1) - LBound(arr, 1)) + 1
    
End Function

'===================================================
' ADDS ONLY UNIQUE ITEMS TO A COLLECTION
'===================================================
Public Function CollectionAddUnique(ByRef Target As collection, Value As String) As Boolean
    
    Dim L As Long
    
    'SEE IF COLLECTION HAS ANY VALUES
    If Target.Count = 0 Then
        Target.Add Value
        Exit Function
    End If
    
    'SEE IF VALUE EXISTS IN COLLECTION
    For L = 1 To Target.Count
        If Target(L) = Value Then
            Exit Function
        End If
    Next L
    
    'NOT IN COLLECTION, ADD VALUE TO COLLECTION
    Target.Add Value
    CollectionAddUnique = True
    
End Function

'===================================================
' TAKES A STRING AND PUTS IT INTO THE CLIPBOARD
'===================================================
Sub ClipboardSet(ByVal inText As String)
  
    Dim objClipboard As Object
    Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    objClipboard.SetText inText
    objClipboard.PutInClipboard

    Set objClipboard = Nothing
  
End Sub

'======================================================
' RETRIVES TEXT FROM CLIPBOARD - SETS VALUE TO FUNCTION
'======================================================
Function ClipboardGet() As String
  
    Dim objClipboard As Object
    Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 
    objClipboard.GetFromClipboard
    ClipboardGet = objClipboard.GetText
 
    Set objClipboard = Nothing
  
End Function

'======================================================
' INSERT A SLICER TO A TABLE
'======================================================
Public Function InsertSlicer(Ws As Worksheet, tbl As ListObject, Target As Range, Value As String)
    
    Dim Wb As Workbook
    Set Wb = Ws.Parent
    
    Wb.SlicerCaches.Add2(tbl, Value). _
    Slicers.Add Ws, , Value, Value, Target.Top, Target.Left, Target.Width, Target.Height

End Function

'===================================================
'SEARCH FOR HEADING, THEN FINDS COLUMN DATA
'===================================================
Function FindColumnData(sht As Worksheet, What As String, Optional IncludeHead As Boolean = False) As Range
    
    Set FindColumnData = _
    sht.Range("1:1").Find(What:=What, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Offset(Abs(Not IncludeHead))
    
    Set FindColumnData = _
        sht.Range(FindColumnData.Address, sht.Cells(sht.Rows.Count, FindColumnData.Column).End(xlUp))
    
End Function

'===================================================
' FIND GROUPS OF VALUES IN A COLUMN
' COLUMN MUST BE SORTED FIRST
'===================================================
Public Function FindGroup(sht As Worksheet, What As String, Rng As Range) As Range

    Dim R1 As Range
    Dim r2 As Range
    
    Set R1 = sht.Range(Rng.Address).Find(What, , xlValues, xlWhole)
    Set r2 = sht.Range(Rng.Address).Find(What, , xlValues, xlWhole, , xlPrevious)
    
    On Error GoTo Err
    Set FindGroup = sht.Range(R1.Address, r2.Address)
    
Err:
End Function

'===================================================
'SEARCH FOR HEADING IN FIRST ROW
'===================================================
Function FindHeading(sht As Worksheet, What As String) As Range
    
    Set FindHeading = _
    sht.Range("1:1").Find(What:=What, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
End Function

'===================================================
' FREEZE THE TOP ROW OF ACTIVE SHEET
'===================================================
Public Sub FreezeTopRow(Optional Bo As Boolean)
     
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    
End Sub


'===================================================
' TIMER FOR SENDKEYS/APPLICATION INTERACTION
'===================================================
Public Function Pause(Mil_Second As Single) As Single

    Dim X As Single
        
    X = Timer
    While Timer - X < Mil_Second
    Wend

End Function

'=======================================================
' CHECKS TO SEE IF FILE EXISTS, THEN OPENS IT IF IT DOES
'=======================================================
Public Function OpenAnyFile(FileToOpen As String) As Boolean
    
    'CALLS FUNCTION TO SEE IF THE FILE EXISTS, RETURNS TRUE IF IT DOES
    If FileExists(FileToOpen) Then
        OpenAnyFile = True
        
        'API FUNCTION FOR OPENING FILES
        Call ShellExecute(0, "Open", FileToOpen & vbNullString, _
        vbNullString, vbNullString, 1)
    End If
    
End Function

'=====================================================
' ENSURES LENGHT OF STRING IS AS LONG AS IT SHOULD BE
'=====================================================
Public Function ProperLength(Value As String, Length As Integer, FillValue As String, Optional AfterString As Boolean = True) As String
    
    If Len(Value) >= Length Then
        Value = Left(Value, Length)
    Else
    
        Do While Len(Value) < Length
            If AfterString = True Then
                Value = Value & FillValue
            ElseIf AfterString = False Then
                Value = FillValue & Value
            End If
        Loop
    End If
    
    ProperLength = Value
    
End Function

'============================================================
' Find Value between two words in a string
'============================================================
Public Function StringBetween(Main As String, Between1 As String, Between2 As String) As String

    Dim I As Integer
    Dim i2 As Integer
    
    I = InStr(Main, Between1)
    I = I + Len(Between1)
    
    If Between2 = "" Then
        i2 = Len(Main) + 1
    Else
        i2 = InStr(I, Main, Between2)
    End If
    
    StringBetween = Trim(Mid(Main, I, i2 - I))
    
End Function

'================================================
' SELECTS TEXT IN A TEXTBOX
'================================================
Public Function TextboxWordSelect(Textbox As MSForms.Textbox) As Boolean
    
    Textbox.SelStart = 0
    Textbox.SelLength = Len(Textbox.Text)
    
End Function

