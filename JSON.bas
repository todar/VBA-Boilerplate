Attribute VB_Name = "JSON"
Option Explicit
Option Compare Text
Option Private Module

'@AUTHOR: ROBERT TODAR
'@AUTHOR: omegastripes@yandex.ru - OPEN SOURCE CODE

'DEPENDENCIES
' -

'PUBLIC FUNCTIONS
' -
' -
' -
' -

'PRIVATE METHODS/FUNCTIONS
' -
' -
' -
' -

'NOTES:
' - TWO MAIN FUNCTIONS TO WORK WITH JSON FILES/STRINGS.
' - EITHER CONVERT VBA ARRAYS/DICTIONARIES INTO A JSON STRING
' - OR PARSE JSON STRING INTO VBA ARRAYS/DICTIONARIES
' -
' -

'TODO:
' - CLEAN UP NOTES
' - CLEAN UP CODE ORGANIZATION.

'EXAMPLES:
' -
' -

'******************************************************************************************
' PUBLIC METHODS
'******************************************************************************************

'TAKE A DICTIONARY OR ARRAY AND RETURN IT AS A JSON STRING
Public Function jsonStringify(json As Variant, Optional FilePath As String) As String

    jsonStringify = BeautifyJson(json)
    
    'WRITE JSON STRING TO FILEPATH IF FILEPATH IS PASSED THROUGH PARAMETERS
    If FilePath <> vbNullString Then
        If CreateFilePath(FilePath) = True Then
            writeToTextFile FilePath, jsonStringify
        End If
    End If

End Function

'PARSE JSON (READ JSON STRING AND MAKE IT INTO EITHER AN ARRAY OR DICTIONARY), AND RETURN OBJECT TYPE/ERROR
Public Function jsonParse(ByVal json As String, Optional returnVar As Variant) As Variant
    
    Dim strState As String
    Dim varItem As Variant
    Dim Fso As New FileSystemObject
    
    If Fso.FileExists(json) Then
        json = ReadTextFile(json)
    End If
    
    If json = "" Then
        Set returnVar = New Scripting.Dictionary
        Set jsonParse = returnVar
        Exit Function
    End If
    ' parse JSON string to object
    ' root element can be the object {} or the array []
    privateParseJson json, returnVar, strState
    
    If IsObject(returnVar) Then
        Set jsonParse = returnVar
    Else
        jsonParse = returnVar
    End If
    
End Function



'******************************************************************************************
' PRIVATE METHODS
'******************************************************************************************
Private Function ReadTextFile(FilePath As String) As String
    
    Dim Fso As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = Fso.OpenTextFile(FilePath, ForReading, False)
    On Error Resume Next
    ReadTextFile = ts.ReadAll
    
    Set Fso = Nothing
    Set ts = Nothing
    
End Function

Private Function writeToTextFile(FilePath As String, Value As String)
    
    Dim Fso As New FileSystemObject
    Dim ts As TextStream
    
    If CreateFilePath(FilePath) = True Then
        Set ts = Fso.OpenTextFile(FilePath, ForWriting, True)
        ts.Write Value
    End If
    
    Set Fso = Nothing
    Set ts = Nothing
    
End Function

Private Function CreateFilePath(FullPath As String) As Boolean

    Dim Fso As New FileSystemObject
    Dim i As Integer
    Dim sPath() As String
    Dim CurPath As String
    
    On Error GoTo catch
    sPath = Split(FullPath, "\")
    
    'CREATES EACH FOLDER PATH IF NEEDED
    For i = LBound(sPath, 1) To UBound(sPath, 1) - 1
        CurPath = CurPath & sPath(i) & "\"
        If Not Fso.FolderExists(CurPath) Then
            Fso.createFolder CurPath
        End If
    Next i
    
    'CREATES FILE IF NEEDED
    If Not Fso.FileExists(FullPath) Then
        Fso.CreateTextFile FullPath
    End If
    
    CreateFilePath = True
    Exit Function
catch:
    'RETURNS FALSE
End Function

'CALL TO OPEN SOURCE CODE
Private Function BeautifyJson(varJson As Variant) As String
    Dim strResult As String
    Dim lngIndent As Long
    BeautifyJson = ""
    lngIndent = 0
    BeautyTraverse BeautifyJson, lngIndent, varJson, "  ", 2
End Function


'******************************************************************************************
' PRIVATE METHODS (OPEN SOURCE)
'******************************************************************************************

'DECIDED TO GO WITH THIS ONE, RUNS MORE EFFICIENTLY THAN MY CODE.
'https://stackoverflow.com/questions/6627652/parsing-json-in-excel-vba
' Copyright (C) 2015-2017 omegastripes
' omegastripes@yandex.ru
'answered May 27 '15 at 22:45 omegastripes
Private Sub privateParseJson(ByVal strContent As String, varJson As Variant, strState As String)
    ' strContent - source JSON string
    ' varJson - created object or array to be returned as result
    ' strState - Object|Array|Error depending on processing to be returned as state
    Dim objTokens As Object
    Dim objRegEx As Object
    Dim bMatched As Boolean

    Set objTokens = CreateObject("Scripting.Dictionary")
    Set objRegEx = CreateObject("VBScript.RegExp")
    With objRegEx
        ' specification http://www.json.org/
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = """(?:\\""|[^""])*""(?=\s*(?:,|\:|\]|\}))"
        Tokenize objTokens, objRegEx, strContent, bMatched, "str"
        .Pattern = "(?:[+-])?(?:\d+\.\d*|\.\d+|\d+)e(?:[+-])?\d+(?=\s*(?:,|\]|\}))"
        Tokenize objTokens, objRegEx, strContent, bMatched, "num"
        .Pattern = "(?:[+-])?(?:\d+\.\d*|\.\d+|\d+)(?=\s*(?:,|\]|\}))"
        Tokenize objTokens, objRegEx, strContent, bMatched, "num"
        .Pattern = "\b(?:true|false|null)(?=\s*(?:,|\]|\}))"
        Tokenize objTokens, objRegEx, strContent, bMatched, "cst"
        .Pattern = "\b[A-Za-z_]\w*(?=\s*\:)" ' unspecified name without quotes
        Tokenize objTokens, objRegEx, strContent, bMatched, "nam"
        .Pattern = "\s"
        strContent = .Replace(strContent, "")
        .MultiLine = False
        Do
            bMatched = False
            .Pattern = "<\d+(?:str|nam)>\:<\d+(?:str|num|obj|arr|cst)>"
            Tokenize objTokens, objRegEx, strContent, bMatched, "prp"
            .Pattern = "\{(?:<\d+prp>(?:,<\d+prp>)*)?\}"
            Tokenize objTokens, objRegEx, strContent, bMatched, "obj"
            .Pattern = "\[(?:<\d+(?:str|num|obj|arr|cst)>(?:,<\d+(?:str|num|obj|arr|cst)>)*)?\]"
            Tokenize objTokens, objRegEx, strContent, bMatched, "arr"
        Loop While bMatched
        .Pattern = "^<\d+(?:obj|arr)>$" ' unspecified top level array
        If Not (.test(strContent) And objTokens.exists(strContent)) Then
            varJson = Null
            strState = "Error"
        Else
            Retrieve objTokens, objRegEx, strContent, varJson
            strState = IIf(IsObject(varJson), "Object", "Array")
        End If
    End With
End Sub

Private Sub Tokenize(objTokens, objRegEx, strContent, bMatched, strType)
    Dim strKey As String
    Dim strRes As String
    Dim lngCopyIndex As Long
    Dim objMatch As Object

    strRes = ""
    lngCopyIndex = 1
    With objRegEx
        For Each objMatch In .Execute(strContent)
            strKey = "<" & objTokens.Count & strType & ">"
            bMatched = True
            With objMatch
                objTokens(strKey) = .Value
                strRes = strRes & Mid(strContent, lngCopyIndex, .FirstIndex - lngCopyIndex + 1) & strKey
                lngCopyIndex = .FirstIndex + .length + 1
            End With
        Next
        strContent = strRes & Mid(strContent, lngCopyIndex, Len(strContent) - lngCopyIndex + 1)
    End With
End Sub

Private Sub Retrieve(objTokens, objRegEx, strTokenKey, varTransfer)
    Dim strContent As String
    Dim strType As String
    Dim objMatches As Object
    Dim objMatch As Object
    Dim strName As String
    Dim varValue As Variant
    Dim objArrayElts As Object

    strType = Left(Right(strTokenKey, 4), 3)
    strContent = objTokens(strTokenKey)
    With objRegEx
        .Global = True
        Select Case strType
            Case "obj"
                .Pattern = "<\d+\w{3}>"
                Set objMatches = .Execute(strContent)
                Set varTransfer = CreateObject("Scripting.Dictionary")
                For Each objMatch In objMatches
                    Retrieve objTokens, objRegEx, objMatch.Value, varTransfer
                Next
            Case "prp"
                .Pattern = "<\d+\w{3}>"
                Set objMatches = .Execute(strContent)

                Retrieve objTokens, objRegEx, objMatches(0).Value, strName
                Retrieve objTokens, objRegEx, objMatches(1).Value, varValue
                If IsObject(varValue) Then
                    Set varTransfer(strName) = varValue
                Else
                    varTransfer(strName) = varValue
                End If
            Case "arr"
                .Pattern = "<\d+\w{3}>"
                Set objMatches = .Execute(strContent)
                Set objArrayElts = CreateObject("Scripting.Dictionary")
                For Each objMatch In objMatches
                    Retrieve objTokens, objRegEx, objMatch.Value, varValue
                    If IsObject(varValue) Then
                        Set objArrayElts(objArrayElts.Count) = varValue
                    Else
                        objArrayElts(objArrayElts.Count) = varValue
                    End If
                    varTransfer = objArrayElts.items
                Next
            Case "nam"
                varTransfer = strContent
            Case "str"
                varTransfer = Mid(strContent, 2, Len(strContent) - 2)
                varTransfer = Replace(varTransfer, "\""", """")
                varTransfer = Replace(varTransfer, "\\", "\")
                varTransfer = Replace(varTransfer, "\/", "/")
                varTransfer = Replace(varTransfer, "\b", Chr(8))
                varTransfer = Replace(varTransfer, "\f", Chr(12))
                varTransfer = Replace(varTransfer, "\n", vbLf)
                varTransfer = Replace(varTransfer, "\r", vbCr)
                varTransfer = Replace(varTransfer, "\t", vbTab)
                .Global = False
                .Pattern = "\\u[0-9a-fA-F]{4}"
                Do While .test(varTransfer)
                    varTransfer = .Replace(varTransfer, ChrW(("&H" & Right(.Execute(varTransfer)(0).Value, 4)) * 1))
                Loop
            Case "num"
                varTransfer = Evaluate(strContent)
            Case "cst"
                Select Case LCase(strContent)
                    Case "true"
                        varTransfer = True
                    Case "false"
                        varTransfer = False
                    Case "null"
                        varTransfer = Null
                End Select
        End Select
    End With
End Sub

Private Sub BeautyTraverse(strResult As String, lngIndent As Long, varElement As Variant, strIndent As String, lngStep As Long)
    Dim arrKeys() As Variant
    Dim lngIndex As Long
    Dim strTemp As String

    Select Case VarType(varElement)
        Case vbObject
            If varElement.Count = 0 Then
                strResult = strResult & "{}"
            Else
                strResult = strResult & "{" & vbCrLf
                lngIndent = lngIndent + lngStep
                arrKeys = varElement.Keys
                For lngIndex = 0 To UBound(arrKeys)
                    strResult = strResult & String(lngIndent, strIndent) & """" & arrKeys(lngIndex) & """" & ": "
                    BeautyTraverse strResult, lngIndent, varElement(arrKeys(lngIndex)), strIndent, lngStep
                    If Not (lngIndex = UBound(arrKeys)) Then strResult = strResult & ","
                    strResult = strResult & vbCrLf
                Next
                lngIndent = lngIndent - lngStep
                strResult = strResult & String(lngIndent, strIndent) & "}"
            End If
        Case Is >= vbArray
            If UBound(varElement) = -1 Then
                strResult = strResult & "[]"
            Else
                strResult = strResult & "[" & vbCrLf
                lngIndent = lngIndent + lngStep
                For lngIndex = 0 To UBound(varElement)
                    strResult = strResult & String(lngIndent, strIndent)
                    BeautyTraverse strResult, lngIndent, varElement(lngIndex), strIndent, lngStep
                    If Not (lngIndex = UBound(varElement)) Then strResult = strResult & ","
                    strResult = strResult & vbCrLf
                Next
                lngIndent = lngIndent - lngStep
                strResult = strResult & String(lngIndent, strIndent) & "]"
            End If
        Case vbInteger, vbLong, vbSingle, vbDouble
            strResult = strResult & varElement
        Case vbNull
            strResult = strResult & "Null"
        Case vbBoolean
            strResult = strResult & IIf(varElement, "true", "false")
        Case Else
            strTemp = Replace(varElement, "\""", """")
            strTemp = Replace(strTemp, "\", "\\")
            strTemp = Replace(strTemp, "/", "\/")
            strTemp = Replace(strTemp, Chr(8), "\b")
            strTemp = Replace(strTemp, Chr(12), "\f")
            strTemp = Replace(strTemp, vbLf, "\n")
            strTemp = Replace(strTemp, vbCr, "\r")
            strTemp = Replace(strTemp, vbTab, "\t")
            strResult = strResult & """" & strTemp & """"
    End Select

End Sub

