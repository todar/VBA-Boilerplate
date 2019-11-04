VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalyticsTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * This is intended to gather data in how VBA code is being used. This can capture
' * code use, new users, code completion, errors, timing, trends and other metrics.
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @ref {Library} Microsoft Scripting Runtime
' * @ref {Class Module} JSON
' * @ref {Module} FileSystemsUtilities
' * @ref {Function} FileSystemsUtilities.BuildOutFilePath
' */
Option Explicit

'/**
' * State is a {Type} containing all the private variables for this class.
' */
Private Type state
    stats As Scripting.Dictionary
    startTime As Double
    wasStarted As Boolean
    wasReported As Boolean
    isDisabled As Boolean
End Type

'/**
' * GUILD is unique ids. These are used to access specifc record to append end times to it.
' */
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

'/**
' * property {Long} CoCreateGuid - API Needed to generate GUID.
' * property {state} this - The private properties stored in one place.
' * property {String} filePath - This is where the xml is stored.
' */
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Private this As state
Private Const FILEPATH As String = "C:\test\Analytics.json"

'/**
' * Create new objects on init. Need to make sure xml is valid, will disable if not.
' */
Private Sub Class_Initialize()
    Set this.stats = New Scripting.Dictionary
    
    ' Make sure there is a file to write to.
    createAnalyticsFileIfNeeded
End Sub

Private Sub createAnalyticsFileIfNeeded()
    Dim fso As New FileSystemObject
    If Not fso.FileExists(FILEPATH) Then
        If Not BuildOutFilePath(FILEPATH) = True Then
            this.isDisabled = True
            Debug.Print "Analytics are not being tracked. Unable to create filepath: " & FILEPATH
        End If
    End If
End Sub

'/**
' * If analytics are being tracked and this object is destroyed then it will complete on it's own.
' */
Private Sub Class_Terminate()
    ' Only finalize if not already reported.
    If this.wasStarted And Not this.wasReported Then
        FinalizeStats
    End If
End Sub

'/**
' * This is the main init by the user. It logs the codeName and starts everything off.
' * @param {String} codeName - The name of the code (function) being run.
' * @param {Double} [minutesToSave] - Allows user to mark ahead of time the time saved.
' */
Public Sub TrackStats(ByVal codeName As String, Optional ByVal minutesToSave As Variant)
    If this.isDisabled Then
        Exit Sub
    End If
    
    ' Note that is was started to make sure the stats are logged.
    this.wasStarted = True
    AddStat "event", "onUse"
    AddStat "id", UUID
    AddStat "codeName", codeName
    AddStat "username", Environ("UserName")
    AddStat "date", CStr(Date)
    AddStat "time", CStr(Time)
    If Not IsMissing(minutesToSave) Then
        AddStat "minutesSaved", minutesToSave
    End If
    this.startTime = timer
End Sub

Public Sub AddStat(ByVal Key As String, ByVal value As Variant)
    this.stats(Key) = value
End Sub

Public Sub LogEvent(ByVal eventName As String, ParamArray keyValuePairs() As Variant)
    ' Check to see that key/value pairs passed in (an even number).
    If (UBound(keyValuePairs) - LBound(keyValuePairs) + 1) Mod 2 <> 0 Then
        Debug.Print "[Analitics.LogEvent] Unable to log. Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
        Exit Sub
    End If
    
    ' Add standard stats
    Dim stats As New Scripting.Dictionary
    Set stats = New Scripting.Dictionary
    stats.Add "event", eventName
    stats.Add "id", UUID
    stats.Add "username", Environ("UserName")
    stats.Add "date", CStr(Date)
    stats.Add "time", CStr(Time)
    
    ' Add key values to the return Dictionary.
    Dim index As Long
    For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        stats.Add keyValuePairs(index), keyValuePairs(index + 1)
    Next index
    
    Debug.Print JSON.Stringify(stats, 2)
    
    AppendToTextFile FILEPATH, JSON.Stringify(stats) & ","
End Sub

'/**
' * This finilized the analytics and adds all the finished runtimes.
' * This is also called from Class_Terminate if user doesn't manully call this method.
' *
' * @param {Double} [minutesToSave] - Allows user to add time saved.
' */
Public Sub FinalizeStats(Optional ByVal minutesToSave As Variant)
    If this.isDisabled Then
        Exit Sub
    End If
    
    ' Mark as finalized to not run twice
    this.wasReported = True
    
    ' Get the total runtime, rounding this to the third decimal place.
    AddStat "runtimeSeconds", Math.Round(timer - this.startTime, 3)

    If Not IsMissing(minutesToSave) Then
        AddStat "minutesSaved", minutesToSave
    End If

    Debug.Print JSON.Stringify(this.stats, 2)
    
    ' Append record to log file. Note, this file is JSON but will not have needed
    ' Array [] syntax due to appending is much faster then rewriting the entire file.
    ' Will need to run function to fix this.
    AppendToTextFile FILEPATH, JSON.Stringify(this.stats) & ","
End Sub

'/**
' * Creates an unique id. Used to find xml record in xml file. Needed to add end times.
' *
' * @author Gus Molina (open sourced)
' * @copywrite 2000
' * @see <https://stackoverflow.com/questions/7031347/how-can-i-generate-guids-in-excel>
' */
Private Function UUID() As String
    Dim udtGUID As GUID
    If (CoCreateGuid(udtGUID) = 0) Then
        UUID = _
            String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
            String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
            String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
            IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
            IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
            IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
            IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
            IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
            IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
            IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
            IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function
