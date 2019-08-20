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
' * @module {Class} MicroStopWatch
' * @ref {Microsoft xml, v6.0} MSXML2.DOMDocument60, IXMLDOMElement
' */
Option Explicit

'/**
' * State is a {Type} containing all the private variables for this class.
' */
Private Type state
    id As String
    timer As MicroStopWatch
    doc As MSXML2.DOMDocument60
    record As IXMLDOMElement
    isComplete As Boolean
    wasStarted As Boolean
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
Private Const filePath As String = "C:\rtmacros\Data\Analytics.xml"

'/**
' * Create new objects on init. Need to make sure xml is valid, will disable if not.
' */
Private Sub Class_Initialize()
    Set this.timer = New MicroStopWatch
    Set this.doc = New MSXML2.DOMDocument60
    Set this.record = this.doc.CreateElement("record")
    If this.doc.Load(filePath) = False Then
        this.isDisabled = True
        Debug.Print "Analytics are not being tracked. Possible the xml file is corrupt or missing."
    End If
End Sub

'/**
' * If analytics are being tracked and this object is destroyed then it will complete on it's own.
' */
Private Sub Class_Terminate()
    If Not this.isComplete And this.wasStarted Then
        CompleteStats
    End If
End Sub

'/**
' * Adds custom xml elements to the current record.
' * @param {String} name - This is the xml Tag name to be added.
' * @param {Double} [minutesToSave] - Allows user to mark ahead of time the time saved.
' */
Public Sub AddField(ByVal name As String, ByVal value As Variant)
    CreateElement name, value
End Sub

'/**
' * Get count of all records from filepath that don't have an endtime.
' */
Public Property Get CountOfInteruptedCode() As Long
    this.doc.Load filePath
    CountOfInteruptedCode = this.doc.SelectNodes("//analytics/record[not(endTime)]").length
End Property

'/**
' * This is the main init by the user. It logs the codeName and starts everything off.
' * @param {String} codeName - The name of the code (function) being run.
' * @param {Double} [minutesToSave] - Allows user to mark ahead of time the time saved.
' */
Public Sub Track(ByVal codeName As String, Optional ByVal minutesToSave As Double)
    If this.isDisabled Then
        Exit Sub
    End If
    
    this.wasStarted = True
    this.id = GetGUID
    this.record.setAttribute "id", this.id
    CreateElement "codeName", codeName
    CreateElement "username", Environ("UserName")
    CreateElement "date", Date
    CreateElement "startTime", Time
    CreateElement "minutesSaved", minutesToSave
    LogRecord
    this.timer.StartTimer
End Sub

'/**
' * Private Helper function to easily create new xml elements
' * and add them to the record element.
' *
' * @param {String} name - This is the xml Tag name to be added.
' * @param {Variant} value - The text value to be added to the xml element.
' */
Private Sub CreateElement(ByVal name As String, ByVal value As Variant)
    Dim element As IXMLDOMElement
    Set element = this.record.SelectSingleNode(name)
    If element Is Nothing Then
        Set element = this.doc.CreateElement(name)
        this.record.appendChild element
    End If
    element.Text = value
End Sub

'/**
' * This finilized the analytics and adds all the finished runtimes.
' * This is also called from Class_Terminate if user doesn't manully call this method.
' *
' * @param {Double} [minutesToSave] - Allows user to add time saved.
' */
Public Sub CompleteStats(Optional ByVal minutesToSave As Double)
    If this.isDisabled Then
        Exit Sub
    End If
    
    CreateElement "endTime", Time
    CreateElement "runtimeSeconds", Round(this.timer.Seconds, 3)

    this.isComplete = True
    If minutesToSave <> 0 Then
        CreateElement "minutesSaved", minutesToSave
    End If
    
    AppendRecord
    Debug.Print this.record.xml
End Sub

'/**
' * Private helper to find record based on ID and append that xml element.
' */
Private Sub AppendRecord()
    If this.isDisabled Then
        Exit Sub
    End If

    ' check if xml is loaded correctly...
    If this.doc.Load(filePath) = False Then
        Exit Sub
    End If
    
    Dim oldRecord As IXMLDOMElement
    Set oldRecord = this.doc.SelectSingleNode("//analytics/record[@id='" & this.id & "']")
    
    this.doc.FirstChild.replaceChild this.record, oldRecord
    this.doc.Save filePath
End Sub

'/**
' * Private helper to add xml record element to the xml file stored at const filepath.
' */
Private Sub LogRecord()
    If this.isDisabled Then
        Exit Sub
    End If

    ' double check if xml is loaded correctly. Checked at the beggining .
    If this.doc.Load(filePath) = False Then
        Debug.Print "Analytics are not being tracked. Possible the xml file is corrupt or missing."
        Exit Sub
    End If
    
    ' Add new node to root element.
    this.doc.FirstChild.appendChild this.record
    this.doc.Save filePath
End Sub

'/**
' * Creates an unique id. Used to find xml record in xml file. Needed to add end times.
' *
' * @author Gus Molina (open sourced)
' * @copywrite 2000
' * @see <https://stackoverflow.com/questions/7031347/how-can-i-generate-guids-in-excel>
' */
Private Function GetGUID() As String
    Dim udtGUID As GUID
    If (CoCreateGuid(udtGUID) = 0) Then

        GetGUID = _
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
