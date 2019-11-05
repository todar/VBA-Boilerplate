Attribute VB_Name = "MainTesting"
'**/
' * This is the main testing section. Contains factory method to easily implement
' * tests. Also has a place to run all tests.
' * Naming convention for tests should start with `Test` to easily find them.
' *
' * @see TestSampleTests for an example.
' * Global Classes [Console, JSON, LocalStorage] These can be called globally.
' * @ref {Class Module} TestSuite - Collection of tests.
' * @ref {Class Module} TestCase - A group of related tests.
' * @ref {Class Module} TestMatcher - The test functions.
' */
Option Explicit

'**/
' * TestSuites Factory method. This creates an easy way to run tests
' * without needing to create an object everytime.
' *
' * @see TestSampleTests for an example.
' */
Public Function test(ByVal description As String) As TestSuite
    Set test = New TestSuite
    test.description = description
End Function

'**/ A list of project tests */
Public Function TestAll()
    TestSampleTests
    TestTestMatcherFunctions
End Function

'**/ This is just a sample how to use the unit testing tools. */
Public Function TestSampleTests() As TestSuite
    With test("Sample Math tests")
        With .test("two plus two")
            .Expect(2 + 2).ToEqual 4
            .Expect(2 + 2).ToBeLessThan 5
        End With
        
        With .test("six times one")
            .Expect(6 * 1).ToEqual 6
            .Expect(6 * 1).ToBeGreaterThanOrEqual 6
        End With
        
        With .test("five minus two")
            .Expect(5 - 2).ToEqual 3
            .Expect(5 - 2).ToBeGreaterThan 2
        End With
        
        With .test("Adding 1 + 1 equals 2")
            .Expect(1 + 1).ToEqual 2
        End With
    End With
End Function

'**/ Unit tests for checking the functions within the unit tester. */
Public Function TestTestMatcherFunctions() As TestSuite
    With test("TestMatcher Functions")
        With .test("Approximate For Doubles")
            .Expect(1.001).ToBeApproximate 1.003, 1
            .Expect(1.001).ToNotBeApproximate 1.003, 5
        End With
        
        With .test("Defined & Undefined")
            .Expect("").ToBeDefined
            
            .Expect().ToBeUndefined
            Dim notDefined As Variant
            .Expect(notDefined).ToBeUndefined
        End With
        
        With .test("Truth & Falsy")
            .Expect(True).ToBeTruthy
            .Expect(1 = 1).ToBeTruthy
            
            .Expect(False).ToBeFalsy
            .Expect(1 = 4).ToBeFalsy
        End With
        
        With .test("equality")
            .Expect(False).ToEqual False
            .Expect(1).ToEqual 1
            .Expect("a").ToEqual "a"
            
            .Expect(False).ToNotEqual "Something Else"
            .Expect(1).ToNotEqual 2
            .Expect("a").ToNotEqual "b"
            
            .Expect(2).ToBeGreaterThan 1
            .Expect(2).ToBeGreaterThanOrEqual 1
            .Expect(2).ToBeGreaterThanOrEqual 2
            
            .Expect(1).ToBeLessThan 2
            .Expect(1).ToBeLessThanOrEqual 2
            .Expect(2).ToBeLessThanOrEqual 2
        End With
    End With
End Function

'/**
' * Sample of how to track and use Analytics class.
' * @ref {Class Module} AnalyticsTracker
' * @ref {Class Module} JSON
' * @ref {Module} FileSystemUtilities
' * @ref {Library} Microsoft Scripting Runtime
' */
Private Sub howToTrackAnalytics()
    ' This tracks to a JSON file and the immediate window.
    ' To be effecent this appends to the text file.
    ' Because of this the JSON file is missing the outer array
    ' brackets []. Also includes a comma after each object {},
    ' So to use this as JSON you must edit those two things.
    Dim analytics As New AnalyticsTracker
    
    ' You can track standard stats for code use!
    ' This collects codeName, username, date, time, timesaved, runtime
    analytics.TrackStats "test", 5
    
    ' Can also add custom stats to the main thread.
    analytics.AddStat "customStat", "I'm custom!"
    
    ' Also have the ability to log your own custom events. This by default
    ' still adds things like date, time, username.
    analytics.LogEvent "onCustom", "name", "Robert", "age", 31
    
    ' Optional. You can either call this function, or let the
    ' terminate event in the class to run it.
    ' An example log looks like: {"event":"onUse", ...},
    analytics.FinalizeStats
End Sub


'/**
' * Sample of how to use JSON
' * @ref {Class Module} JSON
' * @ref {Library} Microsoft Scripting Runtime
' */
Private Sub howToUseJSON()
    ' Here is some sample data. The outer object must be an Array, Dictionary, or
    ' a Collection. This is to make it valid JSON.
    ' A note is that Arrays will be parsed as collections for performance reasons.
    Dim config As New Scripting.Dictionary
    config.Add "name", "JSONSample"
    config.Add "users", Array("robert", "Fred", "Mark")
    
    ' Stringify will turn a Dictionary or Array into a string of JSON.
    ' This can be stored in a text file a parsed back into an object
    ' using JSON.Parse. Example of that below.
    ' This example: ~> {"name":"JSONSample","users":["robert","Fred","Mark"]}
    Dim JSONString As String
    JSONString = JSON.Stringify(config)
    Debug.Print JSONString
    
    ' This example uses the JSON string from above and converts it into
    ' VBA Dictionary. Note that the Array from above is converted into
    ' a Collection. This is due to performance in iterating over the list
    ' while parsing.
    Dim clone As Scripting.Dictionary
    Set clone = JSON.Parse(JSONString)
End Sub

'/**
' * Sample of how to use Console
' * @ref {Modlue} FileSystemUtilities
' * @ref {Function} FileSystemUtilities.BuildOutFilePath
' * @ref {Function} FileSystemUtilities.ReadTextFile
' * @ref {File} `./templates/console.html`
' */
Private Sub howToUseConsole()
    ' Each method takes in a source so the log knows where the log occured
    ' and a message, and this can be whatever makes the most sense.
    ' The various methods have a different style logo for each one and
    ' would be used to filter on each one.
    ' This also logs to the immediate window as well.
    Console.Log "Testing.howToUseConsole", "This is a sample info log!"
    Console.Warn "Testing.howToUseConsole", "This is a sample warning log!"
    Console.Error "Testing.howToUseConsole", "This is a sample error log!"
End Sub

'/**
' * Sample of how to use LocalStorage
' * @ref {Library} Microsoft Scripting Runtime
' * @ref {Class Module} Console - Used to log to immediate window and log files.
' * @ref {Class Module} JSON - Used to store key value pairs.
' * @ref {Modlue} FileSystemUtilities
' * @ref {Function} FileSystemUtilities.BuildOutFilePath
' * @ref {Function} FileSystemUtilities.ReadTextFile
' */
Private Sub howToUseLocalStorage()
    ' Setting an item is easy, it stores info based on key value pairs.
    ' This will either override what's in there or create a new pair if
    ' it doesn't already exist.
    LocalStorage.SetItem "name", "Robert"
    
    ' Getting an item you simply call it by the key you stored it as.
    ' By default you have to pass in a fallback value. This value is used
    ' only if there is some error in retriving the value or if it doesn't
    ' exist. This helps prevent any errors.
    Dim name As String
    name = LocalStorage.GetItem("name", "Robert")
    
    ' Easily show the storage in the immediate window.
    LocalStorage.DisplayToImmediateWindow
    
    ' Easily remove a sinlge item by it's key
    LocalStorage.RemoveItem "name"
    
    ' Or you can remove all items and reset the storage back to an empty object {}
    LocalStorage.Clear
End Sub






















