Attribute VB_Name = "MainApplication"
'**/
' * This is the main entry point for the app. Notes for the app should be stored
' * within this section. List all reference and Global Classes here as well.
' *
' * Global Classes [Console, JSON, LocalStorage] These can all be called globally.
' * @ref {Library} Microsoft Scripting Runtime
' * @ref {Library} Microsoft Internet Controls
' * @ref {Library} Microsoft VBScript Regular Expressions 5.5 [RegExp, Match]
' * @ref {Library} Microsoft ActiveX Data Objects 6.1 Library
' */
Option Explicit

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
