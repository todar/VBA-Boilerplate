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
