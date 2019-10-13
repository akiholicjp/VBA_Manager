Attribute VB_Name = "Z_VBA_UnitTest"
''
' VBAUnittest v1.0.2
' Copyright(c) 2016 takus - https://github.com/takus69/VBAUnittest
'
' @author takus4649@gmail.com
' @license MIT (https://opensource.org/licenses/MIT)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'VBA: Import ../Dump.bas: Private

Option Explicit

Dim tests As Object
Dim runCount As Integer
Dim failedCount As Integer
Dim assertFailedCount As Integer
Dim assertCount As Integer
Dim assertMessage As String
Dim runningTest As String
Dim excludedTests As Object
Dim excludedModules As Object
Const setUpMethodName As String = "setUp"
Const tearDownMethodName As String = "tearDown"
Const C_COMP_NAME_PREFIX As String = ""
Const xUnitTest_GUI_ENABLE As Boolean = False
Const C_PROC_NAME_PREFIX_NORMAL As String = "xUnitTest"
Const C_PROC_NAME_PREFIX_GUI As String = "xUnitTestGUI"

Private Enum xUnitTestType
   Normal = 1
   GUI = 2
End Enum

' Public procedure
' Setting excluding tests or modules
Sub setExcludedTests()
'    addExcludedTest "TestModule.testExcludedTest"
'    addExcludedTest "TestModule2.testExcludedTest"
End Sub

Sub setExcludedModules()
'    addExcludedModule "TestExcludedModule"
End Sub

' Test runner
Sub testRun(test As String)
   testInit
   oneTestRun test

   showResult
End Sub

Sub testInit()
   runCount = 0
   failedCount = 0
   Set tests = CreateObject("Scripting.Dictionary")
   Set excludedTests = CreateObject("Scripting.Dictionary")
   Set excludedModules = CreateObject("Scripting.Dictionary")
   setExcludedTests
   setExcludedModules
End Sub

Sub addTest(test As String)
   tests.add tests.Count, test
End Sub

Sub suiteRun()
   Dim i As Integer

   For i = 0 To tests.Count - 1
      oneTestRun tests.Item(i)
   Next i

   showResult
End Sub

Sub testModuleRun(TestModule As String)
   testInit
   addTestsInTestModule TestModule
   suiteRun
End Sub

Sub allTestRun()
   testInit
   addAllTest
   suiteRun
End Sub

' Assertion
Function assertBaseTrue(status As Boolean) As Boolean
   assertCount = assertCount + 1
   If Not status Then
      assertFailedCount = assertFailedCount + 1
   End If
   assertBaseTrue = status
End Function

Function assertTrue(status As Boolean) As Boolean
   Dim ret As Boolean
   ret = assertBaseTrue(status)
   If Not ret Then
      addAssertMessage setAssert(True, status)
   End If
   assertTrue = ret
End Function

Function assertFalse(status As Boolean) As Boolean
   Dim ret As Boolean
   ret = assertBaseTrue(Not status)
   If Not ret Then
      addAssertMessage setAssert(False, status)
   End If
   assertFalse = ret
End Function

Function assertObject(expected, actual) As Boolean
   Dim ret As Boolean
   ret = assertBaseTrue(expected Is actual)
   If Not ret Then
      addAssertMessage setAssert(expected, actual)
   End If
   assertObject = ret
End Function

Function assert(expected, actual) As Boolean
   Dim ret As Boolean
   If IsObject(expected) Or IsObject(actual) Then
      If IsObject(expected) And IsObject(actual) Then
         ret = assertBaseTrue(expected Is actual)
      Else
         ret = assertBaseTrue(False)
      End If
   ElseIf IsNull(expected) Or IsNull(actual) Then
      If IsNull(expected) And IsNull(actual) Then
         ret = assertBaseTrue(True)
      Else
         ret = assertBaseTrue(False)
      End If
   ElseIf IsArray(expected) Or IsArray(actual) Then
      If IsArray(expected) And IsArray(actual) Then
         ret = assertBaseTrue(Dump(expected) = Dump(actual))
      Else
         ret = assertBaseTrue(False)
      End If
   Else
      ret = assertBaseTrue(expected = actual)
   End If

   If Not ret Then
      addAssertMessage setAssert(Dump(expected), Dump(actual))
   End If
   assert = ret
End Function

Function assertNe(expected, actual) As Boolean
   Dim ret As Boolean
   If IsObject(expected) Or IsObject(actual) Then
      If IsObject(expected) And IsObject(actual) Then
         ret = assertBaseTrue(Not expected Is actual)
      Else
         ret = assertBaseTrue(True)
      End If
   Else
      ret = assertBaseTrue(expected <> actual)
   End If
   If Not ret Then
      addAssertMessage setAssert(expected, actual, False)
   End If
   assertNe = ret
End Function

' Messages
Function testSummary() As String
   testSummary = runCount & " run, " & failedCount & " failed"
End Function

Function failedMessage() As String
   failedMessage = runningTest & ", Count of assertion is " & assertCount & vbCrLf & assertMessage
End Function

Function addAssertMessage(ByVal str As String)
   str = ", at assertion " & assertCount & ", " & str
   If assertMessage = "" Then
      assertMessage = str
   Else
      assertMessage = assertMessage & vbCrLf & str
   End If
End Function

Function isSetUp(TestModule As String) As Boolean
   Dim methodName As String, i As Long

   With ThisWorkbook.VBProject.VBComponents(TestModule).CodeModule
      For i = 1 To .CountOfLines
         methodName = .ProcOfLine(i, 0)
         If methodName = setUpMethodName Then
               isSetUp = True
               Exit Function
         End If
      Next i
   End With

   isSetUp = False
End Function

Function isTearDown(TestModule As String) As Boolean
   Dim methodName As String, i As Long

   With ThisWorkbook.VBProject.VBComponents(TestModule).CodeModule
      For i = 1 To .CountOfLines
         methodName = .ProcOfLine(i, 0)
         If methodName = tearDownMethodName Then
               isTearDown = True
               Exit Function
         End If
      Next i
   End With

   isTearDown = False
End Function

' Private procedure
Private Sub oneTestRun(test As String)
   Dim runStatus As Boolean, arr() As String, runningModule As String
   assertFailedCount = 0
   assertCount = 0
   assertMessage = ""
   runningTest = test
   runningModule = fetchModule(test)

   If isSetUp(runningModule) Then
      Application.Run runningModule & "." & setUpMethodName
   End If

   runCount = runCount + 1
   Application.Run test

   If isTearDown(runningModule) Then
      Application.Run runningModule & "." & tearDownMethodName
   End If

   If assertFailedCount > 0 Then
      failedCount = failedCount + 1
      showFailed
   End If
End Sub

Private Function fetchModule(testMethod)
   Dim runningModule As String, arr() As String

   runningModule = ""
   arr = Split(testMethod, ".")
   If UBound(arr) = 1 Then
      runningModule = arr(0)
   End If

   fetchModule = runningModule
End Function

Private Sub showResult()
   Dim result As String

   If failedCount = 0 Then
      result = "green"
   Else
      result = "red"
   End If
   Debug.Print result & " : " & testSummary
End Sub

Private Sub showFailed()
   Debug.Print failedMessage
End Sub

Private Function setAssert(expected, actual, Optional ByVal eq As Boolean = True) As String
   If eq Then
      setAssert = "Expected:" & expected & ", " & "Actual:" & actual
   Else
      setAssert = "UnExpected:" & expected & ", " & "Actual:" & actual
   End If
End Function

Private Function fetchProcs(TestModule As String, eType As xUnitTestType) As String()
   Dim buf As String, testName As String, procNames() As String, i As Long, cnt As Integer
   Dim sType As String
   cnt = -1
   Select Case eType
   Case xUnitTestType.Normal
      sType = C_PROC_NAME_PREFIX_NORMAL
   Case xUnitTestType.GUI
      sType = C_PROC_NAME_PREFIX_GUI
   Case Else
      sType = ""
   End Select
   With ThisWorkbook.VBProject.VBComponents(TestModule).CodeModule
      For i = 1 To .CountOfLines
         testName = TestModule & "." & .ProcOfLine(i, 0)
         If buf <> testName And .ProcOfLine(i, 0) Like (sType & "_*") Then
               buf = testName
               If Not isTestExcluded(testName) Then
                  cnt = cnt + 1
                  ReDim Preserve procNames(cnt)

                  procNames(cnt) = testName
               End If
         End If
      Next i
   End With

   fetchProcs = procNames
End Function

Private Function isTestExcluded(testName As String) As Boolean
   isTestExcluded = excludedTests.exists(testName)
End Function

Private Function isModuleExcluded(moduleName As String) As Boolean
   isModuleExcluded = excludedModules.exists(moduleName)
End Function

Private Sub addTestsInTestModule(TestModule As String)
   Call addTestsInTestModuleSub(TestModule, xUnitTestType.Normal)
   If xUnitTest_GUI_ENABLE Then
      Call addTestsInTestModuleSub(TestModule, xUnitTestType.GUI)
   End If
End Sub

Private Sub addTestsInTestModuleSub(TestModule As String, eType As xUnitTestType)
   Dim procNames() As String
   Dim i As Integer

   procNames = fetchProcs(TestModule, eType)
   If (Not procNames) <> -1 Then ' Check no array data
      For i = 0 To UBound(procNames)
         addTest procNames(i)
      Next i
   End If
End Sub

Private Sub addAllTest()
   Dim comp As Object, procNames() As String

   For Each comp In ThisWorkbook.VBProject.VBComponents
      If comp.Name Like (C_COMP_NAME_PREFIX & "*") And Not isModuleExcluded(comp.Name) Then
         addTestsInTestModule comp.Name
      End If
   Next comp
End Sub

Private Sub addExcludedTest(excludedTest As String)
   excludedTests.add excludedTest, True
End Sub

Private Sub addExcludedModule(excludedModule As String)
   excludedModules.add excludedModule, True
End Sub
