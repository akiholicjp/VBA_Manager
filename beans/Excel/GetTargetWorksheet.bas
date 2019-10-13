' VBA: Import GetWorkbook.bas
' VBA: Import GetWorksheet.bas

Function GetTargetWorksheet(ByVal strSheet As String, Optional ByVal bForceOpen As Boolean = False, Optional ByRef bOpened As Boolean, Optional ByRef oBaseBook As Object = Nothing) As Worksheet
   Dim oReg As Object
   Dim sBook As String, sSheet As String
   Dim oBook As Workbook
   Dim v As Variant

   Set GetTargetWorksheet = Nothing
   bOpened = False

   Set oReg = CreateObject("VBScript.RegExp")
   oReg.Pattern = "'?(\[[^\]]+\])?([^'\[\]]+)'?"
   oReg.IgnoreCase = False
   oReg.Global = False

   Set oBook = oBaseBook
   If oReg.test(strSheet) Then
      sBook = Trim(oReg.Replace(strSheet, "$1"))
      sSheet = Trim(oReg.Replace(strSheet, "$2"))

      If sSheet = "" Then GoTo ExitProc

      If sBook <> "" Then
         sBook = Mid(sBook, 2, Len(sBook) - 2)

         Set oBook = GetWorkbook(sBook, oDefault:=oBook)
         If oBook Is Nothing And bForceOpen Then
            Set oBook = Workbooks.Open(sBook)
            bOpened = True
         End If
      End If
      If oBook Is Nothing Then GoTo ExitProc

      Set GetTargetWorksheet = GetWorksheet(oBook, sSheet)
   End If
ExitProc:
   Exit Function
ErrExit:
   Set GetTargetWorksheet = Nothing
   Resume ExitProc
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetTargetWorksheet()
   Dim o As Object
   Dim bOpened As Boolean

   Set o = GetTargetWorksheet("[test]TestSheet", bForceOpen:=False)
   If Not o Is Nothing Then o.Parent.Close

   Set o = GetTargetWorksheet("")
   assert "Nothing", TypeName(o)
   Set o = GetTargetWorksheet("[TestBook]GetTargetWorksheet")
   assert "Worksheet", TypeName(o)
   Set o = GetTargetWorksheet("'[TestBook]GetTargetWorksheet'")
   assert "Worksheet", TypeName(o)
   Set o = GetTargetWorksheet("'[" & ThisWorkbook.Path & "/test.xlsx]TestSheet'", bForceOpen:=True, bOpened:=bOpened)
   assert "Worksheet", TypeName(o)
   assert True, bOpened

   Set o = GetTargetWorksheet("TestSheet", bOpened:=bOpened, oBaseBook:=o.Parent)
   assert "Worksheet", TypeName(o)
   assert False, bOpened

   o.Parent.Close
End Sub
