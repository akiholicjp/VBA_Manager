' Function ExistsWorksheet(ByRef oBook As Workbook, ByVal strSheetName As String) As Boolean
'    Dim o As Worksheet
'    ExistsWorksheet = False
'    If oBook Is Nothing Then Exit Function
'    For Each o In oBook.Worksheets
'       If o.Name = strSheetName Then
'          ExistsWorksheet = True
'          Exit For
'       End If
'    Next o
' End Function

Function ExistsWorksheet(ByRef oBook As Workbook, ByVal strSheetName As String) As Boolean
   Dim o As Worksheet
   On Error Resume Next
   Set o = oBook.Worksheets(strSheetName)
   ExistsWorksheet = Not (o Is Nothing)
   On Error GoTo 0
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ExistsWorksheet()
   Dim oWb As Workbook
   Dim oWs As Worksheet

   Set oWb = Application.Workbooks.Add

   assert False, ExistsWorksheet(oWb, "ExistsWorksheet")

   Set oWs = oWb.Worksheets.Add
   oWs.Name = "ExistsWorksheet"

   assert True, ExistsWorksheet(oWb, "ExistsWorksheet")

   oWb.Close SaveChanges:=False
End Sub
