' Function ExistsWorkbook(ByVal strBookName As String) As Boolean
'    Dim o As Workbook
'    ExistsWorkbook = False
'    For Each o In Workbooks
'       If o.Name = strBookName Then
'          ExistsWorkbook = True
'          Exit For
'       ElseIf o.FullName = strBookName Then
'          ExistsWorkbook = True
'          Exit For
'       End If
'    Next o
' End Function

Function ExistsWorkbook(ByVal strBookName As String) As Boolean
   Dim o As Workbook
   On Error Resume Next
   Set o = Workbooks(strBookName)
   ExistsWorkbook = Not (o Is Nothing)
   On Error GoTo 0
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ExistsWorkbook()
   Dim oWb As Workbook

   On Error Resume Next
   Application.Workbooks("Book1").Close
   On Error GoTo 0

   assert False, ExistsWorkbook("Book1")
   Set oWb = Application.Workbooks.Add

   assert True, ExistsWorkbook(oWb.Name)
   Application.Workbooks(oWb.Name).Close
End Sub
