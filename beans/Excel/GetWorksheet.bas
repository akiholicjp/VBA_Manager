Function GetWorksheet(oBook As Workbook, ByVal strSheetName As String, Optional ByRef oDefault As Object = Nothing) As Worksheet
   Set GetWorksheet = oDefault
   On Error Resume Next
   Set GetWorksheet = oBook.Worksheets(strSheetName)
End Function

' Function GetWorksheet(oBook As Workbook, ByVal strSheetName As String) As Worksheet
'    Dim v As Worksheet
'    Set GetWorksheet = Nothing
'    For Each v In oBook.Worksheets
'       If v.Name = strSheetName Then
'          Set GetWorksheet = v
'          Exit For
'       End If
'    Next v
' End Function
