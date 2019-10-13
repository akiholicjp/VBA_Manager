' VBA: Import ../GetFSO.bas

Function GetWorkbook(ByVal strBookName As String, Optional ByRef oDefault As Object = Nothing, Optional FSO As Object = Nothing) As Workbook
   Dim v As Workbook
   Set GetWorkbook = oDefault
   ' Workbooks.Item(hoge) ���ƁA�g���q�܂ň�v����K�v������̂Ō��������Ƃ���B
   With GetFSO()
      For Each v In Workbooks
         If v.Name = strBookName Or v.FullName = strBookName Or .GetBaseName(v.Name) = strBookName Then
            Set GetWorkbook = v
            Exit For
         End If
      Next v
   End With
End Function
