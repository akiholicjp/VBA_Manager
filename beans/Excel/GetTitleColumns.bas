Function GetTitleColumns(oRange As Range, Optional ByVal column_From As Long = -1, Optional ByVal column_End As Long = -1, Optional ByVal idx_Row As Long = 1) As Object
   Dim dic As Object
   Dim sLabel As String
   Dim i As Long

   With oRange
      If column_From < 0 Then column_From = 1
      If column_End < 0 Then column_End = .Columns.Count
      Set dic = CreateObject("Scripting.Dictionary")
      For i = column_From To column_End
         sLabel = Trim(.Cells(idx_Row, i).Text)
         If sLabel <> "" Then
            If dic.Exists(sLabel) Then
               ' èdï°ÇµÇΩèÍçáÇÕñ≥éã
            Else
               dic.Add Key:=sLabel, Item:=i
            End If
         End If
      Next i
   End With
   Set GetTitleColumns = dic
End Function
