Function UpdateTitleColumns(TitleList As Variant, oRange As Range, Optional ByVal idx_From As Long = -1, Optional ByVal idx_End As Long = -1, Optional ByVal idx_Row As Long = 1, Optional bLast As Boolean = True, Optional ByRef dic As Object) As Object
   Dim vLabel As Variant
   Dim s As String
   Dim i As Long
   Dim v As Variant
   Dim iC_From As Long, iC_End As Long, iR As Long
   Dim iC_Last As Long

   If idx_From < 0 Then idx_From = 1
   If idx_End < 0 Then idx_End = oRange.Columns.Count
   iC_From = oRange.Columns(idx_From).Column
   iC_End = oRange.Columns(idx_End).Column
   iR = oRange.Rows(idx_Row).Row

   If dic Is Nothing Then Set dic = CreateObject("Scripting.Dictionary")

   With oRange.Parent
      vLabel = .Range(.Cells(iR, iC_From), .Cells(iR, iC_End))
   End With
   If Not IsEmpty(vLabel) Then
      iC_Last = iC_From - 1
      For i = 1 To UBound(vLabel, 2)
         s = Trim(vLabel(1, i))
         If s <> "" Then
            If dic.Exists(s) Then
               If bLast Then dic(s) = (i + iC_From - 1)
            Else
               dic.Add Key:=s, Item:=(i + iC_From - 1)
            End If
            iC_Last = i + iC_From - 1
         End If
      Next i
   End If
   For Each v In TitleList
      If Not dic.Exists(v) Then
         iC_Last = iC_Last + 1
         dic.Add Key:=v, Item:=iC_Last
         oRange.Parent.Cells(iR, iC_Last).Value = v
      End If
   Next v
   Set UpdateTitleColumns = dic
End Function
