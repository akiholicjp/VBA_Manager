Function GetRangeIgnoreHidden(ByRef rng As Range) As Variant
   Dim oSel As Range
   Dim v As Variant
   For Each v In rng
      If (Not v.Rows.Hidden) And (Not v.Columns.Hidden) Then
         If oSel Is Nothing Then Set oSel = v Else Set oSel = Union(oSel, v)
      End If
   Next v
   Set GetRangeIgnoreHidden = oSel
End Function
