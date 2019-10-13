Function RangeToArray(ParamArray Target() As Variant)
   Dim v As Variant
   Dim vv As Variant
   Dim i As Long
   Dim ary() As Variant
   ReDim ary(UBound(Target))

   i = 0
   For Each v In Target
      If IsArray(v) Then
         ReDim Preserve ary(UBound(ary) + v.Count)
         For Each vv In v
            ary(i) = vv.Text
            i = i + 1
         Next vv
      Else
         ary(i) = v.Text
         i = i + 1
      End If
   Next v
   RangeToArray = ary
End Function
