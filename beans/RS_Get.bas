Function RS_Get(ByRef oRS As Object, ByVal field As String, Optional ByVal vNull As Variant = "") As Variant
   Dim v As Variant
   v = oRS(field)
   If IsNull(v) Then
      RS_Get = vNull
   Else
      RS_Get = v
   End If
End Function
