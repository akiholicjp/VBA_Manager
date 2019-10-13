Function RS_Set(ByRef oRS As Object, ByVal field As String, ByRef val As Variant) As Variant
   oRS(field) = val
   RS_Set = oRS(field)
End Function
