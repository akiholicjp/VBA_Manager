Function RS_Del(ByRef oRS As Object, ByVal field As String, ByRef target As Variant) As Variant
   Do While True
      oRS.FindFirst field & "='" & target & "'"
      If oRS.NoMatch Then
         oRS.FindFirst field & "=" & target
         If oRS.NoMatch Then
            Exit Do
         End If
      End If
      oRS.Delete
   Loop
End Function
