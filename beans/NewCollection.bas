Function NewCollection(ParamArray Target() As Variant) As Object
   Dim oCol As New Collection
   Dim v As Variant
   For Each v In Target
      oCol.Add v
   Next v
   Set NewCollection = oCol
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_NewCollection()
   assert "[1%,""ABC"",2#]", Dump(NewCollection(1, "ABC", 2.0))
End Sub
