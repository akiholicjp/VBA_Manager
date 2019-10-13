' VBA: Import UniRegExp.bas
Function RegExe(ByVal sTest As String, ByVal sPattern As String, Optional ByRef oResults As Object, Optional ByRef RegExp As Variant) As Boolean
   Dim oMatch As Variant
   Dim oMatches As Object
   Dim v As Variant
   Set oResults = New Collection
   If IsMissing(RegExp) Then
      Set oMatches = UniRegExp(sPattern).Execute(sTest)
   Else
      RegExp.Pattern = sPattern
      Set oMatches = RegExp.Execute(sTest)
   End If
   RegExe = (oMatches.Count > 0)
   For Each oMatch In oMatches
      For Each v In oMatch.SubMatches
         oResults.Add v
      Next v
   Next oMatch
End Function
