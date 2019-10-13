Function GetWorksheetProperty(tSheet As Worksheet, ByVal sKEY As String) As String
   Dim v As Variant
   For Each v In tSheet.Names
      If v.Name = sKEY Or v.Name Like "*!" & sKEY Then Exit For
   Next
   If IsEmpty(v) Then
      GetWorksheetProperty = ""
   ElseIf Not (v.Name = sKEY Or v.Name Like "*!" & sKEY) Then
      GetWorksheetProperty = ""
   Else
      GetWorksheetProperty = tSheet.Names(sKEY).Comment
   End If
End Function
