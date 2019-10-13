Sub SetWorksheetProperty(tSheet As Worksheet, ByVal sKEY As String, ByVal sComment As String)
   Dim v As Variant
   For Each v In tSheet.Names
      If v.Name = sKEY Or v.Name Like "*!" & sKEY Then Exit For
   Next
   If IsEmpty(v) Then
      tSheet.Names.Add Name:="TYPE", RefersTo:="=$A$1"
   ElseIf Not (v.Name = sKEY Or v.Name Like "*!" & sKEY) Then
      tSheet.Names.Add Name:="TYPE", RefersTo:="=$A$1"
   End If
   tSheet.Names(sKEY).Comment = sComment
End Sub

