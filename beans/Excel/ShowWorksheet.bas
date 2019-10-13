Sub ShowWorksheet(oSheet As Worksheet)
   If oSheet.Parent.Name Like "*.xla*" Then
      oSheet.Parent.IsAddin = False
   End If
   oSheet.Visible = True
End Sub
