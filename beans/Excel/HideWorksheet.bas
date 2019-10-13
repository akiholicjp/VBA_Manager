Sub HideWorksheet(oSheet As Worksheet)
   If oSheet.Parent.Name Like "*.xla*" Then
      oSheet.Parent.IsAddin = True
   Else
      oSheet.Visible = False
   End If
End Sub
