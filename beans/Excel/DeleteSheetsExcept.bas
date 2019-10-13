Sub DeleteSheetsExcept(oSheet As Worksheet)
   Dim sName As String
   Dim v As Variant
   sName= oSheet.Name
   For Each v In oSheet.Parent.Worksheets
      If v.Name <> sName Then v.Delete
   Next v
End Sub
