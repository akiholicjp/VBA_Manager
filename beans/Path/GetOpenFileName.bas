Public Function GetOpenFileName(Optional ByVal FileFilter As Variant, Optional ByVal FilterIndex As Long = 1, Optional ByVal Title As String = "ファイルを開く", Optional ByVal MultiSelect As Boolean = False, Optional ByVal InitialFileName As String) As Variant
   Dim intRet As Integer
   Dim v As Variant, vv As Variant

   With Application.FileDialog(msoFileDialogOpen)
      .Title = Title
      .Filters.Clear
      For Each v In FileFilter
         vv = Split(v, ",")
         .Filters.Add vv(0), vv(1)
      Next v
      .Filters.Add "すべてのファイル", "*.*"
      .FilterIndex = FilterIndex
      .AllowMultiSelect = MultiSelect
      If IsEmpty(InitialFileName) Then
         .InitialFileName = CurrentProject.Path
      Else
         .InitialFileName = InitialFileName
      End If
      .Show
      Set GetOpenFileName = .SelectedItems
   End With
End Function
