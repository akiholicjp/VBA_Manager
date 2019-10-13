Function SelectFolderName(ByVal dispTitle As String) As String
   Dim filePath As String
   Dim i As Integer

   With Application.FileDialog(msoFileDialogFolderPicker)
      .Title = dispTitle
      .InitialView = msoFileDialogViewList
      If .Show = -1 Then
         filePath = .SelectedItems.Item(1)
      Else
         filePath = "False"
      End If

   End With
   SelectFolderName = filePath
End Function
