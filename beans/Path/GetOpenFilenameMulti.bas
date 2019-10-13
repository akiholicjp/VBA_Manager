Function GetOpenFilenameMulti(Optional FileFilter As Variant, Optional FilterIndex As Variant, Optional Title As Variant, Optional ButtonText As Variant) As Collection
   Dim tmpText As Variant, sTmp As Variant
   Dim j As Long

   Set GetOpenFilenameMulti = New Collection

   tmpText = Application.GetOpenFilename(FileFilter, FilterIndex, Title, ButtonText, True)
   If VarType(tmpText) = vbBoolean Then Exit Function

   '----- 処理ファイルのソート(GetOpenFilenameは順序が明確ではないため) -----
   With GetOpenFilenameMulti
      For Each sTmp In tmpText
         For j = 1 To .Count
            If StrComp(sTmp, .Item(j), vbTextCompare) < 0 Then
               Exit For
            End If
         Next
         If j > .Count Then
            .Add sTmp
         Else
            .Add sTmp, Before:=j
         End If
      Next
   End With
End Function
