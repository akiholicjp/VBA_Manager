' VBA: Import ../Wrapper/_SHBrowseForFolder.bas
' VBA: Import ../Wrapper/_SHGetPathFromIDList.bas

Function GetFolderSimple(Optional ByVal strTitle As String = "Select Folder", Optional ByRef strPath As String) As String
   Dim bif As BROWSEINFO
   Dim pidl As Long
   On Error GoTo ErrGetFolder
   strPath = ""
   With bif
      .pidlRoot = &H0 'デスクトップ
      .ulFlags = &H1 'フォルダのみ選択可能
      .lpszTitle = strTitle
   End With
   pidl = SHBrowseForFolder(bif)
   If pidl <> 0 Then
      strPath = String$(256, vbNullChar)
      SHGetPathFromIDList pidl, strPath
      strPath = Left(strPath, InStr(strPath, vbNullChar) - 1)
   End If
   GetFolderSimple = strPath
   Exit Function
ErrGetFolder:
   GetFolderSimple = ""
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTestGUI_beans_GetFolderSimple()
   ' Usage
   Dim strPath As String, s As String
   strPath = ""
   s = GetFolderSimple(strTitle:="TEST", strPath:=strPath)
   MsgBox strPath, Title:=s

   s = GetFolderSimple(strTitle:="TEST", strPath:="c:\")
   MsgBox s, Title:=s

   s = GetFolderSimple()
   MsgBox s, Title:=s
End Sub
