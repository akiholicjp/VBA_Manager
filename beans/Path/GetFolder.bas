' VBA: Import ../Wrapper/_SHBrowseForFolder.bas
' VBA: Import ../Wrapper/_SHGetPathFromIDList.bas
' VBA: Import ../Wrapper/_SendMessage.bas
' VBA: Import ../Wrapper/_FARPROC.bas


Function GetFolder(Optional ByVal strTitle As String = "Select Folder", Optional ByRef strPath As String) As String
   Dim bif As BROWSEINFO
   Dim pidl As Long
   Dim s As String
   On Error GoTo ErrGetFolder
   With bif
      .pidlRoot = &H0 'デスクトップ
      .ulFlags = &H1 'フォルダのみ選択可能
      .lpszTitle = strTitle
      .lpfn = FARPROC(AddressOf BrowseCallbackProc)
      If IsMissing(strPath) Then
         .lParam = CurDir & Chr(0)
      ElseIf Trim(strPath) = "" Then
         .lParam = CurDir & Chr(0)
      Else
         .lParam = strPath & Chr(0)
      End If
   End With
   strPath = ""
   pidl = SHBrowseForFolder(bif)
   If pidl <> 0 Then
      s = String$(256, vbNullChar)
      SHGetPathFromIDList pidl, s
      strPath = Left(s, InStr(s, vbNullChar) - 1)
   End If
   GetFolder = strPath
   Exit Function
ErrGetFolder:
   GetFolder = ""
End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
   Const BFFM_SETSELECTIONA = (&H400 + 102) ' WM_USER[&H400]
   Const BFFM_INITIALIZED = 1
   If uMsg = BFFM_INITIALIZED Then
      SendMessage hWnd, BFFM_SETSELECTIONA, 1, ByVal lpData
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTestGUI_beans_GetFolder()
   ' Usage
   Dim strPath As String, s As String
   strPath = ""
   s = GetFolder(strTitle:="TEST", strPath:=strPath)
   MsgBox strPath, Title:=s

   s = GetFolder(strTitle:="TEST", strPath:="c:\")
   MsgBox s, Title:=s

   s = GetFolder()
   MsgBox s, Title:=s
End Sub
