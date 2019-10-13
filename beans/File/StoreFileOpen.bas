' VBA: Import StoreFileCopy.bas
Sub StoreFileOpen(ByVal sPath As String, ByVal sTmpPath As String, ByVal bForce As Boolean)
   sFilename = StoreFileCopy(sPath, sTmpPath, bForce)
   If sFilename <> "" Then
      Call ShellExecute(sFilename, sDir:=sTmpPath)
   End If
End Sub
