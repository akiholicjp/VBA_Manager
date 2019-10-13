' VBA: Import CoolCopy.bas
' VBA: Import ../GetFSO.bas
' VBA: Import ../Win/ShellExecute.bas

Function StoreFileCopy(ByVal sPath As String, ByVal sCopyDir As String, ByVal bForce As Boolean) As String
   Dim pos As Long
   Dim sBaseDirname As String, sDirName As String, sFilename As String
   If Not GetFSO().FolderExists(sCopyDir) Then
      MsgBox "[" & sCopyDir & "] is not exists."
      StoreFileCopy = ""
      Exit Function
   End If
   With CreateObject("Shell.Application")
      pos = InStr(sPath, "\")
      sBaseDirname = Left(sPath, pos - 1)
      sPath = Mid(sPath, pos + 1)
      pos = InStrRev(sPath, "\")
      sDirName = Left(sPath, pos - 1)
      sFilename = Mid(sPath, pos + 1)
      Call CoolCopy(ThisWorkbook.Path & "\" & sBaseDirname, sDirName, sFilename, sCopyDir, bForce)
      StoreFileCopy = sFilename
   End With
End Function
