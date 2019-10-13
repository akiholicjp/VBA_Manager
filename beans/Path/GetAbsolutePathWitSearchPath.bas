' VBA: Import ../GetFSO.bas
' VBA: Import IsAbsolutePath.bas
' VBA: Import GetBaseFileName.bas

Function GetAbsolutePathWithSearchPath(ByVal pathFile As String, ByVal aSearchPath As Variant, Optional ByRef sDir As String, Optional ByRef sFile As String) As String
   Dim sPath As String
   Dim b As Boolean
   Dim vDir As Variant
   With GetFSO()
      If IsAbsolutePath(pathFile) Then
         If .FileExists(pathFile) Then
            sPath = pathFile
            sDir = .GetParentFolderName(pathFile)
            sFile = GetBaseFileName(pathFile)
            b = True
         ElseIf .FolderExists(pathFile) Then
            sPath = pathFile
            sDir = .GetParentFolderName(pathFile)
            sFile = ""
            b = True
         Else
            sPath = ""
            sDir = ""
            sFile = ""
            b = False
         End If
      Else
         b = False
         For Each vDir In aSearchPath
            sPath = .GetAbsolutePathName(.BuildPath(vDir, pathFile))
            If .FileExists(sPath) Then
               sDir = .GetParentFolderName(sPath)
               sFile = GetBaseFileName(sPath)
               b = True
               Exit For
            ElseIf .FolderExists(sPath) Then
               sDir = .GetParentFolderName(sPath)
               sFile = ""
               b = True
               Exit For
            End If
         Next vDir
      End If
   End With
   If b Then
      GetAbsolutePathWithSearchPath = sPath
   Else
      GetAbsolutePathWithSearchPath = ""
   End If
End Function

' =================== VBA: TEST: Begin ===================

' VBA: Import GetOwnPath.bas
' VBA: Import SetCurrentPath.bas

Public Sub xUnitTest_beans_GetAbsolutePathWithSearchPath()
   Dim o As Object
   Set o = New Collection
   o.Add "."
   assert StrConv(GetOwnPath(), vbUpperCase), StrConv(GetAbsolutePathWithSearchPath(GetOwnPath(), Array(".")), vbUpperCase)
   SetCurrentPath(GetOwnPath())
   assert StrConv(GetOwnPath(), vbUpperCase), StrConv(GetAbsolutePathWithSearchPath(".", o), vbUpperCase)
End Sub
