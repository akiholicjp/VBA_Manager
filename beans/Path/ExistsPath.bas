' VBA: Import ../GetFSO.bas

Function ExistsPath(ByVal sPath As String)
   With GetFSO()
      ExistsPath = .FileExists(sPath) Or .FolderExists(sPath)
   End With
End Function
