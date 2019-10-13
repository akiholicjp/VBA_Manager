' VBA: Import ../GetFSO.bas

Function BuildPath(ByVal sBasePath As String, ByVal sPath As String)
   BuildPath = GetFSO().BuildPath(sBasePath, sPath)
End Function
