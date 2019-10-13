' VBA: Import ../GetShell.bas
Function SetCurrentPath(ByVal sDir As String) As String
   GetShell().CurrentDirectory = sDir
End Function
