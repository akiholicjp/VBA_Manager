' VBA: Import GetFSO.bas: Private
' VBA: Import Path/GetOwnPath.bas: Private
' VBA: Import Dump.bas: Private

Sub LogEasy(ByVal str As String, Optional ByVal sFile As String = "", Optional ByVal bDate As Boolean = False, Optional ByVal bClose As Boolean = False, Optional dump As Variant)
   Static oFile As Object
   Dim s As String
   If oFile Is Nothing Then
      With GetFSO()
         If sFile = "" Then sFile = "ezy_log.log"
         If bDate Then sFile = .GetBaseName(sFile) & "_" & Format(Now, "yyyymmss_hhmmss") & "." & .GetExtensionName(sFile)
         sFile = .BuildPath(GetOwnPath(), sFile)
         Set oFile = .OpenTextFile(Filename:=sFile, IOMode:=8, Create:=True) ' AppendMode
      End With
   End If

   s = Format(Now, "yyyy/mm/dd hh:mm:ss") & ", " & str
   If Not IsMissing(dump) Then s = s & ", " & Dump(dump)
   s = s & "."

   oFile.WriteLine s

   If bClose Then
      oFile.Close
      Set oFile = Nothing
   End If
End Sub
