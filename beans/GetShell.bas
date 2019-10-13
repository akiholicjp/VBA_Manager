Function GetShell() As Object
   Static G_Shell As Object
   If G_Shell Is Nothing Then Set G_Shell = CreateObject("WScript.Shell")
   Set GetShell = G_Shell
End Function
