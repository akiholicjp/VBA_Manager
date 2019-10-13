Function VBA_GetModuleType(ByRef oMod As Object, Optional ByVal IgnoreBlankDcm As Boolean = False) As String
   Dim sType As String
   Select Case oMod.Type
   Case 1 ' Module.Standard
      sType = "bas"
   Case 2 ' Module.Class
      sType = "cls"
   Case 3 ' Module.Forms
      sType = "frm"
   Case 100 ' Module.Document
      If Not IgnoreBlankDcm Or oMod.CodeModule.CountOfLines > 0 Then
         sType = "dcm"
      Else
         sType = ""
      End If
   End Select
   VBA_GetModuleType = sType
End Function
