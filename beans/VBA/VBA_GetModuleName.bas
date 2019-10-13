Function VBA_GetModuleName(ByRef oComp As Object, Optional ByVal IgnoreBlankDcm As Boolean = False) As String
   Dim sName As String
   Select Case oComp.Type
   Case 1 ' Module.Standard
      sName = oComp.Name & ".bas"
   Case 2 ' Module.Class
      sName = oComp.Name & ".cls"
   Case 3 ' Module.Forms
      sName = oComp.Name & ".frm"
   Case 100 ' Module.Document
      If Not IgnoreBlankDcm Or oComp.CodeModule.CountOfLines > 0 Then
         sName = oComp.Name & ".dcm"
      Else
         sName = ""
      End If
   End Select
   VBA_GetModuleName = sName
End Function
