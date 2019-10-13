Function VBA_ExportCodeModule(ByRef oComp As Object, ByVal sDir As String, Optional ByVal IgnoreBlankDcm As Boolean = True) As Boolean
   Dim sName As String
   VBA_ExportCodeModule = False
   Select Case oComp.Type
   Case 1 ' Module.Standard
      sName = oComp.Name & ".bas"
   Case 2 ' Module.Class
      sName = oComp.Name & ".cls"
   Case 3 ' Module.Forms
      sName = oComp.Name & ".frm"
   Case 100 ' Module.Document
      If (Not IgnoreBlankDcm) Or (oComp.CodeModule.CountOfLines > 0) Then
         sName = oComp.Name & ".dcm"
      Else
         sName = ""
      End If
   End Select
   If sName <> "" Then
      sDir = Trim(sDir)
      If Right(sDir, 1) = "/" Then
         oComp.Export sDir & sName
      Else
         oComp.Export sDir & "/" & sName
      End If
   End If
   VBA_ExportCodeModule = True
End Function
