Function VBA_GetProject() As Variant
#If ACCESS_VBA <> 1 Then
   Set VBA_GetProject = ThisWorkbook.VBProject
#Else
   Set VBA_GetProject = Application.VBE.ActiveVBProject
#End If
End Function
