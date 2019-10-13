Function VBA_GetComponents() As Variant
#If ACCESS_VBA <> 1 Then
   Set VBA_GetComponents = ThisWorkbook.VBProject.VBComponents
#Else
   Set VBA_GetComponents = Application.VBE.ActiveVBProject.VBComponents
#End If
End Function
