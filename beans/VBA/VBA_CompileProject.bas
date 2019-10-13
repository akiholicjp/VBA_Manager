Public Sub VBA_CompileProject()
#If ACCESS_VBA <> 1 Then
   Dim oCtrl As Object
   Set oCtrl = Nothing
   On Error Resume Next
   Set oCtrl = Application.VBE.CommandBars.FindControl(, 578)
   On Error GoTo 0
   If Not oCtrl Is Nothing Then
      If oCtrl.Enabled = True Then oCtrl.Execute
   End If
#End If
End Sub
