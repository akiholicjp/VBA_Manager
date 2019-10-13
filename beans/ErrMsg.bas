' VBA: Import ErrMsgString.bas

Sub ErrMsg(ByVal sErr As String, Optional ByVal sProc As String = "", Optional ByVal sModule As String = "", Optional bErr As Boolean = False, Optional bErl As Boolean = False, Optional bMsgBox As Boolean = False)
   sErr = ErrMsgString(sErr, sProc, sModule, bErr, bErl)
   Debug.Print sErr
   If bMsgBox Then
      If bErr And Err.Number <> 0 Then
         MsgBox sErr, vbCritical + vbOKOnly, "Message", Err.Helpfile, Err.HelpContext
      Else
         MsgBox sErr, vbCritical + vbOKOnly, "Message"
      End If
   End If
End Sub

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTestGUI_beans_ErrMsg()
   ' Usage

   Dim v As Variant
   Call ErrMsg("TEST")
   Call ErrMsg(sErr:="TEST", sProc:="A")
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X")
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True)
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True, bErl:=True)
   On Error Resume Next
   Call ErrMsg("TEST")
   Call ErrMsg(sErr:="TEST", sProc:="A")
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X")
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True)
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True, bErl:=True)
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True, bErl:=True, bMsgBox:=True)
   On Error GoTo ErrHandle
   v = Nothing
   On Error GoTo ErrHandle2
   99 v = Nothing
   Exit Sub
ErrHandle:
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True)
   Resume Next
ErrHandle2:
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True, bErl:=True)
   Call ErrMsg(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True, bErl:=True, bMsgBox:=True)
   Resume Next
End Sub
