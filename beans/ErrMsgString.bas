Function ErrMsgString(ByVal sErr As String, Optional ByVal sProc As String = "", Optional ByVal sModule As String = "", Optional bErr As Boolean = False, Optional bErl As Boolean = False) As String
   If bErr And Err.Number <> 0 Then
      sErr = sErr & vbCrLf & "> " & "[ERR#" & Str(Err.Number) & "] " & Err.Description & "<"
   End If
   If (sProc <> "") Or (sModule <> "") Then
      sErr = sErr & vbCrLf & ">> occurred"
      If (bErl And VBA.Erl <> 0) Then
         sErr = sErr & " @ " & VBA.Erl & " in [" & sProc & "] on [" & sModule & "]"
      Else
         sErr = sErr & " in [" & sProc & "] on [" & sModule & "]"
      End If
      If bErr And Err.Number <> 0 Then
         sErr = sErr & " of [" & Err.Source &"]"
      End If
      sErr = sErr & " <<"
   End If
   ErrMsgString = sErr
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ErrMsgString()
   Dim v As Variant
   Err.Clear

   assert _
   "TEST", _
   ErrMsgString("TEST")

   assert _
   "TEST" & vbCrLf & ">> occurred in [A] on [] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A")
   assert _
   "TEST" & vbCrLf & ">> occurred in [A] on [X] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X")
   assert _
   "TEST" & vbCrLf & ">> occurred in [A] on [X] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True)
   assert _
   "TEST" & vbCrLf & ">> occurred in [A] on [X] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True, bErl:=True)
   ' assert _
   ' "TEST" & vbCrLf & "> [ERR# 424] オブジェクトが必要です。<" & vbCrLf & ">> occurred in [A] on [X] of [VBAProject] <<", _
   ' ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True)
   ' assert _
   ' "TEST" & vbCrLf & "> [ERR# 424] オブジェクトが必要です。<" & vbCrLf & ">> occurred in [A] on [X] of [VBAProject] <<", _
   ' ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True, bErl:=True)
   On Error Resume Next
   assert _
   "TEST", _
   ErrMsgString("TEST")
   assert _
   "TEST" & vbCrLf & ">> occurred in [A] on [] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A")
   assert _
   "TEST" & vbCrLf & ">> occurred in [A] on [X] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X")
   assert _
   "TEST" & vbCrLf & ">> occurred in [A] on [X] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True)
   assert _
   "TEST" & vbCrLf & ">> occurred in [A] on [X] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True, bErl:=True)
   On Error GoTo ErrHandle
   v = Nothing
   On Error GoTo ErrHandle2
99 v = Nothing
   Exit Sub
ErrHandle:
   assert _
   "TEST" & vbCrLf & "> [ERR# 91] オブジェクト変数または With ブロック変数が設定されていません。<" & vbCrLf & ">> occurred in [A] on [X] of [VBAProject] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True)
   Resume Next
ErrHandle2:
   assert _
   "TEST" & vbCrLf & "> [ERR# 91] オブジェクト変数または With ブロック変数が設定されていません。<" & vbCrLf & ">> occurred @ 99 in [A] on [X] of [VBAProject] <<", _
   ErrMsgString(sErr:="TEST", sProc:="A", sModule:="X", bErr:=True, bErl:=True)
   Resume Next
End Sub
