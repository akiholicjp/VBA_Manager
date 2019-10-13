Sub SSS_ERR_TRAP_CALL(sProc As String, sModule As String)
    Dim sErr As String
    If Err.Number <> 0 Then
       sErr = sErr & vbCrLf & "> " & "[ERR#" & str(Err.Number) & "] " & Err.Description & "<"
    End If
    If (sProc <> "") Or (sModule <> "") Then
       sErr = sErr & vbCrLf & ">> occurred"
       If VBA.Erl <> 0 Then
          sErr = sErr & " @ " & VBA.Erl & " in [" & sProc & "] on [" & sModule & "]"
       Else
          sErr = sErr & " in [" & sProc & "] on [" & sModule & "]"
       End If
       If Err.Number <> 0 Then
          sErr = sErr & " of [" & Err.Source & "]"
       End If
       sErr = sErr & " <<"
    End If
    Debug.Print sErr
    MsgBox sErr
End Sub
