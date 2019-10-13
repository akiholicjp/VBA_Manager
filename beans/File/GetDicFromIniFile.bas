' VBA: Import ../RegExe.bas
' VBA: Import ../RegTest.bas
' VBA: Import ../NewDic.bas

Function GetDicFromIniFile(ByVal sIniFile As String, Optional ByVal iDuplicateMode As Long = 0) As Object
   Dim regResults As Object
   Dim dF As Long
   Dim sLine As String
   Dim sSec As String
   Dim oDic As Object
   Dim oSec As Object
   Dim sKey As String, sVal As String

   dF = 0
   On Error GoTo Err_Proc

   dF = FreeFile
   Open sIniFile For Input As #dF

   Set oDic = NewDic()
   Do Until EOF(dF)
      Line Input #dF, sLine: sLine = Trim(sLine)
      If sLine = "" Then GoTo CONTINUE

      If RegTest(sLine, "^[;'#].*") Then
         GoTo CONTINUE
      ElseIf RegExe(sLine, "^\[([^\]]+)\]$", regResults) Then
         sSec = Trim(regResults(1))
         If oDic.Exists(sSec) Then
            Set oSec = oDic(sSec)
         Else
            Set oSec = NewDic()
            oDic.Add Key:=sSec, Item:=oSec
         End If
         GoTo CONTINUE
      ElseIf RegExe(sLine, "^([^=\s]+)\s*=(.*)$", regResults) Then
         If oSec Is Nothing Then
            sSec = ""
            Set oSec = NewDic()
            oDic.Add Key:=sSec, Item:=oSec
         End If
         sKey = regResults(1)
         sVal = Trim(regResults(2))
         If oSec.Exists(sKey) Then
            Select Case iDuplicateMode
            Case 2 ' IniDupTakeLast
               oSec.Remove sKey
               oSec.Add Key:=sKey, Item:=sVal
            Case 1 ' IniDupTakeFirst
               ' Do Nothing
            Case Else '0 IniDupErr
               Err.Raise Number:=1000, Description:="Section["& sSec & "]ÇÃKey[" & sKey & "]Ç™èdï°ÇµÇƒÇ¢Ç‹Ç∑ÅB"
            End Select
         Else
            oSec.Add Key:=sKey, Item:=sVal
         End If
      Else
         If Not oSec.Exists("") Then
            oSec.Add Key:="", Item:=New Collection
         End If
         oSec("").Add sLine
      End If

CONTINUE:
   Loop

Exit_Proc:
   If dF <> 0 Then Close #dF
   Set GetDicFromIniFile = oDic
   Exit Function
Err_Proc:
   Set oDic = Nothing
   Resume Exit_Proc
End Function
