' VBA: Import ../GetFSO.bas

' Microsoft Internet Controls,Microsoft HTML Object Library
Function CoolCopy(sBaseFolder As String, sFolder As String, sFilename As String, sCopyTo As String, bForce As Boolean) As Boolean
   Dim oShell As Object
   Dim sCopyFrom As String
   Dim sTarget As String
   Set oShell = CreateObject("Shell.Application")
   With GetFSO()
      If .FileExists(sCopyTo + "\" + sFilename) Then
         Call .DeleteFile(sCopyTo + "\" + sFilename, bForce)
      End If
      sCopyFrom = sBaseFolder + "\" + sFolder
      sTarget = sFilename
      If .FolderExists(sCopyFrom) Then
         Call oShell.Namespace((sCopyTo)).CopyHere(oShell.Namespace((sCopyFrom)).Items.Item((sTarget)), 4 + 16 + 512)
      End If
      If Not .FileExists(sCopyTo + "\" + sFilename) Then
         sCopyFrom = sBaseFolder + "\" + sFolder + ".zip"
         If .FileExists(sCopyFrom) Then
            sTarget = sFilename
            Call oShell.Namespace((sCopyTo)).CopyHere(oShell.Namespace((sCopyFrom)).Items.Item((sTarget)), 4 + 16 + 512)
            If Not .FileExists(sCopyTo + "\" + sFilename) Then
               sTarget = sFolder + "\" + sFilename
               Call oShell.Namespace((sCopyTo)).CopyHere(oShell.Namespace((sCopyFrom)).Items.Item((sTarget)), 4 + 16 + 512)
            End If
         End If
      End If

      If .FileExists(sCopyTo + "\" + sFilename) Then
         CoolCopy = True
      Else
         CoolCopy = False
      End If
      If Not oShell Is Nothing Then
         Set oShell = Nothing
      End If
   End With
End Function
