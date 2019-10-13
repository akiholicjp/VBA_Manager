Type BROWSEINFO
   hWndOwner As Long
   pidlRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfn As LongPtr
   lParam As String
   iImage As Long
End Type

#If Win64 Then
Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
#Else
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
#End If
