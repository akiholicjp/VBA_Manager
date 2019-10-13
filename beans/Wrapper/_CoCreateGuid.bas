Private Type GUID
 Data1 As Long
 Data2 As Integer
 Data3 As Integer
 Data4(7) As Byte
End Type

#If Win64 Then
Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
#Else
Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
#End If
