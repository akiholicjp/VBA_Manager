#If VBA7 Then
#If Win64 Then
Public Declare PtrSafe _
Function rtcCallByName Lib "VBE7.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As LongPtr, _
    ByVal CallType As VbCallType, _
    ByRef Args() As Any, _
    Optional ByVal lcid As Long _
    ) As Variant
#Else
Public Declare _
Function rtcCallByName Lib "VBE7.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As Long, _
    ByVal CallType As VbCallType, _
    ByRef Args() As Any, _
    Optional ByVal lcid As Long _
    ) As Variant
#End If
#Else
Public Declare _
Function rtcCallByName Lib "VBE6.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As Long, _
    ByVal CallType As VbCallType, _
    ByRef Args() As Any, _
    Optional ByVal lcid As Long _
    ) As Variant
#End If
