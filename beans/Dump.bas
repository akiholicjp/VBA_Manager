Function Dump(ByRef x As Variant, Optional ByVal WithPtr As Boolean = False) As String
   Dim dicContainer As Object
   Set dicContainer = CreateObject("Scripting.Dictionary")
   Dump = DumpSub(x, dicContainer, WithPtr)
   Set dicContainer = Nothing
End Function

Private Function DumpSub(ByRef x As Variant, ByRef dicContainer As Object, ByVal WithPtr As Boolean) As String
   Dim sType As String
   Dim d As String, t As String
   If IsObject(x) Then
      Select Case TypeName(x)
      Case "Dictionary":
         DumpSub = DictionaryToStr(x, dicContainer, WithPtr)
      Case "Collection":
         DumpSub = CollectionToStr(x, dicContainer, WithPtr)
      Case Else
         DumpSub = ObjectToStr(x)
      End Select
      Exit Function
   End If

   sType = TypeName(x)
   Select Case sType
   Case "Boolean":     DumpSub = x
   Case "Integer":     DumpSub = x & "%"
   Case "Long":        DumpSub = x & "&"
   #If VBA7 And Win64 Then
   Case "LongLong":    DumpSub = x & "^"
   #End If
   Case "Single":      DumpSub = x & "!"
   Case "Double":      DumpSub = x & "#"
   Case "Currency":    DumpSub = x & "@"
   Case "Byte":        DumpSub = "CByte(" & x & ")"
   Case "Decimal":     DumpSub = "CDec(" & x & ")"
   Case "Date":
      If Abs(x) >= 1 Then
         DumpSub = "#" & Format(x, "YYYY/MM/DD hh:mm:ss") & "#"
      Else
         DumpSub = "#" & Format(x, "hh:mm:ss") & "#"
      End If
   Case "String"
      If StrPtr(x) = 0 Then
         DumpSub = "<vbNullString>"
      Else
         DumpSub = """" & x & """"
      End If
   Case "Empty", "Null", "Nothing"
      DumpSub = "<" & sType & ">"
   Case "Error"
      If IsMissing(x) Then
         DumpSub = "<Missing>"
      Else
         DumpSub = "<Error>"
      End If
   Case "ErrObject"
      DumpSub = "<Err:" & x.Number & ">"
   Case "Unknown"
      DumpSub = "<unknown:" & sType & ">"
   Case Else
      If IsArray(x) Then
         DumpSub = ArrayToStr(x, dicContainer, WithPtr)
      Else
         DumpSub = ""
         Exit Function
      End If
   End Select
End Function

Private Function ObjectToStr(ByVal v As Variant) As String
   If IsObject(v) Then
      On Error GoTo ErrHandle
      ObjectToStr = v.ToStr()
      On Error GoTo 0
   Else
      ObjectToStr = v
   End If
   Exit Function
ErrHandle:
   ObjectToStr = "<obj:" & TypeName(v) & ">"
   Resume Next
End Function

Private Function DictionaryToStr(ByVal x As Variant, ByRef dicContainer As Object, ByVal WithPtr As Boolean) As String
   If dicContainer.Exists(ObjPtr(x)) Then
      If WithPtr Then
         DictionaryToStr = "<Dictionary:" & ObjPtr(x) & ">"
      Else
         DictionaryToStr = "<Dictionary>"
      End If
      Exit Function
   Else
      dicContainer.Add Key:=ObjPtr(x), Item:=x
   End If

   Dim s As String
   Dim i As Long
   Dim vKeys() As Variant, vItems() As Variant

   vKeys = x.Keys
   vItems = x.Items
   If WithPtr Then
      s = ObjPtr(x) & "{"
   Else
      s = "{"
   End If
   For i = 0 To x.Count - 1
      s = s & DumpSub(vKeys(i), dicContainer, WithPtr) & ":" & DumpSub(vItems(i), dicContainer, WithPtr)
      If i <> x.Count - 1 Then s = s & ","
   Next i
   s = s & "}"
   DictionaryToStr = s
End Function

Private Function CollectionToStr(ByVal x As Variant, ByRef dicContainer As Object, ByVal WithPtr As Boolean) As String
   If dicContainer.Exists(ObjPtr(x)) Then
      If WithPtr Then
         CollectionToStr = "<Collection:" & ObjPtr(x) & ">"
      Else
         CollectionToStr = "<Collection>"
      End If
      Exit Function
   Else
      dicContainer.Add Key:=ObjPtr(x), Item:=x
   End If

   Dim s As String
   Dim i As Long
   If WithPtr Then
      s = ObjPtr(x) & "["
   Else
      s = "["
   End If
   For i = 1 To x.Count
      s = s & DumpSub(x(i), dicContainer, WithPtr)
      If i <> x.Count Then s = s & ","
   Next i
   s = s & "]"
   CollectionToStr = s
End Function

Private Function ArrayToStr(ByVal x As Variant, ByRef dicContainer As Object, ByVal WithPtr As Boolean) As String
   Dim s As String
   Dim i As Long

   s = "("
   For i = LBound(x) To UBound(x)
      s = s & DumpSub(x(i), dicContainer, WithPtr)
      If i <> UBound(x) Then s = s & ","
   Next i
   s = s & ")"
   ArrayToStr = s
End Function
