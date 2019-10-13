Public Function parse(ByVal strFilename As String, ByVal strFilename2 As String) As Boolean
   Dim oParse As New Collection
   Dim fp As Long
   Dim s As String
   Dim c As String
   Dim i As Long
   Dim iQuote As Long, bComment As Boolean, bContLine As Boolean, bNormalCont As Boolean
   Dim sLine As String
   ' Dim reg As Object
   ' Dim bHead As Boolean, bHeadHead As Boolean
   ' Set reg = CreateObject("VBScript.RegExp")
   ' reg.Global = True
   ' reg.MultiLine = True
   ' reg.IgnoreCase = False

   s = ""

   iQuote = -1
   bComment = False
   bContLine = False
   bNormalCont = False

   fp = FreeFile
   Open strFilename For Input As fp
   Do Until EOF(fp)
      Line Input #fp, sLine

      For i = 1 To Len(sLine)
         c = Mid(sLine, i, 1)
         If bNormalCont Then
            s = s & c
         ElseIf (iQuote < 0 And c = "'") Or bComment Then
            If s <> "" Then oParse.Add Array("Normal", s, False): s = ""
            If bContLine Then
               oParse.Add Array("CommentCont", sLine, True)
            Else
               oParse.Add Array("CommentBegin", Mid(sLine, i + 1), True)
            End If
            If Right(Trim(sLine), 1) = "_" Then
               bComment = True
            Else
               bComment = True
            End If
            Exit For
         ElseIf iQuote > 0 Then
            If c = """" Then
               oParse.Add Array("String", Mid(sLine, iQuote + 1, (i - iQuote) - 1), False)
               iQuote = -1
            End If
         ElseIf c = """" Then
            If s <> "" Then oParse.Add Array("Normal", s, False): s = ""
            iQuote = i
         ElseIf c = "_" Then
            If i = 1 Or Trim(Mid(sLine, i - 1, 1)) = "" Then
               If s <> "" Then oParse.Add Array("Normal", s, False): s = ""
               oParse.Add Array("Cont", c, True)
               bNormalCont = True
            Else
               s = s & c
            End If
         Else
            s = s & c
         End If
      Next

      If bComment Then
         ' Pass
      ElseIf iQuote > 0 Then
         oParse.Add Array("ErrorNotQuoteEnd", Mid(sLine, iQuote + 1, True))
      ElseIf bNormalCont Then
         If Trim(s) <> "" Then
            oParse.Add Array("ErrorContEnd", s, True): s = ""
         Else
            If Len(s) > 0 Then
               oParse.Add Array("ContEnd", s, True): s = ""
            End If
         End If
      Else
         oParse.Add Array("Normal", s, True): s = ""
      End If

      If Right(Trim(sLine), 1) = "_" Then
         iQuote = -1
         ' bComment: AsIs
         bContLine = True
         bNormalCont = False
      Else
         iQuote = -1
         bComment = False
         bContLine = False
         bNormalCont = False
      End If
   Loop
   Close fp

   fp = FreeFile
   Open strFilename2 For Output As fp
   Dim v As Variant, s2 As String
   s = ""
   s2 = ""
   For Each v In oParse
      s2 = s2 & v(0) & ":"
      Select Case v(0)
      Case "Normal": s = s & v(1)
      Case "CommentCont": s = s & v(1)
      Case "CommentBegin": s = s & "'" & v(1)
      Case "String": s = s & """" & v(1) & """"
      Case "ErrorNotQuoteEnd": s = s & """" & v(1)
         MsgBox "ErrorNotQuotedEnd"
      Case "Cont": s = s & v(1)
      Case "ContEnd": s = s & v(1)
      Case "ErrorContEnd": s = s & v(1)
         MsgBox "ErrorContEnd"
      End Select
      If v(2) Then
         Print #fp, s
         Debug.Print s2
         s = "": s2 = ""
      End If
   Next v
   Close fp

   ' bHead = True
   ' bHeadHead = False
'       If bHead Then
'          If bHeadHead Then
'             reg.Pattern = "^\s*END\s*"
'             If reg.Test(s) Then
'                bHeadHead = False
'             Else
'                GoTo Continue
'             End If
'          End If
'          reg.Pattern = "^\s*BEGIN\s*.*?END\s*"
'          If reg.Test(s) Then s = reg.Replace(s, "")
'          reg.Pattern = "^\s*Attribute\s*"
'          If reg.Test(s) Then GoTo Continue
'          reg.Pattern = "^\s*(Private\s+|Public\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property)\s*"
'          If reg.Test(s) Then
'             bHead = False
'          Else
'             If Trim(s) <> "" Then
'                oHeadList.Add s
'             End If
'          End If
'       End If
'       If Not bHead Then
'          If Trim(s) <> "" Then
'             oMainList.Add s
'          End If
'       End If
'       ' reg.Pattern = "^\s*'\s*BeanImport\s+([^\s:]+)\s*:?\s*([^\s]+)?\s*"
'       ' If reg.Test(sLine) Then
'       ' End If
' Continue:
'    Loop
'    Close fp
End Function
