Private Function EditCodeModule_Autocontrols( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   ' 標準モジュールのPREFIX
   Const C_CTRL_MENU = "MENU"
   Const C_CTRL_COMMAND = "COMMAND"
   Const C_CTRL_CELL = "CELL"
   Const C_CTRL_ROW = "ROW"
   Const C_CTRL_COLUMN = "COLUMN"
   Const C_CTRL_SHEET = "SHEET"

   ' 追加メニューのタイプ
   Const C_TYPE_MENU = "MENU"
   Const C_TYPE_BUTTON = "BUTTON"
   Const C_TYPE_EDIT = "EDIT"
   Const C_TYPE_HOOK = "HOOK"

   Dim S_CTRL_REGEXP As String
   Dim S_TYPE_REGEXP As String
   S_CTRL_REGEXP = C_CTRL_MENU & "|" & C_CTRL_COMMAND & "|" & C_CTRL_CELL & "|" & C_CTRL_ROW & "|" & C_CTRL_COLUMN & "|" & C_CTRL_SHEET
   S_TYPE_REGEXP = C_TYPE_MENU & "|" & C_TYPE_BUTTON & "|" & C_TYPE_EDIT & "|" & C_TYPE_HOOK

   Dim dicCtrl As Object
   Dim regResults As Object
   Dim v As Variant
   Dim sProcName As String
   Dim sGroupName As String
   If propState("CODE_STATE") = "<INIT>" Then
      If RegExe(propMod("NAME"), "^\s*(" & S_CTRL_REGEXP & ")([0-9]*)_(.*)\s*$", regResults) Then
         If Not propGlobal.Exists("AUTOCONTROLS") Then propGlobal.Add Key:="AUTOCONTROLS", Item:=NewDic()
         sGroupName = propMod("NAME")
         If Not propGlobal("AUTOCONTROLS").Exists(sGroupName) Then
            propGlobal("AUTOCONTROLS").Add Key:=sGroupName, Item:=DicProp("ID", sGroupName, "Type", regResults(1), "No", CLng(regResults(2)), "Caption", Trim(regResults(3)), "Controls", NewDic())
            propState.Add Key:="Autocontrols", Item:=propGlobal("AUTOCONTROLS")(sGroupName)("Controls")
            propState("Autocontrols_cont") = False
         End If
      End If
   End If

   If Trim(sComment) <> "" Then
      If RegExe(sComment, "^'[-=\s]*VBA\s*:\s*AutoControls Definition\s*:\s*(" & S_CTRL_REGEXP & ")([0-9]*)\s*:(.*?)\s*[-=\s]*$", regResults) Then
         If Not propGlobal.Exists("AUTOCONTROLS") Then propGlobal.Add Key:="AUTOCONTROLS", Item:=NewDic()
         sGroupName = propMod("NAME") & "_" & CStr(propGlobal("AUTOCONTROLS").Count)
         If Not propGlobal("AUTOCONTROLS").Exists(sGroupName) Then
            propGlobal("AUTOCONTROLS").Add Key:=sGroupName, Item:=DicProp("ID", sGroupName, "Type", regResults(1), "No", CLng(regResults(2)), "Caption", Trim(regResults(3)), "Controls", NewDic())
            If Not propState.Exists("Autocontrols") Then
               propState.Add Key:="Autocontrols", Item:=propGlobal("AUTOCONTROLS")(sGroupName)("Controls")
            Else
               Set propState("Autocontrols") = propGlobal("AUTOCONTROLS")(sGroupName)("Controls")
            End If
            propState("Autocontrols_cont") = False
         End If
      End If
   End If

   If propState("CODE_STATE") Like "<PROC:*:True>" And propState.Exists("Autocontrols") Then
      Set dicCtrl = propState("Autocontrols")
      Set v = SplitEx(propState("CODE_STATE"), Delim:=":"): sProcName = v(2)
      If Not dicCtrl.Exists(sProcName) Then
         If RegExe(sProcName, "(" & S_TYPE_REGEXP & ")([0-9]*)_(.*)", regResults) Then
            dicCtrl.Add Key:=sProcName, Item:=DicProp("ID", sProcName, "Type", regResults(1), "No", CLng(regResults(2)), "Name", regResults(3), "OnAction", propMod("NAME") & "." & sProcName, "Attr", NewDic())
            propState("Autocontrols_cont") = True
         End If
      ElseIf propState("Autocontrols_cont") Then
         If sComment = "" Then
            propState("Autocontrols_cont") = False
         End If
         If RegExe(sComment, "^'\s*([A-Za-z_]+):(.*)\s*", regResults) Then
            dicCtrl(sProcName)("Attr").Add Key:=Trim(regResults(1)), Item:=Trim(regResults(2))
         End If
      End If
   End If
   EditCodeModule_Autocontrols = True
End Function

Private Function EditCodeModule_AutocontrolsInsert( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean

   Dim s As String
   Dim sBlank As String
   Dim dicModule As Variant, dicCtrl As Object
   Dim vMod As Variant, vPropMod As Variant, vCtrls As Variant, vPropCtrl As Variant, vPropAttr As Variant
   Dim o As Object, o2 As Object

   If RegTest(sComment, "^'[-=\s]+VBA\s*:\s*Autocontrols Definition[-=\s]*$") Then
      If Not propGlobal.Exists("AUTOCONTROLS") Then EditCodeModule_AutocontrolsInsert = True: Exit Function
      If propGlobal("AUTOCONTROLS").Count = 0 Then EditCodeModule_AutocontrolsInsert = True: Exit Function

      sBlank = vbCrLf & String(Len(sLine), " ")
      s = String(Len(sLine), " ") & "Set G_dicModules = CreateObject(""Scripting.Dictionary"")" & sBlank

      For Each vMod In IterateDictionary(propGlobal("AUTOCONTROLS"))
         Set o = New Collection
         For Each vPropMod In IterateDictionary(vMod(1))
            If vPropMod(0) = "Controls" Then o.Add """" & vPropMod(0) & """, DicProp()" Else o.Add """" & vPropMod(0) & """, """ & vPropMod(1) & """"
         Next vPropMod
         s = s & "G_dicModules.Add Key:=""" & vMod(0) & """, Item:=DicProp(" & Join(o, ", ") & ")" & sBlank
         s = s & "With G_dicModules(""" & vMod(0) & """)(""Controls"")" & sBlank
         For Each vCtrls In IterateDictionary(vMod(1)("Controls"))
            Set o = New Collection
            Set o2 = New Collection
            For Each vPropAttr In IterateDictionary(vCtrls(1)("Attr"))
               o2.Add """" & vPropAttr(0) & """, """ & vPropAttr(1) & """"
            Next vPropAttr
            For Each vPropCtrl In IterateDictionary(vCtrls(1))
               If vPropCtrl(0) = "Attr" Then o.Add """" & vPropCtrl(0) & """, DicProp(" & Join(o2, ", ") & ")" Else o.Add """" & vPropCtrl(0) & """, """ & vPropCtrl(1) & """"
            Next vPropCtrl
            s = s & "   " & ".Add Key:=""" & vCtrls(0) & """, Item:=DicProp(" & Join(o, ", ") & ")" & sBlank
         Next vCtrls
         s = s & "End With" & sBlank
      Next vMod

      sLine = s
   End If
   EditCodeModule_AutocontrolsInsert = True
End Function
