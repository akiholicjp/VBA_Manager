Attribute VB_Name = "Z_VBA_Manager"
Option Explicit
Private Const C_SHORTCUT_CLEAR = "C"
Private Const C_SHORTCUT_RELOAD = "L"
Private Const C_SHORTCUT_COMPILE = "X"
Private Const C_CONF_INIFILE = "VBA_Manager"
Private Const C_MOD_VBAMANAGER = "Z_VBA_Manager"
Private Const C_MOD_UNMANAGED = "Zc_*"
Private G_CurDirStack As New Collection
Public Sub Step0_clearModules()
   Call clearModules
End Sub
Public Sub Step1_LoadModules()
   Call loadModules
   Call VBA_CompileProject
End Sub
Public Sub Step2_compileModules()
   Call VBA_CompileProject
End Sub
Public Sub Step3_AddDebugCode()
   Dim o As Object
   For Each o In GetManagedModules(VBA_GetComponents())
      Call VBA_AddErrorTrap(o)
      Call VBA_AddLineNumbers(o)
   Next o
End Sub
Public Sub Step4_Release()
   Call VBA_RemoveCodeModule(VBA_GetComponents(), VBA_GetModule(VBA_GetComponents(), "Z_VBA_Manager"))
End Sub
Public Sub RemoveLineNumbers()
   Dim o As Object
   For Each o In GetManagedModules(VBA_GetComponents())
      Call VBA_RemoveLineNumbers(o)
   Next o
End Sub
Public Sub AddLineNumbers()
   Dim o As Object
   For Each o In GetManagedModules(VBA_GetComponents())
      Call VBA_AddLineNumbers(o)
   Next o
End Sub
Public Sub clearModules()
   Dim o As Object
   For Each o In GetManagedModules(VBA_GetComponents())
      Call VBA_RemoveCodeModule(VBA_GetComponents(), o)
      DoEvents
   Next o
End Sub
Public Sub exportModulesAndList(Optional ByVal sProcList As String = "proclist")
   Dim v As Variant
   Dim oProcDic As Object
   Dim sName As String
   Dim oMod As Object
   Dim fp1 As Long
   Dim fp2 As Long
   fp1 = FreeFile
   Open GetOwnPath() & "/src/" & C_CONF_INIFILE & "-" & ThisWorkbook.Name & ".txt" For Output As fp1
   fp2 = FreeFile
   Open GetOwnPath() & "/" & sProcList & "-" & ThisWorkbook.Name & ".txt" For Output As fp2
   For Each oMod In GetManagedModules(VBA_GetComponents())
      Call VBA_ExportCodeModule(oMod, GetOwnPath() & "/src/", IgnoreBlankDcm:=True)
      sName = VBA_GetModuleName(oMod, IgnoreBlankDcm:=True)
      If sName <> "" Then
         Print #fp1, "." & "/src/" & sName
         Set oProcDic = VBA_GetDicProc(oMod)
         For Each v In oProcDic.Keys
            Print #fp2, sName & "," & v & "," & oProcDic(v)(0) & "," & oProcDic(v)(1)
         Next v
      End If
   Next oMod
   Close fp1
   Close fp2
End Sub
Public Sub loadModules()
   Dim b As Boolean
   Dim s As String
   Dim v As Variant
   Dim o As Object
   Dim msgError As String: msgError = ""
   Dim sPath As String
   Dim oModList As Object
   Dim arrayLoadModules As Object
   Dim oMod As Object
   Dim propGlobal As Object
   SetCurrentPath (GetOwnPath())
   For Each v In GetLibListArray()
      msgError = ""
      If loadIniFile(v, propGlobal, oModList, msgError) Then Exit For
   Next v
   If oModList Is Nothing Then msgError = msgError & "モジュールの構築リストが見つかりません"
   If msgError <> "" Then GoTo exit_loadModules
   If oModList.Count > 0 Then
      If Not BuildModules(oModList, msgError, propGlobal) Then
         msgError = msgError & "Modulesのインポートに失敗しました" & vbCrLf
      End If
   End If
   If msgError <> "" Then GoTo exit_loadModules
exit_loadModules:
   If msgError <> "" Then
      MsgBox msgError
   End If
End Sub
Public Sub saveModules()
   Call SaveThis
End Sub
Private Sub Auto_Open()
   If Workbooks.Count = 0 Then Workbooks.Add
   If C_SHORTCUT_CLEAR <> "" Then Application.MacroOptions Macro:="Z_VBA_Manager.Step0_clearModules", ShortcutKey:=C_SHORTCUT_CLEAR
   If C_SHORTCUT_RELOAD <> "" Then Application.MacroOptions Macro:="Z_VBA_Manager.Step1_loadModules", ShortcutKey:=C_SHORTCUT_RELOAD
   If C_SHORTCUT_COMPILE <> "" Then Application.MacroOptions Macro:="Z_VBA_Manager.Step2_compileModules", ShortcutKey:=C_SHORTCUT_COMPILE
End Sub
Private Sub Auto_Close()
   Application.MacroOptions Macro:="Z_VBA_Manager.Step0_clearModules", ShortcutKey:=""
   Application.MacroOptions Macro:="Z_VBA_Manager.Step1_loadModules", ShortcutKey:=""
   Application.MacroOptions Macro:="Z_VBA_Manager.Step2_compileModules", ShortcutKey:=""
End Sub
Private Function loadIniFile(ByVal sPath As String, ByRef propGlobal As Object, ByRef oModList As Object, ByRef msgError As String) As Boolean
   Dim dicIni As Object
   Dim dicSetup As Object
   Dim sDir As String, sFile As String
   Dim o As Object
   Dim v As Variant
   Dim b As Boolean
   b = False
   Set oModList = Nothing
   If propGlobal Is Nothing Then
      Set propGlobal = DicProp( _
         "DEBUG", False, _
         "TEST", False, _
         "IGNORE_BLANK", True, _
         "IGNORE_COMMENT", True, _
         "APPEND_INFO", True, _
         "SEARCH_PATH", Array(".") _
      )
   End If
   sPath = GetAbsolutePathWithSearchPath(sPath, Array("."), sDir, sFile)
   If sPath = "" Then
      msgError = msgError & "ファイルが見つかりません: " & sPath & vbCrLf
      GoTo Exit_Proc
   End If
   Set dicIni = GetDicFromIniFile(sFile)
   If dicIni Is Nothing Then
      msgError = msgError & "Iniファイルの読み込みに失敗しました: " & sPath & vbCrLf
      GoTo Exit_Proc
   End If
   If dicIni.Exists("SETUP") Then
      Set dicSetup = dicIni("SETUP")
      For Each v In Array("DEBUG", "TEST", "IGNORE_BLANK", "IGNORE_COMMENT", "APPEND_INFO")
         If dicSetup.Exists(v) Then propGlobal(v) = CBool(dicSetup(v))
      Next v
      For Each v In SplitEx(dicSetup("SEARCH_PATH"), Delim:=";")
         If o Is Nothing Then Set o = New Collection
         If CStr(v) Like "$*" Then
            o.Add GetAbsolutePath(Mid(ExpandEnvironmentStringsWhole(CStr(v)), 2))
         Else
            o.Add ExpandEnvironmentStringsWhole(CStr(v))
         End If
      Next v
      If Not o Is Nothing Then
         propGlobal.Remove "SEARCH_PATH"
         propGlobal.Add Key:="SEARCH_PATH", Item:=o
      End If
   End If
   If dicIni.Exists("BUILD") Then
      If dicIni("BUILD").Exists("") Then
         Set oModList = dicIni("BUILD")("")
         b = True
      End If
   End If
Exit_Proc:
   If Not b Then
      Set propGlobal = Nothing
      Set oModList = Nothing
   End If
   loadIniFile = b
End Function
Private Function IsManagedProcName(ByVal sName As String) As Boolean
   Dim b As Boolean
   Dim v As Variant
   b = (Not sName Like C_MOD_VBAMANAGER)
   If b Then
      For Each v In Array(C_MOD_UNMANAGED)
         If sName Like v Then
            b = False
            Exit For
         End If
      Next v
   End If
   IsManagedProcName = b
End Function
Private Function GetManagedModules(ByRef oComps As Object) As Collection
   Dim oCol As New Collection
   Dim o As Object
   For Each o In oComps
      If IsManagedProcName(o.Name) Then oCol.Add o
   Next o
   Set GetManagedModules = oCol
End Function
Private Function GetLibListArray() As Variant
   GetLibListArray = Array(C_CONF_INIFILE & ".ini")
End Function
Private Sub SaveThis()
#If ACCESS_VBA <> 1 Then
   ThisWorkbook.Save
#End If
End Sub
Private Function ParseLibList(ByRef oModList As Object, ByRef msgError As String, ByRef propGlobal As Object) As Object
   Dim regResults As Object
   Dim v As Variant
   Dim sLine As String
   Dim dicModules As Object
   Dim propMod As Object
   Set dicModules = CreateObject("Scripting.Dictionary")
   For Each v In oModList: sLine = Trim(v)
      If sLine Like "+*" Then
         sLine = Mid(Trim(sLine), 2)
         If GetFSO().GetExtensionName(sLine) = "" And GetFSO().GetParentFolderName(sLine) = "" Then
            Set propMod = addNewModule(dicModules, sLine, propGlobal, msgError)
            If propMod Is Nothing Then
               msgError = msgError & "ライブラリファイルの追加に失敗しました: " & sLine & vbCrLf
               Set dicModules = Nothing
               Exit For
            End If
         Else
            Set propMod = addLoadModule(dicModules, sLine, propGlobal, msgError)
            If propMod Is Nothing Then
               msgError = msgError & "ライブラリファイルの追加に失敗しました: " & sLine & vbCrLf
               Set dicModules = Nothing
               Exit For
            End If
         End If
      Else
         If Not ParseBeanImport(sLine, Nothing, False, propMod, propGlobal, msgError) Then
            msgError = msgError & "Beanファイルの追加に失敗しました: " & sLine & vbCrLf
            Set dicModules = Nothing
            Exit For
         End If
      End If
   Next v
   Set ParseLibList = dicModules
End Function
Private Function addLoadModule(ByRef dicModules As Object, ByVal sPath As String, ByRef propGlobal As Object, ByRef msgError As String) As Object
   Dim sName As String, sType As String
   Dim propMod As Object
   sName = GetFSO().GetBaseName(sPath)
   sType = StrConv(GetFSO().GetExtensionName(sPath), vbLowerCase)
   Set addLoadModule = Nothing
   If sName = "" Then
      msgError = msgError & "モジュール名称の指定がありません" & vbCrLf
      Exit Function
   ElseIf sName <> "" Then
      If dicModules.Exists(sName) Then
         msgError = msgError & "モジュール名が重複しています: " & sPath & vbCrLf
         Exit Function
      End If
   End If
   Set propMod = DicProp("NAME", sName, "TYPE", sType, "MODULE", Nothing, "BEANS", CreateObject("Scripting.Dictionary"), "OPTIONS", DicProp("Base", "", "Compare", "", "Explicit", False, "Private", False))
   If Not addBean(sPath, DicProp("Private", False), False, propMod, propGlobal, msgError) Then
      msgError = msgError & "モジュールの追加に失敗しました" & vbCrLf
      Exit Function
   End If
   dicModules.Add Key:=sName, Item:=propMod
   Set addLoadModule = propMod
End Function
Private Function addNewModule(ByRef dicModules As Object, ByVal sName As String, ByRef propGlobal As Object, ByRef msgError As String) As Object
   Set addNewModule = Nothing
   If sName = "" Then
      msgError = msgError & "モジュール名称の指定がありません" & vbCrLf
      Exit Function
   ElseIf sName <> "" Then
      If dicModules.Exists(sName) Then
         msgError = msgError & "モジュール名が重複しています: " & sName & vbCrLf
         Exit Function
      End If
   End If
   Set addNewModule = DicProp("NAME", sName, "TYPE", "new", "MODULE", Nothing, "BEANS", CreateObject("Scripting.Dictionary"), "OPTIONS", DicProp("Base", "", "Compare", "", "Explicit", False, "Private", False))
   dicModules.Add Key:=sName, Item:=addNewModule
End Function
Private Function addBean(ByVal sPath As String, ByRef dicOptions As Object, ByRef bDuplicate As Boolean, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim sFile As String, sDir As String, sName As String
   sPath = GetAbsolutePathWithSearchPath(sPath, propGlobal("SEARCH_PATH"), sDir, sFile)
   sName = StrConv(sFile, vbLowerCase)
   If propMod("BEANS").Exists(sName) Then
      If Not bDuplicate Then
         msgError = msgError & "モジュールファイルが重複しています: " & sPath & vbCrLf
         addBean = False
      Else
         addBean = True
      End If
      Exit Function
   End If
   If Not GetFSO().FileExists(sPath) Then
      msgError = msgError & "モジュールファイルが存在しません: " & sPath & vbCrLf
      addBean = False
      Exit Function
   End If
   propMod("BEANS").Add Key:=sName, Item:=DicProp("NAME", sName, "FILE", sFile, "DIR", sDir, "LOADED", False, "OPTIONS", dicOptions)
   addBean = True
End Function
Private Function ParseBeanImport(ByVal str As String, ByRef dicBeanParentOptions As Variant, ByVal bDuplicate As Boolean, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim regResults As Object
   Dim sPath As String, sOption As String
   Dim dicBeanOptions As Object
   Dim bPrivate As Boolean
   If Not RegExe(str, "^\s*([^\s:]+)\s*:?\s*(Private)?\s*", regResults) Then
      msgError = msgError & "Bean Import形式が不正です: " & str & vbCrLf
      ParseBeanImport = False
      Exit Function
   End If
   sPath = Trim(regResults(1))
   If Not dicBeanParentOptions Is Nothing Then bPrivate = dicBeanParentOptions("Private")
   If Not bPrivate Then bPrivate = Trim(regResults(2)) = "Private"
   Set dicBeanOptions = DicProp("Private", bPrivate)
   ParseBeanImport = addBean(sPath, dicBeanOptions, bDuplicate, propMod, propGlobal, msgError)
End Function
Private Function BuildModules(ByRef oModList As Object, ByRef msgError As String, ByRef propGlobal As Object) As Boolean
   Dim b As Boolean
   Dim bNewRead As Boolean
   Dim sPath As String
   Dim propBean As Object
   Dim propMod As Object
   Dim dicModules As Object
   Dim oMergeMod As Object
   Dim oModule As Object
   Dim v As Variant
   Dim vv As Variant
   Call PushCollection(G_CurDirStack, GetCurrentPath())
   Set dicModules = ParseLibList(oModList, msgError, propGlobal)
   If dicModules Is Nothing Then
      msgError = msgError & "モジュールリストファイルの読み込みに失敗しました " & vbCrLf
      BuildModules = False
      Exit Function
   End If
   For Each v In dicModules.Keys(): Set propMod = dicModules(v)
      Select Case propMod("TYPE")
      Case "new"
         Set oModule = VBA_GetComponents().Add(1)
         Set propMod("MODULE") = oModule
         oModule.Name = propMod("NAME")
      Case Else
         Set propBean = propMod("BEANS")(StrConv(propMod("NAME") & "." & propMod("TYPE"), vbLowerCase))
         Call SetCurrentPath(PushCollection(G_CurDirStack, propBean("DIR")))
         Set oModule = VBA_ImportCodeModule(VBA_GetComponents(), propBean("FILE"), msgError)
         If Not EditCodeModule(oModule, propBean, propMod, propGlobal, msgError) Then
            msgError = msgError & "モジュールの編集に失敗しました: " & propMod("FILE") & vbCrLf
            BuildModules = False
            Call SetCurrentPath(PopCollection(G_CurDirStack))
            Exit Function
         End If
         Set propMod("MODULE") = oModule
         propBean("LOADED") = True
         Call SetCurrentPath(PopCollection(G_CurDirStack))
      End Select
      Do While True
         For Each vv In propMod("BEANS").Items(): Set propBean = vv
            If Not propBean("LOADED") Then
               Call SetCurrentPath(PushCollection(G_CurDirStack, propBean("DIR")))
               Set oMergeMod = VBA_ImportCodeModule(VBA_GetComponents(), propBean("FILE"), msgError)
               If Not EditCodeModule(oMergeMod, propBean, propMod, propGlobal, msgError) Then
                  msgError = msgError & "Beansの編集に失敗しました: " & propBean("FILE") & vbCrLf
                  VBA_GetComponents().Remove oMergeMod
                  BuildModules = False
                  Call SetCurrentPath(PopCollection(G_CurDirStack))
                  Exit Function
               End If
               If propGlobal("APPEND_INFO") Then
                  Call VBA_MergeCodeModule(propMod("MODULE"), oMergeMod, "'==== " & propBean("FILE") & " ====")
               Else
                  Call VBA_MergeCodeModule(propMod("MODULE"), oMergeMod)
               End If
               VBA_GetComponents().Remove oMergeMod
               propBean("LOADED") = True
               Call SetCurrentPath(PopCollection(G_CurDirStack))
            End If
         Next vv
         bNewRead = False
         For Each vv In propMod("BEANS").Items(): Set propBean = vv
            If Not propBean("LOADED") Then bNewRead = True: Exit For
         Next vv
         If Not bNewRead Then Exit Do
      Loop
   Next v
   BuildModules = True
End Function

Private Function EditCodeModule(ByRef oMod As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim i As Long
   Dim sLine As String
   Dim iComment As Long, bContinue As Boolean
   Dim sComment As String
   Dim bRet As Boolean
   Dim bRemove As Boolean
   Dim iRemoveBeg As Long
   Dim iRemoveCnt As Long
   Dim propState As Object
   Dim iDeclLine As Long
   Dim iProcBodyLine As Long
   Dim sProcName As String
   Dim iType As Long
   Set propState = CreateObject("Scripting.Dictionary")
   bRet = True
   With oMod.CodeModule
      iComment = 0
      bContinue = False
      i = 1
      iRemoveBeg = 0
      iRemoveCnt = 0
      iDeclLine = VBA_GetCountOfDeclarationLines(oMod)
      propState("CODE_LINE") = 0
      propState("CODE_STATE") = "<INIT>"
      bRet = bRet And EditCodeModule_RemoveMode("TEST", "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_RemoveMode("DEBUG", "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Import(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_OptionStatements(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Global(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_GlobalInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Autocontrols(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutocontrolsInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutoMacro(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutoMacroInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_ConvPrivateScope(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      Do While i <= .CountOfLines
         bRemove = False
         sLine = .Lines(i, 1)
         If bContinue And iComment > 0 Then iComment = 1 Else iComment = VBA_GetCommentOfLine(sLine)
         If iComment > 0 Then sComment = Mid(sLine, iComment) Else sComment = ""
         If iComment > 0 Then sLine = Mid(sLine, 1, iComment - 1)
         propState("CODE_LINE") = i
         If i <= iDeclLine Then
            propState("CODE_STATE") = "<DECLARE>"
         ElseIf i > iDeclLine Then
            sProcName = .ProcOfLine(i, iType)
            iProcBodyLine = .ProcBodyLine(sProcName, iType)
            If iProcBodyLine > 0 Then
               propState("CODE_STATE") = "<PROC:" & sProcName & ":" & CStr(i >= iProcBodyLine) & ">"
            Else
               propState("CODE_STATE") = "<PROC:" & sProcName & ":>"
            End If
         End If
         bRet = bRet And EditCodeModule_RemoveMode("TEST", sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_RemoveMode("DEBUG", sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_Import(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_OptionStatements(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_Global(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_GlobalInsert(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_Autocontrols(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_AutocontrolsInsert(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_AutoMacro(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_AutoMacroInsert(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         bRet = bRet And EditCodeModule_ConvPrivateScope(Nothing, sLine, sComment, bRemove, propState, propBean, propMod, propGlobal, msgError)
         If bRemove Then GoTo CONTINUE
         If propGlobal("IGNORE_COMMENT") Then
            If propGlobal("APPEND_INFO") And sComment Like "*:[#][#][#]INSERTED[#][#][#]" Then
               sComment = Replace(sComment, ":###INSERTED###", "")
            Else
               sComment = ""
            End If
         End If
         If Trim(sLine & sComment) = "" And propGlobal("IGNORE_BLANK") Then
            bRemove = True
            GoTo CONTINUE
         End If
         If .Lines(i, 1) <> sLine & sComment Then .ReplaceLine i, sLine & sComment
CONTINUE:
         If Not bRet Then Exit Do
         If bRemove Then
            If iRemoveBeg = 0 Then
               iRemoveBeg = i
               iRemoveCnt = 1
            Else
               iRemoveCnt = iRemoveCnt + 1
            End If
         Else
            If iRemoveBeg <> 0 Then
               .DeleteLines iRemoveBeg, iRemoveCnt
               i = i - iRemoveCnt
               iRemoveBeg = 0
               iRemoveCnt = 0
               iDeclLine = VBA_GetCountOfDeclarationLines(oMod)
            End If
         End If
         bContinue = (Right(Trim(.Lines(i, 1)), 1) = "_")
         i = i + 1
      Loop
      If iRemoveBeg <> 0 Then .DeleteLines iRemoveBeg, iRemoveCnt
      propState("CODE_STATE") = "<END>"
      propState("CODE_LINE") = i
      bRet = bRet And EditCodeModule_RemoveMode("TEST", "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_RemoveMode("DEBUG", "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Import(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_OptionStatements(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Global(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_GlobalInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_Autocontrols(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutocontrolsInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutoMacro(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_AutoMacroInsert(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
      bRet = bRet And EditCodeModule_ConvPrivateScope(Nothing, "", "", bRemove, propState, propBean, propMod, propGlobal, msgError)
   End With
   EditCodeModule = bRet
End Function
Private Function EditCodeModule_RemoveMode( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim s As String: s = Trim(sComment)
   Dim sKeyword As String: sKeyword = opt
   Dim sKey As String: sKey = "REMOVE_MODE_" & sKeyword
   If Not propState.Exists(sKey) Then propState(sKey) = False
   If s <> "" Then
      If RegTest(s, "^'?\s*[=-_\s]*VBA\s*:\s*" & sKeyword & "\s*:\s*Begin[=-_\s]*") Then propState(sKey) = True
      ElseIf RegTest(s, "^'?\s*[=-_\s]*VBA\s*:\s*" & sKeyword & "\s*:\s*End[=-_\s]*") Then propState(sKey) = False
   End If
   If propState(sKey) And (Not propState(sKey) Or Not propGlobal(sKeyword)) Then
      bRemove = True
   End If
   EditCodeModule_RemoveMode = True
End Function

Private Function EditCodeModule_Import( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim s As String: s = Trim(sComment)
   Dim regResults As Object
   If s <> "" Then
      If RegExe(s, "^'?\s*VBA\s*:\s*Import\s+(.*)", regResults) Then
         If Not ParseBeanImport(regResults(1), propBean("OPTIONS"), True, propMod, propGlobal, msgError) Then
            msgError = msgError & "Beanモジュールの追加に失敗しました: " & regResults(1) & vbCrLf
            EditCodeModule_Import = False
            Exit Function
         End If
         bRemove = True
      End If
   End If
   EditCodeModule_Import = True
End Function

Private Function EditCodeModule_OptionStatements( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   If propState("CODE_STATE") <> "<DECLARE>" Then EditCodeModule_OptionStatements = True: Exit Function
   Dim regResults As Object
   Dim s As String: s = Trim(sLine)
   Dim oOptions As Object: Set oOptions = propMod("OPTIONS")
   If RegExe(s, "^Option\s+Base\s+(0|1)", regResults) Then
      If Not oOptions.Exists("Base") Then
         oOptions("Base") = regResults(1)
      ElseIf oOptions("Base") = "" Then
         oOptions("Base") = regResults(1)
      ElseIf regResults(1) = oOptions("Base") Then
         bRemove = True
      Else
         msgError = msgError & "Option Base指定が矛盾しています: " & propBean("FILE") & vbCrLf
         EditCodeModule_OptionStatements = False
         Exit Function
      End If
   ElseIf RegExe(s, "^Option\s+Compare\s+(Binary|Text|Database)", regResults) Then
      If Not oOptions.Exists("Compare") Then
         oOptions("Compare") = regResults(1)
      ElseIf oOptions("Compare") = "" Then
         oOptions("Compare") = regResults(1)
      ElseIf regResults(1) = oOptions("Compare") Then
         bRemove = True
      Else
         msgError = msgError & "Option Compare指定が矛盾しています" & propBean("FILE") & vbCrLf
         EditCodeModule_OptionStatements = False
         Exit Function
      End If
   ElseIf RegTest(s, "^Option\s+Explicit") Then
      If Not oOptions.Exists("Explicit") Then
         oOptions("Explicit") = True
      ElseIf Not oOptions("Explicit") Then
         oOptions("Explicit") = True
      Else
         bRemove = True
      End If
   ElseIf RegTest(s, "^Option\s+Private") Then
      If Not oOptions.Exists("Private") Then
         oOptions("Private") = True
      ElseIf Not oOptions("Private") Then
         oOptions("Private") = True
      Else
         bRemove = True
      End If
   End If
   EditCodeModule_OptionStatements = True
End Function

Private Function EditCodeModule_Global( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim s As String: s = Trim(sComment)
   Dim sGVarName As String
   Dim regResults As Object
   Dim oGlobalList As Object
   If s <> "" Then
      If RegExe(s, "^'?\s*(?:VBA\s*:)?\s*Global\s+(Set\s+)?([^\s]+)\s+As\s+([^\s]+)\s*(=.*)", regResults) Then
         If Not propGlobal.Exists("GLOBAL_LIST") Then propGlobal.Add Key:="GLOBAL_LIST", Item:=NewDic()
         Set oGlobalList = propGlobal("GLOBAL_LIST")
         sGVarName = regResults(2)
         If oGlobalList.Exists(sGVarName) Then
            msgError = msgError & "Global定義が重複しています: " & sGVarName & "@" & propBean("FILE") & vbCrLf
            EditCodeModule_Global = False
            Exit Function
         End If
         oGlobalList.Add Key:=sGVarName, Item:=DicProp( _
            "NAME", sGVarName, _
            "FILE", propBean("FILE"), _
            "MODULE", propMod("NAME"), _
            "SET", (regResults(1) Like "Set*"), _
            "TYPE", regResults(3), _
            "INIT", Trim(regResults(4)) _
         )
      End If
   End If
   EditCodeModule_Global = True
End Function
Private Function EditCodeModule_GlobalInsert( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim dic As Variant
   Dim sFile As String
   Dim s As String
   Dim sBlank As String
   If RegTest(sComment, "^'[-=\s]+VBA\s*:\s*Global Variable Definition[-=\s]*$") Then
      s = sLine
      sBlank = vbCrLf & String(Len(sLine), " ")
      sFile = ""
      If Not propGlobal.Exists("GLOBAL_LIST") Then EditCodeModule_GlobalInsert = True: Exit Function
      If propGlobal("GLOBAL_LIST").Count = 0 Then EditCodeModule_GlobalInsert = True: Exit Function
      For Each dic In propGlobal("GLOBAL_LIST").Items
         If propGlobal("APPEND_INFO") Then
            If sFile <> dic("FILE") Then
               s = s & sBlank & "'####" & dic("FILE") & "####:###INSERTED###"
               sFile = dic("FILE")
            End If
         End If
         s = s & sBlank & "Public " & dic("NAME") & " As " & dic("TYPE")
      Next dic
      sLine = s
   ElseIf RegTest(sComment, "^'[-=\s]+VBA\s*:\s*Global Variable Initialize[-=\s]*$") Then
      s = sLine
      sBlank = vbCrLf & String(Len(sLine), " ")
      sFile = ""
      If Not propGlobal.Exists("GLOBAL_LIST") Then EditCodeModule_GlobalInsert = True: Exit Function
      If propGlobal("GLOBAL_LIST").Count = 0 Then EditCodeModule_GlobalInsert = True: Exit Function
      For Each dic In propGlobal("GLOBAL_LIST").Items
         If propGlobal("APPEND_INFO") Then
            If sFile <> dic("FILE") Then
               s = s & sBlank & "'####" & dic("FILE") & "####:###INSERTED###"
               sFile = dic("FILE")
            End If
         End If
         If dic("SET") Then
            s = s & sBlank & "Set " & dic("NAME") & " " & dic("INIT")
         Else
            s = s & sBlank & dic("NAME") & " " & dic("INIT")
         End If
      Next dic
      sLine = s
   End If
   EditCodeModule_GlobalInsert = True
End Function

Private Function EditCodeModule_Autocontrols( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Const C_CTRL_MENU = "MENU"
   Const C_CTRL_COMMAND = "COMMAND"
   Const C_CTRL_CELL = "CELL"
   Const C_CTRL_ROW = "ROW"
   Const C_CTRL_COLUMN = "COLUMN"
   Const C_CTRL_SHEET = "SHEET"
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

Private Function EditCodeModule_ConvPrivateScope( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Const C_TYPE = "Private"
   If Not propBean("OPTIONS")("Private") Then EditCodeModule_ConvPrivateScope = True: Exit Function
   Dim regResults As Object
   If propState("CODE_STATE") = "<DECLARE>" Then
      If RegExe(sLine, "^(\s*)(?:Public\s)?(\s*Declare\s.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      ElseIf RegExe(sLine, "^(\s*)(?:Public\s)?(\s*Enum\s.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      ElseIf RegExe(sLine, "^(\s*)(?:Public\s)?(\s*Type\s.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      ElseIf RegExe(sLine, "^(\s*)(?:Public\s)?(\s*Const\s.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      ElseIf RegExe(sLine, "^(\s*)Public\s(.*)", regResults) Then
         sLine = regResults(1) & C_TYPE & " " & regResults(2)
      End If
   ElseIf propState("CODE_STATE") Like "<PROC:*:True>" Then
      If propState("PrivateScope") <> propState("CODE_STATE") Then
         If RegExe(sLine, "^(\s*)(?:Public\s)?(\s*(?:Sub|Function|Property).*)", regResults) Then
            sLine = regResults(1) & C_TYPE & " " & regResults(2)
         End If
         propState("PrivateScope") = propState("CODE_STATE")
      End If
   ElseIf propState("CODE_STATE") = "<INIT>" Or propState("CODE_STATE") = "<END>" Then
      propState("PrivateScope") = ""
   End If
   EditCodeModule_ConvPrivateScope = True
End Function

Private Function EditCodeModule_AutoMacro( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim s As String: s = Trim(sComment)
   Dim vType As Variant
   Dim sProcName As String
   Dim regResults As Object
   Dim oMacroList As Object
   If s <> "" Then
      If propState("CODE_STATE") Like "<PROC:*:True>" Then
         sProcName = SplitEx(propState("CODE_STATE"), Delim:=":")(2)
         For Each vType In Array("Auto_Open", "Auto_Close", "Auto_Activate", "Auto_Deactivate")
            If RegExe(s, "^'\s*VBA\s*:\s*" & vType & "\s*", regResults) Then
               If Not propGlobal.Exists("AutoMacro_" & vType) Then propGlobal.Add Key:="AutoMacro_" & vType, Item:=DicProp()
               Set oMacroList = propGlobal("AutoMacro_" & vType)
               If oMacroList.Exists(sProcName) Then
                  msgError = msgError & "Auto_Open定義が重複しています: " & sProcName & "@" & propBean("FILE") & vbCrLf
                  EditCodeModule_AutoMacro = False
                  Exit Function
               End If
               oMacroList.Add Key:=sProcName, Item:=propMod("NAME") & "." & sProcName
            End If
         Next vType
      End If
   End If
   EditCodeModule_AutoMacro = True
End Function
Private Function EditCodeModule_AutoMacroInsert( _
   opt As Variant, ByRef sLine As String, ByRef sComment As String, ByRef bRemove As Boolean, _
   ByRef propState As Object, ByRef propBean As Object, ByRef propMod As Object, ByRef propGlobal As Object, ByRef msgError As String) As Boolean
   Dim s As String
   Dim sBlank As String
   Dim v As Variant
   Dim regResults As Object
   If RegExe(sComment, "^'[-=\s]*VBA\s*:\s*Run Auto Macro\s*:\s*(\w+)[-=\s]*$", regResults) Then
      s = sLine
      sBlank = vbCrLf & String(Len(sLine), " ")
      If Not propGlobal.Exists("AutoMacro_" & regResults(1)) Then EditCodeModule_AutoMacroInsert = True: Exit Function
      If propGlobal("AutoMacro_" & regResults(1)).Count = 0 Then EditCodeModule_AutoMacroInsert = True: Exit Function
      For Each v In propGlobal("AutoMacro_" & regResults(1)).Items
         s = s & sBlank & "Call " & v & "()"
      Next v
      sLine = s
   End If
   EditCodeModule_AutoMacroInsert = True
End Function

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
Private Function DicProp(ParamArray Target() As Variant) As Object
   Dim oDic As Object
   Dim iNum As Long, i As Long
   Dim vKey As Variant
   Dim bKey As Boolean
   iNum = UBound(Target) - LBound(Target) + 1
   i = LBound(Target)
   If iNum > 0 Then
      If IsObject(Target(i)) Then
         Set oDic = Target(i)
         i = i + 1
      End If
   End If
   If oDic Is Nothing Then
      Set oDic = CreateObject("Scripting.Dictionary")
   End If
   bKey = True
   Do While i < iNum
      If bKey Then
         If IsObject(Target(i)) Then
            Set vKey = Target(i)
         Else
            vKey = Target(i)
         End If
         bKey = False
      Else
         oDic.Add Key:=vKey, Item:=Target(i)
         bKey = True
      End If
      i = i + 1
   Loop
   Set DicProp = oDic
End Function
Private Function GetFSO() As Object
   Static G_FSO As Object
   If G_FSO Is Nothing Then Set G_FSO = CreateObject("Scripting.FileSystemObject")
   Set GetFSO = G_FSO
End Function
Private Function RegTest(ByVal sTest As String, ByVal sPattern As String, Optional ByRef RegExp As Variant) As Boolean
   If IsMissing(RegExp) Then
      RegTest = UniRegExp(sPattern).Test(sTest)
   Else
      RegExp.Pattern = sPattern
      RegTest = RegExp.Test(sTest)
   End If
End Function
Private Function RegExe(ByVal sTest As String, ByVal sPattern As String, Optional ByRef oResults As Object, Optional ByRef RegExp As Variant) As Boolean
   Dim oMatch As Variant
   Dim oMatches As Object
   Dim v As Variant
   Set oResults = New Collection
   If IsMissing(RegExp) Then
      Set oMatches = UniRegExp(sPattern).Execute(sTest)
   Else
      RegExp.Pattern = sPattern
      Set oMatches = RegExp.Execute(sTest)
   End If
   RegExe = (oMatches.Count > 0)
   For Each oMatch In oMatches
      For Each v In oMatch.SubMatches
         oResults.Add v
      Next v
   Next oMatch
End Function
Private Function Join(vList As Variant, Optional ByVal sDelim As String = " ") As String
   Dim s As String, v As Variant
   s = ""
   For Each v In vList
      If s = "" Then
         s = v
      Else
         s = s & sDelim & v
      End If
   Next v
   Join = s
End Function
Private Function SplitEx(str As String, Optional Delim As String = ",", Optional Quote As String = """") As Collection
   Dim lpos As Long
   Dim i As Long
   Dim ch As String
   Dim QuoteMode As Boolean
   lpos = 0
   QuoteMode = False
   Set SplitEx = New Collection
   For i = 1 To Len(str)
      ch = Mid(str, i, 1)
      If Not QuoteMode Then
         If ch = Delim Then
            SplitEx.Add Mid(str, lpos + 1, i - lpos - 1)
            lpos = i
         ElseIf ch = Quote Then
            QuoteMode = True
         End If
      Else
         If ch = Quote Then
            QuoteMode = False
         End If
      End If
   Next i
   If i <> lpos Then
      SplitEx.Add Mid(str, lpos + 1, i - lpos - 1)
   End If
End Function
Private Sub VBA_AddErrorTrap(oMod As Object)
   Dim i As Long
   Dim sProcName As String
   Dim sLine As String
   Dim iProcStartLine As Long
   Dim bProc As Boolean
   Dim regResults As Object
   Dim bEndProcExist As Boolean
   Dim sProcType As String
   Dim bErrTrapExist As Boolean
   Dim bMultiProcDef As Boolean
   Dim bIgnoreProc As Boolean
   With oMod.CodeModule
      bProc = False
      For i = 1 To .CountOfDeclarationLines
         sLine = .Lines(i, 1)
         If RegTest(sLine, "^\s*'.*\s*VBA\s*:\s*Auto Added Error Trap Code") Then Exit Sub
      Next i
      Call .InsertLines(1, "'!=!=!=!=! VBA: Auto Added Error Trap Code !=!=!=!=!")
      Call .InsertLines(2, "Private Const SSS_MOD_NAME = """ & oMod.Name & """")
      i = .CountOfDeclarationLines + 1
      Do While i <= .CountOfLines
         If .ProcOfLine(i, 0) <> vbNullString Then
            If sProcName <> .ProcOfLine(i, 0) Then
               sProcName = .ProcOfLine(i, 0)
               iProcStartLine = .ProcStartLine(sProcName, 0)
               bEndProcExist = False
               bErrTrapExist = False
               If sProcName Like "*ERR_TRAP_CALL*" Then
                  bIgnoreProc = True
               Else
                  bIgnoreProc = False
               End If
            End If
            If bIgnoreProc Then GoTo CONTINUE
            If iProcStartLine <= i Then
               sLine = .Lines(i, 1)
               If RegTest(sLine, "^\s*(Private\s+|Public\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property)\s*.*") Or bMultiProcDef Then
                  If sLine Like "* _" Then
                     bMultiProcDef = True
                  Else
                     Call .InsertLines(i + 1, "Const SSS_PROC_NAME = """ & sProcName & """: On Error GoTo SSS_ERR_TRAP")
                     i = i + 1
                     bMultiProcDef = False
                     bProc = True
                  End If
                  GoTo CONTINUE
               ElseIf RegExe(sLine, "^\s*End\s+(Sub|Function|Property)\s*", regResults) Then
                  sProcType = regResults(1)
                  If Not bErrTrapExist Then
                     If bEndProcExist Then
                        Call .InsertLines(i, "Exit " & sProcType)
                     Else
                        Call .InsertLines(i, "SSS_END_PROC: Exit " & sProcType)
                     End If
                     Call .InsertLines(i + 1, "SSS_ERR_TRAP: Call SSS_ERR_TRAP_CALL(SSS_PROC_NAME, SSS_MOD_NAME): Resume SSS_END_PROC")
                     i = i + 2
                  End If
                  bProc = False
                  GoTo CONTINUE
               ElseIf RegTest(sLine, "^\s*SSS_END_PROC\s*:.*") Then
                  bEndProcExist = True
               ElseIf RegTest(sLine, "^\s*SSS_ERR_TRAP\s*:") Then
                  bErrTrapExist = True
               End If
            End If
         End If
CONTINUE:
         i = i + 1
      Loop
   End With
End Sub
Private Sub VBA_AddLineNumbers(oMod As Object)
   Dim i As Long, iOff As Long
   Dim sProcName As String
   Dim sLine As String
   Dim iProcStartLine As Long
   Dim iProcEndLine As Long
   Dim bProc As Boolean
   Dim bMultiLine As Boolean
   With oMod.CodeModule
      Select Case oMod.Type
      Case 1: iOff = 10001
      Case 2: iOff = 10009
      Case 3: iOff = 10015
      Case 100: iOff = 10009
      Case Else: iOff = 10000
      End Select
      bProc = False
      bMultiLine = False
      For i = 1 To .CountOfDeclarationLines
         If RegTest(.Lines(i, 1), "^\s*'.*Auto Added Line Numbers") Then Exit Sub
      Next i
      Call .InsertLines(1, "'!=!=!=!=! Auto Added Line Numbers !=!=!=!=!")
      For i = .CountOfDeclarationLines + 1 To .CountOfLines
         If .ProcOfLine(i, 0) <> vbNullString Then
            If sProcName <> .ProcOfLine(i, 0) Then
               sProcName = .ProcOfLine(i, 0)
               iProcStartLine = .ProcStartLine(sProcName, 0)
               iProcEndLine = .ProcCountLines(sProcName, 0) + iProcStartLine - 1
            End If
            If iProcStartLine <= i And i <= iProcEndLine Then
               sLine = .Lines(i, 1)
               If RegTest(sLine, "^\s*(Private\s+|Public\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property)\s*") Then
                  bProc = True
                  If RegTest(sLine, "_\s*$") Then bMultiLine = True
                  GoTo CONTINUE
               ElseIf RegTest(sLine, "^\s*End\s+(Sub|Function|Property)\s*") Then
                  bProc = False
                  If RegTest(sLine, "_\s*$") Then bMultiLine = True
                  GoTo CONTINUE
               End If
               If bProc Then
                  If bMultiLine Then
                     .ReplaceLine i, Space(5) & " " & sLine
                     If Not RegTest(sLine, "_\s*$") Then bMultiLine = False
                     GoTo CONTINUE
                  End If
                  If RegTest(sLine, "_\s*$") Then bMultiLine = True
                  If RegTest(sLine, "^\s*(Case)\s*") Then
                     .ReplaceLine i, Space(5) & " " & sLine
                     GoTo CONTINUE
                  ElseIf RegTest(sLine, "^\s*#") Then
                     .ReplaceLine i, Space(5) & " " & sLine
                     GoTo CONTINUE
                  End If
                  .ReplaceLine i, CStr(i + iOff) & Space(5 - Len(CStr(i + iOff))) & " " & sLine
               End If
            End If
         End If
CONTINUE:
      Next i
   End With
End Sub
Private Sub VBA_CompileProject()
#If ACCESS_VBA <> 1 Then
   Dim oCtrl As Object
   Set oCtrl = Nothing
   On Error Resume Next
   Set oCtrl = Application.VBE.CommandBars.FindControl(, 578)
   On Error GoTo 0
   If Not oCtrl Is Nothing Then
      If oCtrl.Enabled = True Then oCtrl.Execute
   End If
#End If
End Sub
Private Function VBA_ExportCodeModule(ByRef oComp As Object, ByVal sDir As String, Optional ByVal IgnoreBlankDcm As Boolean = True) As Boolean
   Dim sName As String
   VBA_ExportCodeModule = False
   Select Case oComp.Type
   Case 1
      sName = oComp.Name & ".bas"
   Case 2
      sName = oComp.Name & ".cls"
   Case 3
      sName = oComp.Name & ".frm"
   Case 100
      If (Not IgnoreBlankDcm) Or (oComp.CodeModule.CountOfLines > 0) Then
         sName = oComp.Name & ".dcm"
      Else
         sName = ""
      End If
   End Select
   If sName <> "" Then
      sDir = Trim(sDir)
      If Right(sDir, 1) = "/" Then
         oComp.Export sDir & sName
      Else
         oComp.Export sDir & "/" & sName
      End If
   End If
   VBA_ExportCodeModule = True
End Function
Private Function VBA_GetModuleName(ByRef oComp As Object, Optional ByVal IgnoreBlankDcm As Boolean = False) As String
   Dim sName As String
   Select Case oComp.Type
   Case 1
      sName = oComp.Name & ".bas"
   Case 2
      sName = oComp.Name & ".cls"
   Case 3
      sName = oComp.Name & ".frm"
   Case 100
      If Not IgnoreBlankDcm Or oComp.CodeModule.CountOfLines > 0 Then
         sName = oComp.Name & ".dcm"
      Else
         sName = ""
      End If
   End Select
   VBA_GetModuleName = sName
End Function
Private Function VBA_GetModule(ByRef oComps As Object, ByVal sName As String) As Object
   Dim o As Object
   Set VBA_GetModule = Nothing
   For Each o In oComps
      If o.Name = sName Then
         Set VBA_GetModule = o
         Exit For
      End If
   Next o
End Function
Private Function VBA_GetCommentOfLine(ByVal sLine As String) As Long
   Dim b As Boolean
   Dim iComment As Long
   Dim iQuoteBeg As Long, iQuoteEnd As Long
   iComment = 0
   Do While True
      iComment = InStr(iComment + 1, sLine, "'")
      If iComment = 0 Then
         b = False
         Exit Do
      End If
      iQuoteEnd = 0
      Do While True
         iQuoteBeg = InStr(iQuoteEnd + 1, sLine, """")
         iQuoteEnd = InStr(iQuoteBeg + 1, sLine, """")
         If iQuoteBeg = 0 Or iQuoteEnd = 0 Then
            b = True
            Exit Do
         End If
         If iQuoteBeg < iComment And iComment < iQuoteEnd Then
            b = False
            Exit Do
         End If
      Loop
      If b Then Exit Do
   Loop
   If b Then
      VBA_GetCommentOfLine = iComment
   Else
      VBA_GetCommentOfLine = 0
   End If
End Function
Private Function VBA_GetCountOfDeclarationLines(ByRef oMod As Object) As Long
   Dim iDeclPos As Long
   Dim sProcName As String
   Dim i As Long
   Dim iType As Long
   With oMod.CodeModule
      iDeclPos = .CountOfDeclarationLines
      If .CountOfLines - iDeclPos > 0 Then
         sProcName = .ProcOfLine(iDeclPos + 1, iType)
         For i = .ProcStartLine(sProcName, iType) To .ProcBodyLine(sProcName, iType) - 1
            If (Trim(.Lines(i, 1)) Like "[#]End*") Then iDeclPos = i
         Next i
      End If
   End With
   VBA_GetCountOfDeclarationLines = iDeclPos
End Function
Private Function VBA_GetComponents() As Variant
#If ACCESS_VBA <> 1 Then
   Set VBA_GetComponents = ThisWorkbook.VBProject.VBComponents
#Else
   Set VBA_GetComponents = Application.VBE.ActiveVBProject.VBComponents
#End If
End Function
Private Function VBA_GetDicProc(ByRef oMod As Object) As Object
   Dim lp_line As Long
   Dim prcName As String
   Dim sName As String
   Dim dicProc As Object
   Dim fp As Long
   Dim iType As Long
   sName = VBA_GetModuleName(oMod, IgnoreBlankDcm:=True)
   If sName <> "" Then
      Set dicProc = CreateObject("Scripting.Dictionary")
   End If
   With oMod.CodeModule
      prcName = ""
      For lp_line = 1 To .CountOfLines
         If prcName <> .ProcOfLine(lp_line, iType) Then
            prcName = .ProcOfLine(lp_line, iType)
            dicProc.Add Key:=prcName, Item:=Array(lp_line, Trim(.Lines(.ProcBodyLine(prcName, iType), 1)))
         End If
      Next lp_line
   End With
   Set VBA_GetDicProc = dicProc
End Function
Private Function VBA_ImportCodeModule(ByRef oComps As Object, ByVal sPath As String, ByRef msgError As String) As Object
   Set VBA_ImportCodeModule = Nothing
   Dim oComp As Object
   Select Case GetFSO().GetExtensionName(sPath)
   Case "bas", "cls", "frm"
      Set oComp = oComps.Import(sPath)
      If oComp Is Nothing Then Exit Function
   Case "dcm"
      Set oComp = VBA_ImportDocumentCodeModule(oComps, sPath, msgError)
      If oComp Is Nothing Then Exit Function
   Case Else
      msgError = msgError & "VBAモジュールファイルの拡張子が不正です: " & sPath & vbCrLf
      Exit Function
   End Select
   Set VBA_ImportCodeModule = oComp
End Function
Private Sub VBA_MergeCodeModule(ByRef oBaseMod As Object, ByRef oImportMod As Object, Optional ByVal sTag As String = "")
   Dim s As String
   Dim iDeclPos As Long
   With oImportMod.CodeModule
      s = ""
      iDeclPos = VBA_GetCountOfDeclarationLines(oImportMod)
      If iDeclPos > 0 Then
         s = .Lines(1, iDeclPos)
         If sTag <> "" Then s = sTag & vbCrLf & s
      End If
   End With
   If s <> "" Then
      With oBaseMod.CodeModule
         .InsertLines VBA_GetCountOfDeclarationLines(oBaseMod) + 1, s
      End With
   End If
   With oImportMod.CodeModule
      s = ""
      If .CountOfLines - iDeclPos > 0 Then
         s = .Lines(iDeclPos + 1, .CountOfLines - iDeclPos)
         If sTag <> "" Then s = sTag & vbCrLf & s
      End If
   End With
   If s <> "" Then
      With oBaseMod.CodeModule
         .InsertLines .CountOfLines + 1, s
      End With
   End If
End Sub
Private Function VBA_RemoveCodeModule(ByRef oComps As Object, ByRef o As Object) As Boolean
   If o.Type = 100 Then
      With o.CodeModule
         .DeleteLines StartLine:=1, Count:=.CountOfLines
      End With
   Else
      oComps.Remove o
   End If
   VBA_RemoveCodeModule = True
End Function
Private Function VBA_RemoveProcFromCodeModule(ByRef oMod As Object, ByVal sNameLike As String) As Boolean
   Dim sProcName As String
   Dim iLine As Long
   Dim iLineFirst As Long
   Dim iLineEnd As Long
   With oMod.CodeModule
      sProcName = ""
      iLine = 1
      iLineFirst = 1
      iLineEnd = 1
      Do While iLine <= .CountOfLines
         If sProcName <> .ProcOfLine(iLine, 0) Then
            If sProcName Like sNameLike Then
               .DeleteLines StartLine:=iLineFirst, Count:=iLineEnd - iLineFirst + 1
               iLine = iLineFirst
            Else
               iLineFirst = iLine
               iLine = iLine + 1
            End If
            sProcName = .ProcOfLine(iLine, 0)
         Else
            iLineEnd = iLine
            iLine = iLine + 1
         End If
      Loop
      If sProcName Like sNameLike Then
         .DeleteLines StartLine:=iLineFirst, Count:=iLineEnd - iLineFirst + 1
      End If
   End With
   VBA_RemoveProcFromCodeModule = True
End Function
Private Sub VBA_RemoveLineNumbers(oMod As Object)
   Dim i As Long
   Dim sProcName As String
   Dim sLine As String
   Dim iProcStartLine As Long
   Dim iProcEndLine As Long
   Dim bProc As Boolean
   Dim bTarget As Boolean
   With oMod.CodeModule
      bProc = False
      bTarget = False
      For i = 1 To .CountOfDeclarationLines
         If RegTest(.Lines(i, 1), "^\s*'.*Auto Added Line Numbers") Then
            Call .DeleteLines(i)
            bTarget = True
            Exit For
         End If
      Next i
      If Not bTarget Then Exit Sub
      For i = .CountOfDeclarationLines + 1 To .CountOfLines
         If .ProcOfLine(i, 0) <> vbNullString Then
            If sProcName <> .ProcOfLine(i, 0) Then
               sProcName = .ProcOfLine(i, 0)
               iProcStartLine = .ProcStartLine(sProcName, 0)
               iProcEndLine = .ProcCountLines(sProcName, 0) + iProcStartLine - 1
            End If
            If iProcStartLine <= i And i <= iProcEndLine Then
               sLine = .Lines(i, 1)
               If RegTest(sLine, "^\s*(Private\s+|Public\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property)\s*") Then
                  bProc = True
                  GoTo CONTINUE
               ElseIf RegTest(sLine, "^\s*End\s+(Sub|Function|Property)\s*") Then
                  bProc = False
                  GoTo CONTINUE
               End If
               If bProc Then
                  .ReplaceLine i, Mid(sLine, 7)
               End If
            End If
         End If
CONTINUE:
      Next i
   End With
End Sub
Private Function GetOwnPath() As String
#If ACCESS_VBA <> 1 Then
   GetOwnPath = ThisWorkbook.Path
#Else
   GetOwnPath = Application.CurrentProject.Path
#End If
End Function
Private Function GetAbsolutePath(ByVal pathFile As String, Optional ByVal sBaseDir As String = "") As String
   With GetFSO()
      If IsAbsolutePath(pathFile) Then
         GetAbsolutePath = pathFile
         Exit Function
      End If
      If sBaseDir <> "" Then
         pathFile = .BuildPath(sBaseDir, pathFile)
      End If
      GetAbsolutePath = .GetAbsolutePathName(pathFile)
   End With
End Function
Function GetAbsolutePathWithSearchPath(ByVal pathFile As String, ByVal aSearchPath As Variant, Optional ByRef sDir As String, Optional ByRef sFile As String) As String
   Dim sPath As String
   Dim b As Boolean
   Dim vDir As Variant
   With GetFSO()
      If IsAbsolutePath(pathFile) Then
         If .FileExists(pathFile) Then
            sPath = pathFile
            sDir = .GetParentFolderName(pathFile)
            sFile = GetBaseFileName(pathFile)
            b = True
         ElseIf .FolderExists(pathFile) Then
            sPath = pathFile
            sDir = .GetParentFolderName(pathFile)
            sFile = ""
            b = True
         Else
            sPath = ""
            sDir = ""
            sFile = ""
            b = False
         End If
      Else
         b = False
         For Each vDir In aSearchPath
            sPath = .GetAbsolutePathName(.BuildPath(vDir, pathFile))
            If .FileExists(sPath) Then
               sDir = .GetParentFolderName(sPath)
               sFile = GetBaseFileName(sPath)
               b = True
               Exit For
            ElseIf .FolderExists(sPath) Then
               sDir = .GetParentFolderName(sPath)
               sFile = ""
               b = True
               Exit For
            End If
         Next vDir
      End If
   End With
   If b Then
      GetAbsolutePathWithSearchPath = sPath
   Else
      GetAbsolutePathWithSearchPath = ""
   End If
End Function
Private Function SetCurrentPath(ByVal sDir As String) As String
   GetShell().CurrentDirectory = sDir
End Function
Private Function IsAbsolutePath(ByVal s As String) As Boolean
   Dim c As String
   s = Trim(s)
   If s = "" Then
      IsAbsolutePath = False
   ElseIf s Like "*:/" Or s Like "*:\" Then
      IsAbsolutePath = True
   ElseIf s Like "/*" Or s Like "\*" Then
      IsAbsolutePath = True
   Else
      IsAbsolutePath = (StrConv(s, vbUpperCase) = StrConv(GetFSO().GetAbsolutePathName(s), vbUpperCase))
   End If
End Function
Private Function GetBaseFileName(ByVal sPath As String) As String
   With GetFSO()
      GetBaseFileName = .GetBaseName(sPath) & "." & .GetExtensionName(sPath)
   End With
End Function
Private Function GetCurrentPath() As String
   GetCurrentPath = GetShell().CurrentDirectory
End Function
Private Function GetDicFromIniFile(ByVal sIniFile As String, Optional ByVal iDuplicateMode As Long = 0) As Object
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
            Case 2
               oSec.Remove sKey
               oSec.Add Key:=sKey, Item:=sVal
            Case 1
            Case Else
               Err.Raise Number:=1000, Description:="Section[" & sSec & "]のKey[" & sKey & "]が重複しています。"
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
Private Function PopCollection(ByRef o As Object) As Variant
   With o
      If .Count > 0 Then
         If IsObject(.Item(.Count)) Then
            Set PopCollection = .Item(.Count)
         Else
            PopCollection = .Item(.Count)
         End If
         .Remove .Count
      Else
         PopCollection = Null
      End If
   End With
End Function
Private Function PushCollection(ByRef o As Object, newItem As Variant) As Variant
   If o Is Nothing Then Set o = New Collection
   With o
      .Add newItem
      If IsObject(.Item(.Count)) Then
         Set PushCollection = .Item(.Count)
      Else
         PushCollection = .Item(.Count)
      End If
   End With
End Function
Private Function PeekCollection(ByRef o As Object) As Variant
   With o
      If .Count > 0 Then
         If IsObject(.Item(.Count)) Then
            Set PeekCollection = .Item(.Count)
         Else
            PeekCollection = .Item(.Count)
         End If
      Else
         PeekCollection = Null
      End If
   End With
End Function
Private Function IterateDictionary(o As Variant) As Collection
   Dim v As Variant
   Set IterateDictionary = New Collection
   With IterateDictionary
      For Each v In o.Keys
         .Add Array(v, o(v))
      Next v
   End With
End Function
Private Function ExpandEnvironmentStringsWhole(ByVal str As String) As String
   Dim oWSH As Object
   Dim s As String
   Set oWSH = CreateObject("WScript.Shell")
   Do While True
      s = oWSH.ExpandEnvironmentStrings(str)
      If s = str Then Exit Do
      str = s
   Loop
   ExpandEnvironmentStringsWhole = s
   Set oWSH = Nothing
End Function
Sub LogEasy(ByVal str As String, Optional ByVal sFile As String = "", Optional ByVal bDate As Boolean = False, Optional ByVal bClose As Boolean = False, Optional Dump As Variant)
   Static oFile As Object
   Dim s As String
   If oFile Is Nothing Then
      With GetFSO()
         If sFile = "" Then sFile = "ezy_log.log"
         If bDate Then sFile = .GetBaseName(sFile) & "_" & Format(Now, "yyyymmss_hhmmss") & "." & .GetExtensionName(sFile)
         sFile = .BuildPath(GetOwnPath(), sFile)
         Set oFile = .OpenTextFile(Filename:=sFile, IOMode:=8, Create:=True)
      End With
   End If
   s = Format(Now, "yyyy/mm/dd hh:mm:ss") & ", " & str
   If Not IsMissing(Dump) Then s = s & ", " & Dump(Dump)
   s = s & "."
   oFile.WriteLine s
   If bClose Then
      oFile.Close
      Set oFile = Nothing
   End If
End Sub
Private Function UniRegExp(ByVal Pattern As Variant, Optional bSet As Boolean = False, Optional ByVal bGlobal As Boolean = True, Optional ByVal MultiLine As Boolean = True, Optional ByVal IgnoreCase As Boolean = False) As Object
   Static G_RegExp As Object
   If G_RegExp Is Nothing Then
      Set G_RegExp = CreateObject("VBScript.RegExp")
      With G_RegExp
         .Global = bGlobal
         .MultiLine = MultiLine
         .IgnoreCase = IgnoreCase
      End With
   ElseIf bSet Then
      With G_RegExp
         .Global = bGlobal
         .MultiLine = MultiLine
         .IgnoreCase = IgnoreCase
      End With
   End If
   If Not IsNull(Pattern) Then G_RegExp.Pattern = Pattern
   Set UniRegExp = G_RegExp
End Function
Private Function VBA_ImportDocumentCodeModule(ByRef oComps As Object, ByVal sPath As String, ByRef msgError As String) As Object
   Set VBA_ImportDocumentCodeModule = Nothing
   Dim o As Object
   Dim sName As String
   Dim bExist As Boolean
   sName = GetFSO().GetBaseName(sPath)
   bExist = False
   For Each o In oComps
      If o.Name = sName Then
         bExist = True
         Exit For
      End If
   Next o
   If Not bExist Then
      msgError = msgError & "インポート先のDocumentが見つかりません: " & sName & vbCrLf
      Exit Function
   End If
   With o.CodeModule
      .DeleteLines StartLine:=1, Count:=.CountOfLines
      .AddFromFile sPath
      .DeleteLines StartLine:=1, Count:=4
   End With
   Set VBA_ImportDocumentCodeModule = o
End Function
Private Function GetShell() As Object
   Static G_Shell As Object
   If G_Shell Is Nothing Then Set G_Shell = CreateObject("WScript.Shell")
   Set GetShell = G_Shell
End Function
Private Function NewDic() As Object
   Set NewDic = CreateObject("Scripting.Dictionary")
End Function
Private Function Dump(ByRef x As Variant, Optional ByVal WithPtr As Boolean = False) As String
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
