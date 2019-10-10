Option Explicit

' VBA: Import beans/DicProp.bas: Private
' VBA: Import beans/GetFSO.bas: Private
' VBA: Import beans/RegTest.bas: Private
' VBA: Import beans/RegExe.bas: Private
' VBA: Import beans/Join.bas: Private
' VBA: Import beans/SplitEx.bas: Private
' VBA: Import beans/SplitEx.bas: Private
' VBA: Import beans/VBA/VBA_AddErrorTrap.bas: Private
' VBA: Import beans/VBA/VBA_AddLineNumbers.bas: Private
' VBA: Import beans/VBA/VBA_CompileProject.bas: Private
' VBA: Import beans/VBA/VBA_ExportCodeModule.bas: Private
' VBA: Import beans/VBA/VBA_GetModuleName.bas: Private
' VBA: Import beans/VBA/VBA_GetModule.bas: Private
' VBA: Import beans/VBA/VBA_GetCommentOfLine.bas: Private
' VBA: Import beans/VBA/VBA_GetCountOfDeclarationLines.bas: Private
' VBA: Import beans/VBA/VBA_GetComponents.bas: Private
' VBA: Import beans/VBA/VBA_GetDicProc.bas: Private
' VBA: Import beans/VBA/VBA_ImportCodeModule.bas: Private
' VBA: Import beans/VBA/VBA_MergeCodeModule.bas: Private
' VBA: Import beans/VBA/VBA_RemoveCodeModule.bas: Private
' VBA: Import beans/VBA/VBA_RemoveProcFromCodeModule.bas: Private
' VBA: Import beans/VBA/VBA_RemoveLineNumbers.bas: Private
' VBA: Import beans/Path/GetOwnPath.bas: Private
' VBA: Import beans/Path/GetAbsolutePath.bas: Private
' VBA: Import beans/Path/GetAbsolutePathWitSearchPath.bas
' VBA: Import beans/Path/SetCurrentPath.bas: Private
' VBA: Import beans/Path/IsAbsolutePath.bas: Private
' VBA: Import beans/Path/GetBaseFileName.bas: Private
' VBA: Import beans/Path/GetCurrentPath.bas: Private
' VBA: Import beans/Path/GetOwnPath.bas: Private
' VBA: Import beans/File/GetDicFromIniFile.bas: Private
' VBA: Import beans/Data/PopCollection.bas: Private
' VBA: Import beans/Data/PushCollection.bas: Private
' VBA: Import beans/Data/PeekCollection.bas: Private
' VBA: Import beans/Data/IterateDictionary.bas: Private
' VBA: Import beans/Win/ExpandEnvironmentStringsWhole.bas: Private

' VBA: Import beans/LogEasy.bas

Private Const C_SHORTCUT_CLEAR = "C"
Private Const C_SHORTCUT_RELOAD = "L"
Private Const C_SHORTCUT_COMPILE = "X"
Private Const C_CONF_INIFILE = "VBA_Manager" '設定ファイルのファイル名
Private Const C_MOD_VBAMANAGER = "Z_VBA_Manager"
Private Const C_MOD_UNMANAGED = "Zc_*"

Private G_CurDirStack As New Collection

Public Sub Step0_clearModules()
   Call clearModules()
End Sub

Public Sub Step1_LoadModules()
   Call loadModules()
   Call VBA_CompileProject()
End Sub

Public Sub Step2_compileModules()
   Call VBA_CompileProject()
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

   ' カレントパスの設定
   SetCurrentPath(GetOwnPath())

   ' Global Property, モジュールリストの読み込み
   For Each v In GetLibListArray()
      msgError = ""
      If loadIniFile(v, propGlobal, oModList, msgError) Then Exit For
   Next v
   If oModList Is Nothing Then msgError = msgError & "モジュールの構築リストが見つかりません"
   If msgError <> "" Then GoTo exit_loadModules

   'Modules読み込み
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
   Call SaveThis()
End Sub

Private Sub Auto_Open()
   If Workbooks.Count = 0 Then Workbooks.Add
   If C_SHORTCUT_CLEAR <> "" Then Application.MacroOptions Macro:="Z_VBA_Manager.Step0_clearModules", ShortcutKey:=C_SHORTCUT_CLEAR
   If C_SHORTCUT_RELOAD <> "" Then Application.MacroOptions Macro:="Z_VBA_Manager.Step1_loadModules", ShortcutKey:=C_SHORTCUT_RELOAD
   If C_SHORTCUT_COMPILE <> "" Then Application.MacroOptions Macro:="Z_VBA_Manager.Step2_compileModules", ShortcutKey:=C_SHORTCUT_COMPILE
   'LogEasy "Z_VBA_Manager.Auto_Open"
End Sub

Private Sub Auto_Close()
   'LogEasy "Z_VBA_Manager.Auto_Close", bClose:=True
   Application.MacroOptions Macro:="Z_VBA_Manager.Step0_clearModules", ShortcutKey:=""
   Application.MacroOptions Macro:="Z_VBA_Manager.Step1_loadModules", ShortcutKey:=""
   Application.MacroOptions Macro:="Z_VBA_Manager.Step2_compileModules", ShortcutKey:=""
End Sub

'===================================================================================================

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

'===================================================================================================

Private Function GetLibListArray() As Variant
   GetLibListArray =  Array(C_CONF_INIFILE & ".ini")
End Function

Private Sub SaveThis()
#If ACCESS_VBA <> 1 Then
   ThisWorkbook.Save
#End If
End Sub

'===================================================================================================

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
      ' モジュール構築
      Select Case propMod("TYPE")
      Case "new"
         Set oModule = VBA_GetComponents().Add(1) ' Module.Standard
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

      ' Beanの追加
      Do While True ' 追加するBeanモジュールがなくなるまで繰り返す。
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
