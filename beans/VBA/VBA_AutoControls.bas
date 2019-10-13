Option Explicit
Option Base 0

' VBA: Import ../DicProp.bas: Private

' 標準モジュールのPREFIX
Private Const C_CTRL_MENU = "MENU"
Private Const C_CTRL_COMMAND = "COMMAND"
Private Const C_CTRL_CELL = "CELL"
Private Const C_CTRL_ROW = "ROW"
Private Const C_CTRL_COLUMN = "COLUMN"
Private Const C_CTRL_SHEET = "SHEET"

' 追加メニューのタイプ
Private Const C_TYPE_MENU = "MENU"
Private Const C_TYPE_BUTTON = "BUTTON"
Private Const C_TYPE_EDIT = "EDIT"
Private Const C_TYPE_HOOK = "HOOK"

Private Const C_INIT_HOOK = "InitializeHook"

Private G_dicModules As Object

Private Sub loadAutoControlsDefinition()
   '=-=-=-=- VBA: Autocontrols Definition =-=-=-=-
End Sub

Public Sub VBA_AutoControlsOpen() 'VBA: Auto_Open
   Dim cMod As Object
   Dim v As Variant

   If G_dicModules Is Nothing Then Call loadAutoControlsDefinition()
   If G_dicModules Is Nothing Then Exit Sub

   For Each v In G_dicModules.Keys()
      Set cMod = G_dicModules(v)
      If Not cMod Is Nothing Then
         Select Case cMod("Type")
         Case C_CTRL_COMMAND
            Call AutoCtrl(0, cMod)
         Case C_CTRL_MENU
            Call AutoCtrl(Application.CommandBars("Worksheet Menu Bar").Index, cMod)
         Case C_CTRL_CELL
            Call AutoCtrl(Application.CommandBars("Cell").Index, cMod)
            Call AutoCtrl(Application.CommandBars("Cell").Index + 3, cMod) ' For Page Break Preview
         Case C_CTRL_COLUMN
            Call AutoCtrl(Application.CommandBars("Column").Index, cMod)
            Call AutoCtrl(Application.CommandBars("Column").Index + 3, cMod) ' For Page Break Preview
         Case C_CTRL_ROW
            Call AutoCtrl(Application.CommandBars("Row").Index, cMod)
            Call AutoCtrl(Application.CommandBars("Row").Index + 3, cMod) ' For Page Break Preview
         Case C_CTRL_SHEET
            Call AutoCtrl(Application.CommandBars("Ply").Index, cMod)
         End Select
      End If
   Next v
End Sub

Public Sub VBA_AutoControlsClose() 'VBA: Auto_Close
   Dim cMod As Variant
   Dim cCtrl As Variant
   Dim v As Variant
   Dim vv As Variant
   Dim s As String

   If G_dicModules Is Nothing Then Call loadAutoControlsDefinition()
   If G_dicModules Is Nothing Then Exit Sub

   For Each v In G_dicModules.Keys
      Set cMod = G_dicModules(v)
      For Each vv In cMod("Controls").Keys
         Set cCtrl = cMod("Controls")(vv)
         s = cCtrl("ID")
         If cCtrl.Exists("Ctrl") Then
            If Not cCtrl("Ctrl") Is Nothing Then Call cCtrl("Ctrl").Delete()
            cCtrl.Remove "Ctrl"
         End If
         Call cMod("Controls").Remove(s)
      Next vv
      s = cMod("ID")
      If cMod.Exists("Ctrl") Then
         For Each vv In cMod("Ctrl")
            Call vv.Delete()
         Next vv
         cMod.Remove "Ctrl"
      End If
      Call G_dicModules.Remove(s)
   Next v
End Sub

Private Sub AutoCtrl(iType As Long, cMod As Object, Optional bForce As Boolean = True, Optional bTemporary As Boolean = True)
   Dim cMain As Object
   Dim c As Object
   Dim cCtrl As Object
   Dim cCtrltrl As Object
   Dim dicAttr As Object
   Dim v As Variant

   If iType <> 0 Then ' Not C_CTRL_COMMAND
      If bForce Then
         On Error Resume Next
         ThisWorkbook.Application.CommandBars(iType).Controls(cMod("Caption")).Delete
         On Error GoTo 0
      End If
      Set cMain = ThisWorkbook.Application.CommandBars(iType).Controls.Add(Type:=msoControlPopup, Temporary:=bTemporary)
      cMain.Caption = cMod("Caption")
   Else
      If bForce Then
         On Error Resume Next
         ThisWorkbook.Application.CommandBars(cMod("Caption")).Delete
         On Error GoTo 0
      End If
      Set cMain = ThisWorkbook.Application.CommandBars.Add(Name:=cMod("Caption"), Position:=msoBarFloating, MenuBar:=False, Temporary:=bTemporary)
   End If
   If Not cMod.Exists("Ctrl") Then cMod.Add Key:="Ctrl", Item:=New Collection
   cMod("Ctrl").Add cMain

   With cMain
      For Each v In cMod("Controls").Keys()
         Set cCtrl = cMod("Controls")(v)
         Set dicAttr = cCtrl("Attr")
         Set c = Nothing
         Select Case cCtrl("Type")
         Case C_TYPE_HOOK
            If cCtrl("Name") Like C_INIT_HOOK Then Application.Run cCtrl("OnAction")
         Case C_TYPE_MENU
            Set c = .Controls.Add(Type:=msoControlButton)
         Case C_TYPE_BUTTON
            Set c = .Controls.Add(Type:=msoControlButton)
         Case C_TYPE_EDIT
            Set c = .Controls.Add(Type:=msoControlEdit)
         End Select
         If Not c Is Nothing Then
            Set cCtrl("Ctrl") = c
            If dicAttr.Exists("Caption") Then
               c.Caption = dicAttr("Caption")
            Else
               c.Caption = cCtrl("Name")
            End If
            If dicAttr.Exists("FaceId") Then c.FaceId = CLng(dicAttr("FaceId"))
            If dicAttr.Exists("BeginGroup") Then c.BeginGroup = CBool(dicAttr("BeginGroup"))
            c.OnAction = cCtrl("OnAction")
         End If
      Next v
      .Visible = True
   End With
End Sub
