Sub ObjectFit()
    Const myTitle = "オブジェクトをセルに合わせる"
    Dim obj As Object

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub

    Select Case TypeName(Selection)
        Case "DrawingObjects"
            For Each obj In Selection
                ObjectFitDrawingObject obj
            Next
        Case "Range"
            If MsgBox("アクティブシートのすべての図形オブジェクトをセルに合わせます。" & Chr$(10) & "元に戻すことはできません。", _
                vbExclamation Or vbOKCancel, myTitle) <> vbOK Then Exit Sub
            For Each obj In ActiveSheet.DrawingObjects
                ObjectFitDrawingObject obj
            Next
        Case Else
            For Each obj In ActiveSheet.DrawingObjects
                If obj.Name = Selection.Name Then
                    ObjectFitDrawingObject obj
                    Exit For
                End If
            Next
    End Select
End Sub

Sub ObjectFitDrawingObject(obj As Object)
    Dim rng1 As Range, rng2 As Range
    Dim objRight As Double, objBottom As Double

    objRight = obj.Left + obj.Width
    objBottom = obj.Top + obj.Height
    Set rng1 = obj.TopLeftCell
    Set rng2 = obj.BottomRightCell

    If (rng1.Top + rng1.Height / 2# < obj.Top) Then
        Set rng1 = rng1.Offset(1, 0)
    End If
    If (rng1.Left + rng1.Width / 2# < obj.Left) Then
        Set rng1 = rng1.Offset(0, 1)
    End If

    If (rng2.Top + rng2.Height / 2# < objBottom) Then
        Set rng2 = rng2.Offset(1, 0)
    End If
    If (rng2.Left + rng2.Width / 2# < objRight) Then
        Set rng2 = rng2.Offset(0, 1)
    End If

    obj.Left = rng1.Left
    obj.Top = rng1.Top
    obj.Width = rng2.Left - rng1.Left
    obj.Height = rng2.Top - rng1.Top
End Sub
