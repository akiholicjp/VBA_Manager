' VBA: Import beans/Data/ArrayToSet.bas
' VBA: Import beans/Data/InCollection.bas
' VBA: Import beans/Excel/GetWorksheet.bas

Sub PlotStateChartSub(oSheet As Worksheet, ByVal idx As Long, ByVal t1 As Double, ByVal t2 As Double, ByVal dState As Double, Optional t_offset As Double = 0.0, Optional y_offset As Double = 0.0, Optional ByVal sState As String = "", Optional size As Double = 18, Optional pitch As Double = 2, Optional vRGB As Long = 4210752)
   With oSheet
      With .Shapes
         With .AddShape(msoShapeRectangle, t1 + t_offset, (size + pitch) * idx + y_offset, t2 - t1, size)
            With .Fill
               .ForeColor.RGB = vRGB
               .ForeColor.Brightness = 1# - dState
               '.Transparency = 1# - dState
               '.Patterned iState
               '.Patterned msoPatternWideUpwardDiagonal
            End With
            With .Line
               .Weight = 0.1
               .ForeColor.RGB = RGB(0, 0, 0)
            End With
            If sState <> "" Then
               With .TextFrame2
                  .VerticalAnchor = msoAnchorMiddle
                  .HorizontalAnchor = msoAnchorNone
                  .MarginLeft = 2.0
                  .MarginRight = 2.0
                  .MarginTop = 0
                  .MarginBottom = 0
                  With .TextRange.Characters
                     .Text = sState
                     'With .ParagraphFormat
                     '   .FirstLineIndent = 0
                     '   .Alignment = msoAlignLeft
                     'End With
                     With .Font
                        .NameComplexScript = "ÇlÇr ÉSÉVÉbÉN"
                        .NameFarEast = "ÇlÇr ÉSÉVÉbÉN"
                        .Name = "ÇlÇr ÉSÉVÉbÉN"
                        .Size = 11
                        With .Fill
                           .ForeColor.RGB = RGB(0, 0, 0)
                           If dState > 0.6 Then
                              .ForeColor.Brightness = 1
                           Else
                              .ForeColor.Brightness = 0
                           End If
                           .Transparency = 0
                        End With
                     End With
                  End With
               End With
            End If

         End With
      End With
   End With
End Sub

Sub PlotStateChartLabel(oSheet As Worksheet, ByVal idx As Long, ByVal sLabel As String, ByVal x_size As Double, Optional y_offset As Double = 0.0, Optional size As Double = 18, Optional pitch As Double = 2, Optional vRGB As Long = 4210752)
   With oSheet
      With .Shapes
         With .AddShape(msoShapeRectangle, 0.0, (size + pitch) * idx + y_offset, x_size, size)
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Visible = msoFalse
            With .TextFrame2
               .VerticalAnchor = msoAnchorMiddle
               .HorizontalAnchor = msoAnchorNone
               .MarginLeft = 2.0
               .MarginRight = 2.0
               .MarginTop = 0
               .MarginBottom = 0
               With .TextRange.Characters
                  .Text = sLabel
                  With .ParagraphFormat
                     .FirstLineIndent = 1
                     .Alignment = msoAlignRight
                  End With
                  With .Font
                     .NameComplexScript = "ÇlÇr ÉSÉVÉbÉN"
                     .NameFarEast = "ÇlÇr ÉSÉVÉbÉN"
                     .Name = "ÇlÇr ÉSÉVÉbÉN"
                     .Size = 11
                     .Fill.ForeColor.RGB = vRGB
                  End With
               End With
            End With
         End With
      End With
   End With
End Sub

Sub PlotStateChart(oSheet As Worksheet, n, sLabel, xRng, yRng, Optional vRGB As Long = 4210752)
   Dim vSet As New Collection
   Dim v As Variant
   Dim i As Long
   Dim x_from As Double, x_to As Double, x_base As Double
   Dim old_sState As String, sState As String
   Dim bFirst As Boolean
   For Each v In ArrayToSet(yRng.Value)
      vSet.Add CStr(v)
   Next v
   If vSet.Count > 100 Then
      MsgBox "ÉXÉeÅ[É^ÉXÇ™ëΩÇ∑Ç¨Ç‹Ç∑"
      Exit Sub
   End If
   x_base = CDbl(xRng(1).Value)
   bFirst = True
   Call PlotStateChartLabel(oSheet, n, sLabel, 100.0, vRGB:=vRGB)
   For i = 1 To xRng.Count
      If bFirst Then
         If Not IsError(yRng(i).Value) Then
            sState = CStr(yRng(i).Value)
            x_from = CDbl(xRng(i).Value)
            old_sState = sState
            bFirst = False
         End If
      Else
         If Not IsError(yRng(i).Value) Then sState = CStr(yRng(i).Value) Else sState = old_sState
         If sState <> old_sState Then
            x_to = CDbl(xRng(i).Value)
            Call PlotStateChartSub(oSheet, n, x_from - x_base, x_to - x_base, InCollection(old_sState, vSet) / CDbl(vSet.Count), t_offset:=100.0, sState:=old_sState, vRGB:=vRGB)
            x_from = x_to
         End If
         old_sState = sState
      End If
   Next i
   x_to = CDbl(xRng(i).Value)
   Call PlotStateChartSub(oSheet, n, x_from - x_base, x_to - x_base, InCollection(old_sState, vSet) / CDbl(vSet.Count), t_offset:=100.0, sState:=old_sState, vRGB:=vRGB)
End Sub

Sub tt()
   With ActiveSheet
      Dim v As Variant
      Dim oGraph As Object
      Set oGraph = GetWorksheet(ThisWorkbook, "Graph")
      If oGraph Is Nothing Then
         Set oGraph = ThisWorkbook.Worksheets.Add
         oGraph.Name = "Graph"
      End If
      For Each v In oGraph.Shapes
         v.Delete
      Next v
      Call PlotStateChart(oGraph, 1, .[R8].Text, .[D800:D1950], .[R800:R1950])
      Call PlotStateChart(oGraph, 2, .[S8].Text, .[D800:D1950], .[S800:S1950])
      Call PlotStateChart(oGraph, 4, .[T8].Text, .[D800:D1950], .[T800:T1950])
      Call PlotStateChart(oGraph, 5, .[U8].Text, .[D800:D1950], .[U800:U1950])
      Call PlotStateChart(oGraph, 6, .[V8].Text, .[D800:D1950], .[V800:V1950])
      Call PlotStateChart(oGraph, 7, .[W8].Text, .[D800:D1950], .[W800:W1950])
      Call PlotStateChart(oGraph, 9, .[X8].Text, .[D800:D1950], .[X800:X1950])
      Call PlotStateChart(oGraph, 10, .[Y8].Text, .[D800:D1950], .[Y800:Y1950])
      Call PlotStateChart(oGraph, 11, .[Z8].Text, .[D800:D1950], .[Z800:Z1950])
      Call PlotStateChart(oGraph, 12, .[AA8].Text, .[D800:D1950], .[AA800:AA1950])
      Call PlotStateChart(oGraph, 14, .[AB8].Text, .[D800:D1950], .[AB800:AB1950])
      Call PlotStateChart(oGraph, 15, .[AC8].Text, .[D800:D1950], .[AC800:AC1950])
      Call PlotStateChart(oGraph, 16, .[AD8].Text, .[D800:D1950], .[AD800:AD1950])
      Call PlotStateChart(oGraph, 17, .[AE8].Text, .[D800:D1950], .[AE800:AE1950])
      Call PlotStateChart(oGraph, 19, .[AF8].Text, .[D800:D1950], .[AF800:AF1950])
      Call PlotStateChart(oGraph, 20, .[AG8].Text, .[D800:D1950], .[AG800:AG1950])
      Call PlotStateChart(oGraph, 21, .[AH8].Text, .[D800:D1950], .[AH800:AH1950])
      Call PlotStateChart(oGraph, 22, .[AI8].Text, .[D800:D1950], .[AI800:AI1950])
      Call PlotStateChart(oGraph, 24, .[AJ8].Text, .[D800:D1950], .[AJ800:AJ1950])
      Call PlotStateChart(oGraph, 25, .[AK8].Text, .[D800:D1950], .[AK800:AK1950])
      Call PlotStateChart(oGraph, 26, .[AL8].Text, .[D800:D1950], .[AL800:AL1950])
      Call PlotStateChart(oGraph, 27, .[AM8].Text, .[D800:D1950], .[AM800:AM1950])
   End With
End Sub
