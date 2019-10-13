Sub ExcelUpdateStop(Optional ByVal sStatus As String = "Tasking...")
   Application.Calculation = xlCalculationManual
   Application.StatusBar = sStatus
   Application.Cursor = xlWait
   ' Application.EnableEvents = False
   ' Application.Interactive = False
   DoEvents
   Application.ScreenUpdating = False
End Sub

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ExcelUpdateStop()
   Dim vScreenUpdating As Variant
   Dim vCalculation As Variant
   Dim vStatusBar As Variant
   Dim vCursor As Variant
   vScreenUpdating = Application.ScreenUpdating
   vCalculation = Application.Calculation
   vStatusBar = Application.StatusBar
   vCursor = Application.Cursor

   Call ExcelUpdateStop()

   Application.ScreenUpdating = True
   DoEvents
   Application.Calculation = xlCalculationAutomatic
   Application.StatusBar = Empty
   Application.Cursor = xlDefault
   ' Application.EnableEvents = True
   ' Application.Interactive = True

   Application.ScreenUpdating = vScreenUpdating
   Application.Calculation = vCalculation
   Application.StatusBar = vStatusBar
   Application.Cursor = vCursor
End Sub
