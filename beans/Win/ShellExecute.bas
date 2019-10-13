Function ShellExecute(ByVal sFile As String, Optional ByVal sArgs As String = "", Optional ByVal sDir As String = "", Optional ByVal sOpe As String = "", Optional ByVal iShow As Long = 1) As Long
   ShellExecute = CreateObject("Shell.Application").ShellExecute(sFile, sArgs, sDir, sOpe, iShow)
End Function

' <<iShow>>
' 0: Open the application with a hidden window.
' 1: Open the application with a normal window. If the window is minimized or maximized, the system restores it to its original size and position.
' 2: Open the application with a minimized window.
' 3: Open the application with a maximized window.
' 4: Open the application with its window at its most recent size and position. The active window remains active.
' 5: Open the application with its window at its current size and position.
' 7: Open the application with a minimized window. The active window remains active.
' 10: Open the application with its window in the default state specified by the application.
