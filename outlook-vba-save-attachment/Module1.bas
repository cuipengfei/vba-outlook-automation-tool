Attribute VB_Name = "Module1"
Public Sub SubClosingPopUp(PauseTime As Integer, Message As String, Title As String)

Dim WScriptShell As Object
Dim ConfigString As String

Set WScriptShell = CreateObject("WScript.Shell")
ConfigString = "mshta.exe vbscript:close(CreateObject(""WScript.Shell"")." & _
               "Popup(""" & Message & """," & PauseTime & ",""" & Title & """))"

WScriptShell.Run ConfigString

End Sub


