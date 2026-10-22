'------------------------------------------------------------------------------
' EndpointGuy Intune Toolkit - silent launcher
'
' Starts Toolkit.ps1 with NO PowerShell window at all - nothing flashes up
' and nothing appears on the taskbar except the toolkit itself.
'
' powershell.exe is a console program, so Windows draws its window before
' -WindowStyle Hidden can take effect. That is the brief flash. Starting it
' from here with window style 0 (hidden) means the window is never shown.
'
' Shortcut target:
'   C:\Windows\System32\wscript.exe "C:\Program Files\Endpoint Guy\Intune Toolkit\Launch-Toolkit.vbs"
'
' Any arguments are passed through, e.g.  Launch-Toolkit.vbs -TenantId contoso.onmicrosoft.com
' To troubleshoot a toolkit that will not start, use Run-Toolkit.bat instead.
'------------------------------------------------------------------------------
Option Explicit

Dim shell, fso, folder, powershell, script, cmd, i

Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

folder     = fso.GetParentFolderName(WScript.ScriptFullName)
script     = fso.BuildPath(folder, "Toolkit.ps1")
powershell = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"

If Not fso.FileExists(script) Then
    MsgBox "Toolkit.ps1 was not found next to this launcher:" & vbCrLf & vbCrLf & script, vbCritical, "EndpointGuy Intune Toolkit"
    WScript.Quit 1
End If

cmd = """" & powershell & """ -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & script & """"

For i = 0 To WScript.Arguments.Count - 1
    cmd = cmd & " """ & WScript.Arguments(i) & """"
Next

' 0 = hidden window, False = do not wait for the toolkit to close.
shell.Run cmd, 0, False
