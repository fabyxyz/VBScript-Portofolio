Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim qt : qt = """"
Dim src

Call ini()
Sub ini()
	Call search()
End Sub

Sub search()
	src = InputBox("Search filename:","File Searcher")
	src = CStr(src)
	src = LCase(src)
	if src = "" then WScript.Quit
	if not fso.FileExists("Files/" & src & ".txt") then
		if msgBox("Could not find any file with the name " & qt & src & qt,5+16,"File Searcher") = vbCancel then
			WScript.Quit
		else
			search()
		end if
	else
		Dim folderName : folderName = "Files"
		Dim fullPath : fullPath = fso.GetAbsolutePathName(folderName)
		shell.Run(fullPath & "/" & src & ".txt")
	end if
End Sub
