Option Explicit

Dim FileSystem
WScript.Quit Main

Function Main()

	' Always run in console mode...
	If InStr(1, WScript.FullName, "wscript.exe", vbTextCompare) > 0 Then
		CreateObject("WScript.Shell").Run "cmd /c cscript.exe //nologo """ & WScript.ScriptFullName & """", 1, False
		Exit Function
	End If

	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	' Display version info...
	Dim dt0__, tm0__, s0__
	dt0__ = FileSystem.GetFile(WScript.FullName).DateCreated
	tm0__ = Hour(dt0__) & ":" & Minute(dt0__) & ":" & Second(dt0__)
	dt0__ = MonthName(Month(dt0__), True) & " " & Day(dt0__) & " " & Year(dt0__)

	WScript.Echo WScript.Name & " (v" & WScript.Version & ", " & dt0__ & ", " & tm0__ & ")"
	WScript.Echo "Type ""exit"" to exit the interpreter. If 'init.vbs' exists, it will be loaded automatically."

	If FileSystem.FileExists("init.vbs") Then Import "init.vbs"

	' Start REPL...
	Do

		If Len(s0__) > 0 Then

			Select Case LCase(Split(s0__)(0))

				Case "import"
					Import Mid(s0__, Len("import ") + 1)

				Case "class", "do", "for", "function", "if", "select", "sub", "while", "with"
					ExecuteBlock s0__

				Case Else
					ExecuteStatement s0__

			End Select

		End If

		WScript.StdOut.Write ">>> "
		s0__ = WScript.StdIn.ReadLine()

	Loop While LCase(Trim(s0__)) <> "exit"

End Function

Sub Import(strFile)

	If Not FileSystem.FileExists(strFile) Then WScript.Echo "File not found" : Exit Sub

	On Error Resume Next
	ExecuteGlobal FileSystem.OpenTextFile(strFile).ReadAll()

	If Err Then WScript.Echo Err.Description

End Sub

Function ExecuteBlock(ByVal strCode)

	Do
		ExecuteBlock = ExecuteBlock & strCode & vbCrLf
		WScript.StdOut.Write "... "
		strCode = WScript.StdIn.ReadLine()
	Loop While Len(strCode) > 0

	On Error Resume Next
	ExecuteGlobal ExecuteBlock

	If Err Then WScript.Echo Err.Description

End Function

Sub ExecuteStatement(s0__)

	Dim e0__, e1__, r1__, t1__

	On Error Resume Next
	ExecuteGlobal s0__
	e0__ = Err.Description

	Err.Clear
	t1__ = TypeName(Eval(s0__))
	r1__ = Eval(s0__)
	e1__ = Err.Description

	Dim f0__
	If Len(e0__) > 0 Then
		If Len(e1__) > 0 And Len(t1__) = 0 Then f0__ = e0__ Else f0__ = FormatValue(t1__, r1__)
	Else
		If Len(e1__) = 0 Then If InStr(s0__, "=") = 0 Then f0__ = FormatValue(t1__, r1__)
	End If

	' If we need to display output, do so...
	If Len(f0__) > 0 Then WScript.Echo f0__

End Sub

Function FormatValue(t, v)

	Select Case t
		Case "String"
			FormatValue = """" & v & """"
		Case "Variant()"
			FormatValue = FormatArray(v)
		Case "Byte"
			FormatValue = "0x" & Right("0" & Hex(v), 2) & " (" & v & ")"
		Case "Currency"
			FormatValue = FormatCurrency(v)
		Case "Date"
			FormatValue = FormatDateTime(v, vbGeneralDate)
		Case "Boolean"
			' Adding the empty string forces VBScript to display True/False instead of -1/0
			FormatValue = v & ""
		Case "Integer", "Long", "Single", "Double"
			' No special formatting desired
			FormatValue = v
		Case "Empty"
			FormatValue = "Name is undefined"
		Case Else
			' For all non-primitive types, just display the type name instead of its value...
			If Len(v) > 0 Then FormatValue = t & " (" & v & ")" Else FormatValue = "Type " & t
	End Select

End Function

Function FormatArray(a)

	Dim e
	For Each e In a
		FormatArray = FormatArray & ", " & FormatValue(TypeName(e), e)
	Next

	FormatArray = "(" & Mid(FormatArray, 3) & ")"

End Function
