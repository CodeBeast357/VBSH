Class NamespaceVBSH
	Private headingmulti
	Private footermulti
	Private countmulti
	Private cmdexit
	Private prompt
	Private promptmulti

	Private Sub PrintError()
		WScript.StdOut.Write "Error #" & CStr(Err.Number) & ": " & Err.Description & vbCrLf
	End Sub

	Private Function GetFooter(input)
		term = UCase(LTrim(input))
		For i = 0 To countmulti
			heading = headingmulti(i)
			If Left(term, Len(heading)) = heading Then
				GetFooter = footermulti(i)
				Exit Function
			End If
		Next
	End Function

	Private Function LoopInput(input, footer)
		Dim footers()
		LoopInput = input
		ReDim footers(1)
		footers(0) = footer
		Do
			WScript.StdOut.Write promptmulti
			line = WScript.StdIn.ReadLine
			LoopInput = LoopInput & vbCrLf & line
			nextFooter = GetFooter(line)
			If IsEmpty(nextFooter) Then
				If UCase(Left(LTrim(line), Len(footer))) = footer Then
					count = UBound(footers) - 1
					ReDim footers(count)
					If count > 0 Then
						footer = footers(count - 1)
					End If
				End If
			Else
				count = UBound(footers) + 1
				ReDim footers(count)
				footer = nextFooter
				footers(count - 1) = footer
			End If
		Loop Until UBound(footers) = 0
	End Function

	Private Sub Class_Initialize
		headingmulti = Array("_", "DO UNTIL", "DO WHILE", "FOR", "FOR EACH", "FUNCTION", "IF", "SELECT", "SUB", "WHILE")
		footermulti = Array("", "LOOP", "LOOP", "NEXT", "NEXT", "END FUNCTION", "END IF", "END SELECT", "END SUB", "WEND")
		countmulti = UBound(headingmulti)
		cmdexit = "STOP"
		prompt = "VBSH> "
		promptmulti = "...   "
		executing = True
		While executing:
			WScript.StdOut.Write prompt
			input = WScript.StdIn.ReadLine
			Select Case UCase(Trim(input))
				Case cmdexit
					executing = False
				Case Else
					footer = GetFooter(input)
					cmd = input
					If Not IsEmpty(footer) Then
						cmd = LoopInput(input, footer)
					End If
					On Error Resume Next
					WScript.StdOut.Write Execute(cmd)
					If Err.Number <> 0 Then
						PrintError
						Err.Clear
					End If
					On Error GoTo 0
			End Select
		Wend
	End Sub
End Class
Sub VBSH()
	Set sh = New NamespaceVBSH
End Sub
VBSH
