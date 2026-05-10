#tag Module
Protected Module XLSXCellRef
	#tag Method, Flags = &h0, Description = 436F6E7665727420636F6C756D6E206C65747465727320746F206120312D626173656420636F6C756D6E20696E6465782E20224122202D3E20312C20225A22202D3E2032362C2022414122202D3E2032372E2052657475726E732030206F6E20696E76616C696420696E7075742E0A
		Function ColLettersToIndex(letters As String) As Integer
		  Var n As Integer = 0
		  Var up As String = letters.Uppercase
		  For i As Integer = 0 To up.Length - 1
		    Var c As String = up.Middle(i, 1)
		    Var v As Integer = c.Asc - Asc("A") + 1
		    If v < 1 Or v > 26 Then Return 0
		    n = n * 26 + v
		  Next
		  Return n
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 436F6E76657274206120312D626173656420636F6C756D6E20696E64657820746F206C6574746572732E2031202D3E202241222C203237202D3E20224141222E2052657475726E7320656D70747920737472696E6720666F7220636F6C203C20312E0A
		Function IndexToColLetters(col As Integer) As String
		  If col < 1 Then Return ""
		  Var s As String = ""
		  Var n As Integer = col
		  While n > 0
		    Var r As Integer = (n - 1) Mod 26
		    s = Chr(Asc("A") + r) + s
		    n = (n - 1) \ 26
		  Wend
		  Return s
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 53706C697420616E204131207265666572656E636520282241423132222920696E746F20312D626173656420726F7720616E6420636F6C2E2052657475726E732046616C7365206F6E206D616C666F726D656420696E7075742E0A
		Function A1ToRowCol(a1 As String, ByRef row As Integer, ByRef col As Integer) As Boolean
		  Var letters As String = ""
		  Var digits As String = ""
		  For i As Integer = 0 To a1.Length - 1
		    Var c As String = a1.Middle(i, 1)
		    If c >= "0" And c <= "9" Then
		      digits = a1.Middle(i)
		      Exit For
		    Else
		      letters = letters + c
		    End If
		  Next
		  If letters = "" Or digits = "" Then Return False
		  col = ColLettersToIndex(letters)
		  row = Integer.FromString(digits)
		  Return col > 0 And row > 0
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Conversion helpers between Excel A1 cell references ("AB12") and 1-based
		(row, col) integers.
		
		Public functions:
		  ColLettersToIndex("AA")               -> 27
		  IndexToColLetters(27)                 -> "AA"
		  A1ToRowCol("AB12", row, col)          -> True; row=12, col=28
		
		All indices are 1-based to match Excel's convention. ColLettersToIndex returns 0
		on invalid input rather than raising.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
