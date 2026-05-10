#tag Module
Protected Module XLSXWebListboxFiller
	#tag Method, Flags = &h0, Description = 46696C6C2061205765624C697374426F7820776974682074686520636F6E74656E7473206F6620616E20584C535853686565742E204669727374206E6F6E2D656D70747920726F77206265636F6D657320746865206865616465723B206D65726765642D63656C6C20666F6C6C6F776572732072656E64657220626C616E6B2E0A
		Sub Fill(lb As WebListBox, sheet As XLSXSheet, styles As XLSXStyles)
		  lb.RemoveAllRows
		  Var cols As Integer = Max(1, sheet.ColCount)
		  lb.ColumnCount = cols

		  Var headerRow As Integer = FindFirstNonEmptyRow(sheet)
		  If headerRow = 0 Then Return

		  ' First non-empty row -> header.
		  For c As Integer = 1 To cols
		    lb.HeaderAt(c - 1) = sheet.CellAt(headerRow, c).DisplayText(styles)
		  Next

		  ' Body rows. Build texts first so we can skip rows where every cell
		  ' is empty (Excel often leaves styled-but-empty rows that inflate RowCount).
		  For r As Integer = headerRow + 1 To sheet.RowCount
		    Var rowTexts() As String
		    Var anyNonEmpty As Boolean = False
		    For c As Integer = 1 To cols
		      Var text As String
		      If sheet.IsCellMergedFollower(r, c) Then
		        text = ""
		      Else
		        text = sheet.CellAt(r, c).DisplayText(styles)
		      End If
		      rowTexts.Add text
		      If text <> "" Then anyNonEmpty = True
		    Next
		    If Not anyNonEmpty Then Continue
		    lb.AddRow("")
		    Var lbRow As Integer = lb.RowCount - 1
		    For c As Integer = 0 To cols - 1
		      lb.CellTextAt(lbRow, c) = rowTexts(c)
		    Next
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5363616E2074686520666972737420353020726F7773206F662074686520736865657420616E642072657475726E2074686520666972737420726F7720696E64657820746861742068617320616E79206E6F6E2D656D7074792063656C6C2E2052657475726E73203120696620616C6C2070726F62656420726F77732061726520656D7074792E0A
		Private Function FindFirstNonEmptyRow(sheet As XLSXSheet) As Integer
		  Var maxProbe As Integer = Min(sheet.RowCount, 50)
		  For r As Integer = 1 To maxProbe
		    For c As Integer = 1 To sheet.ColCount
		      If Not sheet.CellAt(r, c).IsEmpty Then Return r
		    Next
		  Next
		  Return 1
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Pours one XLSXSheet into a WebListBox.

		Public:
		  Fill(lb, sheet, styles)
		    - Resets the listbox (RemoveAllRows + sets ColumnCount).
		    - Uses the first non-empty row of the sheet as the header.
		    - For each subsequent row, builds a String() of column values
		      and AddRows it. Cells inside a merged range that are NOT the
		      top-left anchor render as empty.

		Lives in the Web project (not Common/) because WebListBox is a
		Web-only type. The Desktop project has its own filler at
		VNS-Desktop-XLSX_Reader/XLSXDesktopListboxFiller.xojo_code.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
