#tag Class
Protected Class XLSXCellRange
	#tag Method, Flags = &h0, Description = 4275696C6420612072616E67652066726F6D2069747320636F726E6572732028616C6C20312D62617365642C20696E636C7573697665292E0A
		Sub Constructor(firstRow As Integer, firstCol As Integer, lastRow As Integer, lastCol As Integer)
		  Me.FirstRow = firstRow
		  Me.FirstCol = firstCol
		  Me.LastRow = lastRow
		  Me.LastCol = lastCol
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E7320547275652069662028726F772C20636F6C292066616C6C7320696E7369646520746869732072616E67652C20696E636C75736976652E0A
		Function Contains(row As Integer, col As Integer) As Boolean
		  Return row >= FirstRow And row <= LastRow And col >= FirstCol And col <= LastCol
		End Function
	#tag EndMethod

	#tag Property, Flags = &h0, Description = 546F7020726F77206F66207468652072616E67652028312D62617365642C20696E636C7573697665292E0A
		FirstRow As Integer
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 4C65667420636F6C756D6E206F66207468652072616E67652028312D62617365642C20696E636C7573697665292E0A
		FirstCol As Integer
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 426F74746F6D20726F77206F66207468652072616E67652028312D62617365642C20696E636C7573697665292E0A
		LastRow As Integer
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 526967687420636F6C756D6E206F66207468652072616E67652028312D62617365642C20696E636C7573697665292E0A
		LastCol As Integer
	#tag EndProperty

	#tag Note, Name = About
		An inclusive 1-based cell range (e.g. A1:C3).
		
		Used to represent merged-cell rectangles parsed from a worksheet's
		<mergeCells> element. Both corners are inclusive.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
