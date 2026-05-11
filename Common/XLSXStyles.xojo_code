#tag Class
Protected Class XLSXStyles
	#tag Method, Flags = &h0, Description = 506172736520746865207374796C65732E786D6C20746578742E20456D70747920696E70757420697320616C6C6F77656420E28094207765207374696C6C2073656564206275696C742D696E206E756D466D74206964732E0A
		Sub Constructor(stylesXml As String)
		  ' stylesXml may be empty when the workbook has no styles.xml part.
		  mNumFmts = New Dictionary
		  mCellXfs = New Dictionary
		  SeedBuiltInNumFmts
		  If stylesXml = "" Then Return
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(stylesXml)
		  Catch
		    Raise New XLSXException(XLSXEnums.eParseError.MalformedXML, "styles.xml")
		  End Try
		  ParseNumFmts(doc)
		  ParseCellXfs(doc)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E732074686520457863656C20666F726D617420636F646520666F722074686520676976656E2063656C6C586620696E6465782C206F7220656D70747920737472696E6720696620616273656E742E0A
		Function NumberFormatCodeAt(cellXfIndex As Integer) As String
		  ' Returns the format code (e.g. "0.00", "dd/mm/yyyy", "General") for the given
		  ' cellXf index, or "" if the index is out of range.
		  If Not mCellXfs.HasKey(cellXfIndex) Then Return ""
		  Var numFmtId As Integer = mCellXfs.Value(cellXfIndex)
		  If mNumFmts.HasKey(numFmtId) Then Return mNumFmts.Value(numFmtId)
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 546F74616C206E756D626572206F66203C78663E20656E7472696573207061727365642066726F6D203C63656C6C5866733E2E0A
		Function CellXfCount() As Integer
		  Return mCellXfs.KeyCount
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 506F70756C61746520746865206275696C742D696E20457863656C206E756D466D74206964732028302E2E343920737562736574292E20437573746F6D20636F6465732066726F6D203C6E756D466D74733E206F766572726964652074686573652E0A
		Private Sub SeedBuiltInNumFmts()
		  ' Excel's built-in numFmtIds 0..49 (subset that we recognize).
		  ' Custom format codes (id >= 164) declared in <numFmts> override these.
		  mNumFmts.Value(0) = "General"
		  mNumFmts.Value(1) = "0"
		  mNumFmts.Value(2) = "0.00"
		  mNumFmts.Value(3) = "#,##0"
		  mNumFmts.Value(4) = "#,##0.00"
		  mNumFmts.Value(9) = "0%"
		  mNumFmts.Value(10) = "0.00%"
		  mNumFmts.Value(14) = "dd/mm/yyyy"
		  mNumFmts.Value(15) = "d-mmm-yy"
		  mNumFmts.Value(16) = "d-mmm"
		  mNumFmts.Value(17) = "mmm-yy"
		  mNumFmts.Value(18) = "h:mm AM/PM"
		  mNumFmts.Value(19) = "h:mm:ss AM/PM"
		  mNumFmts.Value(20) = "hh:mm"
		  mNumFmts.Value(21) = "hh:mm:ss"
		  mNumFmts.Value(22) = "yyyy-mm-dd hh:mm"
		  ' Built-in accounting / negative-in-parens forms (Excel ids 37..44).
		  mNumFmts.Value(37) = "#,##0 ;(#,##0)"
		  mNumFmts.Value(38) = "#,##0 ;[Red](#,##0)"
		  mNumFmts.Value(39) = "#,##0.00;(#,##0.00)"
		  mNumFmts.Value(40) = "#,##0.00;[Red](#,##0.00)"
		  mNumFmts.Value(41) = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
		  mNumFmts.Value(42) = "_(""$""* #,##0_);_(""$""* (#,##0);_(""$""* ""-""_);_(@_)"
		  mNumFmts.Value(43) = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
		  mNumFmts.Value(44) = "_(""$""* #,##0.00_);_(""$""* (#,##0.00);_(""$""* ""-""??_);_(@_)"
		  mNumFmts.Value(48) = "##0.0E+0"
		  mNumFmts.Value(49) = "@"
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365203C6E756D466D74733E2F3C6E756D466D743E20656E74726965732028637573746F6D20666F726D617420636F6465732C20696473203E3D20313634292E0A
		Private Sub ParseNumFmts(doc As XmlDocument)
		  Var nodes As XmlNodeList = doc.Xql("//*[local-name()='numFmts']/*[local-name()='numFmt']")
		  For i As Integer = 0 To nodes.Length - 1
		    Var n As XmlNode = nodes.Item(i)
		    Var idAttr As String = n.GetAttribute("numFmtId")
		    Var codeAttr As String = n.GetAttribute("formatCode")
		    If idAttr <> "" Then
		      mNumFmts.Value(Integer.FromString(idAttr)) = codeAttr
		    End If
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365203C63656C6C5866733E2F3C78663E20656E74726965733B206275696C6420746865206D61702066726F6D207866496E64657820746F206E756D466D7449642E0A
		Private Sub ParseCellXfs(doc As XmlDocument)
		  Var nodes As XmlNodeList = doc.Xql("//*[local-name()='cellXfs']/*[local-name()='xf']")
		  For i As Integer = 0 To nodes.Length - 1
		    Var n As XmlNode = nodes.Item(i)
		    Var fmtAttr As String = n.GetAttribute("numFmtId")
		    Var numFmtId As Integer = If(fmtAttr <> "", Integer.FromString(fmtAttr), 0)
		    mCellXfs.Value(i) = numFmtId
		  Next
		End Sub
	#tag EndMethod

	#tag Property, Flags = &h21, Description = 4D61702066726F6D206E756D466D7449642028496E74656765722920746F20666F726D617420636F64652028537472696E67292E0A
		Private mNumFmts As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4D61702066726F6D2063656C6C586620696E6465782028496E74656765722C20302D62617365642920746F206E756D466D7449642028496E7465676572292E0A
		Private mCellXfs As Dictionary
	#tag EndProperty

	#tag Note, Name = About
		Parses xl/styles.xml from an XLSX archive and resolves a cellXf index
		to a format code.

		Public surface:
		  Constructor(stylesXml As String)
		    Pass the UTF-8 text of styles.xml. Empty input is fine — we still
		    seed built-in numFmt ids.
		  NumberFormatCodeAt(cellXfIndex As Integer) As String
		    Returns "0.00" / "dd/mm/yyyy" / "General" / etc. — feed it to
		    XLSXFormatter.Format().
		  CellXfCount() As Integer
		    Total number of <xf> entries declared in <cellXfs>.

		Resolution chain:
		  cell @s attribute (an integer index into <cellXfs>)
		    -> <xf>.numFmtId
		      -> <numFmts>/<numFmt formatCode=...> if id >= 164
		      -> built-in code from SeedBuiltInNumFmts otherwise
		      -> "" (empty) on miss

		We only read numFmt codes — fonts, fills, borders, alignment are out
		of V1 scope.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
