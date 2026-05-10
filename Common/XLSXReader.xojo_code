#tag Module
Protected Module XLSXReader
	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D20612066696C65206F6E206469736B2E2052616973657320584C5358457863657074696F6E206F6E206661696C7572652E0A
		Function Open(file As FolderItem) As XLSXWorkbook
		  If file Is Nil Or Not file.Exists Then
		    Raise New XLSXException(XLSXEnums.eParseError.MissingPart, "file does not exist")
		  End If
		  Var zip As XLSXZip = XLSXZip.Open(file)
		  Return OpenFromZip(zip, file.Name)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D2072617720627974657320285765622075706C6F6164292E2054686520736F757263654E616D65206973207573656420666F7220646961676E6F737469637320616E64205549207469746C652E0A
		Function Open(data As MemoryBlock, sourceName As String) As XLSXWorkbook
		  If data Is Nil Or data.Size = 0 Then
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "empty data")
		  End If
		  Var zip As XLSXZip = XLSXZip.Open(data)
		  Return OpenFromZip(zip, sourceName)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 496E7465726E616C20706970656C696E653A20706172736520736861726564537472696E67732C207374796C65732C2072656C732C207468656E2065616368207368656574206C697374656420696E20776F726B626F6F6B2E786D6C2E0A
		Private Function OpenFromZip(zip As XLSXZip, sourceName As String) As XLSXWorkbook
		  If Not zip.HasPart("xl/workbook.xml") Then
		    Raise New XLSXException(XLSXEnums.eParseError.MissingPart, "xl/workbook.xml")
		  End If

		  Var wb As New XLSXWorkbook(sourceName)
		  wb.SharedStrings = ParseSharedStrings(zip.ReadPart("xl/sharedStrings.xml"))
		  wb.Styles = New XLSXStyles(zip.ReadPart("xl/styles.xml"))

		  Var sheetMap As Dictionary = ParseRelsToTargets(zip.ReadPart("xl/_rels/workbook.xml.rels"))
		  Var workbookXml As String = zip.ReadPart("xl/workbook.xml")
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(workbookXml)
		  Catch
		    Raise New XLSXException(XLSXEnums.eParseError.MalformedXML, "xl/workbook.xml")
		  End Try

		  Var sheetNodes As XmlNodeList = doc.Xql("//*[local-name()='sheets']/*[local-name()='sheet']")
		  For i As Integer = 0 To sheetNodes.Length - 1
		    Var sn As XmlNode = sheetNodes.Item(i)
		    Var name As String = sn.GetAttribute("name")
		    Var rid As String = sn.GetAttribute("r:id")
		    If rid = "" Then rid = sn.GetAttribute("id")
		    Var target As String = If(sheetMap.HasKey(rid), sheetMap.Value(rid), "")
		    If target = "" Then Continue

		    ' rels target is relative to the rels owner — for xl/_rels/workbook.xml.rels
		    ' that means relative to xl/. So "worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml".
		    Var partPath As String = target
		    If Not partPath.BeginsWith("/") Then
		      partPath = "xl/" + partPath
		    Else
		      partPath = partPath.Middle(1)
		    End If

		    Var sheetXml As String = zip.ReadPart(partPath)
		    Var sheet As New XLSXSheet(name, i + 1, sheetXml, wb.SharedStrings)
		    wb.AddSheet(sheet)
		  Next

		  Return wb
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 506172736520786C2F736861726564537472696E67732E786D6C20696E746F206120537472696E672829206F66207265736F6C76656420737472696E677320286F6E6520656E74727920706572203C73693E2C206A6F696E696E6720616C6C203C743E206368696C6472656E20666F7220726963682D746578742072756E73292E0A
		Private Function ParseSharedStrings(xml As String) As String()
		  Var arr() As String
		  If xml = "" Then Return arr
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(xml)
		  Catch
		    Raise New XLSXException(XLSXEnums.eParseError.MalformedXML, "xl/sharedStrings.xml")
		  End Try
		  Var siNodes As XmlNodeList = doc.Xql("//*[local-name()='sst']/*[local-name()='si']")
		  For i As Integer = 0 To siNodes.Length - 1
		    arr.Add(JoinSiText(siNodes.Item(i)))
		  Next
		  Return arr
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E636174656E617465207468652074657874206F66206576657279203C743E2064657363656E64616E74206F6620616E203C73693E206E6F646520E280942068616E646C657320626F74682073696E676C65203C743E20616E6420726963682D74657874203C723E3C743E2072756E732E0A
		Private Function JoinSiText(siNode As XmlNode) As String
		  ' <si> may be a single <t> or several <r><t>...</t></r> rich-text runs.
		  Var result As String = ""
		  Var ts As XmlNodeList = siNode.Xql(".//*[local-name()='t']")
		  For i As Integer = 0 To ts.Length - 1
		    Var n As XmlNode = ts.Item(i)
		    If n.FirstChild <> Nil Then result = result + n.FirstChild.Value
		  Next
		  Return result
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 506172736520786C2F5F72656C732F776F726B626F6F6B2E786D6C2E72656C7320696E746F20612044696374696F6E617279206B657965642062792052656C6174696F6E73686970204964202D3E2054617267657420706174682E0A
		Private Function ParseRelsToTargets(relsXml As String) As Dictionary
		  Var d As New Dictionary
		  If relsXml = "" Then Return d
		  Var doc As New XmlDocument
		  Try
		    doc.LoadXml(relsXml)
		  Catch
		    Return d
		  End Try
		  Var rels As XmlNodeList = doc.Xql("//*[local-name()='Relationships']/*[local-name()='Relationship']")
		  For i As Integer = 0 To rels.Length - 1
		    Var n As XmlNode = rels.Item(i)
		    Var id As String = n.GetAttribute("Id")
		    Var target As String = n.GetAttribute("Target")
		    If id <> "" And target <> "" Then d.Value(id) = target
		  Next
		  Return d
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Top-level entry point for the XLSX parser.

		Two overloads:
		  Open(file As FolderItem) As XLSXWorkbook
		    For Desktop file dialogs.
		  Open(data As MemoryBlock, sourceName As String) As XLSXWorkbook
		    For Web uploads (WebFileUploader.UploadComplete provides bytes).

		Both return a fully-populated XLSXWorkbook (sharedStrings + styles +
		every sheet parsed eagerly). Callers can then iterate via SheetCount /
		SheetAt(i) and bind cells to Listbox / WebListbox.

		Errors raise XLSXException with one of XLSXEnums.eParseError:
		  NotAZip       - magic bytes wrong, or empty MemoryBlock
		  MissingPart   - file doesn't exist, or xl/workbook.xml absent
		  MalformedXML  - workbook.xml / sharedStrings.xml / sheetN.xml
		                  failed XmlDocument.LoadXml

		Pipeline (private OpenFromZip):
		  1. ZipReader extracts archive (XLSXZip).
		  2. ParseSharedStrings -> wb.SharedStrings()
		  3. New XLSXStyles(...) -> wb.Styles
		  4. ParseRelsToTargets reads xl/_rels/workbook.xml.rels (rId -> target).
		  5. For each <sheet> in workbook.xml, look up its rId target and
		     construct an XLSXSheet from xl/<target>.

		Out of V1 scope: defined names, external links, theme/colors, formulas
		(cached value only), encrypted workbooks.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
