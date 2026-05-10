#tag Module
Protected Module XLSXEnums
	#tag Enum, Name = eCellType, Type = Integer, Flags = &h0, Description = 546865206B696E64206F662076616C75652073746F72656420696E206120776F726B73686565742063656C6C2E0A
		Empty = 0
		  Str = 1
		  Number = 2
		  DateValue = 3
		  Bool = 4
		  FormulaCached = 5
		ErrorVal = 6
	#tag EndEnum

	#tag Enum, Name = eParseError, Type = Integer, Flags = &h0, Description = 4572726F722063617465676F726965732072616973656420696E7369646520616E20584C5358457863657074696F6E2E0A
		NotAZip = 0
		  MissingPart = 1
		  MalformedXML = 2
		  Encrypted = 3
		Unsupported = 4
	#tag EndEnum

	#tag Note, Name = About
		Enums shared by the XLSX parser.
		
		eCellType describes how a worksheet cell value should be interpreted.
		  Empty / Str / Number / DateValue / Bool / FormulaCached / ErrorVal
		
		eParseError tags errors raised by the parser inside an XLSXException:
		  NotAZip       - file does not start with the ZIP magic 50 4B 03 04
		  MissingPart   - required part (e.g. xl/workbook.xml) absent
		  MalformedXML  - an XmlDocument.LoadXml call failed
		  Encrypted     - workbook is in the OLE-wrapped encrypted form
		  Unsupported   - feature out of V1 scope or environment limitation
		
		Always reference cases via the namespace, e.g. XLSXEnums.eCellType.Number.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
