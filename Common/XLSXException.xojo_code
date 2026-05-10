#tag Class
Protected Class XLSXException
Inherits RuntimeException
	#tag Method, Flags = &h0, Description = 4275696C6420616E20657863657074696F6E207769746820612070617273652D6572726F7220636F646520616E6420616E206F7074696F6E616C2064657461696C20737472696E672E20546865204D6573736167652070726F706572747920697320736574206175746F6D61746963616C6C79207573696E67207468652068656C70657220636F64652E546F537472696E672E0A
		Sub Constructor(code As XLSXEnums.eParseError, detail As String = "")
		  Super.Constructor
		  Me.Code = code
		  Me.Detail = detail
		  Me.Message = "XLSXException(" + code.ToString + "): " + detail
		End Sub
	#tag EndMethod

	#tag Property, Flags = &h0, Description = 5468652063617465676F7279206F66207061727365206572726F722E0A
		Code As XLSXEnums.eParseError
	#tag EndProperty

	#tag Property, Flags = &h0, Description = 467265652D666F726D20636F6E746578743A2066696C656E616D652C206D697373696E67207061727420706174682C206D616C666F726D65642D584D4C206C6F636174696F6E2C206574632E0A
		Detail As String
	#tag EndProperty

	#tag Note, Name = About
		The single exception class used by the XLSX parser.
		
		Inherits RuntimeException so callers can catch it polymorphically. Carries:
		  Code   : XLSXEnums.eParseError - which category of failure
		  Detail : String                - free-form context (filename, missing part, etc.)
		
		The Constructor builds Me.Message automatically as
		  XLSXException(<Code symbol>): <Detail>
		using the Extends-based ToString helper from XLSXHelpers.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
