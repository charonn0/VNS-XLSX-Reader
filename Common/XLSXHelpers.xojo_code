#tag Module
Protected Module XLSXHelpers
	#tag Method, Flags = &h0, Description = 457874656E73696F6E206D6574686F643A2072657475726E73207468652073796D626F6C6963206E616D65206F6620616E206543656C6C54797065202822456D707479222C2022537472222C20224E756D626572222C20224461746556616C7565222C2022426F6F6C222C2022466F726D756C61436163686564222C20224572726F7256616C22292E0A
		Function ToString(Extends t As XLSXEnums.eCellType) As String
		  Select Case t
		  Case XLSXEnums.eCellType.Empty
		    Return "Empty"
		  Case XLSXEnums.eCellType.Str
		    Return "Str"
		  Case XLSXEnums.eCellType.Number
		    Return "Number"
		  Case XLSXEnums.eCellType.DateValue
		    Return "DateValue"
		  Case XLSXEnums.eCellType.Bool
		    Return "Bool"
		  Case XLSXEnums.eCellType.FormulaCached
		    Return "FormulaCached"
		  Case XLSXEnums.eCellType.ErrorVal
		    Return "ErrorVal"
		  End Select
		  Return "Unknown(" + Integer(t).ToString + ")"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5061727365207468652073796D626F6C6963206E616D65206F6620616E206543656C6C547970652E2052657475726E73206543656C6C547970652E456D70747920666F7220756E6B6E6F776E20696E7075742E0A
		Function CellTypeFromString(s As String) As XLSXEnums.eCellType
		  Select Case s
		  Case "Empty"
		    Return XLSXEnums.eCellType.Empty
		  Case "Str"
		    Return XLSXEnums.eCellType.Str
		  Case "Number"
		    Return XLSXEnums.eCellType.Number
		  Case "DateValue"
		    Return XLSXEnums.eCellType.DateValue
		  Case "Bool"
		    Return XLSXEnums.eCellType.Bool
		  Case "FormulaCached"
		    Return XLSXEnums.eCellType.FormulaCached
		  Case "ErrorVal"
		    Return XLSXEnums.eCellType.ErrorVal
		  End Select
		  Return XLSXEnums.eCellType.Empty
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 457874656E73696F6E206D6574686F643A2072657475726E73207468652073796D626F6C6963206E616D65206F6620616E206550617273654572726F722028224E6F74415A6970222C20224D697373696E6750617274222C20224D616C666F726D6564584D4C222C2022456E63727970746564222C2022556E737570706F7274656422292E0A
		Function ToString(Extends e As XLSXEnums.eParseError) As String
		  Select Case e
		  Case XLSXEnums.eParseError.NotAZip
		    Return "NotAZip"
		  Case XLSXEnums.eParseError.MissingPart
		    Return "MissingPart"
		  Case XLSXEnums.eParseError.MalformedXML
		    Return "MalformedXML"
		  Case XLSXEnums.eParseError.Encrypted
		    Return "Encrypted"
		  Case XLSXEnums.eParseError.Unsupported
		    Return "Unsupported"
		  End Select
		  Return "Unknown(" + Integer(e).ToString + ")"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5061727365207468652073796D626F6C6963206E616D65206F6620616E206550617273654572726F722E2052657475726E73206550617273654572726F722E556E737570706F7274656420666F7220756E6B6E6F776E20696E7075742E0A
		Function ParseErrorFromString(s As String) As XLSXEnums.eParseError
		  Select Case s
		  Case "NotAZip"
		    Return XLSXEnums.eParseError.NotAZip
		  Case "MissingPart"
		    Return XLSXEnums.eParseError.MissingPart
		  Case "MalformedXML"
		    Return XLSXEnums.eParseError.MalformedXML
		  Case "Encrypted"
		    Return XLSXEnums.eParseError.Encrypted
		  Case "Unsupported"
		    Return XLSXEnums.eParseError.Unsupported
		  End Select
		  Return XLSXEnums.eParseError.Unsupported
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToString(Extends m As XLSXEnums.eOpenMode) As String
		  Select Case m
		  Case XLSXEnums.eOpenMode.Auto
		    Return "Auto"
		  Case XLSXEnums.eOpenMode.Memory
		    Return "Memory"
		  Case XLSXEnums.eOpenMode.Disk
		    Return "Disk"
		  End Select
		  Return "Unknown(" + Integer(m).ToString + ")"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5061727365207468652073796D626F6C6963206E616D65206F6620616E20654F70656E4D6F64652E2052657475726E7320654F70656E4D6F64652E4175746F20666F7220756E6B6E6F776E20696E7075742E0A
		Function OpenModeFromString(s As String) As XLSXEnums.eOpenMode
		  Select Case s
		  Case "Auto"
		    Return XLSXEnums.eOpenMode.Auto
		  Case "Memory"
		    Return XLSXEnums.eOpenMode.Memory
		  Case "Disk"
		    Return XLSXEnums.eOpenMode.Disk
		  End Select
		  Return XLSXEnums.eOpenMode.Auto
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Project-wide helpers for the XLSX parser. Add new helper functions here rather
		than scattering them across the modules whose types they support.
		
		Conventions:
		  - enum -> string : extension method named ToString, declared
		                     `Function ToString(Extends e As <EnumType>) As String`.
		                     Callers write `someEnum.ToString` directly.
		  - string -> enum : regular function named <EnumName>FromString.
		
		Adding a new enum to XLSXEnums? Immediately add a matching ToString overload
		and FromString helper in this module.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
