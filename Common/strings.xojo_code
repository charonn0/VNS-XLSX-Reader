#tag Module
Protected Module strings
	#tag Constant, Name = kStrAppTitle, Type = String, Dynamic = True, Default = \"VNS XLSX Reader", Scope = Public, Description = 4170702077696E646F77202F2070616765207469746C652E0A
	#tag EndConstant

	#tag Constant, Name = kStrMenuFileOpen, Type = String, Dynamic = True, Default = \"Open\xE2\x80\xA6", Scope = Public, Description = 46696C65206D656E75202D204F70656E2E2E2E206974656D206C6162656C2E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorTitle, Type = String, Dynamic = True, Default = \"Cannot open file", Scope = Public, Description = 5469746C65206F6620746865206572726F72206469616C6F672073686F776E207768656E206F70656E696E6720612066696C65206661696C732E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorNotXLSX, Type = String, Dynamic = True, Default = \"This file is not a valid .xlsx workbook.", Scope = Public, Description = 4572726F72206D6573736167653A2066696C65206973206E6F7420612076616C696420584C535820776F726B626F6F6B2E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorMalformed, Type = String, Dynamic = True, Default = \"This workbook has malformed content.", Scope = Public, Description = 4572726F72206D6573736167653A20776F726B626F6F6B20636F6E74656E74206973206D616C666F726D65642E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorEncrypted, Type = String, Dynamic = True, Default = \"Encrypted workbooks are not supported.", Scope = Public, Description = 4572726F72206D6573736167653A20656E6372797074656420776F726B626F6F6B7320617265206E6F7420737570706F727465642E0A
	#tag EndConstant

	#tag Constant, Name = kStrErrorGeneric, Type = String, Dynamic = True, Default = \"Could not read this workbook.", Scope = Public, Description = 4572726F72206D6573736167653A2067656E65726963206661696C757265206E6F7420636F766572656420627920746865206D6F72652073706563696669632063617365732E0A
	#tag EndConstant

	#tag Constant, Name = kStrUploadPrompt, Type = String, Dynamic = True, Default = \"Choose an .xlsx file to view\xE2\x80\xA6", Scope = Public, Description = 5765622066696C652075706C6F616465722070726F6D70742061736B696E6720746865207573657220746F2063686F6F736520616E202E786C73782066696C652E0A
	#tag EndConstant

	#tag Note, Name = About
		Localizable user-visible strings.
		
		All constants are Dynamic so the IDE Language Editor can override them per
		locale at build time. Code must never use hardcoded user-visible strings -
		always reference a kStr... constant from this module.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
