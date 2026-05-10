#tag Module
Protected Module XLSXFormatter
	#tag Method, Flags = &h0, Description = 436F6E7665727420616E20457863656C20646174652073657269616C206E756D62657220746F2061204461746554696D6520696E20746865206C6F63616C2074696D657A6F6E652E20457863656C2065706F636820697320313839392D31322D333020746F206D6174636820457863656C2773204C6F7475732D312D322D332D6C6561702D796561722D627567206265686176696F722E0A
		Function ExcelSerialToDateTime(serial As Double) As DateTime
		  ' Excel epoch is 1899-12-30 00:00:00 (compensates for Lotus 1-2-3 leap-year bug).
		  Var epoch As New DateTime(1899, 12, 30, 0, 0, 0, 0, TimeZone.Current)
		  Var seconds As Double = serial * 86400.0
		  Return New DateTime(epoch.SecondsFrom1970 + seconds, TimeZone.Current)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4465746563742077686574686572206120666F726D617420636F6465206361727269657320646174652F74696D6520746F6B656E732028792F6D2F642F682F7329206F7574736964652071756F746564206C69746572616C732E2052657475726E732046616C736520666F72202222202F202247656E6572616C22202F2070757265206E756D6572696320666F726D6174732E0A
		Function IsDateFormatCode(code As String) As Boolean
		  ' Detects y/m/d/h/s tokens outside quoted "..." literals.
		  ' Returns True for codes like "dd/mm/yyyy", "hh:mm:ss", "yyyy-mm-dd hh:mm".
		  ' Returns False for plain numeric codes ("0.00", "#,##0", "0%").
		  If code = "" Or code = "General" Then Return False
		  Var inQuote As Boolean = False
		  For i As Integer = 0 To code.Length - 1
		    Var c As String = code.Middle(i, 1)
		    If c = """" Then
		      inQuote = Not inQuote
		      Continue
		    End If
		    If inQuote Then Continue
		    Select Case c.Lowercase
		    Case "y", "d", "h", "s", "m"
		      ' "m" is ambiguous (month vs minute) but in either case this is a date/time format.
		      Return True
		    End Select
		  Next
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 5475726E2061207261772063656C6C2076616C7565202B2063656C6C2074797065202B20457863656C20666F726D617420636F646520696E746F2074686520646973706C61792D726561647920746578742E20526F75746573207468726F75676820466F726D61744E756D62657256616C7565202F20466F726D61744461746556616C7565206261736564206F6E207479706520616E6420497344617465466F726D6174436F64652E0A
		Function Format(rawValue As String, cellType As XLSXEnums.eCellType, formatCode As String) As String
		  Select Case cellType
		  Case XLSXEnums.eCellType.Empty
		    Return ""
		  Case XLSXEnums.eCellType.Str
		    Return rawValue
		  Case XLSXEnums.eCellType.Bool
		    If rawValue = "1" Then Return "TRUE"
		    Return "FALSE"
		  Case XLSXEnums.eCellType.ErrorVal
		    Return rawValue
		  Case XLSXEnums.eCellType.Number, XLSXEnums.eCellType.FormulaCached, XLSXEnums.eCellType.DateValue
		    If rawValue = "" Then Return ""
		    Var d As Double = rawValue.ToDouble
		    If IsDateFormatCode(formatCode) Or cellType = XLSXEnums.eCellType.DateValue Then
		      Var dt As DateTime = ExcelSerialToDateTime(d)
		      Return FormatDateValue(dt, formatCode)
		    Else
		      Return FormatNumberValue(d, formatCode)
		    End If
		  End Select
		  Return rawValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4170706C79206120707261676D6174696320737562736574206F6620457863656C206E756D6572696320666F726D617420636F6465732E20556E6B6E6F776E20636F6465732066616C6C206261636B20746F20446F75626C652E546F537472696E672E0A
		Private Function FormatNumberValue(d As Double, formatCode As String) As String
		  ' Pragmatic subset of Excel number format codes; unknown codes pass through as plain string.
		  Select Case formatCode
		  Case "", "General"
		    Return d.ToString
		  Case "0"
		    Return d.ToString("0")
		  Case "0.00"
		    Return d.ToString("0.00")
		  Case "#,##0"
		    Return d.ToString("#,##0")
		  Case "#,##0.00"
		    Return d.ToString("#,##0.00")
		  Case "0%"
		    Var pct As Double = d * 100.0
		    Return pct.ToString("0") + "%"
		  Case "0.00%"
		    Var pct As Double = d * 100.0
		    Return pct.ToString("0.00") + "%"
		  Else
		    Return d.ToString
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 4170706C79206120707261676D6174696320737562736574206F6620457863656C20646174652F74696D6520666F726D617420636F6465732E20556E6B6E6F776E20636F6465732066616C6C206261636B20746F204461746554696D652E53514C4461746554696D652E0A
		Private Function FormatDateValue(dt As DateTime, formatCode As String) As String
		  ' Pragmatic subset of Excel date/time format codes. Unknown codes fall back to SQLDateTime.
		  Var lc As String = formatCode.Lowercase
		  Select Case lc
		  Case "yyyy-mm-dd"
		    Return dt.SQLDate
		  Case "dd/mm/yyyy"
		    Return Pad2(dt.Day) + "/" + Pad2(dt.Month) + "/" + dt.Year.ToString
		  Case "hh:mm"
		    Return Pad2(dt.Hour) + ":" + Pad2(dt.Minute)
		  Case "hh:mm:ss"
		    Return Pad2(dt.Hour) + ":" + Pad2(dt.Minute) + ":" + Pad2(dt.Second)
		  Case "yyyy-mm-dd hh:mm"
		    Return dt.SQLDate + " " + Pad2(dt.Hour) + ":" + Pad2(dt.Minute)
		  Else
		    Return dt.SQLDateTime
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 54776F2D6469676974207A65726F2D70616464656420646563696D616C20737472696E6720666F7220616E20696E746567657220302E2E39392E0A
		Private Function Pad2(n As Integer) As String
		  If n < 10 Then Return "0" + n.ToString
		  Return n.ToString
		End Function
	#tag EndMethod

	#tag Note, Name = About
		Pure-function helpers that turn raw cell values + Excel format codes into
		display-ready strings.

		Public surface:
		  Format(rawValue, cellType, formatCode)         -> String
		  ExcelSerialToDateTime(serial As Double)         -> DateTime
		  IsDateFormatCode(code As String)                -> Boolean

		Excel's epoch is 1899-12-30 (off by one day from the Lotus 1-2-3 leap-year
		bug — that is the value Excel actually uses, so we match it).

		Format-code subset supported in V1:
		  numbers : "" / "General" / "0" / "0.00" / "#,##0" / "#,##0.00"
		            / "0%" / "0.00%"
		  dates   : "yyyy-mm-dd" / "dd/mm/yyyy" / "hh:mm" / "hh:mm:ss"
		            / "yyyy-mm-dd hh:mm"

		Anything else falls through to a sensible default (Double.ToString /
		DateTime.SQLDateTime). To extend support, add a Case branch in
		FormatNumberValue / FormatDateValue.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Module
#tag EndModule
