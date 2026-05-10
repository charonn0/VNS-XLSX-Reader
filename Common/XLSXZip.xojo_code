#tag Class
Protected Class XLSXZip
	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D206120466F6C6465724974656D206F6E206469736B2E20566572696669657320746865205A4950206D616769632062797465732028353020344220303320303429206265666F72652065787472616374696E672E0A
		Shared Function Open(file As FolderItem) As XLSXZip
		  If file Is Nil Or Not file.Exists Then
		    Raise New XLSXException(XLSXEnums.eParseError.MissingPart, "file is Nil or does not exist")
		  End If
		  If Not VerifyZipMagic(file) Then
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, file.Name)
		  End If
		  Var z As New XLSXZip
		  z.mSourceFile = file
		  z.mOwnsSourceFile = False
		  z.ExtractToTemp
		  Return z
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D207261772062797465732028652E672E2057656246696C6555706C6F616465722E55706C6F6164436F6D706C6574652064617461292E205370696C6C7320746F20612074656D702066696C65206265636175736520466F6C6465724974656D2E556E7A6970206E65656473206120466F6C6465724974656D206F6E206469736B2E0A
		Shared Function Open(data As MemoryBlock) As XLSXZip
		  If data Is Nil Or data.Size = 0 Then
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "empty data")
		  End If
		  If data.Size < 4 _
		     Or data.Byte(0) <> &h50 _
		     Or data.Byte(1) <> &h4B _
		     Or data.Byte(2) <> &h03 _
		     Or data.Byte(3) <> &h04 Then
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "bad magic in MemoryBlock")
		  End If
		  Var tempFolder As FolderItem = SpecialFolder.Temporary
		  If tempFolder Is Nil Or Not tempFolder.Exists Then
		    Raise New XLSXException(XLSXEnums.eParseError.Unsupported, "SpecialFolder.Temporary unavailable")
		  End If
		  Var tempZip As FolderItem = tempFolder.Child("xlsxzip_" + System.Microseconds.ToString + ".zip")
		  Var bs As BinaryStream = BinaryStream.Create(tempZip, True)
		  Try
		    bs.Write(data)
		  Finally
		    bs.Close
		  End Try
		  Var z As New XLSXZip
		  z.mSourceFile = tempZip
		  z.mOwnsSourceFile = True
		  Try
		    z.ExtractToTemp
		  Catch err As RuntimeException
		    Try
		      If tempZip <> Nil And tempZip.Exists Then tempZip.Remove
		    Catch
		    End Try
		    Raise err
		  End Try
		  Return z
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E73205472756520696620746865206172636869766520636F6E7461696E7320612066696C652061742074686520676976656E2072656C617469766520706174682028636173652D73656E7369746976652C204F50432D7374796C652022786C2F776F726B626F6F6B2E786D6C22292E0A
		Function HasPart(partPath As String) As Boolean
		  Return PartFile(partPath) <> Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E732074686520656E7472792773205554462D3820746578742C206F7220656D70747920737472696E6720696620746865207061727420697320616273656E742E2043616C6C65727320646563696465207768657468657220616273656E636520697320666174616C2E0A
		Function ReadPart(partPath As String) As String
		  Var f As FolderItem = PartFile(partPath)
		  If f Is Nil Then Return ""
		  Var ti As TextInputStream = TextInputStream.Open(f)
		  Try
		    Var s As String = ti.ReadAll(Encodings.UTF8)
		    Return s
		  Finally
		    ti.Close
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 556E7A69702074686520736F75726365206172636869766520696E746F20612066726573686C792D6372656174656420737562666F6C646572206F66205370656369616C466F6C6465722E54656D706F726172792E0A
		Private Sub ExtractToTemp()
		  ' Use the framework FolderItem.Unzip into a fresh per-instance subfolder
		  ' under SpecialFolder.Temporary. Destructor cleans it up.
		  Var tempRoot As FolderItem = SpecialFolder.Temporary
		  If tempRoot Is Nil Or Not tempRoot.Exists Then
		    Raise New XLSXException(XLSXEnums.eParseError.Unsupported, "SpecialFolder.Temporary unavailable")
		  End If
		  Var dest As FolderItem = tempRoot.Child("xlsxzip_extract_" + System.Microseconds.ToString)
		  If Not dest.Exists Then dest.CreateFolder
		  Try
		    mSourceFile.Unzip(dest)
		  Catch err As RuntimeException
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "Unzip failed: " + err.Message)
		  End Try
		  mExtractRoot = dest
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265736F6C7665206120706172742070617468206C696B652022786C2F776F726B626F6F6B2E786D6C2220746F206120466F6C6465724974656D20696E73696465207468652065787472616374656420747265652E2052657475726E73204E696C20696620616E79207365676D656E74206973206D697373696E672E0A
		Private Function PartFile(partPath As String) As FolderItem
		  ' Resolve "xl/workbook.xml" → mExtractRoot.Child("xl").Child("workbook.xml").
		  ' Returns Nil if any segment doesn't exist.
		  If mExtractRoot Is Nil Then Return Nil
		  Var clean As String = partPath.ReplaceAll("\", "/")
		  If clean.BeginsWith("/") Then clean = clean.Middle(1)
		  Var segments() As String = clean.Split("/")
		  Var cur As FolderItem = mExtractRoot
		  For Each seg As String In segments
		    If seg = "" Then Continue
		    cur = cur.Child(seg)
		    If cur Is Nil Or Not cur.Exists Then Return Nil
		  Next
		  If cur.IsFolder Then Return Nil
		  Return cur
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 52656164207468652066697273742034206279746573206F66206120466F6C6465724974656D20616E64207665726966792074686579206D6174636820746865205A4950206D616769632035302034422030332030342E0A
		Private Shared Function VerifyZipMagic(file As FolderItem) As Boolean
		  If file Is Nil Or Not file.Exists Then Return False
		  Var bs As BinaryStream
		  Try
		    bs = BinaryStream.Open(file, False)
		  Catch
		    Return False
		  End Try
		  Try
		    If bs.Length < 4 Then Return False
		    Var raw As String = bs.Read(4)
		    Var mb As MemoryBlock = raw
		    Return mb.Byte(0) = &h50 _
		      And mb.Byte(1) = &h4B _
		      And mb.Byte(2) = &h03 _
		      And mb.Byte(3) = &h04
		  Finally
		    bs.Close
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265637572736976656C792064656C657465206120666F6C64657220616E642069747320636F6E74656E74732C207377616C6C6F77696E6720616E79206572726F722E20557365642062792044657374727563746F722E0A
		Private Shared Sub RemoveTreeBestEffort(f As FolderItem)
		  If f Is Nil Or Not f.Exists Then Return
		  If f.IsFolder Then
		    Var n As Integer = f.Count
		    For i As Integer = n - 1 DownTo 0
		      RemoveTreeBestEffort(f.ChildAt(i))
		    Next
		  End If
		  Try
		    f.Remove
		  Catch
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 436C65616E2075703A2064656C65746520746865206578747261637465642074656D7020666F6C6465722C20616E64207468652074656D7020736F757263652066696C65206966207765206F776E2069742028637265617465642066726F6D2061204D656D6F7279426C6F636B292E0A
		Sub Destructor()
		  If mExtractRoot <> Nil Then RemoveTreeBestEffort(mExtractRoot)
		  mExtractRoot = Nil
		  If mOwnsSourceFile And mSourceFile <> Nil And mSourceFile.Exists Then
		    Try
		      mSourceFile.Remove
		    Catch
		    End Try
		  End If
		  mSourceFile = Nil
		End Sub
	#tag EndMethod

	#tag Property, Flags = &h21, Description = 54686520736F757263652061726368697665206F6E206469736B202D2065697468657220757365722D737570706C696564206F7220612074656D702066696C6520776520637265617465642E0A
		Private mSourceFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 54727565207768656E2077652063726561746564206D536F7572636546696C65202866726F6D2061204D656D6F7279426C6F636B2920616E64206D7573742072656D6F766520697420696E2044657374727563746F722E0A
		Private mOwnsSourceFile As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 5468652074656D7020666F6C64657220776865726520746865206172636869766520776173206578747261637465642E2052656D6F766564207265637572736976656C7920696E2044657374727563746F722E0A
		Private mExtractRoot As FolderItem
	#tag EndProperty

	#tag Note, Name = About
		Reads named parts (e.g. xl/workbook.xml) from an XLSX or any ZIP archive,
		using the framework's FolderItem.Unzip into a per-instance subfolder of
		SpecialFolder.Temporary.
		
		Two entry points:
		  Open(file As FolderItem)  - for the Desktop file dialog
		  Open(data As MemoryBlock) - for the Web file uploader (writes a temp .zip first)
		
		Both validate the ZIP magic 50 4B 03 04 before unzipping. The destructor
		recursively removes the extracted folder and the temp .zip if we created one.
		
		Read API:
		  HasPart("xl/workbook.xml") As Boolean
		  ReadPart("xl/workbook.xml") As String   - returns UTF-8 text, or empty
		                                            string if the part is absent.
		                                            Callers decide whether absence
		                                            is fatal.
	#tag EndNote

	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
