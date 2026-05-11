#tag Class
Protected Class XLSXZip
	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D206120466F6C6465724974656D206F6E206469736B2E20546865206D6F646520706172616D65746572207069636B7320746865206261636B656E643A204175746F202864656661756C74292C204D656D6F72792C206F72204469736B2E0A
		Shared Function Open(file As FolderItem, mode As XLSXEnums.eOpenMode = XLSXEnums.eOpenMode.Auto) As XLSXZip
		  If file Is Nil Or Not file.Exists Then
		    Raise New XLSXException(XLSXEnums.eParseError.MissingPart, "file is Nil or does not exist")
		  End If
		  Var z As New XLSXZip
		  z.mMode = ResolveMode(mode)
		  Select Case z.mMode
		  Case XLSXEnums.eOpenMode.Memory
		    Var bs As BinaryStream = BinaryStream.Open(file, False)
		    Var data As MemoryBlock = bs.Read(bs.Length)
		    bs.Close
		    z.BuildFromMemoryData(data)
		  Case XLSXEnums.eOpenMode.Disk
		    z.BuildFromDiskFile(file, False)
		  End Select
		  Return z
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D207261772062797465732E20546865206D6F646520706172616D65746572207069636B7320746865206261636B656E643B20756E646572204469736B206D6F646520746865206279746573207370696C6C20746F20612074656D70202E7A69702066697273742E0A
		Shared Function Open(data As MemoryBlock, mode As XLSXEnums.eOpenMode = XLSXEnums.eOpenMode.Auto) As XLSXZip
		  If data Is Nil Or data.Size = 0 Then
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "empty data")
		  End If
		  Var z As New XLSXZip
		  z.mMode = ResolveMode(mode)
		  Select Case z.mMode
		  Case XLSXEnums.eOpenMode.Memory
		    z.BuildFromMemoryData(data)
		  Case XLSXEnums.eOpenMode.Disk
		    ' FolderItem.Unzip needs a file on disk, so spill the MemoryBlock into
		    ' SpecialFolder.Temporary, then unzip that file.
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
		    Try
		      z.BuildFromDiskFile(tempZip, True)
		    Catch err As RuntimeException
		      Try
		        If tempZip <> Nil And tempZip.Exists Then tempZip.Remove
		      Catch
		      End Try
		      Raise err
		    End Try
		  End Select
		  Return z
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E73205472756520696620746865206172636869766520636F6E7461696E7320612066696C652061742074686520676976656E2072656C617469766520706174682028636173652D73656E7369746976652C204F50432D7374796C652022786C2F776F726B626F6F6B2E786D6C22292E0A
		Function HasPart(partPath As String) As Boolean
		  Select Case mMode
		  Case XLSXEnums.eOpenMode.Memory
		    Return PartFileMemory(partPath) <> Nil
		  Case XLSXEnums.eOpenMode.Disk
		    Return PartFileDisk(partPath) <> Nil
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E732074686520656E7472792773205554462D3820746578742C206F7220656D70747920737472696E6720696620746865207061727420697320616273656E742E2043616C6C65727320646563696465207768657468657220616273656E636520697320666174616C2E0A
		Function ReadPart(partPath As String) As String
		  Select Case mMode
		  Case XLSXEnums.eOpenMode.Memory
		    Var fileData As MemoryBlock = PartFileMemory(partPath)
		    If fileData Is Nil Then Return ""
		    Return DefineEncoding(fileData, Encodings.UTF8)
		  Case XLSXEnums.eOpenMode.Disk
		    Var f As FolderItem = PartFileDisk(partPath)
		    If f Is Nil Then Return ""
		    Var ti As TextInputStream = TextInputStream.Open(f)
		    Try
		      Return ti.ReadAll(Encodings.UTF8)
		    Finally
		      ti.Close
		    End Try
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 546865207265736F6C766564206261636B656E642075736564206174204F70656E2074696D6520286E65766572204175746F206166746572204F70656E292E0A
		Function Mode() As XLSXEnums.eOpenMode
		  Return mMode
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 436F6E76657274204175746F20746F204D656D6F7279206F6E20586F6A6F203230323472332B2C20656C7365204469736B2E20506173732D7468726F75676820666F72206578706C69636974206D6F6465732E0A
		Private Shared Function ResolveMode(mode As XLSXEnums.eOpenMode) As XLSXEnums.eOpenMode
		  If mode <> XLSXEnums.eOpenMode.Auto Then Return mode
		  ' Memory backend needs MemoryBlock.Decompress (added in Xojo 2024r3).
		  #If XojoVersion > 2024.02 Then
		    Return XLSXEnums.eOpenMode.Memory
		  #Else
		    Return XLSXEnums.eOpenMode.Disk
		  #EndIf
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5061727365207A6970206C6F63616C2D66696C652D686561646572732066726F6D2061204D656D6F7279426C6F636B20696E746F206D5A69704469726563746F7279206B6579656420627920706174682E204D656D6F72792D6261636B656E642073657475702E0A
		Private Sub BuildFromMemoryData(zipData As MemoryBlock)
		  ' Parses the zip's local file headers, builds mZipDirectory keyed by path.
		  ' (charonn0's implementation from PR #1.)
		  Const ZIP_HEADER_SIG = &h04034b50
		  Const ZIP_FOOTER_SIG = &h08074b50
		  Const FLAG_NAME_UTF = 2048
		  Const FLAG_TRAILER = 8

		  mZipData = zipData
		  mZipDirectory = New Dictionary
		  Var stream As New BinaryStream(zipData)
		  stream.LittleEndian = True

		  Do Until stream.EOF
		    If Not SeekSignature(stream, ZIP_HEADER_SIG) Then Exit Do
		    stream.Position = stream.Position + 4

		    Var entry As New Dictionary
		    entry.Value("version") = stream.ReadUInt16()
		    Var flags As UInt16 = stream.ReadUInt16()
		    entry.Value("flags") = flags
		    entry.Value("method") = stream.ReadUInt16()
		    entry.Value("modtime") = stream.ReadUInt16()
		    entry.Value("moddate") = stream.ReadUInt16()
		    entry.Value("crc") = stream.ReadUInt32()
		    Var compSize As UInt32 = stream.ReadUInt32()
		    entry.Value("compSize") = compSize
		    entry.Value("origSize") = stream.ReadUInt32()
		    Var namelen As UInt16 = stream.ReadUInt16()
		    Var extralen As UInt16 = stream.ReadUInt16()
		    Var path As String = stream.Read(namelen)
		    entry.Value("extra") = stream.Read(extralen)

		    If BitAnd(flags, FLAG_NAME_UTF) = FLAG_NAME_UTF Then
		      path = DefineEncoding(path, Encodings.UTF8).Trim
		    Else
		      path = DefineEncoding(path, Encodings.DOSLatinUS).Trim
		    End If

		    entry.Value("offset") = stream.Position

		    If BitAnd(flags, FLAG_TRAILER) = FLAG_TRAILER And compSize = 0 Then
		      If Not SeekSignature(stream, ZIP_FOOTER_SIG) Then
		        Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "Missing metadata trailer.")
		      End If
		      stream.Position = stream.Position + 4
		      entry.Value("crc") = stream.ReadUInt32()
		      compSize = stream.ReadUInt32()
		      entry.Value("compSize") = compSize
		      entry.Value("origSize") = stream.ReadUInt32()
		    End If

		    mZipDirectory.Value(path) = entry
		  Loop
		  stream.Close
		  If mZipDirectory.KeyCount = 0 Then
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "no entries found")
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 566572696679205A4950206D61676963207468656E20466F6C6465724974656D2E556E7A697020696E746F206120667265736820737562666F6C646572206F66205370656369616C466F6C6465722E54656D706F726172792E204469736B2D6261636B656E642073657475702E206F776E54656D7046696C653D54727565206966207468652063616C6C65722063726561746564207468652066696C65202869742077696C6C2062652072656D6F76656420696E2044657374727563746F72292E0A
		Private Sub BuildFromDiskFile(file As FolderItem, ownTempFile As Boolean)
		  ' FolderItem.Unzip into a fresh per-instance subfolder under SpecialFolder.Temporary.
		  If Not VerifyZipMagic(file) Then
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, file.Name)
		  End If
		  mSourceFile = file
		  mOwnsSourceFile = ownTempFile
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

	#tag Method, Flags = &h21, Description = 4465636F6D707265737320616E20696E2D6D656D6F727920656E74727920746F2061204D656D6F7279426C6F636B2E2052657475726E73204E696C20696620616273656E743B2072616973657320584C5358457863657074696F6E206F6E20756E737570706F7274656420636F6D7072657373696F6E206D6574686F64732E0A
		Private Function PartFileMemory(partPath As String) As MemoryBlock
		  ' Decompress an in-memory entry. Returns Nil if the path doesn't exist.
		  Var clean As String = partPath.ReplaceAll("\", "/")
		  If clean.BeginsWith("/") Then clean = clean.Middle(1)
		  Var entry As Dictionary = mZipDirectory.Lookup(clean, Nil)
		  If entry = Nil Then Return Nil

		  mZipData.LittleEndian = True
		  Var offset As UInt64 = entry.Value("offset")
		  Var length As UInt32 = entry.Value("compSize")
		  Var method As UInt16 = entry.Value("method")

		  Select Case method
		  Case 0 ' stored, not compressed
		    Return mZipData.StringValue(offset, length)
		    #If XojoVersion > 2024.02 Then
		  Case 8 ' deflated
		    ' Xojo 2024r3 added MemoryBlock.Decompress for the GZIP format. The zip
		    ' format uses raw deflate, which is GZIP minus a small header — so we
		    ' wrap the raw deflate bytes in a synthetic GZIP header and feed it in.
		    Var raw As New MemoryBlock(0)
		    Var bs As New BinaryStream(raw)
		    bs.LittleEndian = True
		    bs.WriteUInt16(&h8B1F) ' magic
		    bs.WriteUInt8(8)       ' method
		    bs.WriteUInt8(0)       ' flags
		    bs.WriteUInt32(0)      ' mtime
		    bs.WriteUInt8(0)       ' extra flags
		    bs.WriteUInt8(0)       ' OS
		    bs.Write(mZipData.StringValue(offset, length))
		    bs.WriteUInt32(entry.Value("crc"))
		    bs.WriteUInt32(entry.Value("origSize") Mod (2^32))
		    bs.Close
		    Return raw.Decompress
		    #EndIf
		  Else
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "Unsupported compression algorithm: " + Str(method))
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265736F6C7665206120706172742070617468206C696B652022786C2F776F726B626F6F6B2E786D6C2220746F206120466F6C6465724974656D20756E646572206D45787472616374526F6F742E2052657475726E73204E696C20696620616E79207365676D656E74206973206D697373696E672E0A
		Private Function PartFileDisk(partPath As String) As FolderItem
		  ' Resolve "xl/workbook.xml" → mExtractRoot.Child("xl").Child("workbook.xml").
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

	#tag Method, Flags = &h21, Description = 4C6F63617465206120342D62797465207369676E617475726520696E20612042696E61727953747265616D206279206368756E6B6564207363616E2E2052657475726E73205472756520776974682053747265616D2E506F736974696F6E20617420746865206D617463682C20656C73652046616C736520776974682053747265616D2E506F736974696F6E20726573746F7265642E0A
		Private Shared Function SeekSignature(stream As BinaryStream, signature As UInt32) As Boolean
		  ' Locate Signature in Stream. On success, Stream.Position points at the signature
		  ' and returns True. On miss, Stream.Position is restored.
		  Var pos As UInt64 = stream.Position
		  Var ok As Boolean
		  Var sig As New MemoryBlock(4)
		  sig.LittleEndian = True
		  sig.UInt32Value(0) = signature

		  Do Until stream.EOF
		    Var data As String = stream.Read(65535)
		    Var offset As Integer = InStrB(data, sig)
		    If offset > 0 Then
		      stream.Position = stream.Position - (data.LenB - offset + 1)
		      ok = True
		      Exit Do
		    ElseIf stream.Length - stream.Position >= 4 Then
		      stream.Position = stream.Position - 3
		    End If
		  Loop
		  If Not ok Then stream.Position = pos
		  Return ok
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265637572736976656C792064656C657465206120666F6C64657220616E642069747320636F6E74656E74732C207377616C6C6F77696E6720616E79206572726F722E20557365642062792044657374727563746F72207768656E20696E204469736B206D6F64652E0A
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

	#tag Method, Flags = &h0, Description = 436C65616E2075703A204469736B206D6F64652072656D6F76657320746865206578747261637465642074656D7020666F6C64657220616E6420286966206F776E7354656D7046696C6529207468652074656D70202E7A69703B204D656D6F7279206D6F6465206A7573742064726F7073206D5A697044617461202F206D5A69704469726563746F72792E0A
		Sub Destructor()
		  If mMode = XLSXEnums.eOpenMode.Disk Then
		    If mExtractRoot <> Nil Then RemoveTreeBestEffort(mExtractRoot)
		    mExtractRoot = Nil
		    If mOwnsSourceFile And mSourceFile <> Nil And mSourceFile.Exists Then
		      Try
		        mSourceFile.Remove
		      Catch
		      End Try
		    End If
		    mSourceFile = Nil
		  End If
		  mZipData = Nil
		  mZipDirectory = Nil
		End Sub
	#tag EndMethod

	#tag Note, Name = About
		Reads named parts (e.g. xl/workbook.xml) from an XLSX or any ZIP archive.

		Two backends, picked at Open time via the optional `mode` parameter:

		  Auto    - Memory on Xojo 2024r3+, else Disk. The default.
		  Memory  - Parse zip directory in-memory, decompress entries via
		            MemoryBlock.Decompress (no temp folder, no disk I/O,
		            sandbox-friendly).
		  Disk    - FolderItem.Unzip into SpecialFolder.Temporary, read parts
		            as FolderItems. The destructor removes the temp folder.

		Public surface:

		  Open(file As FolderItem [, mode]) As XLSXZip
		  Open(data As MemoryBlock [, mode]) As XLSXZip
		  HasPart(partPath As String) As Boolean
		  ReadPart(partPath As String) As String   - returns UTF-8 text, or
		                                            "" if the part is absent.
		                                            Callers decide whether
		                                            absence is fatal.
		  Mode() As XLSXEnums.eOpenMode            - the resolved backend
		                                            (never Auto after Open).

		The two backends share HasPart / ReadPart / Destructor — those branch on
		the resolved mode. The Memory backend keeps the compressed zip bytes in
		mZipData for the lifetime of the instance; the Disk backend keeps the
		extracted folder.

		Memory-backend zip parser by @charonn0 (Andrew Lambert) via PR #1.
	#tag EndNote


	#tag Property, Flags = &h21, Description = 546865207265736F6C766564206261636B656E642063686F73656E206174204F70656E2074696D6520284D656D6F7279206F72204469736B292E0A
		Private mMode As XLSXEnums.eOpenMode
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4D656D6F7279206261636B656E643A2074686520656E7469726520636F6D70726573736564207A69702062797465732C2072657461696E656420666F7220746865206C69666574696D65206F662074686520696E7374616E636520736F205061727446696C654D656D6F72792063616E206465636F6D707265737320656E7472696573206F6E2064656D616E642E0A
		Private mZipData As MemoryBlock
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4D656D6F7279206261636B656E643A2070617468202D3E207B6F66667365742C20636F6D7053697A652C206F72696753697A652C206D6574686F642C206372632C20666C6167737D20666F722065616368206C6F63616C2066696C65206865616465722E0A
		Private mZipDirectory As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4469736B206261636B656E643A207468652074656D7020666F6C64657220776865726520746865206172636869766520776173206578747261637465642E2052656D6F766564207265637572736976656C7920696E2044657374727563746F722E0A
		Private mExtractRoot As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4469736B206261636B656E643A2074686520736F757263652061726368697665206F6E206469736B202D2065697468657220757365722D737570706C696564206F7220612074656D70202E7A697020776520637265617465642E0A
		Private mSourceFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h21, Description = 4469736B206261636B656E643A2054727565207768656E2077652063726561746564206D536F7572636546696C65202866726F6D2061204D656D6F7279426C6F636B2920616E64206D7573742072656D6F766520697420696E2044657374727563746F722E0A
		Private mOwnsSourceFile As Boolean = False
	#tag EndProperty


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
