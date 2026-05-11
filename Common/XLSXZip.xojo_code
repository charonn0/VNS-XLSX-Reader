#tag Class
Protected Class XLSXZip
	#tag Method, Flags = &h21
		Private Sub Constructor(ZipData As MemoryBlock)
		  ' This contructor takes raw Zip/XLSX data as a MemoryBlock and
		  ' parses the Zip structure into the mZipDirectory dictionary. This
		  ' dictionary uses full file paths as the key and the values are
		  ' dictionaries of file metadata, including the offset, length,
		  ' compression method, and checksum.
		  ' If the ZipData is empty, corrupt, or not a Zip archive then an
		  ' exception is raised.
		  
		  Const ZIP_HEADER_SIG = &h04034b50
		  Const ZIP_FOOTER_SIG = &h08074b50
		  Const ZIP_DIR_SIG = &h02014b50
		  Const FLAG_NAME_UTF = 2048
		  Const FLAG_TRAILER = 8
		  
		  mZipData = ZipData
		  mZipDirectory = New Dictionary
		  Var stream As New BinaryStream(ZipData)
		  stream.LittleEndian = True
		  
		  Do Until stream.EOF
		    If Not SeekSignature(stream, ZIP_HEADER_SIG) Then Exit Do ' no more entries
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
		    
		    If BitAnd(flags, FLAG_NAME_UTF) = FLAG_NAME_UTF Then ' UTF8 names
		      path = DefineEncoding(path, Encodings.UTF8).Trim
		    Else ' CP437 names
		      path = DefineEncoding(path, Encodings.DOSLatinUS).Trim
		    End If
		    
		    entry.Value("offset") = stream.Position
		    
		    If BitAnd(flags, FLAG_TRAILER) = FLAG_TRAILER And compSize = 0 Then
		      ' a zip entry might have a trailer containing the actual sizes and checksum
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
		  stream.Close()
		  If mZipDirectory.KeyCount = 0 Then
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "empty data")
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E73205472756520696620746865206172636869766520636F6E7461696E7320612066696C652061742074686520676976656E2072656C617469766520706174682028636173652D73656E7369746976652C204F50432D7374796C652022786C2F776F726B626F6F6B2E786D6C22292E0A
		Function HasPart(partPath As String) As Boolean
		  Return PartFile(partPath) <> Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D206120466F6C6465724974656D206F6E206469736B2E20566572696669657320746865205A4950206D616769632062797465732028353020344220303320303429206265666F72652065787472616374696E672E0A
		Shared Function Open(file As FolderItem) As XLSXZip
		  If file Is Nil Or Not file.Exists Then
		    Raise New XLSXException(XLSXEnums.eParseError.MissingPart, "file is Nil or does not exist")
		  End If
		  
		  Var bs As BinaryStream = BinaryStream.Open(file)
		  Var zipdata As MemoryBlock = bs.Read(bs.Length)
		  bs.Close()
		  
		  Return New XLSXZip(zipdata)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 4F70656E20616E20584C53582066726F6D207261772062797465732028652E672E2057656246696C6555706C6F616465722E55706C6F6164436F6D706C6574652064617461292E205370696C6C7320746F20612074656D702066696C65206265636175736520466F6C6465724974656D2E556E7A6970206E65656473206120466F6C6465724974656D206F6E206469736B2E0A
		Shared Function Open(data As MemoryBlock) As XLSXZip
		  Return New XLSXZip(data)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, Description = 5265736F6C7665206120706172742070617468206C696B652022786C2F776F726B626F6F6B2E786D6C2220746F206120466F6C6465724974656D20696E73696465207468652065787472616374656420747265652E2052657475726E73204E696C20696620616E79207365676D656E74206973206D697373696E672E0A
		Private Function PartFile(partPath As String) As MemoryBlock
		  ' Resolve "xl/workbook.xml" → original uncompressed data.
		  ' Returns Nil if the path doesn't exist.
		  ' Raises an exception if the compression algorithm is unsupported.
		  
		  Var clean As String = partPath.ReplaceAll("\", "/")
		  If clean.BeginsWith("/") Then clean = clean.Middle(1)
		  Var entry As Dictionary = mZipDirectory.Lookup(partPath, Nil)
		  If entry = Nil Then Return Nil
		  
		  mZipData.LittleEndian = True
		  Var offset As UInt64 = entry.Value("offset")
		  Var length As UInt32 = entry.Value("compSize")
		  Var method As UInt16 = entry.Value("method")
		  
		  Select Case method
		  Case 0 ' not compressed
		    Return mZipData.StringValue(offset, length)
		    
		    #If XojoVersion > 2024.02 Then            
		  Case 8 ' deflated
		    ' Xojo 2024r3 adds compression and decompression of MemoryBlocks
		    ' using the GZIP format, but the zip format uses raw deflate.
		    ' Fortunately, the GZIP format is just raw deflate with extra 
		    ' fields. So to use Xojo's decompression feature for unzipping
		    ' we just need to add those fields.
		    Var raw As New MemoryBlock(0)
		    Var bs As New BinaryStream(raw)
		    bs.LittleEndian = True
		    bs.WriteUInt16(&h8B1F) ' magic
		    bs.WriteUInt8(8)       ' method
		    bs.WriteUInt8(0)       ' flags
		    bs.WriteUInt32(0)      ' mtime
		    bs.WriteUInt8(0)       ' extra flags
		    bs.WriteUInt8(0)       ' OS
		    bs.Write(mZipData.StringValue(offset, length)) ' compressed data
		    bs.WriteUInt32(entry.Value("crc")) ' checksum
		    bs.WriteUInt32(entry.Value("origSize") Mod (2^32)) ' size
		    bs.Close()
		    Return raw.Decompress()
		    #EndIf  
		    
		  Else
		    Raise New XLSXException(XLSXEnums.eParseError.NotAZip, "Unsupported compression algorithm: " + Str(method))
		  End Select
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, Description = 52657475726E732074686520656E7472792773205554462D3820746578742C206F7220656D70747920737472696E6720696620746865207061727420697320616273656E742E2043616C6C65727320646563696465207768657468657220616273656E636520697320666174616C2E0A
		Function ReadPart(partPath As String) As String
		  Var fileData As MemoryBlock = PartFile(partPath)
		  If fileData Is Nil Then Return ""
		  Return DefineEncoding(fileData, Encodings.UTF8)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function SeekSignature(Stream As BinaryStream, Signature As UInt32) As Boolean
		  ' Locates the Signature in the Stream. If the Signature is found then this method returns True and the Stream.Position
		  ' property reflects the point in the Stream where it was found. If the Signature was not found this method returns False
		  ' and the Stream.Position property is unchanged.
		  
		  Var pos As UInt64 = Stream.Position
		  Var ok As Boolean
		  Var sig As New MemoryBlock(4)
		  sig.LittleEndian = True
		  sig.UInt32Value(0) = Signature
		  
		  Do Until Stream.EOF
		    Var data As String = Stream.Read(65535)
		    Var offset As Integer = InStrB(data, sig)
		    If offset > 0 Then
		      Stream.Position = Stream.Position - (data.LenB - offset + 1)
		      ok = True
		      Exit Do
		    ElseIf Stream.Length - Stream.Position >= 4 Then
		      Stream.Position = Stream.Position - 3
		    End If
		  Loop
		  If Not ok Then Stream.Position = pos
		  Return ok
		End Function
	#tag EndMethod


	#tag Note, Name = About
		Reads named parts (e.g. xl/workbook.xml) from an XLSX or any ZIP archive,
		using the framework's MemoryBlock.Decompress() to extract parts directly
		into memory.
		
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


	#tag Property, Flags = &h21, Description = 54686520736F757263652061726368697665206F6E206469736B202D2065697468657220757365722D737570706C696564206F7220612074656D702066696C6520776520637265617465642E0A
		Private mZipData As MemoryBlock
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mZipDirectory As Dictionary
	#tag EndProperty


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
