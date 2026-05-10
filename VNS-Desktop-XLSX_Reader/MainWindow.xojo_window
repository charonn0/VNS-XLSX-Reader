#tag DesktopWindow
Begin DesktopWindow MainWindow
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   True
   HasMaximizeButton=   True
   HasMinimizeButton=   True
   HasTitleBar     =   True
   Height          =   600
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   482459647
   MenuBarVisible  =   False
   MinimumHeight   =   200
   MinimumWidth    =   400
   Resizeable      =   True
   Title           =   "#strings.kStrAppTitle"
   Type            =   0
   Visible         =   True
   Width           =   900
   Begin DesktopTabPanel TabPanelSheets
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   600
      Index           =   -2147483648
      Italic          =   False
      Left            =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Panels          =   "(no workbook)"
      Scope           =   0
      SmallTabs       =   False
      TabDefinition   =   "Tab 0\rTab 1"
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   0
      Transparent     =   False
      Underline       =   False
      Value           =   0
      Visible         =   True
      Width           =   900
   End
   Begin DesktopListBox ListboxData
      AllowAutoDeactivate=   True
      AllowAutoHideScrollbars=   True
      AllowExpandableRows=   False
      AllowFocusRing  =   True
      AllowResizableColumns=   True
      AllowRowDragging=   False
      AllowRowReordering=   False
      Bold            =   False
      ColumnCount     =   1
      ColumnWidths    =   ""
      DefaultRowHeight=   -1
      DropIndicatorVisible=   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      GridLineStyle   =   0
      HasBorder       =   True
      HasHeader       =   True
      HasHorizontalScrollbar=   True
      HasVerticalScrollbar=   True
      HeadingIndex    =   -1
      Height          =   556
      Index           =   -2147483648
      InitialValue    =   ""
      Italic          =   False
      Left            =   8
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      RequiresSelection=   False
      RowSelectionType=   0
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   32
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   884
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag MenuHandler
		Function FileOpen() As Boolean Handles FileOpen.Action
		  Var dlg As New OpenFileDialog
		  Var t As New FileType
		  t.Name = "Excel Workbook"
		  t.Extensions = "xlsx"
		  dlg.Filter = t
		  Var f As FolderItem = dlg.ShowModal(Self)
		  If f = Nil Then Return True
		  LoadWorkbook(f)
		  Return True
		End Function
	#tag EndMenuHandler


	#tag Method, Flags = &h21
		Private Sub FillCurrentSheet()
		  If mWorkbook Is Nil Then Return
		  Var idx As Integer = TabPanelSheets.SelectedPanelIndex
		  If idx < 0 Or idx >= mWorkbook.SheetCount Then Return
		  Var sheet As XLSXSheet = mWorkbook.SheetAt(idx + 1)
		  XLSXDesktopListboxFiller.Fill(ListboxData, sheet, mWorkbook.Styles)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadWorkbook(file As FolderItem)
		  Try
		    Var wb As XLSXWorkbook = XLSXReader.Open(file)
		    mWorkbook = wb
		    Self.Title = strings.kStrAppTitle + " — " + wb.SourceName + " [" + Str(wb.SheetCount) + "]"
		    RebuildTabs(wb)
		  Catch ex As XLSXException
		    ShowErrorFor(ex)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub RebuildTabs(wb As XLSXWorkbook)
		  ' Clear existing panels.
		  While TabPanelSheets.PanelCount > 0
		    TabPanelSheets.RemovePanelAt(TabPanelSheets.PanelCount - 1)
		  Wend
		  
		  ' Add one panel per sheet.
		  For Each sn As String In wb.SheetNames
		    TabPanelSheets.AddPanel(sn)
		  Next
		  
		  ' Show the first sheet (PanelIndex on child controls is 1-based; 0 = all panels).
		  If wb.SheetCount > 0 Then
		    TabPanelSheets.SelectedPanelIndex = 0
		    ListboxData.PanelIndex = 1
		    FillCurrentSheet
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ShowErrorFor(ex As XLSXException)
		  Var msg As String
		  Select Case ex.Code
		  Case XLSXEnums.eParseError.NotAZip
		    msg = strings.kStrErrorNotXLSX
		  Case XLSXEnums.eParseError.MalformedXML
		    msg = strings.kStrErrorMalformed
		  Case XLSXEnums.eParseError.Encrypted
		    msg = strings.kStrErrorEncrypted
		  Case XLSXEnums.eParseError.MissingPart
		    msg = strings.kStrErrorNotXLSX
		  Else
		    msg = strings.kStrErrorGeneric
		  End Select
		  If ex.Detail <> "" Then msg = msg + " (" + ex.Detail + ")"
		  Var d As New MessageDialog
		  d.Title = strings.kStrErrorTitle
		  d.Message = strings.kStrErrorTitle
		  d.Explanation = msg
		  Call d.ShowModal
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mWorkbook As XLSXWorkbook
	#tag EndProperty


#tag EndWindowCode

#tag Events TabPanelSheets
	#tag Event
		Sub PanelChanged()
		  ListboxData.PanelIndex = TabPanelSheets.SelectedPanelIndex + 1
		  FillCurrentSheet
		End Sub
	#tag EndEvent
#tag EndEvents
