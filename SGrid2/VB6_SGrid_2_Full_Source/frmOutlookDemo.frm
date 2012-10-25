VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form frmOutlookDemo 
   Caption         =   "SGrid Mailbox Style Demonstration"
   ClientHeight    =   5655
   ClientLeft      =   4650
   ClientTop       =   5625
   ClientWidth     =   7335
   Icon            =   "frmOutlookDemo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5655
   ScaleWidth      =   7335
   Begin vbAcceleratorSGrid6.vbalGrid grdOutlook 
      Height          =   5175
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9128
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
   End
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   6480
      Top             =   420
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   21812
      Images          =   "frmOutlookDemo.frx":014A
      Version         =   131072
      KeyCount        =   19
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.Label lblInfo 
      Caption         =   "Mail box demonstration. Right click on the header for sorting and grouping options."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   7215
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Index           =   0
      End
   End
   Begin VB.Menu mnuViewTOP 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Columns"
         Index           =   0
         Begin VB.Menu mnuColumns 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Auto-Preview"
         Index           =   1
         Begin VB.Menu mnuPreview 
            Caption         =   "&None"
            Index           =   0
         End
         Begin VB.Menu mnuPreview 
            Caption         =   "&Unread Messages"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuPreview 
            Caption         =   "&All Messages"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Group Box"
         Index           =   3
      End
   End
   Begin VB.Menu mnuContextTOP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuContext 
         Caption         =   "Sort &Ascending"
         Index           =   0
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Sort &Descending"
         Index           =   1
      End
      Begin VB.Menu mnuContext 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Group by this &Field"
         Index           =   3
      End
      Begin VB.Menu mnuContext 
         Caption         =   "&Group Box"
         Index           =   4
      End
      Begin VB.Menu mnuContext 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContext 
         Caption         =   "&Remove this Column"
         Index           =   6
      End
   End
   Begin VB.Menu mnuMailContextTOP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuMailContext 
         Caption         =   "&Open..."
         Index           =   0
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "&Print..."
         Index           =   1
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "&Reply"
         Index           =   3
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "Reply to &All"
         Index           =   4
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "&Follow Up..."
         Index           =   6
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "Mark as Rea&d"
         Index           =   7
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "Mark as &Unread"
         Index           =   8
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "&Categories..."
         Index           =   9
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "&Delete"
         Index           =   11
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "&Move To Folder..."
         Index           =   12
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuMailContext 
         Caption         =   "&Options..."
         Index           =   14
      End
   End
End
Attribute VB_Name = "frmOutlookDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function decodeHex(ByVal sText As String) As String
Dim i As Long
Dim b() As Byte
Dim lCurrent As Long
Dim lByte As Long
Dim sHex As String
Dim iPos As Long
   
   ReDim b(0 To Len(sText) * 2) As Byte
   For i = 1 To Len(sText)
      sHex = Mid(sText, i, 1)
      If IsNumeric(sHex) Then
         lByte = CInt(sHex)
      Else
         lByte = AscW(sHex) - 55
         If (lByte < 10) Or (lByte > 15) Then
            MsgBox "Error in Hex.", vbExclamation
            Exit Function
         End If
      End If
      If (i Mod 2) = 1 Then
         lCurrent = (lByte * &H10&)
      Else
         b(iPos) = lCurrent Or lByte
         iPos = iPos + 1
      End If
   Next i
   decodeHex = b
   
End Function

Private Sub Form_Load()
Dim iRow As Long
Dim iIconUrgent As Long
Dim iIconAttach As Long
Dim iIconFlag As Long
Dim iIconType As Long
Dim iIdx As Long
Dim dDate As Date
Dim lColour As Long
Dim iCol As Long
Dim lHeight As Long
Dim cS As cGridCell
Dim cSUnread As cGridCell
Dim iMenu As Long

   With grdOutlook
      ' Turn redraw off for speed:
      .Redraw = False
   
      ' Set up the grid:
      
      ' Source of icons.  This can be vbAccelerator ImageList control, class or
      ' a VB ImageList
      .ImageList = ilsIcons
      ' Row mode - select the entire row:
      .RowMode = True
      ' Allow more than one row to be selected:
      .MultiSelect = True
      ' Set the default row height:
      .DefaultRowHeight = 18
      ' Outlook style for the header control:
      .HeaderFlat = True
      ' As it says
      .StretchLastColumnToFit = True
      
      ' Add the columns:
      .AddColumn "urgency", , , 9, 28, , , , False, , , CCLSortIcon
      .AddColumn "type", , , 10, 28, , , , False, , , CCLSortIcon
      .AddColumn "attach", , , 12, 28, , , , False, , , CCLSortIcon
      .AddColumn "flag", , , 11, 28, , , , False, , , CCLSortIcon
      .AddColumn "from", "From", , , 96
      .AddColumn "subject", "Subject", , , 256
      .AddColumn "received", "Received", , , 96, , , , , "dd/mm/yy hh:mm", , CCLSortDate
      .AddColumn "to", "To", , , 96
      .AddColumn "size", "Size", , , 56, , , , , "#,##0", , CCLSortNumeric
      ' Add two invisible columns to cache status information:
      .AddColumn "read", , , , , False
      .AddColumn "ID", , , , , False
      ' The special "rowcolumntext" column must be added to the end
      ' of the available columns.  This never appears as a column
      ' header, but the text in it is drawn underneath the row (assuming
      ' the row is high enough for it, starting at the column
      ' specified by .RowTextStartColumn:
      .AddColumn "body", , , , 96 + 256 + 96 + 96, , , , , , True
      
      ' When the user types a key, this determines which column
      ' the control will search in
      .KeySearchColumn = .ColumnIndex("subject")
      
      ' You can specify specifically at which column the text will start
      ' like this:
      '   .RowTextStartColumn = .ColumnIndex("from")
      ' If you do this you need to track the ColumnOrderChanged event to
      ' ensure you are at the right column if the user moves this column
      ' to the end of the grid.  If you don't specify this setting, the
      ' grid will automatically start drawing rowtext at the position
      ' of the first column included in the select (bIncludeInSelect
      ' parameter of AddColumn)
         
      
      ' Once we have added the columns, we can set the headers up
      ' (if we are using headers)
      .SetHeaders
      
      ' Add some demonstration rows:
      
      ' Set up a bold font:
      Dim sFntUnread As New StdFont
      sFntUnread.Name = "Tahoma"
      sFntUnread.Size = 8
      sFntUnread.Bold = True
      
      Set cS = .NewCellFormatObject
      Set cSUnread = .NewCellFormatObject
      Set cSUnread.Font = sFntUnread
      
      ' Create some pretend text for From, Subject and Body
      Dim sFrom(1 To 10) As String
      sFrom(1) = "Carl Ridenhour"
      sFrom(2) = "Kevin Shields"
      sFrom(3) = "Richard D James"
      sFrom(4) = "Luke Slater"
      sFrom(5) = "Mark Bell"
      sFrom(6) = "Frank Black"
      sFrom(7) = "Richard Clayderman"
      sFrom(8) = "James Last"
      sFrom(9) = "Thurston Moore"
      sFrom(10) = "Beth Gibbons"
      
      Dim sSubject(1 To 10) As String
      sSubject(1) = "Check out this demo"
      sSubject(2) = "RE: Sonic Bubblebath Remix"
      sSubject(3) = "FW: The secret world of plants"
      sSubject(4) = """Make like Ghandi"""
      sSubject(5) = "RE: FW: Feast your eyes on those 'Spirit of 1997' animated GIFs!"
      sSubject(6) = "viz New York Trip"
      sSubject(7) = "Belated Happy Birthday"
      sSubject(8) = "RE: What's the score?"
      sSubject(9) = "vbAccelerator: Excellent site!"
      sSubject(10) = "Pass the peas..."
      
      Dim sBody(1 To 11) As String
      sBody(1) = "Impress passing airline passengers by painting a large " & _
         "blue rectangle in your back garden.  They will think that you " & _
         "have a swimming pool."
      sBody(2) = "Bus drivers: pretend to be an airline pilot by wedging " & _
         "the accelerator pedal down with a brick, tying the steering wheel " & _
         "to your seat with a rope and then walking up and down the aisle " & _
         "asking passengers if they are having a nice trip."
      sBody(3) = "A man walks into a butchers'.  He says ""I bet you £100 that " & _
         "you can't get that meat down from the top shelf"".  " & _
         "The butcher looks up, thinks for a moment, then says ""Sorry mate, " & _
         "can't do it, the steaks are too high""."
      sBody(4) = "A skeleton walks into a bar.  He goes up to the barman and " & _
         "asks for a pint of beer and a mop."
      sBody(5) = "Q: What's the best way to catch a rabbit? A: Hide somewhere " & _
         "and make a noise like a carrot."
      sBody(6) = "Forget the others, this is the real deal - increase the size " & _
         "of your elbows by up to 2 inches, possibly guaranteed in just 'weeks'!  " & _
         "No painful work-outs, no hard to take pills, just a simple injection into " & _
         "your left ear once every three days - and you'll soon have the elbows of " & _
         "your dreams!" & vbCrLf & "So don't delay, write back today and add " & _
         "something special to your arms."
      sBody(7) = "Earn money in your spare time!  Easily earn up to $10,000 a week " & _
         "whilst working from home." & vbCrLf & "This offer may sound too good " & _
         "to be true, but read on, otherwise you might be missing out!!  Join " & _
         "1,000's of others who are earning easy money with our lobster and " & _
         "beaver packaging scheme."
      sBody(8) = "A duck walks into a bar.  The barman says ""I'm sorry, we " & _
         "don't serve ducks in here"".  ""That's ok"", replies the duck, " & _
         """I don't really like duck anyway, it tastes a bit like chicken.  " & _
         "And if we're on the subject, I don't really like oranges either. But a " & _
         "nice steak... that would go down like a dream""."
      sBody(9) = "A man and his giraffe walk into a bar.  He orders two beers, " & _
         "and they both drink up (although the giraffe has some difficulties " & _
         "reaching it's beer).  As they're about to finish the man pulls " & _
         "out a shotgun and shoots the giraffe dead.  It drops to the ground, " & _
         "and suddenly he's walking out the bar.  ""Hey!"", shouts the barman, " & _
         """You can't just leave that lyin' here"".  ""Sorry mate, you must be " & _
         "confused"", says the man.  ""The lion's in the last bar, that was my giraffe..."""
      sBody(10) = "Say goodbye to Y2.038K Fears with the Trouser Press 2038." & _
         vbCrLf & "Top scientists have been working around the clock " & _
         "to find a solution to the most worrying problem post Millenial problem " & _
         "- what happens if your trousers are trapped in their press on " & _
         "Monday, January 18th 2038?" & vbCrLf & "Rest assured that thanks " & _
         "to this miracle of bug-free microchip technology you will be wearing " & _
         "a crisply-creased pair of your favourite trousers to greet the " & _
         "Monday morning.  If you live that long. (Batteries extra)."
      sBody(11) = "A man goes to see an optometrist. The doctor says, " & _
         """You have to stop masturbating"". The guy says, ""Why? Am I going " & _
         "blind?"" The doctor says, ""No, you're upsetting the other patients " & _
         "in the waiting room."""
                           
      ' Now add the rows:
      For iRow = 1 To 200
         
         ' set the urgency:
         iIconUrgent = Rnd * 3
         Select Case iIconUrgent
         Case 1
            iIconUrgent = 7
         Case 2
            iIconUrgent = 8
         Case Else
            iIconUrgent = -1
         End Select
         .CellDetails iRow, 1, , , iIconUrgent
         
         ' set the type:
         If (iRow < 16) Then
            iIconType = 1
         Else
            iIconType = Rnd * 2 + 2
         End If
         .CellIcon(iRow, 2) = iIconType
         
         ' set the attachment:
         If Rnd * 20 > 17 Then
            iIconAttach = 14
         Else
            iIconAttach = -1
         End If
         .CellIcon(iRow, 3) = iIconAttach
         
         ' set the Flag:
         If Rnd * 20 > 18 Then
            iIconFlag = 13
         Else
            iIconFlag = -1
         End If
         .CellIcon(iRow, 4) = iIconFlag
         
         ' mark as irrelevant ("junk mail"):
         iIdx = CInt(Rnd * 9) + 1
         If iIdx = 7 Or iIdx = 8 Then
            lColour = vbGrayText
         Else
            lColour = -1
         End If
         
         ' from:
         If (iRow < 16) Then
            .CellDetails iRow, 5, sFrom(iIdx), , , , lColour, sFntUnread
         Else
            .CellDetails iRow, 5, sFrom(iIdx), , , , lColour
         End If
         
         ' subject:
         iIdx = CInt(Rnd * 9) + 1
         If (iRow < 16) Then
            .CellDetails iRow, 6, sSubject(iIdx), , , , lColour, sFntUnread
         Else
            .CellDetails iRow, 6, sSubject(iIdx), , , , lColour
         End If
         
         ' date:
         dDate = Now
         If (iRow < 16) Then
            dDate = DateAdd("d", -Rnd * 3, dDate)
            dDate = dDate + TimeSerial(Rnd * 24, Rnd * 60, Rnd * 60)
            .CellDetails iRow, 7, dDate, , , , lColour, sFntUnread
         Else
            dDate = DateAdd("m", -Rnd * 12, dDate)
            dDate = DateAdd("d", -Rnd * 31, dDate)
            dDate = dDate + TimeSerial(Rnd * 24, Rnd * 60, Rnd * 60)
            .CellDetails iRow, 7, dDate, , , , lColour
         End If
         
         ' to:
         If (iRow < 16) Then
            .CellDetails iRow, 8, "Steve McMahon", , , , lColour, sFntUnread
         Else
            .CellDetails iRow, 8, "Steve McMahon", , , , lColour
         End If
         
         ' size:
         If (iIconAttach = -1) Then
            .CellDetails iRow, 9, Rnd * 4096, DT_END_ELLIPSIS Or DT_RIGHT Or DT_SINGLELINE
         Else
            .CellDetails iRow, 9, Rnd * 1024 * 1024 + 4096, DT_END_ELLIPSIS Or DT_RIGHT Or DT_SINGLELINE
         End If
         
         iIdx = CInt(Rnd * 9) + 1
         .CellDetails iRow, 12, sBody(iIdx), DT_WORDBREAK, , , RGB(0, 0, &HBF)
         lHeight = .EvaluateTextHeight(iRow, 12) + .DefaultRowHeight + 4

         ' Read/unread marker:
         If (iRow < 16) Then
            .CellDetails iRow, 10, "NOTREAD"
            .RowHeight(iRow) = lHeight
         Else
            .CellDetails iRow, 10, "READ"
         End If
         
         ' ID marker:
         .CellDetails iRow, 11, iRow
                  
         
      Next iRow
      
      
      ' Add the columns to the menu:
      For iCol = 1 To .Columns
         If (.ColumnVisible(iCol)) And (iCol <> 12) Then
            If (iMenu > 0) Then
               Load mnuColumns(iMenu)
               mnuColumns(iMenu).Visible = True
            End If
            If (.ColumnHeader(iCol) = "") Then
               mnuColumns(iMenu).Caption = StrConv(.ColumnKey(iCol), vbProperCase)
            Else
               mnuColumns(iMenu).Caption = .ColumnHeader(iCol)
            End If
            mnuColumns(iMenu).Tag = .ColumnKey(iCol)
            mnuColumns(iMenu).Checked = True
            iMenu = iMenu + 1
         End If
      Next iCol
      
      .Redraw = True
   End With
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
   lblInfo.Width = Me.ScaleWidth - lblInfo.left * 2
   grdOutlook.Move 2 * Screen.TwipsPerPixelX, grdOutlook.top, Me.ScaleWidth - 4 * Screen.TwipsPerPixelX, Me.ScaleHeight - grdOutlook.top - 4 * Screen.TwipsPerPixelY
End Sub

Private Sub grdOutlook_ColumnClick(ByVal lCol As Long)
Dim iCol As Long
Dim iSortCol As Long
Dim sJunk() As String, eJunk() As ECGSortOrderConstants

   With grdOutlook.SortObject
      .ClearNongrouped
      iSortCol = .IndexOf(lCol)
      If (iSortCol <= 0) Then
         iSortCol = .Count + 1
      End If
      
      .SortColumn(iSortCol) = lCol
      If (grdOutlook.ColumnSortOrder(lCol) = CCLOrderNone) Or (grdOutlook.ColumnSortOrder(lCol) = CCLOrderDescending) Then
         .SortOrder(iSortCol) = CCLOrderAscending
      Else
         .SortOrder(iSortCol) = CCLOrderDescending
      End If
      grdOutlook.ColumnSortOrder(lCol) = .SortOrder(iSortCol)
      .SortType(iSortCol) = grdOutlook.ColumnSortType(lCol)
      
      ' Place ascending/descending icon:
      For iCol = 1 To grdOutlook.Columns
         If (iCol <> lCol) Then
            If Not (grdOutlook.ColumnIsGrouped(iCol)) Then
               If grdOutlook.ColumnImage(iCol) > 16 Then
                  grdOutlook.ColumnImage(iCol) = 0
               End If
            End If
         ElseIf grdOutlook.ColumnHeader(iCol) <> "" Then
            grdOutlook.ColumnImageOnRight(iCol) = True
            If (.SortOrder(iSortCol) = CCLOrderAscending) Then
               grdOutlook.ColumnImage(iCol) = 17
            Else
               grdOutlook.ColumnImage(iCol) = 18
            End If
         End If
      Next iCol
      
   End With
   
   Screen.MousePointer = vbHourglass
   grdOutlook.Sort
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub grdOutlook_ColumnOrderChanged()
   '
End Sub

Private Sub grdOutlook_ColumnWidthChanging(ByVal lCol As Long, lWidth As Long, bCancel As Boolean)
   If (lWidth < 26) Then
      lWidth = 26
   End If
End Sub

Private Sub grdOutlook_DblClick(ByVal lRow As Long, ByVal lCol As Long)
   '
   '
End Sub

Private Sub grdOutlook_HeaderRightClick(ByVal x As Single, ByVal y As Single)
Dim lCol As Long
   
   lCol = grdOutlook.ColumnHeaderFromPoint(x, y)
   
   If (lCol > 0) Then
      mnuContext(0).Enabled = True
      mnuContext(1).Enabled = True
      mnuContext(3).Enabled = True
      'Debug.Print grdOutlook.ColumnHeader(lCol), grdOutlook.ColumnIsGrouped(lCol)
      mnuContext(3).Caption = IIf(grdOutlook.ColumnIsGrouped(lCol), "Don't Group By This Field", "Group By This Field")
      mnuContext(6).Enabled = True
            
      mnuContext(0).Checked = (grdOutlook.ColumnSortOrder(lCol) = CCLOrderAscending)
      mnuContext(1).Checked = (grdOutlook.ColumnSortOrder(lCol) = CCLOrderDescending)
      
   Else
      mnuContext(0).Enabled = False
      mnuContext(1).Enabled = False
      mnuContext(3).Enabled = False
      mnuContext(6).Enabled = False
   
      mnuContext(0).Checked = False
      mnuContext(1).Checked = False
   
   End If
      
   x = (x + grdOutlook.ScrollOffsetX) * Screen.TwipsPerPixelX + grdOutlook.left
   y = y * Screen.TwipsPerPixelY + grdOutlook.top
   mnuContextTOP.Tag = lCol
   Me.PopupMenu mnuContextTOP, , x, y
   
   
End Sub

Private Sub grdOutlook_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If (Button = vbRightButton) Then
      Dim lRow As Long
      Dim lCol As Long
      grdOutlook.CellFromPoint x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY, lRow, lCol
      If (lRow > 0) And (lCol > 0) Then
         
         ' Note here I'm not showing a menu for groups.
         ' In Outlook, the behaviour is to perform the action on all
         ' subitems of the group, unless the user has selected individual
         ' items within the group.  This is do-able.
         If Not (grdOutlook.RowIsGroup(lRow)) Then
         
            Dim iSelCount As Long
            iSelCount = grdOutlook.SelectionCount
            If (iSelCount > 0) Then
               ' Show appropriate options depending on the number
               ' of selected mails:
               mnuMailContext(3).Visible = (iSelCount = 1)
               mnuMailContext(4).Visible = (iSelCount = 1)
               mnuMailContext(5).Visible = (iSelCount = 1)
               mnuMailContext(6).Visible = (iSelCount = 1)
               mnuMailContext(13).Visible = (iSelCount = 1)
               mnuMailContext(14).Visible = (iSelCount = 1)
               mnuMailContextTOP.Tag = iSelCount
               x = x + grdOutlook.left
               y = y + grdOutlook.top
               
               ' Show the menu
               Me.PopupMenu mnuMailContextTOP, , x, y
               
            End If
            
         End If
      End If
   End If
End Sub

Private Sub grdOutlook_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Static sSearch As String
   'Debug.Print "RequestEdit"
   If (iKeyAscii <> 0) Then
      'Debug.Print iKeyAscii
      ' Search for the match:
      If (iKeyAscii <> 8) Then
         sSearch = sSearch & Chr$(iKeyAscii)
      Else
         If (Len(sSearch) > 0) Then
            sSearch = left$(sSearch, Len(sSearch) - 1)
         End If
      End If
      'Debug.Print sSearch
   End If
   bCancel = True
End Sub

Private Sub mnuColumns_Click(Index As Integer)
Dim bS As Long
Dim lCol As Long
   bS = Not (mnuColumns(Index).Checked)
   mnuColumns(Index).Checked = bS
   grdOutlook.ColumnVisible(mnuColumns(Index).Tag) = bS
End Sub

Private Sub mnuContext_Click(Index As Integer)
Dim lCol As Long
   Select Case Index
   Case 0
      lCol = CLng(mnuContextTOP.Tag)
      If (mnuContext(0).Checked) Then
         mnuContext(0).Checked = False
         grdOutlook.ColumnSortOrder(lCol) = CCLOrderNone
      Else
         mnuContext(0).Checked = True
         grdOutlook.ColumnSortOrder(lCol) = CCLOrderNone
         grdOutlook_ColumnClick lCol
      End If
   
   Case 1
      lCol = CLng(mnuContextTOP.Tag)
      If (mnuContext(1).Checked) Then
         mnuContext(1).Checked = False
         grdOutlook.ColumnSortOrder(lCol) = CCLOrderNone
      Else
         mnuContext(1).Checked = True
         grdOutlook.ColumnSortOrder(lCol) = CCLOrderAscending
         grdOutlook_ColumnClick lCol
      End If
   
   Case 3
      lCol = CLng(mnuContextTOP.Tag)
      If (grdOutlook.ColumnIsGrouped(lCol)) Then
         ' Ungroup
         grdOutlook.ColumnIsGrouped(lCol) = False
      Else
         ' Group
         grdOutlook.ColumnIsGrouped(lCol) = True
      End If
   Case 4
      ' Group box:
      grdOutlook.AllowGrouping = Not (grdOutlook.AllowGrouping)
      mnuContext(Index).Checked = grdOutlook.AllowGrouping
      mnuView(3).Checked = grdOutlook.AllowGrouping
   Case 6
      lCol = CLng(mnuContextTOP.Tag)
      grdOutlook.ColumnVisible(lCol) = False
      mnuColumns(mnuContextTOP.Tag - 1).Checked = False
   End Select
   
End Sub

Private Sub mnuContextTOP_Click()
   mnuContext(4).Checked = grdOutlook.AllowGrouping
End Sub

Private Sub mnuFile_Click(Index As Integer)
   Unload Me
End Sub

Private Sub mnuMailContext_Click(Index As Integer)
   Select Case Index
   Case 0
      MsgBox "Open selected for " & vbCrLf & getSelectedMailTitles, vbInformation
   Case 1
      MsgBox "Print selected for " & vbCrLf & getSelectedMailTitles, vbInformation
   Case 3
      If (grdOutlook.SelectionCount = 1) Then
         ' Some people can read this without worrying about decode hex.
         Dim sTitle As String
         sTitle = "50006F00730074006300610072006400200063006F006D00700065" & _
                        "0074006900740069006F006E002100"
         Dim sDescription As String
         sDescription = "44006F00200079006F00750020006C0069006B0065002000740068" & _
                        "0069007300200063006F00640065003F0020002000490027006400" & _
                        "20006C0069006B006500200074006F002000680065006100720020" & _
                        "00660072006F006D00200079006F0075002E002000200053006500" & _
                        "6E00640020006D00650020006100200070006F0073007400630061" & _
                        "00720064003A000D000A0020002000530074006500760065002000" & _
                        "4D0063004D00610068006F006E000D000A00200020003200200043" & _
                        "00610072007900730066006F0072007400200052006F0061006400" & _
                        "0D000A0020002000430072006F00750063006800200045006E0064" & _
                        "000D000A00200020004C006F006E0064006F006E000D000A002000" & _
                        "20004E00380020003800520042000D000A002000200055006E0069" & _
                        "0074006500640020004B0069006E00670064006F006D000D000A00" & _
                        "0D000A004200650073007400200070006F00730074006300610072" & _
                        "00640073002000770069006E0020007000720069007A0065007300" & _
                        "21000D000A00"
         MsgBox decodeHex(sDescription), vbInformation, decodeHex(sTitle)
      Else
         MsgBox "Reply selected for " & vbCrLf & getSelectedMailTitles, vbInformation
      End If
   Case 4
      MsgBox "Reply to All selected for " & vbCrLf & getSelectedMailTitles, vbInformation
   Case 6
      ' Follow up
      followUp
   Case 7
      ' Mark as read
      markAsRead True
   Case 8
      ' Mark as unread
      markAsRead False
   Case 9
      MsgBox "Categories selected for " & vbCrLf & getSelectedMailTitles, vbInformation
   Case 11
      ' Delete
      deleteMail
   Case 12
      MsgBox "Move to Folder selected for " & vbCrLf & getSelectedMailTitles, vbInformation
   Case 14
      MsgBox "Show Options dialog here", vbInformation
   End Select
End Sub

Private Function getSelectedMailTitles() As String
Dim i As Long
Dim lRow As Long
Dim sRet As String
   For i = 1 To grdOutlook.SelectionCount
      lRow = grdOutlook.SelectedRowByIndex(i)
      If (Len(sRet) > 0) Then
         sRet = sRet & ", "
      End If
      sRet = sRet & grdOutlook.CellText(lRow, 6)
   Next i
   getSelectedMailTitles = sRet
End Function

Private Sub markAsRead(ByVal bState As Boolean)
Dim i As Long
Dim lRow As Long
Dim iCol As Long
Dim lIcon As Long
Dim sFnt As New StdFont
Dim lHeight As Long

   grdOutlook.Redraw = False

   sFnt.Name = "Tahoma"
   sFnt.Size = 8
   
   ' In real life, you'd actually check the reply state
   ' before setting the icon like this (currently, the
   ' reply "state" is cleared when you mark as read
   ' or unread.
   If (bState) Then
      lIcon = 4
   Else
      sFnt.Bold = True
      lIcon = 1
   End If

   For i = 1 To grdOutlook.SelectionCount
      lRow = grdOutlook.SelectedRowByIndex(i)
      If (grdOutlook.RowIsGroup(lRow)) Then
         '
      Else
         grdOutlook.CellIcon(lRow, 2) = lIcon
         For iCol = 1 To grdOutlook.Columns - 1 ' miss out the preview text
            grdOutlook.CellFont(lRow, iCol) = sFnt
         Next iCol
         grdOutlook.CellText(lRow, 10) = IIf(bState, "READ", "NOTREAD")
         If Not (bState) Then
            lHeight = grdOutlook.EvaluateTextHeight(lRow, 12) + grdOutlook.DefaultRowHeight + 4
            grdOutlook.RowHeight(lRow) = lHeight
         Else
            grdOutlook.RowHeight(lRow) = grdOutlook.DefaultRowHeight
         End If
      End If
   Next i
   
   grdOutlook.Redraw = True
   
End Sub
Private Sub followUp()
Dim i As Long
Dim lRow As Long
   grdOutlook.Redraw = False
   For i = 1 To grdOutlook.SelectionCount
      lRow = grdOutlook.SelectedRowByIndex(i)
      If (grdOutlook.RowIsGroup(lRow)) Then
      Else
         grdOutlook.CellIcon(lRow, 4) = 13
      End If
   Next i
   grdOutlook.Redraw = True
End Sub

Private Sub deleteMail()
Dim i As Long
Dim lRow As Long
   grdOutlook.Redraw = False
   
   For i = grdOutlook.SelectionCount To 1 Step -1
      lRow = grdOutlook.SelectedRowByIndex(i)
      If (grdOutlook.RowIsGroup(lRow)) Then
         If (grdOutlook.RowGroupingState(lRow) = ecgCollapsed) Then
            ' All of the subitems will have been selected already, so just
            ' delete the group
            grdOutlook.RemoveRow lRow
         End If
      Else
         ' Delete this row
         grdOutlook.RemoveRow lRow
      End If
   Next i
   
   
   grdOutlook.Redraw = True
End Sub


Private Sub mnuViewTOP_Click()
   mnuView(3).Checked = grdOutlook.AllowGrouping
End Sub

Private Sub mnuPreview_Click(Index As Integer)
Dim i As Long
Dim lHeight As Long

   For i = 0 To 2
      mnuPreview(i).Checked = (i = Index)
   Next i
   
   grdOutlook.Redraw = False
   If (Index = 0) Then
      ' No preview:
      For i = 1 To grdOutlook.Rows
         If Not grdOutlook.RowIsGroup(i) Then
            grdOutlook.RowHeight(i) = grdOutlook.DefaultRowHeight
         End If
      Next i
   ElseIf (Index = 1) Then
      ' Preview unread only:
      For i = 1 To grdOutlook.Rows
         If Not grdOutlook.RowIsGroup(i) Then
            If (grdOutlook.CellText(i, 10) = "NOTREAD") Then
               lHeight = grdOutlook.EvaluateTextHeight(i, 12) + grdOutlook.DefaultRowHeight
               grdOutlook.RowHeight(i) = lHeight
            Else
               grdOutlook.RowHeight(i) = grdOutlook.DefaultRowHeight
            End If
         End If
      Next i
   Else
      ' All preview:
      For i = 1 To grdOutlook.Rows
         If Not grdOutlook.RowIsGroup(i) Then
            lHeight = grdOutlook.EvaluateTextHeight(i, 12) + grdOutlook.DefaultRowHeight
            grdOutlook.RowHeight(i) = lHeight
         End If
      Next i
   End If
   grdOutlook.Redraw = True
End Sub

Private Sub mnuView_Click(Index As Integer)
   Select Case Index
   Case 3
      mnuContext_Click 4
   End Select
End Sub


