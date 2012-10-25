VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form frmTaskList 
   Caption         =   "SGrid Task List Demonstration"
   ClientHeight    =   4875
   ClientLeft      =   3495
   ClientTop       =   3540
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTaskList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8655
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Text            =   "Edit Box"
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdTasks 
      Height          =   2895
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5106
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
      BorderStyle     =   2
      DisableIcons    =   -1  'True
   End
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   5760
      Top             =   2340
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   11480
      Images          =   "frmTaskList.frx":14F2
      Version         =   131072
      KeyCount        =   10
      Keys            =   "TASKÿICONHEADERÿPRIORITYHEADERÿFLAGHEADERÿCHECKHEADERÿLOWÿHIGHÿFLAGÿCHECKÿUNCHECK"
   End
   Begin VB.Label lblInfo 
      Caption         =   "Task List demonstration. Right click on the header for sorting and grouping options."
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   7215
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuViewTOP 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Columns"
         Index           =   0
         Begin VB.Menu mnuColumn 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Group Box"
         Index           =   2
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
         Caption         =   "&Group By This Field"
         Index           =   3
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Group &Box"
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
End
Attribute VB_Name = "frmTaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function validDate(ByVal sText As String, ByRef dDate As Variant) As Boolean
Dim sDateBit As String
Dim iPos As Long
   If Len(Trim(sText)) = 0 Then
      dDate = Empty
      validDate = True
   Else
      If IsDate(sText) Then
         dDate = CDate(sText)
         validDate = True
      Else
         ' Check if it is of the form day followed by date:
         ' (This is not right - should be using a date picker
         ' here to select the value)
         iPos = InStr(sText, " ")
         If (iPos > 0) Then
            sDateBit = Mid(sText, iPos + 1)
            If (IsDate(sDateBit)) Then
               dDate = CDate(sDateBit)
               validDate = True
            End If
         End If
      End If
   End If
End Function

Private Sub addNewTask()
   '
   If Len(txtEdit.Text) > 0 Then
      Dim lIconIndex As Long
      Dim lCompleteIconIndex As Long
      Dim lPriorityIconIndex As Long
      Dim lFlagIconIndex As Long
      Dim dStartDate As Date
      Dim dCompleteDate As Date
      Dim dDueDate As Date
      Dim fPercentComplete As Single
      Dim sStatus As String
      Dim lRow As Long
      
      lIconIndex = iconIndexForKey("TASK")
      lPriorityIconIndex = iconIndexForKey("LOW")
      lFlagIconIndex = -1
         
      dStartDate = Now
      dDueDate = DateAdd("d", 5, Now)
      dCompleteDate = DateSerial(3000, 6, 6)
      lCompleteIconIndex = iconIndexForKey("UNCHECK")
      fPercentComplete = 0
      sStatus = "Not Started"
         
      lRow = addTask(lIconIndex, lCompleteIconIndex, lPriorityIconIndex, lFlagIconIndex, _
         dStartDate, dDueDate, dCompleteDate, _
         "", txtEdit.Text, sStatus, _
         fPercentComplete, False)
         
      grdTasks.SelectedRow = lRow
      txtEdit.Text = ""
      grdTasks.StartEdit lRow, grdTasks.ColumnIndex("Subject")
      
   End If
   '
End Sub

Private Sub toggleCompletion( _
      ByVal lRow As Long _
   )
Dim bComplete As Boolean
Dim sFnt As StdFont
Dim lColor As OLE_COLOR
Dim dDueDate As Date
Dim iCol As Long

   bComplete = Not (grdTasks.CellIcon(lRow, 2) = iconIndexForKey("CHECK"))
   
   dDueDate = IIf(IsMissing(grdTasks.CellText(lRow, 6)), #6/6/3000#, grdTasks.CellText(lRow, 6))
   
   lColor = vbWindowText
   If (bComplete) Then
      Dim iFnt As IFont
      Set iFnt = grdTasks.Font
      Set sFnt = New StdFont
      iFnt.Clone sFnt
      lColor = RGB(96, 96, 96)
      sFnt.Strikethrough = True
   Else
      Set sFnt = grdTasks.Font
      If (dDueDate < Now) Then
         lColor = RGB(192, 0, 0)
      End If
   End If
   
   
   grdTasks.Redraw = False
   grdTasks.CellIcon(lRow, 2) = IIf(bComplete, iconIndexForKey("CHECK"), iconIndexForKey("UNCHECK"))
   grdTasks.CellText(lRow, 10) = IIf(bComplete, "Completed", "Not Started")
   grdTasks.CellText(lRow, 11) = IIf(bComplete, 1#, 0#)
   For iCol = 5 To 11
      grdTasks.CellFont(lRow, iCol) = sFnt
      grdTasks.CellForeColor(lRow, iCol) = lColor
   Next iCol
   grdTasks.Redraw = True
   
End Sub

Private Function addTask( _
      ByVal lIconIndex As Long, _
      ByVal lCompleteIconIndex As Long, _
      ByVal lPriorityIconIndex As Long, _
      ByVal lFlagIconIndex As Long, _
      ByVal dStartDate As Date, _
      ByVal dDueDate As Date, _
      ByVal dCompleteDate As Date, _
      ByVal sCategories As String, _
      ByVal sSubject As String, _
      ByVal sStatus As String, _
      ByVal fPercentComplete As Single, _
      ByVal bComplete As Boolean _
   )
Dim lRow As Long
Dim sFnt As StdFont
Dim lColor As OLE_COLOR
Dim iCol As Long
Dim bGrouping As Boolean

   lColor = vbWindowText
   If (bComplete) Then
      Dim iFnt As IFont
      Set iFnt = grdTasks.Font
      Set sFnt = New StdFont
      iFnt.Clone sFnt
      sFnt.Strikethrough = True
      lColor = RGB(96, 96, 96)
   Else
      Set sFnt = grdTasks.Font
      If (dDueDate < Now) Then
         lColor = RGB(192, 0, 0)
      End If
   End If
   
   With grdTasks
      .AddRow
      lRow = .Rows
      
      .CellDetails lRow, 1, lIconIndex:=lIconIndex
      .CellDetails lRow, 2, lIconIndex:=lCompleteIconIndex
      .CellDetails lRow, 3, lIconIndex:=lPriorityIconIndex
      .CellDetails lRow, 4, lIconIndex:=lFlagIconIndex
      .CellDetails lRow, 5, dStartDate, oFont:=sFnt, oForeColor:=lColor
      .CellDetails lRow, 6, dDueDate, oFont:=sFnt, oForeColor:=lColor
      If (dCompleteDate < #6/6/3000#) Then
         .CellDetails lRow, 7, dCompleteDate, oFont:=sFnt, oForeColor:=lColor
      End If
      .CellDetails lRow, 8, sCategories, oFont:=sFnt, oForeColor:=lColor
      .CellDetails lRow, 9, sSubject, oFont:=sFnt, oForeColor:=lColor
      .CellDetails lRow, 10, sStatus, oFont:=sFnt, oForeColor:=lColor
      .CellDetails lRow, 11, fPercentComplete, oFont:=sFnt, oForeColor:=lColor
      
      ' Check if we have any groups in effect:
      If (.AllowGrouping) Then
         For iCol = 1 To .Columns
            If (.ColumnIsGrouped(iCol)) Then
               bGrouping = True
               Exit For
            End If
         Next iCol
         
         If (bGrouping) Then
            lRow = .ShiftLastRowToSortLocation()
         End If
         
      End If
      
   End With
   
   addTask = lRow
   
End Function

Private Sub addDummyTask( _
      ByVal sSubject As String, _
      ByVal bComplete As Boolean, _
      ByVal sCategories As String _
   )
Dim lIconIndex As Long
Dim lCompleteIconIndex As Long
Dim lPriorityIconIndex As Long
Dim lFlagIconIndex As Long
Dim dStartDate As Date
Dim dCompleteDate As Date
Dim dDueDate As Date
Dim fPercentComplete As Single
Dim sStatus As String

   lIconIndex = iconIndexForKey("TASK")
   lPriorityIconIndex = IIf(Rnd > 0.5, iconIndexForKey("HIGH"), iconIndexForKey("LOW"))
   lFlagIconIndex = IIf(Rnd > 0.8, iconIndexForKey("FLAG"), -1)
   
   If (bComplete) Then
      dStartDate = DateAdd("d", -Rnd * 60, Now)
      dDueDate = DateAdd("d", Rnd * 60, dStartDate)
      dCompleteDate = DateAdd("d", Rnd * 60, dStartDate)
      lCompleteIconIndex = iconIndexForKey("CHECK")
      fPercentComplete = 1#
      sStatus = "Complete"
   Else
      dStartDate = DateAdd("d", 10 - Rnd * 10, Now)
      dDueDate = DateAdd("d", Rnd * 30, Now)
      dCompleteDate = DateSerial(3000, 6, 6)
      lCompleteIconIndex = iconIndexForKey("UNCHECK")
      If (DateDiff("d", dStartDate, Now) < 0) Then
         fPercentComplete = Rnd
         sStatus = "In Progress"
      Else
         fPercentComplete = 0
         sStatus = "Not Started"
      End If
   End If
   
   addTask lIconIndex, lCompleteIconIndex, lPriorityIconIndex, lFlagIconIndex, _
      dStartDate, dDueDate, dCompleteDate, _
      sCategories, sSubject, sStatus, _
      fPercentComplete, bComplete

End Sub

Private Function iconIndexForKey(ByVal sKey As String)
Dim lIndex
   iconIndexForKey = ilsIcons.ItemIndex(sKey) - 1
End Function

Private Sub addNewTaskRow()
   With grdTasks
      .AddRow
      .CellDetails 1, 9, "Click to add a new task", oForeColor:=vb3DShadow
   End With
End Sub

Private Sub createDummyTaskListData()
Dim taskDate As Date
   
   grdTasks.Clear
   
   addNewTaskRow
      
   addDummyTask "Wear a hat more frequently whilst coding", True, "Website"
   addDummyTask "Complete VB Command Bar control menu support", False, "Website"
   addDummyTask "Purchase LFO 'Sheath'", True, "Personal"
   addDummyTask "Buy cod fillets", False, "Personal"
   addDummyTask "Julia's birthday", False, "Personal"
   addDummyTask "Move all my junk into the loft", False, "Personal"
   addDummyTask "Complete SGrid update", True, "Website"
   addDummyTask "Confirm Komplett Samsung 191T Order", True, "Business"
   addDummyTask "Find supplier for Sharp Transmeta Efficeon laptop", False, "Business"
   addDummyTask "Send presents to the Mollers", False, "Personal"
   addDummyTask "Try and get that bus back without anyone noticing", False, "Personal"
   addDummyTask "Renew Wired subscription", True, "Personal"
   addDummyTask "SGrid .NET completion", False, "Website"
   addDummyTask "Fix Copyright page on site", False, "Website"
   
End Sub

Private Sub configureGrid()
Dim iCol As Long
   
   With grdTasks
      
      .Redraw = False
      
      ' Configure Image List:
      .ImageList = ilsIcons.hIml
      
      ' Set grid lines
      .GridLines = True
      .GridLineMode = ecgGridFillControl
      .GridLineColor = vb3DShadow
      
      ' Various display and behaviour settings
      .HighlightSelectedIcons = False
      .RowMode = True
      .Editable = True
      .SingleClickEdit = True
      ' Currently there's a problem if you set StretchLastColumnToFit = true
      ' when the grid's redraw style is set to true, as the first column
      ' ends up the wrong width.
      .StretchLastColumnToFit = True
      
      ' Set so the first row can be used
      .SplitRow = 1
      
      ' Add columns:
      .AddColumn "Icon", iIconIndex:=iconIndexForKey("ICONHEADER"), eSortType:=CCLSortIcon, lColumnWidth:=20
      .AddColumn "Complete", iIconIndex:=iconIndexForKey("CHECKHEADER"), eSortType:=CCLSortIcon, lColumnWidth:=20
      .AddColumn "Priority", iIconIndex:=iconIndexForKey("PRIORITYHEADER"), bVisible:=False, eSortType:=CCLSortIcon, lColumnWidth:=20
      .AddColumn "Flag", iIconIndex:=iconIndexForKey("FLAGHEADER"), bVisible:=False, eSortType:=CCLSortIcon, lColumnWidth:=20
      .AddColumn "StartDate", "Start Date", bVisible:=False, sFmtString:="dddd dd/mm/yyyy", eSortType:=CCLSortDateDayAccuracy, lColumnWidth:=128
      .AddColumn "DueDate", "Due Date", sFmtString:="dddd dd/mm/yyyy", eSortType:=CCLSortDateDayAccuracy, lColumnWidth:=128
      .AddColumn "DateCompleted", "Date Completed", bVisible:=False, sFmtString:="dddd dd/mm/yyyy", eSortType:=CCLSortDateDayAccuracy, lColumnWidth:=128
      .AddColumn "Categories", "Categories", lColumnWidth:=48
      .AddColumn "Subject", "Subject", lColumnWidth:=192
      .AddColumn "Status", "Status"
      .AddColumn "PercentComplete", "% Complete", sFmtString:="#0%", eSortType:=CCLSortNumeric
            
      ' add to columns menu:
      For iCol = 1 To .Columns
         If (iCol > 1) Then
            Load mnuColumn(iCol - 1)
            mnuColumn(iCol - 1).Visible = True
         End If
         mnuColumn(iCol - 1).Caption = IIf(Len(.ColumnHeader(iCol)) = 0, .ColumnKey(iCol), .ColumnHeader(iCol))
         mnuColumn(iCol - 1).Checked = .ColumnVisible(iCol)
         mnuColumn(iCol - 1).Tag = .ColumnKey(iCol)
      Next iCol
            
      .Redraw = True
            
   End With
   
End Sub

Private Sub Form_Load()
   
   ' Set up the grid:
   configureGrid
   
   ' Add some tasks:
   createDummyTaskListData
   
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   grdTasks.Move grdTasks.left, grdTasks.top, Me.ScaleWidth - grdTasks.left * 2, Me.ScaleHeight - grdTasks.top - 2 * Screen.TwipsPerPixelY
   lblInfo.Width = Me.ScaleWidth - lblInfo.left * 2
End Sub

''' <summary>
''' Clear the edit control when editing is ended in the grid.
''' </summary>
Private Sub grdTasks_CancelEdit()
   
   ' End of edit mode.  Make the text box visible.
   ' Don't use this event to update the cell's text,
   ' since it is fired for all types of cancellation,
   ' including when the user decides to alt-tab off
   ' to another app.
   txtEdit.Visible = False

End Sub


''' <summary>
''' Sort the grid's data in response to a column click.
''' </summary>
''' <param name="lCol">The column which was clicked</param>
Private Sub grdTasks_ColumnClick(ByVal lCol As Long)
Dim iSortIndex As Long
   '
   With grdTasks.SortObject
      .ClearNongrouped
      iSortIndex = .IndexOf(lCol)
      If (iSortIndex < 1) Then
         iSortIndex = .Count + 1
      End If
      .SortColumn(iSortIndex) = lCol
      .SortType(iSortIndex) = grdTasks.ColumnSortType(lCol)
      grdTasks.ColumnSortOrder(lCol) = IIf(grdTasks.ColumnSortOrder(lCol) = CCLOrderAscending, CCLOrderDescending, CCLOrderAscending)
      .SortOrder(iSortIndex) = grdTasks.ColumnSortOrder(lCol)
   End With
   
   grdTasks.Sort
   '
End Sub

Private Sub grdTasks_HeaderRightClick(ByVal x As Single, ByVal y As Single)
Dim lCol As Long
   
   lCol = grdTasks.ColumnHeaderFromPoint(x, y)
   
   If (lCol > 0) Then
      mnuContext(0).Enabled = True
      mnuContext(1).Enabled = True
      mnuContext(3).Enabled = True
      mnuContext(3).Caption = IIf(grdTasks.ColumnIsGrouped(lCol), "Don't Group By This Field", "Group By This Field")
      mnuContext(6).Enabled = True
            
      mnuContext(0).Checked = (grdTasks.ColumnSortOrder(lCol) = CCLOrderAscending)
      mnuContext(1).Checked = (grdTasks.ColumnSortOrder(lCol) = CCLOrderDescending)
      
   Else
      mnuContext(0).Enabled = False
      mnuContext(1).Enabled = False
      mnuContext(3).Enabled = False
      mnuContext(6).Enabled = False
   
      mnuContext(0).Checked = False
      mnuContext(1).Checked = False
   
   End If
      
   x = (x + grdTasks.ScrollOffsetX) * Screen.TwipsPerPixelX + grdTasks.left
   y = y * Screen.TwipsPerPixelY + grdTasks.top
   mnuContextTOP.Tag = lCol
   Me.PopupMenu mnuContextTOP, , x, y
      
End Sub

''' <summary>
''' Allows validation of data prior to cancellation of an edit control for a particular
''' cell.
''' </summary>
''' <param name="lRow">Row being edited.</param>
''' <param name="lCol">Column being edited</param>
''' <param name="newValue">Not currently used</param>
''' <param name="bStayInEditMode">Set to <c>True</c> to prevent the grid from exiting
''' edit mode if the text fails validation.  By default it is <c>False</c>.</param>
Private Sub grdTasks_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, newValue As Variant, bStayInEditMode As Boolean)
Dim sTest As String
Dim iPos As Long
Dim lTest As Long
Dim sMsg As String
Dim dDate As Variant
   '
   If (lCol = 5) Or (lCol = 6) Or (lCol = 7) Then
      If Not validDate(txtEdit.Text, dDate) Then
         sMsg = "You must enter a valid date."
         bStayInEditMode = True
      End If
      
   ElseIf (lCol = 11) Then
      sMsg = "Enter a number between 0 and 100%."
      sTest = txtEdit.Text
      iPos = InStr(sTest, "%")
      If iPos > 1 Then
         sTest = left(sTest, iPos - 1)
      End If
      If Not (IsNumeric(sTest)) Then
         bStayInEditMode = True
      Else
         On Error Resume Next
         lTest = CLng(sTest)
         If (Err.Number = 0) Then
            If (lTest < 0) Or (lTest > 100) Then
               bStayInEditMode = True
            End If
         Else
            bStayInEditMode = True
         End If
      End If
      
   End If
   
   If bStayInEditMode Then
      MsgBox sMsg, vbExclamation
      txtEdit.SetFocus
   Else
      If (lRow = 1) Then
         addNewTask
      Else
         If (lCol = 11) Then
            grdTasks.CellText(lRow, lCol) = lTest / 100#
         ElseIf (lCol = 5) Or (lCol = 6) Or (lCol = 7) Then
            grdTasks.CellText(lRow, lCol) = dDate
         Else
            grdTasks.CellText(lRow, lCol) = txtEdit.Text
         End If
      End If
   End If
   '
End Sub

''' <summary>
''' Fired when the grid detects the user wants to edit a cell.
''' </summary>
''' <param name="lRow">Row being edited.</param>
''' <param name="lCol">Column being edited</param>
''' <param name="iKeyAscii">Key which was pressed if edit mode is being started
''' from a keypress.</param>
''' <param name="bCancel">Set to <c>True</c> to prevent the grid from going
''' into edit mode.  By default it is <c>False</c>.</param>
Private Sub grdTasks_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim sText As String
Dim lLeft As Long
Dim lTop As Long
Dim lWidth As Long
Dim lHeight As Long
   
   If (grdTasks.ColumnSortType(lCol) = CCLSortIcon) Then
      
      bCancel = True
      If Not (lRow = 1) Then
         If (grdTasks.ColumnKey(lCol) = "Complete") Then
            toggleCompletion lRow
         End If
      End If
      
   Else
      ' Is this the split row?
      If (lRow = 1) Then
         If Not (grdTasks.ColumnKey(lCol) = "Subject") Then
            bCancel = True
            Exit Sub
         End If
      End If
   
      ' Get boundary of the cell:
      grdTasks.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
      
      ' Get the text:
      If Not IsMissing(grdTasks.CellText(lRow, lCol)) Then
         sText = grdTasks.CellFormattedText(lRow, lCol)
      Else
         sText = ""
      End If
      
      ' If the user has initiated edit mode by a key, we want
      ' to add this to the text.  This is really a common
      ' thing and should probably be supported automatically
      ' in the grid:
      If Not (iKeyAscii = 0) Then
         If (lRow = 1) Then
            sText = Chr(iKeyAscii)
         Else
            sText = Chr(iKeyAscii) & sText
         End If
         txtEdit.Text = sText
      '    txtEdit.SelStart = 1
      '    txtEdit.SelLength = Len(sText)
      Else
         If (lRow = 1) Then
            sText = ""
         End If
         txtEdit.Text = sText
         'txtEdit.SelStart = 0
         'txtEdit.SelLength = Len(sText)
      End If
      
      ' Set the text properties to match the grid cell being edited:
      Set txtEdit.Font = grdTasks.CellFont(lRow, lCol)
      If grdTasks.CellBackColor(lRow, lCol) = -1 Then
         txtEdit.BackColor = grdTasks.BackColor
      Else
         txtEdit.BackColor = grdTasks.CellBackColor(lRow, lCol)
      End If
      
      ' Move the text box to the edit position, make it visible and give it the focus:
      txtEdit.Move grdTasks.left + Screen.TwipsPerPixelX + lLeft, grdTasks.top + 2 * Screen.TwipsPerPixelY + lTop + (grdTasks.RowHeight(lRow) * Screen.TwipsPerPixelY - txtEdit.Height) \ 2, lWidth - 2 * Screen.TwipsPerPixelX
      txtEdit.Visible = True
      txtEdit.ZOrder
      txtEdit.SetFocus
   
      
   End If
End Sub

Private Sub mnuColumn_Click(Index As Integer)
Dim lCol As Long
   lCol = Index + 1
   mnuColumn(Index).Checked = Not (mnuColumn(Index).Checked)
   grdTasks.ColumnVisible(lCol) = mnuColumn(Index).Checked
End Sub

Private Sub mnuContext_Click(Index As Integer)
Dim lCol As Long
   Select Case Index
   Case 0
      lCol = CLng(mnuContextTOP.Tag)
      If (mnuContext(0).Checked) Then
         mnuContext(0).Checked = False
         grdTasks.ColumnSortOrder(lCol) = CCLOrderNone
      Else
         mnuContext(0).Checked = True
         grdTasks.ColumnSortOrder(lCol) = CCLOrderNone
         grdTasks_ColumnClick mnuContextTOP.Tag
      End If
   
   Case 1
      lCol = CLng(mnuContextTOP.Tag)
      If (mnuContext(1).Checked) Then
         mnuContext(1).Checked = False
         grdTasks.ColumnSortOrder(lCol) = CCLOrderNone
      Else
         mnuContext(1).Checked = True
         grdTasks.ColumnSortOrder(lCol) = CCLOrderAscending
         grdTasks_ColumnClick lCol
      End If
   
   Case 3
      lCol = CLng(mnuContextTOP.Tag)
      If (grdTasks.ColumnIsGrouped(lCol)) Then
         ' Ungroup
         grdTasks.ColumnIsGrouped(lCol) = False
      Else
         ' Group
         grdTasks.ColumnIsGrouped(lCol) = True
      End If
   Case 4
      ' Group box:
      grdTasks.AllowGrouping = Not (grdTasks.AllowGrouping)
      mnuContext(Index).Checked = grdTasks.AllowGrouping
   Case 6
      lCol = CLng(mnuContextTOP.Tag)
      grdTasks.ColumnVisible(lCol) = False
      mnuColumn(mnuContextTOP.Tag - 1).Checked = False
   End Select
   
End Sub

Private Sub mnuFile_Click(Index As Integer)
   Select Case Index
   Case 0
      Unload Me
   End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
   Select Case Index
   Case 2
      grdTasks.AllowGrouping = Not (grdTasks.AllowGrouping)
      mnuView(2).Checked = grdTasks.AllowGrouping
   End Select
End Sub

Private Sub mnuViewTOP_Click()
   mnuView(2).Checked = grdTasks.AllowGrouping
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
   '
   If (KeyAscii = 13) Then
      grdTasks_PreCancelEdit grdTasks.EditRow, grdTasks.EditCol, Empty, False
      KeyAscii = 0 ' stop beeping
   ElseIf (KeyAscii = 27) Then
      ' Get out!
      grdTasks.CancelEdit
      KeyAscii = 0 ' stop beeping
   End If
   '
End Sub


