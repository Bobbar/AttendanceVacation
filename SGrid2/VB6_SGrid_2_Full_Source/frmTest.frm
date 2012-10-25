VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form frmDemo 
   Caption         =   "SGrid Demonstrator"
   ClientHeight    =   9930
   ClientLeft      =   3345
   ClientTop       =   2325
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   8730
   Begin VB.TextBox txtEdit 
      Height          =   375
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   46
      Text            =   "frmTest.frx":0442
      Top             =   5160
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picMisc 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   3540
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   34
      Top             =   6060
      Visible         =   0   'False
      Width           =   2115
      Begin VB.CommandButton cmdCellText 
         Caption         =   "&Cell Text..."
         Height          =   375
         Left            =   1020
         TabIndex        =   37
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkBold 
         Appearance      =   0  'Flat
         Caption         =   "&Bold"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   60
         Width           =   975
      End
      Begin VB.CheckBox chkItalic 
         Appearance      =   0  'Flat
         Caption         =   "&Italic"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   36
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton cmdGetSel 
         Caption         =   "&Selected"
         Height          =   375
         Left            =   1020
         TabIndex        =   38
         Top             =   420
         Width           =   975
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdThis 
      Height          =   4515
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7964
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
   Begin VB.PictureBox picBackground 
      Height          =   1515
      Left            =   4500
      Picture         =   "frmTest.frx":0448
      ScaleHeight     =   1455
      ScaleWidth      =   1515
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3420
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   8670
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   9615
      Width           =   8730
   End
   Begin VB.Frame fraOptions 
      Height          =   9555
      Left            =   6300
      TabIndex        =   1
      Top             =   60
      Width           =   2235
      Begin VB.PictureBox picPopulationGroup 
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   60
         ScaleHeight     =   3075
         ScaleWidth      =   2055
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   6420
         Width           =   2055
         Begin VB.CommandButton cmdAddRow 
            Caption         =   "Add Row"
            Height          =   375
            Left            =   0
            TabIndex        =   42
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdDelRow 
            Caption         =   "Del Row"
            Height          =   375
            Left            =   1020
            TabIndex        =   41
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAutoRowHeight 
            Caption         =   "Fit &Heights"
            Height          =   375
            Left            =   1020
            TabIndex        =   29
            Top             =   1260
            Width           =   975
         End
         Begin VB.CommandButton cmdRemoveCol 
            Caption         =   "&Del Col..."
            Height          =   375
            Left            =   1020
            TabIndex        =   31
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton cmdAddCol 
            Caption         =   "&Add Col..."
            Height          =   375
            Left            =   0
            TabIndex        =   30
            Top             =   2160
            Width           =   975
         End
         Begin VB.CheckBox chkRnd 
            Appearance      =   0  'Flat
            Caption         =   "Ran&dom Row Heights"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   1020
            Width           =   1935
         End
         Begin VB.TextBox txtRows 
            Height          =   285
            Left            =   0
            TabIndex        =   27
            Text            =   "100"
            Top             =   720
            Width           =   2010
         End
         Begin VB.CommandButton cmdRepopulate 
            Caption         =   "&Repopulate"
            Height          =   375
            Left            =   1020
            TabIndex        =   26
            Top             =   300
            Width           =   975
         End
         Begin VB.CommandButton cmdEmpty 
            Caption         =   "&Clear"
            Height          =   375
            Left            =   0
            TabIndex        =   25
            Top             =   300
            Width           =   975
         End
         Begin VB.CheckBox chkCol4 
            Appearance      =   0  'Flat
            Caption         =   "Date Column &Visible"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   33
            ToolTipText     =   "Shows/hides the Date Column in the grid."
            Top             =   2820
            Width           =   1995
         End
         Begin VB.CheckBox chkVisible 
            Appearance      =   0  'Flat
            Caption         =   "Show &Odd Rows only"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   32
            ToolTipText     =   "Shows/Hides all the even rows in the grid using the RowVisible property."
            Top             =   2580
            Width           =   1995
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H80000010&
            Caption         =   " Population"
            ForeColor       =   &H80000016&
            Height          =   240
            Index           =   2
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   60
         ScaleHeight     =   2715
         ScaleWidth      =   2115
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2115
         Begin VB.CheckBox chkBlendSelection 
            Appearance      =   0  'Flat
            Caption         =   "&Alpha Blend Selection"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   47
            ToolTipText     =   "Toggles whether a focus rectangle is drawn around the selection when the grid is in focus."
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CheckBox chkCustomColours 
            Appearance      =   0  'Flat
            Caption         =   "C&ustom Colours"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   43
            ToolTipText     =   "Toggles a custom colour set for the grid."
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CheckBox chkDrawFocusRect 
            Appearance      =   0  'Flat
            Caption         =   "Dra&w Focus Rectangle"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   22
            ToolTipText     =   "Toggles whether a focus rectangle is drawn around the selection when the grid is in focus."
            Top             =   2160
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkHighlightSelectedIcons 
            Appearance      =   0  'Flat
            Caption         =   "Highlight Selected Ico&ns"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   21
            ToolTipText     =   "Toggles whether icons are highlighted when a cell is selected."
            Top             =   1920
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkBackground 
            Appearance      =   0  'Flat
            Caption         =   "&Background Bitmap"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   19
            ToolTipText     =   "Sets a bitmap to use as the background texture behind the grid."
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox chkOptions 
            Appearance      =   0  'Flat
            Caption         =   "&Fill Grid"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "Select if you want grid lines to repeat below the last row of the grid."
            Top             =   960
            Width           =   1755
         End
         Begin VB.CheckBox chkOptions 
            Appearance      =   0  'Flat
            Caption         =   "&Vertical"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "Toggle whether vertical grid lines are displayed."
            Top             =   720
            Value           =   1  'Checked
            Width           =   1755
         End
         Begin VB.CheckBox chkOptions 
            Appearance      =   0  'Flat
            Caption         =   "&Horizontal"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   16
            ToolTipText     =   "Toggle whether horizontal grid lines are displayed."
            Top             =   480
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkOptions 
            Appearance      =   0  'Flat
            Caption         =   "&Grid-Lines"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   15
            ToolTipText     =   "Toggle whether grid lines are shown."
            Top             =   240
            Width           =   1995
         End
         Begin VB.CheckBox chkAlternateRowColour 
            Appearance      =   0  'Flat
            Caption         =   "Alternate Row Colours"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   20
            ToolTipText     =   "Makes alternate rows render in a different colour."
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H80000010&
            Caption         =   " Appearance"
            ForeColor       =   &H80000016&
            Height          =   240
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.PictureBox picBehaviourGroup 
         BorderStyle     =   0  'None
         Height          =   2955
         Left            =   60
         ScaleHeight     =   2955
         ScaleWidth      =   2115
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   780
         Width           =   2115
         Begin VB.CheckBox chkHotTrack 
            Appearance      =   0  'Flat
            Caption         =   "&Hot Track"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   48
            ToolTipText     =   "Setting the SplitRow property causes the grid to always display the specified rows at the top of it's display."
            Top             =   2640
            Width           =   1935
         End
         Begin VB.CheckBox chkSingleClickEdit 
            Appearance      =   0  'Flat
            Caption         =   "Single Clic&k Edit"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   45
            ToolTipText     =   "In single-click edit mode, selecting a cell immediately fires a RequestEdit event and puts the cell into edit mode if required."
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chkSplitRow 
            Appearance      =   0  'Flat
            Caption         =   "&Split Row"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   44
            ToolTipText     =   "Setting the SplitRow property causes the grid to always display the specified rows at the top of it's display."
            Top             =   2400
            Width           =   1935
         End
         Begin VB.CheckBox chkAutoGrouping 
            Appearance      =   0  'Flat
            Caption         =   "A&uto Grouping"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   12
            ToolTipText     =   $"frmTest.frx":0CA4
            Top             =   2160
            Width           =   1515
         End
         Begin VB.CheckBox chkFlatHeader 
            Appearance      =   0  'Flat
            Caption         =   "&Flat Header"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "When FlatHeader is selected, the grid overdraws the 3D borders of the header items to make them appear flatter."
            Top             =   1920
            Width           =   1515
         End
         Begin VB.CheckBox chkEnabled 
            Appearance      =   0  'Flat
            Caption         =   "E&nabled"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   8
            ToolTipText     =   "Enable or disable the grid."
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkEditable 
            Appearance      =   0  'Flat
            Caption         =   "&Editable"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "Select to make the grid editable."
            Top             =   720
            Width           =   1815
         End
         Begin VB.CheckBox chkHeaderButtons 
            Appearance      =   0  'Flat
            Caption         =   "Header Bu&ttons"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   "When the grid has header buttons, you can sort the rows by clicking the header columns. Disable it by turning HeaderButtons off."
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkHeader 
            Appearance      =   0  'Flat
            Caption         =   "&Header"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   9
            ToolTipText     =   "Set whether the grid's header should be shown."
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkOptions 
            Appearance      =   0  'Flat
            Caption         =   "&Row Mode"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   6
            ToolTipText     =   "Normally, you can select single cells in the grid.  In RowMode entire rows are selected."
            Top             =   480
            Width           =   1995
         End
         Begin VB.CheckBox chkOptions 
            Appearance      =   0  'Flat
            Caption         =   "&Multi-Select"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   $"frmTest.frx":0D2F
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H80000010&
            Caption         =   " Behaviour"
            ForeColor       =   &H80000016&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   60
         Picture         =   "frmTest.frx":0DBB
         ScaleHeight     =   540
         ScaleWidth      =   2115
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   2115
      End
   End
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   5280
      Top             =   240
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   24
      IconSizeY       =   24
      ColourDepth     =   24
      Size            =   59040
      Images          =   "frmTest.frx":499D
      Version         =   131072
      KeyCount        =   24
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save..."
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuDemoTOP 
      Caption         =   "&Demos"
      Begin VB.Menu mnuDemo 
         Caption         =   "&Mailbox Style..."
         Index           =   0
      End
      Begin VB.Menu mnuDemo 
         Caption         =   "&Task List..."
         Index           =   1
      End
      Begin VB.Menu mnuDemo 
         Caption         =   "Matrix E&ditor..."
         Index           =   2
      End
      Begin VB.Menu mnuDemo 
         Caption         =   "&Rows on Demand..."
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelpTOP 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&vbAccelerator.com..."
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About..."
         Index           =   2
      End
   End
   Begin VB.Menu mnuContextTOP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuContext 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuContext 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContext 
         Caption         =   "&Copy Text"
         Index           =   2
      End
      Begin VB.Menu mnuContext 
         Caption         =   "C&lear"
         Index           =   3
      End
      Begin VB.Menu mnuContext 
         Caption         =   "&Delete Row"
         Index           =   4
      End
      Begin VB.Menu mnuContext 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContext 
         Caption         =   "&Font..."
         Index           =   6
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Foreground &Colour..."
         Index           =   7
      End
      Begin VB.Menu mnuContext 
         Caption         =   "&Background Colour..."
         Index           =   8
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================
' Name:     vbAcceleratorSGrid Control Demo
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     10 January 2004
'
' Requires: SSubTmr.DLL
'           vbalSGrid.OCX
'
' Copyright © 1998-2004 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Demonstrates the features of the vbAccelerator grid control.
'
' Features:
'
'  * Hierarchial grouping
'  * Drag-drop columns
'  * Visible or invisible columns
'  * Fixed Rows
'  * Row height can be set independently for each row
'  * MS Common Controls or vbAccelerator ImageList support
'  * Up to two icons per cell (e.g. a check box and a standard icon)
'  * Indent text within any cell
'  * Many cell text formatting options including multi-line text
'  * Owner-draw cells
'  * Mouse-over hot-tracking of cells
'  * Alpha-blended selections
'  * Show/Hide rows to allow filtering options
'  * Show/Hide columns
'  * Scroll bars implemented using true API scroll bars.
'  * Up to 2 billion rows and columns (although practically about 20,000 is the limit)
'  * Full row sorting by any number of columns at once, allows sorting by icon, text,
'    date/time or number.
'  * Autosize columns
'
' FREE SOURCE CODE - ENJOY!
' ======================================================================================

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

' Current status text:
Private m_sStatus As String
' Current progress value:
Private m_iValue As Long
' Progress Max value:
Private m_iMax As Long

' Some API calls to make the border of a object thin:
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum
' Add to translate RGB - OleColor
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub TestVeryLongText()
Dim sOut As String
Dim i As Long
   For i = 1 To 4096
      If Rnd < 0.2 Then
         sOut = sOut & " "
      Else
         sOut = sOut & Chr$(Rnd * 26 + Asc("A"))
      End If
   Next i
   grdThis.CellText(1, 5) = sOut
   
   ' test visible...
   grdThis.Redraw = False
   grdThis.CellSelected(48, 2) = True
   grdThis.Redraw = True
   
End Sub

''' <summary>
''' Switch on/off thin border on any Window with a handle and a 3D border.
''' </summary>
''' <param name="hWnd">Window handle</param>
''' <param name="bState"><c>True</c> to set thin borders, <c>False</c> to remove them.</param>
Private Sub ThinBorder(ByVal hWnd As Long, ByVal bState As Boolean)
Dim lStyle As Long
   ' Thin border:
   lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
   If bState Then
      lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
   Else
      lStyle = lStyle Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
   End If
   SetWindowLong hWnd, GWL_EXSTYLE, lStyle
   ' Make the style 'take':
   SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE

End Sub

''' <summary>
''' Draw the current status text in the "status bar" picture box.
''' </summary>
Private Function DrawStatus()
   picStatus.Cls
   If (m_iValue <> 0) Then
      picStatus.Line (0, 0)-(picStatus.ScaleWidth * m_iValue \ m_iMax, picStatus.ScaleHeight), vbButtonShadow, BF
      picStatus.ForeColor = vb3DHighlight
   Else
      picStatus.ForeColor = vbWindowText
   End If
   If (m_sStatus <> "") Then
      picStatus.CurrentX = 4 * Screen.TwipsPerPixelX
      picStatus.CurrentY = 2 * Screen.TwipsPerPixelY
      picStatus.Print m_sStatus
   End If
   picStatus.Refresh
End Function

''' <summary>
''' Set the maximum for the "progress bar" rendered in the status area.
''' </summary>
''' <param name="iMax">New maximum value</param>
Public Property Let Max(ByVal iMax As Long)
   m_iMax = iMax
   DrawStatus
End Property

''' <summary>
''' Set the value for the "progress bar" rendered in the status area.
''' </summary>
''' <param name="iValue">New value</param>
Public Property Let Value(ByVal iValue As Long)
   m_iValue = iValue
   DrawStatus
End Property

''' <summary>
''' Get the value for the "progress bar" rendered in the status area.
''' </summary>
''' <returns>progress value</returns>
Public Property Get Value() As Long
   Value = m_iValue
End Property

''' <summary>
''' Set the text to display in the status area.
''' </summary>
''' <param name="iValue">New value</param>
Public Property Let Status(ByVal sText As String)
   m_sStatus = sText
   DrawStatus
End Property

''' <summary>
''' Set up the grid properties and add the columns
''' </summary>
Private Sub initialiseGrid()
   With grdThis
   
      ' Set the grid image list.  This property can also be
      ' set to a Microsoft ImageList object:
      .ImageList = ilsIcons.hIml
      ' By default, the header uses the same IML as the grid
      .HeaderImageList = 0
      
      ' Add the columns we will use:
      .AddColumn "file", "Name", , , 32, , True, , False, , , CCLSortIcon
      .AddColumn "size", "Size", ecgHdrTextALignRight, , 48, , , , , "#,##0", , CCLSortNumeric
      .AddColumn "type", "Type"
      .AddColumn "date", "Modified", , , 64, False, , , , "Long Date", , CCLSortDate
      .AddColumn "col5", "Col 5", , , 196
      .AddColumn "col6", "Col 6"
      .AddColumn "col7", "Col 7"
      .AddColumn "col8", "Col 8", , , , , , , , , , CCLSortIcon
      .AddColumn "col9", "Col 9"
      .AddColumn "col10", "Col 10"
      
      .KeySearchColumn = .ColumnIndex("size")
      
      .DefaultRowHeight = 30

   End With
   
End Sub

''' <summary>
''' Add some demonstration data to the grid
''' </summary>
Private Sub addData()
Dim lRow As Long, lCol As Long, lIndent As Long
      
   Dim sFnt2 As New StdFont
   sFnt2.Name = "Times New Roman"
   sFnt2.Bold = True
   sFnt2.Size = 12
     
   
   With grdThis
      
      .Redraw = False
      .Rows = CLng(txtRows.Text)
      Max = .Rows
      
      ' For performance, look up the column indices once rather than
      ' each time around in the loop.  This can make population more
      ' than twice as fast
      Dim lFileCol As Long
      Dim lCol8 As Long
      Dim lSizeCol As Long
      Dim lTypeCol As Long
      Dim lDateCol As Long
      Dim lCol5 As Long
      On Error Resume Next ' Because some columns may have been deleted through the UI
      lFileCol = .ColumnIndex("file")
      lCol8 = .ColumnIndex("col8")
      lSizeCol = .ColumnIndex("size")
      lTypeCol = .ColumnIndex("type")
      lDateCol = .ColumnIndex("date")
      lCol5 = .ColumnIndex("col5")
      On Error GoTo 0
      
      For lRow = 1 To .Rows
         If (chkRnd.Value = Checked) Then
            .RowHeight(lRow) = Rnd * 48 + 16
         Else
            .RowHeight(lRow) = .DefaultRowHeight
         End If
         For lCol = 1 To .Columns
            If (lCol = lFileCol Or lCol = lCol8) Then
               .CellDetails lRow, lCol, , , Rnd * (ilsIcons.ImageCount - 1)
            ElseIf (lCol = lSizeCol) Then
               .CellDetails lRow, lCol, Int(Rnd * 1024 * 1024&), DT_RIGHT Or DT_SINGLELINE Or DT_END_ELLIPSIS
            ElseIf (lCol = lTypeCol) Then
               .CellDetails lRow, lCol, "Type " & lRow & ",Col" & lCol
            ElseIf (lCol = lDateCol) Then
               .CellDetails lRow, lCol, DateSerial(Year(Now) + Rnd * 8 - 1, Rnd * 12, Rnd * 31)
            ElseIf (lCol = lCol5) Then
               ' Icons + text
               If (lRow Mod 2) = 0 Then
                  lIndent = 24
               Else
                  lIndent = 0
               End If
               .CellDetails lRow, lCol, "This is a longer piece of text which can wrap onto a second line if the default cell format is changed so the DT_SINGLELINE option is removed. Test ampersands: Autos & Auto Parts.", DT_LEFT Or DT_MODIFYSTRING Or DT_WORDBREAK Or DT_END_ELLIPSIS, Rnd * ilsIcons.ImageCount - 1, , , , lIndent
            Else
               ' Text:
               .CellDetails lRow, lCol, "Row" & lRow & ",Col" & lCol
            End If

            ' Demonstrating multiple forecolor, backcolor and fonts for cells
            If (lRow Mod 42) = 0 Then
               .CellFont(lRow, lCol) = sFnt2
            ElseIf (lRow Mod 35) = 0 Then
               If (lCol = 4) Then
                  .CellBackColor(lRow, lCol) = &HCC9966
               Else
                  .CellBackColor(lRow, lCol) = &HEECC99
               End If
            ElseIf (lRow Mod 10) = 0 Then
               .CellForeColor(lRow, lCol) = &HFF&
            End If

         Next lCol
         If (lRow Mod 50) = 0 Then
            Value = Value + 50
            Status = lRow & " of " & .Rows
         End If
      Next lRow
      Value = 0
      .Redraw = True
   End With
   
End Sub

''' <summary>
''' Sets whether to render alternate rows in a different colour.
''' </summary>
Private Sub chkAlternateRowColour_Click()
   If (chkAlternateRowColour.Value = vbChecked) Then
      grdThis.AlternateRowBackColor = RGB(252, 252, 230)
   Else
      grdThis.AlternateRowBackColor = -1
   End If
End Sub

''' <summary>
''' Sets whether the grid allows automatic grouping or not.
''' </summary>
Private Sub chkAutoGrouping_Click()
   grdThis.AllowGrouping = (chkAutoGrouping.Value = vbChecked)
   ' Making rows visible is not (currently) allowed
   ' whilst grouping is in effect
   chkVisible.Enabled = Not (grdThis.AllowGrouping)
End Sub

''' <summary>
''' Sets whether the grid shows a background bitmap or not.
''' </summary>
Private Sub chkBackground_Click()
   If chkBackground.Value = Checked Then
      Set grdThis.BackgroundPicture = picBackground.Picture
      ' work around vb bug for JPG and GIF - picture is 2 pixels larger than expected
      grdThis.BackgroundPictureHeight = grdThis.BackgroundPictureHeight - 3
   Else
      Set grdThis.BackgroundPicture = Nothing
   End If
End Sub

Private Sub chkBlendSelection_Click()
   grdThis.SelectionAlphaBlend = chkBlendSelection.Value
   grdThis.SelectionOutline = chkBlendSelection.Value
   If (grdThis.SelectionAlphaBlend) Then
      grdThis.DrawFocusRectangle = False
      grdThis.HighlightForeColor = vbWindowText
   Else
      grdThis.DrawFocusRectangle = chkDrawFocusRect.Value
      grdThis.HighlightForeColor = vbHighlightText
   End If
End Sub

''' <summary>
''' Toggles whether the selected cell's text is bold or not
''' </summary>
Private Sub chkBold_Click()
Dim sFnt As New StdFont
   If (chkBold.Tag = "") Then
      With grdThis.CellFont(grdThis.SelectedRow, grdThis.SelectedCol)
         sFnt.Name = .Name
         sFnt.Size = .Size
         sFnt.Bold = (chkBold.Value = Checked)
         sFnt.Italic = (chkItalic.Value = Checked)
         grdThis.CellFont(grdThis.SelectedRow, grdThis.SelectedCol) = sFnt
      End With
   Else
      chkBold.Tag = ""
   End If
End Sub

''' <summary>
''' Toggles whether date column in the grid is visible.
''' </summary>
Private Sub chkCol4_Click()
   grdThis.ColumnVisible("date") = (chkCol4.Value = Checked)
End Sub

''' <summary>
''' Toggles a custom colour set
''' </summary>
Private Sub chkCustomColours_Click()

   ' Best to turn redraw off if setting multiple appearance properties
   grdThis.Redraw = False
   
   ' Set the colours:
   If (chkCustomColours.Value = Checked) Then
      grdThis.AlternateRowBackColor = RGB(86, 35, 87)
      grdThis.BackColor = RGB(72, 29, 73)
      grdThis.GridLineColor = RGB(150, 97, 153)
      grdThis.GridFillLineColor = grdThis.GridLineColor
      grdThis.ForeColor = RGB(155, 122, 158)
      grdThis.GroupingAreaBackColor = RGB(110, 46, 112)
      grdThis.GroupRowBackColor = RGB(135, 102, 138)
      grdThis.GroupRowForeColor = RGB(220, 202, 222)
      grdThis.HighlightBackColor = RGB(196, 170, 126)
      grdThis.HighlightForeColor = RGB(72, 29, 73)
      grdThis.NoFocusHighlightBackColor = RGB(135, 102, 138)
      grdThis.NoFocusHighlightForeColor = RGB(220, 202, 222)
   Else
      grdThis.AlternateRowBackColor = -1
      grdThis.BackColor = vbWindowBackground
      grdThis.GridLineColor = vbButtonFace
      grdThis.GridFillLineColor = grdThis.GridLineColor
      grdThis.ForeColor = vbWindowText
      grdThis.GroupingAreaBackColor = vbButtonShadow
      grdThis.GroupRowBackColor = vbButtonFace
      grdThis.GroupRowForeColor = vbWindowText
      grdThis.HighlightBackColor = vbHighlight
      grdThis.HighlightForeColor = vbHighlightText
      grdThis.NoFocusHighlightBackColor = vbButtonFace
      grdThis.NoFocusHighlightForeColor = vbWindowText
   End If
   
   ' Turn redraw back on
   grdThis.Redraw = True
   
End Sub

''' <summary>
''' Toggles whether the selected cell has a focus rectangle when selected
''' </summary>
Private Sub chkDrawFocusRect_Click()
   grdThis.DrawFocusRectangle = (chkDrawFocusRect.Value = Checked)
   grdThis.Draw
End Sub

''' <summary>
''' Toggles whether the grid is editable or not.
''' </summary>
Private Sub chkEditable_Click()
   grdThis.Editable = (chkEditable = Checked)
   chkSingleClickEdit.Enabled = grdThis.Editable
End Sub

''' <summary>
''' Toggles whether the grid is enabled or not.
''' </summary>
Private Sub chkEnabled_Click()
   grdThis.Enabled = (chkEnabled.Value = Checked)
End Sub

''' <summary>
''' Toggles whether the grid's header is flattened
''' </summary>
Private Sub chkFlatHeader_Click()
   grdThis.HeaderFlat = (chkFlatHeader.Value = Checked)
End Sub

''' <summary>
''' Toggles whether a header is displayed in the grid or not.
''' </summary>
Private Sub chkHeader_Click()
Dim bState As Boolean
   bState = (chkHeader.Value = Checked)
   grdThis.Header = bState
   chkHeaderButtons.Enabled = bState
   chkFlatHeader.Enabled = bState
End Sub

''' <summary>
''' Toggles whether the grid's header has buttons or not
''' </summary>
Private Sub chkHeaderButtons_Click()
   grdThis.HeaderButtons = (chkHeaderButtons.Value = Checked)
End Sub


''' <summary>
''' Toggles whether the icons are highlighted using the selection colour
''' when selected
''' </summary>
Private Sub chkHighlightSelectedIcons_Click()
   grdThis.HighlightSelectedIcons = (chkHighlightSelectedIcons.Value = Checked)
   grdThis.Draw
End Sub

Private Sub chkHotTrack_Click()
   grdThis.HotTrack = (chkHotTrack.Value = vbChecked)
End Sub

''' <summary>
''' Toggles whether the selected cell's text is italic or not
''' </summary>
Private Sub chkItalic_Click()
   chkBold_Click
End Sub

''' <summary>
''' Toggles various multi-select, row mode or grid line options
''' </summary>
Private Sub chkOptions_Click(Index As Integer)
Dim bState As Boolean
   
   bState = (chkOptions(Index).Value = vbChecked)
   Select Case Index
   Case 0
      grdThis.MultiSelect = bState
   Case 1
      grdThis.RowMode = bState
   Case 2
      grdThis.GridLines = bState
      chkOptions(3).Enabled = bState
      chkOptions(4).Enabled = bState
      chkOptions(5).Enabled = bState
   Case 3
      grdThis.NoHorizontalGridLines = Not (bState)
   Case 4
      grdThis.NoVerticalGridLines = Not (bState)
   Case 5
      If (bState) Then
         grdThis.GridLineMode = ecgGridFillControl
      Else
         grdThis.GridLineMode = ecgGridStandard
      End If
   End Select
   
End Sub

''' <summary>
''' Toggles whether cells go immediately into edit mode
''' </summary>
Private Sub chkSingleClickEdit_Click()
   grdThis.SingleClickEdit = (chkSingleClickEdit.Value = Checked)
End Sub

''' <summary>
''' Toggles whether the first row in the grid is set as the
''' split row, i.e. it shows regardless of where the grid
''' has been scrolled to.
''' </summary>
Private Sub chkSplitRow_Click()
   grdThis.SplitRow = IIf(chkSplitRow.Value = Checked, 1, 0)
End Sub

Private Sub chkVisible_Click()
Dim bS As Boolean
Dim lRow As Long
   bS = (chkVisible.Value = Unchecked)
   With grdThis
      .Redraw = False
      For lRow = 1 To .Rows
         If (lRow Mod 2) = 0 Then
            .RowVisible(lRow) = bS
         End If
      Next lRow
      .Redraw = True
   End With
End Sub

''' <summary>
''' Adds a new column to the grid
''' </summary>
Private Sub cmdAddCol_Click()
Static s_iItem As Long
   If s_iItem = 0 Then
      s_iItem = grdThis.Columns
   End If
   With grdThis
      .AddColumn "New" & s_iItem, "New:" & s_iItem
   End With
End Sub

''' <summary>
''' Inserts a new row into the grid at position 1
''' </summary>
Private Sub cmdAddRow_Click()
   '
   If (grdThis.Rows > 0) Then
      grdThis.AddRow 1
   Else
      grdThis.AddRow
   End If
   '
End Sub

''' <summary>
''' Auto-sizes all of the rows to fit their contents
''' given the current column sizes.
''' </summary>
Private Sub cmdAutoRowHeight_Click()
Dim lRow As Long
   Screen.MousePointer = vbHourglass
   With grdThis
      .Redraw = False
      For lRow = 1 To .Rows
         .AutoHeightRow lRow
      Next lRow
      .Redraw = True
   End With
   Screen.MousePointer = vbDefault
End Sub

''' <summary>
''' Allows the selected cell's text to be changed through a dialog.
''' </summary>
Private Sub cmdCellText_Click()
Dim sText As String
Dim sI As String
Dim iCol As Long

   If (grdThis.RowMode) Then
      ' When in row mode use the long text column
      iCol = 5
   Else
      ' Otherwise use the selected column:
      iCol = grdThis.SelectedCol
   End If
   
   ' Get the current text
   sText = grdThis.CellText(grdThis.SelectedRow, iCol)
   ' Use nasty VB input box dialog to get text:
   sI = InputBox$("Enter text", , sText)
   If (Len(sI) > 0) Then ' surely some way to determine whether cancel clicked is *obviously* required?
      ' Change the text
      grdThis.CellText(grdThis.SelectedRow, iCol) = sI
   End If
   
End Sub

''' <summary>
''' Delete the first row in the grid
''' </summary>
Private Sub cmdDelRow_Click()
   '
   If (grdThis.Rows > 0) Then
      grdThis.RemoveRow 1
   End If
   '
End Sub

''' <summary>
''' Clear all grid content
''' </summary>
Private Sub cmdEmpty_Click()
   grdThis.Clear
End Sub

''' <summary>
''' Get information about the selection and print to debug
''' </summary>
Private Sub cmdGetSel_Click()
Dim iRow As Long, iCol As Long
   With grdThis
      For iRow = 1 To .Rows
         If .RowMode Then
            If .CellSelected(iRow, 1) Then
               Debug.Print "SELECTED:" & iRow
            End If
         Else
            For iCol = 1 To .Columns
               If .CellSelected(iRow, iCol) Then
                  Debug.Print "SELECTED:" & iRow, iCol
               End If
            Next iCol
         End If
      Next iRow
   End With
End Sub

''' <summary>
''' Remove a column from the grid;
''' </summary>
Private Sub cmdRemoveCol_Click()
Dim iCol As Long
Dim sKey As String
Dim sI As String
Dim sDefault As String
   If (grdThis.Columns > 0) Then
      For iCol = 1 To grdThis.Columns
         sKey = sKey & grdThis.ColumnKey(iCol) & ","
      Next iCol
      sKey = left$(sKey, Len(sKey) - 1)
      sI = InputBox$("Enter column to delete" & vbCrLf & vbCrLf & "Available columns: " & sKey, , grdThis.ColumnKey(1))
      If (sI <> "") Then
         grdThis.RemoveColumn sI
      End If
   Else
      MsgBox "No columns to delete.", vbInformation
   End If
End Sub

''' <summary>
''' Repopulate with data
''' </summary>
Private Sub cmdRepopulate_Click()
   
   Dim lT As Long
   lT = timeGetTime()

   ' Add some data:
   addData
   
   m_sStatus = grdThis.Rows & "rows, " & timeGetTime() - lT & "ms"
   DrawStatus
   
End Sub

''' <summary>
''' Initialise the status bar and load some demonstration data
''' into the grid.
''' </summary>
Private Sub Form_Load()
   
   ThinBorder picStatus.hWnd, True
   
   Me.Show
   Me.Refresh
   
   grdThis.Redraw = False
   
   initialiseGrid
   addData
   
   grdThis.Redraw = True
   
End Sub

''' <summary>
''' Resize the controls on the form
''' </summary>
Private Sub Form_Resize()
Dim lSize As Long
Dim lHeight As Long

   On Error Resume Next

   lHeight = Me.ScaleHeight - picStatus.Height - 4 * Screen.TwipsPerPixelY
   lSize = fraOptions.Width + grdThis.left
   grdThis.Move 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, Me.ScaleWidth - grdThis.left - lSize, lHeight
   fraOptions.Move Me.ScaleWidth - lSize, grdThis.top - 6 * Screen.TwipsPerPixelY, fraOptions.Width, lHeight + 6 * Screen.TwipsPerPixelY
   picStatus.Move grdThis.left, Me.ScaleHeight - picStatus.Height - Screen.TwipsPerPixelY, Me.ScaleWidth - grdThis.left * 2
   
End Sub

''' <summary>
''' Clear the edit control when editing is ended in the grid.
''' </summary>
Private Sub grdThis_CancelEdit()
   
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
Private Sub grdThis_ColumnClick(ByVal lCol As Long)
Dim sTag As String
Dim iSortIndex As Long
      
   With grdThis.SortObject
      
      ' This demo allows grouping.  When a column is clicked
      ' for sorting, we only want to remove any grouped rows:
      .ClearNongrouped
      
      ' See if this column is already in the sort object:
      iSortIndex = .IndexOf(lCol)
      If (iSortIndex = 0) Then
         ' If not, we add it:
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lCol
      End If
   
      ' Determine which sort order to apply:
      sTag = grdThis.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      grdThis.ColumnTag(lCol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = grdThis.ColumnSortType(lCol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   grdThis.Sort
   Screen.MousePointer = vbDefault
   
End Sub

''' <summary>
''' Respond to column width changes.
''' </summary>
''' <param name="lCol">Column whose size is being changed.</param>
''' <param name="lWidth">New width of the column</param>
''' <param name="bCancel">Whether to cancel sizing or not</param>
Private Sub grdThis_ColumnWidthChanging(ByVal lCol As Long, lWidth As Long, bCancel As Boolean)
   ' If column 1 then prevent size change;
   If (grdThis.ColumnKey(lCol) = "file") Then
      bCancel = True
   End If

End Sub

Private Sub grdThis_HotItemChange(ByVal lRow As Long, ByVal lCol As Long)
   '
   'Debug.Print "HotItem: " & grdThis.CellText(lRow, lCol)
   '
End Sub

''' <summary>
''' Respond to mouse down events:
''' </summary>
''' <param name="Button">Mouse buttons.</param>
''' <param name="Shift">Shift keys pressed, if any.</param>
''' <param name="X">X position of the mouse relative to the control</param>
''' <param name="Y">Y position of the mouse relative to the control</param>
''' <param name="bDoDefault">Whether to perform the default action or not</param>
Private Sub grdThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, bDoDefault As Boolean)
   
   ' This would allow you to have a sort of simple select mode
   ' where any selection is added to the existing selection:
   'Shift = vbCtrlMask
   
End Sub

Private Sub grdThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim bSelection As Boolean
   bSelection = ((grdThis.SelectedRow > 0) And (grdThis.SelectedCol > 0))
   If (bSelection) Then
      If (Button = vbRightButton) Then
         If (bSelection) Then
            mnuContext(0).Enabled = grdThis.Editable
            mnuContext(2).Enabled = Not IsMissing(grdThis.CellText(grdThis.SelectedRow, grdThis.SelectedCol))
            
            Me.PopupMenu mnuContextTOP, , x + grdThis.left, y + grdThis.top
         End If
      Else
         ' Check the cell boundary:
         Dim lLeft As Long
         Dim lTop As Long
         Dim lWidth As Long
         Dim lHeight As Long
         grdThis.CellBoundary grdThis.SelectedRow, grdThis.SelectedCol, lLeft, lTop, lWidth, lHeight
         'Debug.Print lLeft, lTop, lWidth, lHeight
      End If
   End If
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
Private Sub grdThis_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, newValue As Variant, bStayInEditMode As Boolean)
      
   If (txtEdit.Text = "") Then
      Status = "Enter some text."
      ' This would be a good place for a popup message bubble
      ' either use the OS or use a VB window that's
      ' transparent to the mouse by subclassing WM_NCHITTEST = HT_NOWHERE
      MsgBox "Please enter some text into the cell.", vbExclamation
      bStayInEditMode = True
   Else
      Status = "Ready"
      grdThis.CellText(grdThis.EditRow, grdThis.EditCol) = txtEdit.Text
   End If
   
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
Private Sub grdThis_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
Dim sText As String
   
   
   ' Don't allow editing the icon-only columns:
   If (grdThis.ColumnKey(lCol) = "file") Or (grdThis.ColumnKey(lCol) = "col8") Then
      bCancel = True
      Exit Sub
   End If
   
   ' Get boundary of the cell:
   grdThis.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
   
   ' Get the text:
   If Not IsMissing(grdThis.CellText(lRow, lCol)) Then
      sText = grdThis.CellFormattedText(lRow, lCol)
   Else
      sText = ""
   End If
   
   ' If the user has initiated edit mode by a key, we want
   ' to add this to the text.  This is really a common
   ' thing and should probably be supported automatically
   ' in the grid:
   If Not (iKeyAscii = 0) Then
      sText = Chr$(iKeyAscii) & sText
      txtEdit.Text = sText
      txtEdit.SelStart = 1
      txtEdit.SelLength = Len(sText)
   Else
      txtEdit.Text = sText
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(sText)
   End If
   
   ' Set the text properties to match the grid cell being edited:
   Set txtEdit.Font = grdThis.CellFont(lRow, lCol)
   If grdThis.CellBackColor(lRow, lCol) = -1 Then
      txtEdit.BackColor = grdThis.BackColor
   Else
      txtEdit.BackColor = grdThis.CellBackColor(lRow, lCol)
   End If
   
   ' Move the text box to the edit position, make it visible and give it the focus:
   txtEdit.Move lLeft + grdThis.left, lTop + grdThis.top + Screen.TwipsPerPixelY, lWidth, lHeight
   txtEdit.Visible = True
   txtEdit.ZOrder
   txtEdit.SetFocus
   
End Sub

''' <summary>
''' Raised when the grid's selection changes.
''' </summary>
''' <param name="lRow">New selected row.</param>
''' <param name="lCol">New selected column.</param>
Private Sub grdThis_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
   Status = "Selected: " & lRow & "," & lCol
   chkBold.Tag = "CODE"
   chkBold.Value = Abs(grdThis.CellFont(lRow, lCol).Bold)
   chkBold.Tag = ""
   chkItalic.Tag = "CODE"
   chkItalic.Value = Abs(grdThis.CellFont(lRow, lCol).Italic)
   chkItalic.Tag = ""
End Sub

Private Sub mnuContext_Click(Index As Integer)
Dim cD As cCommonDialog
   Select Case Index
   Case 0
      ' edit mode
      grdThis.StartEdit grdThis.SelectedRow, grdThis.SelectedCol
   Case 2
      Clipboard.Clear
      Clipboard.SetText grdThis.CellText(grdThis.SelectedRow, grdThis.SelectedCol)
   Case 3
      grdThis.CellText(grdThis.SelectedRow, grdThis.SelectedCol) = Empty
   Case 4
      grdThis.RemoveRow grdThis.SelectedRow
   Case 6
      Set cD = New cCommonDialog
      Dim iFnt As IFont
      Dim sFnt As StdFont
      Set iFnt = grdThis.CellFont(grdThis.SelectedRow, grdThis.SelectedCol)
      iFnt.Clone sFnt
      If cD.VBChooseFont(sFnt, Owner:=Me.hWnd) Then
         grdThis.CellFont(grdThis.SelectedRow, grdThis.SelectedCol) = sFnt
      End If
   Case 7
      Set cD = New cCommonDialog
      Dim lColor As Long
      lColor = grdThis.CellForeColor(grdThis.SelectedRow, grdThis.SelectedCol)
      If (lColor = -1) Then
         lColor = grdThis.ForeColor
      End If
      OleTranslateColor lColor, 0, lColor
      If cD.VBChooseColor(lColor, FullOpen:=True, Owner:=Me.hWnd) Then
         grdThis.CellForeColor(grdThis.SelectedRow, grdThis.SelectedCol) = lColor
      End If
   Case 8
      Set cD = New cCommonDialog
      lColor = grdThis.CellBackColor(grdThis.SelectedRow, grdThis.SelectedCol)
      If (lColor = -1) Then
         lColor = grdThis.BackColor
      End If
      OleTranslateColor lColor, 0, lColor
      If cD.VBChooseColor(lColor, FullOpen:=True, Owner:=Me.hWnd) Then
         grdThis.CellBackColor(grdThis.SelectedRow, grdThis.SelectedCol) = lColor
      End If
   End Select
End Sub

''' <summary>
''' Fired when the demo menu subitems are clicked.
''' cell.
''' </summary>
''' <param name="Index">Index of clicked menu item.</param>
Private Sub mnuDemo_Click(Index As Integer)

   ' Show other demonstration forms:
   Select Case Index
   Case 0
      frmOutlookDemo.Show
   Case 1
      frmTaskList.Show
   Case 2
      frmMatrixDemo.Show
   Case 3
      frmOnDemand.Show
   End Select
   
   
End Sub

Private Sub mnuFile_Click(Index As Integer)
   Select Case Index
   Case 0
      LoadGridData
   Case 1
      SaveGridData
   Case 3
      Unload Me
   End Select
End Sub

Private Sub LoadGridData()
Dim sFile As String
Dim cC As New cCommonDialog
Dim iFIle As Integer
On Error GoTo ErrorHandler
   If (cC.VBGetOpenFileName(sFile, Filter:="SGrid Data Files (*.sgd)|*.sgd|All Files (*.*)|*.*", DefaultExt:="SGD", Owner:=Me.hWnd)) Then
      grdThis.LoadGridData sFile
   End If
   Exit Sub
ErrorHandler:
   MsgBox Err.Description, vbExclamation
   Exit Sub
End Sub
Private Sub SaveGridData()
Dim sFile As String
Dim cC As New cCommonDialog
Dim iFIle As Integer
On Error GoTo ErrorHandler
   If (cC.VBGetSaveFileName(sFile, Filter:="SGrid Data Files (*.sgd)|*.sgd|All Files (*.*)|*.*", DefaultExt:="SGD", Owner:=Me.hWnd)) Then
      killFileIfExists sFile
      grdThis.SaveGridData sFile
   End If
   Exit Sub
ErrorHandler:
   MsgBox Err.Description, vbExclamation
   Exit Sub
End Sub
Private Sub killFileIfExists(ByVal sFile As String)
Dim sDir As String
   On Error Resume Next
   sDir = Dir(sFile)
   If Len(sDir) > 0 Then
      On Error GoTo 0
      Kill sDir
   End If
End Sub

Private Sub mnuHelp_Click(Index As Integer)
   Select Case Index
   Case 0
      ShellExecute Me.hWnd, "open", "http://vbaccelerator.com/", "", "", SW_SHOWNORMAL
   Case 2
      frmAbout.Show vbModal, Me
   End Select
End Sub

''' <summary>
''' Customise edit cancellation or end when a key down event occurs
''' in the edit control.
''' </summary>
''' <param name="KeyCode">Key which was pressed.</param>
''' <param name="Shift">Shift state.</param>
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If (KeyCode = vbKeyReturn) Then
      ' Request Commit edit.  This will fire the
      ' grid's PreCancelEdit event, which gives you
      ' an opportunity to validate the data and put
      ' it in the cell if good.  The CancelEdit
      ' event will then fire afterwards.
      grdThis.EndEdit
   ElseIf (KeyCode = vbKeyEscape) Then
      ' Cancel edit.  This skips PreCancelEdit and
      ' fires the CancelEdit event
      grdThis.CancelEdit
   ElseIf (grdThis.SingleClickEdit) Then
      Select Case KeyCode
      
      End Select
   End If
   
End Sub

