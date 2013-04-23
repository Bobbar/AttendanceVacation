VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmVacations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vacations"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVacations.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   9840
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   2400
      TabIndex        =   44
      Top             =   1020
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Frame Frame3 
      Caption         =   "Functions"
      Height          =   1335
      Left            =   4680
      TabIndex        =   38
      Top             =   5520
      Visible         =   0   'False
      Width           =   4755
      Begin VB.CommandButton cmdAttenReports 
         Caption         =   "Attendance Reports"
         Height          =   360
         Left            =   1740
         TabIndex        =   41
         Top             =   540
         Width           =   1770
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Close"
         Height          =   360
         Left            =   3660
         TabIndex        =   40
         Top             =   540
         Width           =   990
      End
      Begin VB.CommandButton cmdVacaReports 
         Caption         =   "Vacation Reports"
         Height          =   360
         Left            =   120
         TabIndex        =   39
         Top             =   540
         Width           =   1470
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Queries"
      Height          =   1335
      Left            =   2580
      TabIndex        =   24
      Top             =   3900
      Width           =   4755
      Begin VB.CheckBox chkTaken 
         Caption         =   "Taken"
         Height          =   195
         Left            =   2040
         TabIndex        =   36
         Top             =   840
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkReScheduled 
         Caption         =   "ReScheduled"
         Height          =   195
         Left            =   2040
         TabIndex        =   35
         Top             =   600
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkRequested 
         Caption         =   "Requested"
         Height          =   195
         Left            =   2040
         TabIndex        =   34
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   360
         Left            =   3540
         TabIndex        =   33
         Top             =   540
         Width           =   990
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All Periods"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1020
         Width           =   1695
      End
      Begin VB.TextBox txtXYears 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Text            =   "2"
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optXYears 
         Caption         =   "Option1"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   195
      End
      Begin VB.OptionButton optPrevYear 
         Caption         =   "Previous Period"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optCurrentYear 
         Caption         =   "Current Period"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periods Ago"
         Height          =   195
         Left            =   780
         TabIndex        =   29
         Top             =   765
         Width           =   855
      End
   End
   Begin VB.Frame frmEntries 
      Caption         =   "Entries"
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   5700
      Width           =   9615
      Begin vbAcceleratorSGrid6.vbalGrid Grid1 
         Height          =   4095
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7223
         RowMode         =   -1  'True
         GridLines       =   -1  'True
         BackgroundPicture=   "frmVacations.frx":0CCA
         BackgroundPictureHeight=   128
         BackgroundPictureWidth=   128
         BackColor       =   -2147483633
         GridLineColor   =   4210752
         HighlightForeColor=   4210752
         NoFocusHighlightBackColor=   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScrollBarStyle  =   2
         Editable        =   -1  'True
         DisableIcons    =   -1  'True
         HotTrack        =   -1  'True
         SelectionAlphaBlend=   -1  'True
         SelectionOutline=   -1  'True
      End
      Begin VB.Shape Shape1 
         Height          =   4095
         Left            =   120
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Entries"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3960
         TabIndex        =   17
         Top             =   2040
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   9615
      Begin VB.TextBox txtHours 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4380
         TabIndex        =   51
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Timer tmrButtonEnabler 
         Interval        =   100
         Left            =   120
         Top             =   1020
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   360
         Left            =   4860
         TabIndex        =   47
         Top             =   1860
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.ComboBox cmbStatus2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7620
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   480
         Width           =   1635
      End
      Begin VB.Timer tmrLiveSearch 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   120
         Top             =   1500
      End
      Begin VB.CommandButton cmdSpellChk 
         Caption         =   "Spell Check"
         Height          =   240
         Left            =   6120
         TabIndex        =   23
         Top             =   1800
         Width           =   990
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3300
         TabIndex        =   18
         Top             =   1860
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   360
         Left            =   8520
         TabIndex        =   12
         Top             =   1860
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   11
         Top             =   1860
         Width           =   1215
      End
      Begin VB.ComboBox cmbStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtNotes 
         Height          =   855
         Left            =   2280
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker DTEndDate 
         Height          =   375
         Left            =   2340
         TabIndex        =   7
         Top             =   480
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483635
         Format          =   246022145
         CurrentDate     =   40935
      End
      Begin MSComCtl2.DTPicker DTStartDate 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483635
         Format          =   246022145
         CurrentDate     =   40935
      End
      Begin VB.Label lblLastModified 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   7680
         TabIndex        =   53
         Top             =   1200
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour"
         Height          =   195
         Left            =   4380
         TabIndex        =   52
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paid/UnPaid"
         Height          =   195
         Left            =   7620
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   195
         Left            =   1800
         TabIndex        =   22
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   5460
         TabIndex        =   21
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   195
         Left            =   2340
         TabIndex        =   8
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame frmEmployee 
      Caption         =   "Employee"
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.TextBox txtEmpName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   43
         Text            =   "EmpName"
         Top             =   540
         Width           =   3435
      End
      Begin VB.TextBox txtEmpNum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         TabIndex        =   42
         Text            =   "EmpNum"
         Top             =   540
         Width           =   1395
      End
      Begin VB.CommandButton cmdOverride 
         Appearance      =   0  'Flat
         Caption         =   "Edit"
         Height          =   240
         Left            =   8940
         TabIndex        =   37
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         Height          =   195
         Left            =   2340
         TabIndex        =   50
         Top             =   300
         Width           =   3360
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #:"
         Height          =   195
         Left            =   300
         TabIndex        =   49
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label lblTakenHours 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HoursTaken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7860
         TabIndex        =   20
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hours Taken:"
         Height          =   195
         Left            =   6420
         TabIndex        =   19
         Top             =   720
         Width           =   1230
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   6060
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hours Available:"
         Height          =   195
         Left            =   6480
         TabIndex        =   16
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label lblLabel1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Years Worked:"
         Height          =   195
         Left            =   6420
         TabIndex        =   15
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label lblHireLable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hire Date:"
         Height          =   195
         Left            =   6420
         TabIndex        =   14
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label lblYearsWorked 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YearsWorked"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7860
         TabIndex        =   13
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label lblVacaHours 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VacaHours"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7860
         TabIndex        =   4
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblHireDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HireDate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7860
         TabIndex        =   3
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Line Line2 
      X1              =   5520
      X2              =   7980
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Period"
      Height          =   195
      Left            =   4380
      TabIndex        =   31
      Top             =   5280
      Width           =   1050
   End
   Begin VB.Line Line5 
      X1              =   1920
      X2              =   1920
      Y1              =   5760
      Y2              =   5400
   End
   Begin VB.Line Line3 
      X1              =   7980
      X2              =   7980
      Y1              =   5820
      Y2              =   5400
   End
   Begin VB.Label lblCurrentPeriod 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CurrentPeriod"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   30
      Top             =   5460
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   1920
      X2              =   4260
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
   End
End
Attribute VB_Name = "frmVacations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strGUID             As String, strSelGUID As String
Public intPeriod           As Integer
Private strNotes           As String
Private UpdateMode         As Boolean
Private intSearchWaitTicks As Integer
Private Sub PrintSGrid(FlexGrid As vbalGrid, _
                       Optional sTitle As String, _
                       Optional sSubTitle As String)
    FlexGrid.Redraw = False
    Dim intPadding      As Integer
    Dim PrevX           As Integer, PrevY As Integer, intMidStart As Integer, intMidLen As Integer, intTotLen As Integer, intPossibleLen As Integer
    Dim strSizedTxt     As String, strOrigTxt As String
    Dim bolLongLine     As Boolean, bolFirstLoop As Boolean
    Dim TwipPix         As Long
    Dim intCenterOffset As Long
    Dim intColumns      As Integer
    Dim lngYTopOfGrid   As Long
    Dim lngStartY       As Long, lngStartX As Long, lngEndX As Long, lngEndY As Long
    intColumns = 6 'FlexGrid.Columns
    bolLongLine = False
    'On Error Resume Next
    Printer.Font.Underline = False
    Printer.ScaleMode = 1
    Printer.Orientation = vbPRORLandscape
    Printer.DrawWidth = 1
    Printer.DrawStyle = vbSolid
    Dim xmax As Single, xmin As Single
    xmin = 300
    xmax = 14800
    Dim ymax As Single, ymin As Single
    ymin = 1500
    ymax = 10800
    With Printer
        .ScaleMode = 1
        Printer.Print
        .FontSize = 20
        Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(sTitle) / 2)
        Printer.Print sTitle
        Printer.FontSize = 8
    End With
    Printer.FontSize = 7
    Printer.Print "    " & sSubTitle
    Printer.Print ""
    Printer.Print "    Report date: " & Date & " " & Time & "      Printed by: " & UCase$(Environ$("USERNAME"))
    Const GAP = 40
    With Printer.Font
        .Name = FlexGrid.Font.Name
        .Size = 9
    End With
    Printer.Print ""
    Printer.DrawStyle = vbDash
    Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY), vbBlack
    Printer.DrawStyle = vbSolid
    Printer.Print ""
    If optCurrentYear Then
        PrevY = Printer.CurrentY
        Dim xBoxEnd As Single, lngCenterXStartPos As Long
        'strReportInfo = "Hire Date: " & strCurrentEmpInfo.HireDate
        Printer.Font.Size = 7
        lngCenterXStartPos = (xmax / 2) - (2000 / 2) 'Printer.TextWidth(strReportInfo)
        xBoxEnd = lngCenterXStartPos + 2000 'Printer.TextWidth(strReportInfo)
        Printer.Line (lngCenterXStartPos, PrevY)-(xBoxEnd, Printer.CurrentY + (Printer.TextHeight(strReportInfo) * 5)), &H80000016, BF
        Printer.Font.Bold = True
        Printer.CurrentX = (xmax / 2) - (Printer.TextWidth("Vacation Stats") / 2)
        Printer.CurrentY = PrevY
        Printer.Print "Vacation Stats"
        Printer.Font.Bold = False
        strReportInfo = "Hire Date: " & strCurrentEmpInfo.HireDate
        Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(strReportInfo) / 2)
        Printer.Print strReportInfo
        strReportInfo = "Years Worked: " & CalcYearsWorked(strCurrentEmpInfo.Number).YearsWorked
        Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(strReportInfo) / 2)
        Printer.Print strReportInfo
        strReportInfo = "Hours Taken: " & lblTakenHours.Caption
        Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(strReportInfo) / 2)
        Printer.Print strReportInfo
        strReportInfo = "Hours Available: " & lblVacaHours.Caption
        Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(strReportInfo) / 2)
        Printer.Print strReportInfo
        '    Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(strReportEntryCount) / 2)
        '    Printer.Print strReportEntryCount
        Printer.Line (lngCenterXStartPos, PrevY)-(xBoxEnd, Printer.CurrentY), vbBlack, B
        Printer.Print ""
        Printer.Font.Size = 9
    End If
    Printer.DrawStyle = vbSolid
    Dim X As Single, XFirstColumn As Single
    Dim c As Integer, cc As Integer
    Dim R As Integer
    intMidStart = 1
    With FlexGrid
        PrevX = Printer.CurrentX
        PrevY = Printer.CurrentY
        Printer.CurrentX = xmax - 600
        Printer.CurrentY = ymax + 300
        Printer.ForeColor = vbBlack
        Printer.Font.Underline = False
        Printer.Print "Page " & Printer.Page
        Printer.CurrentX = PrevX
        Printer.CurrentY = PrevY
        intPadding = 150
        frmPBar.PBar1.Max = .Rows
        frmPBar.PBar1.Value = 0
        Form1.tmrUpdateTimeRemaining.Enabled = False
        frmPBar.lblInfo.Caption = "Spooling..."
        'TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
        XFirstColumn = xmin '+ TwipPix * GAP
        X = xmin + GAP
        lngYTopOfGrid = Printer.CurrentY
        If FlexGrid.Header = True Then
            For c = 1 To intColumns
                Printer.CurrentX = X
                TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                PrevY = Printer.CurrentY
                If c = intColumns Then '- 1
                    lngStartY = Printer.CurrentY + 5
                    lngStartX = Printer.CurrentX - GAP + 5
                    lngEndX = xmax
                    lngEndY = Printer.CurrentY + Printer.TextHeight(.ColumnHeader(c)) + GAP
                    Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), &H80000016, BF
                Else
                    lngStartY = Printer.CurrentY + 5
                    lngStartX = Printer.CurrentX - GAP + 5
                    lngEndX = Printer.CurrentX + TwipPix + GAP
                    lngEndY = Printer.CurrentY + Printer.TextHeight(.ColumnHeader(c)) + GAP
                    Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), &H80000016, BF
                End If
                Printer.CurrentX = lngStartX + GAP
                Printer.CurrentY = PrevY
                Printer.Print BoundedText(Printer, .ColumnHeader(c), TwipPix);
                TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                X = X + TwipPix + 2 * GAP
            Next c
            Printer.CurrentY = Printer.CurrentY + GAP
        End If
        Printer.Print
        For R = 1 To .Rows - 1
            If bolStop = True Then
                Printer.EndDoc
                bolStop = False
                frmPBar.Visible = False
                Exit Sub
            End If
            frmPBar.PBar1.Value = R
            frmPBar.lblQryTime = "Row " & R & " of " & .Rows
            DoEvents
            ' Draw a line above this row.
            If R > 0 Then
                Printer.Line (XFirstColumn, Printer.CurrentY)-(xmax, Printer.CurrentY), vbBlack
            End If
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Print the entries on this row.
            X = xmin + GAP
            For c = 1 To intColumns
                If frmPrinters.optCenterJust And c < intColumns Then
                    intCenterOffset = ((.ColumnWidth(c) * Screen.TwipsPerPixelX) / 2) - (Printer.TextWidth(.CellText(R, c)) / 2)
                Else
                    intCenterOffset = 0
                End If
                Printer.CurrentX = X
                If .CellText(R, c) <> "" And Printer.TextWidth(.CellText(R, c)) + intPadding >= xmax - Printer.CurrentX Then
                    lngStartY = Printer.CurrentY + Printer.TextHeight(.CellText(R, c))
                    strOrigTxt = .CellText(R, c)
                    intTotLen = 1
                    bolFirstLoop = True
                    Do Until intTotLen >= Len(strOrigTxt)
                        intMidLen = Len(strOrigTxt) - intMidStart + 1
                        If Not bolFirstLoop Then
                            strSizedTxt = Mid$(strOrigTxt, intMidStart, intPossibleLen)
                        Else
                            strSizedTxt = strOrigTxt
                        End If
                        Dim lngColEnd  As Long
                        Dim intOrigLen As Integer
                        intOrigLen = Len(strOrigTxt)
                        lngColEnd = xmax - Printer.CurrentX
                        Do Until Printer.TextWidth(strSizedTxt) + intPadding <= lngColEnd Or intTotLen >= intOrigLen
                            strSizedTxt = Mid$(strOrigTxt, intMidStart, intMidLen)
                            intMidLen = intMidLen - 1
                        Loop
                        If Not bolFirstLoop Then
                            intMidStart = intMidStart + intPossibleLen
                        Else
                            intPossibleLen = Len(strSizedTxt)
                            intMidStart = intMidStart + intMidLen + 1
                        End If
                        bolFirstLoop = False
                        intTotLen = intTotLen + Len(strSizedTxt) + 1
                        Printer.Font.Underline = .CellFont(R, c).Underline
                        If .CellFont(R, c).Underline = True Then
                            Printer.ForeColor = vbBlack
                        Else
                            Printer.ForeColor = &H404040
                        End If
                        Printer.Print strSizedTxt
                        lngEndY = Printer.CurrentY + GAP
                        PrevY = Printer.CurrentY
                        Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor(R, 4), BF
                        Printer.CurrentY = PrevY + 5
                        If Printer.CurrentY >= ymax Then ' new page
                            Printer.Line (XFirstColumn, lngYTopOfGrid)-(xmax, Printer.CurrentY + GAP), vbBlack, B
                            X = xmin
                            For cc = 1 To intColumns - 1
                                TwipPix = .ColumnWidth(cc) * Screen.TwipsPerPixelX
                                X = X + TwipPix + 2 * GAP
                                Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack
                            Next cc
                            Printer.NewPage
                            Printer.CurrentX = xmax - 600
                            Printer.CurrentY = ymax + 300
                            Printer.ForeColor = vbBlack
                            Printer.Font.Underline = False
                            Printer.Print "Page " & Printer.Page
                            Printer.CurrentX = xmin
                            ymin = 400
                            lngYTopOfGrid = ymin
                            Printer.CurrentY = ymin
                            lngStartY = Printer.CurrentY
                        End If
                        Printer.CurrentX = X + GAP
                        strSizedTxt = ""
                    Loop
                    intMidStart = 1
                    intMidLen = 0
                    intTotLen = 0
                    strSizedTxt = ""
                    bolLongLine = True
                Else
                    PrevY = Printer.CurrentY - GAP
                    If c = 4 Then
                        lngStartY = Printer.CurrentY - GAP + 5
                        lngStartX = Printer.CurrentX - GAP + 5
                        lngEndX = Printer.CurrentX + .ColumnWidth(c) * Screen.TwipsPerPixelX + GAP - 5
                        lngEndY = Printer.CurrentY + Printer.TextHeight(.CellText(R, c)) + GAP - 5
                        Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor(R, c), BF
                    End If
                    Printer.CurrentX = X + intCenterOffset
                    TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                    Printer.Font.Underline = .CellFont(R, c).Underline
                    If .CellFont(R, c).Underline = True Then
                        Printer.ForeColor = vbBlack
                    Else
                        Printer.ForeColor = &H404040
                    End If
                    Printer.CurrentX = X + intCenterOffset
                    Printer.CurrentY = PrevY + GAP
                    Printer.Print BoundedText(Printer, .CellText(R, c), TwipPix);
                End If
                TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                X = X + TwipPix + 2 * GAP
            Next c
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Move to the next line.
            If bolLongLine = True Then
                bolLongLine = False
            Else
                Printer.Print
                bolLongLine = False
            End If
            ' if near end of page, start a new one
            If Printer.CurrentY >= ymax And R < .Rows Then
                Printer.Line (XFirstColumn, lngYTopOfGrid)-(xmax, Printer.CurrentY), vbBlack, B
                X = xmin
                For c = 1 To intColumns - 1
                    TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                    X = X + TwipPix + 2 * GAP
                    Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack
                Next c
                Printer.NewPage
                Printer.CurrentX = xmax - 600
                Printer.CurrentY = ymax + 300
                Printer.ForeColor = vbBlack
                Printer.Font.Underline = False
                Printer.Print "Page " & Printer.Page
                Printer.CurrentX = xmin
                ymin = 400
                lngYTopOfGrid = ymin
                Printer.CurrentY = ymin
            End If
        Next R
        ymax = Printer.CurrentY
        'Draw a box around everything.
        Printer.Line (XFirstColumn, lngYTopOfGrid)-(xmax, ymax), vbBlack, B
        X = xmin
        ' Draw lines between the columns.
        For c = 1 To intColumns - 1 '3
            TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
            X = X + TwipPix + 2 * GAP
            Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack
        Next c
        ' End If
    End With
    frmPBar.Hide
    Form1.tmrUpdateTimeRemaining.Enabled = True
    Printer.EndDoc
    FlexGrid.Redraw = True
End Sub
Public Sub SetStart()
    If strCurrentEmpInfo.Number <> "" Then
        txtEmpName = strCurrentEmpInfo.Name
        txtEmpNum = strCurrentEmpInfo.Number
        lblHireDate.Caption = strCurrentEmpInfo.HireDate
        lblYearsWorked.Caption = CalcYearsWorked(strCurrentEmpInfo.Number).YearsWorked
        DTStartDate.Value = Date
        DTEndDate.Value = Date
        cmbStatus.Text = "REQUESTED"
        cmbStatus2.Text = "PAID"
    Else
        ClearAll
    End If
End Sub
Public Sub LoadEntries(Optional ShowAll As Boolean)
    If strCurrentEmpInfo.Number = "" Then Exit Sub
    Call SetStart
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    If ShowAll Then
        strSQL1 = "SELECT * " & "FROM attendb.vacations vacations_0" & " WHERE (vacations_0.idEmpNum='" & strCurrentEmpInfo.Number & "')" & IIf(chkRequested.Value = 0, "AND vacations_0.idStatus <> 'REQUESTED'", "") & IIf(chkReScheduled.Value = 0, "AND vacations_0.idStatus <> 'RESCHEDULED'", "") & IIf(chkTaken.Value = 0, "AND vacations_0.idStatus <> 'TAKEN'", "") & " Order By vacations_0.idStartDate"
    Else
        'strSQL1 = "SELECT * " & "FROM attendb.vacations vacations_0" & " WHERE (vacations_0.idStartDate>={d '" & Format$(dtAnniDate(strCurrentEmpInfo.HireDate, intPeriod).PreviousYearSub1Week, strDBDateFormat) & "'})" & "AND (vacations_0.idStartDate<={d '" & Format$(dtAnniDate(strCurrentEmpInfo.HireDate, intPeriod).CurrentYearPlus1Week, strDBDateFormat) & "'}) AND (vacations_0.idEmpNum='" & strCurrentEmpInfo.Number & "')" & IIf(chkRequested.Value = 0, "" & "AND vacations_0.idStatus <> 'REQUESTED'", "") & IIf(chkReScheduled.Value = 0, "AND vacations_0.idStatus <> 'RESCHEDULED'", "") & IIf(chkTaken.Value = 0, "AND vacations_0.idStatus <> 'TAKEN'", "") & " Order By vacations_0.idStartDate"
        strSQL1 = "SELECT * " & "FROM attendb.vacations vacations_0" & " WHERE (vacations_0.idStartDate>={d '" & Format$(DateAdd("yyyy", -intPeriod, dtVacaPeriod.StartDate), strDBDateFormat) & "'})" & "AND (vacations_0.idStartDate<={d '" & Format$(DateAdd("yyyy", -intPeriod, dtVacaPeriod.EndDate), strDBDateFormat) & "'}) AND (vacations_0.idEmpNum='" & strCurrentEmpInfo.Number & "')" & IIf(chkRequested.Value = 0, "" & "AND vacations_0.idStatus <> 'REQUESTED'", "") & IIf(chkReScheduled.Value = 0, "AND vacations_0.idStatus <> 'RESCHEDULED'", "") & IIf(chkTaken.Value = 0, "AND vacations_0.idStatus <> 'TAKEN'", "") & " Order By vacations_0.idStartDate"
    End If
    cn_Global.CursorLocation = adUseClient
    Set rs = cn_Global.Execute(strSQL1)
    If optAll Then
        lblCurrentPeriod.Caption = "All Periods"
    Else
        lblCurrentPeriod.Caption = DateAdd("yyyy", -intPeriod, dtVacaPeriod.StartDate) & " - " & DateAdd("yyyy", -intPeriod, dtVacaPeriod.EndDate)
    End If
    With rs
        If .RecordCount < 1 Then
            Grid1.Visible = False
            'Grid1.Clear
            CalcWeeksAvail
            Exit Sub
        End If
    End With
    Grid1.Redraw = False
    'Grid1.Visible = False
    Grid1.Clear
    Grid1.Rows = rs.RecordCount + 1
    With rs
        Do Until .EOF
            Grid1.CellDetails .AbsolutePosition, 1, Format$(!idStartDate, strUserDateFormat), DT_CENTER
            Grid1.CellDetails .AbsolutePosition, 2, Format$(!idEndDate, strUserDateFormat), DT_CENTER
            Grid1.CellDetails .AbsolutePosition, 3, !idHours, DT_CENTER 'DateDiffW(!idStartDate, !idEndDate) * 8, DT_CENTER
            If !idStatus = "REQUESTED" Then
                Grid1.CellDetails .AbsolutePosition, 4, !idStatus, DT_CENTER, , &H8080FF
            ElseIf !idStatus = "RESCHEDULED" Then
                Grid1.CellDetails .AbsolutePosition, 4, !idStatus, DT_CENTER, , &H8080FF
            ElseIf !idStatus = "TAKEN" Then
                Grid1.CellDetails .AbsolutePosition, 4, !idStatus, DT_CENTER, , &H80FF80
            End If
            Grid1.CellDetails .AbsolutePosition, 5, !idStatus2, DT_CENTER
            Grid1.CellDetails .AbsolutePosition, 6, !idNotes, DT_WORDBREAK
            Grid1.CellDetails .AbsolutePosition, 7, !idGUID, DT_CENTER
            Grid1.CellDetails .AbsolutePosition, 8, !idLastModified
            Grid1.CellDetails .AbsolutePosition, 9, !idLastModifiedBy
            Grid1.ColumnVisible(7) = False
            .MoveNext
        Loop
    End With
    Grid1.RowVisible(Grid1.Rows) = False
    ReSizeSGrid
    Grid1.Redraw = True
    Grid1.Visible = True
    CalcWeeksAvail
    bolOpenEmp = True
    If txtEmpNum <> Form1.txtAttenEmpNum Then
        Form1.txtAttenEmpNum.Text = strCurrentEmpInfo.Number 'intEmpNum(List1.ListIndex)
        Form1.GetEntries
        frmVacations.SetFocus
        'GetCurrentEmp (txtEmpNum.Text)
    End If
End Sub
Private Sub ReSizeSGrid()
    On Error Resume Next
    ' Grid1.Redraw = False
    Dim c As Integer, R As Integer, intCellPadding As Integer
    intCellPadding = 20
    For c = 1 To Grid1.Columns
        Grid1.AutoWidthColumn c
        Grid1.ColumnWidth(c) = Grid1.ColumnWidth(c) + intCellPadding
    Next c
    Grid1.ColumnWidth(3) = 50
    Grid1.ColumnWidth(6) = 500
    Grid1.ColumnWidth(5) = 100
    Grid1.ColumnWidth(8) = 200
    Grid1.ColumnWidth(9) = 150
    For R = 1 To Grid1.Rows
        Grid1.AutoHeightRow R
    Next R
    Grid1.HeaderHotTrack = True
    'Grid1.Redraw = True
End Sub
Public Sub CalcWeeksAvail()
    Dim rs            As New ADODB.Recordset
    Dim strSQL1       As String
    Dim intTakenWeeks As Integer
    Dim lngHoursTaken As Long
    Dim intDays       As Integer
    Dim DTStartDate   As Date, DTEndDate As Date
    intTakenWeeks = 0
    DTStartDate = dtVacaPeriod.StartDate
    DTEndDate = dtVacaPeriod.EndDate
    strSQL1 = "SELECT * " & "FROM attendb.vacations vacations_0" & " WHERE (vacations_0.idStartDate>={d '" & Format$(DTStartDate, strDBDateFormat) & "'}) AND (vacations_0.idendDate<={d '" & Format$(DTEndDate, strDBDateFormat) & "'}) AND (vacations_0.idEmpNum='" & strCurrentEmpInfo.Number & "') AND (vacations_0.idStatus='TAKEN') AND (vacations_0.idStatus2='PAID')"
    cn_Global.CursorLocation = adUseClient
    Set rs = cn_Global.Execute(strSQL1)
    With rs
        If .RecordCount < 1 Then
            lblTakenHours = 0
            If strCurrentEmpInfo.VacaHours <> 0 Then
                lblVacaHours = strCurrentEmpInfo.VacaHours
            Else
                lblVacaHours = CalcYearsWorked(strCurrentEmpInfo.Number).VacaHoursAvail
            End If
            Exit Sub
        Else
            Do Until rs.EOF
                lngHoursTaken = lngHoursTaken + !idHours
                .MoveNext
            Loop
        End If
    End With
    lblTakenHours = lngHoursTaken
    If strCurrentEmpInfo.VacaHours <> 0 Then
        lblVacaHours = strCurrentEmpInfo.VacaHours - lngHoursTaken
    Else
        lblVacaHours = CalcYearsWorked(strCurrentEmpInfo.Number).VacaHoursAvail - lngHoursTaken
    End If
End Sub
Private Sub chkTaken_Click()
    Call LoadEntries(optAll)
End Sub
Private Sub chkRequested_Click()
    Call LoadEntries(optAll)
End Sub
Private Sub chkReScheduled_Click()
    Call LoadEntries(optAll)
End Sub
Private Sub cmdAdd_Click()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_Global.CursorLocation = adUseClient
    strSQL1 = "select * from vacations"
    rs.Open strSQL1, cn_Global, adOpenUnspecified, adLockOptimistic
    With rs
        rs.AddNew
        !idStartDate = Format$(DTStartDate.Value, strDBDateFormat)
        !idEndDate = Format$(DTEndDate.Value, strDBDateFormat)
        !idHours = Int(txtHours.Text)
        !idStatus = cmbStatus.Text
        !idStatus2 = cmbStatus2.Text
        !idNotes = strNotes
        !idEmpNum = strCurrentEmpInfo.Number
        !idLastModifiedBy = strLocalUser
        rs.Update
    End With
    ClearAllButEmpInfo
    Call LoadEntries(optAll)
End Sub
Private Sub cmdAttenReports_Click()
    frmReport.LoadEmpList
    frmReport.Show
End Sub
Private Sub cmdCancel_Click()
    ClearAllButEmpInfo
End Sub
Private Sub cmdClear_Click()
    ClearAll
    ResetFilters
End Sub
Private Sub ClearAllButEmpInfo()
    DTStartDate = Date
    DTEndDate = Date
    txtHours.Text = 0
    txtNotes.Text = ""
    cmdUpdate.Visible = False
    UpdateMode = False
    cmdCancel.Visible = False
    cmdAdd.Visible = True
    cmbStatus.Text = "REQUESTED"
    cmbStatus2.Text = "PAID"
    lblLastModified.Visible = False
    'cmbStatus.Enabled = False
    Frame1.BackColor = vbButtonFace
End Sub
Private Sub ClearAllButEmpName()
    lblLastModified.Visible = False
    DTStartDate = Date
    DTEndDate = Date
    txtHours.Text = 0
    txtNotes.Text = ""
    cmdUpdate.Visible = False
    UpdateMode = False
    cmdCancel.Visible = False
    cmdAdd.Visible = True
    cmbStatus.Text = "REQUESTED"
    cmbStatus2.Text = "PAID"
    Grid1.Visible = False
    'Grid1.Clear
    lblHireDate = "0"
    lblYearsWorked = "0"
    lblTakenHours = "0"
    lblVacaHours = "0"
    'cmbStatus.Enabled = False
    Frame1.BackColor = vbButtonFace
End Sub
Private Sub ClearAll()
    lblLastModified.Visible = False
    DTStartDate = Date
    DTEndDate = Date
    txtNotes.Text = ""
    txtHours.Text = 0
    txtEmpNum.Text = ""
    txtEmpName.Text = ""
    lblHireDate.Caption = ""
    lblYearsWorked.Caption = ""
    lblVacaHours.Caption = ""
    lblTakenHours.Caption = ""
    cmdUpdate.Visible = False
    UpdateMode = False
    bolOpenEmp = False
    cmdCancel.Visible = False
    cmdAdd.Visible = True
    cmbStatus.Text = "REQUESTED"
    'cmbStatus.Enabled = False
    Frame1.BackColor = vbButtonFace
    'Form1.ClearFields
    'Grid1.Clear
    Grid1.Visible = False
    lblHireDate = "0"
    lblYearsWorked = "0"
    lblTakenHours = "0"
    lblVacaHours = "0"
    ClearEmpInfo
End Sub
Private Sub cmdExit_Click()
    bolVacationOpen = False
    Unload frmVacations
End Sub
Private Sub ResetFilters()
    optCurrentYear.Value = True
    txtXYears.Text = "2"
    chkRequested.Value = 1
    chkReScheduled.Value = 1
    chkTaken.Value = 1
End Sub
Private Sub cmdOverride_Click()
    If strCurrentEmpInfo.Number = "" Then Exit Sub
    Dim VacaHours    As String
    Dim intVacaHours As Integer
    VacaHours = InputBox("Enter number of vacation HOURS available." & vbCrLf & vbCrLf & "Current vacation available: " & IIf(strCurrentEmpInfo.VacaHours = 0, CalcYearsWorked(strCurrentEmpInfo.Number).VacaHoursAvail, strCurrentEmpInfo.VacaHours) & vbCrLf & vbCrLf & "Set this to 0 to have vacations calculated automatically.", "Vacation Override", "0")
    VacaHours = Trim$(VacaHours)
    If VacaHours = "" Then Exit Sub
    If Not IsNumeric(VacaHours) Then
        Dim blah
        blah = MsgBox(Chr(34) & VacaHours & Chr(34) & " is not a number... (>.<)", vbExclamation, "Uh...")
        Exit Sub
    End If
    intVacaHours = Int(VacaHours)
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    strSQL1 = "SELECT * From emplist Where idNumber = '" & strCurrentEmpInfo.Number & "'"
    cn_Global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_Global, adOpenKeyset, adLockOptimistic
    With rs
        If intVacaHours = !idVacaHours Then
            Exit Sub
        End If
        !idVacaHours = intVacaHours
        .Update
    End With
    strCurrentEmpInfo.VacaHours = intVacaHours
    Call LoadEntries(optAll)
End Sub
Private Sub cmdReset_Click()
    ResetFilters
End Sub
Private Sub cmdSpellChk_Click()
    txtNotes.Text = SpellMe(txtNotes.Text)
End Sub
Private Sub cmdUpdate_Click()
    UpdateEntry
End Sub
Private Sub UpdateEntry()
    On Error GoTo errs
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_Global.CursorLocation = adUseClient
    strSQL1 = "SELECT * From vacations Where idGUID = '" & strGUID & "'"
    rs.Open strSQL1, cn_Global, adOpenKeyset, adLockOptimistic
    With rs
        !idStartDate = Format$(DTStartDate.Value, strDBDateFormat)
        !idEndDate = Format$(DTEndDate.Value, strDBDateFormat)
        !idHours = Int(txtHours.Text)
        !idStatus = cmbStatus.Text
        !idStatus2 = cmbStatus2.Text
        !idNotes = strNotes
        !idEmpNum = strCurrentEmpInfo.Number
        !idLastModifiedBy = strLocalUser
        rs.Update
    End With
    ClearAllButEmpInfo
    UpdateMode = False
    Call LoadEntries(optAll)
    Exit Sub
errs:
    Dim blah
    If Err.Number = "-2147217864" Then blah = MsgBox("It looks like no changes were made." & vbCrLf & vbCrLf & "I cannot update the database with identical information.", vbExclamation, "Database Error")
    Resume Next
End Sub
Private Sub cmdVacaReports_Click()
    frmVacationReports.Show
End Sub
Private Sub DTStartDate_Change()
    DTEndDate.Value = DTStartDate.Value
End Sub
Private Sub DTStartDate_Click()
    DTEndDate.Value = DTStartDate.Value
End Sub
Private Sub Form_Load()
    bolVacationOpen = True
    intPeriod = 0
    mnuPopup.Visible = False
    cmbStatus.AddItem "", 0
    cmbStatus.AddItem "REQUESTED", 1
    cmbStatus.AddItem "RESCHEDULED", 2
    cmbStatus.AddItem "TAKEN", 3
    cmbStatus2.AddItem "", 0
    cmbStatus2.AddItem "PAID", 1
    cmbStatus2.AddItem "UNPAID", 2
    Grid1.AddColumn 1, "Start Date"
    Grid1.AddColumn 2, "End Date"
    Grid1.AddColumn 3, "Hours"
    Grid1.AddColumn 4, "Status"
    Grid1.AddColumn 5, "Paid/UnPaid"
    Grid1.AddColumn 6, "Notes"
    Grid1.AddColumn 7
    Grid1.ColumnVisible(7) = False
    Grid1.AddColumn 8, "Last Modified"
    Grid1.AddColumn 9, "Last Modified By"
    'Grid1.ColumnVisible(6) = False
    SetStart
    Call LoadEntries(optAll)
End Sub
Private Sub DeleteEntry()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_Global.CursorLocation = adUseClient
    strSQL1 = "SELECT * From Vacations Where idGUID = '" & strSelGUID & "'"
    rs.Open strSQL1, cn_Global, adOpenKeyset, adLockOptimistic
    With rs
        .Delete
        .Update
    End With
    Call LoadEntries(optAll)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolVacationOpen = False
End Sub
Private Sub Grid1_ColumnClick(ByVal lCol As Long)
    Dim sTag As String
    With Grid1.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = Grid1.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderDescending
        End If
        Grid1.ColumnTag(lCol) = sTag
        Select Case Grid1.ColumnKey(lCol)
            Case "file", "col8"
                ' sort by icon:
                .SortType(1) = CCLSortIcon
            Case "date"
                ' sort by date:
                .SortType(1) = CCLSortDate
            Case Else
                ' sort by text:
                .SortType(1) = CCLSortString
        End Select
    End With
    Screen.MousePointer = vbHourglass
    Grid1.Sort
    Screen.MousePointer = vbDefault
End Sub
Private Sub Grid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    On Error GoTo errs
    With Grid1
        DTStartDate.Value = .CellText(.SelectedRow, 1)
        DTEndDate.Value = .CellText(.SelectedRow, 2)
        txtHours.Text = .CellText(.SelectedRow, 3)
        cmbStatus.Text = .CellText(.SelectedRow, 4)
        cmbStatus2.Text = .CellText(.SelectedRow, 5)
        txtNotes.Text = .CellText(.SelectedRow, 6)
        strGUID = .CellText(.SelectedRow, 7)
        lblLastModified.Caption = "Last Modified:" & vbCrLf & .CellText(.SelectedRow, 8)
    End With
    cmdAdd.Visible = False
    cmdUpdate.Visible = True
    cmbStatus.Enabled = True
    cmdCancel.Visible = True
    lblLastModified.Visible = True
    Frame1.BackColor = &HC0FFC0
    UpdateMode = True
    Exit Sub
errs:
    If Err.Number = 9 Then
        ClearAllButEmpInfo
    End If
End Sub
Private Sub Grid1_KeyPress(KeyAscii As Integer)
    On Error GoTo errs
    If KeyAscii = 100 Then
        Call LoadEntries(optAll)
    End If
    Exit Sub
errs:
End Sub
Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next 'GoTo errs
    If Button = 2 Then 'Button 2 is "Right Click"
        'If Grid1.SelectedRow <> 0 Then
        strSelGUID = Grid1.CellText(Grid1.SelectedRow, 7)
        Me.PopupMenu mnuPopup
        'End If
    End If
    Exit Sub
errs:
End Sub

Private Sub lblVacaHours_Change()
    On Error Resume Next
    cmdOverride.Left = lblVacaHours.Left + lblVacaHours.Width + 100
    If lblVacaHours < 1 Then
        lblVacaHours.ForeColor = vbRed
    Else
        lblVacaHours.ForeColor = vbBlack
    End If
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        GetCurrentEmp (intEmpNum(List1.ListIndex))
        Call LoadEntries(optAll)
        frmVacations.SetFocus
        List1.Visible = False
        List1.Clear
    End If
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errs
    GetCurrentEmp (intEmpNum(List1.ListIndex))
    Call LoadEntries(optAll)
    frmVacations.SetFocus
    List1.Visible = False
    List1.Clear
    Exit Sub
errs:
    If Err.Number = 9 Then Exit Sub
End Sub
Private Sub List1_GotFocus()
    tmrLiveSearch.Enabled = False
    intSearchWaitTicks = 0
End Sub
Private Sub mnuDelete_Click()
    On Error GoTo errs
    Dim blah
    With Grid1
        blah = MsgBox("Start Date: " & .CellText(.SelectedRow, 1) & vbCrLf & "End Date: " & .CellText(.SelectedRow, 2) & vbCrLf & "Current Status: " & .CellText(.SelectedRow, 4) & vbCrLf & "GUID: " & strSelGUID & vbCrLf & vbCrLf & "Are you sure you want to delete this entry?", vbYesNo + vbExclamation, "Delete Entry")
    End With
    If blah = vbNo Then
        ClearAllButEmpInfo
        Call LoadEntries(optAll)
    ElseIf blah = vbYes Then
        ClearAllButEmpInfo
        DeleteEntry
    End If
    Exit Sub
errs:
    blah = MsgBox(Err.Description, vbExclamation + vbOKOnly, "Error")
End Sub
Private Sub mnuPrint_Click()
    frmPrinters.Show 1
    If bolCancelPrint = True Then
        bolCancelPrint = False
        Exit Sub
    End If
    strReportTitle = "Emp #: " & strCurrentEmpInfo.Number & "   Name: " & strCurrentEmpInfo.Name
    frmPBar.Show
    DoEvents
    'Grid1.RemoveColumn 7
    PrintSGrid Grid1, strReportTitle
    'Grid1.AddColumn 7
    Call LoadEntries(optAll)
End Sub
Private Sub optAll_Click()
    Call LoadEntries(optAll)
End Sub
Private Sub optCurrentYear_Click()
    intPeriod = 0
    Call LoadEntries(optAll)
End Sub
Private Sub optPrevYear_Click()
    intPeriod = 1
    Call LoadEntries(optAll)
End Sub
Private Sub optXYears_Click()
    If txtXYears.Text <> 0 Then
        intPeriod = txtXYears.Text
        Call LoadEntries(optAll)
    End If
End Sub
Private Sub LiveSearch(ByVal strSearchString As String)
    Dim rs               As New ADODB.Recordset
    Dim strSQL1          As String
    Dim strUsedJobNums() As String
    Dim Row              As Integer
    On Error Resume Next
    Row = 0
    List1.Clear
    Erase intEmpNum
    cn_Global.CursorLocation = adUseClient
    strSQL1 = "SELECT idName,idNumber From emplist Where idName Like '%" & strSearchString & "%' Order By EmpList.idName"
    Set rs = cn_Global.Execute(strSQL1)
    ReDim intEmpNum(rs.RecordCount)
    ReDim strUsedJobNums(rs.RecordCount + 1)
    Do Until rs.EOF
        With rs
            List1.AddItem !idName, Row
            intEmpNum(Row) = !idNumber
            Row = Row + 1
            rs.MoveNext
        End With
    Loop
    If rs.RecordCount >= 1 Then
        List1.Visible = True
    ElseIf rs.RecordCount <= 0 Then
        List1.Visible = False
    End If
End Sub
Private Sub tmrButtonEnabler_Timer()
    If bolOpenEmp And IsNumeric(txtHours.Text) Then
        If Int(txtHours.Text) > 0 Then
            cmdAdd.Enabled = True
        Else
            cmdAdd.Enabled = False
        End If
    Else
        cmdAdd.Enabled = False
    End If
    If cmbStatus.Text = "" Or cmbStatus2.Text = "" Then cmdAdd.Enabled = False
    If UpdateMode Then
        If cmbStatus.Text = "" Or cmbStatus2.Text = "" Then
            cmdUpdate.Enabled = False
        Else
            cmdUpdate.Enabled = True
        End If
    End If
End Sub
Private Sub tmrLiveSearch_Timer()
    On Error Resume Next
    intSearchWaitTicks = intSearchWaitTicks + 1
    If intSearchWaitTicks >= intSearchWait Then
        LiveSearch (txtEmpName.Text)
        intSearchWaitTicks = 0
        tmrLiveSearch.Enabled = False
    End If
End Sub
Private Sub txtEmpName_Change()
    ClearAllButEmpName
End Sub
Private Sub txtEmpName_GotFocus()
    With txtEmpName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
    If List1.ListCount <= 0 Then List1.Visible = False
    If Len(txtEmpName.Text) >= 2 Then
        tmrLiveSearch.Enabled = True
        intSearchWaitTicks = 0
    Else
        List1.Visible = False
    End If
    If KeyCode = vbKeyDown Then
        List1.SetFocus
        List1.Selected(0) = True
    End If
End Sub
Private Sub txtEmpNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And CheckForEmp(txtEmpNum.Text) Then
        GetCurrentEmp (txtEmpNum.Text)
        bolOpenEmp = True
        Call LoadEntries(optAll)
        frmVacations.SetFocus
    ElseIf KeyAscii = 13 And Not CheckForEmp(txtEmpNum.Text) Then
        Dim blah
        blah = MsgBox("Employee not found.", vbOKOnly, "Error")
        ClearAll
    End If
End Sub
Private Sub txtEmpNum_LostFocus()
    If GetTabState And txtEmpNum.Text <> "" Then
        If CheckForEmp(txtEmpNum.Text) Then
            GetCurrentEmp (txtEmpNum.Text)
            bolOpenEmp = True
            Call LoadEntries(optAll)
            frmVacations.SetFocus
        Else
            Dim blah
            blah = MsgBox("Employee not found.", vbOKOnly, "Error")
            Form1.ClearFields
            ClearAll
        End If
    ElseIf txtEmpNum.Text <> "" Then
    End If
End Sub
Private Sub txtNotes_Change()
    strNotes = Replace$(txtNotes.Text, vbCrLf, "")
End Sub
Private Sub txtNotes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And UpdateMode = True Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 And UpdateMode = False And cmdAdd.Enabled = True Then
        KeyAscii = 0
        txtNotes.Text = Replace$(txtNotes.Text, vbCrLf, "")
    End If
End Sub
Private Sub txtXYears_Change()
    On Error Resume Next
    If optXYears Then
        intPeriod = txtXYears.Text
        Call LoadEntries(optAll)
    End If
End Sub
