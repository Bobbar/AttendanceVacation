VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance"
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10725
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrButtonEnable 
      Interval        =   30
      Left            =   10080
      Top             =   4440
   End
   Begin VB.Timer tmrUpdateTimeRemaining 
      Interval        =   150
      Left            =   10080
      Top             =   3960
   End
   Begin VB.Timer tmrLiveSearch 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   10080
      Top             =   4980
   End
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
      Left            =   2520
      TabIndex        =   27
      Top             =   900
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Frame Frame3 
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   10575
      Begin RichTextLib.RichTextBox txtNotes 
         Height          =   855
         Left            =   1920
         TabIndex        =   32
         Top             =   1680
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         MaxLength       =   200
         TextRTF         =   $"Form1.frx":0CCA
      End
      Begin VB.CommandButton cmdSpellCheck 
         Caption         =   "Spell Check"
         Height          =   240
         Left            =   7755
         TabIndex        =   31
         Top             =   2520
         Width           =   990
      End
      Begin VB.TextBox txtHoursLate 
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
         Left            =   8280
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cmbTimeOffType 
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
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   840
         Width           =   2775
      End
      Begin VB.ComboBox cmbExcused 
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
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Add Entry"
         Enabled         =   0   'False
         Height          =   480
         Left            =   4620
         TabIndex        =   16
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   360
         Left            =   9480
         TabIndex        =   15
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Entry"
         Height          =   480
         Left            =   3840
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   480
         Left            =   5220
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker DTEntryDate 
         Height          =   345
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
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
         Format          =   198639617
         CurrentDate     =   40484
      End
      Begin MSComCtl2.DTPicker DTEntryDateTo 
         Height          =   345
         Left            =   1560
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   198639617
         CurrentDate     =   40484
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   0
         TabIndex        =   45
         Top             =   60
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excused?"
         Height          =   195
         Left            =   3375
         TabIndex        =   26
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excuse Type"
         Height          =   195
         Left            =   5400
         TabIndex        =   25
         Top             =   600
         Width           =   2715
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hours"
         Height          =   195
         Left            =   8280
         TabIndex        =   24
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   195
         Left            =   1560
         TabIndex        =   21
         Top             =   960
         Width           =   570
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Entries"
      Height          =   4515
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   10575
      Begin vbAcceleratorSGrid6.vbalGrid GridAtten 
         Height          =   4155
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   7329
         RowMode         =   -1  'True
         GridLines       =   -1  'True
         BackgroundPicture=   "Form1.frx":0D47
         BackgroundPictureHeight=   128
         BackgroundPictureWidth=   128
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
         HeaderFlat      =   -1  'True
         DisableIcons    =   -1  'True
         DrawFocusRectangle=   0   'False
         HotTrack        =   -1  'True
         SelectionAlphaBlend=   -1  'True
         SelectionOutline=   -1  'True
      End
      Begin VB.Label lblGrid 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   10425
      End
      Begin VB.Shape Shape1 
         Height          =   4155
         Left            =   120
         Top             =   240
         Width           =   10335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   10575
      Begin VB.Frame Frame5 
         Caption         =   "Fuctions"
         Height          =   735
         Left            =   2700
         TabIndex        =   40
         Top             =   1680
         Width           =   4635
         Begin VB.CommandButton cmdAttenReports 
            Caption         =   "Attendance Reports"
            Height          =   360
            Left            =   2880
            TabIndex        =   43
            Top             =   240
            Width           =   1650
         End
         Begin VB.CommandButton cmdVacaReports 
            Caption         =   "Vacation Reports"
            Height          =   360
            Left            =   1200
            TabIndex        =   42
            Top             =   240
            Width           =   1590
         End
         Begin VB.CommandButton cmdVacations 
            Caption         =   "Vacations"
            Height          =   360
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.ComboBox cmbLocation2 
         Enabled         =   0   'False
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
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         Height          =   360
         Left            =   9480
         TabIndex        =   28
         Top             =   2040
         Width           =   990
      End
      Begin VB.TextBox txtAttenEmpNum 
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
         Height          =   405
         Left            =   300
         TabIndex        =   0
         Text            =   "EmpNum"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtAttenEmpName 
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
         Height          =   405
         Left            =   2400
         TabIndex        =   1
         Text            =   "EmpName"
         Top             =   480
         Width           =   4155
      End
      Begin VB.ComboBox cmbLocation 
         Enabled         =   0   'False
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
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox chkIsActive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Caption         =   "Inactive"
         DisabledPicture =   "Form1.frx":3099
         DownPicture     =   "Form1.frx":5699
         Enabled         =   0   'False
         Height          =   375
         Left            =   975
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Alarms"
         Height          =   1635
         Left            =   7440
         TabIndex        =   33
         Top             =   300
         Width           =   2955
         Begin VB.Label lblAcked 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ack"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblPartialUn 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partial Un -"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   945
            TabIndex        =   37
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblFullUn 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Full Un -"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1170
            TabIndex        =   36
            Top             =   480
            Width           =   750
         End
         Begin VB.Label lblFullEx 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Full Ex -"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1155
            TabIndex        =   35
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label lblPartialEx 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partial Ex -"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   945
            TabIndex        =   34
            Top             =   960
            Width           =   1065
         End
      End
      Begin VB.Label lblHireDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%HIRE DATE%"
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
         Left            =   5400
         TabIndex        =   46
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   375
         TabIndex        =   38
         Top             =   1890
         Width           =   465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location 2"
         Height          =   195
         Left            =   3060
         TabIndex        =   30
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hire Date"
         Height          =   195
         Left            =   5460
         TabIndex        =   5
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         Height          =   195
         Left            =   3900
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label lblAppVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%APP VERSION%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   180
      TabIndex        =   48
      Top             =   10800
      Width           =   1290
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by Bobby Lovell"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   9180
      TabIndex        =   47
      Top             =   10800
      Width           =   1470
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuRepo 
         Caption         =   "Reports"
         Begin VB.Menu mnuVacaReports 
            Caption         =   "Vacations"
         End
         Begin VB.Menu mnuReport 
            Caption         =   "Attendance"
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Filter Entries"
      End
      Begin VB.Menu mnuAddHours 
         Caption         =   "Sum Hours"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sndPlaySound _
                Lib "WINMM.DLL" _
                Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                       ByVal uFlags As Long) As Long
Private intSearchWaitTicks As Integer
Private UpdateMode         As Boolean
Private strNotes           As String

Private Sub DrawSGridForPrint()
    Dim i     As Integer
    Dim Hours As Single
    On Error Resume Next
    AttenVals.FullExcused = 0
    AttenVals.FullUnExcused = 0
    AttenVals.PartialExcused = 0
    AttenVals.PartialUnExcused = 0
    For i = 1 To GridAtten.Rows ' - 1
        Hours = Hours + GridAtten.CellText(i, 5)
        CountExcusesAll GridAtten.CellText(i, 3), GridAtten.CellText(i, 4)
    Next
    GridAtten.Rows = GridAtten.Rows + 1
    If Hours <= 0 Then
        GridAtten.CellDetails GridAtten.Rows, 5, "| Total Hrs: None |", DT_CENTER
    Else
        GridAtten.CellDetails GridAtten.Rows, 5, "| Total Hrs: " & Hours & " |", DT_CENTER
    End If
    If intRowsAdded = 0 Then
        strReportEntryCount = " Total Entries: 0 "
    Else
        strReportEntryCount = " Total Entries: " & intRowsAdded & " "
    End If
    strReportInfo = " " & AttenVals.PartialUnExcused & " Partial Unexcused | " & AttenVals.PartialExcused & " Partial Excused | " & AttenVals.FullUnExcused & " Full Unexcused | " & AttenVals.FullExcused & " Full Excused | " & intOtherRowsAdded & " Other "
End Sub
Public Sub CheckForDLLS()
    Dim blah
    Dim strLastRunVersion As String, strCurrentVersion As String
    strLastRunVersion = GetSetting(App.EXEName, "Version", "LastRun", "0.0.0")
    strCurrentVersion = App.Major & App.Minor & App.Revision
    If strLastRunVersion <> strCurrentVersion Then
        blah = Shell(App.Path & "\InstallDLLs.bat", vbMinimizedNoFocus)
        SaveSetting App.EXEName, "Version", "LastRun", strCurrentVersion
        'they be there
    Else
        'blah = msgbox("
        'SaveSetting App.EXEName, "Version", "LastRun", strCurrentVersion
    End If
End Sub
Public Sub CheckForAlarm(EmpNum As String)
    Dim SoundName As String
    Dim Found     As Boolean
    Dim NewAlarm  As Boolean
    intPartialUnExcused = AttenVals.PartialUnExcused
    intFullUnExcused = AttenVals.FullUnExcused
    intFullExcused = AttenVals.FullExcused
    intPartialExcused = AttenVals.PartialExcused
    Dim bolSetPartialUn As Integer, bolSetFullUn As Integer, bolSetFullEx As Integer, bolSetPartialEx As Integer
    NewAlarm = False
    Found = InStr(1, vbNullChar & Join(strConfirmedAlarms(), vbNullChar) & vbNullChar, vbNullChar & txtAttenEmpNum.Text & vbNullChar) > 0
    bolSetPartialUn = GetSetting(App.EXEName, EmpNum, "PartialUn", 0)
    bolSetFullUn = GetSetting(App.EXEName, EmpNum, "FullUn", 0)
    bolSetFullEx = GetSetting(App.EXEName, EmpNum, "FullEx", 0)
    bolSetPartialEx = GetSetting(App.EXEName, EmpNum, "PartialEx", 0)
    If Found = True Then
        Exit Sub
    Else
        If intPartialUnExcused >= intPartialUnExcusedAllowed Then bolPartialUnExcusedExceeded = True
        If intFullUnExcused >= intFullUnExcusedAllowed Then bolFullUnExcusedExceeded = True
        If intFullExcused >= intFullExcusedAllowed Then bolFullExcusedExceeded = True
        If intPartialExcused >= intPartialExcusedAllowed Then bolPartialExcusedExceeded = True
        If bolPartialUnExcusedExceeded = True Then
            lblPartialUn.ForeColor = &HFF&
            lblPartialUn.Caption = "Partial Un - " & intPartialUnExcused & " of " & intPartialUnExcusedAllowed
        Else
            lblPartialUn.ForeColor = vbBlack
            lblPartialUn.Caption = "Partial Un - " & intPartialUnExcused & " of " & intPartialUnExcusedAllowed
        End If
        If bolFullUnExcusedExceeded = True Then
            lblFullUn.ForeColor = &HFF&
            lblFullUn.Caption = "Full Un - " & intFullUnExcused & " of " & intFullUnExcusedAllowed
        Else
            lblFullUn.ForeColor = vbBlack
            lblFullUn.Caption = "Full Un - " & intFullUnExcused & " of " & intFullUnExcusedAllowed
        End If
        If bolFullExcusedExceeded = True Then
            lblFullEx.ForeColor = &HFF&
            lblFullEx.Caption = "Full Ex - " & intFullExcused & " of " & intFullExcusedAllowed
        Else
            lblFullEx.ForeColor = vbBlack
            lblFullEx.Caption = "Full Ex - " & intFullExcused & " of " & intFullExcusedAllowed
        End If
        If bolPartialExcusedExceeded = True Then
            lblPartialEx.ForeColor = &HFF&
            lblPartialEx.Caption = "Partial Ex - " & intPartialExcused & " of " & intPartialExcusedAllowed
        Else
            lblPartialEx.ForeColor = vbBlack
            lblPartialEx.Caption = "Partial Ex - " & intPartialExcused & " of " & intPartialExcusedAllowed
        End If
        Call SetAckLabel(EmpNum)
        If bolPartialUnExcusedExceeded = True And bolSetPartialUn = 0 Then NewAlarm = True
        If bolFullUnExcusedExceeded = True And bolSetFullUn = 0 Then NewAlarm = True
        If bolFullExcusedExceeded = True And bolSetFullEx = 0 Then NewAlarm = True
        If bolPartialExcusedExceeded = True And bolSetPartialEx = 0 Then NewAlarm = True
        If bolPartialUnExcusedExceeded = False And bolFullUnExcusedExceeded = False And bolFullExcusedExceeded = False And bolPartialExcusedExceeded = False Then
            strAlarmTitleString = Form1.txtAttenEmpName.Text
        Else
            strAlarmTitleString = Form1.txtAttenEmpName.Text & " is over the limit!"
        End If
        If NewAlarm = True And bolAlarmOKed = False Then
            SoundName = App.Path & "\Sounds\Siren.wav"
            sndPlaySound SoundName$, &H1 'play alarm sound
            bolAlarmsCalled = False
            List1.Visible = False
            frmAlarm.Show vbModal
        End If
    End If
End Sub
Public Sub SetAckLabel(EmpNum As String)
    Dim bolSetPartialUn As Integer, bolSetFullUn As Integer, bolSetFullEx As Integer, bolSetPartialEx As Integer
    bolSetPartialUn = GetSetting(App.EXEName, EmpNum, "PartialUn", 0)
    bolSetFullUn = GetSetting(App.EXEName, EmpNum, "FullUn", 0)
    bolSetFullEx = GetSetting(App.EXEName, EmpNum, "FullEx", 0)
    bolSetPartialEx = GetSetting(App.EXEName, EmpNum, "PartialEx", 0)
    Dim i As Integer
    For i = 1 To lblAcked.UBound
        Unload lblAcked(i)
    Next
    If bolPartialUnExcusedExceeded = True And bolSetPartialUn = 1 Then
        Load lblAcked(lblAcked.UBound + 1)
        lblAcked(lblAcked.UBound).Top = lblPartialUn.Top
        lblAcked(lblAcked.UBound).Visible = True
    End If
    If bolFullUnExcusedExceeded = True And bolSetFullUn = 1 Then
        Load lblAcked(lblAcked.UBound + 1)
        lblAcked(lblAcked.UBound).Top = lblFullUn.Top
        lblAcked(lblAcked.UBound).Visible = True
    End If
    If bolFullExcusedExceeded = True And bolSetFullEx = 1 Then
        Load lblAcked(lblAcked.UBound + 1)
        lblAcked(lblAcked.UBound).Top = lblFullEx.Top
        lblAcked(lblAcked.UBound).Visible = True
    End If
    If bolPartialExcusedExceeded = True And bolSetPartialEx = 1 Then
        Load lblAcked(lblAcked.UBound + 1)
        lblAcked(lblAcked.UBound).Top = lblPartialEx.Top
        lblAcked(lblAcked.UBound).Visible = True
    End If
    Frame4.Refresh
End Sub
Public Sub SetEmpLocation(Location As String, Location2 As String)
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    On Error Resume Next
    If txtAttenEmpNum.Text = "" Or NewEmp = True Then Exit Sub
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From EmpList Where idNumber = '" & txtAttenEmpNum.Text & "'"
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    With rs
        !idLocation1 = Location
        !idLocation2 = Location2
        rs.Update
    End With
    rs.Close
    cn.Close
    GetEmpInfo
End Sub
Public Sub SetEmpActive(IsActive As String)
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    If txtAttenEmpNum.Text = "" Or chkIsActive.Enabled = False Or NewEmp = True Then Exit Sub
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From EmpList Where idNumber = '" & txtAttenEmpNum.Text & "'"
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    With rs
        !idIsActive = IsActive
    End With
    rs.Update
    rs.Close
    cn.Close
    GetEmpInfo
End Sub
Private Sub LiveSearch(ByVal strSearchString As String)
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim Row     As Integer
    On Error Resume Next
    Row = 0
    List1.Clear
    Erase intEmpNum
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT idName,idNumber From emplist Where idName Like '%" & strSearchString & "%' Order By EmpList.idName"
    rs.Open strSQL1, cn, adOpenKeyset
    ReDim intEmpNum(rs.RecordCount)
    
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
    rs.Close
    cn.Close
End Sub
Public Sub DateRangeReport()
    GridAtten.Redraw = False
    GridAtten.Visible = False
    Dim rs               As New ADODB.Recordset
    Dim cn               As New ADODB.Connection
    Dim strSQL1          As String
    Dim strUsedJobNums() As String
    Dim dtTicketDate     As Date
    Dim Row              As Integer, Line               As Integer
    Dim Found            As Boolean
    Dim EmpNum           As String
    Dim Qrys(4)          As String
    Dim Qry              As String
    Dim BuiltQrys        As Integer
    Dim sFntUnder        As New StdFont
    sFntUnder.Underline = True
    sFntUnder.name = "Tahoma"
    Dim sFntNormal As New StdFont
    sFntNormal.Underline = False
    sFntNormal.name = "Tahoma"
    On Error Resume Next
    intUnExcusedRowsAdded = 0
    intExcusedRowsAdded = 0
    intOtherRowsAdded = 0
    SelectedFilters = 0
    If frmFilters.chkExcused.Value = 1 Then SelectedFilters = SelectedFilters + 1
    If frmFilters.chkExcusedPartial.Value = 1 Then SelectedFilters = SelectedFilters + 1
    If frmFilters.chkUnexcused.Value = 1 Then SelectedFilters = SelectedFilters + 1
    If frmFilters.chkUnExcusedPartial.Value = 1 Then SelectedFilters = SelectedFilters + 1
    If frmFilters.chkOther.Value = 1 Then SelectedFilters = SelectedFilters + 1
    EmpNum = txtAttenEmpNum.Text
    Qry = "SELECT * FROM attendb.attenentries attenentries_0 WHERE"
    'UnExcusedPartial
    Qrys(0) = "(attenentries_0.idAttenTimeOffType='Left Early') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='UNEXCUSED') " & "OR (attenentries_0.idAttenTimeOffType='Left & Came Back') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='UNEXCUSED') " & "OR (attenentries_0.idAttenTimeOffType='Late for DAY') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='UNEXCUSED') OR " & "(attenentries_0.idAttenTimeOffType='Late from LUNCH') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='UNEXCUSED')"
    'ExcusedPartial
    Qrys(1) = "(attenentries_0.idAttenTimeOffType='Left Early') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='EXCUSED')" & " OR (attenentries_0.idAttenTimeOffType='Left & Came Back') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='EXCUSED')" & " OR (attenentries_0.idAttenTimeOffType='Late for DAY') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='EXCUSED') OR" & " (attenentries_0.idAttenTimeOffType='Late from LUNCH') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='EXCUSED')"
    'UnexcusedFull
    Qrys(2) = "(attenentries_0.idAttenTimeOffType='Called Off') AND (attenentries_0.idAttenExcused='UNEXCUSED') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "')" & " OR (attenentries_0.idAttenTimeOffType='No Call, No Show') AND (attenentries_0.idAttenExcused='UNEXCUSED') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "')" & " OR (attenentries_0.idAttenTimeOffType='Requested Day Off') AND (attenentries_0.idAttenExcused='UNEXCUSED') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "')"
    'ExcusedFull
    Qrys(3) = "(attenentries_0.idAttenTimeOffType='Called Off') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='EXCUSED')" & " OR (attenentries_0.idAttenTimeOffType='No Call, No Show') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='EXCUSED')" & " OR (attenentries_0.idAttenTimeOffType='Requested Day Off') AND (attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='EXCUSED')"
    'Other
    Qrys(4) = "(attenentries_0.idAttenEmpNum='" & EmpNum & "') AND (attenentries_0.idAttenExcused='OTHER')"
    If frmFilters.chkUnExcusedPartial.Value = 1 And frmFilters.chkAll.Value = 0 Then
        BuiltQrys = BuiltQrys + 1
        If BuiltQrys < SelectedFilters Then
            Qry = Qry + Qrys(0) + " OR "
        ElseIf BuiltQrys = SelectedFilters Then
            Qry = Qry + Qrys(0) + " Order By attenentries_0.idAttenEntryDate Desc"
        End If
    End If
    If frmFilters.chkExcusedPartial.Value = 1 And frmFilters.chkAll.Value = 0 Then
        BuiltQrys = BuiltQrys + 1
        If BuiltQrys < SelectedFilters Then
            Qry = Qry + Qrys(1) + " OR "
        ElseIf BuiltQrys = SelectedFilters Then
            Qry = Qry + Qrys(1) + " Order By attenentries_0.idAttenEntryDate Desc"
        End If
    End If
    If frmFilters.chkUnexcused.Value = 1 And frmFilters.chkAll.Value = 0 Then
        BuiltQrys = BuiltQrys + 1
        If BuiltQrys < SelectedFilters Then
            Qry = Qry + Qrys(2) + " OR "
        ElseIf BuiltQrys = SelectedFilters Then
            Qry = Qry + Qrys(2) + " Order By attenentries_0.idAttenEntryDate Desc"
        End If
    End If
    If frmFilters.chkExcused.Value = 1 And frmFilters.chkAll.Value = 0 Then
        BuiltQrys = BuiltQrys + 1
        If BuiltQrys < SelectedFilters Then
            Qry = Qry + Qrys(3) + " OR "
        ElseIf BuiltQrys = SelectedFilters Then
            Qry = Qry + Qrys(3) + " Order By attenentries_0.idAttenEntryDate Desc"
        End If
    End If
    If frmFilters.chkOther.Value = 1 And frmFilters.chkAll.Value = 0 Then
        BuiltQrys = BuiltQrys + 1
        If BuiltQrys < SelectedFilters Then
            Qry = Qry + Qrys(4) + " OR "
        ElseIf BuiltQrys = SelectedFilters Then
            Qry = Qry + Qrys(4) + " Order By attenentries_0.idAttenEntryDate Desc"
        End If
    End If
    If frmFilters.chkAll.Value = 1 Then Qry = "SELECT * From AttenEntries where idAttenEmpNum like '" & txtAttenEmpNum.Text & "' Order By attenentries.idAttenEntryDate Desc"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = Qry
    rs.Open strSQL1, cn, adOpenKeyset
    Row = 0
    Line = 1
    GridAtten.Clear
    GridAtten.Rows = rs.RecordCount ' + 1
    DTStartDate = Format$(frmFilters.DTStart.Value, "MM/DD/YYYY")
    DTEndDate = Format$(frmFilters.DTEnd.Value, "MM/DD/YYYY")
    ReDim strUsedJobNums(rs.RecordCount + 1)
    Do Until rs.EOF
        With rs
            dtTicketDate = Format$(!idAttenEntryDate, strUserDateFormat)
            If frmFilters.chkAllDate.Value = 0 Then
                If dtTicketDate < DTStartDate Or dtTicketDate > DTEndDate Then  'Date range filter
                    'If dtTicketDate > DTStartDate And dtTicketDate < DTEndDate Then
                    ReDim Preserve strUsedJobNums(UBound(strUsedJobNums) + 1)
                    strUsedJobNums(Row) = !idGUID
                    Row = Row + 1
                Else
                    'let the ticket be displayed
                End If
            Else
            End If
            Found = InStr(1, vbNullChar & Join$(strUsedJobNums(), vbNullChar) & vbNullChar, vbNullChar & !idGUID & vbNullChar) > 0
            If Found = False Then
                strUsedJobNums(Row) = !idGUID
                GridAtten.CellDetails Line, 7, !idGUID
                GridAtten.CellDetails Line, 1, Format$(!idAttenEntryDate, strUserDateFormat), DT_CENTER
                If Format$(!idAttenEntryDateTo, strUserDateFormat) <> Format$("2000-01-01", strUserDateFormat) Then
                    GridAtten.CellDetails Line, 2, Format$(!idAttenEntryDateTo, strUserDateFormat), DT_CENTER
                Else
                    GridAtten.CellDetails Line, 2, ""
                End If
                If !idAttenExcused = "EXCUSED" Then
                    GridAtten.CellDetails Line, 3, !idAttenExcused, DT_CENTER, , &H80FF80
                    intExcusedRowsAdded = intExcusedRowsAdded + 1
                ElseIf !idAttenExcused = "UNEXCUSED" Then
                    GridAtten.CellDetails Line, 3, !idAttenExcused, DT_CENTER, , &H8080FF
                    intUnExcusedRowsAdded = intUnExcusedRowsAdded + 1
                ElseIf !idAttenExcused = "OTHER" Then
                    GridAtten.CellDetails Line, 3, !idAttenExcused, DT_CENTER, , &HFFFF80
                    intOtherRowsAdded = intOtherRowsAdded + 1
                End If
                GridAtten.CellDetails Line, 4, !idAttenTimeOffType, DT_CENTER
                GridAtten.CellDetails Line, 5, !idAttenPartialDay, DT_CENTER
                GridAtten.CellDetails Line, 6, !idAttenNotes
                Line = Line + 1
                Row = Row + 1
            ElseIf Found = True Then
            End If
            'SkipHeader:
            rs.MoveNext
        End With
    Loop
    rs.Close
    cn.Close
    Erase strUsedJobNums
    GridAtten.Rows = Line - 1
    intRowsAdded = Line - 1
    ReSizeSGrid
    GridAtten.Redraw = True
    GridAtten.Visible = True
End Sub
Private Sub ClearAllButEmpNum()
    bolOpenEmp = False
    strNotes = ""
    txtAttenEmpNum.Enabled = True
    txtAttenEmpName.Enabled = True
    txtAttenEmpName.Visible = True
    DTEntryDate.Enabled = True
    DTEntryDateTo.Enabled = False
    cmbExcused.Enabled = True
    cmbTimeOffType.Enabled = True
    txtHoursLate.Enabled = True
    txtNotes.Enabled = True
    lblHireDate.Caption = Date
    cmbExcused.ListIndex = 0
    cmbTimeOffType.ListIndex = 0
    cmbLocation.ListIndex = 0
    cmbLocation2.ListIndex = 0
    txtHoursLate.Text = ""
    txtNotes.Text = ""
    GridAtten.Visible = False
    GridAtten.Redraw = False
    cmdUpdate.Visible = False
    cmdSubmit.Visible = True
    NewEmp = False
    UpdateMode = False
    chkIsActive.Value = 0
    chkIsActive.Caption = "Inactive"
    chkIsActive.Enabled = False
    bolPartialUnExcusedExceeded = False
    bolFullUnExcusedExceeded = False
    bolFullExcusedExceeded = False
    bolPartialExcusedExceeded = False
    lblPartialUn.ForeColor = vbBlack
    lblPartialEx.ForeColor = vbBlack
    lblFullUn.ForeColor = vbBlack
    lblFullEx.ForeColor = vbBlack
    lblPartialUn.Caption = "Partial Un -"
    lblPartialEx.Caption = "Partial Ex -"
    lblFullUn.Caption = "Full Un -"
    lblFullEx.Caption = "Full Ex -"
    bolAlarmOKed = False
    Dim i As Integer
    For i = 1 To lblAcked.UBound
        Unload lblAcked(i)
    Next
End Sub
Public Sub ClearFields()
    strNotes = ""
    txtAttenEmpNum.Enabled = True
    txtAttenEmpName.Enabled = True
    txtAttenEmpName.Visible = True
    DTEntryDate.Enabled = True
    DTEntryDate.Value = Date
    DTEntryDateTo.Value = Date
    DTEntryDateTo.Enabled = False
    cmbExcused.Enabled = True
    cmbTimeOffType.Enabled = True
    txtHoursLate.Enabled = True
    txtNotes.Enabled = True
    cmdCancel.Visible = False
    cmdSubmit.Caption = "Add Entry"
    txtAttenEmpNum.Text = ""
    txtAttenEmpName.Text = ""
    lblHireDate.Caption = Date
    cmbExcused.ListIndex = 0
    cmbTimeOffType.ListIndex = 0
    txtHoursLate.Text = ""
    txtNotes.Text = ""
    GridAtten.Visible = False
    cmdUpdate.Visible = False
    UpdateMode = False
    cmdSubmit.Visible = True
    NewEmp = False
    UpdateMode = False
    List1.Visible = False
    List1.Clear
    cmbLocation.ListIndex = 0
    cmbLocation2.ListIndex = 0
    chkIsActive.Value = 0
    chkIsActive.Caption = "Inactive"
    chkIsActive.Enabled = False
    bolIsDateRange = False
    DTEntryDateTo.Value = Date
    DTEntryDateTo.Enabled = False
    lblPartialUn.ForeColor = vbBlack
    lblPartialEx.ForeColor = vbBlack
    lblFullUn.ForeColor = vbBlack
    lblFullEx.ForeColor = vbBlack
    lblPartialUn.Caption = "Partial Un -"
    lblPartialEx.Caption = "Partial Ex -"
    lblFullUn.Caption = "Full Un -"
    lblFullEx.Caption = "Full Ex -"
    bolAlarmOKed = False
    ClearEmpInfo
    Dim i As Integer
    For i = 1 To lblAcked.UBound
        Unload lblAcked(i)
    Next
    Frame4.Refresh
End Sub
Private Sub ClearBottomFields()
    cmbExcused.ListIndex = 0
    cmbTimeOffType.ListIndex = 0
    txtHoursLate.Text = ""
    txtNotes.Text = ""
    strNotes = ""
    txtAttenEmpName.Visible = True
    bolIsDateRange = False
    DTEntryDateTo.Value = Date
    DTEntryDateTo.Enabled = False
    DTEntryDate.Value = Date
End Sub
Public Sub GetEntries()
    On Error Resume Next
    Dim rs        As New ADODB.Recordset
    Dim cn        As New ADODB.Connection
    Dim strSQL1   As String
    Dim sFntUnder As New StdFont
    sFntUnder.Underline = True
    sFntUnder.name = "Tahoma"
    Dim sFntNormal As New StdFont
    sFntNormal.Underline = False
    sFntNormal.name = "Tahoma"
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    AttenVals.FullExcused = 0
    AttenVals.FullUnExcused = 0
    AttenVals.PartialExcused = 0
    AttenVals.PartialUnExcused = 0
    Call FillHeader(txtAttenEmpNum.Text)
    strSQL1 = "SELECT * From AttenEntries Where idAttenEmpNum = '" & txtAttenEmpNum.Text & "' Order By attenentries.idAttenEntryDate Desc"
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset
    If rs.RecordCount < 1 Then ' No entries
        GridAtten.Visible = False
        GridAtten.Redraw = False
        GridAtten.Clear
        lblGrid.Visible = True
        rs.Close
        cn.Close
        If bolVacationOpen Then
            Call frmVacations.SetStart
            Call frmVacations.LoadEntries(frmVacations.optAll)
        End If
        Exit Sub
    ElseIf rs.RecordCount >= 1 Then
        txtAttenEmpNum.TabIndex = ControlFocus(0)
        txtAttenEmpName.TabIndex = ControlFocus(1)
        DTEntryDate.TabIndex = ControlFocus(3)
        cmbExcused.TabIndex = ControlFocus(4)
        cmbTimeOffType.TabIndex = ControlFocus(5)
        txtHoursLate.TabIndex = ControlFocus(6)
        txtNotes.TabIndex = ControlFocus(7)
        GridAtten.Redraw = False
        txtNotes.Enabled = True
        txtHoursLate.Enabled = True
        cmbTimeOffType.Enabled = True
        cmbExcused.Enabled = True
        DTEntryDate.Enabled = True
        DTEntryDate.SetFocus
        lblGrid.Visible = False
    End If
    GridAtten.Clear
    GridAtten.Rows = rs.RecordCount '+ 1
    intUnExcusedRowsAdded = 0
    intExcusedRowsAdded = 0
    intOtherRowsAdded = 0
    intRowsAdded = 0
    Do Until rs.EOF
        With rs
            GridAtten.CellDetails .AbsolutePosition, 7, !idGUID
            GridAtten.ColumnVisible(7) = False
            GridAtten.CellDetails .AbsolutePosition, 1, Format$(!idAttenEntryDate, strUserDateFormat), DT_CENTER
            If Format$(!idAttenEntryDateTo, strUserDateFormat) <> Format$("2000-01-01", strUserDateFormat) Then
                GridAtten.CellDetails .AbsolutePosition, 2, Format$(!idAttenEntryDateTo, strUserDateFormat), DT_CENTER
            Else
                GridAtten.CellDetails .AbsolutePosition, 2, ""
            End If
            If !idAttenExcused = "EXCUSED" Then
                GridAtten.CellDetails .AbsolutePosition, 3, !idAttenExcused, DT_CENTER, , &H80FF80
                intExcusedRowsAdded = intExcusedRowsAdded + 1
            ElseIf !idAttenExcused = "UNEXCUSED" Then
                GridAtten.CellDetails .AbsolutePosition, 3, !idAttenExcused, DT_CENTER, , &H8080FF
                intUnExcusedRowsAdded = intUnExcusedRowsAdded + 1
            ElseIf !idAttenExcused = "OTHER" Then
                GridAtten.CellDetails .AbsolutePosition, 3, !idAttenExcused, DT_CENTER, , &HFFFF80
                intOtherRowsAdded = intOtherRowsAdded + 1
            End If
            intRowsAdded = intRowsAdded + 1
            GridAtten.CellDetails .AbsolutePosition, 4, !idAttenTimeOffType, DT_CENTER
            GridAtten.CellDetails .AbsolutePosition, 5, !idAttenPartialDay, DT_CENTER
            GridAtten.CellDetails .AbsolutePosition, 6, !idAttenNotes, DT_WORDBREAK
            CountExcusesInYear !idAttenExcused, !idAttenTimeOffType, !idAttenEntryDate
            rs.MoveNext
        End With
    Loop
    rs.Close
    ReSizeSGrid
    GridAtten.Redraw = True
    GridAtten.Visible = True
    DTEntryDate.Value = Date
    CheckForAlarm txtAttenEmpNum.Text
    If bolVacationOpen Then
        If frmVacations.txtEmpNum <> txtAttenEmpNum Then
            Call frmVacations.SetStart
            Call frmVacations.LoadEntries(frmVacations.optAll)
        End If
    End If
End Sub
Private Sub ReSizeSGridForPrint()
    GridAtten.Redraw = False
    Dim i As Integer, intCellPadding As Integer
    intCellPadding = 10
    For i = 1 To GridAtten.Columns
        GridAtten.AutoWidthColumn i
        GridAtten.ColumnWidth(i) = GridAtten.ColumnWidth(i) + intCellPadding
    Next i
    If GridAtten.ColumnWidth(2) < GridAtten.ColumnWidth(1) Then GridAtten.ColumnWidth(2) = GridAtten.ColumnWidth(1)
    GridAtten.Redraw = True
End Sub
Private Sub ReSizeSGrid()
    Dim R As Integer, intCellPadding As Integer
    intCellPadding = 20
    GridAtten.AutoWidthColumn 4
    GridAtten.ColumnWidth(4) = GridAtten.ColumnWidth(4) + intCellPadding
    GridAtten.ColumnWidth(1) = 112
    GridAtten.ColumnWidth(2) = 112
    GridAtten.ColumnWidth(3) = 104
    GridAtten.ColumnWidth(6) = 500
    GridAtten.ColumnWidth(2) = GridAtten.ColumnWidth(1)
    GridAtten.ColumnWidth(5) = 70
    For R = 1 To GridAtten.Rows '
        GridAtten.AutoHeightRow R
    Next R
    GridAtten.HeaderHotTrack = True
End Sub
Public Sub FillHeader(EmpNum As String)
    txtAttenEmpNum.Text = EmpNum
    strCurrentEmpInfo.Number = EmpNum
    txtAttenEmpName.Text = ReturnEmpInfo(EmpNum).name
    strCurrentEmpInfo.name = ReturnEmpInfo(EmpNum).name
    lblHireDate.Caption = ReturnEmpInfo(EmpNum).HireDate
    strCurrentEmpInfo.HireDate = ReturnEmpInfo(EmpNum).HireDate
    strCurrentEmpInfo.VacaWeeks = ReturnEmpInfo(EmpNum).VacaWeeks
    
    If ReturnEmpInfo(EmpNum).IsActive = "TRUE" Then
        chkIsActive.Enabled = True
        chkIsActive.Value = 1
        chkIsActive.Caption = "Active"
    ElseIf ReturnEmpInfo(EmpNum).IsActive = "FALSE" Then
        chkIsActive.Enabled = True
        chkIsActive.Value = 0
        chkIsActive.Caption = "Inactive"
    End If
    
    cmbLocation2.Text = ReturnEmpInfo(EmpNum).Location2
    strCurrentEmpInfo.Location2 = ReturnEmpInfo(EmpNum).Location2
    cmbLocation.Text = ReturnEmpInfo(EmpNum).Location1
    strCurrentEmpInfo.Location1 = ReturnEmpInfo(EmpNum).Location1
    cmbLocation.Enabled = False
    cmbLocation2.Enabled = False
    bolOpenEmp = True
End Sub
Private Sub chkIsActive_Click()
    If bolOpenEmp Then
        If chkIsActive.Value = 0 Then
            chkIsActive.Caption = "Inactive"
            SetEmpActive "FALSE"
        Else
            chkIsActive.Caption = "Active"
            SetEmpActive "TRUE"
        End If
    End If
End Sub
Private Sub cmbLocation_Click()
    If bolOpenEmp = True Then
        SetEmpLocation cmbLocation.Text, cmbLocation2.Text
        cmbLocation.Enabled = False
    End If
End Sub
Private Sub cmbLocation2_Click()
    If bolOpenEmp = True Then
        SetEmpLocation cmbLocation.Text, cmbLocation2.Text
        cmbLocation2.Enabled = False
    End If
End Sub
Private Sub cmdAttenReports_Click()
    frmReport.LoadEmpList
    frmReport.Show
End Sub
Private Sub cmdClear_Click()
    'ClearFields
    ClearBottomFields
End Sub
Private Sub cmdClear2_Click()
    ClearFields
    'ClearVacation
End Sub
Public Sub AddEmpToDB(name As Variant, _
                      Num As String, _
                      HireDate As String, _
                      Location1 As String, _
                      Location2 As String, _
                      IsActive As String)
    Dim rs         As New ADODB.Recordset
    Dim cn         As New ADODB.Connection
    Dim strSQL2    As String
    Dim FormatDate As String
    On Error GoTo errs
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    FormatDate = Format$(HireDate, strDBDateFormat)
    strSQL2 = "INSERT INTO attendb.emplist (idName,idLocation1,idLocation2,idNumber,idHireDate,idIsActive) VALUES ('" & name & "','" & Location1 & "','" & Location2 & "','" & Num & "','" & FormatDate & "','" & IsActive & "')"
    rs.Open strSQL2, cn, adOpenKeyset, adLockOptimistic
    Exit Sub
errs:
    MsgBox Err.Description
End Sub
Private Sub cmdSubmit_Click()
    On Error GoTo errs
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "select * from attenentries"
    rs.Open strSQL1, cn, adOpenUnspecified, adLockOptimistic
    With rs
        rs.AddNew
        !idAttenEmpNum = Trim$(txtAttenEmpNum.Text)
        If Trim$(txtHoursLate.Text) = "" Then
            !idAttenPartialDay = Round("0.00", 2)
        Else
            !idAttenPartialDay = Round(txtHoursLate.Text, 2)
        End If
        !idAttenExcused = cmbExcused.Text
        !idAttenTimeOffType = cmbTimeOffType.Text
        If bolIsDateRange = True Then
            !idAttenEntryDate = Format$(DTEntryDate.Value, strDBDateFormat)
            !idAttenEntryDateTo = Format$(DTEntryDateTo.Value, strDBDateFormat)
        Else
            !idAttenEntryDate = Format$(DTEntryDate.Value, strDBDateFormat)
            !idAttenEntryDateTo = "2000-01-01"
        End If
        !idAttenNotes = strNotes
        rs.Update
    End With
    rs.Close
    cn.Close
    GetEntries
    ClearBottomFields
    Exit Sub
errs:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next
End Sub
Private Sub cmdUpdate_Click()
    On Error GoTo errs
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From AttenEntries Where idGUID Like '" & SelGUID & "'"
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    With rs
        If bolIsDateRange = True Then
            !idAttenEntryDate = Format$(DTEntryDate.Value, strDBDateFormat)
            !idAttenEntryDateTo = Format$(DTEntryDateTo.Value, strDBDateFormat)
        Else
            !idAttenEntryDate = Format$(DTEntryDate.Value, strDBDateFormat)
            !idAttenEntryDateTo = "2000-01-01"
        End If
        !idAttenExcused = cmbExcused.Text
        !idAttenTimeOffType = cmbTimeOffType.Text
        If Trim$(txtHoursLate.Text) = "" Then
            !idAttenPartialDay = Round("0.00", 2)
        Else
            !idAttenPartialDay = Round(txtHoursLate.Text, 2)
        End If
        !idAttenNotes = strNotes
        rs.Update
    End With
    rs.Close
    cn.Close
    txtAttenEmpNum.Enabled = True
    txtAttenEmpName.Enabled = True
    GetEntries
    cmdUpdate.Visible = False
    cmdCancel.Visible = False
    cmdSubmit.Visible = True
    ClearBottomFields
    txtAttenEmpNum.SetFocus
    UpdateMode = False
    Exit Sub
errs:
    Dim blah
    If Err.Number = "-2147217864" Then blah = MsgBox("It looks like no changes were made." & vbCrLf & vbCrLf & "I cannot update the database with identical information.", vbExclamation, "Database Error")
    Resume Next
End Sub
Private Sub cmdCancel_Click()
    txtAttenEmpNum.Enabled = True
    txtAttenEmpName.Enabled = True
    txtAttenEmpName.Visible = True
    DTEntryDate.Enabled = True
    cmbExcused.Enabled = True
    cmbTimeOffType.Enabled = True
    txtHoursLate.Enabled = True
    txtNotes.Enabled = True
    cmbExcused.ListIndex = 0
    cmbTimeOffType.ListIndex = 0
    txtHoursLate.Text = ""
    txtNotes.Text = ""
    cmdUpdate.Visible = False
    cmdCancel.Visible = False
    UpdateMode = False
    cmdSubmit.Visible = True
    NewEmp = False
    UpdateMode = False
End Sub
Private Sub cmdSpellCheck_Click()
    txtNotes.Text = SpellMe(txtNotes.Text)
End Sub
Private Sub cmdVacaReports_Click()
    frmVacationReports.Show
End Sub
Private Sub cmdVacations_Click()
    frmVacations.Show
End Sub
Private Sub Form_Initialize()
    CheckForDLLS
End Sub
Private Sub Form_Load()

FindMySQLDriver

lblAppVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
    intQryIndex = 0
    ReDim lngTimeRemainingArray(1)
    strServerAddress = "10.0.1.232"
    Dim strFullAccessUser As String, strFullAccessPass As String, strROAccessUser As String, strROAccessPass As String
    strFullAccessUser = "AttenUser"
    strFullAccessPass = "y2zq3T21Ejia"
    strROAccessUser = "AttenUserRO"
    strROAccessPass = "8f0DYyS7y056"
 
    Select Case UCase$(Environ$("USERNAME"))
        Case "SORRELLJ"
            strUsername = strFullAccessUser
            strPassword = strFullAccessPass
        Case "HOWARDS"
            strUsername = strFullAccessUser
            strPassword = strFullAccessPass
        Case "SCHRINERT"
            strUsername = strFullAccessUser
            strPassword = strFullAccessPass
        Case "LOVELLB"
            strUsername = strFullAccessUser
            strPassword = strFullAccessPass
        Case Else
            strUsername = strROAccessUser
            strPassword = strROAccessPass
    End Select

    cmbExcused.AddItem "", 0
    cmbExcused.AddItem "EXCUSED", 1
    cmbExcused.AddItem "UNEXCUSED", 2
    cmbExcused.AddItem "OTHER", 3
    FillCombos
    bolIsDateRange = False
    mnuPopup.Visible = False
    lblHireDate.Caption = Date
    DTEntryDate.Value = Date
    DTEntryDateTo.Value = Date
    ClearFields
    ControlFocus(0) = txtAttenEmpNum.TabIndex
    ControlFocus(1) = txtAttenEmpName.TabIndex
    ControlFocus(3) = DTEntryDate.TabIndex
    ControlFocus(4) = cmbExcused.TabIndex
    ControlFocus(5) = cmbTimeOffType.TabIndex
    ControlFocus(6) = txtHoursLate.TabIndex
    ControlFocus(7) = txtNotes.TabIndex
    frmFilters.DTEnd.Value = Date
    frmFilters.DTStart.Value = Date
    bolAlarmOKed = False
    bolCancelPrint = False
    Flashes = 0
    InitializeMe 'Word Spell checker
    ReDim Preserve strConfirmedAlarms(1)
    dtFiscalYearEnd = "5/31/" & DateTime.Year(Now)
    GetEmpInfo
    SetupGrid
End Sub
Private Sub SetupGrid()
    GridAtten.AddColumn 1, "Date"
    GridAtten.AddColumn 2, "Date To"
    GridAtten.AddColumn 3, "Excuse"
    GridAtten.AddColumn 4, "Type"
    GridAtten.AddColumn 5, "Hours"
    GridAtten.AddColumn 6, "Notes"
    GridAtten.AddColumn 7
    GridAtten.ColumnVisible(7) = False
    GridAtten.Gridlines = True
End Sub
Public Sub FillCombos()
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From comboitems Order By idTimeOffType"
    rs.Open strSQL1, cn, adOpenKeyset
    cmbTimeOffType.Clear
    cmbLocation.Clear
    cmbLocation2.Clear
    frmAddNewEmp.cmbLocation1.Clear
    frmAddNewEmp.cmbLocation2.Clear
    cmbTimeOffType.AddItem "", 0
    cmbLocation.AddItem "", 0
    cmbLocation2.AddItem "", 0
    frmAddNewEmp.cmbLocation1.AddItem "", 0
    frmAddNewEmp.cmbLocation2.AddItem "", 0
    Do Until rs.EOF
        With rs
            If !idLocation1 <> "" Then
                cmbLocation.AddItem !idLocation1, .AbsolutePosition
                frmAddNewEmp.cmbLocation1.AddItem !idLocation1, .AbsolutePosition
            End If
            If !idLocation2 <> "" Then
                cmbLocation2.AddItem !idLocation2, .AbsolutePosition
                frmAddNewEmp.cmbLocation2.AddItem !idLocation2, .AbsolutePosition
            End If
            If !idTimeOffType <> "" Then cmbTimeOffType.AddItem !idTimeOffType, .AbsolutePosition
            .MoveNext
        End With
    Loop
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If KillMe = True Then
        moApp.Quit False
    End If
    Set moApp = Nothing
    Call EndProgram
    'End
End Sub
Sub EndProgram()
    Dim tmpForm As Form
    For Each tmpForm In Forms
        If tmpForm.name <> "Form1" Then
            Unload tmpForm
            Set tmpForm = Nothing
        End If
    Next
End Sub
Private Sub Frame1_Click()
    List1.Visible = False
    List1.Clear
End Sub
Private Sub Frame1_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    If X > DTEntryDateTo.Left And X < DTEntryDateTo.Left + DTEntryDateTo.Width And Y > DTEntryDateTo.Top And Y < DTEntryDateTo.Top + DTEntryDateTo.Height Then
        DTEntryDateTo.Enabled = True
        bolIsDateRange = True
    End If
    If X > cmbLocation.Left And X < cmbLocation.Left + cmbLocation.Width And Y > cmbLocation.Top And Y < cmbLocation.Top + cmbLocation.Height Then
        cmbLocation.Enabled = True
    End If
    If X > cmbLocation2.Left And X < cmbLocation2.Left + cmbLocation2.Width And Y > cmbLocation2.Top And Y < cmbLocation2.Top + cmbLocation2.Height Then
        cmbLocation2.Enabled = True
    End If
End Sub
Private Sub Frame2_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    If Button = 2 Then 'Button 2 is "Right Click"
        Me.PopupMenu mnuPopup
    End If
End Sub
Private Sub Frame3_Click()
    List1.Visible = False
    List1.Clear
End Sub
Private Sub Frame3_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    If X > DTEntryDateTo.Left And X < DTEntryDateTo.Left + DTEntryDateTo.Width And Y > DTEntryDateTo.Top And Y < DTEntryDateTo.Top + DTEntryDateTo.Height Then
        DTEntryDateTo.Enabled = True
        bolIsDateRange = True
    End If
    If X > cmbLocation.Left And X < cmbLocation.Left + cmbLocation.Width And Y > cmbLocation.Top And Y < cmbLocation.Top + cmbLocation.Height Then
        cmbLocation.Enabled = True
    End If
End Sub
Private Sub Frame4_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    If txtAttenEmpNum.Text = "" Then Exit Sub
    bolAlarmsCalled = True
    frmAlarm.Show vbModal
End Sub
Private Sub GridAtten_ColumnClick(ByVal lCol As Long)
    Dim sTag As String
    With GridAtten.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = GridAtten.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderDescending
        End If
        GridAtten.ColumnTag(lCol) = sTag
        Select Case GridAtten.ColumnKey(lCol)
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
    GridAtten.Sort
    Screen.MousePointer = vbDefault
End Sub
Private Sub GridAtten_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    cmdUpdate.Visible = True
    UpdateMode = True
    cmdSubmit.Visible = False
    cmdCancel.Visible = True
    txtAttenEmpNum.Enabled = False
    txtAttenEmpName.Enabled = False
    DTEntryDate.Value = Date
    DTEntryDateTo.Value = Date
    With GridAtten
        If Len(.CellText(.SelectedRow, 2)) > 1 Then
            DTEntryDate.Value = .CellText(.SelectedRow, 1)
            DTEntryDateTo.Value = .CellText(.SelectedRow, 2)
            bolIsDateRange = True
            DTEntryDateTo.Enabled = True
        Else
            DTEntryDate.Value = .CellText(.SelectedRow, 1)
            bolIsDateRange = False
            DTEntryDateTo.Enabled = False
        End If
        cmbExcused.Text = .CellText(.SelectedRow, 3)
        cmbTimeOffType.Text = .CellText(.SelectedRow, 4)
        txtHoursLate.Text = .CellText(.SelectedRow, 5)
        txtNotes.Text = .CellText(.SelectedRow, 6)
        SelGUID = .CellText(.SelectedRow, 7)
    End With
End Sub
Private Sub GridAtten_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then 'Key is "Del"
        If GridAtten.SelectedRow <> 0 Then
            SelGUID = GridAtten.CellText(GridAtten.SelectedRow, 7)
            SelDate = GridAtten.CellText(GridAtten.SelectedRow, 1)
            SelExcuse = GridAtten.CellText(GridAtten.SelectedRow, 3)
            SelType = GridAtten.CellText(GridAtten.SelectedRow, 4)
            SelHours = GridAtten.CellText(GridAtten.SelectedRow, 5)
        End If
        DeleteEntry
    End If
End Sub
Private Sub GridAtten_MouseUp(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    If Button = 2 Then 'Button 2 is "Right Click"
        If GridAtten.SelectedRow <> 0 Then
            SelGUID = GridAtten.CellText(GridAtten.SelectedRow, 7)
            SelDate = GridAtten.CellText(GridAtten.SelectedRow, 1)
            SelExcuse = GridAtten.CellText(GridAtten.SelectedRow, 3)
            SelType = GridAtten.CellText(GridAtten.SelectedRow, 4)
            SelHours = GridAtten.CellText(GridAtten.SelectedRow, 5)
        End If
        Me.PopupMenu mnuPopup
    End If
End Sub
Private Sub Label13_Click()

    StartTimer
    Dim TotAttenEnt As Long
    Dim blah
    GetDataBaseStats
    Dim rs      As New ADODB.Recordset
    Dim rs2     As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String, strSQL2 As String
    ReDim AttenStats(0)
    strSQL2 = "SELECT COUNT(*) FROM attendb.attenentries attenentries_0 where idAttenTimeOffType = 'Called Off'"
    strSQL1 = "SELECT * FROM attendb.comboitems comboitems_0"
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenForwardOnly, adLockReadOnly
    With rs
        Do Until .EOF
            AttenStats(UBound(AttenStats)).ExTypeName = !idTimeOffType
            strSQL2 = "SELECT COUNT(*) FROM attendb.attenentries attenentries_0 where idAttenTimeOffType = '" & AttenStats(UBound(AttenStats)).ExTypeName & "'"
            rs2.Open strSQL2, cn, adOpenForwardOnly, adLockReadOnly
            AttenStats(UBound(AttenStats)).ExTypeCount = rs2.Fields(0)
            rs2.Close
            ReDim Preserve AttenStats(UBound(AttenStats) + 1)
            .MoveNext
        Loop
    End With
    rs.Close
    cn.Close
    TotAttenEnt = DataBaseStats.TotalAttenEntries
    Dim i As Integer
    For i = 0 To UBound(AttenStats) - 1
        AttenStats(i).ExTypePct = Round((AttenStats(i).ExTypeCount / TotAttenEnt) * 100, 2)
    Next i
    Dim strBreakdown As String
    strBreakdown = vbNullString
    For i = 0 To UBound(AttenStats) - 1
        strBreakdown = strBreakdown + AttenStats(i).ExTypeName & ": " & AttenStats(i).ExTypeCount & " (" & AttenStats(i).ExTypePct & "%)" & vbCrLf
    Next i
    ReDim strChartData(UBound(AttenStats) - 1, 1)
   'ReDim strChartData(1, UBound(AttenStats) - 1)
    Dim c As Integer
  
 
     For c = 0 To UBound(AttenStats) - 1
         strChartData(c, 0) = AttenStats(c).ExTypeName & "(" & AttenStats(c).ExTypeCount & ")"
         
      
     Next c
     For c = 0 To UBound(AttenStats) - 1
         strChartData(c, 1) = AttenStats(c).ExTypeCount
         

     Next c
     
    MySort strChartData
  
    
    
    blah = MsgBox("---DBStats---" & vbCrLf & vbCrLf & "Emps: " & DataBaseStats.TotalEmployees & vbCrLf & "Vaca Entries: " _
    & DataBaseStats.TotalVacaEntries & vbCrLf & "Atten Entries: " & DataBaseStats.TotalAttenEntries & vbCrLf & vbCrLf & _
    "---Breakdown (# and % of Tot.)---" & vbCrLf & vbCrLf & strBreakdown & vbCrLf & "------" & vbCrLf & vbCrLf & "Query Time: " _
    & StopTimer & "ms" & vbCrLf & vbCrLf & "[View stats on a chart?]", vbOKCancel, "DB Stats")
     
     
    If blah = vbOK Then
      frmChart.Show
'        strBreakdown = vbNullString
'        For i = 0 To UBound(AttenStats) - 1
'            strBreakdown = strBreakdown + AttenStats(i).ExTypeName & vbTab & AttenStats(i).ExTypeCount & vbCrLf
'        Next i
'        Clipboard.Clear
'        Clipboard.SetText strBreakdown
    Else
    End If
End Sub
Private Sub MySort(ByRef pvarArray As Variant)
    Dim i               As Long
    Dim c               As Integer
    Dim v               As Integer
    Dim lngHighValIndex As Long
    Dim varSwap()       As Variant
    Dim lngMax          As Long
    ReDim varSwap(UBound(pvarArray, 2))
    lngMax = UBound(pvarArray, 1)
    For c = 0 To lngMax
        lngHighValIndex = lngMax - c
        For v = 0 To UBound(varSwap)
            varSwap(v) = pvarArray(lngMax - c, v)
        Next v
        For i = 0 To lngMax - c
            If pvarArray(i, 1) > pvarArray(lngHighValIndex, 1) Then lngHighValIndex = i
        Next
        For v = 0 To UBound(varSwap)
            pvarArray(lngMax - c, v) = pvarArray(lngHighValIndex, v)
            pvarArray(lngHighValIndex, v) = varSwap(v)
        Next v
    Next c
End Sub
Private Sub lblAcked_Click(index As Integer)
    If txtAttenEmpNum.Text = "" Then Exit Sub
    bolAlarmsCalled = True
    frmAlarm.Show vbModal
End Sub
Private Sub lblFullEx_Click()
    If txtAttenEmpNum.Text = "" Then Exit Sub
    bolAlarmsCalled = True
    frmAlarm.Show vbModal
End Sub
Private Sub lblFullUn_Click()
    If txtAttenEmpNum.Text = "" Then Exit Sub
    bolAlarmsCalled = True
    frmAlarm.Show vbModal
End Sub
Private Sub lblPartialEx_Click()
    If txtAttenEmpNum.Text = "" Then Exit Sub
    bolAlarmsCalled = True
    frmAlarm.Show vbModal
End Sub
Private Sub lblPartialUn_Click()
    If txtAttenEmpNum.Text = "" Then Exit Sub
    bolAlarmsCalled = True
    frmAlarm.Show vbModal
End Sub
Private Sub List1_GotFocus()
    tmrLiveSearch.Enabled = False
    intSearchWaitTicks = 0
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtAttenEmpNum.Text = intEmpNum(List1.ListIndex)
        GetEntries
        List1.Visible = False
        List1.Clear
    End If
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errs
    txtAttenEmpNum.Text = intEmpNum(List1.ListIndex)
    List1.Visible = False
    List1.Clear
    DoEvents
    GetEntries
    Exit Sub
errs:
    If Err.Number = 9 Then Exit Sub
End Sub
'Private Sub mnuAddHours_Click()
'
'    Dim start_row, stop_row, i, Rows As Integer
'
'    Dim Hours As Single
'
'    on error Resume Next
'
'    If MSHFlexGridAtten.Row > MSHFlexGridAtten.RowSel Then
'        start_row = MSHFlexGridAtten.RowSel
'        stop_row = MSHFlexGridAtten.Row
'    Else
'        start_row = MSHFlexGridAtten.Row
'        stop_row = MSHFlexGridAtten.RowSel
'
'    End If
'
'    Rows = stop_row - start_row + 1
'
'    For i = start_row To stop_row
'        Hours = Hours + MSHFlexGridAtten.TextMatrix(i, 5)
'    Next
'    MsgBox Hours & " hours in " & Rows & " entries.", , "Total Hours"
'
'End Sub
Private Sub mnuDelete_Click()
    DeleteEntry
End Sub
Private Sub DeleteEntry()
    Dim rs      As New ADODB.Recordset
    Dim cn      As New ADODB.Connection
    Dim strSQL1 As String
    Dim blah
    On Error GoTo errs
    If UpdateMode Then Call cmdCancel_Click
    If txtAttenEmpNum.Text = "" Or SelGUID = "" Or SelDate = "" Or SelExcuse = "" Or SelType = "" Or SelHours = "" Then
        MsgBox ("Invalid row selected, please try again.")
        Exit Sub
    Else
    End If
    blah = MsgBox("Emp #: " & txtAttenEmpNum.Text & vbCrLf & "Date: " & SelDate & vbCrLf & "Excuse: " & SelExcuse & vbCrLf & "Type: " & SelType & vbCrLf & "Hours: " & SelHours & vbCrLf & "GUID: " & SelGUID & vbCrLf & vbCrLf & "Are you sure you want to delete this entry?", vbExclamation + vbYesNo, "Delete Entry")
    If blah = vbNo Then
        Exit Sub
    ElseIf blah = vbYes Then
    End If
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * From  AttenEntries Where idGUID = '" & SelGUID & "'"
    rs.Open strSQL1, cn, adOpenKeyset, adLockOptimistic
    With rs
        .Delete
    End With
    rs.Close
    cn.Close
    SelDate = vbNullString
    SelExcuse = vbNullString
    SelType = vbNullString
    SelHours = vbNullString
    SelGUID = vbNullString
    MsgBox ("Single entry deleted successfully.")
    GetEntries
    Exit Sub
errs:
    blah = MsgBox(Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "No rows were selected for deletion. Please click on a valid row and try again.", vbOKOnly + vbCritical, "Oops!")
    Err.Clear
End Sub
Private Sub mnuFilter_Click()
    frmFilters.Show
End Sub
Private Sub mnuPrint_Click()
    On Error Resume Next
    frmPrinters.Show 1
    If bolCancelPrint = True Then
        bolCancelPrint = False
        Exit Sub
    End If
    strReportType = "SINGLE"
    strReportName = txtAttenEmpName.Text
    strReportNum = txtAttenEmpNum.Text
    'GridAtten.RemoveColumn 7
    DrawSGridForPrint
    ReSizeSGridForPrint
    PrintSGrid GridAtten, "Emp #: " & txtAttenEmpNum.Text & "   Name: " & txtAttenEmpName.Text, IIf(frmFilters.chkAllDate.Value = 0, "Entries between " & frmFilters.DTStart.Value & " and " & frmFilters.DTEnd.Value, "")
    'GridAtten.AddColumn 7
    GetEntries
    Unload frmFilters
    
End Sub
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
    intColumns = 6
    bolLongLine = False
    On Error Resume Next
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
        .name = FlexGrid.Font.name
        .Size = 9
    End With
    Printer.Print ""
    Printer.DrawStyle = vbDash
    Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY), vbBlack
    Printer.DrawStyle = vbSolid
    Printer.Print ""
    PrevY = Printer.CurrentY
    Dim xBoxEnd As Single, lngCenterXStartPos As Long
    Printer.Font.Size = 7
    lngCenterXStartPos = (xmax / 2) - (Printer.TextWidth(strReportInfo) / 2)
    xBoxEnd = lngCenterXStartPos + Printer.TextWidth(strReportInfo)
    Printer.Line (lngCenterXStartPos, PrevY)-(xBoxEnd, Printer.CurrentY + (Printer.TextHeight(strReportInfo) * 3)), &H80000016, BF
    Printer.Font.Bold = True
    Printer.CurrentX = (xmax / 2) - (Printer.TextWidth("Attendance Stats") / 2)
    Printer.CurrentY = PrevY
    Printer.Print "Attendance Stats"
    Printer.Font.Bold = False
    Printer.CurrentX = lngCenterXStartPos
    Printer.Print strReportInfo
    Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(strReportEntryCount) / 2)
    Printer.Print strReportEntryCount
    Printer.Line (lngCenterXStartPos, PrevY)-(xBoxEnd, Printer.CurrentY), vbBlack, B
    Printer.Print ""
    Printer.Font.Size = 9
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
        'TwipPix = .ColumnWidth(1) * Screen.TwipsPerPixelX
        'XFirstColumn = xmin + TwipPix * GAP
        XFirstColumn = xmin '+ TwipPix * GAP
        X = xmin + GAP
        lngYTopOfGrid = Printer.CurrentY
        Printer.CurrentY = Printer.CurrentY + GAP
        If FlexGrid.Header = True Then
            For c = 1 To .Columns - 1
                Printer.CurrentX = X
                TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                PrevY = Printer.CurrentY
                If c = .Columns - 1 Then
                    lngStartY = Printer.CurrentY - GAP + 5
                    lngStartX = Printer.CurrentX - GAP + 5
                    lngEndX = xmax
                    lngEndY = Printer.CurrentY + Printer.TextHeight(.ColumnHeader(c)) + GAP
                    Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), &H80000016, BF
                Else
                    lngStartY = Printer.CurrentY - GAP + 5
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
        For R = 1 To .Rows
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
                        Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor(R, 3), BF
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
                    If c = 3 Then
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
Private Sub mnuReport_Click()
    frmReport.LoadEmpList
    frmReport.Show
End Sub
Private Sub mnuVacaReports_Click()
    frmVacationReports.Show
End Sub
Private Sub tmrButtonEnable_Timer()
    If NewEmp = False And UpdateMode = False And txtAttenEmpNum.Text <> "" And txtAttenEmpName.Text <> "" And DTEntryDate.Value <> "" And cmbExcused.Text <> "" And cmbTimeOffType.Text <> "" Then
        cmdSubmit.Enabled = True
    ElseIf NewEmp = False And UpdateMode = False Then
        cmdSubmit.Enabled = False
    End If
    If UpdateMode = True And NewEmp = False And txtAttenEmpNum.Text <> "" And txtAttenEmpName.Text <> "" And DTEntryDate.Value <> "" And cmbExcused.Text <> "" And cmbTimeOffType.Text <> "" Then
        cmdUpdate.Enabled = True
    ElseIf UpdateMode = True And NewEmp = False Then
        cmdUpdate.Enabled = False
    End If
    'cmdVacations.Enabled = bolOpenEmp
End Sub
Private Sub tmrLiveSearch_Timer()
    On Error Resume Next
    intSearchWaitTicks = intSearchWaitTicks + 1
    If bolOpenEmp Then
        tmrLiveSearch.Enabled = False
        intSearchWaitTicks = 0
    End If
    If intSearchWaitTicks >= intSearchWait Then
        LiveSearch (txtAttenEmpName.Text)
        intSearchWaitTicks = 0
        tmrLiveSearch.Enabled = False
    End If
End Sub
Private Sub tmrUpdateTimeRemaining_Timer()
    frmPBar.lblQryTime.Caption = strTimeRemaining
End Sub
Private Sub txtAttenEmpName_Change()
    ClearAllButEmpNum
End Sub
Private Sub txtAttenEmpName_GotFocus()
    With txtAttenEmpName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtAttenEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If List1.ListCount <= 0 Then List1.Visible = False
    If Len(txtAttenEmpName.Text) >= 2 Then
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
Private Sub txtAttenEmpNum_Change()
    ClearAllButEmpNum
End Sub
Private Sub txtAttenEmpNum_GotFocus()
    With txtAttenEmpNum
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtAttenEmpNum_KeyPress(KeyAscii As Integer)
    Dim blah
    If KeyAscii = 13 Then
        If CheckForEmp(txtAttenEmpNum.Text) Then
            GetEntries
        Else
            blah = MsgBox("Employee not found. Add new?", vbOKCancel + vbCritical, "Error")
            If blah = vbCancel Then
                ClearFields
                Exit Sub
            End If
            NewEmp = True
            frmAddNewEmp.lblEmpNum.Caption = txtAttenEmpNum.Text
            frmAddNewEmp.Show vbModal
        End If
    End If
End Sub
Private Sub txtAttenEmpNum_LostFocus()
    Dim blah
    If GetTabState And txtAttenEmpNum.Text <> "" Then
        frmAlarm.Hide
        If CheckForEmp(txtAttenEmpNum.Text) Then
            GetEntries
        Else
            blah = MsgBox("Employee not found. Add new?", vbOKCancel + vbCritical, "Error")
            If blah = vbCancel Then
                ClearFields
                Exit Sub
            End If
            NewEmp = True
            frmAddNewEmp.lblEmpNum.Caption = txtAttenEmpNum.Text
            frmAddNewEmp.Show vbModal
        End If
    ElseIf txtAttenEmpNum.Text <> "" Then
    End If
End Sub
Private Sub txtHoursLate_GotFocus()
    With txtHoursLate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtHoursLate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And UpdateMode = True Then
        Call cmdUpdate_Click
    ElseIf KeyAscii = 13 And UpdateMode = False And cmdSubmit.Enabled = True Then
        Call cmdSubmit_Click
    End If
End Sub
Private Sub txtHoursLate_LostFocus()
    On Error Resume Next
    If Trim$(txtHoursLate.Text) = "" Then
        txtHoursLate.Text = Round("0.00", 2)
    Else
        txtHoursLate.Text = Round(txtHoursLate.Text, 2)
    End If
End Sub
Private Sub txtNotes_Change()
    strNotes = Replace$(Trim$(txtNotes.Text), vbCrLf, "")
End Sub
Private Sub txtNotes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And UpdateMode = True Then
        Call cmdUpdate_Click
    ElseIf KeyAscii = 13 And UpdateMode = False And cmdSubmit.Enabled = True Then
        Call cmdSubmit_Click
    End If
End Sub
