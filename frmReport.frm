VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Attendance Reports"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   6315
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   7515
      Begin VB.Frame Frame7 
         Height          =   6015
         Left            =   4200
         TabIndex        =   22
         Top             =   180
         Width           =   3195
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove Selected"
            Height          =   420
            Left            =   480
            TabIndex        =   25
            Top             =   5460
            Width           =   990
         End
         Begin VB.ListBox lstEmpReport 
            Appearance      =   0  'Flat
            Height          =   4905
            Left            =   240
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   420
            Width           =   2715
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear All"
            Height          =   420
            Left            =   1740
            TabIndex        =   23
            Top             =   5460
            Width           =   990
         End
         Begin VB.Label lblReportList 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Report List"
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
            Left            =   240
            TabIndex        =   26
            Top             =   180
            Width           =   2670
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1035
         Left            =   3420
         TabIndex        =   11
         Top             =   2580
         Width           =   675
         Begin VB.CommandButton cmdAddAll 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   70
            TabIndex        =   13
            ToolTipText     =   "Add All"
            Top             =   540
            Width           =   510
         End
         Begin VB.CommandButton cmdAddOne 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   70
            TabIndex        =   12
            ToolTipText     =   "Add Selected"
            Top             =   180
            Width           =   510
         End
      End
      Begin TabDlg.SSTab SSTab 
         Height          =   6015
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   10610
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Office"
         TabPicture(0)   =   "frmReport.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblOfficeEmp"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lstOfficeEmp"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Shop"
         TabPicture(1)   =   "frmReport.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblShopEmp"
         Tab(1).Control(1)=   "lstShopEmp"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Wooster"
         TabPicture(2)   =   "frmReport.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblWoosterEmp"
         Tab(2).Control(1)=   "lstWoosterShopEmp"
         Tab(2).ControlCount=   2
         Begin VB.ListBox lstOfficeEmp 
            Appearance      =   0  'Flat
            Height          =   4905
            ItemData        =   "frmReport.frx":0D1E
            Left            =   240
            List            =   "frmReport.frx":0D25
            MultiSelect     =   2  'Extended
            TabIndex        =   18
            Top             =   720
            Width           =   2655
         End
         Begin VB.CommandButton cmdShopSel 
            Caption         =   "Selection"
            Height          =   360
            Left            =   -74160
            TabIndex        =   17
            Top             =   5460
            Width           =   990
         End
         Begin VB.ListBox lstShopEmp 
            Appearance      =   0  'Flat
            Height          =   4905
            ItemData        =   "frmReport.frx":0D37
            Left            =   -74760
            List            =   "frmReport.frx":0D3E
            MultiSelect     =   2  'Extended
            TabIndex        =   16
            Top             =   720
            Width           =   2655
         End
         Begin VB.ListBox lstWoosterShopEmp 
            Appearance      =   0  'Flat
            Height          =   4905
            ItemData        =   "frmReport.frx":0D4E
            Left            =   -74760
            List            =   "frmReport.frx":0D55
            MultiSelect     =   2  'Extended
            TabIndex        =   15
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label lblWoosterEmp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Wooster Employees"
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
            Left            =   -74940
            TabIndex        =   21
            Top             =   420
            Width           =   3060
         End
         Begin VB.Label lblOfficeEmp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Office Employees"
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
            Left            =   0
            TabIndex        =   20
            Top             =   420
            Width           =   3120
         End
         Begin VB.Label lblShopEmp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Shop Employees"
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
            Left            =   -74880
            TabIndex        =   19
            Top             =   420
            Width           =   3000
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Settings"
      Height          =   2835
      Left            =   720
      TabIndex        =   0
      Top             =   6360
      Width           =   6135
      Begin VB.Frame Frame1 
         Caption         =   "Per Page"
         Height          =   855
         Left            =   3180
         TabIndex        =   28
         Top             =   1860
         Width           =   1095
         Begin VB.OptionButton optSingle 
            Caption         =   "Single"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   540
            Width           =   855
         End
         Begin VB.OptionButton optMulti 
            Caption         =   "Multiple"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Zero Entries"
         Height          =   855
         Left            =   4380
         TabIndex        =   9
         Top             =   1860
         Width           =   1155
         Begin VB.OptionButton optHide 
            Caption         =   "Hide"
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   250
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optShow 
            Caption         =   "Show"
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   540
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Print Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1260
         TabIndex        =   7
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Frame Frame3 
         Caption         =   "Date Range"
         Height          =   1455
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   5055
         Begin VB.CheckBox chkDateRange 
            Caption         =   "Use Date Range"
            Height          =   255
            Left            =   1920
            TabIndex        =   8
            Top             =   1080
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DTEnd 
            Height          =   375
            Left            =   3000
            TabIndex        =   2
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   -2147483635
            CustomFormat    =   "MM-dd-yyyy"
            Format          =   316342273
            CurrentDate     =   40487
         End
         Begin MSComCtl2.DTPicker DTStart 
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   -2147483635
            CustomFormat    =   "MM-dd-yyyy"
            Format          =   316342273
            CurrentDate     =   40487
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Date:"
            Height          =   195
            Left            =   600
            TabIndex        =   6
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ending Date:"
            Height          =   195
            Left            =   3360
            TabIndex        =   5
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "è"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   21.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2280
            TabIndex        =   4
            Top             =   540
            Width           =   480
         End
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid Grid1 
      Height          =   1095
      Left            =   7020
      TabIndex        =   27
      Top             =   6420
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1931
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      NoFocusHighlightBackColor=   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
      HotTrack        =   -1  'True
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bolShowZeroEntries As Boolean
Public Sub AddEmpToReportSingle(ByVal EmpNum As String)
    Dim rs        As New ADODB.Recordset
    Dim strSQL1   As String
    Dim sFntUnder As New StdFont
    sFntUnder.Underline = True
    sFntUnder.Name = "Tahoma"
    Dim sFntNormal As New StdFont
    sFntNormal.Underline = False
    sFntNormal.Name = "Tahoma"
    On Error Resume Next
    intRowsAdded = 0
    intUnExcusedRowsAdded = 0
    intExcusedRowsAdded = 0
    intOtherRowsAdded = 0
    cn_Global.CursorLocation = adUseClient
    DTStartDate = Format$(frmReport.DTStart.Value, "MM/DD/YYYY")
    DTEndDate = Format$(frmReport.DTEnd.Value, "MM/DD/YYYY")
    strSQL1 = "SELECT * FROM attendb.attenentries attenentries_0" & " WHERE (attenentries_0.idAttenEmpNum='" & EmpNum & "')" & (IIf(chkDateRange.Value = 1, " AND (attenentries_0.idAttenEntryDate>={d '" & Format$(frmReport.DTStart.Value, strDBDateFormat) & "'} AND attenentries_0.idAttenEntryDate<={d '" & Format$(frmReport.DTEnd.Value, strDBDateFormat) & "'})", "")) & " AND (attenentries_0.idAttenEmpNum='" & EmpNum & "')" & " ORDER BY attenentries_0.idAttenEntryDate Desc"
    Set rs = cn_Global.Execute(strSQL1)
    With rs
        strReportNum = EmpNum
        strReportName = ReturnEmpInfo(EmpNum).Name
    End With
    If Not bolShowZeroEntries And rs.RecordCount < 1 Then
        NoEntries = True
        Exit Sub
    End If
    Do Until rs.EOF
        With rs
            Grid1.Rows = Grid1.Rows + 1
            Grid1.CellDetails intGridRow, 1, Format$(!idAttenEntryDate, strUserDateFormat), DT_CENTER
            If Format$(!idAttenEntryDateTo, strUserDateFormat) <> Format$("2000-01-01", strUserDateFormat) Then Grid1.CellDetails intGridRow, 2, Format$(!idAttenEntryDateTo, strUserDateFormat), DT_CENTER
            Select Case !idAttenExcused
                Case "EXCUSED"
                    Grid1.CellDetails intGridRow, 3, !idAttenExcused, DT_CENTER, , &H80FF80
                    intExcusedRowsAdded = intExcusedRowsAdded + 1
                Case "UNEXCUSED"
                    Grid1.CellDetails intGridRow, 3, !idAttenExcused, DT_CENTER, , &H8080FF
                    intUnExcusedRowsAdded = intUnExcusedRowsAdded + 1
                Case "OTHER"
                    Grid1.CellDetails intGridRow, 3, !idAttenExcused, DT_CENTER, , &HFFFF80
                    intOtherRowsAdded = intOtherRowsAdded + 1
            End Select
            Grid1.CellDetails intGridRow, 4, !idAttenTimeOffType, DT_CENTER
            Grid1.CellDetails intGridRow, 5, !idAttenPartialDay, DT_CENTER
            Grid1.CellDetails intGridRow, 6, !idAttenNotes, DT_CENTER
            intGridRow = intGridRow + 1
            intRowsAdded = intRowsAdded + 1
            .MoveNext ' goto next entry
        End With
    Loop
End Sub
Public Sub AddEmpToReportMulti(ByVal EmpNum As String)
    Dim rs        As New ADODB.Recordset
    Dim strSQL1   As String
    Dim sFntUnder As New StdFont
    sFntUnder.Underline = True
    sFntUnder.Name = "Tahoma"
    Dim sFntNormal As New StdFont
    sFntNormal.Underline = False
    sFntNormal.Name = "Tahoma"
    On Error Resume Next
    intRowsAdded = 0
    intUnExcusedRowsAdded = 0
    intExcusedRowsAdded = 0
    intOtherRowsAdded = 0
    cn_Global.CursorLocation = adUseClient
    DTStartDate = Format$(frmReport.DTStart.Value, "MM/DD/YYYY")
    DTEndDate = Format$(frmReport.DTEnd.Value, "MM/DD/YYYY")
    strSQL1 = "SELECT * FROM attendb.attenentries attenentries_0" & " WHERE (attenentries_0.idAttenEmpNum='" & EmpNum & "')" & (IIf(chkDateRange.Value = 1, " AND (attenentries_0.idAttenEntryDate>={d '" & Format$(frmReport.DTStart.Value, strDBDateFormat) & "'} AND attenentries_0.idAttenEntryDate<={d '" & Format$(frmReport.DTEnd.Value, strDBDateFormat) & "'})", "")) & " AND (attenentries_0.idAttenEmpNum='" & EmpNum & "')" & " ORDER BY attenentries_0.idAttenEntryDate Desc"
    Set rs = cn_Global.Execute(strSQL1)
    With rs
        strReportNum = EmpNum
        strReportName = ReturnEmpInfo(EmpNum).Name
    End With
    If Not bolShowZeroEntries And rs.RecordCount < 1 Then
        NoEntries = True
        Exit Sub
    End If
    Do Until rs.EOF
        With rs
            Grid1.Rows = Grid1.Rows + 1
            Grid1.CellDetails intGridRow, 1, Format$(!idAttenEntryDate, strUserDateFormat), DT_CENTER
            If Format$(!idAttenEntryDateTo, strUserDateFormat) <> Format$("2000-01-01", strUserDateFormat) Then Grid1.CellDetails intGridRow, 2, Format$(!idAttenEntryDateTo, strUserDateFormat), DT_CENTER
            Select Case !idAttenExcused
                Case "EXCUSED"
                    Grid1.CellDetails intGridRow, 3, !idAttenExcused, DT_CENTER, , &H80FF80
                    intExcusedRowsAdded = intExcusedRowsAdded + 1
                Case "UNEXCUSED"
                    Grid1.CellDetails intGridRow, 3, !idAttenExcused, DT_CENTER, , &H8080FF
                    intUnExcusedRowsAdded = intUnExcusedRowsAdded + 1
                Case "OTHER"
                    Grid1.CellDetails intGridRow, 3, !idAttenExcused, DT_CENTER, , &HFFFF80
                    intOtherRowsAdded = intOtherRowsAdded + 1
            End Select
            Grid1.CellDetails intGridRow, 4, !idAttenTimeOffType, DT_CENTER
            Grid1.CellDetails intGridRow, 5, !idAttenPartialDay, DT_CENTER
            Grid1.CellDetails intGridRow, 6, !idAttenNotes, DT_CENTER
            intGridRow = intGridRow + 1
            intRowsAdded = intRowsAdded + 1
            .MoveNext ' goto next entry
        End With
    Loop
End Sub
Public Sub EmpListReportMulti()
    Dim i As Integer, PagesPrinted As Integer
    Grid1.Visible = False
    Grid1.Redraw = False
    Grid1.Header = True
    Grid1.Clear
    Grid1.Rows = 1
    NoEntries = False
    intGridRow = 1
    frmPBar.PBar1.Max = UBound(strListLine)
    frmPBar.PBar1.Value = 0
    frmPBar.lblInfo.Caption = "Print Job Spooling..." & vbCrLf & "Printer = " & Printer.DeviceName
    frmPBar.Show
    DoEvents
    Printer.Font.Underline = False
    Printer.ScaleMode = 1
    Printer.Orientation = vbPRORLandscape
    Printer.DrawWidth = 1
    Printer.Print ""
    Printer.Print "    " & strReportMsg
    Printer.Print ""
    Printer.Print "    Report date: " & Date & " " & Time & "      Printed by: " & UCase$(Environ$("USERNAME"))
    For i = 0 To UBound(strListLine) - 1
        If bolStop = True Then
            Printer.KillDoc
            bolStop = False
            frmPBar.Hide
            'Grid1.Visible = True
            Grid1.Redraw = True
            Exit Sub
        End If
        NoEntries = False
        ShowData
        AddEmpToReportMulti strListLine(i)
        If NoEntries = False Then
            PagesPrinted = PagesPrinted + 1
            DrawSGridForPrint
            ReSizeSGrid
            'PrintSGridSingle Grid1
            PrintSGridMulti Grid1
        Else
            ' do not print
        End If
        Grid1.Clear
        Grid1.Rows = 1
        intGridRow = 1
        frmPBar.PBar1.Value = i
        HideData
        DoEvents
    Next
    Printer.EndDoc
    frmPBar.Hide
    ClearAvgQryTimes
    Dim blah
    blah = MsgBox(Printer.Page - 1 & " pages have been sent to " & Printer.DeviceName, vbOKOnly + vbInformation, "Print job complete")
    frmReport.SetFocus
    'Grid1.Visible = True
    Grid1.Redraw = True
End Sub
Public Sub GetEmpNumList()
    Dim i                  As Integer
    Dim strListLineSplit() As String
    lstEmpReport.Visible = False
    ReDim strListLine(lstEmpReport.ListCount)
    For i = 0 To lstEmpReport.ListCount - 1
        lstEmpReport.ListIndex = i
        strListLineSplit = Split(lstEmpReport.Text, " - ")
        strListLine(i) = strListLineSplit(1)
    Next i
    lstEmpReport.Visible = True
End Sub
Private Sub ReSizeSGrid()
    Grid1.Redraw = False
    Dim i As Integer, intCellPadding As Integer
    intCellPadding = 20
    For i = 1 To Grid1.Columns
        Grid1.AutoWidthColumn i
        Grid1.ColumnWidth(i) = Grid1.ColumnWidth(i) + intCellPadding
    Next i
    Grid1.Redraw = True
End Sub
Private Sub DrawSGridForPrint()
    Dim i     As Integer
    Dim Hours As Single
    On Error Resume Next
    AttenVals.PartialUnExcused = 0
    AttenVals.PartialExcused = 0
    AttenVals.FullExcused = 0
    AttenVals.FullUnExcused = 0
    For i = 1 To Grid1.Rows ' - 1
        Hours = Hours + Grid1.CellText(i, 5)
        CountExcusesAll Grid1.CellText(i, 3), Grid1.CellText(i, 4)
    Next
    If Hours <= 0 Then
        Grid1.CellDetails Grid1.Rows, 5, "| Total Hrs: None |", DT_CENTER
    Else
        Grid1.CellDetails Grid1.Rows, 5, "| Total Hrs: " & Hours & " |", DT_CENTER
    End If
    If intRowsAdded = 0 Then
        strReportEntryCount = " Total Entries: 0 "
    Else
        strReportEntryCount = " Total Entries: " & intRowsAdded & " "
    End If
    strReportInfo = " " & AttenVals.PartialUnExcused & " Partial Unexcused | " & AttenVals.PartialExcused & " Partial Excused | " & AttenVals.FullUnExcused & " Full Unexcused | " & AttenVals.FullExcused & " Full Excused | " & intOtherRowsAdded & " Other "
End Sub
Public Sub EmpListReportSingle()
    Dim i As Integer, PagesPrinted As Integer
    Grid1.Visible = False
    Grid1.Redraw = False
    Grid1.Header = True
    Grid1.Clear
    Grid1.Rows = 1
    NoEntries = False
    intGridRow = 1
    frmPBar.PBar1.Max = UBound(strListLine)
    frmPBar.PBar1.Value = 0
    frmPBar.lblInfo.Caption = "Print Job Spooling..." & vbCrLf & "Printer = " & Printer.DeviceName
    frmPBar.Show
    DoEvents
    For i = 0 To UBound(strListLine) - 1
        If bolStop = True Then
            Printer.KillDoc
            bolStop = False
            frmPBar.Hide
            'Grid1.Visible = True
            Grid1.Redraw = True
            Exit Sub
        End If
        NoEntries = False
        ShowData
        AddEmpToReportSingle strListLine(i)
        If NoEntries = False Then
            PagesPrinted = PagesPrinted + 1
            DrawSGridForPrint
            ReSizeSGrid
            PrintSGridSingle Grid1
        Else
            ' do not print
        End If
        Grid1.Clear
        Grid1.Rows = 1
        intGridRow = 1
        frmPBar.PBar1.Value = i
        HideData
        DoEvents
    Next
    Printer.EndDoc
    frmPBar.Hide
    ClearAvgQryTimes
    Dim blah
    blah = MsgBox(Printer.Page - 1 & " pages have been sent to " & Printer.DeviceName, vbOKOnly + vbInformation, "Print job complete")
    frmReport.SetFocus
    ' Grid1.Visible = True
    Grid1.Redraw = True
End Sub
Private Sub PrintSGridMulti(FlexGrid As vbalGrid)
    On Error Resume Next
    FlexGrid.Redraw = False
    Dim intPadding    As Integer
    Dim PrevX         As Integer, PrevY As Integer, intMidStart As Integer, intMidLen As Integer, intTotLen As Integer
    Dim strSizedTxt   As String, strOrigTxt As String
    Dim bolLongLine   As Boolean
    Dim TwipPix       As Long
    Dim lngYTopOfGrid As Long
    bolLongLine = False
    Dim sMsg            As String
    Dim intCenterOffset As Long
    Dim lngStartY       As Long, lngStartX As Long, lngEndX As Long, lngEndY As Long
    Dim xmax            As Single, xmin As Single
    xmin = 300
    xmax = 14800
    Dim ymax As Single, ymin As Single
    ymin = 1500
    ymax = 10800
    Printer.Font.Underline = False
    Printer.ScaleMode = 1
    Printer.Orientation = vbPRORLandscape
    Printer.DrawWidth = 1
    Printer.DrawStyle = vbSolid
    If Printer.CurrentY + (Printer.TextHeight("####") * 8) >= ymax Then  ' new page
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
    Printer.Print ""
    Printer.DrawStyle = vbDash
    Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY), vbBlack
    Printer.DrawStyle = vbSolid
    Printer.ForeColor = vbBlack
    sMsg = "Emp #: " & strReportNum & "   Name: " & strReportName
    With Printer
        .ScaleMode = 1
        'Printer.Print
        Printer.CurrentY = Printer.CurrentY + 100
        .FontSize = 10
        Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(sMsg) / 2)
        Printer.Print sMsg
        Printer.FontSize = 8
    End With
    Printer.FontSize = 7
    'Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 100
    Const GAP = 40
    With Printer.Font
        .Name = FlexGrid.Font.Name
        .Size = 9
    End With
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
    Printer.CurrentY = Printer.CurrentY + 100
    'Printer.Print ""
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
        TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
        XFirstColumn = xmin + TwipPix * GAP
        lngYTopOfGrid = Printer.CurrentY
        Printer.CurrentY = Printer.CurrentY + GAP
        If FlexGrid.Header = True Then
            X = xmin + GAP
            For c = 1 To .Columns ' - 1
                Printer.CurrentX = X
                TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                PrevY = Printer.CurrentY
                If c = .Columns Then
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
            Printer.Print
        End If
        For R = 1 To .Rows ' - 1
            If bolStop = True Then
                Printer.EndDoc
                bolStop = False
                frmPBar.Visible = False
                Exit Sub
            End If
            ' Draw a line above this row.
            If R > 0 Then
                Printer.Line (XFirstColumn, Printer.CurrentY)-(xmax, Printer.CurrentY), vbBlack
            End If
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Print the entries on this row.
            X = xmin + GAP
            For c = 1 To .Columns ' - 1
                If frmPrinters.optCenterJust And c < .Columns Then
                    intCenterOffset = ((.ColumnWidth(c) * Screen.TwipsPerPixelX) / 2) - (Printer.TextWidth(.CellText(R, c)) / 2)
                Else
                    intCenterOffset = 0
                End If
                Printer.CurrentX = X
                If .CellText(R, c) <> "" And Printer.TextWidth(.CellText(R, c)) + intPadding >= xmax - Printer.CurrentX Then           '.ColWidth(c)
                    lngStartY = Printer.CurrentY + Printer.TextHeight(.CellText(R, c))
                    strOrigTxt = .CellText(R, c)
                    Do Until intTotLen >= Len(strOrigTxt)
                        Do Until Printer.TextWidth(strSizedTxt) + intPadding >= xmax - Printer.CurrentX Or intTotLen >= Len(strOrigTxt)
                            intMidLen = intMidLen + 1
                            intTotLen = intTotLen + 1
                            strSizedTxt = Mid$(strOrigTxt, intMidStart, intMidLen)
                        Loop
                        intMidStart = intMidStart + intMidLen ' - 1
                        intMidLen = 1
                        Printer.Font.Underline = .CellFont(R, c).Underline
                        If .CellFont(R, c).Underline = True Then
                            Printer.ForeColor = vbBlack
                        Else
                            Printer.ForeColor = &H404040
                        End If
                        Printer.Print strSizedTxt 'Left$(.CellText (R, c), i)
                        lngEndY = Printer.CurrentY + GAP
                        PrevY = Printer.CurrentY
                        Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor(R, 3), BF
                        Printer.CurrentY = PrevY + 5
                        If Printer.CurrentY >= ymax Then ' new page
                            Printer.Line (XFirstColumn, lngYTopOfGrid)-(xmax, Printer.CurrentY + GAP), vbBlack, B
                            X = xmin
                            For cc = 1 To .Columns - 1
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
                            lngStartY = Printer.CurrentY '+ Printer.TextHeight(.CellText(R, c))
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
                    'bolLongLine = False
                    PrevY = Printer.CurrentY - GAP ' + 10
                    If c = 3 Then
                        lngStartY = Printer.CurrentY - GAP + 5
                        lngStartX = Printer.CurrentX - GAP + 5
                        lngEndX = Printer.CurrentX + .ColumnWidth(c) * Screen.TwipsPerPixelX + GAP ' - 10
                        lngEndY = Printer.CurrentY + Printer.TextHeight(.CellText(R, c)) + GAP - 5
                        Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor(R, c), BF
                    End If
                    Printer.CurrentX = X + intCenterOffset
                    TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                    Printer.Font.Underline = .CellFont(R, c).Underline
                    If .CellFont(R, c).Underline = True Then
                        Printer.ForeColor = vbBlack
                    Else
                        Printer.ForeColor = &H404040   '&H808080
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
                For c = 1 To .Columns - 1 '3
                    TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX '+ GAP
                    X = X + TwipPix + 2 * GAP
                    Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack 'ymax
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
        For c = 1 To .Columns - 1 '3
            TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
            X = X + TwipPix + 2 * GAP
            'vbBlack
            Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack 'Printer.CurrentY
        Next c
    End With
End Sub
Public Sub PrintSGridSingle(FlexGrid As vbalGrid)
    On Error Resume Next
    FlexGrid.Redraw = False
    'Dim sMsg As String
    Dim intPadding    As Integer
    Dim PrevX         As Integer, PrevY As Integer, intMidStart As Integer, intMidLen As Integer, intTotLen As Integer
    Dim strSizedTxt   As String, strOrigTxt As String
    Dim bolLongLine   As Boolean
    Dim TwipPix       As Long
    Dim lngYTopOfGrid As Long
    bolLongLine = False
    Dim sMsg            As String
    Dim intCenterOffset As Long
    Dim lngStartY       As Long, lngStartX As Long, lngEndX As Long, lngEndY As Long
    Dim xmax            As Single, xmin As Single
    xmin = 300
    xmax = 14800
    Dim ymax As Single, ymin As Single
    ymin = 1500
    ymax = 10800
    Printer.Font.Underline = False
    Printer.ScaleMode = 1
    Printer.Orientation = vbPRORLandscape
    Printer.DrawWidth = 1
    Printer.DrawStyle = vbSolid
    If strReportType = "MULTI" Then
        sMsg = ""
    ElseIf strReportType = "SINGLE" Then
        sMsg = "Emp #: " & strReportNum & "   Name: " & strReportName
    End If
    With Printer
        .ScaleMode = 1
        Printer.Print
        .FontSize = 20
        Printer.CurrentX = (xmax / 2) - (Printer.TextWidth(sMsg) / 2)
        Printer.Print sMsg
        Printer.FontSize = 8
    End With
    Printer.FontSize = 7
    Printer.Print "    " & strReportMsg
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
        TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
        XFirstColumn = xmin + TwipPix * GAP
        lngYTopOfGrid = Printer.CurrentY
        Printer.CurrentY = Printer.CurrentY + GAP
        If FlexGrid.Header = True Then
            X = xmin + GAP
            For c = 1 To .Columns
                Printer.CurrentX = X
                TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                PrevY = Printer.CurrentY
                If c = .Columns Then
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
            Printer.Print
        End If
        For R = 1 To .Rows '- 1
            If bolStop = True Then
                Printer.EndDoc
                bolStop = False
                frmPBar.Visible = False
                Exit Sub
            End If
            ' Draw a line above this row.
            If R > 0 Then
                Printer.Line (XFirstColumn, Printer.CurrentY)-(xmax, Printer.CurrentY), vbBlack
            End If
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Print the entries on this row.
            X = xmin + GAP
            For c = 1 To .Columns ' - 1
                If frmPrinters.optCenterJust And c < .Columns Then
                    intCenterOffset = ((.ColumnWidth(c) * Screen.TwipsPerPixelX) / 2) - (Printer.TextWidth(.CellText(R, c)) / 2)
                Else
                    intCenterOffset = 0
                End If
                Printer.CurrentX = X
                If .CellText(R, c) <> "" And Printer.TextWidth(.CellText(R, c)) + intPadding >= xmax - Printer.CurrentX Then           '.ColWidth(c)
                    lngStartY = Printer.CurrentY + Printer.TextHeight(.CellText(R, c))
                    strOrigTxt = .CellText(R, c)
                    Do Until intTotLen >= Len(strOrigTxt)
                        Do Until Printer.TextWidth(strSizedTxt) + intPadding >= xmax - Printer.CurrentX Or intTotLen >= Len(strOrigTxt)
                            intMidLen = intMidLen + 1
                            intTotLen = intTotLen + 1
                            strSizedTxt = Mid$(strOrigTxt, intMidStart, intMidLen)
                        Loop
                        intMidStart = intMidStart + intMidLen ' - 1
                        intMidLen = 1
                        Printer.Font.Underline = .CellFont(R, c).Underline
                        If .CellFont(R, c).Underline = True Then
                            Printer.ForeColor = vbBlack
                        Else
                            Printer.ForeColor = &H404040
                        End If
                        Printer.Print strSizedTxt 'Left$(.CellText (R, c), i)
                        lngEndY = Printer.CurrentY + GAP
                        PrevY = Printer.CurrentY
                        Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor(R, 3), BF
                        Printer.CurrentY = PrevY + 5
                        If Printer.CurrentY >= ymax Then ' new page
                            Printer.Line (XFirstColumn, lngYTopOfGrid)-(xmax, Printer.CurrentY + GAP), vbBlack, B
                            X = xmin
                            For cc = 1 To .Columns - 1
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
                            lngStartY = Printer.CurrentY '+ Printer.TextHeight(.CellText(R, c))
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
                    'bolLongLine = False
                    PrevY = Printer.CurrentY - GAP ' + 10
                    If c = 3 Then
                        lngStartY = Printer.CurrentY - GAP + 5
                        lngStartX = Printer.CurrentX - GAP + 5
                        lngEndX = Printer.CurrentX + .ColumnWidth(c) * Screen.TwipsPerPixelX + GAP ' - 10
                        lngEndY = Printer.CurrentY + Printer.TextHeight(.CellText(R, c)) + GAP - 5
                        Printer.Line (lngStartX, lngStartY)-(lngEndX, lngEndY), .CellBackColor(R, c), BF
                    End If
                    Printer.CurrentX = X + intCenterOffset
                    TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                    Printer.Font.Underline = .CellFont(R, c).Underline
                    If .CellFont(R, c).Underline = True Then
                        Printer.ForeColor = vbBlack
                    Else
                        Printer.ForeColor = &H404040   '&H808080
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
                For c = 1 To .Columns - 1 '3
                    TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX '+ GAP
                    X = X + TwipPix + 2 * GAP
                    Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack 'ymax
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
        For c = 1 To .Columns - 1 '3
            TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
            X = X + TwipPix + 2 * GAP
            'vbBlack
            Printer.Line (X, lngYTopOfGrid)-(X, Printer.CurrentY), vbBlack 'Printer.CurrentY
        Next c
    End With
    Printer.NewPage
End Sub
Public Sub LoadEmpList()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim iOffice As Integer, iShop As Integer, iWoosterShop As Integer
    cn_Global.CursorLocation = adUseClient
    strSQL1 = "SELECT * From EmpList Where idIsActive = 'TRUE' Order By idName"
    Set rs = cn_Global.Execute(strSQL1)
    lstOfficeEmp.Clear
    lstShopEmp.Clear
    lstWoosterShopEmp.Clear
    iOffice = 0
    iShop = 0
    iWoosterShop = 0
    Do Until rs.EOF
        With rs
            If !idLocation2 = "OFFICE" And !idLocation1 = "BREMEN" Or !idLocation2 = "OFFICE" And !idLocation1 = "ROCKY MTN" Or !idLocation2 = "OFFICE" And !idLocation1 = "NUCLEAR" Then
                lstOfficeEmp.AddItem !idName & " - " & !idNumber, iOffice
                iOffice = iOffice + 1
            ElseIf !idLocation2 = "SHOP" And !idLocation1 = "BREMEN" Then
                lstShopEmp.AddItem !idName & " - " & !idNumber, iShop
                iShop = iShop + 1
            ElseIf !idLocation1 = "WOOSTER" Then
                lstWoosterShopEmp.AddItem !idName & " - " & !idNumber, iWoosterShop
                iWoosterShop = iWoosterShop + 1
            End If
            .MoveNext
        End With
    Loop
    lblOfficeEmp.Caption = "Office Employees - " & lstOfficeEmp.ListCount
    lblShopEmp.Caption = "Shop Employees - " & lstShopEmp.ListCount
    lblWoosterEmp.Caption = "Wooster Employees - " & lstWoosterShopEmp.ListCount
End Sub
Private Sub chkDateRange_Click()
    If chkDateRange.Value = 1 Then
        DTStart.Enabled = True
        DTEnd.Enabled = True
    Else
        DTStart.Enabled = False
        DTEnd.Enabled = False
    End If
End Sub
Private Sub cmdAddAll_Click()
    Call AddEmps("ADDALL")
End Sub
Private Sub cmdAddOne_Click()
    Call AddEmps("ADDSEL")
End Sub
Private Sub AddEmps(Adds As String)
    Dim i As Integer
    Select Case SSTab.Tab
        Case 0 'Office
            Select Case Adds
                Case "ADDSEL"
                    For i = 0 To lstOfficeEmp.ListCount - 1
                        If lstOfficeEmp.Selected(i) = True Then
                            lstEmpReport.AddItem lstOfficeEmp.List(i)
                        End If
                    Next
                    lstEmpReport.Refresh
                    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
                Case "ADDALL"
                    For i = 0 To lstOfficeEmp.ListCount - 1
                        lstEmpReport.AddItem lstOfficeEmp.List(i)
                    Next
                    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
            End Select
        Case 1 'Shop
            Select Case Adds
                Case "ADDSEL"
                    For i = 0 To lstShopEmp.ListCount - 1
                        If lstShopEmp.Selected(i) = True Then
                            lstEmpReport.AddItem lstShopEmp.List(i)
                        End If
                    Next
                    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
                Case "ADDALL"
                    For i = 0 To lstShopEmp.ListCount - 1
                        lstEmpReport.AddItem lstShopEmp.List(i)
                    Next
                    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
            End Select
        Case 2 'Wooster
            Select Case Adds
                Case "ADDSEL"
                    For i = 0 To lstWoosterShopEmp.ListCount - 1
                        If lstWoosterShopEmp.Selected(i) = True Then
                            lstEmpReport.AddItem lstWoosterShopEmp.List(i)
                        End If
                    Next
                    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
                Case "ADDALL"
                    For i = 0 To lstWoosterShopEmp.ListCount - 1
                        lstEmpReport.AddItem lstWoosterShopEmp.List(i)
                    Next
                    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
            End Select
    End Select
End Sub
Private Sub cmdClear_Click()
    Grid1.Clear
    Grid1.Rows = 1
    lstEmpReport.Clear
    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
End Sub
Private Sub cmdGo_Click()
    Dim blah
    If lstEmpReport.ListCount < 1 Then
        blah = MsgBox("Please select some employees.", vbExclamation + vbOKOnly, "No employees in list")
        Exit Sub
    End If
    If optMulti Then
        GetEmpNumList
        frmPrinters.Show 1
        If bolCancelPrint Then
            bolCancelPrint = False
            Exit Sub
        Else
            If chkDateRange.Value = 0 Then
                strReportMsg = "All entries."
            ElseIf chkDateRange.Value = 1 Then
                strReportMsg = "Date range from " & DTStart.Value & " to " & DTEnd.Value
            End If
            EmpListReportMulti
        End If
    ElseIf Not optMulti Then
        If chkDateRange.Value = 0 Then
            strReportMsg = "All entries."
        ElseIf chkDateRange.Value = 1 Then
            strReportMsg = "Date range from " & DTStart.Value & " to " & DTEnd.Value
        End If
        GetEmpNumList
        blah = MsgBox("This will print up to " & UBound(strListLine) & " individual reports." & vbCrLf & vbCrLf & "OK to continue?", vbOKCancel + vbExclamation, "Print individual reports")
        If blah = vbCancel Then
            Exit Sub
        Else
            strReportType = "SINGLE"
            frmPrinters.Show 1
            If bolCancelPrint Then
                bolCancelPrint = False
                Exit Sub
            Else
                EmpListReportSingle
            End If
        End If
    End If
End Sub
Private Sub cmdRemove_Click()
    Dim A As Integer
    Do Until lstEmpReport.SelCount = 0
        If lstEmpReport.Selected(A) Then lstEmpReport.RemoveItem A: A = A - 1
        A = A + 1
    Loop
    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
End Sub
Private Sub cmdShopSel_Click()
    Dim i As Integer
    For i = 0 To lstShopEmp.ListCount - 1
        If lstShopEmp.Selected(i) = True Then
            lstEmpReport.AddItem lstShopEmp.List(i)
        End If
    Next
    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
End Sub
Private Sub Form_Click()
    frmPBar.Top = frmReport.Top + frmReport.Height / 2 - (frmPBar.Height / 2)
    frmPBar.Left = frmReport.Left + frmReport.Width / 2 - (frmPBar.Width / 2)
End Sub
Private Sub Form_Load()
    Grid1.AddColumn 0, "Date"
    Grid1.AddColumn 1, "Date To"
    Grid1.AddColumn 2, "Excuse"
    Grid1.AddColumn 3, "Type"
    Grid1.AddColumn 4, "Hours"
    Grid1.AddColumn 5, "Notes"
    Grid1.Gridlines = True
    Grid1.Header = False
    DTStart.Value = Date
    DTEnd.Value = Date
    LoadEmpList
    bolShowZeroEntries = optShow
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmReport.Hide
End Sub
Private Sub lstEmpReport_DblClick()
    Dim A As Integer
    Do Until lstEmpReport.SelCount = 0
        If lstEmpReport.Selected(A) Then lstEmpReport.RemoveItem A: A = A - 1
        A = A + 1
    Loop
    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
End Sub
Private Sub lstOfficeEmp_DblClick()
    Dim i As Integer
    For i = 0 To lstOfficeEmp.ListCount - 1
        If lstOfficeEmp.Selected(i) = True Then
            lstEmpReport.AddItem lstOfficeEmp.List(i)
        End If
    Next
    lstEmpReport.Refresh
    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
End Sub
Private Sub lstShopEmp_DblClick()
    Dim i As Integer
    For i = 0 To lstShopEmp.ListCount - 1
        If lstShopEmp.Selected(i) = True Then
            lstEmpReport.AddItem lstShopEmp.List(i)
        End If
    Next
    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
End Sub
Private Sub lstShopEmp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim i As Integer
        For i = 0 To lstShopEmp.ListCount - 1
            If lstShopEmp.Selected(i) = True Then
                lstEmpReport.AddItem lstShopEmp.List(i)
            End If
        Next
        lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
    End If
End Sub
Private Sub lstWoosterShopEmp_DblClick()
    Dim i As Integer
    For i = 0 To lstWoosterShopEmp.ListCount - 1
        If lstWoosterShopEmp.Selected(i) = True Then
            lstEmpReport.AddItem lstWoosterShopEmp.List(i)
        End If
    Next
    lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
End Sub
Private Sub lstWoosterShopEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim i As Integer
        For i = 0 To lstWoosterShopEmp.ListCount - 1
            If lstWoosterShopEmp.Selected(i) = True Then
                lstEmpReport.AddItem lstWoosterShopEmp.List(i)
            End If
        Next
        lblReportList.Caption = "Report List - " & lstEmpReport.ListCount
    End If
End Sub
Private Sub optShow_Click()
    bolShowZeroEntries = optShow
End Sub
