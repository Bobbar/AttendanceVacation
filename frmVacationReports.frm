VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmVacationReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vacation Reports"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVacationReports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   9240
   StartUpPosition =   1  'CenterOwner
   Begin vbAcceleratorSGrid6.vbalGrid Grid1 
      Height          =   5295
      Left            =   60
      TabIndex        =   21
      Top             =   4020
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   9340
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   16777215
      GridLineColor   =   4210752
      NoFocusHighlightBackColor=   -2147483635
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
      HotTrack        =   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   300
      Left            =   5880
      TabIndex        =   4
      Top             =   3480
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reports"
      Height          =   3795
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8715
      Begin VB.CommandButton cmdToExcel 
         Caption         =   "To Excel"
         Height          =   300
         Left            =   7740
         TabIndex        =   22
         Top             =   3360
         Width           =   810
      End
      Begin VB.Frame Frame4 
         Caption         =   "Filters"
         Height          =   855
         Left            =   3600
         TabIndex        =   13
         Top             =   2400
         Width           =   4935
         Begin VB.CheckBox chkRockyMtn 
            Caption         =   "Rocky Mtn"
            Height          =   195
            Left            =   2460
            TabIndex        =   20
            Top             =   540
            Value           =   1  'Checked
            Width           =   1155
         End
         Begin VB.CheckBox chkNuclear 
            Caption         =   "Nuclear"
            Height          =   195
            Left            =   3840
            TabIndex        =   19
            Top             =   540
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkWooster 
            Caption         =   "Wooster"
            Height          =   195
            Left            =   1200
            TabIndex        =   18
            Top             =   540
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox chkBremen 
            Caption         =   "Bremen"
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   540
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.CheckBox chkTaken 
            Caption         =   "Taken"
            Height          =   195
            Left            =   3900
            TabIndex        =   16
            Top             =   240
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.CheckBox chkReScheduled 
            Caption         =   "ReScheduled"
            Height          =   195
            Left            =   2100
            TabIndex        =   15
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkRequested 
            Caption         =   "Requested"
            Height          =   195
            Left            =   420
            TabIndex        =   14
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Accruals"
         Height          =   855
         Left            =   3600
         TabIndex        =   10
         Top             =   180
         Width           =   4935
         Begin VB.CommandButton cmdAccruals 
            Caption         =   "Run"
            Enabled         =   0   'False
            Height          =   360
            Left            =   3540
            TabIndex        =   11
            Top             =   300
            Width           =   990
         End
         Begin VB.Label lblAccrualRange 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AccrualRange"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1020
            TabIndex        =   12
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Date Range"
         Height          =   1335
         Left            =   3600
         TabIndex        =   5
         Top             =   1080
         Width           =   4935
         Begin VB.CommandButton cmdDateRange 
            Caption         =   "Run"
            Height          =   360
            Left            =   2040
            TabIndex        =   8
            Top             =   840
            Width           =   1110
         End
         Begin MSComCtl2.DTPicker DTStartDate 
            Height          =   375
            Left            =   300
            TabIndex        =   6
            Top             =   300
            Width           =   1875
            _ExtentX        =   3307
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
            Format          =   213123073
            CurrentDate     =   40941
         End
         Begin MSComCtl2.DTPicker DTEndDate 
            Height          =   375
            Left            =   2940
            TabIndex        =   7
            Top             =   300
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
            Format          =   213123073
            CurrentDate     =   40941
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "è"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   24
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   2340
            TabIndex        =   9
            Top             =   180
            Width           =   510
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Who's On Vacation"
         Height          =   3075
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   3435
         Begin VB.CommandButton cmdWhosOnVaca 
            Caption         =   "Run"
            Height          =   360
            Left            =   2160
            TabIndex        =   3
            Top             =   1320
            Width           =   990
         End
         Begin VB.ListBox lstMonths 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2580
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   2
            Top             =   300
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmVacationReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strEmpsToRun() As String
Private Sub cmdAccruals_Click()
    Dim rs        As New ADODB.Recordset
    Dim cn        As New ADODB.Connection
    Dim strSQL1   As String
    Dim i         As Integer
    Dim sFntUnder As New StdFont
    sFntUnder.Underline = True
    sFntUnder.Name = "Tahoma"
    Dim sFntNormal As New StdFont
    sFntNormal.Underline = False
    sFntNormal.Name = "Tahoma"
    EmpListByCompany
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    Grid1.BackColor = colGridBusy
    Screen.MousePointer = vbHourglass
    Grid1.Redraw = False
    Grid1.Clear
    Grid1.Rows = 1
    For i = 0 To UBound(strEmpsToRun) '(strEmpInfo)
        strSQL1 = "SELECT * " & "FROM attendb.vacations vacations_0" & " WHERE (vacations_0.idStartDate>={d '" & Format$(dtAnniDate(strEmpInfo(i).HireDate).PreviousYear, strDBDateFormat) & "'}) AND (vacations_0.idEndDate<={d '" & Format$(dtFiscalYearEnd, strDBDateFormat) & "'})" & " AND (vacations_0.idEmpNum='" & strEmpsToRun(i) & "') " & IIf(chkRequested.Value = 0, "AND vacations_0.idStatus <> 'REQUESTED'", "") & IIf(chkReScheduled.Value = 0, "AND vacations_0.idStatus <> 'RESCHEDULED'", "") & IIf(chkTaken.Value = 0, "AND vacations_0.idStatus <> 'TAKEN'", "")
        rs.Open strSQL1, cn, adOpenKeyset
        With rs
            If .RecordCount = 0 Then GoTo NextLoop
            Grid1.Rows = Grid1.Rows + 1
            Grid1.CellDetails Grid1.Rows - 1, 1, ReturnEmpInfo(strEmpsToRun(i)).Name, DT_CENTER, , , , sFntUnder
            Grid1.CellDetails Grid1.Rows - 1, 2, "Anni. Date: " & dtAnniDate(ReturnEmpInfo(strEmpsToRun(i)).HireDate).PreviousYear, DT_CENTER, , , , sFntUnder
            'Grid1.CellDetails Grid1.Rows - 1, 3, "Weeks Avail: " & CalcYearsWorked(strEmpsToRun(i)).VacaWeeksAvail, DT_CENTER, , , , sFntUnder
            Grid1.CellDetails Grid1.Rows - 1, 3, "Weeks Avail: " & VacaAvailAccrual(strEmpsToRun(i)), DT_CENTER, , , , sFntUnder
            Do Until .EOF
                Grid1.Rows = Grid1.Rows + 1
                Grid1.CellDetails Grid1.Rows - 1, 1, !idStartDate, DT_CENTER
                Grid1.CellDetails Grid1.Rows - 1, 2, !idEndDate, DT_CENTER
                Grid1.CellDetails Grid1.Rows - 1, 3, !idStatus, DT_CENTER
                Grid1.CellDetails Grid1.Rows - 1, 4, !idNotes, DT_CENTER
                .MoveNext
            Loop
            Grid1.Rows = Grid1.Rows + 1
NextLoop:
            .Close
        End With
    Next i
    ReSizeSGrid
    Grid1.Redraw = True
    Screen.MousePointer = vbDefault
    Grid1.BackColor = &H80000005
    strReportTitle = "Accruals"
    strReportSubTitle = IIf(chkBremen, "Bremen ", "") & IIf(chkWooster, " Wooster ", "") & IIf(chkRockyMtn, " RockyMtn ", "") & IIf(chkNuclear, " Nuclear ", "")
End Sub
Private Sub CustomDateRange(StartDate As Date, EndDate As Date)
    Dim rs          As New ADODB.Recordset
    Dim cn          As New ADODB.Connection
    Dim strSQL1     As String
    Dim i           As Integer, b           As Integer
    Dim SortArray() As Variant
    ReDim SortArray(6, 0)
    Dim DateSortArray() As Variant
    ReDim DateSortArray(6, 0)
    Dim clStartDate   As Date, clEndDate As Date
    Dim bolFinalEntry As Boolean
    bolFinalEntry = False
    EmpListByCompany
    Dim CurEmp         As String, NextEmp As String, FirstEmp As String
    Dim bolHeaderAdded As Boolean
    Dim Found          As String
    bolHeaderAdded = False
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    Dim sFntUnder As New StdFont
    sFntUnder.Underline = True
    sFntUnder.Name = "Tahoma"
    Dim sFntNormal As New StdFont
    sFntNormal.Underline = False
    sFntNormal.Name = "Tahoma"
    Grid1.Redraw = False
    Grid1.Clear
    Grid1.Rows = 1
    strSQL1 = "SELECT * " & "FROM attendb.vacations vacations_0" & " WHERE (vacations_0.idStartDate>={d '" & Format$(StartDate, strDBDateFormat) & "'}) AND (vacations_0.idStartDate<={d '" & Format$(EndDate, strDBDateFormat) & "'})" & IIf(chkRequested.Value = 0, "AND vacations_0.idStatus <> 'REQUESTED'", "") & IIf(chkReScheduled.Value = 0, "AND vacations_0.idStatus <> 'RESCHEDULED'", "") & IIf(chkTaken.Value = 0, "AND vacations_0.idStatus <> 'TAKEN'", "")
    rs.Open strSQL1, cn, adOpenKeyset
    With rs
        If .RecordCount = 0 Then Exit Sub
        Do Until .EOF
            Found = InStr(1, vbNullChar & Join(strEmpsToRun(), vbNullChar) & vbNullChar, vbNullChar & !idEmpNum & vbNullChar) > 0
            If Found Then
                'make room in array for new data
                SortArray(0, UBound(SortArray, 2)) = ReturnEmpInfo(!idEmpNum).Name 'add data to array
                SortArray(1, UBound(SortArray, 2)) = !idEmpNum
                SortArray(2, UBound(SortArray, 2)) = !idStartDate
                SortArray(3, UBound(SortArray, 2)) = !idEndDate
                SortArray(4, UBound(SortArray, 2)) = !idStatus
                SortArray(5, UBound(SortArray, 2)) = !idStatus2
                SortArray(6, UBound(SortArray, 2)) = !idNotes
                ReDim Preserve SortArray(6, UBound(SortArray, 2) + 1)
                .MoveNext 'move to next position
            Else
                .MoveNext
            End If
        Loop
        .Close
        If UBound(SortArray, 2) = 0 Then GoTo leavesub
        Call MedianThreeQuickSort1(SortArray) 'Modified quick sort that supports multidimentional arrays (sorts by element 0 (name))
        i = 1
        Do
            CurEmp = SortArray(1, i)
            If i + 1 <= UBound(SortArray, 2) Then
                NextEmp = SortArray(1, i + 1)
            End If
            bolHeaderAdded = False
            If CurEmp = NextEmp Then
                FirstEmp = SortArray(1, i)
                If Not bolHeaderAdded Then
                    Grid1.Rows = Grid1.Rows + 1
                    Grid1.CellDetails Grid1.Rows - 1, 1, SortArray(0, i), DT_CENTER, , , , sFntUnder
                    Grid1.CellDetails Grid1.Rows - 1, 2, "Hire Date: " & ReturnEmpInfo(SortArray(1, i)).HireDate, DT_CENTER, , , , sFntUnder
                    Grid1.CellDetails Grid1.Rows - 1, 3, "Hours Avail: " & VacaAvail(SortArray(1, i)), DT_CENTER, , , , sFntUnder
                    bolHeaderAdded = True
                End If
                Do Until FirstEmp <> CurEmp Or i = UBound(SortArray, 2)
                    ReDim Preserve DateSortArray(6, UBound(DateSortArray, 2) + 1) 'make room in array for new data
                    DateSortArray(0, UBound(DateSortArray, 2)) = SortArray(2, i) 'ReturnEmpInfo("NAME", SortArray(1, i)) 'add data to array
                    DateSortArray(1, UBound(DateSortArray, 2)) = SortArray(1, i)
                    DateSortArray(2, UBound(DateSortArray, 2)) = SortArray(0, i) 'ReturnEmpInfo(SortArray(1, i)).Name 'SortArray(2, i)
                    DateSortArray(3, UBound(DateSortArray, 2)) = SortArray(3, i)
                    DateSortArray(4, UBound(DateSortArray, 2)) = SortArray(4, i)
                    DateSortArray(5, UBound(DateSortArray, 2)) = SortArray(5, i)
                    DateSortArray(6, UBound(DateSortArray, 2)) = SortArray(6, i)
                    i = i + 1
                    If i = UBound(SortArray, 2) And FirstEmp = SortArray(1, i) Then
                        ReDim Preserve DateSortArray(6, UBound(DateSortArray, 2) + 1) 'make room in array for new data
                        DateSortArray(0, UBound(DateSortArray, 2)) = SortArray(2, i) 'ReturnEmpInfo("NAME", SortArray(1, i)) 'add data to array
                        DateSortArray(1, UBound(DateSortArray, 2)) = SortArray(1, i)
                        DateSortArray(2, UBound(DateSortArray, 2)) = SortArray(0, i) 'ReturnEmpInfo(SortArray(1, i)).Name 'SortArray(2, i)
                        DateSortArray(3, UBound(DateSortArray, 2)) = SortArray(3, i)
                        DateSortArray(4, UBound(DateSortArray, 2)) = SortArray(4, i)
                        DateSortArray(5, UBound(DateSortArray, 2)) = SortArray(5, i)
                        DateSortArray(6, UBound(DateSortArray, 2)) = SortArray(6, i)
                        bolFinalEntry = True
                    Else
                        CurEmp = SortArray(1, i)
                    End If
                Loop
                If UBound(DateSortArray, 2) > 25 Then
                    Call MedianThreeQuickSort1(DateSortArray) 'Modified quick sort that supports multidimentional arrays (sorts by element 0 (name))
                Else
                    Call MySort(DateSortArray) 'Custom Selection Sort (Works faster on smaller sorts than the quicksort)
                End If
                For b = 1 To UBound(DateSortArray, 2)
                    Grid1.Rows = Grid1.Rows + 1
                    Grid1.CellDetails Grid1.Rows - 1, 1, DateSortArray(0, b), DT_CENTER 'StartDate
                    Grid1.CellDetails Grid1.Rows - 1, 2, DateSortArray(3, b), DT_CENTER 'EndDate
                    clStartDate = DateSortArray(0, b)
                    clEndDate = DateSortArray(3, b)
                    Grid1.CellDetails Grid1.Rows - 1, 3, DateDiffW(clStartDate, clEndDate) * 8 & " Hours", DT_CENTER
                    Grid1.CellDetails Grid1.Rows - 1, 4, DateSortArray(4, b), DT_CENTER 'Status
                    Grid1.CellDetails Grid1.Rows - 1, 5, DateSortArray(5, b), DT_CENTER 'IsPaid?
                    Grid1.CellDetails Grid1.Rows - 1, 6, DateSortArray(6, b), DT_CENTER 'Notes
                Next b
                'Erase DateSortArray
                ReDim DateSortArray(6, 0)
                Grid1.Rows = Grid1.Rows + 1
            Else
                If Not bolHeaderAdded Then
                    Grid1.Rows = Grid1.Rows + 1
                    Grid1.CellDetails Grid1.Rows - 1, 1, SortArray(0, i), DT_CENTER, , , , sFntUnder
                    Grid1.CellDetails Grid1.Rows - 1, 2, "Hire Date: " & ReturnEmpInfo(SortArray(1, i)).HireDate, DT_CENTER, , , , sFntUnder
                    Grid1.CellDetails Grid1.Rows - 1, 3, "Hours Avail: " & VacaAvail(SortArray(1, i)), DT_CENTER, , , , sFntUnder
                    bolHeaderAdded = True
                End If
                Grid1.Rows = Grid1.Rows + 1
                Grid1.CellDetails Grid1.Rows - 1, 1, SortArray(2, i), DT_CENTER 'StartDate
                Grid1.CellDetails Grid1.Rows - 1, 2, SortArray(3, i), DT_CENTER 'EndDate
                clStartDate = SortArray(2, i)
                clEndDate = SortArray(3, i)
                Grid1.CellDetails Grid1.Rows - 1, 3, DateDiffW(clStartDate, clEndDate) * 8 & " Hours", DT_CENTER
                Grid1.CellDetails Grid1.Rows - 1, 4, SortArray(4, i), DT_CENTER 'Status
                Grid1.CellDetails Grid1.Rows - 1, 5, SortArray(5, i), DT_CENTER 'IsPaid?
                Grid1.CellDetails Grid1.Rows - 1, 6, SortArray(6, i), DT_CENTER 'Notes
                Grid1.Rows = Grid1.Rows + 1
                i = i + 1
            End If
            If i = UBound(SortArray, 2) And Not bolFinalEntry Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.CellDetails Grid1.Rows - 1, 1, SortArray(0, i), DT_CENTER, , , , sFntUnder
                Grid1.CellDetails Grid1.Rows - 1, 2, "Hire Date: " & ReturnEmpInfo(SortArray(1, i)).HireDate, DT_CENTER, , , , sFntUnder
                Grid1.CellDetails Grid1.Rows - 1, 3, "Hours Avail: " & VacaAvail(SortArray(1, i)), DT_CENTER, , , , sFntUnder
                Grid1.Rows = Grid1.Rows + 1
                Grid1.CellDetails Grid1.Rows - 1, 1, SortArray(2, i), DT_CENTER 'StartDate
                Grid1.CellDetails Grid1.Rows - 1, 2, SortArray(3, i), DT_CENTER 'EndDate
                clStartDate = SortArray(2, i)
                clEndDate = SortArray(3, i)
                Grid1.CellDetails Grid1.Rows - 1, 3, DateDiffW(clStartDate, clEndDate) * 8 & " Hours", DT_CENTER
                Grid1.CellDetails Grid1.Rows - 1, 4, SortArray(4, i), DT_CENTER 'Status
                Grid1.CellDetails Grid1.Rows - 1, 5, SortArray(5, i), DT_CENTER 'IsPaid?
                Grid1.CellDetails Grid1.Rows - 1, 6, SortArray(6, i), DT_CENTER 'Notes
                bolHeaderAdded = True
            End If
        Loop Until i >= UBound(SortArray, 2)
        Grid1.Rows = Grid1.Rows + 1
leavesub:
        Erase SortArray
        ReDim SortArray(6, 0)
    End With
    ' Grid1.Redraw = True
    strReportTitle = "Vacations between " & StartDate & " and " & EndDate
    strReportSubTitle = IIf(chkBremen, "Bremen ", "") & IIf(chkWooster, " Wooster ", "") & IIf(chkRockyMtn, " RockyMtn ", "") & IIf(chkNuclear, " Nuclear ", "")
End Sub
Private Function VacaAvailAccrual(EmpNum) As Integer
    Dim rs            As New ADODB.Recordset
    Dim cn            As New ADODB.Connection
    Dim strSQL1       As String
    Dim intTakenWeeks As Integer
    Dim DTStartDate   As Date, DTEndDate As Date
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    VacaAvailAccrual = 0
    intTakenWeeks = 0
    DTStartDate = dtAnniDate(ReturnEmpInfo(EmpNum).HireDate).PreviousYear
    DTEndDate = dtFiscalYearEnd
    strSQL1 = "SELECT * " & "FROM attendb.vacations vacations_0" & " WHERE (vacations_0.idStartDate>={d '" & Format$(DTStartDate, strDBDateFormat) & "'}) AND (vacations_0.idEndDate<={d '" & Format$(DTEndDate, strDBDateFormat) & "'}) AND (vacations_0.idEmpNum='" & EmpNum & "') AND (vacations_0.idStatus='TAKEN') AND (vacations_0.idStatus2='PAID')"
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset
    With rs
        If .RecordCount < 1 Then
            VacaAvailAccrual = 0
            If ReturnEmpInfo(EmpNum).VacaWeeks <> 0 Then
                VacaAvailAccrual = ReturnEmpInfo(EmpNum).VacaWeeks
            Else
                VacaAvailAccrual = CalcYearsWorked(EmpNum).VacaWeeksAvail
            End If
            Exit Function
        Else
            intTakenWeeks = .RecordCount
        End If
    End With
    If ReturnEmpInfo(EmpNum).VacaWeeks <> 0 Then
        VacaAvailAccrual = ReturnEmpInfo(EmpNum).VacaWeeks - intTakenWeeks
    Else
        VacaAvailAccrual = CalcYearsWorked(EmpNum).VacaWeeksAvail - intTakenWeeks
    End If
End Function
Private Function VacaAvail(EmpNum) As Integer
    Dim rs            As New ADODB.Recordset
    Dim cn            As New ADODB.Connection
    Dim strSQL1       As String
    Dim intTakenWeeks As Integer
    Dim lngHoursTaken As Long
    Dim intDays       As Integer
    Dim DTStartDate   As Date, DTEndDate As Date
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    VacaAvail = 0
    intTakenWeeks = 0
    DTStartDate = dtVacaPeriod.StartDate 'dtAnniDate(ReturnEmpInfo(EmpNum).HireDate).PreviousYear
    DTEndDate = dtVacaPeriod.EndDate 'dtAnniDate(ReturnEmpInfo(EmpNum).HireDate).CurrentYear
    strSQL1 = "SELECT * " & "FROM attendb.vacations vacations_0" & " WHERE (vacations_0.idStartDate>={d '" & Format$(DTStartDate, strDBDateFormat) & "'}) AND (vacations_0.idEndDate<={d '" & Format$(DTEndDate, strDBDateFormat) & "'}) AND (vacations_0.idEmpNum='" & EmpNum & "') AND (vacations_0.idStatus='TAKEN') AND (vacations_0.idStatus2='PAID')"
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    rs.Open strSQL1, cn, adOpenKeyset
    With rs
        If .RecordCount < 1 Then
            VacaAvail = 0
            If ReturnEmpInfo(EmpNum).VacaHours <> 0 Then
                VacaAvail = strCurrentEmpInfo.VacaHours
            Else
                VacaAvail = CalcYearsWorked(EmpNum).VacaHoursAvail
            End If
            Exit Function
        Else
            Do Until rs.EOF
                intDays = DateDiffW(!idStartDate, !idEndDate) '+ 1
                lngHoursTaken = lngHoursTaken + intDays * 8
                .MoveNext
            Loop
        End If
    End With
    If ReturnEmpInfo(EmpNum).VacaHours <> 0 Then
        VacaAvail = ReturnEmpInfo(EmpNum).VacaHours - lngHoursTaken
    Else
        VacaAvail = CalcYearsWorked(EmpNum).VacaHoursAvail - lngHoursTaken
    End If
End Function
Public Sub MySort(ByRef pvarArray As Variant)
    Dim i               As Long
    Dim c               As Integer
    Dim v               As Integer
    Dim lngHighValIndex As Long
    Dim varSwap()       As Variant
    Dim lngMax          As Long
    ReDim varSwap(UBound(pvarArray, 1))
    lngMax = UBound(pvarArray, 2)
    For c = 0 To lngMax
        lngHighValIndex = lngMax - c
        For v = 0 To UBound(varSwap)
            varSwap(v) = pvarArray(v, lngMax - c)
        Next v
        For i = 0 To lngMax - c
            If pvarArray(0, i) > pvarArray(0, lngHighValIndex) Then lngHighValIndex = i
        Next
        For v = 0 To UBound(varSwap)
            pvarArray(v, lngMax - c) = pvarArray(v, lngHighValIndex)
            pvarArray(v, lngHighValIndex) = varSwap(v)
        Next v
    Next c
End Sub
Public Sub MedianThreeQuickSort1(ByRef pvarArray As Variant, _
                                 Optional ByVal plngLeft As Long, _
                                 Optional ByVal plngRight As Long)
    Dim lngFirst  As Long
    Dim lngLast   As Long
    Dim varMid    As Variant
    Dim lngIndex  As Long
    Dim varSwap() As Variant
    Dim A         As Long
    Dim b         As Long
    Dim c         As Long
    Dim i         As Integer
    ReDim varSwap(UBound(pvarArray, 1))
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray, 2)
        plngRight = UBound(pvarArray, 2)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    lngIndex = plngRight - plngLeft + 1
    A = Int(lngIndex * Rnd) + plngLeft
    b = Int(lngIndex * Rnd) + plngLeft
    c = Int(lngIndex * Rnd) + plngLeft
    If pvarArray(0, A) <= pvarArray(0, b) And pvarArray(0, b) <= pvarArray(0, c) Then
        lngIndex = b
    Else
        If pvarArray(0, b) <= pvarArray(0, A) And pvarArray(0, A) <= pvarArray(0, c) Then
            lngIndex = A
        Else
            lngIndex = c
        End If
    End If
    varMid = pvarArray(0, lngIndex)
    Do
        Do While pvarArray(0, lngFirst) < varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid < pvarArray(0, lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            For i = 0 To UBound(varSwap)
                varSwap(i) = pvarArray(i, lngFirst)
                pvarArray(i, lngFirst) = pvarArray(i, lngLast)
                pvarArray(i, lngLast) = varSwap(i)
            Next i
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If lngLast - plngLeft < plngRight - lngFirst Then
        If plngLeft < lngLast Then MedianThreeQuickSort1 pvarArray, plngLeft, lngLast
        If lngFirst < plngRight Then MedianThreeQuickSort1 pvarArray, lngFirst, plngRight
    Else
        If lngFirst < plngRight Then MedianThreeQuickSort1 pvarArray, lngFirst, plngRight
        If plngLeft < lngLast Then MedianThreeQuickSort1 pvarArray, plngLeft, lngLast
    End If
End Sub
'WhosOnVaca determines which emps will be on vacation by month, ordered desc by week, and alphebetically by name (This was a tough one!)
Private Sub WhosOnVaca(Month As Integer)
    Dim rs              As New ADODB.Recordset
    Dim cn              As New ADODB.Connection
    Dim strSQL1         As String
    Dim Found           As Boolean
    Dim bolMonthAdded   As Boolean
    Dim CurDate         As Date, NextDate As Date, FirstDate As Date
    Dim intStartingRows As Integer
    intStartingRows = Grid1.Rows
    Dim SortArray() As Variant
    ReDim SortArray(4, 0)
    Dim i         As Integer
    Dim sFntUnder As New StdFont
    sFntUnder.Underline = True
    sFntUnder.Name = "Tahoma"
    Dim sFntNormal As New StdFont
    sFntNormal.Underline = False
    sFntNormal.Name = "Tahoma"
    bolMonthAdded = False
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    cn.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=attendb;dsn=;"
    cn.CursorLocation = adUseClient
    strSQL1 = "SELECT * " & "FROM attendb.vacations vacations_0" & " WHERE (vacations_0.idStartDate>={d '" & Format$(GetMonthDates(Month, DateTime.Year(Now)).StartDate, strDBDateFormat) & "'}) AND (vacations_0.idStartDate<={d '" & Format$(GetMonthDates(Month, DateTime.Year(Now)).EndDate, strDBDateFormat) & "'}) " & IIf(chkRequested.Value = 0, "AND vacations_0.idStatus <> 'REQUESTED'", "") & IIf(chkReScheduled.Value = 0, "AND vacations_0.idStatus <> 'RESCHEDULED'", "") & IIf(chkTaken.Value = 0, "AND vacations_0.idStatus <> 'TAKEN'", "")
    rs.Open strSQL1, cn, adOpenKeyset
    With rs
        If rs.RecordCount = 0 Then Exit Sub
        Do Until .EOF
            Found = InStr(1, vbNullChar & Join(strEmpsToRun(), vbNullChar) & vbNullChar, vbNullChar & !idEmpNum & vbNullChar) > 0
            If Found Then
                If Not bolMonthAdded Then
                    Grid1.Rows = Grid1.Rows + 1
                    Grid1.CellDetails Grid1.Rows - 1, 1, lstMonths.List(Month - 1), DT_CENTER, , , , sFntUnder
                    bolMonthAdded = True
                End If
                CurDate = !idStartDate 'set to current position date
                If .AbsolutePosition + 1 <= .RecordCount Then 'if not on last position, get date of next position
                    .MoveNext
                    NextDate = !idStartDate
                    .MovePrevious
                Else
                End If
                If CurDate = NextDate Then 'if the current position date is the same as the next one, add it to an array to be sorted
                    FirstDate = !idStartDate 'get date of initial match
                    Do Until FirstDate <> CurDate Or .EOF ' loop through positions until we get to the end, or we find one that doesnt match
                        Found = InStr(1, vbNullChar & Join(strEmpsToRun(), vbNullChar) & vbNullChar, vbNullChar & !idEmpNum & vbNullChar) > 0
                        If Not Found Then GoTo filteremp
                        SortArray(0, UBound(SortArray, 2)) = ReturnEmpInfo(!idEmpNum).Name 'add data to array
                        SortArray(1, UBound(SortArray, 2)) = !idStartDate
                        SortArray(2, UBound(SortArray, 2)) = !idEndDate
                        SortArray(3, UBound(SortArray, 2)) = !idStatus2
                        SortArray(4, UBound(SortArray, 2)) = !idNotes
                        ReDim Preserve SortArray(4, UBound(SortArray, 2) + 1) 'make room in array for new data
filteremp:
                        .MoveNext 'move to next position
                        If .EOF Then
                            GoTo printarray 'if at the end of dataset, add what data we collected to the flexgrid, after sorting the array
                        Else
                            CurDate = !idStartDate 'otherwise set to the current position date
                        End If
                    Loop
printarray:
                    If UBound(SortArray, 2) > 25 Then
                        Call MedianThreeQuickSort1(SortArray) 'Modified quick sort that supports multidimentional arrays (sorts by element 0 (name))
                    Else
                        Call MySort(SortArray)
                    End If
                    For i = 1 To UBound(SortArray, 2) 'cycle through the now sorted array and add it to the grid
                        Grid1.Rows = Grid1.Rows + 1
                        Grid1.CellDetails Grid1.Rows - 1, 1, SortArray(0, i), DT_CENTER
                        Grid1.CellDetails Grid1.Rows - 1, 2, SortArray(1, i), DT_CENTER
                        Grid1.CellDetails Grid1.Rows - 1, 3, SortArray(2, i), DT_CENTER
                        Grid1.CellDetails Grid1.Rows - 1, 4, SortArray(3, i), DT_CENTER
                        Grid1.CellDetails Grid1.Rows - 1, 5, SortArray(4, i), DT_CENTER
                    Next i
                    Grid1.Rows = Grid1.Rows + 1
                    'Erase SortArray 'erase that shit
                    ReDim SortArray(4, 0) 'make it an array again
                Else 'if the current date and next dates dont match, add them directly to the grid
                    Grid1.Rows = Grid1.Rows + 1
                    Grid1.CellDetails Grid1.Rows - 1, 1, ReturnEmpInfo(!idEmpNum).Name, DT_CENTER
                    Grid1.CellDetails Grid1.Rows - 1, 2, !idStartDate, DT_CENTER
                    Grid1.CellDetails Grid1.Rows - 1, 3, !idEndDate, DT_CENTER
                    Grid1.CellDetails Grid1.Rows - 1, 4, !idStatus2, DT_CENTER
                    Grid1.CellDetails Grid1.Rows - 1, 5, !idNotes, DT_CENTER
                    Grid1.Rows = Grid1.Rows + 1
                    .MoveNext
                    If .EOF Then 'if we can set the current date, or leave the loop
                        GoTo getout
                    Else
                        CurDate = !idStartDate
                    End If
                End If
            Else
                .MoveNext
            End If
        Loop
getout:
        .Close
    End With
    If intStartingRows < Grid1.Rows Then 'if we added rows to the grid, put a blank row at the end for easier reading
        Grid1.Rows = Grid1.Rows + 1
    Else
    End If
    bolMonthAdded = False
End Sub
Private Sub cmdToExcel_Click()
    Dim XcLApp As Object  'used for excel application'
    Dim XcLWB  As Object 'used for excel work book'
    Dim XcLWS  As Object 'used for excel work sheet'
    Dim i      As Integer, c As Integer ' counter for the rows of the flexgrid'
    Set XcLApp = CreateObject("Excel.Application")  'creating new excel application'
    Set XcLWB = XcLApp.Workbooks.Add                'opening new excel work book'
    Set XcLWS = XcLWB.Worksheets.Add                'opening new excel worksheet'
    'taking data from flexgrid and sendting it to excel'
    With Grid1
        For i = 1 To .Rows - 1
            For c = 1 To .Columns
                XcLWS.Range(Addres_Excel(i, c)).Value = .Cell(i, c).Text
            Next c
        Next i
    End With
    XcLApp.Visible = True
End Sub
Private Sub cmdWhosOnVaca_Click()
    Dim intMonths() As Integer
    ReDim intMonths(1)
    Dim i As Integer
    Grid1.BackColor = colGridBusy
    Screen.MousePointer = vbHourglass
    Grid1.Redraw = False
    Grid1.Clear
    Grid1.Rows = 1
    For i = 0 To lstMonths.ListCount - 1
        If lstMonths.Selected(i) = True Then
            intMonths(UBound(intMonths)) = i + 1
            ReDim Preserve intMonths(UBound(intMonths) + 1)
        End If
    Next
    EmpListByCompany
    For i = 1 To UBound(intMonths) - 1
        Call WhosOnVaca(intMonths(i))
    Next i
    ReSizeSGrid
    Grid1.Redraw = True
    Screen.MousePointer = vbDefault
    Grid1.BackColor = &H80000005
    strReportTitle = "Who's On Vacation"
    strReportSubTitle = IIf(chkBremen, "Bremen ", "") & IIf(chkWooster, " Wooster ", "") & IIf(chkRockyMtn, " RockyMtn ", "") & IIf(chkNuclear, " Nuclear ", "")
End Sub
Private Sub EmpListByCompany()
    ReDim strEmpsToRun(0)
    Dim i As Integer
    For i = 0 To UBound(strEmpInfo)
        If strEmpInfo(i).IsActive = True Then
            If chkBremen And strEmpInfo(i).Location1 = "BREMEN" Then
                ReDim Preserve strEmpsToRun(UBound(strEmpsToRun) + 1)
                strEmpsToRun(UBound(strEmpsToRun)) = strEmpInfo(i).Number
            End If
            If chkNuclear And strEmpInfo(i).Location1 = "NUCLEAR" Then
                ReDim Preserve strEmpsToRun(UBound(strEmpsToRun) + 1)
                strEmpsToRun(UBound(strEmpsToRun)) = strEmpInfo(i).Number
            End If
            If chkRockyMtn And strEmpInfo(i).Location1 = "ROCKY MTN" Then
                ReDim Preserve strEmpsToRun(UBound(strEmpsToRun) + 1)
                strEmpsToRun(UBound(strEmpsToRun)) = strEmpInfo(i).Number
            End If
            If chkWooster And strEmpInfo(i).Location1 = "WOOSTER" Then
                ReDim Preserve strEmpsToRun(UBound(strEmpsToRun) + 1)
                strEmpsToRun(UBound(strEmpsToRun)) = strEmpInfo(i).Number
            End If
        End If
    Next i
End Sub
Private Sub cmdDateRange_Click()
    Grid1.BackColor = colGridBusy
    Screen.MousePointer = vbHourglass
    Grid1.Redraw = False
    Call CustomDateRange(DTStartDate, DTEndDate)
    ReSizeSGrid
    Grid1.Redraw = True
    Screen.MousePointer = vbDefault
    Grid1.BackColor = &H80000005
End Sub
Private Sub cmdPrint_Click()
    frmPrinters.Show 1
    If bolCancelPrint = True Then
        bolCancelPrint = False
        Exit Sub
    End If
    Grid1.Redraw = False
    frmPBar.Show
    DoEvents
    ReSizeSGrid
    PrintFlexGridSGrid Grid1, strReportTitle, strReportSubTitle
    Grid1.Redraw = True
End Sub
Private Sub Form_Load()
    lstMonths.AddItem "January", 0
    lstMonths.AddItem "Feburary", 1
    lstMonths.AddItem "March", 2
    lstMonths.AddItem "April", 3
    lstMonths.AddItem "May", 4
    lstMonths.AddItem "June", 5
    lstMonths.AddItem "July", 6
    lstMonths.AddItem "August", 7
    lstMonths.AddItem "September", 8
    lstMonths.AddItem "October", 9
    lstMonths.AddItem "November", 10
    lstMonths.AddItem "December", 11
    DTStartDate.Value = Date
    DTEndDate.Value = Date
    lblAccrualRange.Caption = "Prev. Anni. Date to " & dtFiscalYearEnd
    Grid1.AddColumn "0"
    Grid1.AddColumn "1"
    Grid1.AddColumn "2"
    Grid1.AddColumn "3"
    Grid1.AddColumn "4"
    Grid1.AddColumn "5"
    'Grid1.AddColumn "6"
    ' Grid1.AddColumn "7"
    Grid1.Header = False
    Grid1.Gridlines = True
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
