VERSION 5.00
Begin VB.Form frmAlarm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Warning"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlarm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Alarms"
      Height          =   2835
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   4995
      Begin VB.CheckBox chkAckPartialUn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Caption         =   "Acknowledge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CheckBox chkAckFullUn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Caption         =   "Acknowledge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CheckBox chkAckPartialEx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Caption         =   "Acknowledge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1620
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CheckBox chkAckFullEx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Caption         =   "Acknowledge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partial Unexcused: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1980
         TabIndex        =   15
         Top             =   360
         Width           =   2040
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Unexcused:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   900
         Width           =   2040
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Partial Excused:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2115
         TabIndex        =   13
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Excused:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2280
         TabIndex        =   12
         Top             =   2220
         Width           =   1260
      End
      Begin VB.Label lblPartialUn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3960
         TabIndex        =   11
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblFullUn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3960
         TabIndex        =   10
         Top             =   900
         Width           =   600
      End
      Begin VB.Label lblPartialEx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3840
         TabIndex        =   9
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblFullEx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3840
         TabIndex        =   8
         Top             =   2220
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Name"
      Height          =   735
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4995
      Begin VB.Label lblEmpName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EmpName"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1740
      TabIndex        =   0
      Top             =   3780
      Width           =   1635
   End
   Begin VB.Timer tmrFlasher 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   180
      Top             =   3840
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sndPlaySound _
                Lib "WINMM.DLL" _
                Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                       ByVal uFlags As Long) As Long
Const SND_PURGE = &H40
Private Sub GetAlarms(EmpNum As Integer)
    chkAckPartialUn.Value = GetSetting(App.EXEName, EmpNum, "PartialUn", 0)
    chkAckFullUn.Value = GetSetting(App.EXEName, EmpNum, "FullUn", 0)
    chkAckPartialEx.Value = GetSetting(App.EXEName, EmpNum, "PartialEx", 0)
    chkAckFullEx.Value = GetSetting(App.EXEName, EmpNum, "FullEx", 0)
End Sub
Private Sub chkAckFullEx_Click()
    With chkAckFullEx
        If .Value = 0 Then
            .BackColor = vbRed
            .Caption = "Acknowledge"
            SaveSetting App.EXEName, Form1.txtAttenEmpNum.Text, "FullEx", .Value
        ElseIf .Value = 1 Then
            .BackColor = vbGreen
            .Caption = "Acknowledged"
            SaveSetting App.EXEName, Form1.txtAttenEmpNum.Text, "FullEx", .Value
        End If
    End With
End Sub
Private Sub chkAckFullUn_Click()
    With chkAckFullUn
        If .Value = 0 Then
            .BackColor = vbRed
            .Caption = "Acknowledge"
            SaveSetting App.EXEName, Form1.txtAttenEmpNum.Text, "FullUn", .Value
        ElseIf .Value = 1 Then
            .BackColor = vbGreen
            .Caption = "Acknowledged"
            SaveSetting App.EXEName, Form1.txtAttenEmpNum.Text, "FullUn", .Value
        End If
    End With
End Sub
Private Sub chkAckPartialEx_Click()
    With chkAckPartialEx
        If .Value = 0 Then
            .BackColor = vbRed
            .Caption = "Acknowledge"
            SaveSetting App.EXEName, Form1.txtAttenEmpNum.Text, "PartialEx", .Value
        ElseIf .Value = 1 Then
            .BackColor = vbGreen
            .Caption = "Acknowledged"
            SaveSetting App.EXEName, Form1.txtAttenEmpNum.Text, "PartialEx", .Value
        End If
    End With
End Sub
Private Sub chkAckPartialUn_Click()
    With chkAckPartialUn
        If .Value = 0 Then
            .BackColor = vbRed
            .Caption = "Acknowledge"
            SaveSetting App.EXEName, Form1.txtAttenEmpNum.Text, "PartialUn", .Value
        ElseIf .Value = 1 Then
            .BackColor = vbGreen
            .Caption = "Acknowledged"
            SaveSetting App.EXEName, Form1.txtAttenEmpNum.Text, "PartialUn", .Value
        End If
    End With
End Sub
Private Sub cmdOk_Click()
    bolAlarmOKed = True
    Call Form1.SetAckLabel(Form1.txtAttenEmpNum.Text)
    sndPlaySound vbNullString, SND_PURGE 'Stop sound
    frmAlarm.Hide
End Sub
Private Sub Form_Activate()
    lblEmpName.Caption = strAlarmTitleString
    cmdOk.Visible = True
    chkAckPartialUn.Visible = False
    chkAckFullUn.Visible = False
    chkAckPartialEx.Visible = False
    chkAckFullEx.Visible = False
    If bolPartialUnExcusedExceeded = True Then
        lblPartialUn.ForeColor = &HFF&
        lblPartialUn.Caption = intPartialUnExcused & " of " & intPartialUnExcusedAllowed
    Else
        lblPartialUn.ForeColor = vbBlack
        lblPartialUn.Caption = intPartialUnExcused & " of " & intPartialUnExcusedAllowed
    End If
    If bolFullUnExcusedExceeded = True Then
        lblFullUn.ForeColor = &HFF&
        lblFullUn.Caption = intFullUnExcused & " of " & intFullUnExcusedAllowed
    Else
        lblFullUn.ForeColor = vbBlack
        lblFullUn.Caption = intFullUnExcused & " of " & intFullUnExcusedAllowed
    End If
    If bolFullExcusedExceeded = True Then
        lblFullEx.ForeColor = &HFF&
        lblFullEx.Caption = intFullExcused & " of " & intFullExcusedAllowed
    Else
        lblFullEx.ForeColor = vbBlack
        lblFullEx.Caption = intFullExcused & " of " & intFullExcusedAllowed
    End If
    If bolPartialExcusedExceeded = True Then
        lblPartialEx.ForeColor = &HFF&
        lblPartialEx.Caption = intPartialExcused & " of " & intPartialExcusedAllowed
    Else
        lblPartialEx.ForeColor = vbBlack
        lblPartialEx.Caption = intPartialExcused & " of " & intPartialExcusedAllowed
    End If
    GetAlarms Form1.txtAttenEmpNum
    If bolAlarmsCalled = True Then
        frmAlarm.BackColor = &H8000000F
        tmrFlasher.Enabled = False
        Flashes = 0
        cmdOk.Visible = True
        If bolPartialUnExcusedExceeded = True Then
            chkAckPartialUn.Visible = True
        Else
            chkAckPartialUn.Visible = False
        End If
        If bolFullUnExcusedExceeded = True Then
            chkAckFullUn.Visible = True
        Else
            chkAckFullUn.Visible = False
        End If
        If bolPartialExcusedExceeded = True Then
            chkAckPartialEx.Visible = True
        Else
            chkAckPartialEx.Visible = False
        End If
        If bolFullExcusedExceeded = True Then
            chkAckFullEx.Visible = True
        Else
            chkAckFullEx.Visible = False
        End If
        'frmAlarm.Visible = True
    ElseIf bolAlarmsCalled = False Then
        tmrFlasher.Enabled = True
        Flashes = 0
    End If
End Sub
Private Sub tmrFlasher_Timer()
    Dim MaxFlashes As Integer
    MaxFlashes = 10
    If frmAlarm.BackColor = &H8000000F And Flashes < MaxFlashes Then
        frmAlarm.BackColor = &HFF&
        Frame1.BackColor = &HFF&
        Frame2.BackColor = &HFF&
    ElseIf frmAlarm.BackColor = &HFF& And Flashes < MaxFlashes Then
        frmAlarm.BackColor = &H8000000F
        Frame1.BackColor = &H8000000F
        Frame2.BackColor = &H8000000F
    End If
    If Flashes >= MaxFlashes Then
        frmAlarm.BackColor = &H8000000F
        Frame1.BackColor = &H8000000F
        Frame2.BackColor = &H8000000F
        tmrFlasher.Enabled = False
        Flashes = 0
        cmdOk.Visible = True
        chkAckPartialUn.Visible = bolPartialUnExcusedExceeded
        chkAckFullUn.Visible = bolFullUnExcusedExceeded
        chkAckPartialEx.Visible = bolPartialExcusedExceeded
        chkAckFullEx.Visible = bolFullExcusedExceeded
    End If
    Flashes = Flashes + 1
End Sub
