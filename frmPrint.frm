VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Filter Criteria"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Frame Frame2 
         Caption         =   "Date Range"
         Height          =   1455
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   5055
         Begin VB.CheckBox chkAllDate 
            Caption         =   "All"
            Height          =   255
            Left            =   2280
            TabIndex        =   11
            Top             =   1080
            Value           =   1  'Checked
            Width           =   855
         End
         Begin MSComCtl2.DTPicker DTEnd 
            Height          =   375
            Left            =   3000
            TabIndex        =   6
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
            CustomFormat    =   "MM-dd-yyyy"
            Format          =   197722113
            CurrentDate     =   40487
         End
         Begin MSComCtl2.DTPicker DTStart 
            Height          =   375
            Left            =   240
            TabIndex        =   7
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
            CustomFormat    =   "MM-dd-yyyy"
            Format          =   197722113
            CurrentDate     =   40487
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
            Height          =   480
            Left            =   2265
            TabIndex        =   10
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ending Date:"
            Height          =   195
            Left            =   3360
            TabIndex        =   9
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Date:"
            Height          =   195
            Left            =   600
            TabIndex        =   8
            Top             =   360
            Width           =   1020
         End
      End
      Begin VB.CheckBox chkExcused 
         Caption         =   "Excused"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   2040
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkUnexcused 
         Caption         =   "Unexcused"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   2760
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Apply Filter"
         Height          =   480
         Left            =   2400
         TabIndex        =   1
         Top             =   2400
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkAll.Value = 0 Then

chkExcused.Value = 1
chkExcused.Enabled = False

chkUnexcused.Value = 1
chkUnexcused.Enabled = False

Else

chkExcused.Value = 0
chkExcused.Enabled = True

chkUnexcused.Value = 0
chkUnexcused.Enabled = True
End If
End Sub

Private Sub chkAllDate_Click()
If chkAllDate.Value = 0 Then
DTStart.Enabled = True
DTEnd.Enabled = True
Else

DTStart.Enabled = False
DTEnd.Enabled = False
End If

End Sub

Private Sub cmdExecute_Click()

Form1.DateRangeReport


frmReport.Hide

End Sub

Private Sub DTEnd_Change()
dtEndDate = DTEnd.Value
End Sub

Private Sub DTStart_Change()
dtStartDate = DTStart.Value
End Sub

