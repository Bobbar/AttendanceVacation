VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAddNewEmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Employee"
   ClientHeight    =   2445
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddNewEmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTHireDate 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   209387521
      CurrentDate     =   40934
   End
   Begin VB.ComboBox cmbLocation2 
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
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1500
      Width           =   2175
   End
   Begin VB.ComboBox cmbLocation1 
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
      TabIndex        =   4
      Top             =   1500
      Width           =   2235
   End
   Begin VB.TextBox txtLastName 
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
      Left            =   2880
      TabIndex        =   3
      Text            =   "LastName"
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtFirstName 
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
      Left            =   240
      TabIndex        =   2
      Text            =   "FirstName"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location 2"
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location 1"
      Height          =   195
      Left            =   300
      TabIndex        =   12
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hire Date"
      Height          =   195
      Left            =   5280
      TabIndex        =   11
      Top             =   480
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   195
      Left            =   2880
      TabIndex        =   10
      Top             =   480
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   765
   End
   Begin VB.Label lblEmpNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3660
      TabIndex        =   8
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblEmployeeNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1980
      TabIndex        =   7
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmAddNewEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub CancelButton_Click()
    Form1.ClearFields
    frmAddNewEmp.Hide
End Sub
Private Sub cmdAddNew_Click()
    Dim rs         As New ADODB.Recordset
    Dim cn         As New ADODB.Connection
    Dim strEmpName As String, FormatDate As String
    Dim blah
    If txtFirstName.Text = "" Or txtLastName.Text = "" Or cmbLocation1.Text = "" Or cmbLocation2.Text = "" Then
        MsgBox "One or more fields are empty!  Please fill all fields.", vbOKOnly + vbCritical, "Missing information"
        Exit Sub
    Else
    End If
    txtFirstName.Text = Trim$(UCase$(txtFirstName.Text))
    txtLastName.Text = Trim$(UCase$(txtLastName.Text))
    strEmpName = txtLastName.Text & "," & txtFirstName.Text
    If CheckForName(strEmpName).Exist Then
    blah = MsgBox(strEmpName & " is already in the database under Employee #" & CheckForName(strEmpName).Number, vbExclamation + vbOKOnly, "Duplicate entry")
    Exit Sub
    End If
    
    
    FormatDate = Format$(DTHireDate.Value, strDBDateFormat)
    Form1.AddEmpToDB strEmpName, Form1.txtAttenEmpNum.Text, FormatDate, cmbLocation1.Text, cmbLocation2.Text, "TRUE"
    blah = MsgBox("New employee added to database.", vbOKOnly, "Done")
    frmAddNewEmp.Hide
    GetEmpInfo
    Form1.GetEntries
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    txtFirstName.Text = "FirstName"
    txtLastName.Text = "LastName"
    cmbLocation1.ListIndex = 0
    cmbLocation2.ListIndex = 0
    DTHireDate.Value = DateTime.Date
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
    Else
    End If
    frmAddNewEmp.Hide
End Sub
Private Sub txtFirstName_GotFocus()
    With txtFirstName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtLastName_GotFocus()
    With txtLastName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
