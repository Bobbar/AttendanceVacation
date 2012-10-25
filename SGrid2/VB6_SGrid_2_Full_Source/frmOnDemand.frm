VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmOnDemand 
   Caption         =   "SGrid AddRowsOnDemand Mode Demonstration"
   ClientHeight    =   4005
   ClientLeft      =   3120
   ClientTop       =   2820
   ClientWidth     =   8520
   Icon            =   "frmOnDemand.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   8520
   Begin vbAcceleratorSGrid6.vbalGrid grdVirtual 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6800
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      GridLineColor   =   12632256
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
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VB.Label lblTarget 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   480
      Width           =   1155
   End
End
Attribute VB_Name = "frmOnDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iRowsToAdd As Long

Private Sub cmdClear_Click()
   m_iRowsToAdd = Rnd * 512 + 16
   lblTarget.Caption = m_iRowsToAdd
   grdVirtual.Clear
End Sub

Private Sub Form_Load()
   m_iRowsToAdd = Rnd * 512 + 16
   lblTarget.Caption = m_iRowsToAdd
   With grdVirtual
      .AddColumn , "Test1"
      .AddColumn , "Test2"
      .AddRowsOnDemand = True
   End With
End Sub

Private Sub grdVirtual_RequestRow(ByVal lRow As Long, lItemData As Long, bVisible As Boolean, lHeight As Long, bNoMoreRows As Boolean)
   If (lRow > m_iRowsToAdd) Then
      bNoMoreRows = True
   Else
      lblInfo.Caption = lRow
   End If
End Sub

Private Sub grdVirtual_RequestRowData(ByVal lRow As Long)
   With grdVirtual
      .CellText(lRow, 1) = "Row" & lRow & ";Col1"
      .CellText(lRow, 2) = "Row" & lRow & ";Col2"
   End With
End Sub
