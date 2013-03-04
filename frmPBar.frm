VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPBar 
   BorderStyle     =   0  'None
   Caption         =   "Progress..."
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPBar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin MSComctlLib.ProgressBar PBar1 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   360
         Left            =   2400
         TabIndex        =   1
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label lblQryTime 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%AVG QRY TIME%"
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   2400
         Width           =   5835
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Progress"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   3
         Top             =   1320
         Width           =   5805
      End
   End
End
Attribute VB_Name = "frmPBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    bolStop = True
    DoEvents
    ClearAvgQryTimes
End Sub
