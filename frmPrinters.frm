VERSION 5.00
Begin VB.Form frmPrinters 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Printer"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrinters.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Printers"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Frame Frame2 
         Caption         =   "Justification"
         Height          =   915
         Left            =   3960
         TabIndex        =   4
         Top             =   1260
         Width           =   1695
         Begin VB.OptionButton optCenterJust 
            Caption         =   "Center Justify"
            Height          =   255
            Left            =   180
            TabIndex        =   6
            Top             =   540
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton optLeftJust 
            Caption         =   "Left Justify"
            Height          =   195
            Left            =   180
            TabIndex        =   5
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmbPrinters 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   5175
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Printing"
         Height          =   480
         Left            =   2280
         TabIndex        =   1
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please select a printer:"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   420
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub LoadPrinters()
    Dim i              As Integer
    Dim DefaultPrinter As Integer
    cmbPrinters.Clear
    For i = 0 To Printers.Count - 1
        cmbPrinters.AddItem Printers(i).DeviceName, i
        If Printer.DeviceName = Printers(i).DeviceName Then DefaultPrinter = i
    Next
    cmbPrinters.ListIndex = DefaultPrinter
End Sub
Private Sub cmbPrinters_Click()
    Set Printer = Printers(cmbPrinters.ListIndex)
End Sub
Private Sub cmdStart_Click()
    If cmbPrinters.Text <> "" Then
        bolCancelPrint = False
        frmPrinters.Hide
    Else
        MsgBox "Please select a printer."
    End If
End Sub
Private Sub Form_Activate()
    LoadPrinters
    GetSettings
End Sub
Private Sub Form_Load()
    LoadPrinters
    GetSettings
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolCancelPrint = True
End Sub
Private Sub GetSettings()
    On Error Resume Next
    Dim bolLeft As Variant, bolCenter As Variant
    bolLeft = GetSetting(App.EXEName, "PrinterJustify", "Left", 0)
    bolCenter = GetSetting(App.EXEName, "PrinterJustify", "Center", 0)
    If bolLeft = bolCenter Then
        SaveSetting App.EXEName, "PrinterJustify", "Left", optLeftJust.Value
        SaveSetting App.EXEName, "PrinterJustify", "Center", optCenterJust.Value
    End If
    optLeftJust.Value = GetSetting(App.EXEName, "PrinterJustify", "Left", 0)
    optCenterJust.Value = GetSetting(App.EXEName, "PrinterJustify", "Center", 0)
End Sub
Private Sub optCenterJust_Click()
    SaveSetting App.EXEName, "PrinterJustify", "Center", optCenterJust.Value
    SaveSetting App.EXEName, "PrinterJustify", "Left", optLeftJust.Value
End Sub
Private Sub optLeftJust_Click()
    SaveSetting App.EXEName, "PrinterJustify", "Left", optLeftJust.Value
    SaveSetting App.EXEName, "PrinterJustify", "Center", optCenterJust.Value
End Sub
