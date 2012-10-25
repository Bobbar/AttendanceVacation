VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   15990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vbAcceleratorSGrid6.vbalGrid Grid1 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15915
      _ExtentX        =   28072
      _ExtentY        =   16113
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPicture=   "Form2.frx":0000
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
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Grid1.AddColumn 1, "Date"
    Grid1.AddColumn 2, "Date To"
    Grid1.AddColumn 3, "Excuse"
    Grid1.AddColumn 4, "Type"
    Grid1.AddColumn 5, "Hours"
    Grid1.AddColumn 6, "Notes"
    Grid1.AddColumn 7
    Grid1.ColumnVisible(7) = False
    Grid1.Gridlines = True


Grid1.LoadGridData "C:\GridData.dat"
ReSizeSGrid


End Sub
Private Sub ReSizeSGrid()
    'Grid1.Redraw = False

    Dim i, R, intCellPadding As Integer

    intCellPadding = 20

    For i = 1 To Grid1.Columns
        Grid1.AutoWidthColumn i
        Grid1.ColumnWidth(i) = Grid1.ColumnWidth(i) + intCellPadding
    Next i

    Grid1.ColumnWidth(6) = 500
    Grid1.ColumnWidth(2) = Grid1.ColumnWidth(1)
    Grid1.ColumnWidth(5) = 70

    For R = 1 To Grid1.Rows
        Grid1.AutoHeightRow R
    Next R

    Grid1.HeaderHotTrack = True
    'Grid1.Redraw = True

End Sub

