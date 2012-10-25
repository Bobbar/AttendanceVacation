VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmChart 
   Caption         =   "Stats Chart"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChart.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6345
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6855
      Left            =   -300
      OleObjectBlob   =   "frmChart.frx":0CCA
      TabIndex        =   0
      Top             =   -60
      Width           =   11115
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private HeightOffset As Long, WidthOffset As Long

Private Sub Form_Load()
HeightOffset = frmChart.Height - (MSChart1.Top + MSChart1.Height)
WidthOffset = frmChart.Width - (MSChart1.Left + MSChart1.Width)



MSChart1.ChartData = strChartData
MSChart1.Repaint = True
MSChart1.Refresh


End Sub

Private Sub Form_Resize()
MSChart1.Width = frmChart.Width - WidthOffset
MSChart1.Height = frmChart.Height - HeightOffset

End Sub
Private Sub MSChart1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then 'Button 2 is "Right Click"
        Dim blah
        
        blah = MsgBox("Copy data to clipboard for input into Excel?", vbOKCancel, "Copy Data")
        
        If blah = vbOK Then
            Dim strBreakdown
            Dim i As Integer
            
            For i = 0 To UBound(AttenStats) - 1
                strBreakdown = strBreakdown + AttenStats(i).ExTypeName & vbTab & AttenStats(i).ExTypeCount & vbCrLf
            Next i
            Clipboard.Clear
            Clipboard.SetText strBreakdown
        End If
    End If
End Sub
