VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Com Monitor (by: Paw Jershauge)"
   ClientHeight    =   8895
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11655
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Main"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS Say 
      Height          =   255
      Left            =   0
      OleObjectBlob   =   "Main.frx":0442
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   120
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   3
      ItemData        =   "Main.frx":049A
      Left            =   10080
      List            =   "Main.frx":04A7
      TabIndex        =   14
      Text            =   "1 Stop Bit"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   2
      ItemData        =   "Main.frx":04D1
      Left            =   10080
      List            =   "Main.frx":04E4
      TabIndex        =   13
      Text            =   "8 Bit"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      ItemData        =   "Main.frx":050B
      Left            =   10080
      List            =   "Main.frx":0518
      TabIndex        =   12
      Text            =   "Com Port 2"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      ItemData        =   "Main.frx":0540
      Left            =   10080
      List            =   "Main.frx":057A
      TabIndex        =   11
      Text            =   "9600 Baud"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      ItemData        =   "Main.frx":0645
      Left            =   10080
      List            =   "Main.frx":0652
      TabIndex        =   10
      Text            =   "1 Stop Bit"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      ItemData        =   "Main.frx":067C
      Left            =   10080
      List            =   "Main.frx":068F
      TabIndex        =   9
      Text            =   "8 Bit"
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      ItemData        =   "Main.frx":06B6
      Left            =   10080
      List            =   "Main.frx":06C3
      TabIndex        =   8
      Text            =   "Com Port 1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer StatusTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11160
      Top             =   8400
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      ItemData        =   "Main.frx":06EB
      Left            =   10080
      List            =   "Main.frx":0725
      TabIndex        =   7
      Text            =   "9600 Baud"
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   255
      Left            =   11160
      TabIndex        =   6
      Top             =   3240
      Width           =   375
   End
   Begin MSComctlLib.StatusBar Sts 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   8595
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5997
            MinWidth        =   5997
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3140
            MinWidth        =   3140
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5997
            MinWidth        =   5997
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3140
            MinWidth        =   3140
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox IncommingData 
      Appearance      =   0  'Flat
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      ToolTipText     =   "Only Data from the upper graph is showen here..."
      Top             =   3600
      Width           =   11415
   End
   Begin VB.TextBox SaveAs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   3240
      Width           =   10215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   655
      TabIndex        =   1
      Top             =   1680
      Width           =   9855
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   600
         Top             =   120
      End
      Begin MSCommLib.MSComm MSComm2 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   655
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   600
         Top             =   120
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
   End
   Begin VB.Line Line2 
      X1              =   11640
      X2              =   9960
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Log file :"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   600
   End
   Begin VB.Line Line1 
      X1              =   11640
      X2              =   9960
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Menu mnu 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveHEX 
         Caption         =   "Save data convertet into HEX"
      End
      Begin VB.Menu m 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuStart 
      Caption         =   "&Start"
   End
   Begin VB.Menu mnuStop 
      Caption         =   "S&top"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public P1X1 As Integer, P1X As Integer
Public P2X1 As Integer, P2X As Integer
Public P1Y1 As Integer, P1Y2 As Integer, P1Y3 As Integer, P1Y4 As Integer, P1Y5 As Integer, P1Y6 As Integer, P1Y7 As Integer, P1Y8 As Integer
Public P2Y1 As Integer, P2Y2 As Integer, P2Y3 As Integer, P2Y4 As Integer, P2Y5 As Integer, P2Y6 As Integer, P2Y7 As Integer, P2Y8 As Integer
Public sp As Boolean, sp1 As Boolean
Public tmptimer As Double, Datalen1 As Double, Datalen2 As Double

Private Sub cmdOpen_Click()
CMD.Filter = "Log File(*.log)|*.log|Log File Text(*.txt)|*.txt"
CMD.ShowSave
If CMD.FileName <> "" Then SaveAs.Text = CMD.FileName
SaveSetting App.EXEName, "Setting", "Savepath", SaveAs.Text
End Sub

Private Sub Combo1_Click(Index As Integer)
MSComm1.CommPort = Replace(Combo1(0).Text, "Com Port ", "")
MSComm1.Settings = Replace(Combo1(1).Text, " Baud", ",n,") & Replace(Combo1(2).Text, " Bit", ",") & Replace(Combo1(3).Text, " Stop Bit", "")
Sts.Panels(2).Text = "P" & Replace(Combo1(0).Text, "Com Port ", "") & ":" & Replace(Combo1(1).Text, " Baud", ",n,") & Replace(Combo1(2).Text, " Bit", ",") & Replace(Combo1(3).Text, " Stop Bit", "")
End Sub

Private Sub Combo2_Click(Index As Integer)
MSComm2.CommPort = Replace(Combo2(0).Text, "Com Port ", "")
MSComm2.Settings = Replace(Combo2(1).Text, " Baud", ",n,") & Replace(Combo2(2).Text, " Bit", ",") & Replace(Combo2(3).Text, " Stop Bit", "")
Sts.Panels(4).Text = "P" & Replace(Combo2(0).Text, "Com Port ", "") & ":" & Replace(Combo2(1).Text, " Baud", ",n,") & Replace(Combo2(2).Text, " Bit", ",") & Replace(Combo2(3).Text, " Stop Bit", "")
End Sub

Private Sub Form_Load()
Call Combo1_Click(0)
Call Combo2_Click(0)
SaveAs.Text = GetSetting(App.EXEName, "Setting", "Savepath")
End Sub

Function SaveLog(SaveHex As Boolean)
Dim TmpHexData As String
If SaveAs.Text <> "" Then
If IncommingData.Text <> "" Then
If SaveHex = False Then
 Open SaveAs.Text For Output As #1
 Print #1, IncommingData.Text
 Close #1
Else
 For a = 1 To Len(IncommingData.Text)
  If 2 = Len(Hex(Asc(Mid(IncommingData.Text, a, 1)))) Then
   TmpHexData = TmpHexData & Hex(Asc(Mid(IncommingData.Text, a, 1)))
  Else
   TmpHexData = TmpHexData & "0" & Hex(Asc(Mid(IncommingData.Text, a, 1)))
  End If
 Next a
 Open Left(SaveAs.Text, Len(SaveAs.Text) - 4) & "HEX" & Right(SaveAs.Text, 4) For Output As #1
 Print #1, TmpHexData
 Close #1
End If
End If
End If
End Function

Private Sub Form_Terminate()
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
If MSComm2.PortOpen = True Then MSComm2.PortOpen = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
If MSComm2.PortOpen = True Then MSComm2.PortOpen = False
End Sub

Private Sub mnuAbout_Click()
MsgBox "Com monitor was programmed to help people get information from the communication ports on a computer..." & Chr(13) & Chr(10) & "Dev. Paw jershauge (Paw.Jershauge@pc.dk)", vbInformation, "About"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuSaveHEX_Click()
SaveLog True
End Sub

Private Sub mnuStart_Click()
sp = True
sp1 = True
If Combo1(0).Text = Combo2(0).Text Then
MsgBox "Cant monitor the same communication port on both graphs", vbCritical, "Communication port error"
Else
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
If MSComm2.PortOpen = True Then MSComm2.PortOpen = False
P1Y1 = 24
P1Y3 = 47
P1Y5 = 71
P1Y7 = 94
P2Y1 = 24
P2Y3 = 47
P2Y5 = 71
P2Y7 = 94
Picture1.Line (0, 24)-(655, 24), vbWhite
Picture1.Line (0, 47)-(655, 47), vbWhite
Picture1.Line (0, 71)-(655, 71), vbWhite
Picture2.Line (0, 24)-(655, 24), vbWhite
Picture2.Line (0, 47)-(655, 47), vbWhite
Picture2.Line (0, 71)-(655, 71), vbWhite
tmptimer = Timer
Timer1.Enabled = True
Timer2.Enabled = True
StatusTimer.Enabled = True
For a = 0 To 3
Combo1(a).Enabled = False
Combo2(a).Enabled = False
Next a
End If
mnuSaveHEX.Enabled = False
End Sub

Function Draw1(L1 As Boolean, L2 As Boolean, L3 As Boolean, L4 As Boolean)
P1X1 = P1X
P1Y2 = P1Y1
P1Y4 = P1Y3
P1Y6 = P1Y5
P1Y8 = P1Y7
P1X = P1X + 1
If L1 = True Then P1Y1 = 0
If L1 = False Then P1Y1 = 24
If L2 = True Then P1Y3 = 25
If L2 = False Then P1Y3 = 47
If L3 = True Then P1Y5 = 48
If L3 = False Then P1Y5 = 71
If L4 = True Then P1Y7 = 72
If L4 = False Then P1Y7 = 94
Picture1.Line (P1X1, P1Y2)-(P1X, P1Y1), vbRed
Picture1.Line (P1X1, P1Y4)-(P1X, P1Y3), vbBlue
Picture1.Line (P1X1, P1Y6)-(P1X, P1Y5), vbGreen
Picture1.Line (P1X1, P1Y8)-(P1X, P1Y7), &H80FF&
If P1X = 655 Then
Picture1.Cls
Picture1.Line (0, 24)-(655, 24), vbWhite
Picture1.Line (0, 47)-(655, 47), vbWhite
Picture1.Line (0, 71)-(655, 71), vbWhite
P1X = 0
End If
End Function

Function Draw2(L1 As Boolean, L2 As Boolean, L3 As Boolean, L4 As Boolean)
P2X1 = P2X
P2Y2 = P2Y1
P2Y4 = P2Y3
P2Y6 = P2Y5
P2Y8 = P2Y7
P2X = P2X + 1
If L1 = True Then P2Y1 = 0
If L1 = False Then P2Y1 = 24
If L2 = True Then P2Y3 = 25
If L2 = False Then P2Y3 = 47
If L3 = True Then P2Y5 = 48
If L3 = False Then P2Y5 = 71
If L4 = True Then P2Y7 = 72
If L4 = False Then P2Y7 = 94
Picture2.Line (P2X1, P2Y2)-(P2X, P2Y1), vbRed
Picture2.Line (P2X1, P2Y4)-(P2X, P2Y3), vbBlue
Picture2.Line (P2X1, P2Y6)-(P2X, P2Y5), vbGreen
Picture2.Line (P2X1, P2Y8)-(P2X, P2Y7), &H80FF&
If P2X = 655 Then
Picture2.Cls
Picture2.Line (0, 24)-(655, 24), vbWhite
Picture2.Line (0, 47)-(655, 47), vbWhite
Picture2.Line (0, 71)-(655, 71), vbWhite
P2X = 0
End If
End Function

Private Sub mnuStop_Click()
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
If MSComm2.PortOpen = True Then MSComm2.PortOpen = False
Timer1.Enabled = False
Timer2.Enabled = False
StatusTimer.Enabled = False
For a = 0 To 3
Combo1(a).Enabled = True
Combo2(a).Enabled = True
Next a
mnuSaveHEX.Enabled = True
End Sub

Private Sub StatusTimer_Timer()
Sts.Panels(5).Text = Round(Timer - tmptimer, 4)
End Sub

Private Sub Timer1_Timer()
Dim TmpData1 As String
If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
TmpData1 = MSComm1.Input
IncommingData.Text = IncommingData.Text & TmpData1
Datalen1 = Datalen1 + Len(TmpData1)
If TmpData1 <> "" Then
Draw1 True, MSComm1.CDHolding, MSComm1.CTSHolding, MSComm1.DSRHolding
Sts.Panels(1).Text = "Receiving DATA [" & Datalen1 & "]"
If sp = True Then sp = False: Say.Speak "In comeing data in the upper graph"
Else
Draw1 False, MSComm1.CDHolding, MSComm1.CTSHolding, MSComm1.DSRHolding
Sts.Panels(1).Text = "No DATA [" & Datalen1 & "]"
End If
SaveLog False
End Sub

Private Sub Timer2_Timer()
Dim TmpData2 As String
If MSComm2.PortOpen = False Then MSComm2.PortOpen = True
TmpData2 = MSComm2.Input
Datalen2 = Datalen2 + Len(TmpData2)
If TmpData2 <> "" Then
Draw2 True, MSComm2.CDHolding, MSComm2.CTSHolding, MSComm2.DSRHolding
Sts.Panels(3).Text = "Receiving DATA [" & Datalen2 & "]"
If sp1 = True Then Say.Speak "In comeing data in the lower graph": sp1 = False
Else
Draw2 False, MSComm2.CDHolding, MSComm2.CTSHolding, MSComm2.DSRHolding
Sts.Panels(3).Text = "No DATA [" & Datalen2 & "]"
End If
End Sub
