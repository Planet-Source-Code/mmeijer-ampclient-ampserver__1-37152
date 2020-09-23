VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   0  'None
   Caption         =   "Amp Client"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Connection 
      Left            =   3960
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame frmPlayer 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.Label cmdFastrewind 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FRW"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.Label cmdFastforward 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FFW"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.Label cmdVolumeDown 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Volume -"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label cmdVolumeUp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Volume +"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label cmdRepeat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Repeat"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label cmdShuffle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Shuffle"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label cmdLast 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Last"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
      Begin VB.Label cmdFirst 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label cmdFadout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fadout"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label cmdPrevious 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Previous"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.Label cmdNext 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Next"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.Label cmdPause 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pause"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label cmdStop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stop"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label cmdPlay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Play"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2160
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   4440
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line7 
         X1              =   4440
         X2              =   4440
         Y1              =   0
         Y2              =   2160
      End
      Begin VB.Line Line6 
         X1              =   720
         X2              =   4440
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame frmConfig 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   4455
      Begin VB.TextBox txtRemotehost 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtRemotePort 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remote host"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remote port"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Line Line5 
         X1              =   720
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2160
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   4440
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         X1              =   4440
         X2              =   4440
         Y1              =   2160
         Y2              =   0
      End
      Begin VB.Line Line1 
         X1              =   1440
         X2              =   4440
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AMP Client"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label cmdExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cmdMinimize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cmdConfig 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Config"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Label cmdPlayer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LNG_TCP_PORT_DEFAULT As Long = 2235
Private Const LNG_TCP_PORT_MAX As Long = 65000
Private Const LNG_TCP_PORT_MIN As Long = 1

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub cmdConfig_Click()
    frmPlayer.ZOrder vbSendToBack
    frmConfig.ZOrder vbBringtoback
End Sub

Private Sub cmdFadout_Click()
    SendCommand "fadeout"
End Sub

Private Sub cmdFastforward_Click()
    SendCommand "fastforward"
End Sub

Private Sub cmdFastrewind_Click()
    SendCommand "fastrewind"
End Sub

Private Sub cmdFirst_Click()
    SendCommand "first"
End Sub

Private Sub cmdLast_Click()
    SendCommand "last"
End Sub

Private Sub cmdNext_Click()
    SendCommand "next"
End Sub

Private Sub cmdPause_Click()
    SendCommand "pause"
End Sub

Private Sub cmdPlay_Click()
    SendCommand "play"
End Sub

Private Sub cmdPlayer_Click()
    frmConfig.ZOrder vbSendToBack
    frmPlayer.ZOrder vbBringToFront
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMinimize_Click()
    Me.WindowState = 1
End Sub

Private Sub cmdPrevious_Click()
    SendCommand "previous"
End Sub

Private Sub cmdRepeat_Click()
    SendCommand "repeat"
End Sub

Private Sub cmdShuffle_Click()
    SendCommand "shuffle"
End Sub

Private Sub cmdStop_Click()
    SendCommand "stop"
End Sub

Private Sub cmdVolumeDown_Click()
    SendCommand "volumedown"
End Sub

Private Sub cmdVolumeUp_Click()
    SendCommand "volumeup"
End Sub

Private Sub Connection_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation
End Sub

Private Sub Form_Load()
    txtRemotePort.Text = LNG_TCP_PORT_DEFAULT
End Sub

Private Sub txtRemoteport_Change()
    Dim bytCount As Byte
    Dim strChar As String * 1
    If IsNumeric(txtRemotePort.Text) Then
        If txtRemotePort.Text > LNG_TCP_PORT_MAX Then txtRemotePort.Text = LNG_TCP_PORT_MAX
        If txtRemotePort.Text < LNG_TCP_PORT_MIN Then txtRemotePort.Text = LNG_TCP_PORT_MIN
        txtRemotePort.Text = CLng(txtRemotePort.Text)
    Else
        For bytCount = 1 To Len(txtRemotePort.Text)
            strChar = Mid(txtRemotePort.Text, bytCount, 1)
            If Not IsNumeric(strChar) Then txtRemotePort.Text = Replace(txtRemotePort.Text, strChar, "")
        Next
    End If
End Sub

Private Sub SendCommand(ByVal sCommand As String)
    'play,pause,next,previous,fadeout,first
    'last,shuffle,repeat,volumeup,volumedown
    'fastforward,fastrewind
    If sCommand = "" Then
        Exit Sub
    End If
    If txtRemotehost.Text = "" Then
        MsgBox "Specify a host to connect to!", vbExclamation
        Exit Sub
    End If
    If Connection.State <> sckClosed Then Connection.Close
    Connection.Connect txtRemotehost.Text, txtRemotePort.Text
    Do Until Connection.State = sckConnected
        DoEvents: DoEvents
        If Connection.State = sckError Then
            Exit Do
        End If
    Loop
    If Connection.State = sckConnected Then
        Connection.SendData (sCommand)
    End If
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub
