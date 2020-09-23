VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   0  'None
   Caption         =   "AMP Server"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   1920
   End
   Begin MSWinsockLib.Winsock Connection 
      Left            =   120
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2235
   End
   Begin VB.Frame frmLog 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.TextBox txtLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   120
         Width           =   4215
      End
      Begin VB.Line Line10 
         X1              =   720
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line4 
         X1              =   4440
         X2              =   4440
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
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2160
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   4440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label cmdClearLog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clear"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.Frame frmConfig 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   4455
      Begin VB.TextBox txtLocalPort 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MaxLength       =   5
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local port"
         Height          =   255
         Left            =   120
         TabIndex        =   8
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
         X1              =   4440
         X2              =   4440
         Y1              =   0
         Y2              =   2160
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   4440
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line6 
         X1              =   1920
         X2              =   4440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line5 
         X1              =   1320
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Label cmdStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Start"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label cmdConfig 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Config"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.Label cmdLog 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Log"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   615
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LNG_TCP_PORT_DEFAULT As Long = 2235
Private Const LNG_TCP_PORT_MAX As Long = 65000
Private Const LNG_TCP_PORT_MIN As Long = 1
Private Const STR_MESSAGE_INACTIVE As String = "Doing nothing.."
Private Const STR_MESSAGE_ACTIVE As String = "Activated, listening.."
Private Const STR_COMMANDS As String = "play,stop,pause,next,previous,fadeout,first,last,shuffle,repeat,volumeup,volumedown,fastforward,fastrewind"

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub cmdClearLog_Click()
    txtLog.Text = ""
End Sub

Private Sub cmdConfig_Click()
    frmLog.ZOrder vbSendToBack
    
End Sub

Private Sub cmdLog_Click()
    frmConfig.ZOrder vbSendToBack
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMinimize_Click()
    Me.WindowState = 1
End Sub

Private Sub cmdStart_Click()
    If cmdStart.Caption = "Start" Then
        If Listen(txtLocalPort.Text) Then
            cmdStart.Caption = "Stop"
            Timer.Enabled = False
            cmdStart.BackColor = &H8080FF
            lblTitle.Caption = STR_MESSAGE_ACTIVE
            txtLog.Text = Now() & " - Service started" & vbNewLine & txtLog.Text
            txtLocalPort.Enabled = False
        Else
            MsgBox "Port error!" & vbNewLine & _
                   "The port you are using is already in use." & _
                   "Change the localport and try again.", vbExclamation
            txtLog.Text = Now() & " - Service failed to start" & vbNewLine & txtLog.Text
        End If
    Else
        If Connection.State <> sckClosed Then Connection.Close
        cmdStart.Caption = "Start"
        Timer.Enabled = True
        lblTitle.Caption = STR_MESSAGE_INACTIVE
        txtLog.Text = Now() & " - Service stopped" & vbNewLine & txtLog.Text
        txtLocalPort.Enabled = True
    End If
End Sub

Private Sub Connection_ConnectionRequest(ByVal requestID As Long)
    If Connection.State <> sckClosed Then Connection.Close
    Connection.Accept requestID
End Sub

Private Sub Connection_DataArrival(ByVal bytesTotal As Long)
    Dim strCommand As String
    Connection.GetData strCommand, , 15
    If isCommand(strCommand) Then
        Call ExecCommand(strCommand)
        txtLog.Text = Now() & " - " & Connection.RemoteHostIP & " - exec: " & strCommand & vbNewLine & txtLog.Text
    Else
        txtLog.Text = Now() & " - " & Connection.RemoteHostIP & " - invalid command: " & strCommand & vbNewLine & txtLog.Text
    End If
    If Connection.State <> sckClosed Then Connection.Close
    Connection.LocalPort = txtLocalPort.Text
    Connection.Listen
End Sub

Private Sub Timer_Timer()
    If cmdStart.Caption = "Start" Then
        If cmdStart.BackColor = &HC0C0FF Then
            cmdStart.BackColor = &H8080FF
        Else
            cmdStart.BackColor = &HC0C0FF
        End If
    End If
End Sub

Private Sub txtLocalport_Change()
    Dim bytCount As Byte
    Dim strChar As String * 1
    If IsNumeric(txtLocalPort.Text) Then
        If txtLocalPort.Text > LNG_TCP_PORT_MAX Then txtLocalPort.Text = LNG_TCP_PORT_MAX
        If txtLocalPort.Text < LNG_TCP_PORT_MIN Then txtLocalPort.Text = LNG_TCP_PORT_MIN
        txtLocalPort.Text = CLng(txtLocalPort.Text)
    Else
        For bytCount = 1 To Len(txtLocalPort.Text)
            strChar = Mid(txtLocalPort.Text, bytCount, 1)
            If Not IsNumeric(strChar) Then txtLocalPort.Text = Replace(txtLocalPort.Text, strChar, "")
        Next
    End If
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Function Listen(ByVal lPort As Long) As Boolean
    If Connection.State <> sckClosed Then Connection.Close
    Connection.LocalPort = lPort
    On Error Resume Next
    Connection.Listen
    If Err.Number = 0 Then
        Listen = True
    Else
        Listen = False
    End If
End Function

Private Sub Form_Load()
    txtLocalPort.Text = LNG_TCP_PORT_DEFAULT
    lblTitle.Caption = STR_MESSAGE_INACTIVE
    Timer.Enabled = True
    
End Sub

Function isCommand(ByVal sCommand As String) As Boolean
    Dim arrCommands() As String
    Dim bytCount As Byte
    arrCommands = Split(STR_COMMANDS, ",")
    isCommand = False
    For bytCount = 0 To UBound(arrCommands)
        If StrComp(arrCommands(bytCount), sCommand) = 0 Then
            isCommand = True
            Exit For
        End If
    Next
End Function

Private Sub ExecCommand(ByVal sCommand As String)
    Dim objWinamp As New Winamp
    Set objWinamp = New Winamp
    If Not objWinamp.isAlive Then
        txtLog.Text = Now() & " - Winamp is not active, command: """ & sCommand & """ not executed." & vbNewLine & txtLog.Text
        Exit Sub
    End If
    'play,pause,next,previous,fadeout,first
    'last,shuffle,repeat,volumeup,volumedown
    'fastforward,fastrewind
    Select Case sCommand
        Case "stop"
            objWinamp.Halt
        Case "play"
            objWinamp.Play
        Case "pause"
            objWinamp.Pause
        Case "next"
            objWinamp.NextTrack
        Case "previous"
            objWinamp.PreviousTrack
        Case "fadeout"
            objWinamp.FadeOut
        Case "first"
            objWinamp.FirstTrack
        Case "last"
            objWinamp.LastTrack
        Case "shuffle"
            objWinamp.Shuffle
        Case "repeat"
            objWinamp.Repeat
        Case "volumeup"
            objWinamp.VolumeIncrease
        Case "volumedown"
            objWinamp.VolumeDecrease
        Case "fastforward"
            objWinamp.FastForward
        Case "fastrewind"
            objWinamp.FastRewind
    End Select
    Set objWinamp = Nothing
End Sub
