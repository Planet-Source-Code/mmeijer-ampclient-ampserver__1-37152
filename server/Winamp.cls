VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Winamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_AMP_WINDOW_NAME = "Winamp v1.x"
Private Const WM_COMMAND = &H111
Private Const WM_USER = &H400
    
Private Const WA_PREVIOUS As Long = 40044
Private Const WA_NEXT As Long = 40048
Private Const WA_PLAY As Long = 40045
Private Const WA_PAUSE As Long = 40046
Private Const WA_STOP As Long = 40047
Private Const WA_FADEOUT As Long = 40147
Private Const WA_STOP_AFTER_TRACK As Long = 40157
Private Const WA_FFW As Long = 40148
Private Const WA_FRW As Long = 40144
Private Const WA_FIRST_TRACK As Long = 40154
Private Const WA_LAST_TRACK As Long = 40158
Private Const WA_VOLUME_INCREASE As Long = 40058
Private Const WA_VOLUME_DECREASE As Long = 40059
Private Const WA_TOGGLE_REPEAT As Long = 40022
Private Const WA_TOGGLE_SHUFFLE As Long = 40023
Private Const WA_CLOSE As Long = 40001
Private Const WA_PLAY_CD As Long = 40323
    
Private Const WA_GETVERSION As Long = 0
Private Const WA_CLEARPLAYLIST As Long = 101
Private Const WA_GETSTATUS As Long = 104
Private Const WA_GETTRACKPOSITION = 105
Private Const WA_GETTRACKLENGTH = 105
Private Const WA_SEEKTOPOSITION = 106
Private Const WA_SETVOLUME = 122
Private Const WA_SETBALANCE = 123

Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private hwndWinamp As Long

Public Sub SetBalance(ByVal lBalance As Long)
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Exit Sub
    End If
    If lBalance > 255 Then lBalance = 255
    If lBalance < 0 Then lBalance = 0
    SendMessage hwndWinamp, WM_USER, lBalance, WA_SETBALANCE
End Sub

Public Sub SetVolume(ByVal lVolume As Long)
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Exit Sub
    End If
    If lVolume > 255 Then lVolume = 255
    If lVolume < 0 Then lVolume = 0
    SendMessage hwndWinamp, WM_USER, lVolume, WA_SETVOLUME
End Sub

Public Function Search(ByVal lSeconds As Long) As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Search = 0
        Exit Function
    End If
    Search = SendMessage(hwndWinamp, WM_USER, (lSeconds * 1000), WA_SEEKTOPOSITION)
End Function

Public Function TrackLength() As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        TrackLength = 0
        Exit Function
    End If
    TrackLength = SendMessage(hwndWinamp, WM_USER, 0, WA_GETTRACKLENGTH)
End Function

Public Function TrackPosition() As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        TrackPosition = 0
        Exit Function
    End If
    TrackPosition = SendMessage(hwndWinamp, WM_USER, 0, WA_GETTRACKPOSITION)
End Function

Public Function Status() As Long
    'status: 0=not playing, 1=playing, 3=paused
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Status = 0
        Exit Function
    End If
    Status = SendMessage(hwndWinamp, WM_USER, 0, WA_GETSTATUS)
End Function

Public Function PlayerVersion() As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        PlayerVersion = 0
        Exit Function
    End If
    PlayerVersion = SendMessage(hwndWinamp, WM_USER, 0, WA_GETVERSION)
End Function

Public Function ClearPlayList() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        ClearPlayList = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_USER, 0, WA_CLEARPLAYLIST)
    If lngReturn = 0 Then
        ClearPlayList = False
    Else
        ClearPlayList = True
    End If
End Function

Public Function PlayCD() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        PlayCD = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_PLAY_CD, 0)
    If lngReturn = 0 Then
        PlayCD = False
    Else
        PlayCD = True
    End If
End Function

Public Function Quit() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Quit = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_CLOSE, 0)
    If lngReturn = 0 Then
        Quit = False
    Else
        Quit = True
    End If
End Function

Public Function Shuffle() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Shuffle = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_TOGGLE_SHUFFLE, 0)
    If lngReturn = 0 Then
        Shuffle = False
    Else
        Shuffle = True
    End If
End Function

Public Function StopAfterTrack() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        StopAfterTrack = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_STOP_AFTER_TRACK, 0)
    If lngReturn = 0 Then
        StopAfterTrack = False
    Else
        StopAfterTrack = True
    End If
End Function

Public Function Repeat() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Repeat = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_TOGGLE_REPEAT, 0)
    If lngReturn = 0 Then
        Repeat = False
    Else
        Repeat = True
    End If
End Function

Public Function VolumeDecrease() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        VolumeDecrease = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_VOLUME_DECREASE, 0)
    If lngReturn = 0 Then
        VolumeDecrease = False
    Else
        VolumeDecrease = True
    End If
End Function

Public Function VolumeIncrease() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        VolumeIncrease = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_VOLUME_INCREASE, 0)
    If lngReturn = 0 Then
        VolumeIncrease = False
    Else
        VolumeIncrease = True
    End If
End Function

Public Function LastTrack() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        LastTrack = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_LAST_TRACK, 0)
    If lngReturn = 0 Then
        LastTrack = False
    Else
        LastTrack = True
    End If
End Function

Public Function FirstTrack() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        FirstTrack = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_FIRST_TRACK, 0)
    If lngReturn = 0 Then
        FirstTrack = False
    Else
        FirstTrack = True
    End If
End Function

Public Function FastRewind() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        FastRewind = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_FRW, 0)
    If lngReturn = 0 Then
        FastRewind = False
    Else
        FastRewind = True
    End If
End Function

Public Function FastForward() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        FastForward = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_FFW, 0)
    If lngReturn = 0 Then
        FastForward = False
    Else
        FastForward = True
    End If
End Function

Public Function FadeOut() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        FadeOut = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_FADEOUT, 0)
    If lngReturn = 0 Then
        FadeOut = False
    Else
        FadeOut = True
    End If
End Function

Public Function Pause() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Pause = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_PAUSE, 0)
    If lngReturn = 0 Then
        Pause = False
    Else
        Pause = True
    End If
End Function

Public Function Play() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Play = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_PLAY, 0)
    If lngReturn = 0 Then
        Play = False
    Else
        Play = True
    End If
End Function

Public Function Halt() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        Halt = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_STOP, 0)
    If lngReturn = 0 Then
        Halt = False
    Else
        Halt = True
    End If
End Function

Public Function NextTrack() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        NextTrack = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_NEXT, 0)
    If lngReturn = 0 Then
        NextTrack = False
    Else
        NextTrack = True
    End If
End Function

Public Function PreviousTrack() As Boolean
    Dim lngReturn As Long
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp = 0 Then
        PreviousTrack = False
        Exit Function
    End If
    lngReturn = SendMessage(hwndWinamp, WM_COMMAND, WA_PREVIOUS, 0)
    If lngReturn = 0 Then
        PreviousTrack = False
    Else
        PreviousTrack = True
    End If
End Function

Public Function isAlive() As Boolean
    hwndWinamp = FindWindow(STR_AMP_WINDOW_NAME, vbNullString)
    If hwndWinamp <> 0 Then
        isAlive = True
    Else
        isAlive = False
    End If
End Function


