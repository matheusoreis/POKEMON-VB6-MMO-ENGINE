Attribute VB_Name = "modDxShow"
Option Explicit

Private m_objBasicAudio As IBasicAudio
Private m_objMediaEvent As IMediaEvent
Private m_objMediaControl As IMediaControl
Private m_objMediaPosition As IMediaPosition

Sub RemoveDShow()
    On Local Error GoTo RemoveDShowError

    'If a MediaControl instance exists, then stop it from playing
    If ObjPtr(m_objMediaControl) > 0 Then
        m_objMediaControl.Stop
    End If

    'Destroy all objects
    If ObjPtr(m_objBasicAudio) > 0 Then Set m_objBasicAudio = Nothing
    If ObjPtr(m_objMediaControl) > 0 Then Set m_objMediaControl = Nothing
    If ObjPtr(m_objMediaPosition) > 0 Then Set m_objMediaPosition = Nothing
    Exit Sub

RemoveDShowError:
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenDShowFile(filename As String)
    On Local Error GoTo OpenFileError

    Set m_objMediaControl = New FilgraphManager
    Call m_objMediaControl.RenderFile(filename)

    Set m_objBasicAudio = m_objMediaControl
    m_objBasicAudio.volume = 0    'Loudest
    m_objBasicAudio.Balance = 0    'Centered

    Set m_objMediaEvent = m_objMediaControl

    Set m_objMediaPosition = m_objMediaControl
    m_objMediaPosition.Rate = 1    'Normal forward playback speed

    Exit Sub

OpenFileError:
    Err.Clear
    Resume Next
End Sub

Public Sub PlayMp3()
'Play if DShow is initialized and a file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then
        If CLng(m_objMediaPosition.CurrentPosition) = CLng(m_objMediaPosition.Duration) Then
            m_objMediaPosition.CurrentPosition = 0
        End If
        Call m_objMediaControl.Run
    End If
End Sub

Public Sub PauseMp3()
'Pause if DShow is initialized and a file is loaded
    If ObjPtr(m_objMediaControl) > 0 Then
        Call m_objMediaControl.Pause
    End If
End Sub

Public Sub StopMp3()
'Stop if DShow is initialized and a file is loaded
    If (ObjPtr(m_objMediaControl) > 0) And (ObjPtr(m_objMediaPosition) > 0) Then
        m_objMediaControl.Stop
        m_objMediaPosition.CurrentPosition = 0
    End If
End Sub

Public Sub LoopMp3()
'Check only if DShow is initialized and a file is loaded
    If ObjPtr(m_objMediaEvent) > 0 Then
        If Mp3_StillPlaying(0) = False Then
            PlayMp3
        End If
    End If
End Sub

Public Sub SeekPosition(Amount As Double)
    If m_objMediaPosition.CurrentPosition + Amount < 0 Then
        m_objMediaPosition.CurrentPosition = 0
    ElseIf m_objMediaPosition.CurrentPosition + Amount > _
           m_objMediaPosition.Duration Then
        m_objMediaPosition.CurrentPosition = m_objMediaPosition.Duration
    Else
        m_objMediaPosition.CurrentPosition = m_objMediaPosition.CurrentPosition + Amount
    End If
End Sub

Public Function Mp3_StillPlaying(msTimeout As Long) As Boolean
    On Local Error Resume Next
    Dim EvCode As Long
    'Check only if DShow is initialized and a file is loaded
    If ObjPtr(m_objMediaEvent) > 0 Then
        m_objMediaEvent.WaitForCompletion msTimeout, EvCode
        If EvCode = 0 Then
            Mp3_StillPlaying = True
        Else
            Mp3_StillPlaying = False
        End If
    End If
End Function

Public Sub Mp3_SetRate(newRate As Double)
'rate must not be <= 0!
    If newRate <= 0 Then newRate = 0.1
    'Set rate if DShow is initialized and file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then m_objMediaPosition.Rate = newRate
End Sub

Public Function Mp3_GetDuration() As Double
'Get duration if DShow is initialized and file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then Mp3_GetDuration = m_objMediaPosition.Duration
End Function

Public Function Mp3_GetPosition() As Double
'Get position if DShow is initialized and file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then Mp3_GetPosition = m_objMediaPosition.CurrentPosition
End Function

Public Sub Mp3_SetPosition(newPosition As Double)
'Set position if DShow is initialized and file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then
        'if newPosition is out of bounds then correct
        If newPosition < 0 Then newPosition = 0
        If newPosition > m_objMediaPosition.Duration Then newPosition = m_objMediaPosition.Duration
        m_objMediaPosition.CurrentPosition = newPosition
    End If
End Sub

Public Function Mp3_GetVolume() As Long
'Get volume if DShow is initialized and file has audiostream
    If ObjPtr(m_objBasicAudio) > 0 Then Mp3_GetVolume = m_objBasicAudio.volume
End Function

Public Sub Mp3_SetVolume(newVolume As Long)
'Volume must be between 0 (loudest) and -10000 (disabled)
    If newVolume > 0 Then newVolume = 0
    If newVolume < -10000 Then newVolume = -10000
    'Set volume if DShow is initialized and file has audiostream
    If ObjPtr(m_objBasicAudio) > 0 Then m_objBasicAudio.volume = newVolume
End Sub

Public Function Mp3_GetBalance() As Long
'Get balance if DShow is initialized and file has audiostream
    If ObjPtr(m_objBasicAudio) > 0 Then Mp3_GetBalance = m_objBasicAudio.Balance
End Function

Public Sub Mp3_SetBalance(newBalance As Long)
'balance must be between -10000 (left) and +10000 (right)
    If newBalance < -10000 Then newBalance = -10000
    If newBalance > 10000 Then newBalance = 10000
    'Set balance if DShow is initialized and file has audiostream
    If ObjPtr(m_objBasicAudio) > 0 Then m_objBasicAudio.Balance = newBalance
End Sub


