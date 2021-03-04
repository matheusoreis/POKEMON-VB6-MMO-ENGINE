Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub CheckKeys()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
        If ShiftRun = True Then GoTo Continue
        ShiftRun = True
        Call SendPlayerRun(True)
    Else
        ShiftDown = False
        If ShiftRun = False Then GoTo Continue
        ShiftRun = False
        Call SendPlayerRun(False)
    End If

Continue:

    If GetKeyState(vbKeySpace) < 0 Then
        CheckMapGetItem
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If

    'Confusão
    If GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon) > 0 Then
        If GetPlayerEquipmentNgt(MyIndex, weapon, 5) > 0 Then
            ConfuseKeys
            Exit Sub
        End If
    Else
        If GetPlayerEquipmentNgt(MyIndex, weapon, 5) > 0 Then
            Call SetPlayerEquipmentNgt(MyIndex, 5, weapon, 0)
        End If
    End If

    If frmMain.optWOn.value = True Then
        If frmMain.txtMyChat.Visible = False Then
            'Move Up
            If GetKeyState(vbKeyW) < 0 Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
                Exit Sub
            Else
                DirUp = False
            End If

            'Move Right
            If GetKeyState(vbKeyD) < 0 Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
                Exit Sub
            Else
                DirRight = False
            End If

            'Move down
            If GetKeyState(vbKeyS) < 0 Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
                Exit Sub
            Else
                DirDown = False
            End If

            'Move left
            If GetKeyState(vbKeyA) < 0 Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
                Exit Sub
            Else
                DirLeft = False
            End If
        End If
    Else
        'Move Up
        If GetKeyState(vbKeyUp) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If

        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If

        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If

        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If
    End If


    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ConfuseKeys()
'Move Up
    If GetKeyState(vbKeyDown) < 0 Then
        DirUp = True
        DirDown = False
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirUp = False
    End If

    'Move Right
    If GetKeyState(vbKeyLeft) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = True
        Exit Sub
    Else
        DirRight = False
    End If

    'Move down
    If GetKeyState(vbKeyUp) < 0 Then
        DirUp = False
        DirDown = True
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirDown = False
    End If

    'Move left
    If GetKeyState(vbKeyRight) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = True
        DirRight = False
        Exit Sub
    Else
        DirLeft = False
    End If
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
    Dim ChatText As String
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim Command() As String
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ChatText = Trim$(MyText)

    If LenB(ChatText) = 0 Then
        If KeyAscii = vbKeyReturn Then
            chaton = Not chaton
            SetFocusOnGame
            frmMain.PicChat.Visible = True
        End If
        Exit Sub
    End If
    If chaton = False Then Exit Sub
    MyText = LCase$(ChatText)

    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        chaton = False
        SetFocusOnGame

        ' Handle when the player presses the return key
        If KeyAscii = vbKeyReturn Then

            ' Broadcast message
            If Left$(ChatText, 1) = "'" Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)

                If Len(ChatText) > 0 Then
                    Call BroadcastMsg(ChatText)
                End If

                MyText = vbNullString
                frmMain.txtMyChat.text = vbNullString
                Exit Sub
            End If

            ' Emote message
            If Left$(ChatText, 1) = "-" Then
                MyText = Mid$(ChatText, 2, Len(ChatText) - 1)

                If Len(ChatText) > 0 Then
                    Call EmoteMsg(ChatText)
                End If

                MyText = vbNullString
                frmMain.txtMyChat.text = vbNullString
                Exit Sub
            End If

            ' Player message
            If Left$(ChatText, 1) = "!" Then
                If Mid$(ChatText, 1, 2) = "! " Then GoTo Continue
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                Name = vbNullString

                ' Get the desired player from the user text
                For i = 1 To Len(ChatText)

                    If Mid$(ChatText, i, 1) <> Space(1) Then
                        Name = Name & Mid$(ChatText, i, 1)
                    Else
                        Exit For
                    End If

                Next

                ' Make sure they are actually sending something
                If Len(ChatText) - i > 0 Then
                    ChatText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                    ' Send the message to the player
                    Call PlayerMsg(ChatText, Name)
                Else
                    Call AddText("Usage: !playername (message)", AlertColor)
                End If

                MyText = vbNullString
                frmMain.txtMyChat.text = vbNullString
                Exit Sub
            End If

            If Left$(MyText, 1) = "/" Then
                Command = Split(MyText, Space(1))

                Select Case Command(0)

                Case "/honra"
                    Call AddText("Honra: " & GetPlayerHonra(MyIndex), HelpColor)

                Case "/ajuda"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Broadcast Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /info, /who, /fps, /fpslock", HelpColor)
                Case "/vit"
                   ' If txtMyChat.Visible = False Then
                If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                    SendTradeRequest
                    ' play sound
                    PlaySound Sound_ButtonClick, -1, -1
                Else
                    AddText "Invalid trade target.", BrightRed
                End If
           ' End If

                Case "/batalhar"
                    frmMain.PicBatalha.Visible = Not frmMain.PicBatalha.Visible
                    frmMain.PicBatalha.top = (frmMain.ScaleHeight / 2) - (frmMain.PicBatalha.Height / 2)
                    frmMain.PicBatalha.Left = (frmMain.ScaleWidth / 2) - (frmMain.PicBatalha.Width / 2)
                    frmMain.optEscolha(1).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")

                    frmMain.cmbTipo.ListIndex = 1
                    frmMain.optArena(0).value = True
                    Arena = 1

                    frmMain.PicArena(1).Picture = LoadPicture(App.Path & "\data files\graphics\arenas\1.bmp")
                    If Player(MyIndex).Arena(1) = 1 Then
                        frmMain.lblArena(3).Caption = "Arena: 1 - (Ocupada)"
                        frmMain.lblArena(3).ForeColor = QBColor(BrightRed)
                    Else
                        frmMain.lblArena(3).Caption = "Arena: 1 - (Livre)"
                        frmMain.lblArena(3).ForeColor = &HFF00&
                    End If

                Case "/leilao"
                    If frmLeilao.Visible = True Then
                        frmLeilao.Visible = False
                    Else
                        SendALeilao
                        frmLeilao.Visible = True
                    End If
                
                Case "/top"
                    frmDailyLogin.Visible = True
                    
                Case "/info"
                    ' Checks to make sure we have more than one string in the array

                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    Set buffer = New clsBuffer
                    buffer.WriteLong CPlayerInfoRequest
                    buffer.WriteString Command(1)
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                Case "/stats"
                    Set buffer = New clsBuffer
                    buffer.WriteLong CGetStats
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    BLoc = Not BLoc
                    ' Map Editor
                Case "/editmap"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapReport
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendBanList
                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If

                    SendBan Command(1)
                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditItem
                    ' Editing animation request
                Case "/editanimation"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditNpc
                Case "/editresource"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditResource
                    ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditSpell
                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                Case "/mute"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                    If UBound(Command) < 2 Then
                        AddText "Use: /mute (Nome) (Minutos)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(2)) Then
                        AddText "Use: /mute (Nome) (Minutos)", AlertColor
                        GoTo Continue
                    End If

                    SendMutePlayer Command(1), Command(2)

                    ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    SendBanDestroy
                    ' Packet debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    DEBUG_MODE = (Not DEBUG_MODE)
                Case Else
                    AddText "Not a valid command!", HelpColor
                End Select

                'continue label where we go instead of exiting the sub
Continue:
                MyText = vbNullString
                frmMain.txtMyChat.text = vbNullString
                Exit Sub
            End If

            ' Say message
            If Len(ChatText) > 0 Then
                Call SayMsg(ChatText)
            End If

            MyText = vbNullString
            frmMain.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Handle when the user presses the backspace key
        If (KeyAscii = vbKeyBack) Then
            If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        End If

        ' And if neither, then add the character to the user's text buffer
        If (KeyAscii <> vbKeyReturn) Then
            If (KeyAscii <> vbKeyBack) Then
                MyText = MyText & ChrW$(KeyAscii)
            End If
        End If

        ' Error handler
        Exit Sub
errorhandler:
        HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
    End If
    Exit Sub
End Sub
