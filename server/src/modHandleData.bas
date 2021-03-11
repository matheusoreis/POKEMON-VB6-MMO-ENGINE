Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CRequestEditPokemon) = GetAddress(AddressOf HandleRequestEditPokemon)
    HandleDataSub(CSavePokemon) = GetAddress(AddressOf HandleSavePokemon)
    HandleDataSub(CRequestPokemon) = GetAddress(AddressOf HandleRequestPokemon)
    HandleDataSub(CSelectPoke) = GetAddress(AddressOf HandleSelectPoke)
    HandleDataSub(CMutePlayer) = GetAddress(AddressOf HandleMutePlayer)
    HandleDataSub(CEvolCommand) = GetAddress(AddressOf HandleEvolCommand)
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CRequestQuests) = GetAddress(AddressOf HandleRequestQuests)
    HandleDataSub(CQuestCommand) = GetAddress(AddressOf HandleQuestCommand)
    HandleDataSub(CLeiloar) = GetAddress(AddressOf HandleLeiloar)
    HandleDataSub(CComprar) = GetAddress(AddressOf HandleComprar)
    HandleDataSub(CRetirar) = GetAddress(AddressOf HandleRetirar)
    HandleDataSub(CChatComando) = GetAddress(AddressOf HandleChatComando)
    HandleDataSub(CSendSurfInit) = GetAddress(AddressOf HandleSendSurfInit)
    HandleDataSub(CLutarComando) = GetAddress(AddressOf handleLutarComando)
    HandleDataSub(CAprender) = GetAddress(AddressOf handleAprenderHab)
    HandleDataSub(CSetOrg) = GetAddress(AddressOf HandleSetOrg)
    HandleDataSub(CAbrir) = GetAddress(AddressOf HandleAbrir)
    HandleDataSub(CBuyOrgShop) = GetAddress(AddressOf HandleBuyOrgShop)
    HandleDataSub(CRecoverPass) = GetAddress(AddressOf HandleRecoverPass)
    HandleDataSub(CNewPass) = GetAddress(AddressOf HandleNewPass)
    HandleDataSub(CGiveVip) = GetAddress(AddressOf HandleObterVip)
    HandleDataSub(CSetVisual) = GetAddress(AddressOf HandleSetVisual)
    HandleDataSub(CPlayerRun) = GetAddress(AddressOf HandlePlayerRun)
    HandleDataSub(CComandGym) = GetAddress(AddressOf HandleComandoGym)
    HandleDataSub(CGrupoMsg) = GetAddress(AddressOf HandleGrupoMsg)
End Sub

Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim RecoveryKey As String
    Dim Email As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            RecoveryKey = Buffer.ReadString
            Email = Buffer.ReadString
            
             ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(RecoveryKey)) < 5 Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 5 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Or Len(Trim$(RecoveryKey)) > NAME_LENGTH Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Email Valido
            If IsValidEmail(Trim$(Email)) = False Then
                Call AlertMsg(Index, "Email Inválido.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password, RecoveryKey, Email)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(Index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(Player(Index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar Index
                Else
                    ' send new char shit
                    If Not IsPlaying(Index) Then
                        Call SendNewCharClasses(Index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(Index, Name)

            If LenB(Trim$(Player(Index).Name)) > 0 Then
                Call DeleteName(Player(Index).Name)
            End If

            Call ClearPlayer(Index)
            ' Everything went ok
            Call Kill(App.Path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Trim$(Buffer.ReadString)
            Password = Buffer.ReadString

            ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(Index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(Index, "Multiple account logins is not authorized.")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(Index, Name)
            ClearBank Index
            LoadBank Index, Name
            
            ' Check if character data has been created
            If LenB(Trim$(Player(Index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar Index
            Else
                ' send new char shit
                If Not IsPlaying(Index) Then
                    Call SendNewCharClasses(Index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        Sprite = Buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(Index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(Index) Then
            Call AlertMsg(Index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(Index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Class, Sprite)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        HandleUseChar Index
        
        Set Buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    If Player(Index).MutedTime > 0 Then
        PlayerMsg Index, "Você não pode falar!", BrightRed
        Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(Index), Index, Msg, QBColor(White))
    Msg = Trim$(GetPlayerName(Index)) & ": " & Msg
    Call SendChatBubble(GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER, Msg, White)
    Set Buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim S As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    If Player(Index).MutedTime > 0 Then
    PlayerMsg Index, "Você não pode falar!", BrightRed
    Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    S = "[Global]" & GetPlayerName(Index) & ": " & Msg
    Call SayMsg_Global(Index, Msg, QBColor(White))
    Call AddLog(S, PLAYER_LOG)
    Call TextAdd(S)
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
            Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Cannot message yourself.", BrightRed)
    End If
    
    Set Buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    movement = Buffer.ReadLong 'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(Index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(Index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(Index).StunDuration > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    'If Surf Decision
    If Player(Index).InSurf = 3 Then
        Player(Index).InSurf = 0
        SendSurfInit Index
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(Index).InShop > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(Index) <> tmpX Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    If GetPlayerY(Index) <> tmpY Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    If Player(Index).Flying = 1 Then
        PlayerMoveFly Index, Dir, movement
    Else
        Call PlayerMove(Index, Dir, movement)
    End If
    
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim InvNum As Long
Dim Buffer As clsBuffer
    
    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong
    Set Buffer = Nothing

    UseItem Index, InvNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim X As Long, Y As Long
    
    ' can't attack whilst casting
    If TempPlayer(Index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
     SendAttack Index

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then
            TryPlayerAttackPlayer Index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc Index, i
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            X = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
            X = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)
    End Select
    
    CheckResource Index, X, Y
    If GetPlayerEquipment(Index, weapon) > 0 Then SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, weapon)
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim Buffer As clsBuffer
Dim sMes As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PointType = Buffer.ReadByte 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(Index, PointType) >= 255 Then
            PlayerMsg Index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(Index, Stats.Strength, GetPlayerRawStat(Index, Stats.Strength) + 1)
                sMes = "ATTACK"
            Case Stats.Endurance
                Call SetPlayerStat(Index, Stats.Endurance, GetPlayerRawStat(Index, Stats.Endurance) + 1)
                sMes = "DEFENSE"
            Case Stats.Intelligence
                Call SetPlayerStat(Index, Stats.Intelligence, GetPlayerRawStat(Index, Stats.Intelligence) + 1)
                sMes = "SP.ATK"
            Case Stats.Agility
                Call SetPlayerStat(Index, Stats.Agility, GetPlayerRawStat(Index, Stats.Agility) + 1)
                sMes = "SPEED"
            Case Stats.Willpower
                Call SetPlayerStat(Index, Stats.Willpower, GetPlayerRawStat(Index, Stats.Willpower) + 1)
                sMes = "SP.DEF"
        End Select
        
        SendActionMsg GetPlayerMap(Index), "+1 " & sMes, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
       Dim i As Long

    For i = 1 To Vitals.Vital_Count - 1
        SendVital Index, i
    Next
    
    'Call SendStats(Index)
    SendPlayerData Index
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Name As String
Dim i As Long, Tempo As Long
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString
    Set Buffer = Nothing
    i = FindPlayer(Name)
    PlayerMsg Index, "Cartão do Treinador em Construção", White

    IniciarBatalharGym Index, 1
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(Index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp to yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
            Call PlayerMsg(Index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
    Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long, i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The sprite
    i = FindPlayer(Buffer.ReadString)
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    
    If i = 0 Or i > MAX_PLAYERS Then Exit Sub
    Call SetPlayerSprite(i, n)
    Call SendPlayerData(i)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim MapNum As Long
    Dim X As Long
    Dim Y As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Index)
    i = Map(MapNum).Revision + 1
    Call ClearMap(MapNum)
    
    Map(MapNum).Name = Buffer.ReadString
    Map(MapNum).Music = Buffer.ReadString
    Map(MapNum).Revision = i
    Map(MapNum).Moral = Buffer.ReadByte
    Map(MapNum).Up = Buffer.ReadLong
    Map(MapNum).Down = Buffer.ReadLong
    Map(MapNum).Left = Buffer.ReadLong
    Map(MapNum).Right = Buffer.ReadLong
    Map(MapNum).BootMap = Buffer.ReadLong
    Map(MapNum).BootX = Buffer.ReadByte
    Map(MapNum).BootY = Buffer.ReadByte
    Map(MapNum).MaxX = Buffer.ReadByte
    Map(MapNum).MaxY = Buffer.ReadByte
    Map(MapNum).Weather = Buffer.ReadLong
    Map(MapNum).Intensity = Buffer.ReadLong
    
    For X = 1 To 2
        Map(MapNum).LevelPoke(X) = Buffer.ReadLong
    Next
    
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(X, Y).Layer(i).X = Buffer.ReadLong
                Map(MapNum).Tile(X, Y).Layer(i).Y = Buffer.ReadLong
                Map(MapNum).Tile(X, Y).Layer(i).Tileset = Buffer.ReadLong
            Next
            Map(MapNum).Tile(X, Y).Type = Buffer.ReadByte
            Map(MapNum).Tile(X, Y).Data1 = Buffer.ReadLong
            Map(MapNum).Tile(X, Y).Data2 = Buffer.ReadLong
            Map(MapNum).Tile(X, Y).Data3 = Buffer.ReadLong
            Map(MapNum).Tile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(X) = Buffer.ReadLong
        Call ClearMapNpc(X, MapNum)
    Next

    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    ' Save the map
    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    Call ClearTempTile(MapNum)
    Call CacheResources(MapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i

    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim S As String
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Get yes/no value
    S = Buffer.ReadLong 'Parse(1)
    Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If S = 1 Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)

    'send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next

    TempPlayer(Index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapDone
    SendDataTo Index, Buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(Index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong 'CLng(Parse(1))
    Amount = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(Index, InvNum) < 1 Or GetPlayerInvItemNum(Index, InvNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(Index, InvNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(Index, InvNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
        Call SendMapNpcsToMap(GetPlayerMap(Index))
    Next

    CacheResources GetPlayerMap(Index)
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim S As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    S = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If LenB(Trim$(Map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                S = S & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If

    Next

    S = S & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    S = Mid$(S, 1, Len(S) - 2)
    S = S & "."
    Call PlayerMsg(Index, S, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(Index) & "!", White)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim F As Long
    Dim S As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    F = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, S
        Input #F, Name
        Call PlayerMsg(Index, n & ": Banned IP " & S & " by " & Name, White)
        n = n + 1
    Loop

    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim File As Long
    Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.Path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    Kill filename
    Call PlayerMsg(Index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEditMap
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SItemEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimationEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(Index) & " saved Animation #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim NpcNum As Long
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    NpcNum = Buffer.ReadLong

    ' Prevent hacking
    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = Buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(Npc(NpcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(NpcNum)
    Call SaveNpc(NpcNum)
    Call AddLog(GetPlayerName(Index) & " saved Npc #" & NpcNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(Index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SShopEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    shopNum = Buffer.ReadLong

    ' Prevent hacking
    If shopNum < 0 Or shopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set Buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopNum)
    Call SaveShop(shopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & shopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    SpellNum = Buffer.ReadLong

    ' Prevent hacking
    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call AddLog(GetPlayerName(Index) & " saved Spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If Trim$(GetPlayerName(Index)) = "Alifer" Then GoTo Continue
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

Continue:
    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    i = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then
    
    If Trim$(GetPlayerName(Index)) = "Alifer" Then
        If GetPlayerAccess(n) <= 0 Then
            Call GlobalMsg(GetPlayerName(n) & " você obteve Acesso: " & i & " Agora você faz parte da administração.", BrightCyan)
        End If

        Call SetPlayerAccess(n, i)
        Call SendPlayerData(n)
    End If

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, "Invalid access level. Access: " & i, White)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " você obteve Acesso: " & i & " Agora você faz parte da administração.", BrightCyan)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Invalid access level. Access:" & i, White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString) 'Parse(1))
    SaveOptions
    Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong 'CLng(Parse(1))
    Y = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Prevent subscript out of range
    If X < 0 Or X > Map(GetPlayerMap(Index)).MaxX Or Y < 0 Or Y > Map(GetPlayerMap(Index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        ' Change target
                        If TempPlayer(Index).targetType = TARGET_TYPE_PLAYER And TempPlayer(Index).target = i Then
                            TempPlayer(Index).target = 0
                            TempPlayer(Index).targetType = TARGET_TYPE_NONE
                            ' send target to player
                            SendTarget Index, 0
                        Else
                            TempPlayer(Index).target = i
                            TempPlayer(Index).targetType = TARGET_TYPE_PLAYER
                            ' send target to player
                            SendTarget Index, 0
                        End If
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next

    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index)).Npc(i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).Npc(i).X = X Then
                If MapNpc(GetPlayerMap(Index)).Npc(i).Y = Y Then
                    If TempPlayer(Index).target = i And TempPlayer(Index).targetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(Index).target = 0
                        TempPlayer(Index).targetType = TARGET_TYPE_NONE
                        ' send target to player
                        SendTarget Index, 0
                    Else
                        ' Change target
                        TempPlayer(Index).target = i
                        TempPlayer(Index).targetType = TARGET_TYPE_NPC
                        ' send target to player
                        SendTarget Index, i
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(Index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchInvSlots Index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        PlayerMsg Index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(Index).SpellCD(n) > GetTickCount Then
            PlayerMsg Index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchSpellSlots Index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendPing
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleUnequip(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerUnequipItem Index, Buffer.ReadLong
    Set Buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData Index
End Sub

Sub HandleRequestItems(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems Index
End Sub

Sub HandleRequestAnimations(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations Index
End Sub

Sub HandleRequestNPCS(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs Index
End Sub

Sub HandleRequestResources(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources Index
End Sub

Sub HandleRequestSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells Index
End Sub

Sub HandleRequestShops(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops Index
End Sub

Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Command As Byte
    Dim tmpItem As Long
    Dim tmpAmount As Long
    Dim Pokemon As Long, Pokeball As Long, Level As Long, EXP As Long
    Dim Vital(1 To Vitals.Vital_Count - 1) As Long, MaxVital(1 To Vitals.Vital_Count - 1) As Long
    Dim Stat(1 To Stats.Stat_Count - 1) As Long, Spell(1 To MAX_POKE_SPELL), i As Long
    Dim Felicidade As Long, Sexo As Byte, Shiny As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' item
    Command = Buffer.ReadByte
    
    If Command > 1 Then Exit Sub
    
    If Command = 0 Then
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
    
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
        SpawnItem tmpItem, tmpAmount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index)
    Else
    
    Pokemon = Buffer.ReadLong
    Pokeball = Buffer.ReadLong
    Level = Buffer.ReadLong
    EXP = Buffer.ReadLong
    
    Vital(1) = Buffer.ReadLong
    Vital(2) = Buffer.ReadLong
    MaxVital(1) = Buffer.ReadLong
    MaxVital(2) = Buffer.ReadLong
    
    Stat(1) = Buffer.ReadLong
    Stat(2) = Buffer.ReadLong
    Stat(3) = Buffer.ReadLong
    Stat(4) = Buffer.ReadLong
    Stat(5) = Buffer.ReadLong
    
    Spell(1) = Buffer.ReadLong
    Spell(2) = Buffer.ReadLong
    Spell(3) = Buffer.ReadLong
    Spell(4) = Buffer.ReadLong
    
    Felicidade = Buffer.ReadLong
    Sexo = Buffer.ReadByte
    Shiny = Buffer.ReadByte
    
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
        SpawnItem 3, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index), Pokemon, Pokeball, Level, EXP, Vital(1), Vital(2), MaxVital(1), MaxVital(2), Stat(1), Stat(4), Stat(3), Stat(2), Stat(5), Spell(1), Spell(2), Spell(3), Spell(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Felicidade, Sexo, Shiny
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SetPlayerExp Index, GetPlayerNextLevel(Index)
    CheckPlayerLevelUp Index
End Sub

Sub HandleForgetSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellslot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellslot = Buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(Index).spellBuffer.Spell = spellslot Then
        PlayerMsg Index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(Index).Spell(spellslot) = 0
    SendPlayerSpells Index
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(Index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemamount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopslot = Buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(Index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        itemamount = HasItem(Index, .costitem)
        If itemamount = 0 Or itemamount < .costvalue Then
            PlayerMsg Index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction Index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem Index, .costitem, .costvalue
        GiveInvItem Index, .Item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Trade successful.", BrightGreen
    ResetShopAction Index
    
    Set Buffer = Nothing
End Sub

Sub HandleSellItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invslot As Long
    Dim ItemNum As Long
    Dim Price As Long
    Dim multiplier As Double
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invslot = Buffer.ReadLong
    
    ' if invalid, exit out
    If invslot < 1 Or invslot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(Index, invslot) < 1 Or GetPlayerInvItemNum(Index, invslot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    ItemNum = GetPlayerInvItemNum(Index, invslot)
    
    ' work out price
    multiplier = Shop(TempPlayer(Index).InShop).BuyRate / 100
    Price = Item(ItemNum).Price * multiplier
    
    ' item has cost?
    If Price <= 0 Then
        PlayerMsg Index, "The shop doesn't want that item.", BrightRed
        ResetShopAction Index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem Index, ItemNum, 1
    GiveInvItem Index, 1, Price
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Trade successful.", BrightGreen
    ResetShopAction Index
    
    Set Buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots Index, oldSlot, newSlot
    
    Set Buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    If GetPlayerBankItemNum(Index, BankSlot) = 0 Then Exit Sub
    
    If GetPlayerBankItemPokemon(Index, BankSlot) > 0 Or Item(GetPlayerBankItemNum(Index, BankSlot)).Type = ITEM_TYPE_ROD Then
    TakeBankItemPokemon Index, BankSlot
    Else
    TakeBankItem Index, BankSlot, Amount
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invslot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invslot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    If GetPlayerInvItemNum(Index, invslot) = 0 Then Exit Sub
    
    If GetPlayerInvItemPokeInfoPokemon(Index, invslot) > 0 Or Item(GetPlayerInvItemNum(Index, invslot)).Type = ITEM_TYPE_ROD Then
        GiveBankItemPokemon Index, invslot, Amount
    Else
        GiveBankItem Index, invslot, Amount
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SaveBank Index
    SavePlayer Index
    
    TempPlayer(Index).InBank = False
    
    Set Buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long
    Dim Y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    
    If GetPlayerAccess(Index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX Index, X
        SetPlayerY Index, Y
        SendPlayerXYToMap Index
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(Index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(Index).target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = Index Then
        PlayerMsg Index, "Você não pode trocar com você mesmo!", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(Index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).X
    tY = Player(tradeTarget).Y
    sX = Player(Index).X
    sY = Player(Index).Y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg Index, "Você precisa estar próximo da pessoa para trocar itens!", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg Index, "Você precisa estar próximo da pessoa para trocar itens!", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg Index, "Jogador alvo já está realizando uma troca.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = Index
    SendTradeRequest tradeTarget, Index
End Sub

Sub HandleAcceptTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long
Dim X As Long

    If TempPlayer(Index).InTrade > 0 Then Exit Sub

    If TempPlayer(Index).InTrade > 0 Then
        TempPlayer(Index).TradeRequest = 0
    Else

    tradeTarget = TempPlayer(Index).TradeRequest
    ' let them know they're trading
    PlayerMsg Index, "Você aceitou a troca com " & Trim$(GetPlayerName(tradeTarget)) & ".", BrightGreen
    PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has accepted your trade request.", BrightGreen
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
    TempPlayer(tradeTarget).TradeRequest = 0
    ' set that they're trading with each other
    TempPlayer(Index).InTrade = tradeTarget
    TempPlayer(tradeTarget).InTrade = Index
    ' clear out their trade offers
    For i = 1 To MAX_INV
        'Você
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).Value = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Pokemon = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Pokeball = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Level = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.EXP = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Felicidade = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Sexo = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Shiny = 0
        
        For X = 1 To Vitals.Vital_Count - 1
                TempPlayer(Index).TradeOffer(i).PokeInfo.Vital(X) = 0
                TempPlayer(Index).TradeOffer(i).PokeInfo.MaxVital(X) = 0
        Next
        
        For X = 1 To Stats.Stat_Count - 1
                TempPlayer(Index).TradeOffer(i).PokeInfo.Stat(X) = 0
        Next
        
        For X = 1 To MAX_POKE_SPELL
                TempPlayer(Index).TradeOffer(i).PokeInfo.Spells(X) = 0
        Next
        
        'Outro Jogador
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Pokemon = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Pokeball = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Level = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.EXP = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Felicidade = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Sexo = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Shiny = 0
        
        For X = 1 To Vitals.Vital_Count - 1
                TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Vital(X) = 0
                TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.MaxVital(X) = 0
        Next
        
        For X = 1 To Stats.Stat_Count - 1
                TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Stat(X) = 0
        Next
        
        For X = 1 To MAX_POKE_SPELL
                TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Spells(X) = 0
        Next

    Next
    ' Used to init the trade window clientside
    SendTrade Index, tradeTarget
    SendTrade tradeTarget, Index
    
    ' Send the offer data - Used to clear their client
    SendTradeUpdate Index, 0
    SendTradeUpdate Index, 1
    SendTradeUpdate tradeTarget, 0
    SendTradeUpdate tradeTarget, 1
    End If
End Sub

Sub HandleDeclineTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(Index).TradeRequest, GetPlayerName(Index) & " has declined your trade request.", BrightRed
    PlayerMsg Index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long, X As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim ItemNum As Long
    
    If GetPlayerMap(Index) <> GetPlayerMap(TempPlayer(Index).InTrade) Then Exit Sub
    
    TempPlayer(Index).AcceptTrade = True
    
    tradeTarget = TempPlayer(Index).InTrade
    
    If tradeTarget > 0 Then
    
    ' if not both of them accept, then exit
    If Not TempPlayer(tradeTarget).AcceptTrade Then
        SendTradeStatus Index, 2
        SendTradeStatus tradeTarget, 1
        Exit Sub
    End If
    
     ' if not have space in inventory of tradetarget
        If IsInventoryFull(tradeTarget, Index) Then
            TempPlayer(Index).InTrade = 0
            TempPlayer(tradeTarget).InTrade = 0
            TempPlayer(Index).AcceptTrade = False
            TempPlayer(tradeTarget).AcceptTrade = False
            PlayerMsg tradeTarget, "Você não tem espaço suficiente no inventário.", BrightRed
            PlayerMsg Index, GetPlayerName(tradeTarget) & " não tem espaço suficiente no inventário.", BrightRed
            SendCloseTrade Index
            SendCloseTrade tradeTarget
            Exit Sub '
        End If
        
        ' if not have space in inventory of index
        If IsInventoryFull(Index, tradeTarget) Then
            TempPlayer(Index).InTrade = 0
            TempPlayer(tradeTarget).InTrade = 0
            TempPlayer(Index).AcceptTrade = False
            TempPlayer(tradeTarget).AcceptTrade = False
            PlayerMsg Index, "Você não tem espaço suficiente no inventário.", BrightRed
            PlayerMsg tradeTarget, GetPlayerName(Index) & " não tem espaço suficiente no inventário.", BrightRed
            SendCloseTrade Index
            SendCloseTrade tradeTarget
            Exit Sub
        End If
    
    ' take their items
    For i = 1 To MAX_INV
        ' player
        If TempPlayer(Index).TradeOffer(i).Num > 0 Then
            ItemNum = Player(Index).Inv(TempPlayer(Index).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem(i).Num = ItemNum
                tmpTradeItem(i).Value = TempPlayer(Index).TradeOffer(i).Value
                tmpTradeItem(i).PokeInfo.Pokemon = GetPlayerInvItemPokeInfoPokemon(Index, TempPlayer(Index).TradeOffer(i).Num)
                tmpTradeItem(i).PokeInfo.Pokeball = GetPlayerInvItemPokeInfoPokeball(Index, TempPlayer(Index).TradeOffer(i).Num)
                tmpTradeItem(i).PokeInfo.Level = GetPlayerInvItemPokeInfoLevel(Index, TempPlayer(Index).TradeOffer(i).Num)
                tmpTradeItem(i).PokeInfo.EXP = GetPlayerInvItemPokeInfoExp(Index, TempPlayer(Index).TradeOffer(i).Num)
                tmpTradeItem(i).PokeInfo.Felicidade = GetPlayerInvItemFelicidade(Index, TempPlayer(Index).TradeOffer(i).Num)
                tmpTradeItem(i).PokeInfo.Sexo = GetPlayerInvItemSexo(Index, TempPlayer(Index).TradeOffer(i).Num)
                tmpTradeItem(i).PokeInfo.Shiny = GetPlayerInvItemShiny(Index, TempPlayer(Index).TradeOffer(i).Num)
                
                For X = 1 To Vitals.Vital_Count - 1
                    tmpTradeItem(i).PokeInfo.Vital(X) = GetPlayerInvItemPokeInfoVital(Index, TempPlayer(Index).TradeOffer(i).Num, X)
                    tmpTradeItem(i).PokeInfo.MaxVital(X) = GetPlayerInvItemPokeInfoMaxVital(Index, TempPlayer(Index).TradeOffer(i).Num, X)
                Next
                
                For X = 1 To Stats.Stat_Count - 1
                    tmpTradeItem(i).PokeInfo.Stat(X) = GetPlayerInvItemPokeInfoStat(Index, TempPlayer(Index).TradeOffer(i).Num, X)
                Next
                
                For X = 1 To MAX_POKE_SPELL
                    tmpTradeItem(i).PokeInfo.Spells(X) = GetPlayerInvItemPokeInfoSpell(Index, TempPlayer(Index).TradeOffer(i).Num, X)
                Next
                
                For X = 1 To MAX_BERRYS
                    tmpTradeItem(i).PokeInfo.Berry(X) = GetPlayerInvItemBerry(Index, TempPlayer(Index).TradeOffer(i).Num, X)
                Next
                
                ' take item
                TakeInvSlot Index, TempPlayer(Index).TradeOffer(i).Num, tmpTradeItem(i).Value
            End If
        End If
        ' target
        If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
            ItemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem2(i).Num = ItemNum
                tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                tmpTradeItem2(i).PokeInfo.Pokemon = GetPlayerInvItemPokeInfoPokemon(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                tmpTradeItem2(i).PokeInfo.Pokeball = GetPlayerInvItemPokeInfoPokeball(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                tmpTradeItem2(i).PokeInfo.Level = GetPlayerInvItemPokeInfoLevel(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                tmpTradeItem2(i).PokeInfo.EXP = GetPlayerInvItemPokeInfoExp(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                tmpTradeItem2(i).PokeInfo.Felicidade = GetPlayerInvItemFelicidade(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                tmpTradeItem2(i).PokeInfo.Sexo = GetPlayerInvItemSexo(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                tmpTradeItem2(i).PokeInfo.Shiny = GetPlayerInvItemShiny(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                
                For X = 1 To Vitals.Vital_Count - 1
                    tmpTradeItem2(i).PokeInfo.Vital(X) = GetPlayerInvItemPokeInfoVital(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
                    tmpTradeItem2(i).PokeInfo.MaxVital(X) = GetPlayerInvItemPokeInfoMaxVital(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
                Next
                
                For X = 1 To Stats.Stat_Count - 1
                    tmpTradeItem2(i).PokeInfo.Stat(X) = GetPlayerInvItemPokeInfoStat(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
                Next
                
                For X = 1 To MAX_POKE_SPELL
                    tmpTradeItem2(i).PokeInfo.Spells(X) = GetPlayerInvItemPokeInfoSpell(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
                Next
                
                For X = 1 To MAX_BERRYS
                    tmpTradeItem2(i).PokeInfo.Berry(X) = GetPlayerInvItemBerry(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
                Next
                
                ' take item
                TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).Value
            End If
        End If
    Next
    
    ' taken all items. now they can't not get items because of no inventory space.
    For i = 1 To MAX_INV
        ' player
        If tmpTradeItem2(i).Num > 0 Then
            ' give away!
            GiveInvItem Index, tmpTradeItem2(i).Num, tmpTradeItem2(i).Value, False, tmpTradeItem2(i).PokeInfo.Pokemon, tmpTradeItem2(i).PokeInfo.Pokeball, tmpTradeItem2(i).PokeInfo.Level, tmpTradeItem2(i).PokeInfo.EXP, tmpTradeItem2(i).PokeInfo.Vital(1), tmpTradeItem2(i).PokeInfo.Vital(2), tmpTradeItem2(i).PokeInfo.MaxVital(1), tmpTradeItem2(i).PokeInfo.MaxVital(2), tmpTradeItem2(i).PokeInfo.Stat(1), tmpTradeItem2(i).PokeInfo.Stat(4), tmpTradeItem2(i).PokeInfo.Stat(2), tmpTradeItem2(i).PokeInfo.Stat(3), tmpTradeItem2(i).PokeInfo.Stat(5), _
            tmpTradeItem2(i).PokeInfo.Spells(1), tmpTradeItem2(i).PokeInfo.Spells(2), tmpTradeItem2(i).PokeInfo.Spells(3), tmpTradeItem2(i).PokeInfo.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
            tmpTradeItem2(i).PokeInfo.Felicidade, tmpTradeItem2(i).PokeInfo.Sexo, tmpTradeItem2(i).PokeInfo.Shiny
        End If
        ' target
        If tmpTradeItem(i).Num > 0 Then
            ' give away!
            GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).Value, False, tmpTradeItem(i).PokeInfo.Pokemon, tmpTradeItem(i).PokeInfo.Pokeball, tmpTradeItem(i).PokeInfo.Level, tmpTradeItem(i).PokeInfo.EXP, tmpTradeItem(i).PokeInfo.Vital(1), tmpTradeItem(i).PokeInfo.Vital(2), tmpTradeItem(i).PokeInfo.MaxVital(1), tmpTradeItem(i).PokeInfo.MaxVital(2), tmpTradeItem(i).PokeInfo.Stat(1), tmpTradeItem(i).PokeInfo.Stat(4), tmpTradeItem(i).PokeInfo.Stat(2), tmpTradeItem(i).PokeInfo.Stat(3), tmpTradeItem(i).PokeInfo.Stat(5), _
            tmpTradeItem(i).PokeInfo.Spells(1), tmpTradeItem(i).PokeInfo.Spells(2), tmpTradeItem(i).PokeInfo.Spells(3), tmpTradeItem(i).PokeInfo.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
            tmpTradeItem(i).PokeInfo.Felicidade, tmpTradeItem(i).PokeInfo.Sexo, tmpTradeItem(i).PokeInfo.Shiny
        End If
    Next
    
    SendInventory Index
    SendInventory tradeTarget
    
    ' they now have all the items. Clear out values + let them out of the trade.
    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).Value = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Pokemon = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Pokeball = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Level = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.EXP = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Felicidade = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Sexo = 0
        TempPlayer(Index).TradeOffer(i).PokeInfo.Shiny = 0
            
        For X = 1 To Vitals.Vital_Count - 1
            TempPlayer(Index).TradeOffer(i).PokeInfo.Vital(X) = 0
            TempPlayer(Index).TradeOffer(i).PokeInfo.MaxVital(X) = 0
        Next
            
        For X = 1 To Stats.Stat_Count - 1
            TempPlayer(Index).TradeOffer(i).PokeInfo.Stat(X) = 0
        Next
            
        For X = 1 To MAX_POKE_SPELL
            TempPlayer(Index).TradeOffer(i).PokeInfo.Spells(X) = 0
        Next
            
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Pokemon = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Pokeball = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Level = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.EXP = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Felicidade = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Sexo = 0
        TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Shiny = 0
            
        For X = 1 To Vitals.Vital_Count - 1
            TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Vital(X) = 0
            TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.MaxVital(X) = 0
        Next
            
        For X = 1 To Stats.Stat_Count - 1
            TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Stat(X) = 0
        Next
            
        For X = 1 To MAX_POKE_SPELL
            TempPlayer(tradeTarget).TradeOffer(i).PokeInfo.Spells(X) = 0
        Next
    Next

    TempPlayer(Index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg Index, "Troca Completa.", BrightGreen
    PlayerMsg tradeTarget, "Troca Completa.", BrightGreen
    
    SendCloseTrade Index
    SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleDeclineTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(Index).InTrade

    If tradeTarget > 0 Then

    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).Value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).Value = 0
    Next

    TempPlayer(Index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(tradeTarget).AcceptTrade = False
    
    PlayerMsg Index, "You declined the trade.", BrightRed
    PlayerMsg tradeTarget, GetPlayerName(Index) & " has declined the trade.", BrightRed
    
    SendCloseTrade Index
    SendCloseTrade tradeTarget
    
    End If
End Sub

Sub HandleTradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invslot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Long
    Dim i As Long, X As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invslot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If Not TempPlayer(Index).InTrade > 0 Then Exit Sub
    If invslot <= 0 Or invslot > MAX_INV Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(Index, invslot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invslot) Then
        Exit Sub
    End If
    
    If Item(ItemNum).NTrade = True Then
        PlayerMsg Index, "Este item não pode ser negociado", BrightRed
        Exit Sub
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Then Exit Sub
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invslot Then
                ' add amount
                TempPlayer(Index).TradeOffer(i).Value = TempPlayer(Index).TradeOffer(i).Value + Amount
                ' clamp to limits
                If TempPlayer(Index).TradeOffer(i).Value > GetPlayerInvItemValue(Index, invslot) Then
                    TempPlayer(Index).TradeOffer(i).Value = GetPlayerInvItemValue(Index, invslot)
                End If
                
                'Colocar valores
                TempPlayer(Index).TradeOffer(i).PokeInfo.Pokemon = GetPlayerInvItemPokeInfoPokemon(Index, invslot)
                TempPlayer(Index).TradeOffer(i).PokeInfo.Pokeball = GetPlayerInvItemPokeInfoPokeball(Index, invslot)
                TempPlayer(Index).TradeOffer(i).PokeInfo.Level = GetPlayerInvItemPokeInfoLevel(Index, invslot)
                TempPlayer(Index).TradeOffer(i).PokeInfo.EXP = GetPlayerInvItemPokeInfoExp(Index, invslot)
                TempPlayer(Index).TradeOffer(i).PokeInfo.Felicidade = GetPlayerInvItemFelicidade(Index, invslot)
                TempPlayer(Index).TradeOffer(i).PokeInfo.Sexo = GetPlayerInvItemSexo(Index, invslot)
                TempPlayer(Index).TradeOffer(i).PokeInfo.Shiny = GetPlayerInvItemShiny(Index, invslot)
                
                For X = 1 To Vitals.Vital_Count - 1
                    TempPlayer(Index).TradeOffer(i).PokeInfo.Vital(X) = GetPlayerInvItemPokeInfoVital(Index, invslot, X)
                    TempPlayer(Index).TradeOffer(i).PokeInfo.MaxVital(X) = GetPlayerInvItemPokeInfoMaxVital(Index, invslot, X)
                Next
                
                For X = 1 To Stats.Stat_Count - 1
                    TempPlayer(Index).TradeOffer(i).PokeInfo.Stat(X) = GetPlayerInvItemPokeInfoStat(Index, invslot, X)
                Next
                
                For X = 1 To MAX_POKE_SPELL
                    TempPlayer(Index).TradeOffer(i).PokeInfo.Spells(X) = GetPlayerInvItemPokeInfoSpell(Index, invslot, X)
                Next
                
                For X = 1 To MAX_BERRYS
                    TempPlayer(Index).TradeOffer(i).PokeInfo.Berry(X) = GetPlayerInvItemBerry(Index, invslot, X)
                Next
                
                ' cancel any trade agreement
                TempPlayer(Index).AcceptTrade = False
                TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
                
                SendTradeStatus Index, 0
                SendTradeStatus TempPlayer(Index).InTrade, 0
                
                SendTradeUpdate Index, 0
                SendTradeUpdate TempPlayer(Index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invslot Then
                PlayerMsg Index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(Index).TradeOffer(EmptySlot).Num = invslot
    TempPlayer(Index).TradeOffer(EmptySlot).Value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeSlot = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(Index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(Index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(Index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(Index).AcceptTrade Then TempPlayer(Index).AcceptTrade = False
    If TempPlayer(TempPlayer(Index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    sType = Buffer.ReadLong
    Slot = Buffer.ReadLong
    hotbarNum = Buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(Index).Hotbar(hotbarNum).Slot = 0
            Player(Index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(Index).Inv(Slot).Num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Inv(Slot).Num
                        Player(Index).Hotbar(hotbarNum).sType = sType
                        Player(Index).Hotbar(hotbarNum).Pokemon = GetPlayerInvItemPokeInfoPokemon(Index, Slot)
                        Player(Index).Hotbar(hotbarNum).Pokeball = GetPlayerInvItemPokeInfoPokeball(Index, Slot)
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(Index).Spell(Slot) > 0 Then
                    If Len(Trim$(Spell(Player(Index).Spell(Slot)).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Spell(Slot)
                        Player(Index).Hotbar(hotbarNum).sType = sType
                        Player(Index).Hotbar(hotbarNum).Pokemon = 0
                        Player(Index).Hotbar(hotbarNum).Pokeball = 0
                    End If
                End If
            End If
    End Select
    
    SendHotbar Index
    
    Set Buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long
    Dim PokeX As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    
    Select Case Player(Index).Hotbar(Slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(Index).Inv(i).Num > 0 Then
                    If Player(Index).Hotbar(Slot).Pokemon > 0 Then
                    PokeX = PokeX + 1
                    If Player(Index).Hotbar(Slot).Pokemon = GetPlayerInvItemPokeInfoPokemon(Index, i) Then
                    If Player(Index).Hotbar(Slot).Pokeball = GetPlayerInvItemPokeInfoPokeball(Index, i) Then
                        UseItem Index, i
                        Exit Sub
                    End If
                    End If
                End If
                End If
            Next
            
            If PokeX > 0 Then Exit Sub
            
            For i = 1 To MAX_INV
                If Player(Index).Inv(i).Num > 0 Then
                    If Player(Index).Inv(i).Num = Player(Index).Hotbar(Slot).Slot Then
                        UseItem Index, i
                        Exit Sub
                    End If
                End If
            Next
            
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(Index).Spell(i) > 0 Then
                    If Player(Index).Spell(i) = Player(Index).Hotbar(Slot).Slot Then
                        BufferSpell Index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 3 ' pokemon
    End Select
    
    Set Buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target
    If TempPlayer(Index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(Index).target = Index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(Index).target) Or Not IsPlaying(TempPlayer(Index).target) Then Exit Sub
    If GetPlayerMap(TempPlayer(Index).target) <> GetPlayerMap(Index) Then Exit Sub
    
    ' init the request
    Party_Invite Index, TempPlayer(Index).target
End Sub

Sub HandleAcceptParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsConnected(TempPlayer(Index).partyInvite) Or Not IsPlaying(TempPlayer(Index).partyInvite) Then
        TempPlayer(Index).partyInvite = 0
        Exit Sub
    End If
End Sub

Sub HandleDeclineParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(Index).partyInvite, Index
End Sub

Sub HandlePartyLeave(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave Index
End Sub

Sub HandleMutePlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String, Tempo As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    Tempo = Buffer.ReadLong
    
    If GetPlayerAccess(Index) < ADMIN_MONITOR Then Exit Sub
    If Name = vbNullString Then Exit Sub
    If Not IsNumeric(Tempo) Then Exit Sub
    
    If FindPlayer(Name) = Index Then
    If Tempo > 0 Then
    PlayerMsg Index, "Você não pode se calar", BrightRed
    Exit Sub
    End If
    End If

    If Tempo > 0 Then
    Player(FindPlayer(Name)).MutedTime = (Tempo * 60000) + GetTickCount
    PlayerMsg FindPlayer(Name), "Você foi proibido de falar por " & Tempo & " Minuto(s).", Yellow
    Else
    Player(FindPlayer(Name)).MutedTime = 0
    PlayerMsg FindPlayer(Name), "Você pode falar novamente!", Yellow
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleEvolCommand(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Command As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Command = Buffer.ReadLong
    
    If Command > 1 Then Exit Sub
    If TempPlayer(Index).EvolTimer > 0 Or Player(Index).EvolTimerStone > 0 Then Exit Sub
    
    If Command = 0 Then
        TempPlayer(Index).EvolTimer = 9000 + GetTickCount
        If Player(Index).Flying = 1 Then
            SetPlayerFlying Index, 0
            SendPlayerData Index
        End If
    Else
        Player(Index).EvolPermition = 0
        TempPlayer(Index).EvolTimer = 0
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleRequestEditQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SQuestEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_QUESTS Then
        Exit Sub
    End If

    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateQuestToAll(n)
    Call SaveQuest(n)
    Call AddLog(GetPlayerName(Index) & " saved Quest #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendQuests Index
End Sub

Sub HandleQuestCommand(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, Command As Byte, Value As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Command = Buffer.ReadByte
    Value = Buffer.ReadLong
    Set Buffer = Nothing
    
    Select Case Command
        Case 1
            AceitarQuest Index
        Case 2
            ChecarReqQuest Index, TempPlayer(Index).QuestSelect, Value
    End Select
End Sub

Sub HandleLeiloar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim InvNum As Long, Price As Long, LeilaoNum As Long, Tempo As Long, Tipo As Long
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    InvNum = Buffer.ReadLong
    Price = Buffer.ReadLong
    LeilaoNum = FindLeilao
    Tempo = Buffer.ReadLong
    Tipo = Buffer.ReadLong
    
    If InvNum = 0 Then Exit Sub
    
    If Item(GetPlayerInvItemNum(Index, InvNum)).NTrade = True Then
        Exit Sub
    End If
    
    If GetPlayerInvItemPokeInfoPokemon(Index, InvNum) Then
        If Player(Index).PokeQntia <= 1 Then
            PlayerMsg Index, "Você não pode Leiloar seu unico Pokémon em mãos!", BrightRed
            Exit Sub
        Else
            Player(Index).PokeQntia = Player(Index).PokeQntia - 1
        End If
    End If
    
    If LeilaoNum > 0 Then
        Leilao(LeilaoNum).Vendedor = GetPlayerName(Index)
        Leilao(LeilaoNum).ItemNum = GetPlayerInvItemNum(Index, InvNum)
        Leilao(LeilaoNum).Price = Price
        Leilao(LeilaoNum).Tipo = Tipo
        
        '#Pokemon#
        Leilao(LeilaoNum).Poke.Pokemon = GetPlayerInvItemPokeInfoPokemon(Index, InvNum)
        Leilao(LeilaoNum).Poke.Pokeball = GetPlayerInvItemPokeInfoPokeball(Index, InvNum)
        Leilao(LeilaoNum).Poke.Level = GetPlayerInvItemPokeInfoLevel(Index, InvNum)
        Leilao(LeilaoNum).Poke.EXP = GetPlayerInvItemPokeInfoExp(Index, InvNum)
        Leilao(LeilaoNum).Poke.Felicidade = GetPlayerInvItemFelicidade(Index, InvNum)
        Leilao(LeilaoNum).Poke.Sexo = GetPlayerInvItemSexo(Index, InvNum)
        Leilao(LeilaoNum).Poke.Shiny = GetPlayerInvItemShiny(Index, InvNum)
        
        For i = 1 To Vitals.Vital_Count - 1
            Leilao(LeilaoNum).Poke.Vital(i) = GetPlayerInvItemPokeInfoVital(Index, InvNum, i)
            Leilao(LeilaoNum).Poke.MaxVital(i) = GetPlayerInvItemPokeInfoMaxVital(Index, InvNum, i)
        Next
        
        For i = 1 To Stats.Stat_Count - 1
            Leilao(LeilaoNum).Poke.Stat(i) = GetPlayerInvItemPokeInfoStat(Index, InvNum, i)
        Next
        
        For i = 1 To MAX_POKE_SPELL
            Leilao(LeilaoNum).Poke.Spells(i) = GetPlayerInvItemPokeInfoSpell(Index, InvNum, i)
        Next
        
        For i = 1 To MAX_BERRYS
            Leilao(LeilaoNum).Poke.Berry(i) = GetPlayerInvItemBerry(Index, InvNum, i)
        Next
        
        '#########
        
        Select Case Tempo
            Case 1: Leilao(LeilaoNum).Tempo = 10
            Case 2: Leilao(LeilaoNum).Tempo = 20
            Case 3: Leilao(LeilaoNum).Tempo = 30
            Case 4: Leilao(LeilaoNum).Tempo = 40
            Case 5: Leilao(LeilaoNum).Tempo = 50
            Case 6: Leilao(LeilaoNum).Tempo = 60
            Case 7: Leilao(LeilaoNum).Tempo = 70
            Case 8: Leilao(LeilaoNum).Tempo = 80
            Case 9: Leilao(LeilaoNum).Tempo = 90
            Case 10: Leilao(LeilaoNum).Tempo = 100
            Case 11: Leilao(LeilaoNum).Tempo = 110
            Case 12: Leilao(LeilaoNum).Tempo = 120
        End Select
        
        SendAttLeilao
        SaveLeilão LeilaoNum
        TakeInvSlot Index, InvNum, 1
        SendInventory Index
    Else
        PlayerMsg Index, "Leilão está cheio! Tente mais tarde...", BrightRed
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleComprar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim LeilaoNum As Long, i As Long, Amount As Long, InvNum As Long, PendNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    LeilaoNum = Buffer.ReadLong
    PendNum = FindPend
    
    If Leilao(LeilaoNum).Vendedor = GetPlayerName(Index) Then
        PlayerMsg Index, "Você não pode comprar um item que já é seu!", BrightRed
        Exit Sub
    End If
    
    Select Case Leilao(LeilaoNum).Tipo
        Case 1
            For i = 1 To MAX_INV
                If GetPlayerInvItemNum(Index, i) = 1 Then ' Dollar
                    Amount = GetPlayerInvItemValue(Index, i)
                    InvNum = i
                    Exit For
                End If
            Next
        Case 2 ' Moeda
            For i = 1 To MAX_INV
                If GetPlayerInvItemNum(Index, i) = 2 Then ' Moeda
                    Amount = GetPlayerInvItemValue(Index, i)
                    InvNum = i
                    Exit For
                End If
            Next
        Case Else
    End Select
    
    If Leilao(LeilaoNum).ItemNum = 0 Then Exit Sub
    
    If Amount >= Leilao(LeilaoNum).Price Then
    
        If Leilao(LeilaoNum).Poke.Pokemon > 0 Then
        
        If Player(Index).PokeQntia <= 5 Then
            GiveInvItem Index, Leilao(LeilaoNum).ItemNum, 1, True, _
            Leilao(LeilaoNum).Poke.Pokemon, Leilao(LeilaoNum).Poke.Pokeball, _
            Leilao(LeilaoNum).Poke.Level, Leilao(LeilaoNum).Poke.EXP, _
            Leilao(LeilaoNum).Poke.Vital(1), Leilao(LeilaoNum).Poke.Vital(2), _
            Leilao(LeilaoNum).Poke.MaxVital(1), Leilao(LeilaoNum).Poke.MaxVital(2), _
            Leilao(LeilaoNum).Poke.Stat(1), Leilao(LeilaoNum).Poke.Stat(4), _
            Leilao(LeilaoNum).Poke.Stat(2), Leilao(LeilaoNum).Poke.Stat(3), _
            Leilao(LeilaoNum).Poke.Stat(5), Leilao(LeilaoNum).Poke.Spells(1), _
            Leilao(LeilaoNum).Poke.Spells(2), Leilao(LeilaoNum).Poke.Spells(3), _
            Leilao(LeilaoNum).Poke.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Leilao(LeilaoNum).Poke.Felicidade, Leilao(LeilaoNum).Poke.Sexo, Leilao(LeilaoNum).Poke.Shiny, _
            Leilao(LeilaoNum).Poke.Berry(1), Leilao(LeilaoNum).Poke.Berry(2), Leilao(LeilaoNum).Poke.Berry(3), Leilao(LeilaoNum).Poke.Berry(4), Leilao(LeilaoNum).Poke.Berry(5)
            Player(Index).PokeQntia = Player(Index).PokeQntia + 1
        Else
            DirectBankItemPokemon Index, Leilao(LeilaoNum).ItemNum, _
            Leilao(LeilaoNum).Poke.Pokemon, Leilao(LeilaoNum).Poke.Pokeball, _
            Leilao(LeilaoNum).Poke.Level, Leilao(LeilaoNum).Poke.EXP, _
            Leilao(LeilaoNum).Poke.Vital(1), Leilao(LeilaoNum).Poke.Vital(2), _
            Leilao(LeilaoNum).Poke.MaxVital(1), Leilao(LeilaoNum).Poke.MaxVital(2), _
            Leilao(LeilaoNum).Poke.Stat(1), Leilao(LeilaoNum).Poke.Stat(4), _
            Leilao(LeilaoNum).Poke.Stat(2), Leilao(LeilaoNum).Poke.Stat(3), _
            Leilao(LeilaoNum).Poke.Stat(5), Leilao(LeilaoNum).Poke.Spells(1), _
            Leilao(LeilaoNum).Poke.Spells(2), Leilao(LeilaoNum).Poke.Spells(3), _
            Leilao(LeilaoNum).Poke.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Leilao(LeilaoNum).Poke.Felicidade, Leilao(LeilaoNum).Poke.Sexo, Leilao(LeilaoNum).Poke.Shiny, _
            Leilao(LeilaoNum).Poke.Berry(1), Leilao(LeilaoNum).Poke.Berry(2), Leilao(LeilaoNum).Poke.Berry(3), Leilao(LeilaoNum).Poke.Berry(4), Leilao(LeilaoNum).Poke.Berry(5)
        End If
        
        Else
            If Item(Leilao(LeilaoNum).ItemNum).Type = ITEM_TYPE_ROD Then
                GiveInvItem Index, Leilao(LeilaoNum).ItemNum, 1, True, _
                Leilao(LeilaoNum).Poke.Pokemon, Leilao(LeilaoNum).Poke.Pokeball, _
                Leilao(LeilaoNum).Poke.Level, Leilao(LeilaoNum).Poke.EXP, _
                Leilao(LeilaoNum).Poke.Vital(1), Leilao(LeilaoNum).Poke.Vital(2), _
                Leilao(LeilaoNum).Poke.MaxVital(1), Leilao(LeilaoNum).Poke.MaxVital(2), _
                Leilao(LeilaoNum).Poke.Stat(1), Leilao(LeilaoNum).Poke.Stat(4), _
                Leilao(LeilaoNum).Poke.Stat(2), Leilao(LeilaoNum).Poke.Stat(3), _
                Leilao(LeilaoNum).Poke.Stat(5), Leilao(LeilaoNum).Poke.Spells(1), _
                Leilao(LeilaoNum).Poke.Spells(2), Leilao(LeilaoNum).Poke.Spells(3), _
                Leilao(LeilaoNum).Poke.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Leilao(LeilaoNum).Poke.Felicidade, Leilao(LeilaoNum).Poke.Sexo, Leilao(LeilaoNum).Poke.Shiny, _
                Leilao(LeilaoNum).Poke.Berry(1), Leilao(LeilaoNum).Poke.Berry(2), Leilao(LeilaoNum).Poke.Berry(3), Leilao(LeilaoNum).Poke.Berry(4), Leilao(LeilaoNum).Poke.Berry(5)
            Else
                GiveInvItem Index, Leilao(LeilaoNum).ItemNum, 1
            End If
        End If
        
        TakeInvItem Index, GetPlayerInvItemNum(Index, InvNum), Leilao(LeilaoNum).Price
        
        If IsPlaying(FindPlayer(Leilao(LeilaoNum).Vendedor)) = True Then
            GiveInvItem FindPlayer(Leilao(LeilaoNum).Vendedor), GetPlayerInvItemNum(Index, InvNum), Leilao(LeilaoNum).Price
            PlayerMsg FindPlayer(Leilao(LeilaoNum).Vendedor), "O item " & Trim$(Item(Leilao(LeilaoNum).ItemNum).Name) & " foi vendido com sucesso pelo preço " & Leilao(LeilaoNum).Price, BrightGreen
        Else
            Pendencia(PendNum).Vendedor = Leilao(LeilaoNum).Vendedor
            Pendencia(PendNum).ItemNum = GetPlayerInvItemNum(Index, InvNum)
            Pendencia(PendNum).Price = Leilao(LeilaoNum).Price
            Pendencia(PendNum).Tipo = Leilao(LeilaoNum).Tipo
            SavePendencia PendNum
        End If
        
        LimparLeilaoSlot LeilaoNum
        SaveLeilão LeilaoNum
        ArrumaLeilao
        SendAttLeilao
        PlayerMsg Index, "Item comprado com sucesso!", BrightGreen
    Else
        PlayerMsg Index, "Você não tem dinheiro suficiente!", BrightRed
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleRetirar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim LeilaoNum As Long, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    LeilaoNum = Buffer.ReadLong
    
    If Trim$(Leilao(LeilaoNum).Vendedor) = GetPlayerName(Index) Then
        
        If Leilao(LeilaoNum).Poke.Pokemon > 0 Then
        
        If Player(Index).PokeQntia <= 5 Then
            GiveInvItem Index, Leilao(LeilaoNum).ItemNum, 1, True, _
            Leilao(LeilaoNum).Poke.Pokemon, Leilao(LeilaoNum).Poke.Pokeball, _
            Leilao(LeilaoNum).Poke.Level, Leilao(LeilaoNum).Poke.EXP, _
            Leilao(LeilaoNum).Poke.Vital(1), Leilao(LeilaoNum).Poke.Vital(2), _
            Leilao(LeilaoNum).Poke.MaxVital(1), Leilao(LeilaoNum).Poke.MaxVital(2), _
            Leilao(LeilaoNum).Poke.Stat(1), Leilao(LeilaoNum).Poke.Stat(4), _
            Leilao(LeilaoNum).Poke.Stat(2), Leilao(LeilaoNum).Poke.Stat(3), _
            Leilao(LeilaoNum).Poke.Stat(5), Leilao(LeilaoNum).Poke.Spells(1), _
            Leilao(LeilaoNum).Poke.Spells(2), Leilao(LeilaoNum).Poke.Spells(3), _
            Leilao(LeilaoNum).Poke.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Leilao(LeilaoNum).Poke.Felicidade, Leilao(LeilaoNum).Poke.Sexo, Leilao(LeilaoNum).Poke.Shiny, _
            Leilao(LeilaoNum).Poke.Berry(1), Leilao(LeilaoNum).Poke.Berry(2), Leilao(LeilaoNum).Poke.Berry(3), Leilao(LeilaoNum).Poke.Berry(4), Leilao(LeilaoNum).Poke.Berry(5)
            Player(Index).PokeQntia = Player(Index).PokeQntia + 1
        Else
            DirectBankItemPokemon Index, Leilao(LeilaoNum).ItemNum, _
            Leilao(LeilaoNum).Poke.Pokemon, Leilao(LeilaoNum).Poke.Pokeball, _
            Leilao(LeilaoNum).Poke.Level, Leilao(LeilaoNum).Poke.EXP, _
            Leilao(LeilaoNum).Poke.Vital(1), Leilao(LeilaoNum).Poke.Vital(2), _
            Leilao(LeilaoNum).Poke.MaxVital(1), Leilao(LeilaoNum).Poke.MaxVital(2), _
            Leilao(LeilaoNum).Poke.Stat(1), Leilao(LeilaoNum).Poke.Stat(4), _
            Leilao(LeilaoNum).Poke.Stat(2), Leilao(LeilaoNum).Poke.Stat(3), _
            Leilao(LeilaoNum).Poke.Stat(5), Leilao(LeilaoNum).Poke.Spells(1), _
            Leilao(LeilaoNum).Poke.Spells(2), Leilao(LeilaoNum).Poke.Spells(3), _
            Leilao(LeilaoNum).Poke.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Leilao(LeilaoNum).Poke.Felicidade, Leilao(LeilaoNum).Poke.Sexo, Leilao(LeilaoNum).Poke.Shiny, _
            Leilao(LeilaoNum).Poke.Berry(1), Leilao(LeilaoNum).Poke.Berry(2), Leilao(LeilaoNum).Poke.Berry(3), Leilao(LeilaoNum).Poke.Berry(4), Leilao(LeilaoNum).Poke.Berry(5)
            PlayerMsg Index, "Você já tem 6 Pokémon em mãos o pokémon foi enviado para o Computador!", White
        End If
        
        Else
        
        If Item(Leilao(LeilaoNum).ItemNum).Type = ITEM_TYPE_ROD Then
            GiveInvItem Index, Leilao(LeilaoNum).ItemNum, 1, True, _
            Leilao(LeilaoNum).Poke.Pokemon, Leilao(LeilaoNum).Poke.Pokeball, _
            Leilao(LeilaoNum).Poke.Level, Leilao(LeilaoNum).Poke.EXP, _
            Leilao(LeilaoNum).Poke.Vital(1), Leilao(LeilaoNum).Poke.Vital(2), _
            Leilao(LeilaoNum).Poke.MaxVital(1), Leilao(LeilaoNum).Poke.MaxVital(2), _
            Leilao(LeilaoNum).Poke.Stat(1), Leilao(LeilaoNum).Poke.Stat(4), _
            Leilao(LeilaoNum).Poke.Stat(2), Leilao(LeilaoNum).Poke.Stat(3), _
            Leilao(LeilaoNum).Poke.Stat(5), Leilao(LeilaoNum).Poke.Spells(1), _
            Leilao(LeilaoNum).Poke.Spells(2), Leilao(LeilaoNum).Poke.Spells(3), _
            Leilao(LeilaoNum).Poke.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Leilao(LeilaoNum).Poke.Felicidade, Leilao(LeilaoNum).Poke.Sexo, Leilao(LeilaoNum).Poke.Shiny, _
            Leilao(LeilaoNum).Poke.Berry(1), Leilao(LeilaoNum).Poke.Berry(2), Leilao(LeilaoNum).Poke.Berry(3), Leilao(LeilaoNum).Poke.Berry(4), Leilao(LeilaoNum).Poke.Berry(5)
        Else
            GiveInvItem Index, Leilao(LeilaoNum).ItemNum, 1
        End If
        
        End If
        
        Leilao(LeilaoNum).Vendedor = vbNullString
        Leilao(LeilaoNum).ItemNum = 0
        Leilao(LeilaoNum).Price = 0
        Leilao(LeilaoNum).Tempo = 0
        Leilao(LeilaoNum).Poke.Pokemon = 0
        Leilao(LeilaoNum).Poke.Pokeball = 0
        Leilao(LeilaoNum).Poke.Level = 0
        Leilao(LeilaoNum).Poke.EXP = 0
        
        For i = 1 To Vitals.Vital_Count - 1
            Leilao(LeilaoNum).Poke.Vital(i) = 0
            Leilao(LeilaoNum).Poke.MaxVital(i) = 0
        Next
        
        For i = 1 To Stats.Stat_Count - 1
            Leilao(LeilaoNum).Poke.Stat(i) = 0
        Next
        
        For i = 1 To MAX_POKE_SPELL
            Leilao(LeilaoNum).Poke.Spells(i) = 0
        Next
        
        For i = 1 To MAX_BERRYS
            Leilao(LeilaoNum).Poke.Berry(i) = 0
        Next
        
        SaveLeilão LeilaoNum
        ArrumaLeilao
        
        SendAttLeilao
        
        PlayerMsg Index, "Você retirou seu item do leilão!", BrightGreen
        SendInventory Index
    Else
        PlayerMsg Index, "Você não pode retirar um item que não é seu!", BrightRed
    End If

    Set Buffer = Nothing
End Sub

Sub ArrumaLeilao()
Dim i As Long, X As Long

    For i = 1 To MAX_LEILAO
        If i = 20 Then Exit For
        If Leilao(i).Vendedor = vbNullString And Leilao(i + 1).Vendedor <> vbNullString Then
            Leilao(i).Vendedor = Leilao(i + 1).Vendedor
            Leilao(i + 1).Vendedor = vbNullString
            Leilao(i).ItemNum = Leilao(i + 1).ItemNum
            Leilao(i + 1).ItemNum = 0
            Leilao(i).Price = Leilao(i + 1).Price
            Leilao(i + 1).Price = 0
            Leilao(i).Tempo = Leilao(i + 1).Tempo
            Leilao(i + 1).Tempo = 0
            Leilao(i).Tipo = Leilao(i + 1).Tipo
            Leilao(i + 1).Tipo = 0
            '####Pokemon###
            Leilao(i).Poke.Pokemon = Leilao(i + 1).Poke.Pokemon
            Leilao(i + 1).Poke.Pokemon = 0
            
            Leilao(i).Poke.Pokeball = Leilao(i + 1).Poke.Pokeball
            Leilao(i + 1).Poke.Pokeball = 0
            
            Leilao(i).Poke.Level = Leilao(i + 1).Poke.Level
            Leilao(i + 1).Poke.Level = 0
            
            Leilao(i).Poke.EXP = Leilao(i + 1).Poke.EXP
            Leilao(i + 1).Poke.EXP = 0
            
            Leilao(i).Poke.Felicidade = Leilao(i + 1).Poke.Felicidade
            Leilao(i + 1).Poke.Felicidade = 0
            
            Leilao(i).Poke.Sexo = Leilao(i + 1).Poke.Sexo
            Leilao(i + 1).Poke.Sexo = 0
            
            Leilao(i).Poke.Shiny = Leilao(i + 1).Poke.Shiny
            Leilao(i + 1).Poke.Shiny = 0
            
            For X = 1 To Vitals.Vital_Count - 1
            Leilao(i).Poke.Vital(X) = Leilao(i + 1).Poke.Vital(X)
            Leilao(i + 1).Poke.Vital(X) = 0
            
            Leilao(i).Poke.MaxVital(X) = Leilao(i + 1).Poke.MaxVital(X)
            Leilao(i + 1).Poke.MaxVital(X) = 0
            Next
            
            For X = 1 To Stats.Stat_Count - 1
            Leilao(i).Poke.Stat(X) = Leilao(i + 1).Poke.Stat(X)
            Leilao(i + 1).Poke.Stat(X) = 0
            Next
            
            For X = 1 To MAX_POKE_SPELL
            Leilao(i).Poke.Spells(X) = Leilao(i + 1).Poke.Spells(X)
            Leilao(i + 1).Poke.Spells(X) = 0
            Next
            
        End If
    Next
End Sub

Public Sub LimparLeilaoSlot(ByVal SlotNum As Long)
Dim i As Long

    Leilao(SlotNum).Vendedor = vbNullString
    Leilao(SlotNum).ItemNum = 0
    Leilao(SlotNum).Price = 0
    Leilao(SlotNum).Tipo = 0
    
    'PokeClear
    Leilao(SlotNum).Poke.Pokemon = 0
    Leilao(SlotNum).Poke.Pokeball = 0
    Leilao(SlotNum).Poke.Level = 0
    Leilao(SlotNum).Poke.EXP = 0
    Leilao(SlotNum).Poke.Felicidade = 0
    Leilao(SlotNum).Poke.Sexo = 0
    Leilao(SlotNum).Poke.Shiny = 0
    
    For i = 1 To Vitals.Vital_Count - 1
        Leilao(SlotNum).Poke.Vital(i) = 0
        Leilao(SlotNum).Poke.MaxVital(i) = 0
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Leilao(SlotNum).Poke.Stat(i) = 0
    Next
    
    For i = 1 To MAX_POKE_SPELL
        Leilao(SlotNum).Poke.Spells(i) = 0
    Next
    
End Sub

Sub HandleChatComando(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, C, D As Long, S As String
Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    C = Buffer.ReadLong
    If C = 3 Then
        S = FindPlayer(Buffer.ReadString)
    ElseIf C = 4 Then
        S = Buffer.ReadString
    End If
    
Select Case C
Case 1 'Aceitar

    If TempPlayer(Index).Conversando > 0 Then
        Call SendChat(5, Index, 0, vbNullString)
        Call SendChat(5, TempPlayer(Index).Conversando, 0, vbNullString)
    Else
        PlayerMsg Index, "O jogador selecionado está deslogado no momento!", Red
    End If

    TempPlayer(Index).ConversandoC = 1
    TempPlayer(TempPlayer(Index).Conversando).ConversandoC = 1

    SendChat 4, TempPlayer(Index).Conversando, Index, vbNullString
Case 2 'Recusar

    If TempPlayer(Index).Conversando > 0 Then
        PlayerMsg TempPlayer(Index).Conversando, "Seu pedido de chat privado enviado para: " & GetPlayerName(Index) & " , foi recusado.", BrightRed
        PlayerMsg Index, "Pedido de chat privado enviado por: " & GetPlayerName(TempPlayer(Index).Conversando) & " foi recusado , com sucesso!!", BrightRed
    Else
        PlayerMsg Index, "O jogador selecionado está deslogado no momento!", Red
    End If
    TempPlayer(Index).Conversando = 0
    SendChat 2, TempPlayer(Index).Conversando, Index, vbNullString
Case 3 ' Convidar

    If TempPlayer(Index).ConversandoC > 0 Then
        PlayerMsg Index, "Você já está conversando com " & GetPlayerName(TempPlayer(Index).Conversando), BrightRed
        Exit Sub
    End If
     
    If TempPlayer(S).ConversandoC > 0 Then
        PlayerMsg Index, "O jogador escolhido já está conversando com " & GetPlayerName(TempPlayer(S).Conversando), BrightRed
        Exit Sub
    End If
     
    If S = 0 Then
        PlayerMsg Index, "O Jogador selecionado não está online!", BrightRed
        Exit Sub
    End If
    
    If GetPlayerName(Index) = GetPlayerName(S) Then
        PlayerMsg Index, "Você não pode convidar você mesmo, para o chat privado!", BrightRed
        Exit Sub
    End If
    
    TempPlayer(Index).Conversando = S
    TempPlayer(S).Conversando = Index
    SendChat 1, TempPlayer(Index).Conversando, Index, vbNullString
    PlayerMsg Index, "Seu convite de chat privado foi enviado com sucesso!", BrightRed
Case 4 ' Enviar Menssagem
    
    If TempPlayer(Index).Conversando > 0 Then
        Call SendChat(6, TempPlayer(Index).Conversando, Index, S)
        Call SendChat(7, Index, Index, S)
    Else
        PlayerMsg Index, "o jogador está off e não pode responder!", BrightRed
    End If
Case 5 ' Fechar meu Chat
    If TempPlayer(Index).Conversando > 0 Then
        SendChat 2, TempPlayer(Index).Conversando, Index, vbNullString
        TempPlayer(Index).Conversando = 0
        TempPlayer(Index).ConversandoC = 0
    Else
        PlayerMsg Index, "o jogador está off e não pode responder!", BrightRed
    End If
Case 6 ' Fechar Chat do meu parceiro
    If TempPlayer(Index).Conversando > 0 Then
       TempPlayer(Index).Conversando = 0
       TempPlayer(Index).ConversandoC = 0
       SendChat 3, Index, Index, vbNullString
    Else
       PlayerMsg Index, "o jogador está off e não pode responder!", BrightRed
    End If
End Select
Set Buffer = Nothing
End Sub

Sub HandleSelectPoke(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim PokeSelect As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PokeSelect = Buffer.ReadByte
    
    If Player(Index).PokeInicial = 1 Then
    
    Select Case PokeSelect
    Case 1 'Bulbasaur
        GiveInvItem Index, 3, 1, True, 1, 1, 5, 1, 65, 65, 65, 65, 61, 57, 61, 77, 77, 0, 0, 0, 0
        Player(Index).Pokedex(1) = 1
        SendPlayerPokedex Index
    Case 2 'Charmander
        GiveInvItem Index, 3, 1, True, 4, 1, 5, 1, 59, 59, 59, 59, 64, 77, 55, 72, 62, 0, 0, 0, 0
        Player(Index).Pokedex(4) = 1
        SendPlayerPokedex Index
    Case 3 'Squirtle
        GiveInvItem Index, 3, 1, True, 7, 1, 5, 1, 64, 64, 64, 64, 60, 55, 77, 62, 78, 0, 0, 0, 0
        Player(Index).Pokedex(7) = 1
        SendPlayerPokedex Index
        
    Case 4 'Pikachu
        GiveInvItem Index, 3, 1, True, 25, 1, 5, 1, 35, 35, 35, 35, 55, 90, 50, 50, 55, 0, 0, 0, 0
        Player(Index).Pokedex(25) = 1
        SendPlayerPokedex Index
    End Select
    
    Player(Index).PokeQntia = Player(Index).PokeQntia + 1
    Player(Index).PokeInicial = 0
    GiveInvItem Index, 22, 1
    GiveInvItem Index, 3, 5
    
    'Quest Completar a Pokédex
    TempPlayer(Index).QuestInvite = 100
    AceitarQuest Index
    
    Call PlayerMsg(Index, "[Profº Oak]: Aqui está seu pokémon e sua Pokédex na qual dou a missão de você completar com as informações de 251 pokémons.", White)
    End If
    
    Set Buffer = Nothing
    
End Sub

Sub HandleSendSurfInit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Command As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Command = Buffer.ReadLong
    
    If Player(Index).InSurf <> 3 Then
        Player(Index).InSurf = 0
        SendSurfInit Index
        Exit Sub
    End If
    
    If Command = 1 Then
    If CanSurfPokemonInv(Index) = True Then
        Player(Index).InSurf = 1
        SendSurfInit Index
        ForcePlayerMove Index, MOVING_WALKING, TempPlayer(Index).SurfSlideTo
        ForcePlayerMove Index, MOVING_WALKING, TempPlayer(Index).SurfSlideTo
    Else
        PlayerMsg Index, "Você não possui nenhum pokémon com Habilidade Surf", BrightRed
        Player(Index).InSurf = 0
        SendSurfInit Index
    End If
    
    Else
    
    Player(Index).InSurf = 0
        SendSurfInit Index
    End If
    
    Set Buffer = Nothing
    
End Sub

Function CanSurfPokemonInv(ByVal Index As Long) As Boolean
Dim i As Long

CanSurfPokemonInv = False

    For i = 1 To MAX_INV
    
        If Player(Index).Inv(i).Num = 3 Then
            If Player(Index).Inv(i).PokeInfo.Pokemon > 0 Then
                'If Pokemon(Player(Index).Inv(i).PokeInfo.Pokemon).FRS = 3 Then
                    CanSurfPokemonInv = True
                    Exit For
                'End If
            End If
        End If
    
    Next

End Function

Sub handleLutarComando(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, C, T, A, i, Pok As Long, p As String
Dim Desafiado As Long, Numero As Byte

Set Buffer = New clsBuffer

Buffer.WriteBytes Data()

C = Buffer.ReadLong
T = Buffer.ReadLong
A = Buffer.ReadLong
Pok = Buffer.ReadLong
p = FindPlayer(Buffer.ReadString)

Select Case C
    Case 1
    
        If p = 0 Then
            PlayerMsg Index, "O Jogador selecionado está offline", BrightRed
            Exit Sub
        End If
        
        If TempPlayer(Index).Lutando > 0 Then
            PlayerMsg Index, "Não pode lutar com dois ao mesmo tempo", BrightRed
            Exit Sub
        End If
        
        If TempPlayer(p).Lutando > 0 Then
            PlayerMsg Index, "O Jogador está lutando contra: " & GetPlayerName(TempPlayer(p).Lutando) & ", no momento", BrightRed
            Exit Sub
        End If
       
        If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) = 0 Then
            PlayerMsg Index, "Solte o Pokémon que você vai usar no Duelo!", BrightRed
            Exit Sub
        End If
        
        If GetPlayerEquipmentPokeInfoPokemon(p, weapon) = 0 Then
            PlayerMsg p, "Solte o Pokémon que você vai usar no Duelo!", BrightRed
            PlayerMsg Index, "Jogador não está com o pokémon que vai usar no Duelo.", BrightRed
            Exit Sub
        End If
        
        'Verificar se Você tem Quantia de Pokémons...
        If PokDispBattle(Index, Numero) = False Then
            PlayerMsg Index, "Você não possui " & Numero & " Pokémon(s) para batalhar!", BrightRed
            Exit Sub
        End If
        
        'Verificar se Desafiado tem Quantia de Pokémons...
        If PokDispBattle(p, Numero) = False Then
            PlayerMsg Index, "O jogador " & Trim$(GetPlayerName(p)) & " não possui " & Numero & " Pokémon(s) para batalhar!", BrightRed
            PlayerMsg p, "Você não possui " & Numero & " Pokémon(s) para batalhar!", BrightRed
            Exit Sub
        End If
        
        'Confirmação de Envio
        PlayerMsg Index, "Convite Enviado", BrightGreen
        
        Select Case A
            Case 1
                For i = 1 To Player_HighIndex
                    If Player(i).Map = 90 Then
                        PlayerMsg Index, "Arena 1 Ocupada", BrightRed
                        Exit Sub
                    End If
                Next i
            Case 2
                For i = 1 To Player_HighIndex
                    If Player(i).Map = 3 Then
                        PlayerMsg Index, "Arena 2 Ocupada", BrightRed
                        Exit Sub
                    End If
                Next i
        End Select
        
        Select Case T
            Case 0
                Select Case Pok
                Case 0
                    Numero = 1
                Case 1
                    Numero = 3
                Case 2
                    Numero = 6
                End Select
            
                TempPlayer(Index).Lutando = p
                TempPlayer(Index).LutandoA = A
                TempPlayer(Index).LutandoT = 1
                TempPlayer(Index).LutQntPoke = Numero - 1
                
                TempPlayer(p).Lutando = Index
                TempPlayer(p).LutandoA = A
                TempPlayer(p).LutandoT = 1
                TempPlayer(p).LutQntPoke = Numero - 1
                
                SendLutarComando 1, T, p, Index, A, Numero
            Case 1
                If TempPlayer(Index).inParty > 0 Then
                    Else: PlayerMsg Index, "Precisa estar em grupo para esse modo", BrightRed
                    Exit Sub
                End If
                
                If Party(TempPlayer(Index).inParty).Leader = Index Then
                    Else: PlayerMsg Index, "Somente o líder pode iniciar esse modo", BrightRed
                    Exit Sub
                End If
                
                If TempPlayer(Index).inParty > 0 Then
                    Else: PlayerMsg Index, "O Jogador selecionado não esta em grupo", BrightRed
                    Exit Sub
                End If
                
                If Party(TempPlayer(Index).inParty).Leader = Index Then
                    Else: PlayerMsg Index, "O jogador selecionado não e o lider do grupo", BrightRed
                    Exit Sub
                End If
            
                TempPlayer(Index).Lutando = p
                TempPlayer(Index).Lutando = A
                TempPlayer(Index).Lutando = 2
                
                TempPlayer(p).Lutando = Index
                TempPlayer(p).LutandoA = A
                TempPlayer(p).LutandoT = T
                
                Case 2
                
                ' futura organização guild sei la
                
                End Select
                
Case 2

Select Case TempPlayer(Index).LutandoT
    Case 1
        If TempPlayer(Index).Lutando > 0 Then
            Player(Index).Dir = DIR_LEFT
            Player(TempPlayer(Index).Lutando).Dir = DIR_RIGHT
            GlobalMsg "" & GetPlayerName(Index) & " Vs " & GetPlayerName(TempPlayer(Index).Lutando) & " estão lutando na arena: " & TempPlayer(Index).LutandoA, BrightCyan
            GlobalMsg "Arena: " & TempPlayer(Index).LutandoA & " oculpada!", BrightCyan
            SendArenaStatus TempPlayer(Index).LutandoA, 1
            
            Desafiado = TempPlayer(Index).Lutando
            
            'Salvar ponto de Retorno...
            Player(Index).MyMap(1) = GetPlayerMap(Index)
            Player(Index).MyMap(2) = GetPlayerX(Index)
            Player(Index).MyMap(3) = GetPlayerY(Index)
                    
            Player(Desafiado).MyMap(1) = GetPlayerMap(Desafiado)
            Player(Desafiado).MyMap(2) = GetPlayerX(Desafiado)
            Player(Desafiado).MyMap(3) = GetPlayerY(Desafiado)
            
            Select Case TempPlayer(Index).LutandoA 'Arena...
                Case 1
                    'Teleportar para a Arena 1 Mapa 90
                    PlayerWarp Index, 90, 20, 9
                    PlayerWarp Desafiado, 90, 4, 9
                    
                    'Trainer Point Index/Desafiado
                    Player(Index).TPX = 21
                    Player(Index).TPY = 9
                    Player(Index).TPDir = DIR_LEFT
                    
                    Player(Desafiado).TPX = 3
                    Player(Desafiado).TPY = 9
                    Player(Desafiado).TPDir = DIR_RIGHT
                Case 2
                    'Teleportar para a Arena 1 Mapa 91
                    PlayerWarp Index, 91, 12, 6
                    PlayerWarp Desafiado, 91, 12, 14
                    
                    'Trainer Point Index/Desafiado
                    Player(Index).TPX = 12
                    Player(Index).TPY = 5
                    Player(Index).TPDir = DIR_DOWN
                    
                    Player(Desafiado).TPX = 12
                    Player(Desafiado).TPY = 15
                    Player(Desafiado).TPDir = DIR_UP
            End Select
        Else
            PlayerMsg Index, "O jogador está offline", BrightRed
            TempPlayer(Index).Lutando = 0
            TempPlayer(Index).LutandoA = 0
            TempPlayer(Index).LutandoT = 0
            TempPlayer(Index).LutQntPoke = 0
        End If

    Case 2
        If TempPlayer(Index).Lutando > 0 And TempPlayer(Index).inParty > 0 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).inParty = TempPlayer(Index).inParty Then
                        Party(TempPlayer(Index).inParty).PT = Party(TempPlayer(Index).inParty).MemberCount
                        Select Case TempPlayer(Index).LutandoA
                            Case 1
                                PlayerWarp i, 2, 5, 5
                            Case 2
                                PlayerWarp i, 3, 15, 15
                        End Select
                    End If
                End If
            Next
            
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If TempPlayer(i).inParty = TempPlayer(TempPlayer(Index).Lutando).inParty Then
                    Party(TempPlayer(TempPlayer(Index).Lutando).inParty).PT = Party(TempPlayer(TempPlayer(Index).Lutando).inParty).MemberCount
                    Select Case TempPlayer(Index).LutandoA
                        Case 1
                            PlayerWarp i, 2, 5, 5
                        Case 2
                            PlayerWarp i, 3, 15, 15
                    End Select
                End If
            End If
        Next
            GlobalMsg "O grupo dos jogadores " & GetPlayerName(Index) & " Vs " & GetPlayerName(TempPlayer(Index).Lutando) & " estão lutando na arena: " & TempPlayer(Index).LutandoA, BrightCyan
            GlobalMsg "Arena: " & TempPlayer(Index).LutandoA & " ocupada!", BrightCyan
            Else
                PlayerMsg Index, "O jogador que você convidou para a luta está off no momento", BrightCyan
                TempPlayer(Index).Lutando = 0
                TempPlayer(Index).LutandoA = 0
                TempPlayer(Index).LutandoT = 0
            End If
            
    Case 3
       
           ' em breve
        End Select
     
Case 3 'Recusar
    If TempPlayer(Index).Lutando > 0 Then
        PlayerMsg Index, "Convite de luta recusado com sucesso!", BrightCyan
        PlayerMsg TempPlayer(Index).Lutando, "Seu convite de luta enviado para: " & GetPlayerName(TempPlayer(Index).Lutando) & " , foi recusado!!", BrightCyan
        
        'Limpar
        TempPlayer(TempPlayer(Index).Lutando).Lutando = 0
        TempPlayer(TempPlayer(Index).Lutando).LutandoA = 0
        TempPlayer(TempPlayer(Index).Lutando).LutandoT = 0
        
        TempPlayer(Index).Lutando = 0
        TempPlayer(Index).LutandoA = 0
        TempPlayer(Index).LutandoT = 0
    Else
        TempPlayer(TempPlayer(Index).Lutando).Lutando = 0
        TempPlayer(TempPlayer(Index).Lutando).LutandoA = 0
        TempPlayer(TempPlayer(Index).Lutando).LutandoT = 0
    
        TempPlayer(Index).Lutando = 0
        TempPlayer(Index).LutandoA = 0
        TempPlayer(Index).LutandoT = 0
    End If

    Case 4 'Acabar luta
       ' GlobalMsg "O jogador: " & GetPlayerName(TempPlayer(Index).Lutando) & " , desistiu da luta contra: " & GetPlayerName(Index), BrightCyan
        PlayerWarp Index, 350, 11, 8
        
        TempPlayer(TempPlayer(Index).Lutando).Lutando = 0
        TempPlayer(TempPlayer(Index).Lutando).LutandoA = 0
        TempPlayer(TempPlayer(Index).Lutando).LutandoT = 0
        
        TempPlayer(Index).Lutando = 0
        TempPlayer(Index).LutandoA = 0
        TempPlayer(Index).LutandoT = 0
    End Select
    
    Set Buffer = Nothing
End Sub

Private Sub handleAprenderHab(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Action As Byte, i As Long
' ???
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Action = Buffer.ReadByte
    
    If Player(Index).LearnSpell(1) = 0 Then
        PlayerMsg Index, "Fail", BrightRed
        Exit Sub
    End If
    
    Select Case Action
    Case 1 To 4
        Call SetPlayerEquipmentPokeInfoSpell(Index, Player(Index).LearnSpell(2), weapon, Action)
        Call SetPlayerSpell(Index, Action, Player(Index).LearnSpell(2))
        
        If Player(Index).LearnSpell(3) > 0 Then
            TakeInvItem Index, Player(Index).LearnSpell(3), 1
        End If
        
        For i = 1 To 3
            Player(Index).LearnSpell(i) = 0
        Next
        
        For i = 1 To 10
            If Player(Index).LearnFila(i) > 0 Then
                Player(Index).LearnSpell(1) = 1
                Player(Index).LearnSpell(2) = Player(Index).LearnFila(i)
                Player(Index).LearnFila(i) = 0
                SendAprenderSpell Index, 0
                Exit For
            End If
        Next
        
        SendPlayerSpells Index
        Call SendWornEquipment(Index)
        Call SendMapEquipment(Index)
    Case 5
        For i = 1 To 3
            Player(Index).LearnSpell(i) = 0
        Next
        
        For i = 1 To 10
            If Player(Index).LearnFila(i) > 0 Then
                Player(Index).LearnSpell(1) = 1
                Player(Index).LearnSpell(2) = Player(Index).LearnFila(i)
                Player(Index).LearnFila(i) = 0
                SendAprenderSpell Index, 0
                Exit For
            End If
        Next
        
        SendPlayerSpells Index
        Call SendWornEquipment(Index)
        Call SendMapEquipment(Index)
    End Select
    
    Set Buffer = Nothing

End Sub

Sub HandleSetOrg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim U As String
    Dim n As Long
    Dim i As Long
    Dim l As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    i = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If IsPlaying(n) = False Then Exit Sub
    Select Case i
        Case 1
            U = "Equipe 1"
        Case 2
            U = "Team 2"
        Case 3
            U = "Team 3"
        Case 4
            U = "Team 4"
        Case 0
            U = "Nothing"
    Case Else
        Exit Sub
    End Select
    
    'Setar Valor 0
    If i = 0 Then
        If Player(n).ORG > 0 Then
        
            For l = 1 To MAX_ORG_MEMBERS
                If Trim$(Organization(Player(n).ORG).OrgMember(l).User_Name) = Trim$(GetPlayerName(n)) Then
                    ClearOrgMemberSlot Player(n).ORG, l
                    PlayerMsg n, "Você saiu da Organização", BrightRed
                End If
            Next
            
            Player(n).ORG = 0
            GoTo Continue
        Else
            Player(n).ORG = 0
            PlayerMsg n, "Sua Organização foi Setada para 0!", BrightRed
            GoTo Continue
        End If
    End If
    
    'Caso já tenha organização!
    If Player(n).ORG > 0 Then
        If n <> Index Then
            PlayerMsg Index, "Jogador " & Trim$(GetPlayerName(n)) & " já possui uma organização!", BrightRed
        End If
            PlayerMsg n, "Você já está em uma organização", BrightRed
        Exit Sub
    End If
    
    'Verificar Vaga e Setar numero de Membro na organização!
    l = FindOpenOrgMemberSlot(i)
    Select Case FindOpenOrgMemberSlot(i)
        Case 0
            PlayerMsg Index, "Não há vagas na organização: " & U & "!", BrightRed
            PlayerMsg n, "Não há vagas na organização: " & U & "!", BrightRed
            Exit Sub
        Case 1
            Player(n).ORG = i
            Organization(i).Lider = Trim$(GetPlayerName(n))
            Organization(i).OrgMember(l).Used = True
            Organization(i).OrgMember(l).User_Login = Trim$(GetPlayerLogin(n))
            Organization(i).OrgMember(l).User_Name = Trim$(GetPlayerName(n))
            Organization(i).OrgMember(l).Online = True
            PlayerMsg n, "Você é o lider da organização: " & U, BrightCyan
        Case Else
            Organization(i).OrgMember(l).Used = True
            Organization(i).OrgMember(l).User_Login = Trim$(GetPlayerLogin(n))
            Organization(i).OrgMember(l).User_Name = Trim$(GetPlayerName(n))
            Organization(i).OrgMember(l).Online = True
            Player(n).ORG = i
            GlobalMsg GetPlayerName(n) & " acaba de entrar para organização:  " & U & "!", BrightCyan
    End Select
    
Continue:
    For l = 1 To Player_HighIndex
        If IsPlaying(l) = True Then
            If Player(l).ORG = Player(n).ORG Then
            Call SendOrganização(l, True)
            End If
        End If
    Next

    Call SendOrganização(n, True)
    SaveOrg Player(n).ORG
    SendPlayerData n
    SavePlayer n
End Sub

Sub HandleAbrir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(Index).ORG > 0 Then
        Call SendOrganização(Index)
    End If
End Sub

Sub HandleBuyOrgShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim OrgShopSlot As Byte
Dim Quantia As Long

    'Receber Dados do Cliente
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    OrgShopSlot = Buffer.ReadByte
    Quantia = Buffer.ReadLong
    Set Buffer = Nothing
    
    'Para não haver problemas na entrega do item
    If Quantia = 0 Then Quantia = 1
    
    'Para Itens não currency!
    If Not Item(OrgShop(OrgShopSlot).Item).Type = ITEM_TYPE_CURRENCY Then Quantia = 1
    
    'Evitar OverFlow
    If OrgShopSlot = 0 Or OrgShopSlot > MAX_ORG_SHOP Then Exit Sub
        
    'Jogador sem Organização
    If Player(Index).ORG = 0 Then
        PlayerMsg Index, "Você não faz parte de nenhuma organização!", BrightRed
        Exit Sub
    End If
        
    'Slot Vazio
    If OrgShop(OrgShopSlot).Item = 0 Or OrgShop(OrgShopSlot).Item > MAX_ITEMS Then
        PlayerMsg Index, "OrgShopSlot Vazio.", White
        Exit Sub
    End If
    
    'Sem Org Level Suficiente
    If Organization(Player(Index).ORG).Level < OrgShop(OrgShopSlot).Level Then
        PlayerMsg Index, "Organização abaixo do level requerido!", BrightRed
        Exit Sub
    End If
    
    'Sem Honra Suficiente
    If Player(Index).Honra < OrgShop(OrgShopSlot).Valor Then
        PlayerMsg Index, "Você não possui pontos de Honra o Suficiente para comprar este Item!", BrightRed
        Exit Sub
    End If
    
    'Não bugar Quantia de itens Currency!
    If Item(OrgShop(OrgShopSlot).Item).Type = ITEM_TYPE_CURRENCY Then
    If OrgShop(OrgShopSlot).Quantia = 0 Then OrgShop(OrgShopSlot).Quantia = 1
    End If
    
    'Comprar Caso esteja tudo de Acordo!
    GiveInvItem Index, OrgShop(OrgShopSlot).Item, Quantia * OrgShop(OrgShopSlot).Quantia
    PlayerMsg Index, "Você comprou o item " & Trim$(Item(OrgShop(OrgShopSlot).Item).Name) & " pelo preço de " & (Quantia * OrgShop(OrgShopSlot).Valor) & " pontos de Honra!", BrightGreen
    Call SetPlayerHonra(Index, GetPlayerHonra(Index) - (Quantia * OrgShop(OrgShopSlot).Valor))
    SendPlayerData Index
    
End Sub

Sub HandleRecoverPass(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Account As String, RecoveryKey As String, Email As String
Dim NovaSenha As Long

    'Receber Dados do Cliente
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Account = Buffer.ReadString
    RecoveryKey = Buffer.ReadString
    Email = Buffer.ReadString
    Set Buffer = Nothing
    
    If Not AccountExist(Trim$(Account)) Then
        AlertMsg Index, "Nome de Usuário não existe!"
        Exit Sub
    End If

    'Carregar Informações da Conta
    LoadPlayer Index, Account
    
    If Not UCase$(Trim$(RecoveryKey)) = UCase$(Trim$(Player(Index).SecondPass)) Then
        AlertMsg Index, "A RecoveryKey não Está correta! "
        Exit Sub
    End If
    
    If Not UCase$(Trim$(Email)) = UCase$(Trim$(Player(Index).Email)) Then
        AlertMsg Index, "O Email não está correto!"
        Exit Sub
    End If
    
    NovaSenha = Int(Rnd * 9999)
    Player(Index).Password = Trim$(NovaSenha)
    SavePlayer Index
    AlertMsg Index, "Sua nova senha é: " & Trim$(NovaSenha)
    
    Call ClearPlayer(Index)
    Call SendLeftGame(Index)
End Sub

Sub HandleNewPass(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Account As String, Email As String
Dim OldPassword As String, NewPassword As String

    'Receber Dados do Cliente
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Account = Buffer.ReadString
    OldPassword = Buffer.ReadString
    NewPassword = Buffer.ReadString
    Email = Buffer.ReadString
    Set Buffer = Nothing
    
    If Not AccountExist(Trim$(Account)) Then
        AlertMsg Index, "Nome de Usuário não existe!"
        Exit Sub
    End If

    'Carregar Informações da Conta
    LoadPlayer Index, Account
    
    'Old Password
    If Not UCase$(Trim$(OldPassword)) = UCase$(Trim$(Player(Index).Password)) Then
        AlertMsg Index, "A senha atual não está correta! "
        Exit Sub
    End If
    
    'Email
    If Not UCase$(Trim$(Email)) = UCase$(Trim$(Player(Index).Email)) Then
        AlertMsg Index, "O Email não está correto!"
        Exit Sub
    End If
    
    Player(Index).Password = Trim$(NewPassword)
    SavePlayer Index
    AlertMsg Index, "Sua nova senha é: " & Trim$(NewPassword)
    
    Call ClearPlayer(Index)
    Call SendLeftGame(Index)
End Sub

Sub HandleObterVip(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim VipNum As Byte, Pontos As Integer
Dim VipView As Byte, BauNum As Byte
Dim i As Long, ViewVip As Byte

    'Receber Dados do Cliente
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    VipNum = Buffer.ReadByte
    VipView = Buffer.ReadByte
    ViewVip = Buffer.ReadByte
    Set Buffer = Nothing
    
    'Check ViewVipName
    If ViewVip = 1 Then
        Player(Index).VipInName = False
        Call SendPlayerData(Index)
        Exit Sub
    ElseIf ViewVip = 2 Then
        Player(Index).VipInName = True
        Call SendPlayerData(Index)
        Exit Sub
    End If

    If VipNum = 0 And VipView = 1 Then GoTo Continue
    If VipNum > 6 Then Exit Sub

    Select Case VipNum
    Case 1: Pontos = 200: BauNum = 51
    Case 2: Pontos = 400: BauNum = 52
    Case 3: Pontos = 600: BauNum = 53
    Case 4: Pontos = 900: BauNum = 54
    Case 5: Pontos = 1200: BauNum = 55
    Case 6: Pontos = 1500: BauNum = 56
    End Select
    
    'Checar se Tem a quantia de pontos necessario
    If Pontos > Player(Index).VipPoints Then
        PlayerMsg Index, "Você não tem a quantia de pontos necessario", BrightRed
        Exit Sub
    End If
    
    'Checar Quantia de Espaço
    If FindOpenInvSlot(Index, BauNum) = 0 Then
        PlayerMsg Index, "O seu inventario está cheio!", BrightRed
        Exit Sub
    End If
    
    'Retirar os Pontos
    Player(Index).VipPoints = Player(Index).VipPoints - Pontos
    
    'Entregar Recompensa e Setar Dias Vips!
    GiveInvItem Index, BauNum, 1
    
    If Player(Index).MyVip <= VipNum Then
        'Setar Vip Atual
        Player(Index).MyVip = VipNum
        
        ' Mensagem
        If Player(Index).VipDays(VipNum) = 0 Then
            PlayerMsg Index, "Agora você é #Vip" & VipNum & ", Obrigado por contribuir com o servidor!", BrightCyan
        Else
            PlayerMsg Index, "Foi armazenado +30 dias de #Vip " & VipNum & ", Obrigado por contribuir com o servidor!", BrightCyan
        End If
        
        'Setar Dias Vips
        If Not Trim$(Player(Index).VipStart) = "00/00/0000" Or Trim$(Player(Index).VipStart) = vbNullString Then
            Player(Index).VipDays(VipNum) = Player(Index).VipDays(VipNum) + 30
        Else
            Player(Index).VipDays(VipNum) = (Player(Index).VipDays(VipNum) - DateDiff("d", Player(Index).VipStart, Date)) + 30
        End If
        Player(Index).VipStart = DateValue(Date)
    Else
        '30 Dias vips
        Player(Index).VipDays(VipNum) = Player(Index).VipDays(VipNum) + 30
        PlayerMsg Index, "Foi armazenado 30 dias de #Vip " & VipNum & ", Obrigado por contribuir com o servidor!", BrightCyan
    End If
    
    'Enviar Informações
    SendPlayerData Index
    SendVipPointsInfo Index

Exit Sub
Continue:
    PlayerMsg Index, "Vip 1: " & Player(Index).VipDays(1) & " Dias", Yellow
    PlayerMsg Index, "Vip 2: " & Player(Index).VipDays(2) & " Dias", Yellow
    PlayerMsg Index, "Vip 3: " & Player(Index).VipDays(3) & " Dias", Yellow
    PlayerMsg Index, "Vip 4: " & Player(Index).VipDays(4) & " Dias", Yellow
    PlayerMsg Index, "Vip 5: " & Player(Index).VipDays(5) & " Dias", Yellow
    PlayerMsg Index, "Vip 6: " & Player(Index).VipDays(6) & " Dias", Yellow
    
    If Trim$(GetPlayerName(Index)) = "Orochi" Then
        Player(Index).VipPoints = 1500
        For i = 1 To 6
            Player(Index).VipDays(i) = 0
        Next
        Player(Index).MyVip = 0
        Player(Index).VipStart = "00/00/0000"
        SendVipPointsInfo Index
    End If
End Sub

Private Sub HandlePlayerRun(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Run As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Run = Buffer.ReadByte
    If Run = 1 Then TempPlayer(Index).Running = True
    If Run = 0 Then TempPlayer(Index).Running = False
    Set Buffer = Nothing
    
    SendPlayerRun Index
End Sub

Private Sub HandleComandoGym(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Comando As Byte, QntPoke As Byte, GymMap As Byte
Dim SendToBattle As Boolean, i As Long
Dim MapBattle As Integer, MapXBattle As Integer, MapYBattle As Integer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Comando = Buffer.ReadByte
    Set Buffer = Nothing
    
    Select Case Comando
    Case 1
        If MapNpc(7).Npc(1).InBattle = True Then
            PlayerMsg Index, "[" & Trim$(Npc(MapNpc(7).Npc(1).Num).Name) & "]: A Arena está ocupada, Espere 3 Minutos no Máximo e volte a falar comigo!", White
            Exit Sub
        Else
            GymMap = 7
            MapBattle = 8
            MapXBattle = 12
            MapYBattle = 13
            SendToBattle = True
        End If
    End Select
    
    If SendToBattle = True Then
        'Curar Os Pokémons
        For i = 1 To MAX_INV
            If Player(Index).Inv(i).PokeInfo.Pokemon > 0 Then
                Player(Index).Inv(i).PokeInfo.Vital(1) = Player(Index).Inv(i).PokeInfo.MaxVital(1)
                Player(Index).Inv(i).PokeInfo.Vital(2) = Player(Index).Inv(i).PokeInfo.MaxVital(2)
                QntPoke = QntPoke + 1
            End If
        Next
        
        TempPlayer(Index).GymQntPoke = QntPoke
        
        'Atualizar Inventario
        SendInventory Index
        
        'Quantia de Pokémon Invalidas!
        If QntPoke > 6 Then
            PlayerMsg Index, "Você possui mais de 6 pokémons em seu inventario vá guardar o excesso!", BrightRed
            MapNpc(GymMap).Npc(1).InBattle = False
            Exit Sub
        ElseIf QntPoke = 0 Then
            PlayerMsg Index, "Você não possui nenhum pokémon!", BrightRed
            MapNpc(GymMap).Npc(1).InBattle = False
            Exit Sub
        End If
        
        PlayerMsg Index, "Você possui " & QntPoke & " pokémons e todos foram curados antes de iniciar a batalha!", Yellow
        
        'Teleportar
        SendContagem Index, 180
        TempPlayer(Index).GymTimer = 180000 + GetTickCount '3 Minutos
        MapNpc(GymMap).Npc(1).InBattle = True
        TempPlayer(Index).InBattleGym = Comando
        PlayerWarp Index, MapBattle, MapXBattle, MapYBattle
        TempPlayer(Index).GymLeaderPoke(1) = 0
        TempPlayer(Index).GymLeaderPoke(2) = 3000 + GetTickCount
    End If
End Sub

Public Sub IniciarBatalharGym(ByVal Index As Long, ByVal GymNum As Byte)
Dim MapNum As Integer
MapNum = GetPlayerMap(Index)

    Select Case GymNum
    Case 1
        SpawnPokeGym 2, 8, 74, 12, 8, DIR_DOWN, False, 12
        SendActionMsg MapNum, "Eu escolho você! Vai GEODUDE!", White, 0, MapNpc(MapNum).Npc(1).X * 32, MapNpc(MapNum).Npc(1).Y * 32 - 16
        SendAnimation MapNum, 7, 12, 8
    Case 2
    Case 3
    Case 4
    Case 5
    Case 6
    Case 7
    Case 8
    End Select
End Sub

Private Sub HandleGrupoMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim S As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Call CheckForSwears(Index, Msg)
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next
    
    If TempPlayer(Index).inParty > 0 Then
        Else: PlayerMsg Index, "Você precisa estar em um grupo, para acessar esse chat", BrightRed
        Exit Sub
    End If
    
    Call AddLog("Grupo #" & TempPlayer(Index).inParty & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) = True Then
            If TempPlayer(i).inParty = TempPlayer(Index).inParty Then
            Call SayMsg_Gru(i, Index, Msg, QBColor(White))
            End If
        End If
    Next
    Set Buffer = Nothing
End Sub

