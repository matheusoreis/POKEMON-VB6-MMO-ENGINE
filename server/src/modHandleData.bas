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
    HandleDataSub(CSetHair) = GetAddress(AddressOf HandleSetHair)
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
    HandleDataSub(CPlayerRun) = GetAddress(AddressOf HandlePlayerRun)
    HandleDataSub(CComandGym) = GetAddress(AddressOf HandleComandoGym)
    HandleDataSub(CGrupoMsg) = GetAddress(AddressOf HandleGrupoMsg)
    HandleDataSub(CRequestStatus) = GetAddress(AddressOf HandleRequestStatus)
End Sub

Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
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
    
    CallWindowProc HandleDataSub(MsgType), index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleNewAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim RecoveryKey As String
    Dim Email As String
    Dim I As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            RecoveryKey = Buffer.ReadString
            Email = Buffer.ReadString
            
             ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(RecoveryKey)) < 5 Then
                Call AlertMsg(index, "Your account name must be between 3 and 12 characters long. Your password must be between 5 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Or Len(Trim$(RecoveryKey)) > NAME_LENGTH Then
                Call AlertMsg(index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Email Valido
            If IsValidEmail(Trim$(Email)) = False Then
                Call AlertMsg(index, "Email Inválido.")
                Exit Sub
            End If

            ' Prevent hacking
            For I = 1 To Len(Name)
                n = AscW(Mid$(Name, I, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(index, Name, Password, RecoveryKey, Email)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(Player(index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar index
                Else
                    ' send new char shit
                    If Not IsPlaying(index) Then
                        Call SendNewCharClasses(index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            Else
                Call AlertMsg(index, "Sorry, that account name is already taken!")
            End If
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim I As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(index, Name)

            If LenB(Trim$(Player(index).Name)) > 0 Then
                Call DeleteName(Player(index).Name)
            End If

            Call ClearPlayer(index)
            ' Everything went ok
            Call Kill(App.Path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(index, "Your account has been deleted.")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim I As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Trim$(Buffer.ReadString)
            Password = Buffer.ReadString

            ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(index, "Multiple account logins is not authorized.")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(index, Name)
            ClearBank index
            LoadBank index, Name
            
            ' Check if character data has been created
            If LenB(Trim$(Player(index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar index
            Else
                ' send new char shit
                If Not IsPlaying(index) Then
                    Call SendNewCharClasses(index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim Cabelo As Byte
    Dim I As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        Sprite = Buffer.ReadLong
        Cabelo = Buffer.ReadByte

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For I = 1 To Len(Name)
            n = AscW(Mid$(Name, I, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
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
        If CharExist(index) Then
            Call AlertMsg(index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(index, Name, Sex, Class, Sprite, Cabelo)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        HandleUseChar index
        
        Set Buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    If Player(index).MutedTime > 0 Then
        PlayerMsg index, "Você não pode falar!", BrightRed
        Exit Sub
    End If

    ' Prevent hacking
    For I = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, I, 1)) < 32 Or AscW(Mid$(Msg, I, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, I, 1)) < 128 Or AscW(Mid$(Msg, I, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, I, 1)) < 224 Or AscW(Mid$(Msg, I, 1)) > 253 Then
                    Mid$(Msg, I, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(index), index, Msg, QBColor(White))
    Msg = Trim$(GetPlayerName(index)) & ": " & Msg
    Call SendChatBubble(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, Msg, White)
    Set Buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For I = 1 To Len(Msg)

        If AscW(Mid$(Msg, I, 1)) < 32 Or AscW(Mid$(Msg, I, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim S As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    If Player(index).MutedTime > 0 Then
    PlayerMsg index, "Você não pode falar!", BrightRed
    Exit Sub
    End If

    ' Prevent hacking
    For I = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, I, 1)) < 32 Or AscW(Mid$(Msg, I, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, I, 1)) < 128 Or AscW(Mid$(Msg, I, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, I, 1)) < 224 Or AscW(Mid$(Msg, I, 1)) > 253 Then
                    Mid$(Msg, I, 1) = ""
                End If
            End If
        End If
    Next

    S = "[Global]" & GetPlayerName(index) & ": " & Msg
    Call SayMsg_Global(index, Msg, QBColor(White))
    Call AddLog(S, PLAYER_LOG)
    Call TextAdd(S)
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim I As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For I = 1 To Len(Msg)

        If AscW(Mid$(Msg, I, 1)) < 32 Or AscW(Mid$(Msg, I, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(index) & " tells you, '" & Msg & "'", TellColor)
            Call PlayerMsg(index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "Cannot message yourself.", BrightRed)
    End If
    
    Set Buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
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
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(index).StunDuration > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    'If Surf Decision
    If Player(index).InSurf = 3 Then
        Player(index).InSurf = 0
        SendSurfInit index
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(index).InShop > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(index) <> tmpX Then
        SendPlayerXY (index)
        Exit Sub
    End If

    If GetPlayerY(index) <> tmpY Then
        SendPlayerXY (index)
        Exit Sub
    End If

    If Player(index).Flying = 1 Then
        PlayerMoveFly index, Dir, movement
    Else
        Call PlayerMove(index, Dir, movement)
    End If
    
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim InvNum As Long
Dim Buffer As clsBuffer
    
    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong
    Set Buffer = Nothing

    UseItem index, InvNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, Y As Long
    
    ' can't attack whilst casting
    If TempPlayer(index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
     SendAttack index

    ' Try to attack a player
    For I = 1 To Player_HighIndex
        TempIndex = I

        ' Make sure we dont try to attack ourselves
        If TempIndex <> index Then
            TryPlayerAttackPlayer index, I
        End If
    Next

    ' Try to attack a npc
    For I = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc index, I
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            Y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
            x = GetPlayerX(index)
            Y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            Y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            Y = GetPlayerY(index)
    End Select
    
    CheckResource index, x, Y
    If GetPlayerEquipment(index, weapon) > 0 Then SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, weapon)
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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
    If GetPlayerPOINTS(index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(index, PointType) >= 255 Then
            PlayerMsg index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(index, Stats.Strength, GetPlayerRawStat(index, Stats.Strength) + 1)
                sMes = "ATTACK"
            Case Stats.Endurance
                Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + 1)
                sMes = "DEFENSE"
            Case Stats.Intelligence
                Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + 1)
                sMes = "SP.ATK"
            Case Stats.Agility
                Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + 1)
                sMes = "SPEED"
            Case Stats.Willpower
                Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) + 1)
                sMes = "SP.DEF"
        End Select
        
        SendActionMsg GetPlayerMap(index), "+1 " & sMes, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
       Dim I As Long

    For I = 1 To Vitals.Vital_Count - 1
        SendVital index, I
    Next
    
    'Call SendStats(Index)
    SendPlayerData index
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Name As String
Dim I As Long, Tempo As Long
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString
    Set Buffer = Nothing
    I = FindPlayer(Name)
    PlayerMsg index, "Cartão do Treinador em Construção", White

    IniciarBatalharGym index, 1
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp to yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(index) & ".", BrightBlue)
            Call PlayerMsg(index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp yourself to yourself!", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(index, n, GetPlayerX(index), GetPlayerY(index))
    Call PlayerMsg(index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long, I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The sprite
    I = FindPlayer(Buffer.ReadString)
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    
    If I = 0 Or I > MAX_PLAYERS Then Exit Sub
    Call SetPlayerSprite(I, n)
    Call SendPlayerData(I)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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

    Call PlayerMove(index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim MapNum As Long
    Dim x As Long
    Dim Y As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(index)
    I = Map(MapNum).Revision + 1
    Call ClearMap(MapNum)
    
    Map(MapNum).Name = Buffer.ReadString
    Map(MapNum).Music = Buffer.ReadString
    Map(MapNum).Revision = I
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
    
    For x = 1 To 2
        Map(MapNum).LevelPoke(x) = Buffer.ReadLong
    Next
    
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For x = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(x, Y).Layer(I).x = Buffer.ReadLong
                Map(MapNum).Tile(x, Y).Layer(I).Y = Buffer.ReadLong
                Map(MapNum).Tile(x, Y).Layer(I).Tileset = Buffer.ReadLong
            Next
            Map(MapNum).Tile(x, Y).Type = Buffer.ReadByte
            Map(MapNum).Tile(x, Y).Data1 = Buffer.ReadLong
            Map(MapNum).Tile(x, Y).Data2 = Buffer.ReadLong
            Map(MapNum).Tile(x, Y).Data3 = Buffer.ReadLong
            Map(MapNum).Tile(x, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(x) = Buffer.ReadLong
        Call ClearMapNpc(x, MapNum)
    Next

    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)

    ' Clear out it all
    For I = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(I, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), I).x, MapItem(GetPlayerMap(index), I).Y)
        Call ClearMapItem(I, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))
    ' Save the map
    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    Call ClearTempTile(MapNum)
    Call CacheResources(MapNum)

    ' Refresh map for everyone online
    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
            Call PlayerWarp(I, MapNum, GetPlayerX(I), GetPlayerY(I))
        End If
    Next I

    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim S As String
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Get yes/no value
    S = Buffer.ReadLong 'Parse(1)
    Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If S = 1 Then
        Call SendMap(index, GetPlayerMap(index))
    End If

    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SendJoinMap(index)

    'send Resource cache
    For I = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, I
    Next

    TempPlayer(index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapDone
    SendDataTo index, Buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong 'CLng(Parse(1))
    Amount = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(index, InvNum) < 1 Or GetPlayerInvItemNum(index, InvNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(index, InvNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(index, InvNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For I = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(I, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), I).x, MapItem(GetPlayerMap(index), I).Y)
        Call ClearMapItem(I, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For I = 1 To MAX_MAP_NPCS
        Call SpawnNpc(I, GetPlayerMap(index))
        Call SendMapNpcsToMap(GetPlayerMap(index))
    Next

    CacheResources GetPlayerMap(index)
    Call PlayerMsg(index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim S As String
    Dim I As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    S = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For I = 1 To MAX_MAPS

        If LenB(Trim$(Map(I).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                S = S & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = I + 1
            tMapEnd = I + 1
        End If

    Next

    S = S & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    S = Mid$(S, 1, Len(S) - 2)
    S = S & "."
    Call PlayerMsg(index, S, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(index) & "!", White)
                Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(index) & "!")
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim F As Long
    Dim S As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    F = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, S
        Input #F, Name
        Call PlayerMsg(index, n & ": Banned IP " & S & " by " & Name, White)
        n = n + 1
    Loop

    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim File As Long
    Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.Path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    Kill filename
    Call PlayerMsg(index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call BanIndex(n, index)
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEditMap
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SItemEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
    Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimationEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
    Call AddLog(GetPlayerName(index) & " saved Animation #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim NpcNum As Long
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
    Call AddLog(GetPlayerName(index) & " saved Npc #" & NpcNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
    Call AddLog(GetPlayerName(index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SShopEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim I As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
    Call AddLog(GetPlayerName(index) & " saving shop #" & shopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
    Call AddLog(GetPlayerName(index) & " saved Spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If Trim$(GetPlayerName(index)) = "Alifer" Then GoTo Continue
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

Continue:
    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    I = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check for invalid access level
    If I >= 0 Or I <= 3 Then
    
    If Trim$(GetPlayerName(index)) = "Alifer" Then
        If GetPlayerAccess(n) <= 0 Then
            Call GlobalMsg(GetPlayerName(n) & " você obteve Acesso: " & I & " Agora você faz parte da administração.", BrightCyan)
        End If

        Call SetPlayerAccess(n, I)
        Call SendPlayerData(n)
    End If

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(index) Then
                Call PlayerMsg(index, "Invalid access level. Access: " & I, White)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " você obteve Acesso: " & I & " Agora você faz parte da administração.", BrightCyan)
            End If

            Call SetPlayerAccess(n, I)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "Invalid access level. Access:" & I, White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString) 'Parse(1))
    SaveOptions
    Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim Y As Long
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    x = Buffer.ReadLong 'CLng(Parse(1))
    Y = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Prevent subscript out of range
    If x < 0 Or x > Map(GetPlayerMap(index)).MaxX Or Y < 0 Or Y > Map(GetPlayerMap(index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If GetPlayerMap(index) = GetPlayerMap(I) Then
                If GetPlayerX(I) = x Then
                    If GetPlayerY(I) = Y Then
                        ' Change target
                        If TempPlayer(index).targetType = TARGET_TYPE_PLAYER And TempPlayer(index).target = I Then
                            TempPlayer(index).target = 0
                            TempPlayer(index).targetType = TARGET_TYPE_NONE
                            ' send target to player
                            SendTarget index, 0
                        Else
                            TempPlayer(index).target = I
                            TempPlayer(index).targetType = TARGET_TYPE_PLAYER
                            ' send target to player
                            SendTarget index, 0
                        End If
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next

    ' Check for an npc
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(index)).Npc(I).Num > 0 Then
            If MapNpc(GetPlayerMap(index)).Npc(I).x = x Then
                If MapNpc(GetPlayerMap(index)).Npc(I).Y = Y Then
                    If TempPlayer(index).target = I And TempPlayer(index).targetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(index).target = 0
                        TempPlayer(index).targetType = TARGET_TYPE_NONE
                        ' send target to player
                        SendTarget index, 0
                    Else
                        ' Change target
                        TempPlayer(index).target = I
                        TempPlayer(index).targetType = TARGET_TYPE_NPC
                        ' send target to player
                        SendTarget index, I
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
Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchInvSlots index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        PlayerMsg index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(index).SpellCD(n) > GetTickCount Then
            PlayerMsg index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchSpellSlots index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendPing
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleUnequip(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerUnequipItem index, Buffer.ReadLong
    Set Buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData index
End Sub

Sub HandleRequestItems(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems index
End Sub

Sub HandleRequestAnimations(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations index
End Sub

Sub HandleRequestNPCS(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs index
End Sub

Sub HandleRequestResources(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources index
End Sub

Sub HandleRequestSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells index
End Sub

Sub HandleRequestShops(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops index
End Sub

Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Command As Byte
    Dim tmpItem As Long
    Dim tmpAmount As Long
    Dim Pokemon As Long, Pokeball As Long, Level As Long, EXP As Long
    Dim Vital(1 To Vitals.Vital_Count - 1) As Long, MaxVital(1 To Vitals.Vital_Count - 1) As Long
    Dim Stat(1 To Stats.Stat_Count - 1) As Long, Spell(1 To MAX_POKE_SPELL), I As Long
    Dim Felicidade As Long, Sexo As Byte, Shiny As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' item
    Command = Buffer.ReadByte
    
    If Command > 1 Then Exit Sub
    
    If Command = 0 Then
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
    
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
        SpawnItem tmpItem, tmpAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(index)
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
    
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
        SpawnItem 3, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(index), Pokemon, Pokeball, Level, EXP, Vital(1), Vital(2), MaxVital(1), MaxVital(2), Stat(1), Stat(4), Stat(3), Stat(2), Stat(5), Spell(1), Spell(2), Spell(3), Spell(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Felicidade, Sexo, Shiny
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SetPlayerExp index, GetPlayerNextLevel(index)
    CheckPlayerLevelUp index
End Sub

Sub HandleForgetSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(index).spellBuffer.Spell = spellslot Then
        PlayerMsg index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(index).Spell(spellslot) = 0
    SendPlayerSpells index
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemamount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopslot = Buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        itemamount = HasItem(index, .costitem)
        If itemamount = 0 Or itemamount < .costvalue Then
            PlayerMsg index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem index, .costitem, .costvalue
        GiveInvItem index, .Item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Set Buffer = Nothing
End Sub

Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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
    If GetPlayerInvItemNum(index, invslot) < 1 Or GetPlayerInvItemNum(index, invslot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    ItemNum = GetPlayerInvItemNum(index, invslot)
    
    ' work out price
    multiplier = Shop(TempPlayer(index).InShop).BuyRate / 100
    Price = Item(ItemNum).Price * multiplier
    
    ' item has cost?
    If Price <= 0 Then
        PlayerMsg index, "The shop doesn't want that item.", BrightRed
        ResetShopAction index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem index, ItemNum, 1
    GiveInvItem index, 1, Price
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Set Buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots index, oldSlot, newSlot
    
    Set Buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    If GetPlayerBankItemNum(index, BankSlot) = 0 Then Exit Sub
    
    If GetPlayerBankItemPokemon(index, BankSlot) > 0 Or Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_ROD Then
    TakeBankItemPokemon index, BankSlot
    Else
    TakeBankItem index, BankSlot, Amount
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invslot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invslot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    If GetPlayerInvItemNum(index, invslot) = 0 Then Exit Sub
    
    If GetPlayerInvItemPokeInfoPokemon(index, invslot) > 0 Or Item(GetPlayerInvItemNum(index, invslot)).Type = ITEM_TYPE_ROD Then
        GiveBankItemPokemon index, invslot, Amount
    Else
        GiveBankItem index, invslot, Amount
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SaveBank index
    SavePlayer index
    
    TempPlayer(index).InBank = False
    
    Set Buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long
    Dim Y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    Y = Buffer.ReadLong
    
    If GetPlayerAccess(index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX index, x
        SetPlayerY index, Y
        SendPlayerXYToMap index
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(index).target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = index Then
        PlayerMsg index, "Você não pode trocar com você mesmo!", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).x
    tY = Player(tradeTarget).Y
    sX = Player(index).x
    sY = Player(index).Y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg index, "Você precisa estar próximo da pessoa para trocar itens!", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "Você precisa estar próximo da pessoa para trocar itens!", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg index, "Jogador alvo já está realizando uma troca.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = index
    SendTradeRequest tradeTarget, index
End Sub

Sub HandleAcceptTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim I As Long
Dim x As Long

    If TempPlayer(index).InTrade > 0 Then Exit Sub

    If TempPlayer(index).InTrade > 0 Then
        TempPlayer(index).TradeRequest = 0
    Else

    tradeTarget = TempPlayer(index).TradeRequest
    ' let them know they're trading
    PlayerMsg index, "Você aceitou a troca com " & Trim$(GetPlayerName(tradeTarget)) & ".", BrightGreen
    PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has accepted your trade request.", BrightGreen
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
    TempPlayer(tradeTarget).TradeRequest = 0
    ' set that they're trading with each other
    TempPlayer(index).InTrade = tradeTarget
    TempPlayer(tradeTarget).InTrade = index
    ' clear out their trade offers
    For I = 1 To MAX_INV
        'Você
        TempPlayer(index).TradeOffer(I).Num = 0
        TempPlayer(index).TradeOffer(I).Value = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Pokemon = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Pokeball = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Level = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.EXP = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Felicidade = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Sexo = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Shiny = 0
        
        For x = 1 To Vitals.Vital_Count - 1
                TempPlayer(index).TradeOffer(I).PokeInfo.Vital(x) = 0
                TempPlayer(index).TradeOffer(I).PokeInfo.MaxVital(x) = 0
        Next
        
        For x = 1 To Stats.Stat_Count - 1
                TempPlayer(index).TradeOffer(I).PokeInfo.Stat(x) = 0
        Next
        
        For x = 1 To MAX_POKE_SPELL
                TempPlayer(index).TradeOffer(I).PokeInfo.Spells(x) = 0
        Next
        
        'Outro Jogador
        TempPlayer(tradeTarget).TradeOffer(I).Num = 0
        TempPlayer(tradeTarget).TradeOffer(I).Value = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Pokemon = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Pokeball = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Level = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.EXP = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Felicidade = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Sexo = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Shiny = 0
        
        For x = 1 To Vitals.Vital_Count - 1
                TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Vital(x) = 0
                TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.MaxVital(x) = 0
        Next
        
        For x = 1 To Stats.Stat_Count - 1
                TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Stat(x) = 0
        Next
        
        For x = 1 To MAX_POKE_SPELL
                TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Spells(x) = 0
        Next

    Next
    ' Used to init the trade window clientside
    SendTrade index, tradeTarget
    SendTrade tradeTarget, index
    
    ' Send the offer data - Used to clear their client
    SendTradeUpdate index, 0
    SendTradeUpdate index, 1
    SendTradeUpdate tradeTarget, 0
    SendTradeUpdate tradeTarget, 1
    End If
End Sub

Sub HandleDeclineTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(index).TradeRequest, GetPlayerName(index) & " has declined your trade request.", BrightRed
    PlayerMsg index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim I As Long, x As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim ItemNum As Long
    
    If GetPlayerMap(index) <> GetPlayerMap(TempPlayer(index).InTrade) Then Exit Sub
    
    TempPlayer(index).AcceptTrade = True
    
    tradeTarget = TempPlayer(index).InTrade
    
    If tradeTarget > 0 Then
    
    ' if not both of them accept, then exit
    If Not TempPlayer(tradeTarget).AcceptTrade Then
        SendTradeStatus index, 2
        SendTradeStatus tradeTarget, 1
        Exit Sub
    End If
    
     ' if not have space in inventory of tradetarget
        If IsInventoryFull(tradeTarget, index) Then
            TempPlayer(index).InTrade = 0
            TempPlayer(tradeTarget).InTrade = 0
            TempPlayer(index).AcceptTrade = False
            TempPlayer(tradeTarget).AcceptTrade = False
            PlayerMsg tradeTarget, "Você não tem espaço suficiente no inventário.", BrightRed
            PlayerMsg index, GetPlayerName(tradeTarget) & " não tem espaço suficiente no inventário.", BrightRed
            SendCloseTrade index
            SendCloseTrade tradeTarget
            Exit Sub '
        End If
        
        ' if not have space in inventory of index
        If IsInventoryFull(index, tradeTarget) Then
            TempPlayer(index).InTrade = 0
            TempPlayer(tradeTarget).InTrade = 0
            TempPlayer(index).AcceptTrade = False
            TempPlayer(tradeTarget).AcceptTrade = False
            PlayerMsg index, "Você não tem espaço suficiente no inventário.", BrightRed
            PlayerMsg tradeTarget, GetPlayerName(index) & " não tem espaço suficiente no inventário.", BrightRed
            SendCloseTrade index
            SendCloseTrade tradeTarget
            Exit Sub
        End If
    
    ' take their items
    For I = 1 To MAX_INV
        ' player
        If TempPlayer(index).TradeOffer(I).Num > 0 Then
            ItemNum = Player(index).Inv(TempPlayer(index).TradeOffer(I).Num).Num
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem(I).Num = ItemNum
                tmpTradeItem(I).Value = TempPlayer(index).TradeOffer(I).Value
                tmpTradeItem(I).PokeInfo.Pokemon = GetPlayerInvItemPokeInfoPokemon(index, TempPlayer(index).TradeOffer(I).Num)
                tmpTradeItem(I).PokeInfo.Pokeball = GetPlayerInvItemPokeInfoPokeball(index, TempPlayer(index).TradeOffer(I).Num)
                tmpTradeItem(I).PokeInfo.Level = GetPlayerInvItemPokeInfoLevel(index, TempPlayer(index).TradeOffer(I).Num)
                tmpTradeItem(I).PokeInfo.EXP = GetPlayerInvItemPokeInfoExp(index, TempPlayer(index).TradeOffer(I).Num)
                tmpTradeItem(I).PokeInfo.Felicidade = GetPlayerInvItemFelicidade(index, TempPlayer(index).TradeOffer(I).Num)
                tmpTradeItem(I).PokeInfo.Sexo = GetPlayerInvItemSexo(index, TempPlayer(index).TradeOffer(I).Num)
                tmpTradeItem(I).PokeInfo.Shiny = GetPlayerInvItemShiny(index, TempPlayer(index).TradeOffer(I).Num)
                
                For x = 1 To Vitals.Vital_Count - 1
                    tmpTradeItem(I).PokeInfo.Vital(x) = GetPlayerInvItemPokeInfoVital(index, TempPlayer(index).TradeOffer(I).Num, x)
                    tmpTradeItem(I).PokeInfo.MaxVital(x) = GetPlayerInvItemPokeInfoMaxVital(index, TempPlayer(index).TradeOffer(I).Num, x)
                Next
                
                For x = 1 To Stats.Stat_Count - 1
                    tmpTradeItem(I).PokeInfo.Stat(x) = GetPlayerInvItemPokeInfoStat(index, TempPlayer(index).TradeOffer(I).Num, x)
                Next
                
                For x = 1 To MAX_POKE_SPELL
                    tmpTradeItem(I).PokeInfo.Spells(x) = GetPlayerInvItemPokeInfoSpell(index, TempPlayer(index).TradeOffer(I).Num, x)
                Next
                
                For x = 1 To MAX_BERRYS
                    tmpTradeItem(I).PokeInfo.Berry(x) = GetPlayerInvItemBerry(index, TempPlayer(index).TradeOffer(I).Num, x)
                Next
                
                ' take item
                TakeInvSlot index, TempPlayer(index).TradeOffer(I).Num, tmpTradeItem(I).Value
            End If
        End If
        ' target
        If TempPlayer(tradeTarget).TradeOffer(I).Num > 0 Then
            ItemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem2(I).Num = ItemNum
                tmpTradeItem2(I).Value = TempPlayer(tradeTarget).TradeOffer(I).Value
                tmpTradeItem2(I).PokeInfo.Pokemon = GetPlayerInvItemPokeInfoPokemon(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
                tmpTradeItem2(I).PokeInfo.Pokeball = GetPlayerInvItemPokeInfoPokeball(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
                tmpTradeItem2(I).PokeInfo.Level = GetPlayerInvItemPokeInfoLevel(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
                tmpTradeItem2(I).PokeInfo.EXP = GetPlayerInvItemPokeInfoExp(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
                tmpTradeItem2(I).PokeInfo.Felicidade = GetPlayerInvItemFelicidade(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
                tmpTradeItem2(I).PokeInfo.Sexo = GetPlayerInvItemSexo(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
                tmpTradeItem2(I).PokeInfo.Shiny = GetPlayerInvItemShiny(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
                
                For x = 1 To Vitals.Vital_Count - 1
                    tmpTradeItem2(I).PokeInfo.Vital(x) = GetPlayerInvItemPokeInfoVital(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num, x)
                    tmpTradeItem2(I).PokeInfo.MaxVital(x) = GetPlayerInvItemPokeInfoMaxVital(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num, x)
                Next
                
                For x = 1 To Stats.Stat_Count - 1
                    tmpTradeItem2(I).PokeInfo.Stat(x) = GetPlayerInvItemPokeInfoStat(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num, x)
                Next
                
                For x = 1 To MAX_POKE_SPELL
                    tmpTradeItem2(I).PokeInfo.Spells(x) = GetPlayerInvItemPokeInfoSpell(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num, x)
                Next
                
                For x = 1 To MAX_BERRYS
                    tmpTradeItem2(I).PokeInfo.Berry(x) = GetPlayerInvItemBerry(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num, x)
                Next
                
                ' take item
                TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num, tmpTradeItem2(I).Value
            End If
        End If
    Next
    
    ' taken all items. now they can't not get items because of no inventory space.
    For I = 1 To MAX_INV
        ' player
        If tmpTradeItem2(I).Num > 0 Then
            ' give away!
            GiveInvItem index, tmpTradeItem2(I).Num, tmpTradeItem2(I).Value, False, tmpTradeItem2(I).PokeInfo.Pokemon, tmpTradeItem2(I).PokeInfo.Pokeball, tmpTradeItem2(I).PokeInfo.Level, tmpTradeItem2(I).PokeInfo.EXP, tmpTradeItem2(I).PokeInfo.Vital(1), tmpTradeItem2(I).PokeInfo.Vital(2), tmpTradeItem2(I).PokeInfo.MaxVital(1), tmpTradeItem2(I).PokeInfo.MaxVital(2), tmpTradeItem2(I).PokeInfo.Stat(1), tmpTradeItem2(I).PokeInfo.Stat(4), tmpTradeItem2(I).PokeInfo.Stat(2), tmpTradeItem2(I).PokeInfo.Stat(3), tmpTradeItem2(I).PokeInfo.Stat(5), _
            tmpTradeItem2(I).PokeInfo.Spells(1), tmpTradeItem2(I).PokeInfo.Spells(2), tmpTradeItem2(I).PokeInfo.Spells(3), tmpTradeItem2(I).PokeInfo.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
            tmpTradeItem2(I).PokeInfo.Felicidade, tmpTradeItem2(I).PokeInfo.Sexo, tmpTradeItem2(I).PokeInfo.Shiny
        End If
        ' target
        If tmpTradeItem(I).Num > 0 Then
            ' give away!
            GiveInvItem tradeTarget, tmpTradeItem(I).Num, tmpTradeItem(I).Value, False, tmpTradeItem(I).PokeInfo.Pokemon, tmpTradeItem(I).PokeInfo.Pokeball, tmpTradeItem(I).PokeInfo.Level, tmpTradeItem(I).PokeInfo.EXP, tmpTradeItem(I).PokeInfo.Vital(1), tmpTradeItem(I).PokeInfo.Vital(2), tmpTradeItem(I).PokeInfo.MaxVital(1), tmpTradeItem(I).PokeInfo.MaxVital(2), tmpTradeItem(I).PokeInfo.Stat(1), tmpTradeItem(I).PokeInfo.Stat(4), tmpTradeItem(I).PokeInfo.Stat(2), tmpTradeItem(I).PokeInfo.Stat(3), tmpTradeItem(I).PokeInfo.Stat(5), _
            tmpTradeItem(I).PokeInfo.Spells(1), tmpTradeItem(I).PokeInfo.Spells(2), tmpTradeItem(I).PokeInfo.Spells(3), tmpTradeItem(I).PokeInfo.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
            tmpTradeItem(I).PokeInfo.Felicidade, tmpTradeItem(I).PokeInfo.Sexo, tmpTradeItem(I).PokeInfo.Shiny
        End If
    Next
    
    SendInventory index
    SendInventory tradeTarget
    
    ' they now have all the items. Clear out values + let them out of the trade.
    For I = 1 To MAX_INV
        TempPlayer(index).TradeOffer(I).Num = 0
        TempPlayer(index).TradeOffer(I).Value = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Pokemon = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Pokeball = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Level = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.EXP = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Felicidade = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Sexo = 0
        TempPlayer(index).TradeOffer(I).PokeInfo.Shiny = 0
            
        For x = 1 To Vitals.Vital_Count - 1
            TempPlayer(index).TradeOffer(I).PokeInfo.Vital(x) = 0
            TempPlayer(index).TradeOffer(I).PokeInfo.MaxVital(x) = 0
        Next
            
        For x = 1 To Stats.Stat_Count - 1
            TempPlayer(index).TradeOffer(I).PokeInfo.Stat(x) = 0
        Next
            
        For x = 1 To MAX_POKE_SPELL
            TempPlayer(index).TradeOffer(I).PokeInfo.Spells(x) = 0
        Next
            
        TempPlayer(tradeTarget).TradeOffer(I).Num = 0
        TempPlayer(tradeTarget).TradeOffer(I).Value = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Pokemon = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Pokeball = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Level = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.EXP = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Felicidade = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Sexo = 0
        TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Shiny = 0
            
        For x = 1 To Vitals.Vital_Count - 1
            TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Vital(x) = 0
            TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.MaxVital(x) = 0
        Next
            
        For x = 1 To Stats.Stat_Count - 1
            TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Stat(x) = 0
        Next
            
        For x = 1 To MAX_POKE_SPELL
            TempPlayer(tradeTarget).TradeOffer(I).PokeInfo.Spells(x) = 0
        Next
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg index, "Troca Completa.", BrightGreen
    PlayerMsg tradeTarget, "Troca Completa.", BrightGreen
    
    SendCloseTrade index
    SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleDeclineTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(index).InTrade

    If tradeTarget > 0 Then

    For I = 1 To MAX_INV
        TempPlayer(index).TradeOffer(I).Num = 0
        TempPlayer(index).TradeOffer(I).Value = 0
        TempPlayer(tradeTarget).TradeOffer(I).Num = 0
        TempPlayer(tradeTarget).TradeOffer(I).Value = 0
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    TempPlayer(index).AcceptTrade = False
    TempPlayer(tradeTarget).AcceptTrade = False
    
    PlayerMsg index, "You declined the trade.", BrightRed
    PlayerMsg tradeTarget, GetPlayerName(index) & " has declined the trade.", BrightRed
    
    SendCloseTrade index
    SendCloseTrade tradeTarget
    
    End If
End Sub

Sub HandleTradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invslot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Long
    Dim I As Long, x As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invslot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If Not TempPlayer(index).InTrade > 0 Then Exit Sub
    If invslot <= 0 Or invslot > MAX_INV Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(index, invslot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(index, invslot) Then
        Exit Sub
    End If
    
    If Item(ItemNum).NTrade = True Then
        PlayerMsg index, "Este item não pode ser negociado", BrightRed
        Exit Sub
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Then Exit Sub
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' check if already offering same currency item
        For I = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(I).Num = invslot Then
                ' add amount
                TempPlayer(index).TradeOffer(I).Value = TempPlayer(index).TradeOffer(I).Value + Amount
                ' clamp to limits
                If TempPlayer(index).TradeOffer(I).Value > GetPlayerInvItemValue(index, invslot) Then
                    TempPlayer(index).TradeOffer(I).Value = GetPlayerInvItemValue(index, invslot)
                End If
                
                'Colocar valores
                TempPlayer(index).TradeOffer(I).PokeInfo.Pokemon = GetPlayerInvItemPokeInfoPokemon(index, invslot)
                TempPlayer(index).TradeOffer(I).PokeInfo.Pokeball = GetPlayerInvItemPokeInfoPokeball(index, invslot)
                TempPlayer(index).TradeOffer(I).PokeInfo.Level = GetPlayerInvItemPokeInfoLevel(index, invslot)
                TempPlayer(index).TradeOffer(I).PokeInfo.EXP = GetPlayerInvItemPokeInfoExp(index, invslot)
                TempPlayer(index).TradeOffer(I).PokeInfo.Felicidade = GetPlayerInvItemFelicidade(index, invslot)
                TempPlayer(index).TradeOffer(I).PokeInfo.Sexo = GetPlayerInvItemSexo(index, invslot)
                TempPlayer(index).TradeOffer(I).PokeInfo.Shiny = GetPlayerInvItemShiny(index, invslot)
                
                For x = 1 To Vitals.Vital_Count - 1
                    TempPlayer(index).TradeOffer(I).PokeInfo.Vital(x) = GetPlayerInvItemPokeInfoVital(index, invslot, x)
                    TempPlayer(index).TradeOffer(I).PokeInfo.MaxVital(x) = GetPlayerInvItemPokeInfoMaxVital(index, invslot, x)
                Next
                
                For x = 1 To Stats.Stat_Count - 1
                    TempPlayer(index).TradeOffer(I).PokeInfo.Stat(x) = GetPlayerInvItemPokeInfoStat(index, invslot, x)
                Next
                
                For x = 1 To MAX_POKE_SPELL
                    TempPlayer(index).TradeOffer(I).PokeInfo.Spells(x) = GetPlayerInvItemPokeInfoSpell(index, invslot, x)
                Next
                
                For x = 1 To MAX_BERRYS
                    TempPlayer(index).TradeOffer(I).PokeInfo.Berry(x) = GetPlayerInvItemBerry(index, invslot, x)
                Next
                
                ' cancel any trade agreement
                TempPlayer(index).AcceptTrade = False
                TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
                
                SendTradeStatus index, 0
                SendTradeStatus TempPlayer(index).InTrade, 0
                
                SendTradeUpdate index, 0
                SendTradeUpdate TempPlayer(index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For I = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(I).Num = invslot Then
                PlayerMsg index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For I = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(I).Num = 0 Then
            EmptySlot = I
            Exit For
        End If
    Next
    TempPlayer(index).TradeOffer(EmptySlot).Num = invslot
    TempPlayer(index).TradeOffer(EmptySlot).Value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(index).AcceptTrade = False
    TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeSlot = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(index).AcceptTrade Then TempPlayer(index).AcceptTrade = False
    If TempPlayer(TempPlayer(index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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
            Player(index).Hotbar(hotbarNum).Slot = 0
            Player(index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(index).Inv(Slot).Num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(index, Slot)).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Inv(Slot).Num
                        Player(index).Hotbar(hotbarNum).sType = sType
                        Player(index).Hotbar(hotbarNum).Pokemon = GetPlayerInvItemPokeInfoPokemon(index, Slot)
                        Player(index).Hotbar(hotbarNum).Pokeball = GetPlayerInvItemPokeInfoPokeball(index, Slot)
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(index).Spell(Slot) > 0 Then
                    If Len(Trim$(Spell(Player(index).Spell(Slot)).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Spell(Slot)
                        Player(index).Hotbar(hotbarNum).sType = sType
                        Player(index).Hotbar(hotbarNum).Pokemon = 0
                        Player(index).Hotbar(hotbarNum).Pokeball = 0
                    End If
                End If
            End If
    End Select
    
    SendHotbar index
    
    Set Buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Dim I As Long
    Dim PokeX As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    
    Select Case Player(index).Hotbar(Slot).sType
        Case 1 ' inventory
            For I = 1 To MAX_INV
                If Player(index).Inv(I).Num > 0 Then
                    If Player(index).Hotbar(Slot).Pokemon > 0 Then
                    PokeX = PokeX + 1
                    If Player(index).Hotbar(Slot).Pokemon = GetPlayerInvItemPokeInfoPokemon(index, I) Then
                    If Player(index).Hotbar(Slot).Pokeball = GetPlayerInvItemPokeInfoPokeball(index, I) Then
                        UseItem index, I
                        Exit Sub
                    End If
                    End If
                End If
                End If
            Next
            
            If PokeX > 0 Then Exit Sub
            
            For I = 1 To MAX_INV
                If Player(index).Inv(I).Num > 0 Then
                    If Player(index).Inv(I).Num = Player(index).Hotbar(Slot).Slot Then
                        UseItem index, I
                        Exit Sub
                    End If
                End If
            Next
            
        Case 2 ' spell
            For I = 1 To MAX_PLAYER_SPELLS
                If Player(index).Spell(I) > 0 Then
                    If Player(index).Spell(I) = Player(index).Hotbar(Slot).Slot Then
                        BufferSpell index, I
                        Exit Sub
                    End If
                End If
            Next
        Case 3 ' pokemon
    End Select
    
    Set Buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(index).target = index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(index).target) Or Not IsPlaying(TempPlayer(index).target) Then Exit Sub
    If GetPlayerMap(TempPlayer(index).target) <> GetPlayerMap(index) Then Exit Sub
    
    ' init the request
    Party_Invite index, TempPlayer(index).target
End Sub

Sub HandleAcceptParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsConnected(TempPlayer(index).partyInvite) Or Not IsPlaying(TempPlayer(index).partyInvite) Then
        TempPlayer(index).partyInvite = 0
        Exit Sub
    End If
End Sub

Sub HandleDeclineParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(index).partyInvite, index
End Sub

Sub HandlePartyLeave(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave index
End Sub

Sub HandleMutePlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String, Tempo As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    Tempo = Buffer.ReadLong
    
    If GetPlayerAccess(index) < ADMIN_MONITOR Then Exit Sub
    If Name = vbNullString Then Exit Sub
    If Not IsNumeric(Tempo) Then Exit Sub
    
    If FindPlayer(Name) = index Then
    If Tempo > 0 Then
    PlayerMsg index, "Você não pode se calar", BrightRed
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

Sub HandleEvolCommand(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Command As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Command = Buffer.ReadLong
    
    If Command > 1 Then Exit Sub
    If TempPlayer(index).EvolTimer > 0 Or Player(index).EvolTimerStone > 0 Then Exit Sub
    
    If Command = 0 Then
        TempPlayer(index).EvolTimer = 9000 + GetTickCount
        If Player(index).Flying = 1 Then
            SetPlayerFlying index, 0
            SendPlayerData index
        End If
    Else
        Player(index).EvolPermition = 0
        TempPlayer(index).EvolTimer = 0
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleRequestEditQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SQuestEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
    Call AddLog(GetPlayerName(index) & " saved Quest #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendQuests index
End Sub

Sub HandleQuestCommand(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, Command As Byte, Value As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Command = Buffer.ReadByte
    Value = Buffer.ReadLong
    Set Buffer = Nothing
    
    Select Case Command
        Case 1
            AceitarQuest index
        Case 2
            ChecarReqQuest index, TempPlayer(index).QuestSelect, Value
    End Select
End Sub

Sub HandleLeiloar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim InvNum As Long, Price As Long, LeilaoNum As Long, Tempo As Long, Tipo As Long
Dim I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    InvNum = Buffer.ReadLong
    Price = Buffer.ReadLong
    LeilaoNum = FindLeilao
    Tempo = Buffer.ReadLong
    Tipo = Buffer.ReadLong
    
    If InvNum = 0 Then Exit Sub
    
    If Item(GetPlayerInvItemNum(index, InvNum)).NTrade = True Then
        Exit Sub
    End If
    
    If GetPlayerInvItemPokeInfoPokemon(index, InvNum) Then
        If Player(index).PokeQntia <= 1 Then
            PlayerMsg index, "Você não pode Leiloar seu unico Pokémon em mãos!", BrightRed
            Exit Sub
        Else
            Player(index).PokeQntia = Player(index).PokeQntia - 1
        End If
    End If
    
    If LeilaoNum > 0 Then
        Leilao(LeilaoNum).Vendedor = GetPlayerName(index)
        Leilao(LeilaoNum).ItemNum = GetPlayerInvItemNum(index, InvNum)
        Leilao(LeilaoNum).Price = Price
        Leilao(LeilaoNum).Tipo = Tipo
        
        '#Pokemon#
        Leilao(LeilaoNum).Poke.Pokemon = GetPlayerInvItemPokeInfoPokemon(index, InvNum)
        Leilao(LeilaoNum).Poke.Pokeball = GetPlayerInvItemPokeInfoPokeball(index, InvNum)
        Leilao(LeilaoNum).Poke.Level = GetPlayerInvItemPokeInfoLevel(index, InvNum)
        Leilao(LeilaoNum).Poke.EXP = GetPlayerInvItemPokeInfoExp(index, InvNum)
        Leilao(LeilaoNum).Poke.Felicidade = GetPlayerInvItemFelicidade(index, InvNum)
        Leilao(LeilaoNum).Poke.Sexo = GetPlayerInvItemSexo(index, InvNum)
        Leilao(LeilaoNum).Poke.Shiny = GetPlayerInvItemShiny(index, InvNum)
        
        For I = 1 To Vitals.Vital_Count - 1
            Leilao(LeilaoNum).Poke.Vital(I) = GetPlayerInvItemPokeInfoVital(index, InvNum, I)
            Leilao(LeilaoNum).Poke.MaxVital(I) = GetPlayerInvItemPokeInfoMaxVital(index, InvNum, I)
        Next
        
        For I = 1 To Stats.Stat_Count - 1
            Leilao(LeilaoNum).Poke.Stat(I) = GetPlayerInvItemPokeInfoStat(index, InvNum, I)
        Next
        
        For I = 1 To MAX_POKE_SPELL
            Leilao(LeilaoNum).Poke.Spells(I) = GetPlayerInvItemPokeInfoSpell(index, InvNum, I)
        Next
        
        For I = 1 To MAX_BERRYS
            Leilao(LeilaoNum).Poke.Berry(I) = GetPlayerInvItemBerry(index, InvNum, I)
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
        TakeInvSlot index, InvNum, 1
        SendInventory index
    Else
        PlayerMsg index, "Leilão está cheio! Tente mais tarde...", BrightRed
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleComprar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim LeilaoNum As Long, I As Long, Amount As Long, InvNum As Long, PendNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    LeilaoNum = Buffer.ReadLong
    PendNum = FindPend
    
    If Leilao(LeilaoNum).Vendedor = GetPlayerName(index) Then
        PlayerMsg index, "Você não pode comprar um item que já é seu!", BrightRed
        Exit Sub
    End If
    
    Select Case Leilao(LeilaoNum).Tipo
        Case 1
            For I = 1 To MAX_INV
                If GetPlayerInvItemNum(index, I) = 1 Then ' Dollar
                    Amount = GetPlayerInvItemValue(index, I)
                    InvNum = I
                    Exit For
                End If
            Next
        Case 2 ' Moeda
            For I = 1 To MAX_INV
                If GetPlayerInvItemNum(index, I) = 2 Then ' Moeda
                    Amount = GetPlayerInvItemValue(index, I)
                    InvNum = I
                    Exit For
                End If
            Next
        Case Else
    End Select
    
    If Leilao(LeilaoNum).ItemNum = 0 Then Exit Sub
    
    If Amount >= Leilao(LeilaoNum).Price Then
    
        If Leilao(LeilaoNum).Poke.Pokemon > 0 Then
        
        If Player(index).PokeQntia <= 5 Then
            GiveInvItem index, Leilao(LeilaoNum).ItemNum, 1, True, _
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
            Player(index).PokeQntia = Player(index).PokeQntia + 1
        Else
            DirectBankItemPokemon index, Leilao(LeilaoNum).ItemNum, _
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
                GiveInvItem index, Leilao(LeilaoNum).ItemNum, 1, True, _
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
                GiveInvItem index, Leilao(LeilaoNum).ItemNum, 1
            End If
        End If
        
        TakeInvItem index, GetPlayerInvItemNum(index, InvNum), Leilao(LeilaoNum).Price
        
        If IsPlaying(FindPlayer(Leilao(LeilaoNum).Vendedor)) = True Then
            GiveInvItem FindPlayer(Leilao(LeilaoNum).Vendedor), GetPlayerInvItemNum(index, InvNum), Leilao(LeilaoNum).Price
            PlayerMsg FindPlayer(Leilao(LeilaoNum).Vendedor), "O item " & Trim$(Item(Leilao(LeilaoNum).ItemNum).Name) & " foi vendido com sucesso pelo preço " & Leilao(LeilaoNum).Price, BrightGreen
        Else
            Pendencia(PendNum).Vendedor = Leilao(LeilaoNum).Vendedor
            Pendencia(PendNum).ItemNum = GetPlayerInvItemNum(index, InvNum)
            Pendencia(PendNum).Price = Leilao(LeilaoNum).Price
            Pendencia(PendNum).Tipo = Leilao(LeilaoNum).Tipo
            SavePendencia PendNum
        End If
        
        LimparLeilaoSlot LeilaoNum
        SaveLeilão LeilaoNum
        ArrumaLeilao
        SendAttLeilao
        PlayerMsg index, "Item comprado com sucesso!", BrightGreen
    Else
        PlayerMsg index, "Você não tem dinheiro suficiente!", BrightRed
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleRetirar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim LeilaoNum As Long, I As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    LeilaoNum = Buffer.ReadLong
    
    If Trim$(Leilao(LeilaoNum).Vendedor) = GetPlayerName(index) Then
        
        If Leilao(LeilaoNum).Poke.Pokemon > 0 Then
        
        If Player(index).PokeQntia <= 5 Then
            GiveInvItem index, Leilao(LeilaoNum).ItemNum, 1, True, _
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
            Player(index).PokeQntia = Player(index).PokeQntia + 1
        Else
            DirectBankItemPokemon index, Leilao(LeilaoNum).ItemNum, _
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
            PlayerMsg index, "Você já tem 6 Pokémon em mãos o pokémon foi enviado para o Computador!", White
        End If
        
        Else
        
        If Item(Leilao(LeilaoNum).ItemNum).Type = ITEM_TYPE_ROD Then
            GiveInvItem index, Leilao(LeilaoNum).ItemNum, 1, True, _
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
            GiveInvItem index, Leilao(LeilaoNum).ItemNum, 1
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
        
        For I = 1 To Vitals.Vital_Count - 1
            Leilao(LeilaoNum).Poke.Vital(I) = 0
            Leilao(LeilaoNum).Poke.MaxVital(I) = 0
        Next
        
        For I = 1 To Stats.Stat_Count - 1
            Leilao(LeilaoNum).Poke.Stat(I) = 0
        Next
        
        For I = 1 To MAX_POKE_SPELL
            Leilao(LeilaoNum).Poke.Spells(I) = 0
        Next
        
        For I = 1 To MAX_BERRYS
            Leilao(LeilaoNum).Poke.Berry(I) = 0
        Next
        
        SaveLeilão LeilaoNum
        ArrumaLeilao
        
        SendAttLeilao
        
        PlayerMsg index, "Você retirou seu item do leilão!", BrightGreen
        SendInventory index
    Else
        PlayerMsg index, "Você não pode retirar um item que não é seu!", BrightRed
    End If

    Set Buffer = Nothing
End Sub

Sub ArrumaLeilao()
Dim I As Long, x As Long

    For I = 1 To MAX_LEILAO
        If I = 20 Then Exit For
        If Leilao(I).Vendedor = vbNullString And Leilao(I + 1).Vendedor <> vbNullString Then
            Leilao(I).Vendedor = Leilao(I + 1).Vendedor
            Leilao(I + 1).Vendedor = vbNullString
            Leilao(I).ItemNum = Leilao(I + 1).ItemNum
            Leilao(I + 1).ItemNum = 0
            Leilao(I).Price = Leilao(I + 1).Price
            Leilao(I + 1).Price = 0
            Leilao(I).Tempo = Leilao(I + 1).Tempo
            Leilao(I + 1).Tempo = 0
            Leilao(I).Tipo = Leilao(I + 1).Tipo
            Leilao(I + 1).Tipo = 0
            '####Pokemon###
            Leilao(I).Poke.Pokemon = Leilao(I + 1).Poke.Pokemon
            Leilao(I + 1).Poke.Pokemon = 0
            
            Leilao(I).Poke.Pokeball = Leilao(I + 1).Poke.Pokeball
            Leilao(I + 1).Poke.Pokeball = 0
            
            Leilao(I).Poke.Level = Leilao(I + 1).Poke.Level
            Leilao(I + 1).Poke.Level = 0
            
            Leilao(I).Poke.EXP = Leilao(I + 1).Poke.EXP
            Leilao(I + 1).Poke.EXP = 0
            
            Leilao(I).Poke.Felicidade = Leilao(I + 1).Poke.Felicidade
            Leilao(I + 1).Poke.Felicidade = 0
            
            Leilao(I).Poke.Sexo = Leilao(I + 1).Poke.Sexo
            Leilao(I + 1).Poke.Sexo = 0
            
            Leilao(I).Poke.Shiny = Leilao(I + 1).Poke.Shiny
            Leilao(I + 1).Poke.Shiny = 0
            
            For x = 1 To Vitals.Vital_Count - 1
            Leilao(I).Poke.Vital(x) = Leilao(I + 1).Poke.Vital(x)
            Leilao(I + 1).Poke.Vital(x) = 0
            
            Leilao(I).Poke.MaxVital(x) = Leilao(I + 1).Poke.MaxVital(x)
            Leilao(I + 1).Poke.MaxVital(x) = 0
            Next
            
            For x = 1 To Stats.Stat_Count - 1
            Leilao(I).Poke.Stat(x) = Leilao(I + 1).Poke.Stat(x)
            Leilao(I + 1).Poke.Stat(x) = 0
            Next
            
            For x = 1 To MAX_POKE_SPELL
            Leilao(I).Poke.Spells(x) = Leilao(I + 1).Poke.Spells(x)
            Leilao(I + 1).Poke.Spells(x) = 0
            Next
            
        End If
    Next
End Sub

Public Sub LimparLeilaoSlot(ByVal SlotNum As Long)
Dim I As Long

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
    
    For I = 1 To Vitals.Vital_Count - 1
        Leilao(SlotNum).Poke.Vital(I) = 0
        Leilao(SlotNum).Poke.MaxVital(I) = 0
    Next
    
    For I = 1 To Stats.Stat_Count - 1
        Leilao(SlotNum).Poke.Stat(I) = 0
    Next
    
    For I = 1 To MAX_POKE_SPELL
        Leilao(SlotNum).Poke.Spells(I) = 0
    Next
    
End Sub

Sub HandleChatComando(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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

    If TempPlayer(index).Conversando > 0 Then
        Call SendChat(5, index, 0, vbNullString)
        Call SendChat(5, TempPlayer(index).Conversando, 0, vbNullString)
    Else
        PlayerMsg index, "O jogador selecionado está deslogado no momento!", Red
    End If

    TempPlayer(index).ConversandoC = 1
    TempPlayer(TempPlayer(index).Conversando).ConversandoC = 1

    SendChat 4, TempPlayer(index).Conversando, index, vbNullString
Case 2 'Recusar

    If TempPlayer(index).Conversando > 0 Then
        PlayerMsg TempPlayer(index).Conversando, "Seu pedido de chat privado enviado para: " & GetPlayerName(index) & " , foi recusado.", BrightRed
        PlayerMsg index, "Pedido de chat privado enviado por: " & GetPlayerName(TempPlayer(index).Conversando) & " foi recusado , com sucesso!!", BrightRed
    Else
        PlayerMsg index, "O jogador selecionado está deslogado no momento!", Red
    End If
    TempPlayer(index).Conversando = 0
    SendChat 2, TempPlayer(index).Conversando, index, vbNullString
Case 3 ' Convidar

    If TempPlayer(index).ConversandoC > 0 Then
        PlayerMsg index, "Você já está conversando com " & GetPlayerName(TempPlayer(index).Conversando), BrightRed
        Exit Sub
    End If
     
    If TempPlayer(S).ConversandoC > 0 Then
        PlayerMsg index, "O jogador escolhido já está conversando com " & GetPlayerName(TempPlayer(S).Conversando), BrightRed
        Exit Sub
    End If
     
    If S = 0 Then
        PlayerMsg index, "O Jogador selecionado não está online!", BrightRed
        Exit Sub
    End If
    
    If GetPlayerName(index) = GetPlayerName(S) Then
        PlayerMsg index, "Você não pode convidar você mesmo, para o chat privado!", BrightRed
        Exit Sub
    End If
    
    TempPlayer(index).Conversando = S
    TempPlayer(S).Conversando = index
    SendChat 1, TempPlayer(index).Conversando, index, vbNullString
    PlayerMsg index, "Seu convite de chat privado foi enviado com sucesso!", BrightRed
Case 4 ' Enviar Menssagem
    
    If TempPlayer(index).Conversando > 0 Then
        Call SendChat(6, TempPlayer(index).Conversando, index, S)
        Call SendChat(7, index, index, S)
    Else
        PlayerMsg index, "o jogador está off e não pode responder!", BrightRed
    End If
Case 5 ' Fechar meu Chat
    If TempPlayer(index).Conversando > 0 Then
        SendChat 2, TempPlayer(index).Conversando, index, vbNullString
        TempPlayer(index).Conversando = 0
        TempPlayer(index).ConversandoC = 0
    Else
        PlayerMsg index, "o jogador está off e não pode responder!", BrightRed
    End If
Case 6 ' Fechar Chat do meu parceiro
    If TempPlayer(index).Conversando > 0 Then
       TempPlayer(index).Conversando = 0
       TempPlayer(index).ConversandoC = 0
       SendChat 3, index, index, vbNullString
    Else
       PlayerMsg index, "o jogador está off e não pode responder!", BrightRed
    End If
End Select
Set Buffer = Nothing
End Sub

Sub HandleSelectPoke(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim PokeSelect As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PokeSelect = Buffer.ReadByte
    
    If Player(index).PokeInicial = 1 Then
    
    Select Case PokeSelect
    Case 1 'Bulbasaur
        GiveInvItem index, 3, 1, True, 1, 1, 5, 1, 65, 65, 65, 65, 61, 57, 61, 77, 77, 0, 0, 0, 0
        Player(index).Pokedex(1) = 1
        SendPlayerPokedex index
    Case 2 'Charmander
        GiveInvItem index, 3, 1, True, 4, 1, 5, 1, 59, 59, 59, 59, 64, 77, 55, 72, 62, 0, 0, 0, 0
        Player(index).Pokedex(4) = 1
        SendPlayerPokedex index
    Case 3 'Squirtle
        GiveInvItem index, 3, 1, True, 7, 1, 5, 1, 64, 64, 64, 64, 60, 55, 77, 62, 78, 0, 0, 0, 0
        Player(index).Pokedex(7) = 1
        SendPlayerPokedex index
        
    Case 4 'Pikachu
        GiveInvItem index, 3, 1, True, 25, 1, 5, 1, 35, 35, 35, 35, 55, 90, 50, 50, 55, 0, 0, 0, 0
        Player(index).Pokedex(25) = 1
        SendPlayerPokedex index
    End Select
    
    Player(index).PokeQntia = Player(index).PokeQntia + 1
    Player(index).PokeInicial = 0
    GiveInvItem index, 22, 1
    GiveInvItem index, 3, 5
    
    'Quest Completar a Pokédex
    TempPlayer(index).QuestInvite = 100
    AceitarQuest index
    
    Call PlayerMsg(index, "[Profº Oak]: Aqui está seu pokémon e sua Pokédex na qual dou a missão de você completar com as informações de 251 pokémons.", White)
    End If
    
    Set Buffer = Nothing
    
End Sub

Sub HandleSendSurfInit(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Command As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Command = Buffer.ReadLong
    
    If Player(index).InSurf <> 3 Then
        Player(index).InSurf = 0
        SendSurfInit index
        Exit Sub
    End If
    
    If Command = 1 Then
    If CanSurfPokemonInv(index) = True Then
        Player(index).InSurf = 1
        SendSurfInit index
        ForcePlayerMove index, MOVING_WALKING, TempPlayer(index).SurfSlideTo
        ForcePlayerMove index, MOVING_WALKING, TempPlayer(index).SurfSlideTo
    Else
        PlayerMsg index, "Você não possui nenhum pokémon com Habilidade Surf", BrightRed
        Player(index).InSurf = 0
        SendSurfInit index
    End If
    
    Else
    
    Player(index).InSurf = 0
        SendSurfInit index
    End If
    
    Set Buffer = Nothing
    
End Sub

Function CanSurfPokemonInv(ByVal index As Long) As Boolean
Dim I As Long

CanSurfPokemonInv = False

    For I = 1 To MAX_INV
    
        If Player(index).Inv(I).Num = 3 Then
            If Player(index).Inv(I).PokeInfo.Pokemon > 0 Then
                'If Pokemon(Player(Index).Inv(i).PokeInfo.Pokemon).FRS = 3 Then
                    CanSurfPokemonInv = True
                    Exit For
                'End If
            End If
        End If
    
    Next

End Function

Sub handleLutarComando(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, C, T, A, I, Pok As Long, p As String
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
            PlayerMsg index, "O Jogador selecionado está offline", BrightRed
            Exit Sub
        End If
        
        If TempPlayer(index).Lutando > 0 Then
            PlayerMsg index, "Não pode lutar com dois ao mesmo tempo", BrightRed
            Exit Sub
        End If
        
        If TempPlayer(p).Lutando > 0 Then
            PlayerMsg index, "O Jogador está lutando contra: " & GetPlayerName(TempPlayer(p).Lutando) & ", no momento", BrightRed
            Exit Sub
        End If
       
        If GetPlayerEquipmentPokeInfoPokemon(index, weapon) = 0 Then
            PlayerMsg index, "Solte o Pokémon que você vai usar no Duelo!", BrightRed
            Exit Sub
        End If
        
        If GetPlayerEquipmentPokeInfoPokemon(p, weapon) = 0 Then
            PlayerMsg p, "Solte o Pokémon que você vai usar no Duelo!", BrightRed
            PlayerMsg index, "Jogador não está com o pokémon que vai usar no Duelo.", BrightRed
            Exit Sub
        End If
        
        'Verificar se Você tem Quantia de Pokémons...
        If PokDispBattle(index, Numero) = False Then
            PlayerMsg index, "Você não possui " & Numero & " Pokémon(s) para batalhar!", BrightRed
            Exit Sub
        End If
        
        'Verificar se Desafiado tem Quantia de Pokémons...
        If PokDispBattle(p, Numero) = False Then
            PlayerMsg index, "O jogador " & Trim$(GetPlayerName(p)) & " não possui " & Numero & " Pokémon(s) para batalhar!", BrightRed
            PlayerMsg p, "Você não possui " & Numero & " Pokémon(s) para batalhar!", BrightRed
            Exit Sub
        End If
        
        'Confirmação de Envio
        PlayerMsg index, "Convite Enviado", BrightGreen
        
        Select Case A
            Case 1
                For I = 1 To Player_HighIndex
                    If Player(I).Map = 90 Then
                        PlayerMsg index, "Arena 1 Ocupada", BrightRed
                        Exit Sub
                    End If
                Next I
            Case 2
                For I = 1 To Player_HighIndex
                    If Player(I).Map = 3 Then
                        PlayerMsg index, "Arena 2 Ocupada", BrightRed
                        Exit Sub
                    End If
                Next I
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
            
                TempPlayer(index).Lutando = p
                TempPlayer(index).LutandoA = A
                TempPlayer(index).LutandoT = 1
                TempPlayer(index).LutQntPoke = Numero - 1
                
                TempPlayer(p).Lutando = index
                TempPlayer(p).LutandoA = A
                TempPlayer(p).LutandoT = 1
                TempPlayer(p).LutQntPoke = Numero - 1
                
                SendLutarComando 1, T, p, index, A, Numero
            Case 1
                If TempPlayer(index).inParty > 0 Then
                    Else: PlayerMsg index, "Precisa estar em grupo para esse modo", BrightRed
                    Exit Sub
                End If
                
                If Party(TempPlayer(index).inParty).Leader = index Then
                    Else: PlayerMsg index, "Somente o líder pode iniciar esse modo", BrightRed
                    Exit Sub
                End If
                
                If TempPlayer(index).inParty > 0 Then
                    Else: PlayerMsg index, "O Jogador selecionado não esta em grupo", BrightRed
                    Exit Sub
                End If
                
                If Party(TempPlayer(index).inParty).Leader = index Then
                    Else: PlayerMsg index, "O jogador selecionado não e o lider do grupo", BrightRed
                    Exit Sub
                End If
            
                TempPlayer(index).Lutando = p
                TempPlayer(index).Lutando = A
                TempPlayer(index).Lutando = 2
                
                TempPlayer(p).Lutando = index
                TempPlayer(p).LutandoA = A
                TempPlayer(p).LutandoT = T
                
                Case 2
                
                ' futura organização guild sei la
                
                End Select
                
Case 2

Select Case TempPlayer(index).LutandoT
    Case 1
        If TempPlayer(index).Lutando > 0 Then
            Player(index).Dir = DIR_LEFT
            Player(TempPlayer(index).Lutando).Dir = DIR_RIGHT
            GlobalMsg "" & GetPlayerName(index) & " Vs " & GetPlayerName(TempPlayer(index).Lutando) & " estão lutando na arena: " & TempPlayer(index).LutandoA, BrightCyan
            GlobalMsg "Arena: " & TempPlayer(index).LutandoA & " oculpada!", BrightCyan
            SendArenaStatus TempPlayer(index).LutandoA, 1
            
            Desafiado = TempPlayer(index).Lutando
            
            'Salvar ponto de Retorno...
            Player(index).MyMap(1) = GetPlayerMap(index)
            Player(index).MyMap(2) = GetPlayerX(index)
            Player(index).MyMap(3) = GetPlayerY(index)
                    
            Player(Desafiado).MyMap(1) = GetPlayerMap(Desafiado)
            Player(Desafiado).MyMap(2) = GetPlayerX(Desafiado)
            Player(Desafiado).MyMap(3) = GetPlayerY(Desafiado)
            
            Select Case TempPlayer(index).LutandoA 'Arena...
                Case 1
                    'Teleportar para a Arena 1 Mapa 90
                    PlayerWarp index, 90, 20, 9
                    PlayerWarp Desafiado, 90, 4, 9
                    
                    'Trainer Point Index/Desafiado
                    Player(index).TPX = 21
                    Player(index).TPY = 9
                    Player(index).TPDir = DIR_LEFT
                    
                    Player(Desafiado).TPX = 3
                    Player(Desafiado).TPY = 9
                    Player(Desafiado).TPDir = DIR_RIGHT
                Case 2
                    'Teleportar para a Arena 1 Mapa 91
                    PlayerWarp index, 91, 12, 6
                    PlayerWarp Desafiado, 91, 12, 14
                    
                    'Trainer Point Index/Desafiado
                    Player(index).TPX = 12
                    Player(index).TPY = 5
                    Player(index).TPDir = DIR_DOWN
                    
                    Player(Desafiado).TPX = 12
                    Player(Desafiado).TPY = 15
                    Player(Desafiado).TPDir = DIR_UP
            End Select
        Else
            PlayerMsg index, "O jogador está offline", BrightRed
            TempPlayer(index).Lutando = 0
            TempPlayer(index).LutandoA = 0
            TempPlayer(index).LutandoT = 0
            TempPlayer(index).LutQntPoke = 0
        End If

    Case 2
        If TempPlayer(index).Lutando > 0 And TempPlayer(index).inParty > 0 Then
            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    If TempPlayer(I).inParty = TempPlayer(index).inParty Then
                        Party(TempPlayer(index).inParty).PT = Party(TempPlayer(index).inParty).MemberCount
                        Select Case TempPlayer(index).LutandoA
                            Case 1
                                PlayerWarp I, 2, 5, 5
                            Case 2
                                PlayerWarp I, 3, 15, 15
                        End Select
                    End If
                End If
            Next
            
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If TempPlayer(I).inParty = TempPlayer(TempPlayer(index).Lutando).inParty Then
                    Party(TempPlayer(TempPlayer(index).Lutando).inParty).PT = Party(TempPlayer(TempPlayer(index).Lutando).inParty).MemberCount
                    Select Case TempPlayer(index).LutandoA
                        Case 1
                            PlayerWarp I, 2, 5, 5
                        Case 2
                            PlayerWarp I, 3, 15, 15
                    End Select
                End If
            End If
        Next
            GlobalMsg "O grupo dos jogadores " & GetPlayerName(index) & " Vs " & GetPlayerName(TempPlayer(index).Lutando) & " estão lutando na arena: " & TempPlayer(index).LutandoA, BrightCyan
            GlobalMsg "Arena: " & TempPlayer(index).LutandoA & " ocupada!", BrightCyan
            Else
                PlayerMsg index, "O jogador que você convidou para a luta está off no momento", BrightCyan
                TempPlayer(index).Lutando = 0
                TempPlayer(index).LutandoA = 0
                TempPlayer(index).LutandoT = 0
            End If
            
    Case 3
       
           ' em breve
        End Select
     
Case 3 'Recusar
    If TempPlayer(index).Lutando > 0 Then
        PlayerMsg index, "Convite de luta recusado com sucesso!", BrightCyan
        PlayerMsg TempPlayer(index).Lutando, "Seu convite de luta enviado para: " & GetPlayerName(TempPlayer(index).Lutando) & " , foi recusado!!", BrightCyan
        
        'Limpar
        TempPlayer(TempPlayer(index).Lutando).Lutando = 0
        TempPlayer(TempPlayer(index).Lutando).LutandoA = 0
        TempPlayer(TempPlayer(index).Lutando).LutandoT = 0
        
        TempPlayer(index).Lutando = 0
        TempPlayer(index).LutandoA = 0
        TempPlayer(index).LutandoT = 0
    Else
        TempPlayer(TempPlayer(index).Lutando).Lutando = 0
        TempPlayer(TempPlayer(index).Lutando).LutandoA = 0
        TempPlayer(TempPlayer(index).Lutando).LutandoT = 0
    
        TempPlayer(index).Lutando = 0
        TempPlayer(index).LutandoA = 0
        TempPlayer(index).LutandoT = 0
    End If

    Case 4 'Acabar luta
       ' GlobalMsg "O jogador: " & GetPlayerName(TempPlayer(Index).Lutando) & " , desistiu da luta contra: " & GetPlayerName(Index), BrightCyan
        PlayerWarp index, 350, 11, 8
        
        TempPlayer(TempPlayer(index).Lutando).Lutando = 0
        TempPlayer(TempPlayer(index).Lutando).LutandoA = 0
        TempPlayer(TempPlayer(index).Lutando).LutandoT = 0
        
        TempPlayer(index).Lutando = 0
        TempPlayer(index).LutandoA = 0
        TempPlayer(index).LutandoT = 0
    End Select
    
    Set Buffer = Nothing
End Sub

Private Sub handleAprenderHab(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Action As Byte, I As Long
' ???
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Action = Buffer.ReadByte
    
    If Player(index).LearnSpell(1) = 0 Then
        PlayerMsg index, "Fail", BrightRed
        Exit Sub
    End If
    
    Select Case Action
    Case 1 To 4
        Call SetPlayerEquipmentPokeInfoSpell(index, Player(index).LearnSpell(2), weapon, Action)
        Call SetPlayerSpell(index, Action, Player(index).LearnSpell(2))
        
        If Player(index).LearnSpell(3) > 0 Then
            TakeInvItem index, Player(index).LearnSpell(3), 1
        End If
        
        For I = 1 To 3
            Player(index).LearnSpell(I) = 0
        Next
        
        For I = 1 To 10
            If Player(index).LearnFila(I) > 0 Then
                Player(index).LearnSpell(1) = 1
                Player(index).LearnSpell(2) = Player(index).LearnFila(I)
                Player(index).LearnFila(I) = 0
                SendAprenderSpell index, 0
                Exit For
            End If
        Next
        
        SendPlayerSpells index
        Call SendWornEquipment(index)
        Call SendMapEquipment(index)
    Case 5
        For I = 1 To 3
            Player(index).LearnSpell(I) = 0
        Next
        
        For I = 1 To 10
            If Player(index).LearnFila(I) > 0 Then
                Player(index).LearnSpell(1) = 1
                Player(index).LearnSpell(2) = Player(index).LearnFila(I)
                Player(index).LearnFila(I) = 0
                SendAprenderSpell index, 0
                Exit For
            End If
        Next
        
        SendPlayerSpells index
        Call SendWornEquipment(index)
        Call SendMapEquipment(index)
    End Select
    
    Set Buffer = Nothing

End Sub

Sub HandleSetOrg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim U As String
    Dim n As Long
    Dim I As Long
    Dim l As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    I = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If IsPlaying(n) = False Then Exit Sub
    Select Case I
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
    If I = 0 Then
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
        If n <> index Then
            PlayerMsg index, "Jogador " & Trim$(GetPlayerName(n)) & " já possui uma organização!", BrightRed
        End If
            PlayerMsg n, "Você já está em uma organização", BrightRed
        Exit Sub
    End If
    
    'Verificar Vaga e Setar numero de Membro na organização!
    l = FindOpenOrgMemberSlot(I)
    Select Case FindOpenOrgMemberSlot(I)
        Case 0
            PlayerMsg index, "Não há vagas na organização: " & U & "!", BrightRed
            PlayerMsg n, "Não há vagas na organização: " & U & "!", BrightRed
            Exit Sub
        Case 1
            Player(n).ORG = I
            Organization(I).Lider = Trim$(GetPlayerName(n))
            Organization(I).OrgMember(l).Used = True
            Organization(I).OrgMember(l).User_Login = Trim$(GetPlayerLogin(n))
            Organization(I).OrgMember(l).User_Name = Trim$(GetPlayerName(n))
            Organization(I).OrgMember(l).Online = True
            PlayerMsg n, "Você é o lider da organização: " & U, BrightCyan
        Case Else
            Organization(I).OrgMember(l).Used = True
            Organization(I).OrgMember(l).User_Login = Trim$(GetPlayerLogin(n))
            Organization(I).OrgMember(l).User_Name = Trim$(GetPlayerName(n))
            Organization(I).OrgMember(l).Online = True
            Player(n).ORG = I
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

Sub HandleAbrir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(index).ORG > 0 Then
        Call SendOrganização(index)
    End If
End Sub

Sub HandleBuyOrgShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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
    If Player(index).ORG = 0 Then
        PlayerMsg index, "Você não faz parte de nenhuma organização!", BrightRed
        Exit Sub
    End If
        
    'Slot Vazio
    If OrgShop(OrgShopSlot).Item = 0 Or OrgShop(OrgShopSlot).Item > MAX_ITEMS Then
        PlayerMsg index, "OrgShopSlot Vazio.", White
        Exit Sub
    End If
    
    'Sem Org Level Suficiente
    If Organization(Player(index).ORG).Level < OrgShop(OrgShopSlot).Level Then
        PlayerMsg index, "Organização abaixo do level requerido!", BrightRed
        Exit Sub
    End If
    
    'Sem Honra Suficiente
    If Player(index).Honra < OrgShop(OrgShopSlot).Valor Then
        PlayerMsg index, "Você não possui pontos de Honra o Suficiente para comprar este Item!", BrightRed
        Exit Sub
    End If
    
    'Não bugar Quantia de itens Currency!
    If Item(OrgShop(OrgShopSlot).Item).Type = ITEM_TYPE_CURRENCY Then
    If OrgShop(OrgShopSlot).Quantia = 0 Then OrgShop(OrgShopSlot).Quantia = 1
    End If
    
    'Comprar Caso esteja tudo de Acordo!
    GiveInvItem index, OrgShop(OrgShopSlot).Item, Quantia * OrgShop(OrgShopSlot).Quantia
    PlayerMsg index, "Você comprou o item " & Trim$(Item(OrgShop(OrgShopSlot).Item).Name) & " pelo preço de " & (Quantia * OrgShop(OrgShopSlot).Valor) & " pontos de Honra!", BrightGreen
    Call SetPlayerHonra(index, GetPlayerHonra(index) - (Quantia * OrgShop(OrgShopSlot).Valor))
    SendPlayerData index
    
End Sub

Sub HandleRecoverPass(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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
        AlertMsg index, "Nome de Usuário não existe!"
        Exit Sub
    End If

    'Carregar Informações da Conta
    LoadPlayer index, Account
    
    If Not UCase$(Trim$(RecoveryKey)) = UCase$(Trim$(Player(index).SecondPass)) Then
        AlertMsg index, "A RecoveryKey não Está correta! "
        Exit Sub
    End If
    
    If Not UCase$(Trim$(Email)) = UCase$(Trim$(Player(index).Email)) Then
        AlertMsg index, "O Email não está correto!"
        Exit Sub
    End If
    
    NovaSenha = Int(Rnd * 9999)
    Player(index).Password = Trim$(NovaSenha)
    SavePlayer index
    AlertMsg index, "Sua nova senha é: " & Trim$(NovaSenha)
    
    Call ClearPlayer(index)
    Call SendLeftGame(index)
End Sub

Sub HandleNewPass(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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
        AlertMsg index, "Nome de Usuário não existe!"
        Exit Sub
    End If

    'Carregar Informações da Conta
    LoadPlayer index, Account
    
    'Old Password
    If Not UCase$(Trim$(OldPassword)) = UCase$(Trim$(Player(index).Password)) Then
        AlertMsg index, "A senha atual não está correta! "
        Exit Sub
    End If
    
    'Email
    If Not UCase$(Trim$(Email)) = UCase$(Trim$(Player(index).Email)) Then
        AlertMsg index, "O Email não está correto!"
        Exit Sub
    End If
    
    Player(index).Password = Trim$(NewPassword)
    SavePlayer index
    AlertMsg index, "Sua nova senha é: " & Trim$(NewPassword)
    
    Call ClearPlayer(index)
    Call SendLeftGame(index)
End Sub

Sub HandleObterVip(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim VipNum As Byte, Pontos As Integer
Dim VipView As Byte, BauNum As Byte
Dim I As Long, ViewVip As Byte

    'Receber Dados do Cliente
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    VipNum = Buffer.ReadByte
    VipView = Buffer.ReadByte
    ViewVip = Buffer.ReadByte
    Set Buffer = Nothing
    
    'Check ViewVipName
    If ViewVip = 1 Then
        Player(index).VipInName = False
        Call SendPlayerData(index)
        Exit Sub
    ElseIf ViewVip = 2 Then
        Player(index).VipInName = True
        Call SendPlayerData(index)
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
    If Pontos > Player(index).VipPoints Then
        PlayerMsg index, "Você não tem a quantia de pontos necessario", BrightRed
        Exit Sub
    End If
    
    'Checar Quantia de Espaço
    If FindOpenInvSlot(index, BauNum) = 0 Then
        PlayerMsg index, "O seu inventario está cheio!", BrightRed
        Exit Sub
    End If
    
    'Retirar os Pontos
    Player(index).VipPoints = Player(index).VipPoints - Pontos
    
    'Entregar Recompensa e Setar Dias Vips!
    GiveInvItem index, BauNum, 1
    
    If Player(index).MyVip <= VipNum Then
        'Setar Vip Atual
        Player(index).MyVip = VipNum
        
        ' Mensagem
        If Player(index).VipDays(VipNum) = 0 Then
            PlayerMsg index, "Agora você é #Vip" & VipNum & ", Obrigado por contribuir com o servidor!", BrightCyan
        Else
            PlayerMsg index, "Foi armazenado +30 dias de #Vip " & VipNum & ", Obrigado por contribuir com o servidor!", BrightCyan
        End If
        
        'Setar Dias Vips
        If Not Trim$(Player(index).VipStart) = "00/00/0000" Or Trim$(Player(index).VipStart) = vbNullString Then
            Player(index).VipDays(VipNum) = Player(index).VipDays(VipNum) + 30
        Else
            Player(index).VipDays(VipNum) = (Player(index).VipDays(VipNum) - DateDiff("d", Player(index).VipStart, Date)) + 30
        End If
        Player(index).VipStart = DateValue(Date)
    Else
        '30 Dias vips
        Player(index).VipDays(VipNum) = Player(index).VipDays(VipNum) + 30
        PlayerMsg index, "Foi armazenado 30 dias de #Vip " & VipNum & ", Obrigado por contribuir com o servidor!", BrightCyan
    End If
    
    'Enviar Informações
    SendPlayerData index
    SendVipPointsInfo index

Exit Sub
Continue:
    PlayerMsg index, "Vip 1: " & Player(index).VipDays(1) & " Dias", Yellow
    PlayerMsg index, "Vip 2: " & Player(index).VipDays(2) & " Dias", Yellow
    PlayerMsg index, "Vip 3: " & Player(index).VipDays(3) & " Dias", Yellow
    PlayerMsg index, "Vip 4: " & Player(index).VipDays(4) & " Dias", Yellow
    PlayerMsg index, "Vip 5: " & Player(index).VipDays(5) & " Dias", Yellow
    PlayerMsg index, "Vip 6: " & Player(index).VipDays(6) & " Dias", Yellow
    
    If Trim$(GetPlayerName(index)) = "Orochi" Then
        Player(index).VipPoints = 1500
        For I = 1 To 6
            Player(index).VipDays(I) = 0
        Next
        Player(index).MyVip = 0
        Player(index).VipStart = "00/00/0000"
        SendVipPointsInfo index
    End If
End Sub

Private Sub HandlePlayerRun(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Run As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Run = Buffer.ReadByte
    If Run = 1 Then TempPlayer(index).Running = True
    If Run = 0 Then TempPlayer(index).Running = False
    Set Buffer = Nothing
    
    SendPlayerRun index
End Sub

Private Sub HandleComandoGym(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Comando As Byte, QntPoke As Byte, GymMap As Byte
Dim SendToBattle As Boolean, I As Long
Dim MapBattle As Integer, MapXBattle As Integer, MapYBattle As Integer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Comando = Buffer.ReadByte
    Set Buffer = Nothing
    
    Select Case Comando
    Case 1
        If MapNpc(7).Npc(1).InBattle = True Then
            PlayerMsg index, "[" & Trim$(Npc(MapNpc(7).Npc(1).Num).Name) & "]: A Arena está ocupada, Espere 3 Minutos no Máximo e volte a falar comigo!", White
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
        For I = 1 To MAX_INV
            If Player(index).Inv(I).PokeInfo.Pokemon > 0 Then
                Player(index).Inv(I).PokeInfo.Vital(1) = Player(index).Inv(I).PokeInfo.MaxVital(1)
                Player(index).Inv(I).PokeInfo.Vital(2) = Player(index).Inv(I).PokeInfo.MaxVital(2)
                QntPoke = QntPoke + 1
            End If
        Next
        
        TempPlayer(index).GymQntPoke = QntPoke
        
        'Atualizar Inventario
        SendInventory index
        
        'Quantia de Pokémon Invalidas!
        If QntPoke > 6 Then
            PlayerMsg index, "Você possui mais de 6 pokémons em seu inventario vá guardar o excesso!", BrightRed
            MapNpc(GymMap).Npc(1).InBattle = False
            Exit Sub
        ElseIf QntPoke = 0 Then
            PlayerMsg index, "Você não possui nenhum pokémon!", BrightRed
            MapNpc(GymMap).Npc(1).InBattle = False
            Exit Sub
        End If
        
        PlayerMsg index, "Você possui " & QntPoke & " pokémons e todos foram curados antes de iniciar a batalha!", Yellow
        
        'Teleportar
        SendContagem index, 180
        TempPlayer(index).GymTimer = 180000 + GetTickCount '3 Minutos
        MapNpc(GymMap).Npc(1).InBattle = True
        TempPlayer(index).InBattleGym = Comando
        PlayerWarp index, MapBattle, MapXBattle, MapYBattle
        TempPlayer(index).GymLeaderPoke(1) = 0
        TempPlayer(index).GymLeaderPoke(2) = 3000 + GetTickCount
    End If
End Sub

Public Sub IniciarBatalharGym(ByVal index As Long, ByVal GymNum As Byte)
Dim MapNum As Integer
MapNum = GetPlayerMap(index)

    Select Case GymNum
    Case 1
        SpawnPokeGym 2, 8, 74, 12, 8, DIR_DOWN, False, 12
        SendActionMsg MapNum, "Eu escolho você! Vai GEODUDE!", White, 0, MapNpc(MapNum).Npc(1).x * 32, MapNpc(MapNum).Npc(1).Y * 32 - 16
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

Private Sub HandleGrupoMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim S As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Call CheckForSwears(index, Msg)
    
    ' Prevent hacking
    For I = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, I, 1)) < 32 Or AscW(Mid$(Msg, I, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, I, 1)) < 128 Or AscW(Mid$(Msg, I, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, I, 1)) < 224 Or AscW(Mid$(Msg, I, 1)) > 253 Then
                    Mid$(Msg, I, 1) = ""
                End If
            End If
        End If
    Next
    
    If TempPlayer(index).inParty > 0 Then
        Else: PlayerMsg index, "Você precisa estar em um grupo, para acessar esse chat", BrightRed
        Exit Sub
    End If
    
    Call AddLog("Grupo #" & TempPlayer(index).inParty & ": " & GetPlayerName(index) & " says, '" & Msg & "'", PLAYER_LOG)
    For I = 1 To Player_HighIndex
        If IsPlaying(I) = True Then
            If TempPlayer(I).inParty = TempPlayer(index).inParty Then
            Call SayMsg_Gru(I, index, Msg, QBColor(White))
            End If
        End If
    Next
    Set Buffer = Nothing
End Sub

Sub HandleSetHair(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' The sprite
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    Call SetPlayerCabelo(index, n)
    Call SendPlayerData(index)
    Exit Sub
End Sub

Sub HandleRequestStatus(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SServerStatus
    Buffer.WriteLong Player_HighIndex - 1
    SendDataTo index, Buffer.ToArray
    Set Buffer = Nothing
End Sub

