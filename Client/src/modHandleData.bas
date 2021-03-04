Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetAddress = FunAddr

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub InitMessages()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SPokemonEditor) = GetAddress(AddressOf HandlePokemonEditor)
    HandleDataSub(SUpdatePokemon) = GetAddress(AddressOf HandleUpdatePokemon)
    HandleDataSub(SPlayerPokedex) = GetAddress(AddressOf HandlePlayerPokedex)
    HandleDataSub(SPokeEvo) = GetAddress(AddressOf HandlePokeEvo)
    HandleDataSub(SFishing) = GetAddress(AddressOf handleInFishing)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdateQuest)
    HandleDataSub(SQuestCommand) = GetAddress(AddressOf HandleQuestCommand)
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SDialogue) = GetAddress(AddressOf HandleDialogue)
    HandleDataSub(SLeiloar) = GetAddress(AddressOf HandleAttLeilao)
    HandleDataSub(SCChat) = GetAddress(AddressOf HandleCChat)
    HandleDataSub(SPokeSelect) = GetAddress(AddressOf handlePokeSelect)
    HandleDataSub(SSurfInit) = GetAddress(AddressOf HandleSendSurfInit)
    HandleDataSub(SUpdateRankLevel) = GetAddress(AddressOf HandleUpdateRankLevel)
    HandleDataSub(SCLutar) = GetAddress(AddressOf HandleCLuta)
    HandleDataSub(SArenas) = GetAddress(AddressOf HandleArenas)
    HandleDataSub(SAprender) = GetAddress(AddressOf HandleAprenderSpell)
    HandleDataSub(SNoticia) = GetAddress(AddressOf HandleNoticia)
    HandleDataSub(SOrganização) = GetAddress(AddressOf HandleAttOrg)
    HandleDataSub(SOrgShop) = GetAddress(AddressOf HandleOrgShop)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SVipInfo) = GetAddress(AddressOf HandleVipPlayerInfo)
    'HandleDataSub(SAparencia) = GetAddress(AddressOf HandleAparencia)
    HandleDataSub(SRunning) = GetAddress(AddressOf HandlePlayerRun)
    HandleDataSub(SComandGym) = GetAddress(AddressOf HandleComandoGym)
    HandleDataSub(SContagem) = GetAddress(AddressOf HandleContagem)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Dim MsgType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    MsgType = buffer.ReadLong

    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If

    CallWindowProc HandleDataSub(MsgType), MyIndex, buffer.ReadBytes(buffer.length), 0, 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim buffer As clsBuffer
    Dim i As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    frmLoad.Visible = False
    frmMenu.Visible = True
    frmMenu.picLogin.Visible = True
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.PicRecover.Visible = False
    frmMenu.PicNewPass.Visible = False

    For i = 0 To 2
        frmMenu.txtRecover(i).text = vbNullString
    Next

    For i = 0 To 4
        frmMenu.txtNewPass(i).text = vbNullString
    Next

    Msg = buffer.ReadString    'Parse(1)

    Set buffer = Nothing
    Call MsgBox(Msg, vbOKOnly, Options.Game_Name)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' save options
    Options.SavePass = frmMenu.chkPass.value
    Options.Username = Trim$(frmMenu.txtLUser.text)

    If frmMenu.chkPass.value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.text)
    End If

    SaveOptions

    ' Now we can receive game data
    MyIndex = buffer.ReadLong

    ' player high index
    Player_HighIndex = buffer.ReadLong

    Set buffer = Nothing
    frmLoad.Visible = True
    Call SetStatus("Receiving game data...")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, X As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = buffer.ReadString
            .Vital(Vitals.HP) = buffer.ReadLong
            .Vital(Vitals.MP) = buffer.ReadLong

            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next

            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing

    ' Used for if the player is creating a new character
    frmMenu.Visible = True
    frmMenu.picCharacter.Visible = True
    frmMenu.picLogin.Visible = False
    frmMenu.picRegister.Visible = False
    frmLoad.Visible = False
    frmMenu.cmbClass.Clear
    For i = 1 To Max_Classes
        frmMenu.cmbClass.AddItem Trim$(Class(i).Name)
    Next

    frmMenu.cmbClass.ListIndex = 0
    n = frmMenu.cmbClass.ListIndex + 1

    newCharSprite = 0
    NewCharacterBltSprite

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, X As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong    'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = buffer.ReadString    'Trim$(Parse(n))
            .Vital(Vitals.HP) = buffer.ReadLong    'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = buffer.ReadLong    'CLng(Parse(n + 2))

            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next

            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InGame = True
    Call GameInit
    Call GameLoop

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim X As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemPokeInfoPokemon(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemPokeInfoPokeball(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemPokeInfoLevel(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemPokeInfoExp(MyIndex, i, buffer.ReadLong)

        For X = 1 To Vitals.Vital_Count - 1
            Call SetPlayerInvItemPokeInfoVital(MyIndex, i, buffer.ReadLong, X)
            Call SetPlayerInvItemPokeInfoMaxVital(MyIndex, i, buffer.ReadLong, X)
        Next

        For X = 1 To Stats.Stat_Count - 1
            Call SetPlayerInvItemPokeInfoStat(MyIndex, i, X, buffer.ReadLong)
        Next

        For X = 1 To 4
            Call SetPlayerInvItemPokeInfoSpell(MyIndex, i, buffer.ReadLong, X)
        Next

        For X = 1 To MAX_NEGATIVES
            Call SetPlayerInvItemNgt(MyIndex, i, X, buffer.ReadLong)
        Next

        For X = 1 To MAX_BERRYS
            Call SetPlayerInvItemBerry(MyIndex, i, X, buffer.ReadLong)
        Next

        Call SetPlayerInvItemFelicidade(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemSexo(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemShiny(MyIndex, i, buffer.ReadLong)

        n = n + 2
    Next

    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0    ' clear

    Set buffer = Nothing
    BltInventory

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long, X As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong    'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadLong)    'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong)    'CLng(Parse(3)))

    Call SetPlayerInvItemPokeInfoPokemon(MyIndex, n, buffer.ReadLong)
    Call SetPlayerInvItemPokeInfoPokeball(MyIndex, n, buffer.ReadLong)
    Call SetPlayerInvItemPokeInfoLevel(MyIndex, n, buffer.ReadLong)
    Call SetPlayerInvItemPokeInfoExp(MyIndex, n, buffer.ReadLong)

    For X = 1 To Vitals.Vital_Count - 1
        Call SetPlayerInvItemPokeInfoVital(MyIndex, n, buffer.ReadLong, X)
        Call SetPlayerInvItemPokeInfoMaxVital(MyIndex, n, buffer.ReadLong, X)
    Next

    For X = 1 To Stats.Stat_Count - 1
        Call SetPlayerInvItemPokeInfoStat(MyIndex, n, X, buffer.ReadLong)
    Next

    For X = 1 To 4
        Call SetPlayerInvItemPokeInfoSpell(MyIndex, n, buffer.ReadLong, X)
    Next

    For X = 1 To MAX_NEGATIVES
        Call SetPlayerInvItemNgt(MyIndex, X, weapon, buffer.ReadLong)
    Next

    For X = 1 To MAX_BERRYS
        Call SetPlayerInvItemBerry(MyIndex, n, X, buffer.ReadLong)
    Next

    Call SetPlayerInvItemFelicidade(MyIndex, n, buffer.ReadLong)
    Call SetPlayerInvItemSexo(MyIndex, n, buffer.ReadLong)
    Call SetPlayerInvItemShiny(MyIndex, n, buffer.ReadLong)

    ' changes, clear drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0    ' clear

    BltInventory
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    'PokeInfo Armor
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Armor)

    'PokeInfo Weapon
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, weapon)
    Call SetPlayerEquipmentPokeInfoPokemon(MyIndex, buffer.ReadLong, weapon)
    Call SetPlayerEquipmentPokeInfoPokeball(MyIndex, buffer.ReadLong, weapon)
    Call SetPlayerEquipmentPokeInfoLevel(MyIndex, buffer.ReadLong, weapon)
    Call SetPlayerEquipmentPokeInfoExp(MyIndex, buffer.ReadLong, weapon)

    For i = 1 To Vitals.Vital_Count - 1
        Call SetPlayerEquipmentPokeInfoVital(MyIndex, buffer.ReadLong, weapon, i)
        Call SetPlayerEquipmentPokeInfoMaxVital(MyIndex, buffer.ReadLong, weapon, i)
    Next

    For i = 1 To Stats.Stat_Count - 1
        Call SetPlayerEquipmentPokeInfoStat(MyIndex, buffer.ReadLong, weapon, i)
    Next

    For i = 1 To 4
        Call SetPlayerEquipmentPokeInfoSpell(MyIndex, buffer.ReadLong, weapon, i)
    Next

    For i = 1 To MAX_NEGATIVES
        Call SetPlayerEquipmentNgt(MyIndex, i, weapon, buffer.ReadLong)
    Next

    For i = 1 To MAX_BERRYS
        Call SetPlayerEquipmentBerry(MyIndex, buffer.ReadLong, weapon, i)
    Next

    Call SetPlayerEquipmentFelicidade(MyIndex, weapon, buffer.ReadLong)
    Call SetPlayerEquipmentSexo(MyIndex, weapon, buffer.ReadLong)
    Call SetPlayerEquipmentShiny(MyIndex, weapon, buffer.ReadLong)

    'PokeInfo Helmet
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Helmet)

    'PokeInfo Shield
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Shield)

    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0    ' clear

    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
        For i = 1 To 4
            If GetPlayerEquipmentPokeInfoSpell(MyIndex, weapon, i) > 0 Then
                frmMain.lblPokeSpell(i).Caption = Trim$(Spell(GetPlayerEquipmentPokeInfoSpell(MyIndex, weapon, i)).Name)
            Else
                frmMain.lblPokeSpell(i).Caption = "Null"
            End If
            frmMain.PicPokeSpell(i).Visible = True
        Next

    Else
        For i = 1 To 4
            frmMain.PicPokeSpell(i).Visible = False
        Next
    End If

    If GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon) > 0 Then
        If GetPlayerEquipmentShiny(MyIndex, weapon) = 0 Then
            frmMain.lblCharName = Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)).Name)
        Else
            frmMain.lblCharName = "Shiny " & Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)).Name)
        End If
    Else
        ' Set the character windows
        frmMain.lblCharName = GetPlayerName(MyIndex)
    End If

    BltInventory
    BltEquipment
    BltPokeEquip

    'Resetar Anim
    Player(MyIndex).AnimFrame = 0
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim PlayerNum As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer

    buffer.WriteBytes Data()

    PlayerNum = buffer.ReadLong

    'PokeInfo Armor
    Call SetPlayerEquipment(PlayerNum, buffer.ReadLong, Armor)

    'PokeInfo Weapon
    Call SetPlayerEquipment(PlayerNum, buffer.ReadLong, weapon)
    Call SetPlayerEquipmentPokeInfoPokemon(PlayerNum, buffer.ReadLong, weapon)
    Call SetPlayerEquipmentPokeInfoPokeball(PlayerNum, buffer.ReadLong, weapon)
    Call SetPlayerEquipmentPokeInfoLevel(PlayerNum, buffer.ReadLong, weapon)
    Call SetPlayerEquipmentPokeInfoExp(PlayerNum, buffer.ReadLong, weapon)

    For i = 1 To Vitals.Vital_Count - 1
        Call SetPlayerEquipmentPokeInfoVital(PlayerNum, buffer.ReadLong, weapon, i)
        Call SetPlayerEquipmentPokeInfoMaxVital(PlayerNum, buffer.ReadLong, weapon, i)
    Next

    For i = 1 To Stats.Stat_Count - 1
        Call SetPlayerEquipmentPokeInfoStat(PlayerNum, buffer.ReadLong, weapon, i)
    Next

    For i = 1 To 4
        Call SetPlayerEquipmentPokeInfoSpell(PlayerNum, buffer.ReadLong, weapon, i)
    Next

    For i = 1 To MAX_NEGATIVES
        Call SetPlayerEquipmentNgt(MyIndex, i, weapon, buffer.ReadLong)
    Next

    For i = 1 To MAX_BERRYS
        Call SetPlayerEquipmentBerry(PlayerNum, buffer.ReadLong, weapon, i)
    Next

    Call SetPlayerEquipmentFelicidade(PlayerNum, weapon, buffer.ReadLong)
    Call SetPlayerEquipmentSexo(PlayerNum, weapon, buffer.ReadLong)
    Call SetPlayerEquipmentShiny(PlayerNum, weapon, buffer.ReadLong)

    'PokeInfo Helmet
    Call SetPlayerEquipment(PlayerNum, buffer.ReadLong, Helmet)

    'PokeInfo Shield
    Call SetPlayerEquipment(PlayerNum, buffer.ReadLong, Shield)

    'Resetar Anim
    Player(PlayerNum).AnimFrame = 0
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.HP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
        'frmMain.lblHP.Caption = Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%"
        frmMain.lblHP.Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
        ' hp bar
        frmMain.imgHPBar.Width = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBar_Width)) * HPBar_Width
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerHP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.MP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
        'frmMain.lblMP.Caption = Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%"
        frmMain.lblMp.Caption = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
        ' mp bar
        frmMain.imgMPBar.Width = ((GetPlayerVital(MyIndex, Vitals.MP) / SPRBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / SPRBar_Width)) * SPRBar_Width
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, i, buffer.ReadLong
        'frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim TNL As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    SetPlayerExp MyIndex, buffer.ReadLong

    TNL = buffer.ReadLong

    ' mp bar
    If GetPlayerEquipment(Index, weapon) > 0 Then
        If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
            frmMain.imgEXPBar.Width = ((GetPlayerExp(MyIndex) / EXPBar_Width) / (TNL / EXPBar_Width)) * EXPBar_Width
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerExp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, X As Long
    Dim buffer As clsBuffer
    Dim AntHonra As Long, VipNameView As Byte

    AntHonra = GetPlayerHonra(MyIndex)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Call SetPlayerName(i, buffer.ReadString)
    Call SetPlayerLevel(i, buffer.ReadLong)
    Call SetPlayerPOINTS(i, buffer.ReadLong)
    Call SetPlayerSprite(i, buffer.ReadLong)
    Call SetPlayerMap(i, buffer.ReadLong)
    Call SetPlayerX(i, buffer.ReadLong)
    Call SetPlayerY(i, buffer.ReadLong)
    Call SetPlayerDir(i, buffer.ReadLong)
    Call SetPlayerAccess(i, buffer.ReadLong)
    Call SetPlayerPK(i, buffer.ReadLong)
    Call SetPlayerFlying(i, buffer.ReadLong)
    Player(i).TPX = buffer.ReadLong
    Player(i).TPY = buffer.ReadLong
    Player(i).TPDir = buffer.ReadLong
    Player(i).TPSprite = buffer.ReadLong
    Player(i).Vitorias = buffer.ReadLong
    Player(i).Derrotas = buffer.ReadLong
    Player(i).ORG = buffer.ReadByte
    Call SetPlayerHonra(i, buffer.ReadLong)
    Player(i).MyVip = buffer.ReadByte
    VipNameView = buffer.ReadByte

    If VipNameView = 1 Then
        Player(i).VipInName = True
    Else
        Player(i).VipInName = False
    End If

    If buffer.ReadByte = 1 Then
        Player(i).PokeLight = True
    Else
        Player(i).PokeLight = False
    End If

    For X = 1 To MAX_INSIGNIAS
        Player(i).Insignia(X) = buffer.ReadLong
    Next

    For X = 1 To MAX_QUESTS
        Player(i).Quests(X).status = buffer.ReadByte
        Player(i).Quests(X).Part = buffer.ReadByte
    Next

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False

        ' Set training label visiblity depending on points
        frmMain.lblPoints.Caption = GetPlayerPOINTS(MyIndex)
        If GetPlayerPOINTS(MyIndex) > 0 Then
            For X = 1 To Stats.Stat_Count - 1
                If GetPlayerStat(Index, X) < 255 Then
                    frmMain.lblTrainStat(X).Visible = True
                Else
                    frmMain.lblTrainStat(X).Visible = False
                End If
            Next
        Else
            For X = 1 To Stats.Stat_Count - 1
                frmMain.lblTrainStat(X).Visible = False
            Next
        End If

    End If

    'Vip Check
    If Player(Index).VipInName = True Then
        frmMain.PicChkVipName.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")
    Else
        frmMain.PicChkVipName.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
    End If

    'Quest
    PlayerQuests

    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).XOffset = 0
    Player(i).YOffset = 0

    If GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon) > 0 Then
        If GetPlayerEquipmentShiny(MyIndex, weapon) = 0 Then
            frmMain.lblCharName = Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)).Name)
        Else
            frmMain.lblCharName = "Shiny " & Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)).Name)
        End If
    Else
        frmMain.lblCharName = GetPlayerName(MyIndex)
    End If

    'Carregar Insignias
    Call CarregarInsignia

    'Atualizar Honra se necessario!
    If AntHonra <> GetPlayerHonra(MyIndex) Then
        BltOrgShop
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim n As Byte
    Dim buffer As clsBuffer
    Dim S As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    n = buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, Dir)
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)
    Case DIR_UP
        Player(i).YOffset = PIC_Y
    Case DIR_DOWN
        Player(i).YOffset = PIC_Y * -1
    Case DIR_LEFT
        Player(i).XOffset = PIC_X
    Case DIR_RIGHT
        Player(i).XOffset = PIC_X * -1
    End Select

    'Check to see if the map tile is Grass or not
    If Player(i).Flying = 0 Then
        If Map.Tile(X, Y).Type = TILE_TYPE_GRASS Then
            MeAnimation 10, GetPlayerX(i), GetPlayerY(i)
        End If

        If Map.Tile(X, Y).Type = TILE_TYPE_SLIDE Then
            Player(i).PuloStatus = 1
            Player(i).PuloSlide = 15
        End If
    End If

    'Check to see if the map tile is Grass or not
    If Player(i).InSurf = 1 Then
        MeAnimation 14, GetPlayerX(i), GetPlayerY(i)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNpcNum As Long, FazerRuido As Byte
    Dim X As Long, PokemonId As String
    Dim Y As Long
    Dim Dir As Long
    Dim Movement As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapNpcNum = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Movement = buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
        .Dir = Dir
        .XOffset = 0
        .YOffset = 0
        .Moving = Movement

        Select Case .Dir
        Case DIR_UP
            .YOffset = PIC_Y
        Case DIR_DOWN
            .YOffset = PIC_Y * -1
        Case DIR_LEFT
            .XOffset = PIC_X
        Case DIR_RIGHT
            .XOffset = PIC_X * -1
        End Select

        'Check to see if the map tile is Grass or not
        If Map.Tile(X, Y).Type = TILE_TYPE_GRASS Then
            MeAnimation 10, .X, .Y
        End If

        If MapNpc(MapNpcNum).num > 0 Then
            If Npc(MapNpc(MapNpcNum).num).Pokemon > 0 Then
                'Verificar se Ira fazer Ruido!
                FazerRuido = 100 * Rnd
                If FazerRuido <= 10 Then
                    Select Case Npc(MapNpc(MapNpcNum).num).Pokemon
                    Case 1 To 9
                        PokemonId = "00" & Npc(MapNpc(MapNpcNum).num).Pokemon
                    Case 10 To 99
                        PokemonId = "0" & Npc(MapNpc(MapNpcNum).num).Pokemon
                    Case Else
                        PokemonId = Npc(MapNpc(MapNpcNum).num).Pokemon
                    End Select

                    If isInRange(5, Player(MyIndex).X, Player(MyIndex).Y, X, Y) = True Then
                        PlaySound "PokeSounds\" & PokemonId & ".mp3", -1, -1
                    End If
                End If
            End If
        End If

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Dir As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerDir(i, Dir)

    With Player(i)
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Dir As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong

    With MapNpc(i)
        .Dir = Dir
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, Dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).XOffset = 0
    Player(MyIndex).YOffset = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim buffer As clsBuffer
    Dim thePlayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    thePlayer = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, Dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).XOffset = 0
    Player(thePlayer).YOffset = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong

    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim NeedMap As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap

    ' Get map num
    X = buffer.ReadLong
    ' Get revision
    Y = buffer.ReadLong

    If FileExist(MAP_PATH & "map" & X & MAP_EXT, False) Then
        Call LoadMap(X)
        ' Check to see if the revisions match
        NeedMap = 1

        If Map.Revision = Y Then
            ' We do so we dont need the map
            'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
            NeedMap = 0
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong NeedMap
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        frmMain.picCharacter.Visible = True

        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Dim MapNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer

    buffer.WriteBytes Data()

    MapNum = buffer.ReadLong
    Map.Name = buffer.ReadString
    Map.Music = buffer.ReadString
    Map.Revision = buffer.ReadLong
    Map.Moral = buffer.ReadByte
    Map.Up = buffer.ReadLong
    Map.Down = buffer.ReadLong
    Map.Left = buffer.ReadLong
    Map.Right = buffer.ReadLong
    Map.BootMap = buffer.ReadLong
    Map.BootX = buffer.ReadByte
    Map.BootY = buffer.ReadByte
    Map.MaxX = buffer.ReadByte
    Map.MaxY = buffer.ReadByte
    Map.Weather = buffer.ReadLong
    Map.Intensity = buffer.ReadLong

    For X = 1 To 2
        Map.LevelPoke(X) = buffer.ReadLong
    Next

    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Tileset = buffer.ReadLong
            Next
            Map.Tile(X, Y).Type = buffer.ReadByte
            Map.Tile(X, Y).Data1 = buffer.ReadLong
            Map.Tile(X, Y).Data2 = buffer.ReadLong
            Map.Tile(X, Y).Data3 = buffer.ReadLong
            Map.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.Npc(X) = buffer.ReadLong
        n = n + 1
    Next

    ClearTempTile

    Set buffer = Nothing

    ' Save the map
    Call SaveMap(MapNum)

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False

        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_ITEMS
        With MapItem(i)
            '.playerName = Buffer.ReadString
            .num = buffer.ReadLong
            .value = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
        End With
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .num = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .Dir = buffer.ReadLong
            .Vital(HP) = buffer.ReadLong
            .Sexo = buffer.ReadLong
            .Shiny = buffer.ReadLong
            .Level = buffer.ReadLong
        End With
    Next

    SaveMap i

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapDone()
    Dim i As Long
    Dim MusicFile As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
    Action_HighIndex = 1

    ' load tilesets we need
    LoadTilesets

    MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        PlayMusic MusicFile
    Else
        StopMusic
    End If

    ' re-position the map name
    Call UpdateDrawMapName

    ' get the npc high index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).num > 0 Then
            Npc_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS

    GettingMap = False
    CanMoveNow = True

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte


    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)



    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long, i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong

    With MapItem(n)
        '.playerName = Buffer.ReadString
        .num = buffer.ReadLong
        .value = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .PokeInfo.Pokemon = buffer.ReadLong
        .PokeInfo.Pokeball = buffer.ReadLong
        .PokeInfo.Level = buffer.ReadLong
        .PokeInfo.Exp = buffer.ReadLong

        For i = 1 To Vitals.Vital_Count - 1
            .PokeInfo.Vital(i) = buffer.ReadLong
            .PokeInfo.MaxVital(i) = buffer.ReadLong
        Next

        For i = 1 To Stats.Stat_Count - 1
            .PokeInfo.Stat(i) = buffer.ReadLong
        Next

        For i = 1 To 4
            .PokeInfo.Spells(i) = buffer.ReadLong
        Next

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub HandleItemEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimationEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0    ' clear

    BltInventory

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong

    With MapNpc(n)
        .num = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadLong
        .Sexo = buffer.ReadLong
        .Shiny = buffer.ReadLong
        .Level = buffer.ReadLong

        ' Client use only
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNpcNum As Long
    Dim buffer As clsBuffer
    Dim Morto As Long
    Dim Sumir As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Sumir = buffer.ReadByte
    MapNpcNum = buffer.ReadLong
    Morto = buffer.ReadLong

    If MapNpcNum > 0 Then
        If Morto = 1 Then
            MapNpc(MapNpcNum).Desmaiado = True
        Else
            MapNpc(MapNpcNum).Desmaiado = False
        End If
    End If

    If Sumir = 1 Then
        Call ClearMapNpc(MapNpcNum)
        MapNpc(MapNpcNum).Desmaiado = False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    n = buffer.ReadLong

    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ResourceNum = buffer.ReadLong

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)

    ClearResource ResourceNum

    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    n = buffer.ReadByte
    TempTile(X, Y).DoorOpen = n

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapKey", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditMap()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MapEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleShopEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopnum As Long
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    shopnum = buffer.ReadLong

    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpellEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    SpellNum = buffer.ReadLong

    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing

    ' Update the spells on the pic
    Set buffer = New clsBuffer
    buffer.WriteLong CSpells
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = buffer.ReadLong
    Next

    BltPlayerSpells
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Call ClearPlayer(buffer.ReadLong)
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Resource_Index = buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = buffer.ReadByte
            MapResource(i).X = buffer.ReadLong
            MapResource(i).Y = buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
    Call DrawPing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDoorAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    With TempTile(X, Y)
        .DoorFrame = 1
        .DoorAnimate = 1    ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = GetTickCount
    End With
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDoorAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long, Message As String, color As Long, tmpType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer

    buffer.WriteBytes Data()
    Message = buffer.ReadString
    color = buffer.ReadLong
    tmpType = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing

    CreateActionMsg Message, color, tmpType, X, Y

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer

    buffer.WriteBytes Data()

    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1

    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .lockindex = buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With

    ' play the sound if we've got one
    If isInRange(6, GetPlayerX(MyIndex), GetPlayerY(MyIndex), AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y) = True Then
        PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation
    End If

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim MapNpcNum As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    MapNpcNum = buffer.ReadLong
    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = buffer.ReadLong
    Next

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    Slot = buffer.ReadLong
    SpellCD(Slot) = GetTickCount

    BltPlayerSpells
    blthotbar

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellBuffer = 0
    SpellBufferTimer = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Access As Long
Dim Name As String
Dim Message As String
Dim colour As Long
Dim Header As String
Dim PK As Long
Dim saycolour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    Access = buffer.ReadLong
    PK = buffer.ReadLong
    Message = buffer.ReadString
    Header = buffer.ReadString
    saycolour = buffer.ReadLong
    
    ' Check access level
    If PK = NO Then
        Select Case Access
            Case 0
                colour = RGB(255, 96, 0)
            Case 1
                colour = QBColor(DarkGrey)
            Case 2
                colour = QBColor(Cyan)
            Case 3
                colour = QBColor(BrightGreen)
            Case 4
                colour = QBColor(Yellow)
        End Select
    Else
        colour = QBColor(BrightRed)
    End If
    
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = colour
    frmMain.txtChat.SelText = vbNewLine & Header & Name & ": "
    frmMain.txtChat.SelColor = saycolour
    frmMain.txtChat.SelText = Message
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
        
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim shopnum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    shopnum = buffer.ReadLong

    Set buffer = Nothing

    OpenShop shopnum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ShopAction = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    StunDuration = buffer.ReadLong

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long, X As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_BANK
        Bank.Item(i).num = buffer.ReadLong
        Bank.Item(i).value = buffer.ReadLong
        Bank.Item(i).PokeInfo.Pokemon = buffer.ReadLong
        Bank.Item(i).PokeInfo.Pokeball = buffer.ReadLong
        Bank.Item(i).PokeInfo.Level = buffer.ReadLong
        Bank.Item(i).PokeInfo.Exp = buffer.ReadLong

        For X = 1 To Vitals.Vital_Count - 1
            Bank.Item(i).PokeInfo.Vital(X) = buffer.ReadLong
            Bank.Item(i).PokeInfo.MaxVital(X) = buffer.ReadLong
        Next

        For X = 1 To Stats.Stat_Count - 1
            Bank.Item(i).PokeInfo.Stat(X) = buffer.ReadLong
        Next

        For X = 1 To 4
            Bank.Item(i).PokeInfo.Spells(X) = buffer.ReadLong
        Next

        For X = 1 To MAX_NEGATIVES
            Bank.Item(i).PokeInfo.Negatives(X) = buffer.ReadLong
        Next

        For X = 1 To MAX_BERRYS
            Bank.Item(i).PokeInfo.Berry(X) = buffer.ReadLong
        Next

        Bank.Item(i).PokeInfo.Felicidade = buffer.ReadLong
        Bank.Item(i).PokeInfo.Sexo = buffer.ReadLong
        Bank.Item(i).PokeInfo.Shiny = buffer.ReadLong
    Next

    InBank = True
    frmMain.picBank.Visible = True
    frmMain.picBank.top = (frmMain.ScaleHeight / 2) - (frmMain.picBank.Height / 2)
    frmMain.picBank.Left = (frmMain.ScaleWidth / 2) - (frmMain.picBank.Width / 2)
    BltBank

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    InTrade = buffer.ReadLong
    frmMain.picDialogue.Visible = False
    frmMain.picTrade.Visible = True
    frmMain.picYourTrade.Visible = True
    frmMain.picTheirTrade.Visible = True
    frmMain.lblTradeStatus(0).Caption = "Esperando Confirmação"
    frmMain.lblTradeStatus(1).Caption = "Esperando Confirmação"
    frmMain.lblTradeStatus(0).ForeColor = &HE0E0E0
    frmMain.lblTradeStatus(1).ForeColor = &HE0E0E0
    frmMain.PicTradeOn(0).Visible = False
    frmMain.PicTradeOn(1).Visible = False
    BltTrade

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InTrade = 0
    frmMain.picTrade.Visible = False
    frmMain.picCurrency.Visible = False
    frmMain.lblTradeStatus(0).Caption = "Esperando Confirmação"
    frmMain.lblTradeStatus(1).Caption = "Esperando Confirmação"
    frmMain.lblTradeStatus(0).ForeColor = &HE0E0E0
    frmMain.lblTradeStatus(1).ForeColor = &HE0E0E0
    frmMain.PicTradeOn(0).Visible = False
    frmMain.PicTradeOn(1).Visible = False
    frmMain.picYourTrade.Visible = False
    frmMain.picTheirTrade.Visible = False

    ' re-blt any items we were offering
    BltInventory

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim dataType As Byte
    Dim i As Long, X As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    dataType = buffer.ReadByte

    If dataType = 0 Then    ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).num = buffer.ReadLong
            TradeYourOffer(i).value = buffer.ReadLong
            TradeYourOffer(i).PokeInfo.Pokemon = buffer.ReadLong
            TradeYourOffer(i).PokeInfo.Pokeball = buffer.ReadLong
            TradeYourOffer(i).PokeInfo.Level = buffer.ReadLong
            TradeYourOffer(i).PokeInfo.Exp = buffer.ReadLong
            TradeYourOffer(i).PokeInfo.Felicidade = buffer.ReadLong
            TradeYourOffer(i).PokeInfo.Sexo = buffer.ReadLong
            TradeYourOffer(i).PokeInfo.Shiny = buffer.ReadLong

            For X = 1 To Vitals.Vital_Count - 1
                TradeYourOffer(i).PokeInfo.Vital(X) = buffer.ReadLong
                TradeYourOffer(i).PokeInfo.MaxVital(X) = buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                TradeYourOffer(i).PokeInfo.Stat(X) = buffer.ReadLong
            Next

            For X = 1 To 4
                TradeYourOffer(i).PokeInfo.Spells(X) = buffer.ReadLong
            Next

            For X = 1 To MAX_BERRYS
                TradeYourOffer(i).PokeInfo.Berry(X) = buffer.ReadLong
            Next

        Next
        frmMain.lblYourWorth.Caption = buffer.ReadLong & "g"
        ' remove any items we're offering
        BltInventory
    ElseIf dataType = 1 Then    'theirs
        For i = 1 To MAX_INV

            TradeTheirOffer(i).num = buffer.ReadLong
            TradeTheirOffer(i).value = buffer.ReadLong
            TradeTheirOffer(i).PokeInfo.Pokemon = buffer.ReadLong
            TradeTheirOffer(i).PokeInfo.Pokeball = buffer.ReadLong
            TradeTheirOffer(i).PokeInfo.Level = buffer.ReadLong
            TradeTheirOffer(i).PokeInfo.Exp = buffer.ReadLong
            TradeTheirOffer(i).PokeInfo.Felicidade = buffer.ReadLong
            TradeTheirOffer(i).PokeInfo.Sexo = buffer.ReadLong
            TradeTheirOffer(i).PokeInfo.Shiny = buffer.ReadLong

            For X = 1 To Vitals.Vital_Count - 1
                TradeTheirOffer(i).PokeInfo.Vital(X) = buffer.ReadLong
                TradeTheirOffer(i).PokeInfo.MaxVital(X) = buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                TradeTheirOffer(i).PokeInfo.Stat(X) = buffer.ReadLong
            Next

            For X = 1 To 4
                TradeTheirOffer(i).PokeInfo.Spells(X) = buffer.ReadLong
            Next
        Next
        frmMain.lblTheirWorth.Caption = buffer.ReadLong & "g"
    End If

    BltTrade

    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tradeStatus As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    tradeStatus = buffer.ReadByte

    Set buffer = Nothing

    Select Case tradeStatus
    Case 0    ' clear
        frmMain.lblTradeStatus(0).Caption = "Esperando Confirmação"
        frmMain.lblTradeStatus(1).Caption = "Esperando Confirmação"
        frmMain.lblTradeStatus(0).ForeColor = &HE0E0E0
        frmMain.lblTradeStatus(1).ForeColor = &HE0E0E0
    Case 1    ' they've accepted
        frmMain.lblTradeStatus(1).Caption = "Troca Aceita"
        frmMain.PicTradeOn(1).Visible = True
    Case 2    ' you've accepted
        frmMain.lblTradeStatus(0).Caption = "Troca Aceita"
        frmMain.PicTradeOn(0).Visible = True
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim MapNpcNum As Integer, TargetVital As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    myTarget = buffer.ReadLong
    myTargetType = buffer.ReadLong
    MapNpcNum = buffer.ReadInteger
    TargetVital = buffer.ReadLong

    Set buffer = Nothing

    'Target Info
    If MapNpcNum = 0 Then
        frmMain.PicTarget.Visible = False
    Else
        If MapNpc(MapNpcNum).Shiny = False Then
            frmMain.lblTargetInfo(1).Caption = Trim$(Npc(MapNpc(MapNpcNum).num).Name)
        Else
            frmMain.lblTargetInfo(1).Caption = "S." & Trim$(Npc(MapNpc(MapNpcNum).num).Name)
        End If

        If Npc(MapNpc(MapNpcNum).num).Pokemon > 0 Then
            frmMain.ElementTarget(0).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\signs\" & Pokemon(Npc(MapNpc(MapNpcNum).num).Pokemon).Tipo(1) & ".jpg")
            frmMain.ElementTarget(1).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\signs\" & Pokemon(Npc(MapNpc(MapNpcNum).num).Pokemon).Tipo(2) & ".jpg")
        Else
            frmMain.ElementTarget(0).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\signs\0.jpg")
            frmMain.ElementTarget(1).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\signs\0.jpg")
        End If

        frmMain.lblTargetInfo(0).Caption = TargetVital & "/" & GetPokemonMaxVital(MapNpc(MapNpcNum).num, MapNpc(MapNpcNum).Level)
        bltTargetHp MapNpcNum, TargetVital
        bltPokemonTarget MapNpcNum
        frmMain.PicTarget.ZOrder 0
        frmMain.PicTarget.Visible = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = buffer.ReadLong
        Hotbar(i).sType = buffer.ReadByte
        Hotbar(i).Pokemon = buffer.ReadLong
        Hotbar(i).Pokeball = buffer.ReadLong
    Next
    blthotbar

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    Player_HighIndex = buffer.ReadLong

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long, entityType As Long, entityNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    X = buffer.ReadLong
    Y = buffer.ReadLong
    entityType = buffer.ReadLong
    entityNum = buffer.ReadLong

    PlayMapSound X, Y, entityType, entityNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim theName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    theName = buffer.ReadString

    Dialogue "Trade Request", theName & " has requested a trade. Would you like to accept?", DIALOGUE_TYPE_TRADE, True

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim theName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    theName = buffer.ReadString

    Dialogue "Party Invitation", theName & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, True

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long, inParty As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    inParty = buffer.ReadByte

    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' reset the labels
        For i = 1 To MAX_PARTY_MEMBERS
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        Next
        ' exit out early
        Exit Sub
    End If

    ' carry on otherwise
    Party.Leader = buffer.ReadLong
    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = buffer.ReadLong
        If Party.Member(i) > 0 Then
            frmMain.lblPartyMember(i).Caption = Trim$(GetPlayerName(Party.Member(i)))
            frmMain.imgPartyHealth(i).Visible = True
            frmMain.imgPartySpirit(i).Visible = True
        Else
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        End If
    Next
    Party.MemberCount = buffer.ReadLong

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim PlayerNum As Long, partyIndex As Long
    Dim buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' which player?
    PlayerNum = buffer.ReadLong
    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(PlayerNum).MaxVital(i) = buffer.ReadLong
        Player(PlayerNum).Vital(i) = buffer.ReadLong
    Next

    ' find the party number
    For i = 1 To MAX_PARTY_MEMBERS
        If Party.Member(i) = PlayerNum Then
            partyIndex = i
        End If
    Next

    ' exit out if wrong data
    If partyIndex <= 0 Or partyIndex > MAX_PARTY_MEMBERS Then Exit Sub

    ' hp bar
    frmMain.imgPartyHealth(partyIndex).Width = ((GetPlayerVital(PlayerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(PlayerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
    ' spr bar
    frmMain.imgPartySpirit(partyIndex).Width = ((GetPlayerVital(PlayerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(PlayerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePokeEvo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim PokemonId As Byte, Command As Byte
    PokemonId = GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Command = buffer.ReadByte
    Player(MyIndex).EvolPermition = buffer.ReadByte
    Player(MyIndex).EvoId = buffer.ReadInteger
    Set buffer = Nothing

    frmMain.lblEvol(0).Caption = "Seu " & Trim$(Pokemon(PokemonId).Name) & " está prestes a evoluir para " & Trim$(Pokemon(Pokemon(PokemonId).Evolução(Player(MyIndex).EvoId).Pokemon).Name) & " deseja continuar?"
    frmMain.PicEvolution.Visible = True
    frmMain.PicEvolution.top = (frmMain.ScaleHeight / 2) - (frmMain.PicEvolution.Height / 2)
    frmMain.PicEvolution.Left = (frmMain.ScaleWidth / 2) - (frmMain.PicEvolution.Width / 2)

    bltPokeEvolvePortrait

    If Command = 0 Then
        frmMain.imgClose(9).Visible = True
        frmMain.imgButton(14).Visible = True
    Else
        frmMain.EvolutionTimer = True
        PlaySound Sound_Evolve, -1, -1
        frmMain.imgClose(9).Visible = False
        frmMain.imgButton(14).Visible = False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TrainerPoint", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub handleInFishing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, FishingValue As Byte, ScanValue As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    FishingValue = buffer.ReadByte
    ScanValue = buffer.ReadByte

    If FishingValue = 1 Then
        Player(i).InFishing = GetTickCount    '10000 + GetTickCount
    Else
        Player(i).InFishing = 0
    End If

    If ScanValue = 1 Then
        Player(i).ScanTime = GetTickCount
    Else
        Player(i).ScanTime = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "handleInFishing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleUpdateQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateQuest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleQuestCommand(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Command As Byte, value As Long
    Dim QuestNum As Long, MapNum As Integer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Command = buffer.ReadByte
    value = buffer.ReadLong

    If Command = 2 Then
        QuestNum = buffer.ReadLong
        Player(Index).Quests(QuestNum).KillNpcs = buffer.ReadLong
        Player(Index).Quests(QuestNum).KillPlayers = buffer.ReadLong

        'Atualizar Caso A janela esteja aberta!
        If frmMain.picQuest.Visible = True Then
            If frmMain.lstQuests.ListCount > 0 Then
                UpdateQuestInfo GetQuestNum(Trim$(frmMain.lstQuests.text))
            End If
        End If
    End If

    Set buffer = Nothing

    Select Case Command
        ' Select npc quest
    Case 1
        frmMain.picSelectQuest.Visible = True

        frmMain.picSelectQuest.top = (frmMain.ScaleHeight / 2) - (frmMain.picSelectQuest.Height / 2)
        frmMain.picSelectQuest.Left = (frmMain.ScaleWidth / 2) - (frmMain.picSelectQuest.Width / 2)

        UpdateSelectQuest value
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleQuestCommand", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleQuestEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With frmEditor_Quest
        Editor = EDITOR_QUEST
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_QUESTS
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleQuestEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDialogue(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dialogue buffer.ReadString, buffer.ReadString, buffer.ReadByte, buffer.ReadLong
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDialogue", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttLeilao(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim LeilaoNum As Long, Vendedor As String, itemNum As Long, Price As Long, Tempo As Long, Tipo As Long
    Dim i As Long, X As Long, QntiaItens As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_LEILAO
        Leilao(i).Vendedor = buffer.ReadString
        Leilao(i).itemNum = buffer.ReadLong
        Leilao(i).Price = buffer.ReadLong
        Leilao(i).Tempo = buffer.ReadLong
        Leilao(i).Tipo = buffer.ReadLong
        'Pokemon
        Leilao(i).Poke.Pokemon = buffer.ReadLong
        Leilao(i).Poke.Pokeball = buffer.ReadLong
        Leilao(i).Poke.Level = buffer.ReadLong
        Leilao(i).Poke.Exp = buffer.ReadLong
        Leilao(i).Poke.Felicidade = buffer.ReadLong
        Leilao(i).Poke.Sexo = buffer.ReadLong
        Leilao(i).Poke.Shiny = buffer.ReadLong

        For X = 1 To Vitals.Vital_Count - 1
            Leilao(i).Poke.Vital(X) = buffer.ReadLong
            Leilao(i).Poke.MaxVital(X) = buffer.ReadLong
        Next

        For X = 1 To Stats.Stat_Count - 1
            Leilao(i).Poke.Stat(X) = buffer.ReadLong
        Next

        For X = 1 To 4
            Leilao(i).Poke.Spells(X) = buffer.ReadLong
        Next

        For X = 1 To MAX_NEGATIVES
            Leilao(i).Poke.Negatives(X) = buffer.ReadLong
        Next

        For X = 1 To MAX_BERRYS
            Leilao(i).Poke.Berry(X) = buffer.ReadLong
        Next
    Next

    PageMaxLeilao = 1

    For i = 1 To MAX_LEILAO
        If Leilao(i).itemNum > 0 Then
            QntiaItens = QntiaItens + 1
            If QntiaItens >= 20 Then
                QntiaItens = 0
                PageMaxLeilao = PageMaxLeilao + 1
            End If
        End If
    Next

    frmMain.lblLeilaoInfo(0).Caption = PageLeilao & "/" & PageMaxLeilao

    SendALeilao
    BltLeilao

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAttLeilao", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCChat(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, M, c As Long, T As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    c = buffer.ReadLong
    M = buffer.ReadLong
    T = buffer.ReadString

    Select Case c
    Case 1
        frmEditor_Chat.lblJogador.Caption = "Conversando com: " & GetPlayerName(M)
        Dialogue "Chat Privado", GetPlayerName(M) & " Convidou você para uma conversa privada , desejaria aceitar?", DIALOGUE_TYPE_PM, True
    Case 2
        SendChatComando 6, vbNullString
    Case 3
        frmEditor_Chat.Hide
        frmEditor_Chat.txtChat.text = vbNullString
    Case 4
        frmEditor_Chat.lblJogador.Caption = "Conversando com: " & GetPlayerName(M)
    Case 5
        frmEditor_Chat.Show
        frmEditor_Chat.txtChat.text = vbNullString
    Case 6
        frmEditor_Chat.txtChat.SelText = GetPlayerName(M) & " :" & T & vbNewLine
    Case 7
        frmEditor_Chat.txtChat.SelText = GetPlayerName(M) & " :" & T & vbNewLine
        frmEditor_Chat.txtEChat.text = vbNullString
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCChat", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub handlePokeSelect()
    Dim i As Long

    If frmMain.PicPokeInicial.Visible = True Then Exit Sub

    AddText "[Prof.Oak]: Olá, Seja Bem Vindo! Escolha o seu companheiro para viajar na sua Jornada!", White

    frmMain.PicPokeInicial.Visible = True
    frmMain.PicPokeInicial.top = (frmMain.ScaleHeight / 2) - (frmMain.PicPokeInicial.Height / 2)
    frmMain.PicPokeInicial.Left = (frmMain.ScaleWidth / 2) - (frmMain.PicPokeInicial.Width / 2)

    SelectPokeInicial = 1

    frmMain.lblPokeInicial(1).ForeColor = &HFF00&
    For i = 1 To 4
        If i <> 1 Then
            frmMain.lblPokeInicial(i).ForeColor = &HFFFFFF
        End If
    Next

End Sub

Private Sub HandleSendSurfInit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim PlayerNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PlayerNum = buffer.ReadLong
    Player(PlayerNum).InSurf = buffer.ReadByte

    If PlayerNum = Index Then
        If Player(Index).InSurf = 3 Then
            frmMain.PicSurf.Visible = True
            frmMain.PicSurf.top = (frmMain.ScaleHeight / 2) - (frmMain.PicSurf.Height / 2)
            frmMain.PicSurf.Left = (frmMain.ScaleWidth / 2) - (frmMain.PicSurf.Width / 2)
        Else
            frmMain.PicSurf.Visible = False
        End If
    End If

    Set buffer = Nothing

End Sub


Private Sub HandleUpdateRankLevel(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Position As Long
    Dim i As Byte
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_RANKS
        RankLevel(i).Name = buffer.ReadString
        RankLevel(i).Level = buffer.ReadLong
        RankLevel(i).PokeNum = buffer.ReadLong
    Next

    Set buffer = Nothing
    If RankOpen = 1 Then UpdateRankLevel

End Sub

Private Sub HandleCLuta(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, c, CC, a, T, Pok As Long

    Set buffer = New clsBuffer

    buffer.WriteBytes Data()

    c = buffer.ReadLong
    T = buffer.ReadLong
    CC = buffer.ReadLong
    a = buffer.ReadLong
    Pok = buffer.ReadLong

    Select Case c
    Case 1
        Select Case T
        Case 0
            Dialogue "Duelo " & Pok & " Pokémon(s)", GetPlayerName(CC) & " Desafiou você para uma batalha " & a & ", Aceitar ?", DIALOGUE_TYPE_LT, True
        Case 1
            Dialogue " Grupo x Grupo", GetPlayerName(CC) & " Desafiou você para uma lutar " & a & ", Aceitar ?", DIALOGUE_TYPE_LT, True
        Case 2
            Dialogue " Equipe x Equipe", GetPlayerName(CC) & " Desafiou você para uma lutar " & a & ", Aceitar ?", DIALOGUE_TYPE_LT, True
        End Select
    Case 2
        SendLutarComando 4, 0, 0, 0, vbNullString
    End Select

End Sub

Private Sub HandleArenas(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Arena As Long, status As Long

    Set buffer = New clsBuffer

    buffer.WriteBytes Data()

    Arena = buffer.ReadLong
    status = buffer.ReadLong

    If Arena = 0 Then Exit Sub

    Player(Index).Arena(Arena) = status
    
    'If Player(Index).Arena(Arena) = 1 Then
    '    frmMain.lblArena(3).Caption = "teste"
   ' End If

End Sub

Private Sub HandleAprenderSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Command As Long, i As Integer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    Command = buffer.ReadByte
    Player(Index).LearnSpell(1) = buffer.ReadInteger
    Player(Index).LearnSpell(2) = buffer.ReadInteger

    If Command = 0 Then
        frmMain.PicHabilidade.Visible = True
        frmMain.lblDescHab(0).Caption = "O Pokémon " & Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " quér aprender a habilidade " & Trim$(Spell(Player(Index).LearnSpell(2)).Name) & " mas já possui 4 ataques, qual deseja substituir?"
        For i = 1 To 4
            If GetPlayerEquipmentPokeInfoSpell(Index, weapon, i) > 0 Then
                frmMain.lblDescHab(i).Caption = Trim$(Spell(GetPlayerEquipmentPokeInfoSpell(Index, weapon, i)).Name)
            Else
                frmMain.lblDescHab(i).Caption = vbNullString
            End If
        Next

        frmMain.PicHabilidade.top = (frmMain.ScaleHeight / 2) - (frmMain.PicHabilidade.Height / 2)
        frmMain.PicHabilidade.Left = (frmMain.ScaleWidth / 2) - (frmMain.PicHabilidade.Width / 2)

    Else
        frmMain.PicHabilidade.Visible = False
        frmMain.lblDescHab(0).Caption = vbNullString
        For i = 1 To 4
            frmMain.lblDescHab(i).Caption = vbNullString
        Next
    End If

    Set buffer = Nothing

End Sub

Sub HandleNoticia(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String, i As Long
    Dim color As Byte


    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NoticiaServ(1) = vbNullString Then
        NotX = 0
    End If

    frmMain.tmrNoticia.Enabled = True

    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong

    For i = 1 To MAX_NOTICIAS
        If NoticiaServ(i) = vbNullString Then
            NoticiaServ(i) = Msg
            Exit For
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNoticia", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAttOrg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, P As Long, e As Long, U As String, status As String
    Dim i As Long, Membros As Boolean

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If Player(Index).ORG > 0 Then
        Organization(Player(Index).ORG).Exp = buffer.ReadLong
        Organization(Player(Index).ORG).Level = buffer.ReadLong
        P = buffer.ReadLong
        Membros = buffer.ReadByte
        MaxExpOrg = P

        If Membros = True Then
            For i = 1 To MAX_ORG_MEMBERS
                Organization(Player(Index).ORG).OrgMember(i).User_Name = buffer.ReadString
                Organization(Player(Index).ORG).OrgMember(i).Online = buffer.ReadByte
                Organization(Player(Index).ORG).OrgMember(i).Used = buffer.ReadByte
            Next

            'Limpar
            frmMain.OrgMembers.Clear

            'Adicionar
            For i = 1 To MAX_ORG_MEMBERS

                If Organization(Player(Index).ORG).OrgMember(i).Online = True Then status = "Online"
                If Organization(Player(Index).ORG).OrgMember(i).Online = False Then status = "Offline"

                If Not Organization(Player(Index).ORG).OrgMember(i).Used = 0 Then
                    If i = 1 Then
                        frmMain.OrgMembers.AddItem "Líder:" & Trim$(Organization(Player(Index).ORG).OrgMember(i).User_Name) & " - " & status
                    Else
                        frmMain.OrgMembers.AddItem i - 1 & ": " & Trim$(Organization(Player(Index).ORG).OrgMember(i).User_Name) & " - " & status
                    End If
                End If
            Next
        End If

        frmMain.PicExp.Width = ((Organization(Player(Index).ORG).Exp) / ORG) / (P / ORG) * ORG

        'Atualizar
        BltOrganização

        Select Case Player(Index).ORG
        Case 1
            U = "Equipe Rocket"
        Case 2
            U = "Team Magma"
        Case 3
            U = "Team Aqua"
        Case Else
            Exit Sub
        End Select
    End If
    Set buffer = Nothing
End Sub

Sub HandleOrgShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, P As Long, e As Long, U As String
    Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    For i = 1 To MAX_ORG_SHOP
        OrgShop(i).Item = buffer.ReadLong
        OrgShop(i).Quantia = buffer.ReadLong
        OrgShop(i).Valor = buffer.ReadLong
        OrgShop(i).Level = buffer.ReadLong
    Next
    Set buffer = Nothing

End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, targetType As Long, target As Long, Message As String, colour As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    target = buffer.ReadLong
    targetType = buffer.ReadLong
    Message = buffer.ReadString
    colour = buffer.ReadLong

    AddChatBubble target, targetType, Message, colour
    Set buffer = Nothing
End Sub

Private Sub HandleVipPlayerInfo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, targetType As Long, target As Long, Message As String, colour As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Player(MyIndex).VipPoints = buffer.ReadLong
    Set buffer = Nothing

    frmMain.lblVip(7).Caption = Player(MyIndex).VipPoints
    'Vip Points
    If Player(MyIndex).VipPoints <= 1500 Then
        frmMain.PicVipBar.Width = (Player(MyIndex).VipPoints / 377) / (1500 / 377) * 377
    Else
        frmMain.PicVipBar.Width = 377
    End If
End Sub

Private Sub HandleAparencia(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long
    Dim Sex As String

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Player(i).Sex = buffer.ReadByte

    'Modelos
    Player(i).HairModel = buffer.ReadInteger
    Player(i).ClothModel = buffer.ReadInteger
    Player(i).LegsModel = buffer.ReadInteger

    'Cor
    Player(i).HairColor = buffer.ReadByte
    Player(i).ClothColor = buffer.ReadByte
    Player(i).LegsColor = buffer.ReadByte

    'Numero
    Player(i).HairNum = buffer.ReadInteger
    Player(i).ClothNum = buffer.ReadInteger
    Player(i).LegsNum = buffer.ReadInteger
    Set buffer = Nothing

End Sub

Private Sub HandlePlayerRun(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Run As Byte, i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Run = buffer.ReadByte
    If Run = 1 Then Player(i).Running = True
    If Run = 0 Then Player(i).Running = False
    Set buffer = Nothing
End Sub

Private Sub HandleComandoGym(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Comando As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Comando = buffer.ReadByte
    Set buffer = Nothing

    Select Case Comando
    Case 1    'Brock Invite
        ChatGym = 1    'Brock
        ChatGymStep = 0    'Primeira fala
        frmMain.lblGym(0).Caption = "Então, você está aqui eu sou Brock o líder do ginásio de Pewter, minha força de vontade é uma Rocha Sólida é evidente, mesmo meus pokémons são pura rocha, a verdadeira força de vontade! é isso mesmo... Os meus pokémons são todos do tipo Pedra! HAHAHA! Você vai me desafiar sabendo que você vai perder?"
        frmMain.lblGym(1).Caption = "Sim!"
        frmMain.lblGym(2).Caption = "Cancelar"
        frmMain.lblGym(3).Caption = "Brock"
        frmMain.PicBlank(1).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\gymleader\1.jpg")
        frmMain.PicBlank(0).Visible = True
        frmMain.PicBlank(0).top = (frmMain.ScaleHeight / 2) - (frmMain.PicBlank(0).Height / 2)
        frmMain.PicBlank(0).Left = (frmMain.ScaleWidth / 2) - (frmMain.PicBlank(0).Width / 2)
    End Select
End Sub

Private Sub HandleContagem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Tempo As Integer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ContagemGym = buffer.ReadInteger
    ContagemTick = 1000 + GetTickCount
    Set buffer = Nothing

End Sub
