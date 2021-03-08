Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal Index As Long)
    If Not IsPlaying(Index) Then
        Call JoinGame(Index)
        Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal Index As Long)
    Dim i As Long, X As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    
    ' send the login ok
    SendLoginOk Index
    
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendAnimations(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendResources(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendPlayerSpells(Index)
    Call SendHotbar(Index)
    Call SendPokemonS(Index)
    Call SendQuests(Index)
    Call SendRankLevelTo(Index)
    Call SendOrgShop(Index)
    Call SendEXP(Index)
    Call SendPlayerPokedex(Index)
    Call SendAttLeilaoTo(Index)
    Call ChecarQntiadePokemons(Index)
    Call SendSurfInit(Index)
    Call SendPlayerTeleport(Index)
    'Call SendAparencia(Index)
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
    Next
    
    ' Teleportar jogador para local salvo
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    Select Case GetPlayerMap(Index)
    Case 8
        PlayerWarp Index, 7, 12, 9
        PlayerMsg Index, "[Brock]: Você deslogou na batalha e foi removido da arena!", White
    End Select
    
    ' Msg Administração entrou!
    If GetPlayerAccess(Index) >= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Mensagem de Bem vindo
    Call SendWelcome(Index)

    ' Checar Vip
    CheckVipDays Index, True

    'Enviar Permição para Evoluir
    If Player(Index).EvolPermition = 1 Then
        SendPokeEvolution Index, 0
    End If

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next
    
    'Evoluir Pokémon - Pedra
    If Player(Index).EvolTimerStone > 0 Then
        Player(Index).EvolTimerStone = 9000 + GetTickCount
        SendPokeEvolution Index, 1
    End If
    
    If Player(Index).LearnSpell(1) = 1 Then Player(Index).LearnSpell(1) = 0
    
    'Atualizar Lista de Membros ON!
    If Player(Index).ORG > 0 Then
        For i = 1 To MAX_ORG_MEMBERS
            If Trim$(Organization(Player(Index).ORG).OrgMember(i).User_Name) = Trim$(GetPlayerName(Index)) Then
                Organization(Player(Index).ORG).OrgMember(i).Online = True
                SendOrganizaçãoToOrg Player(Index).ORG
            End If
        Next
    End If
    
    ' Send the flag so they know they can start doing stuff
    SendInGame Index
    VerificarPendencias Index
End Sub

Public Sub VerificarPendencias(ByVal Index As Integer)
Dim i As Long, X As Long

For i = 1 To MAX_LEILAO
        If Pendencia(i).Vendedor = GetPlayerName(Index) Then
            
            'Dar o item
            If Pendencia(i).Poke.Pokemon > 0 Then
            
            
            
            DirectBankItemPokemon Index, Pendencia(i).ItemNum, _
            Pendencia(i).Poke.Pokemon, Pendencia(i).Poke.Pokeball, _
            Pendencia(i).Poke.Level, Pendencia(i).Poke.EXP, _
            Pendencia(i).Poke.Vital(1), Pendencia(i).Poke.Vital(2), _
            Pendencia(i).Poke.MaxVital(1), Pendencia(i).Poke.MaxVital(2), _
            Pendencia(i).Poke.Stat(1), Pendencia(i).Poke.Stat(4), _
            Pendencia(i).Poke.Stat(2), Pendencia(i).Poke.Stat(3), _
            Pendencia(i).Poke.Stat(5), Pendencia(i).Poke.Spells(1), _
            Pendencia(i).Poke.Spells(2), Pendencia(i).Poke.Spells(3), _
            Pendencia(i).Poke.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Pendencia(i).Poke.Felicidade, Pendencia(i).Poke.Sexo, Pendencia(i).Poke.Shiny, _
            Pendencia(i).Poke.Berry(1), Pendencia(i).Poke.Berry(2), Pendencia(i).Poke.Berry(3), Pendencia(i).Poke.Berry(4), Pendencia(i).Poke.Berry(5)
            
            'Enviar Msg
            If Pendencia(i).Tipo > 0 Then
                PlayerMsg Index, "O pokémon " & Trim$(Pokemon(Pendencia(i).Poke.Pokemon).Name) & "foi vendido por " & Pendencia(i).Price & " " & Trim$(Item(Pendencia(i).Tipo).Name), BrightGreen
            Else
                PlayerMsg Index, "O pokémon " & Trim$(Pokemon(Pendencia(i).Poke.Pokemon).Name) & " não foi vendido e foi enviado para o computador!", Yellow
            End If
            
            Else
            DirectBankItemPokemon Index, Pendencia(i).ItemNum, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
            'GiveInvItem Index, Pendencia(I).ItemNum, Pendencia(I).Price, False, _
            Pendencia(I).Poke.Pokemon, Pendencia(I).Poke.Pokeball, _
            Pendencia(I).Poke.Level, Pendencia(I).Poke.EXP, _
            Pendencia(I).Poke.Vital(1), Pendencia(I).Poke.Vital(2), _
            Pendencia(I).Poke.MaxVital(1), Pendencia(I).Poke.MaxVital(2), _
            Pendencia(I).Poke.Stat(1), Pendencia(I).Poke.Stat(4), _
            Pendencia(I).Poke.Stat(2), Pendencia(I).Poke.Stat(3), _
            Pendencia(I).Poke.Stat(5), Pendencia(I).Poke.Spells(1), _
            Pendencia(I).Poke.Spells(2), Pendencia(I).Poke.Spells(3), _
            Pendencia(I).Poke.Spells(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Pendencia(I).Poke.Felicidade, Pendencia(I).Poke.Sexo, Pendencia(I).Poke.Shiny, _
            Pendencia(I).Poke.Berry(1), Pendencia(I).Poke.Berry(2), Pendencia(I).Poke.Berry(3), Pendencia(I).Poke.Berry(4), Pendencia(I).Poke.Berry(5)
            
            'Enviar Msg
            If Pendencia(i).Tipo > 0 Then
                PlayerMsg Index, "Venda Concluida!", White
            Else
                PlayerMsg Index, "O Item " & Trim$(Item(Pendencia(i).ItemNum).Name) & " não foi vendido!", BrightRed
            End If
            
            End If
            
            'Limpar a pendencia em questão...
            Pendencia(i).Vendedor = vbNullString
            Pendencia(i).ItemNum = 0
            Pendencia(i).Price = 0
            Pendencia(i).Tipo = 0
            Pendencia(i).Poke.Pokemon = 0
            Pendencia(i).Poke.Pokeball = 0
            Pendencia(i).Poke.Level = 0
            Pendencia(i).Poke.EXP = 0
            Pendencia(i).Poke.Felicidade = 0
            Pendencia(i).Poke.Sexo = 0
            Pendencia(i).Poke.Shiny = 0
            
            For X = 1 To Vitals.Vital_Count - 1
                Pendencia(i).Poke.Vital(X) = 0
                Pendencia(i).Poke.MaxVital(X) = 0
            Next
            
            For X = 1 To Stats.Stat_Count - 1
                Pendencia(i).Poke.Stat(X) = 0
            Next
            
            For X = 1 To MAX_POKE_SPELL
                Pendencia(i).Poke.Spells(X) = 0
            Next
            
            For X = 1 To MAX_BERRYS
                Pendencia(i).Poke.Berry(X) = 0
            Next
            
            SavePendencia i
        End If
    Next i
End Sub

Sub LeftGame(ByVal Index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    Select Case TempPlayer(Index).InBattleGym
    Case 1
        MapNpc(7).Npc(1).InBattle = False
    End Select
    
    If Player(Index).ORG > 0 Then
        For i = 1 To MAX_ORG_MEMBERS
            If Trim$(Organization(Player(Index).ORG).OrgMember(i).User_Name) = Trim$(GetPlayerName(Index)) Then
                Organization(Player(Index).ORG).OrgMember(i).Online = False
                SendOrganizaçãoToOrg Player(Index).ORG
                Exit For
            End If
        Next
    End If
    
     If TempPlayer(Index).Conversando > 0 Then
            SendChat 2, TempPlayer(Index).Conversando, Index, vbNullString
            TempPlayer(Index).Conversando = 0
        End If
        
        If GetPlayerMap(Index) > 0 Then
            If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_ARENA Then
                PlayerUnequipItem Index, weapon
                
                SendLeaveMap Index, GetPlayerMap(Index)
                Player(Index).Map = Player(Index).MyMap(1)
                Player(Index).X = Player(Index).MyMap(2)
                Player(Index).Y = Player(Index).MyMap(3)
                If TempPlayer(Index).Lutando > 0 Then
                    Dim p As Long
                    p = TempPlayer(Index).Lutando
                    If IsPlaying(p) Then
                        PlayerMsg p, "O jogador " & Trim(GetPlayerName(Index)) & " desistiu da batalha!", Red
                        PlayerWarp p, Player(p).MyMap(1), Player(p).MyMap(2), Player(p).MyMap(3)
                        Player(Index).Derrotas = Player(Index).Derrotas + 1
                        Player(p).Vitorias = Player(p).Vitorias + 1
                    End If
                End If
            End If
        End If
        
        If TempPlayer(Index).Lutando > 0 Then
         Select Case TempPlayer(Index).LutandoT
          Case 1
            SendArenaStatus TempPlayer(Index).LutandoA, 0
            PlayerWarp Index, 350, 11, 8
            SetPlayerDir Index, DIR_DOWN
            SendLutarComando 2, 0, TempPlayer(Index).Lutando, 0, 0, 0
            TempPlayer(Index).Lutando = 0
            TempPlayer(Index).LutandoA = 0
            TempPlayer(Index).LutandoT = 0
          Case 2
            SendArenaStatus TempPlayer(Index).LutandoA, 0
            Party(TempPlayer(Index).inParty).PT = Party(TempPlayer(Index).inParty).PT - 1
            PlayerWarp Index, 350, 11, 8
          Case 3
           ' org
           End Select
        End If
    
    If TempPlayer(Index).InGame Then
    
        TempPlayer(Index).InGame = False
        
          ' Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = Index Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
        ' cancel any trade they're in'123
        If TempPlayer(Index).InTrade > 0 Then
            tradeTarget = TempPlayer(Index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave Index
        
         ' clear target
        For i = 1 To Player_HighIndex
            ' Prevent subscript out range
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(Index) Then
                ' clear players target
                If TempPlayer(i).targetType = TARGET_TYPE_PLAYER And TempPlayer(i).target = Index Then
                    TempPlayer(i).target = 0
                    TempPlayer(i).targetType = TARGET_TYPE_NONE
                    SendTarget i
                End If
            End If
        Next

        ' save and clear data.
        Call SavePlayer(Index)
        Call SaveBank(Index)
        Call ClearBank(Index)

        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(Index)
End Sub

Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(Index, Armor)
    Helm = GetPlayerEquipment(Index, Helmet)
    GetPlayerProtection = (GetPlayerStat(Index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(Index, weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = ((GetPlayerEquipmentPokeInfoStat(Index, weapon, Stats.Agility) + GetPlayerEquipmentBerry(Index, weapon, 4)) \ 2) + (GetPlayerEquipmentPokeInfoLevel(Index, weapon) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(Index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Endurance) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If X > Map(MapNum).MaxX Then X = Map(MapNum).MaxX
    If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    
    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(Index) Then
        SendPlayerXYToMap Index
    End If
    
    ' clear target
    TempPlayer(Index).target = 0
    TempPlayer(Index).targetType = TARGET_TYPE_NONE
    SendTarget Index
    
    ' clear target
    For i = 1 To Player_HighIndex
        ' Prevent subscript out range
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(Index) Then
            If TempPlayer(i).targetType = TARGET_TYPE_PLAYER And TempPlayer(i).target = Index Then
                TempPlayer(i).target = 0
                TempPlayer(i).targetType = TARGET_TYPE_NONE
                SendTarget i
            End If
        End If
    Next

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If OldMap <> MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)
    TempPlayer(Index).StunDuration = 0.6
    TempPlayer(Index).StunTimer = GetTickCount
    SendStunned Index
    
    ' send player's equipment to new map
    SendMapEquipment Index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(MapNum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    SendMapEquipmentTo i, Index
                    SendSurfInit i, Index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(OldMap).Npc(i).Num > 0 Then
                If Npc(MapNpc(OldMap).Npc(i).Num).Pokemon > 0 Then
                    MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetPokemonMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP, MapNpc(OldMap).Npc(i).Level)
                Else
                    MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP)
                End If
            End If
        Next

    End If
    
    'Trocar Status Rock Tunel = 0
    Player(Index).PokeLight = False
    
    ' Check if the player completed a quest
    ChecarTarefasAtuais Index, QUEST_TYPE_GOTOMAP, GetPlayerMap(Index)
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    TempPlayer(Index).GettingMap = YES
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong MapNum
    Buffer.WriteLong Map(MapNum).Revision
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
    
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(MapNum).Npc(i).Num > 0 Then
            If MapNpc(MapNum).Npc(i).Desmaiado = True And Map(MapNum).Npc(i) > 0 Then
                SendNpcDesmaiado MapNum, i, False, Index
                SendMapNpcVitals MapNum, i, Index
            End If
        End If
    Next
    
End Sub

Sub PlayerMoveFly(ByVal Index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, MapNum As Long
    Dim X As Long, Y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, Amount As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If
    
    'Verificar se está voando com Pokémon!
    If GetPlayerEquipment(Index, weapon) = 0 Then
        If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) = 0 Then
            SetPlayerFlying Index, 0
            SendPlayerData Index
            Exit Sub
        End If
    Else
        If GetPlayerEquipmentNgt(Index, weapon, 2) > 0 Then
            SetPlayerFlying Index, 0
            SendPlayerData Index
            Exit Sub
        End If
    End If
    
    Call SetPlayerDir(Index, Dir)
    Moved = NO
    MapNum = GetPlayerMap(Index)
    
        Select Case Dir
        Case DIR_UP
            Call SetPlayerY(Index, GetPlayerY(Index) - 1)
            SendPlayerMove Index, movement, sendToSelf
            Moved = YES
        Case DIR_DOWN
            Call SetPlayerY(Index, GetPlayerY(Index) + 1)
            SendPlayerMove Index, movement, sendToSelf
            Moved = YES
        Case DIR_LEFT
            Call SetPlayerX(Index, GetPlayerX(Index) - 1)
            SendPlayerMove Index, movement, sendToSelf
            Moved = YES
        Case DIR_RIGHT
            Call SetPlayerX(Index, GetPlayerX(Index) + 1)
            SendPlayerMove Index, movement, sendToSelf
            Moved = YES
        End Select
        
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, MapNum As Long
    Dim X As Long, Y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte, NgtNum As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, Amount As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If
    
    'Congelado
    If GetPlayerEquipment(Index, weapon) > 0 Then
        If GetPlayerEquipmentNgt(Index, weapon, 2) > 0 Then
            Exit Sub
        End If
    End If
    
    Call SetPlayerDir(Index, Dir)
    Moved = NO
    MapNum = GetPlayerMap(Index)
    
    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_UP + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_RESOURCE Or CheckResourceStatCut(Index, GetPlayerX(Index), GetPlayerY(Index) - 1) = True Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    If GetPlayerEquipment(Index, weapon) = 0 Or GetPlayerEquipment(Index, weapon) = 247 Then
                    NewMapY = Map(Map(GetPlayerMap(Index)).Up).MaxY
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).target = 0
                    TempPlayer(Index).targetType = TARGET_TYPE_NONE
                    SendTarget Index
                    End If
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < Map(MapNum).MaxY Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_RESOURCE Or CheckResourceStatCut(Index, GetPlayerX(Index), GetPlayerY(Index) + 1) = True Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Down > 0 Then
                    If GetPlayerEquipment(Index, weapon) = 0 Or GetPlayerEquipment(Index, weapon) = 247 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).target = 0
                    TempPlayer(Index).targetType = TARGET_TYPE_NONE
                    SendTarget Index
                    End If
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_RESOURCE Or CheckResourceStatCut(Index, GetPlayerX(Index) - 1, GetPlayerY(Index)) = True Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    If GetPlayerEquipment(Index, weapon) = 0 Or GetPlayerEquipment(Index, weapon) = 247 Then
                    NewMapX = Map(Map(GetPlayerMap(Index)).Left).MaxX
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, NewMapX, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).target = 0
                    TempPlayer(Index).targetType = TARGET_TYPE_NONE
                    SendTarget Index
                    End If
                End If
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < Map(MapNum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_RESOURCE Or CheckResourceStatCut(Index, GetPlayerX(Index) + 1, GetPlayerY(Index)) = True Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                SendPlayerMove Index, movement, sendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Right > 0 Then
                If GetPlayerEquipment(Index, weapon) = 0 Or GetPlayerEquipment(Index, weapon) = 247 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).target = 0
                    TempPlayer(Index).targetType = TARGET_TYPE_NONE
                    SendTarget Index
                    End If
                End If
            End If
    End Select
    
    If GetPlayerFlying(Index) Then Exit Sub
    
    With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            MapNum = .Data1
            X = .Data2
            Y = .Data3
            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) = 0 Then
            Call PlayerWarp(Index, MapNum, X, Y)
            Moved = YES
            End If
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TILE_TYPE_DOOR Then
            MapNum = .Data1
            X = .Data2
            Y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
            Call PlayerWarp(Index, MapNum, X, Y)
            Moved = YES
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            X = .Data1
            Y = .Data2
    
            If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                SendMapKey Index, X, Y, 1
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            X = .Data1
            If X > 0 Then ' shop exists?
                If Len(Trim$(Shop(X).Name)) > 0 Then ' name exists?
                    SendOpenShop Index, X
                    TempPlayer(Index).InShop = X ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank Index
            TempPlayer(Index).InBank = True
            Moved = YES
        End If
        
        ' PokeCentro
        If .Type = TILE_TYPE_HEAL Then
            Dim i As Long
            
            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                Call PlayerUnequipItem(Index, weapon)
                PlayerMsg Index, "Pokémons só podem ser curados dentro da pokéball!", Yellow
                Exit Sub
            End If
            
            'Mexer futuramente
            'SendAnimation MapNum, 1, 12, 4, 0, Index
            SetPlayerVital Index, 1, GetPlayerMaxVital(Index, HP)
            SetPlayerVital Index, 2, GetPlayerMaxVital(Index, MP)
            TempPlayer(Index).StunDuration = 1
            TempPlayer(Index).StunTimer = GetTickCount
            SendStunned Index
            
            Dim PokeQntiaCurado As Byte
            
            For i = 1 To MAX_INV
            If GetPlayerInvItemPokeInfoPokemon(Index, i) > 0 Then
                If GetPlayerInvItemPokeInfoVital(Index, i, 1) < GetPlayerInvItemPokeInfoMaxVital(Index, i, 1) Or GetPlayerInvItemPokeInfoVital(Index, i, 2) < GetPlayerInvItemPokeInfoMaxVital(Index, i, 2) Then
                    SetPlayerInvItemPokeInfoVital Index, i, GetPlayerInvItemPokeInfoMaxVital(Index, i, 1), 1
                    SetPlayerInvItemPokeInfoVital Index, i, GetPlayerInvItemPokeInfoMaxVital(Index, i, 2), 2
                    PokeQntiaCurado = PokeQntiaCurado + 1
                End If
                
                'limpar status negativos
                For NgtNum = 1 To MAX_NEGATIVES
                    If GetPlayerInvItemNgt(Index, i, NgtNum) > 0 Then
                        SetPlayerInvItemNgt Index, i, NgtNum, 0
                    End If
                Next
                
            End If
            Next
            
            If PokeQntiaCurado > 0 Then
            PlayerMsg Index, "Todos seus pokémons foram curado com sucesso!", Yellow
            End If
            
            Call SendInventory(Index)
            Call SendVital(Index, HP)
            Call SendVital(Index, MP)
            
            If GetPlayerMap(Index) = 99 Then
            SendAnimation GetPlayerMap(Index), 1, 10, 4
            SendAnimation GetPlayerMap(Index), 11, 12, 4
            End If
            
            ' send vitals to party if in one
            If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            Moved = YES
        End If
        
       ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            Select Case .Data1
                Case DIR_UP
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_BLOCKED Then Exit Sub
                Case DIR_LEFT
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_BLOCKED Then Exit Sub
                Case DIR_DOWN
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_BLOCKED Then Exit Sub
                Case DIR_RIGHT
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_BLOCKED Then Exit Sub
            End Select
            
            ForcePlayerMove Index, MOVING_WALKING, .Data1
            Moved = YES
        End If
        
        'Script Tile
        If .Type = TILE_TYPE_SCRIPT Then
            Call ScriptedTile(Index, .Data1)
            Moved = YES
        End If
        
        'Water Tile
        If .Type = TILE_TYPE_WATER Then
            
            If GetPlayerEquipment(Index, weapon) > 0 Then
                BlockPlayer Index
                Exit Sub
            End If
            
            If Player(Index).InSurf = 0 Then 'Entrando
                Player(Index).InSurf = 3
                SendSurfInit Index, Index
                TempPlayer(Index).SurfSlideTo = .Data1
                BlockPlayer Index
            Else
                If Player(Index).InSurf = 1 And .Data1 = GetPlayerDir(Index) Then
                Else
                    Select Case .Data1
                    Case DIR_UP
                        ForcePlayerMove Index, MOVING_WALKING, DIR_DOWN
                        ForcePlayerMove Index, MOVING_WALKING, DIR_DOWN
                    Case DIR_DOWN
                        ForcePlayerMove Index, MOVING_WALKING, DIR_UP
                        ForcePlayerMove Index, MOVING_WALKING, DIR_UP
                    Case DIR_LEFT
                        ForcePlayerMove Index, MOVING_WALKING, DIR_RIGHT
                        ForcePlayerMove Index, MOVING_WALKING, DIR_RIGHT
                    Case DIR_RIGHT
                        ForcePlayerMove Index, MOVING_WALKING, DIR_LEFT
                        ForcePlayerMove Index, MOVING_WALKING, DIR_LEFT
                    End Select
                    
                    Player(Index).InSurf = 0
                        SendSurfInit Index
                    End If
            End If
        End If
        
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
    End If

End Sub

Sub ForcePlayerMove(ByVal Index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(Index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(Index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove Index, Direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal Index As Long)
    Dim Slot As Long
    Dim ItemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(Index, i)

        If ItemNum > 0 Then

            Select Case i
                Case Equipment.weapon

                    If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment Index, 0, i
                Case Equipment.Armor

                    If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment Index, 0, i
                Case Equipment.Helmet

                    If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment Index, 0, i
                Case Equipment.Shield

                    If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment Index, 0, i
            End Select

        Else
            SetPlayerEquipment Index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                'PlayerMsg Index, "Slot: " & i, BrightRed
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            'PlayerMsg Index, GetPlayerInvItemNum(Index, i), BrightRed
            'PlayerMsg Index, "Slot: " & i, BrightRed
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long, Optional ByVal Pokemon As Long, Optional ByVal Pokeball As Long, Optional ByVal Level As Long, Optional ByVal EXP As Long, Optional ByVal VitalHP As Long, Optional ByVal VitalMP As Long, Optional ByVal MaxVitalHp As Long, Optional ByVal MaxVitalMp As Long, Optional ByVal StatStr As Long, Optional ByVal StatAgi As Long, Optional ByVal StatEnd As Long, Optional ByVal StatInt As Long, Optional ByVal StatWill As Long, Optional ByVal Spell1 As Long, Optional ByVal Spell2 As Long, Optional ByVal Spell3 As Long, Optional ByVal Spell4 As Long) As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Function
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
    
        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function FindOpenBankSlotPokemon(ByVal Index As Long) As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Function

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, i) = 0 Then
            If GetPlayerBankItemPokemon(Index, i) = 0 Then
                FindOpenBankSlotPokemon = i
                Exit Function
            End If
        End If
    Next i

End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, Optional ByVal QuestMessage As Boolean = True) As Boolean
    Dim i As Long
    Dim n As Long
    Dim X As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
            
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                
                'PokeInfo
                Call SetPlayerInvItemPokeInfoPokemon(Index, i, 0)
                Call SetPlayerInvItemPokeInfoPokeball(Index, i, 0)
                Call SetPlayerInvItemPokeInfoLevel(Index, i, 0)
                Call SetPlayerInvItemPokeInfoExp(Index, i, 0)
                
                'Max/Vitals
                For X = 1 To 2
                    Call SetPlayerInvItemPokeInfoVital(Index, i, 0, X)
                    Call SetPlayerInvItemPokeInfoMaxVital(Index, i, 0, X)
                Next
                
                'Stats Pokemon
                For X = 1 To 5
                    Call SetPlayerInvItemPokeInfoStat(Index, i, X, 0)
                Next
                
                'Spells Pokemon
                For X = 1 To MAX_POKE_SPELL
                    Call SetPlayerInvItemPokeInfoSpell(Index, i, 0, X)
                Next
                
                'Negative
                For X = 1 To MAX_NEGATIVES
                    Call SetPlayerInvItemNgt(Index, i, 0, X)
                Next
                
                'Berry
                For X = 1 To MAX_BERRYS
                    Call SetPlayerInvItemBerry(Index, i, X, 0)
                Next
                
                'Felicidade,Shiny e Sexo
                Call SetPlayerInvItemFelicidade(Index, i, 0)
                Call SetPlayerInvItemShiny(Index, i, 0)
                Call SetPlayerInvItemSexo(Index, i, 0)
                 
                ' Check if the player completed a quest
                ChecarTarefasAtuais Index, QUEST_TYPE_COLLECTITEMS, ItemNum
                
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)

                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal Index As Long, ByVal invslot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim X As Long
    Dim ItemNum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or invslot <= 0 Or invslot > MAX_ITEMS Then
        Exit Function
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, invslot)

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(Index, invslot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(Index, invslot, GetPlayerInvItemValue(Index, invslot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
    
        'Item,Valor
        Call SetPlayerInvItemNum(Index, invslot, 0)
        Call SetPlayerInvItemValue(Index, invslot, 0)
        
        'PokeInfo
        Call SetPlayerInvItemPokeInfoPokemon(Index, invslot, 0)
        Call SetPlayerInvItemPokeInfoPokeball(Index, invslot, 0)
        Call SetPlayerInvItemPokeInfoLevel(Index, invslot, 0)
        Call SetPlayerInvItemPokeInfoExp(Index, invslot, 0)
        
        'Max/Vitals
        For X = 1 To 2
            Call SetPlayerInvItemPokeInfoVital(Index, invslot, 0, X)
            Call SetPlayerInvItemPokeInfoMaxVital(Index, invslot, 0, X)
        Next
        
        'Stats Pokemon
        For X = 1 To 5
            Call SetPlayerInvItemPokeInfoStat(Index, invslot, X, 0)
        Next
        
        'Spells Pokemon
        For X = 1 To MAX_POKE_SPELL
            Call SetPlayerInvItemPokeInfoSpell(Index, invslot, 0, X)
        Next
        
        'Negative
        For X = 1 To MAX_NEGATIVES
            Call SetPlayerInvItemNgt(Index, invslot, 0, X)
        Next
        
        'Berrys
        For X = 1 To MAX_BERRYS
            Call SetPlayerInvItemBerry(Index, invslot, X, 0)
        Next
        
        'Felicidade,Shiny e Sexo
        Call SetPlayerInvItemFelicidade(Index, invslot, 0)
        Call SetPlayerInvItemShiny(Index, invslot, 0)
        Call SetPlayerInvItemSexo(Index, invslot, 0)
        
        ' Check if the player completed a quest
        ChecarTarefasAtuais Index, QUEST_TYPE_COLLECTITEMS, ItemNum
        Exit Function
    End If
End Function

Function GiveInvItem(ByVal Index As Long, ByVal ItemNum As Long, _
ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True, _
Optional ByVal Pokemon As Long, Optional ByVal Pokeball As Long, _
Optional ByVal Level As Long, Optional ByVal EXP As Long, _
Optional ByVal VitalHP As Long, Optional ByVal VitalMP As Long, _
Optional ByVal MaxVitalHp As Long, Optional ByVal MaxVitalMp As Long, _
Optional ByVal StatStr As Long, Optional ByVal StatAgi As Long, _
Optional ByVal StatEnd As Long, Optional ByVal StatInt As Long, _
Optional ByVal StatWill As Long, Optional ByVal Spell1 As Long, _
Optional ByVal Spell2 As Long, Optional ByVal Spell3 As Long, _
Optional ByVal Spell4 As Long, Optional ByVal Ngt1 As Long, _
Optional ByVal Ngt2 As Long, Optional ByVal Ngt3 As Long, _
Optional ByVal Ngt4 As Long, Optional ByVal Ngt5 As Long, _
Optional ByVal Ngt6 As Long, Optional ByVal Ngt7 As Long, _
Optional ByVal Ngt8 As Long, Optional ByVal Ngt9 As Long, _
Optional ByVal Ngt10 As Long, Optional ByVal Ngt11 As Long, _
Optional ByVal Felicidade As Long, Optional ByVal Sexo As Byte, Optional ByVal Shiny As Long, _
Optional ByVal Bry1 As Long, Optional ByVal Bry2 As Long, Optional ByVal Bry3 As Long, Optional ByVal Bry4 As Long, _
Optional ByVal Bry5 As Long) As Boolean

'Função pequena...

    Dim i As Long
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(Index, ItemNum)

    ' Check to see if inventory is full
    If i <> 0 Then
    
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        
        'PokeInfo
        Call SetPlayerInvItemPokeInfoPokemon(Index, i, Pokemon)
        Call SetPlayerInvItemPokeInfoPokeball(Index, i, Pokeball)
        Call SetPlayerInvItemPokeInfoLevel(Index, i, Level)
        Call SetPlayerInvItemPokeInfoExp(Index, i, EXP)
        
        'Max/Vitals
        Call SetPlayerInvItemPokeInfoVital(Index, i, VitalHP, 1)
        Call SetPlayerInvItemPokeInfoVital(Index, i, VitalMP, 2)
        Call SetPlayerInvItemPokeInfoMaxVital(Index, i, MaxVitalHp, 1)
        Call SetPlayerInvItemPokeInfoMaxVital(Index, i, MaxVitalMp, 2)
        
        'Stats Pokemon
        Call SetPlayerInvItemPokeInfoStat(Index, i, 1, StatStr)
        Call SetPlayerInvItemPokeInfoStat(Index, i, 2, StatEnd)
        Call SetPlayerInvItemPokeInfoStat(Index, i, 3, StatInt)
        Call SetPlayerInvItemPokeInfoStat(Index, i, 4, StatAgi)
        Call SetPlayerInvItemPokeInfoStat(Index, i, 5, StatWill)
        
        'Spells Pokemon
        Call SetPlayerInvItemPokeInfoSpell(Index, i, Spell1, 1)
        Call SetPlayerInvItemPokeInfoSpell(Index, i, Spell2, 2)
        Call SetPlayerInvItemPokeInfoSpell(Index, i, Spell3, 3)
        Call SetPlayerInvItemPokeInfoSpell(Index, i, Spell4, 4)
        
        'Negative
        Call SetPlayerInvItemNgt(Index, i, 1, Ngt1)
        Call SetPlayerInvItemNgt(Index, i, 2, Ngt2)
        Call SetPlayerInvItemNgt(Index, i, 3, Ngt3)
        Call SetPlayerInvItemNgt(Index, i, 4, Ngt4)
        Call SetPlayerInvItemNgt(Index, i, 5, Ngt5)
        Call SetPlayerInvItemNgt(Index, i, 6, Ngt6)
        Call SetPlayerInvItemNgt(Index, i, 7, Ngt7)
        Call SetPlayerInvItemNgt(Index, i, 8, Ngt8)
        Call SetPlayerInvItemNgt(Index, i, 9, Ngt9)
        Call SetPlayerInvItemNgt(Index, i, 10, Ngt10)
        Call SetPlayerInvItemNgt(Index, i, 11, Ngt11)
        
        'Berry
        Call SetPlayerInvItemBerry(Index, i, 1, Bry1)
        Call SetPlayerInvItemBerry(Index, i, 2, Bry2)
        Call SetPlayerInvItemBerry(Index, i, 3, Bry3)
        Call SetPlayerInvItemBerry(Index, i, 4, Bry4)
        Call SetPlayerInvItemBerry(Index, i, 5, Bry5)
        
        'Felicidade,Shiny e Sexo
        Call SetPlayerInvItemFelicidade(Index, i, Felicidade)
        Call SetPlayerInvItemShiny(Index, i, Shiny)
        Call SetPlayerInvItemSexo(Index, i, Sexo)
        
        ' Check if the player completed a quest
        ChecarTarefasAtuais Index, QUEST_TYPE_COLLECTITEMS, ItemNum
        
        Call SendInventoryUpdate(Index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim i As Long
    Dim n As Long
    Dim MapNum As Long
    Dim Msg As String
    Dim X As Long
    Dim ItemPokemon As Long, ItemPokeball As Long, ItemLevel As Long
    Dim ItemExp As Long, ItemVital(1 To Vitals.Vital_Count - 1) As Long
    Dim ItemStat(1 To Stats.Stat_Count - 1) As Long, ItemSpell(1 To MAX_POKE_SPELL) As Long
    Dim ItemMaxVital(1 To Vitals.Vital_Count - 1) As Long, Ngt(1 To MAX_NEGATIVES) As Long
    Dim Felicidade As Long, Sexo As Byte, Shiny As Byte, Bry(1 To MAX_BERRYS) As Byte

    If Not IsPlaying(Index) Then Exit Sub
    MapNum = GetPlayerMap(Index)

    For i = MAX_MAP_ITEMS To 1 Step -1
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).X = GetPlayerX(Index)) Then
                If (MapItem(MapNum, i).Y = GetPlayerY(Index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
                        
                        If MapItem(MapNum, i).Num = 3 Then
                        If Player(Index).PokeQntia >= 6 Then
                        PlayerMsg Index, "Você já possui 6 pokémons, o pokémon " & Trim$(Pokemon(MapItem(MapNum, i).PokeInfo.Pokemon).Name) & " foi enviado para o computador!", BrightGreen
                            
                        ItemPokemon = MapItem(MapNum, i).PokeInfo.Pokemon
                        ItemPokeball = MapItem(MapNum, i).PokeInfo.Pokeball
                        ItemLevel = MapItem(MapNum, i).PokeInfo.Level
                        ItemExp = MapItem(MapNum, i).PokeInfo.EXP
                        Felicidade = MapItem(MapNum, i).PokeInfo.Felicidade
                        Sexo = MapItem(MapNum, i).PokeInfo.Sexo
                        Shiny = MapItem(MapNum, i).PokeInfo.Shiny
                        
                        For X = 1 To 2
                            ItemVital(X) = MapItem(MapNum, i).PokeInfo.Vital(X)
                            ItemMaxVital(X) = MapItem(MapNum, i).PokeInfo.MaxVital(X)
                        Next
                        
                        For X = 1 To Stats.Stat_Count - 1
                            ItemStat(X) = MapItem(MapNum, i).PokeInfo.Stat(X)
                        Next
                        
                        For X = 1 To MAX_POKE_SPELL
                            ItemSpell(X) = MapItem(MapNum, i).PokeInfo.Spells(X)
                        Next
                        
                        For X = 1 To MAX_NEGATIVES
                            Ngt(X) = MapItem(MapNum, i).PokeInfo.Negatives(X)
                        Next
                        
                        For X = 1 To MAX_BERRYS
                            Bry(X) = MapItem(MapNum, i).PokeInfo.Berry(X)
                        Next
                        
                        DirectBankItemPokemon Index, 3, ItemPokemon, ItemPokeball, ItemLevel, ItemExp, ItemVital(1), ItemVital(2), ItemMaxVital(1), ItemMaxVital(2), ItemStat(1), ItemStat(4), ItemStat(2), ItemStat(3), ItemStat(5), ItemSpell(1), ItemSpell(2), ItemSpell(3), ItemSpell(4), _
                        Ngt(1), Ngt(2), Ngt(3), Ngt(4), Ngt(5), Ngt(6), Ngt(7), Ngt(8), _
                        Ngt(9), Ngt(10), Ngt(11), Felicidade, Sexo, Shiny, Bry(1), Bry(2), Bry(3), Bry(4), Bry(5)
                             
                        ' Erase item from the map
                        ClearMapItem i, MapNum
                            
                        ' Check if the player completed a quest
                        ChecarTarefasAtuais Index, QUEST_TYPE_COLLECTITEMS, MapItem(MapNum, i).Num
                                
                        Call SendInventoryUpdate(Index, n)
                        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0, vbNullString, False, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                        SendActionMsg GetPlayerMap(Index), Msg, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                        Exit Sub
                    Else
                        Player(Index).PokeQntia = Player(Index).PokeQntia + 1
                    End If
                End If
    
                    ' Open slot available?
                    If n <> 0 Then
                        ' Set item in players inventor
                        Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
    
                        If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Then
                            Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, i).Value)
                            Msg = MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                        Else
                            Call SetPlayerInvItemValue(Index, n, 0)
                            Msg = Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                        End If
                        
                        'PokeInfo
                        Call SetPlayerInvItemPokeInfoPokemon(Index, n, MapItem(MapNum, i).PokeInfo.Pokemon)
                        Call SetPlayerInvItemPokeInfoPokeball(Index, n, MapItem(MapNum, i).PokeInfo.Pokeball)
                        Call SetPlayerInvItemPokeInfoLevel(Index, n, MapItem(MapNum, i).PokeInfo.Level)
                        Call SetPlayerInvItemPokeInfoExp(Index, n, MapItem(MapNum, i).PokeInfo.EXP)
                        
                        'Max/Vitals
                        For X = 1 To Vitals.Vital_Count - 1
                            Call SetPlayerInvItemPokeInfoVital(Index, n, MapItem(MapNum, i).PokeInfo.Vital(X), X)
                            Call SetPlayerInvItemPokeInfoMaxVital(Index, n, MapItem(MapNum, i).PokeInfo.MaxVital(X), X)
                        Next
                        
                        'Stats Pokemon
                        For X = 1 To Stats.Stat_Count - 1
                            Call SetPlayerInvItemPokeInfoStat(Index, n, X, MapItem(MapNum, i).PokeInfo.Stat(X))
                        Next
                        
                        'Spells Pokemon
                        For X = 1 To MAX_POKE_SPELL
                            Call SetPlayerInvItemPokeInfoSpell(Index, n, MapItem(MapNum, i).PokeInfo.Spells(X), X)
                        Next
                        
                        'Negative
                        For X = 1 To MAX_NEGATIVES
                            Call SetPlayerInvItemNgt(Index, n, MapItem(MapNum, i).PokeInfo.Negatives(X), X)
                        Next
                        
                        'Berrys
                        For X = 1 To MAX_BERRYS
                            Call SetPlayerInvItemBerry(Index, n, X, MapItem(MapNum, i).PokeInfo.Berry(X))
                        Next
                        
                        Call SetPlayerInvItemFelicidade(Index, n, MapItem(MapNum, i).PokeInfo.Felicidade)
                        Call SetPlayerInvItemSexo(Index, n, MapItem(MapNum, i).PokeInfo.Sexo)
                        Call SetPlayerInvItemShiny(Index, n, MapItem(MapNum, i).PokeInfo.Shiny)
                        
                        ' Erase item from the map
                        ClearMapItem i, MapNum
                            
                        Call SendInventoryUpdate(Index, n)
                        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0, vbNullString, False, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                        SendActionMsg GetPlayerMap(Index), Msg, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)

                        Exit For
                    Else
                        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                        Exit For
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub ChecarQntiadePokemons(ByVal Index As Long)
Dim i As Long, PokeQntia As Byte

For i = 1 To MAX_INV
    If GetPlayerInvItemPokeInfoPokemon(Index, i) Then
        PokeQntia = PokeQntia + 1
    End If
Next

If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
    PokeQntia = PokeQntia + 1
End If

If PokeQntia > 6 Then
    PlayerMsg Index, "Você tem mais que 6 Pokémons, para não ocorrer bugs com sua conta pedimos para que deposite a quantia de " & PokeQntia & " pokémon(s) no computador!", BrightRed
End If

Player(Index).PokeQntia = PokeQntia

End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim i As Long, X As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Or TempPlayer(Index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
            i = FindOpenMapItemSlot(GetPlayerMap(Index))
            
            If Item(GetPlayerInvItemNum(Index, InvNum)).NDrop = True Then
                Call PlayerMsg(Index, "Você não pode dropar este item", BrightRed) 'ué ta certo .-.
                Exit Sub
            End If
            
            If GetPlayerInvItemNum(Index, InvNum) = 3 Then 'ops
                If Player(Index).PokeQntia >= 2 Then
                    Player(Index).PokeQntia = Player(Index).PokeQntia - 1
                Else
                    PlayerMsg Index, "Você não pode ficar sem pokémons!", BrightRed
                Exit Sub
                End If
            End If

            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).Y = GetPlayerY(Index)
                MapItem(GetPlayerMap(Index), i).canDespawn = True
                MapItem(GetPlayerMap(Index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
                
                'Poke Info
                MapItem(GetPlayerMap(Index), i).PokeInfo.Pokemon = GetPlayerInvItemPokeInfoPokemon(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).PokeInfo.Pokeball = GetPlayerInvItemPokeInfoPokeball(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).PokeInfo.Level = GetPlayerInvItemPokeInfoLevel(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).PokeInfo.EXP = GetPlayerInvItemPokeInfoExp(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).PokeInfo.Felicidade = GetPlayerInvItemFelicidade(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).PokeInfo.Sexo = GetPlayerInvItemSexo(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).PokeInfo.Shiny = GetPlayerInvItemShiny(Index, InvNum)
                
                For X = 1 To Vitals.Vital_Count - 1
                    MapItem(GetPlayerMap(Index), i).PokeInfo.Vital(X) = GetPlayerInvItemPokeInfoVital(Index, InvNum, X)
                    MapItem(GetPlayerMap(Index), i).PokeInfo.MaxVital(X) = GetPlayerInvItemPokeInfoMaxVital(Index, InvNum, X)
                Next
                
                For X = 1 To Stats.Stat_Count - 1
                    MapItem(GetPlayerMap(Index), i).PokeInfo.Stat(X) = GetPlayerInvItemPokeInfoStat(Index, InvNum, X)
                Next
                
                For X = 1 To MAX_POKE_SPELL
                    MapItem(GetPlayerMap(Index), i).PokeInfo.Spells(X) = GetPlayerInvItemPokeInfoSpell(Index, InvNum, X)
                Next
                
                For X = 1 To MAX_NEGATIVES
                    MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(X) = GetPlayerInvItemNgt(Index, InvNum, X)
                Next
                
                For X = 1 To MAX_BERRYS
                    MapItem(GetPlayerMap(Index), i).PokeInfo.Berry(X) = GetPlayerInvItemBerry(Index, InvNum, X)
                Next
                
                If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                        MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                        
                        'PokeInfo
                        Call SetPlayerInvItemPokeInfoPokemon(Index, InvNum, 0)
                        Call SetPlayerInvItemPokeInfoPokeball(Index, InvNum, 0)
                        Call SetPlayerInvItemPokeInfoLevel(Index, InvNum, 0)
                        Call SetPlayerInvItemPokeInfoExp(Index, InvNum, 0)
                        Call SetPlayerInvItemFelicidade(Index, InvNum, 0)
                        Call SetPlayerInvItemSexo(Index, InvNum, 0)
                        Call SetPlayerInvItemShiny(Index, InvNum, 0)
                        
                        'Max/Vitals
                        For X = 1 To Vitals.Vital_Count - 1
                        Call SetPlayerInvItemPokeInfoVital(Index, InvNum, 0, i)
                        Call SetPlayerInvItemPokeInfoMaxVital(Index, InvNum, 0, i)
                        Next
                        
                        'Stats Pokemon
                        For X = 1 To Stats.Stat_Count - 1
                        Call SetPlayerInvItemPokeInfoStat(Index, InvNum, i, 0)
                        Next
                        
                        'Spells Pokemon
                        For X = 1 To MAX_POKE_SPELL
                            Call SetPlayerInvItemPokeInfoSpell(Index, InvNum, 0, i)
                        Next
                        
                        'Negative
                        For X = 1 To MAX_NEGATIVES
                            Call SetPlayerInvItemNgt(Index, InvNum, 0, i)
                        Next
                        
                        'Berry
                        For X = 1 To MAX_BERRYS
                            Call SetPlayerInvItemBerry(Index, InvNum, i, 0)
                        Next
                        
                    Else
                        MapItem(GetPlayerMap(Index), i).Value = Amount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                    End If

                Else
                
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), i).Value = 0
                    
                    ' send message
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    
                    'PokeInfo
                    Call SetPlayerInvItemPokeInfoPokemon(Index, InvNum, 0)
                    Call SetPlayerInvItemPokeInfoPokeball(Index, InvNum, 0)
                    Call SetPlayerInvItemPokeInfoLevel(Index, InvNum, 0)
                    Call SetPlayerInvItemPokeInfoExp(Index, InvNum, 0)
                    Call SetPlayerInvItemFelicidade(Index, InvNum, 0)
                    Call SetPlayerInvItemSexo(Index, InvNum, 0)
                    Call SetPlayerInvItemShiny(Index, InvNum, 0)
                    
                    'Max/Vitals
                    For X = 1 To Vitals.Vital_Count - 1
                        Call SetPlayerInvItemPokeInfoVital(Index, InvNum, 0, i)
                        Call SetPlayerInvItemPokeInfoMaxVital(Index, InvNum, 0, i)
                    Next
                    
                    'Stats Pokemon
                    For X = 1 To Stats.Stat_Count - 1
                        Call SetPlayerInvItemPokeInfoStat(Index, InvNum, i, 0)
                    Next
                    
                    'Spells Pokemon
                    For X = 1 To MAX_POKE_SPELL
                        Call SetPlayerInvItemPokeInfoSpell(Index, InvNum, 0, i)
                    Next
                    
                    'Negatives
                    For X = 1 To MAX_NEGATIVES
                        Call SetPlayerInvItemNgt(Index, InvNum, 0, i)
                    Next
                    
                    'Berrys
                    For X = 1 To MAX_BERRYS
                        Call SetPlayerInvItemBerry(Index, InvNum, i, 0)
                    Next
                    
                End If

                ' Check if the player completed a quest
                ChecarTarefasAtuais Index, QUEST_TYPE_COLLECTITEMS, GetPlayerInvItemNum(Index, InvNum)
                
                ' Send inventory update
                Call SendInventoryUpdate(Index, InvNum)
         
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Amount, _
                GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), _
                Trim$(GetPlayerName(Index)), MapItem(GetPlayerMap(Index), i).canDespawn, _
                MapItem(GetPlayerMap(Index), i).PokeInfo.Pokemon, MapItem(GetPlayerMap(Index), i).PokeInfo.Pokeball, _
                MapItem(GetPlayerMap(Index), i).PokeInfo.Level, MapItem(GetPlayerMap(Index), i).PokeInfo.EXP, _
                MapItem(GetPlayerMap(Index), i).PokeInfo.Vital(1), MapItem(GetPlayerMap(Index), i).PokeInfo.Vital(2), MapItem(GetPlayerMap(Index), i).PokeInfo.MaxVital(1), MapItem(GetPlayerMap(Index), i).PokeInfo.MaxVital(2), MapItem(GetPlayerMap(Index), i).PokeInfo.Stat(1), MapItem(GetPlayerMap(Index), i).PokeInfo.Stat(4), MapItem(GetPlayerMap(Index), i).PokeInfo.Stat(2), MapItem(GetPlayerMap(Index), i).PokeInfo.Stat(3), MapItem(GetPlayerMap(Index), i).PokeInfo.Stat(5), MapItem(GetPlayerMap(Index), i).PokeInfo.Spells(1), MapItem(GetPlayerMap(Index), i).PokeInfo.Spells(2), MapItem(GetPlayerMap(Index), i).PokeInfo.Spells(3), MapItem(GetPlayerMap(Index), i).PokeInfo.Spells(4), _
                MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(1), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(2), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(3), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(4), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(5), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(6), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(7), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(8), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(9), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(10), MapItem(GetPlayerMap(Index), i).PokeInfo.Negatives(11), MapItem(GetPlayerMap(Index), i).PokeInfo.Felicidade, MapItem(GetPlayerMap(Index), i).PokeInfo.Sexo, _
                MapItem(GetPlayerMap(Index), i).PokeInfo.Shiny, MapItem(GetPlayerMap(Index), i).PokeInfo.Berry(1), MapItem(GetPlayerMap(Index), i).PokeInfo.Berry(2), MapItem(GetPlayerMap(Index), i).PokeInfo.Berry(3), MapItem(GetPlayerMap(Index), i).PokeInfo.Berry(4), MapItem(GetPlayerMap(Index), i).PokeInfo.Berry(5))
                
            Else
                Call PlayerMsg(Index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
    Dim i As Long, X As Long, Z As Long, SlotVazio As Boolean
    Dim expRollover As Long
    Dim level_count As Long
    
    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) = 0 Then Exit Sub
    
    level_count = 0
    
    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        expRollover = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
        
        ' can level up?
        If Not SetPlayerLevel(Index, GetPlayerLevel(Index) + 1) Then
            Exit Sub
        End If

        'Evoluir
        If Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Evolução(1).Level = GetPlayerEquipmentPokeInfoLevel(Index, weapon) Then
            Player(Index).EvolPermition = 1
            Player(Index).EvoId = 1
            SendPokeEvolution Index, 0
        End If
        
        'Setar atributos pronto Up = +3 all stats
        For i = 1 To Stats.Stat_Count - 1
            Call SetPlayerEquipmentPokeInfoStat(Index, GetPlayerEquipmentPokeInfoStat(Index, weapon, i) + 3, weapon, i)
        Next
        
        'Setar Felicidade
        Call SetPlayerEquipmentFelicidade(Index, weapon, GetPlayerEquipmentFelicidade(Index, weapon) + 2)
        
        'Setar Novas Habilidade
        For X = 1 To 20
        SlotVazio = False
        
        If Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Habilidades(X).Level > 0 Then
            If GetPlayerEquipmentPokeInfoLevel(Index, weapon) = Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Habilidades(X).Level Then
            
                If Player(Index).LearnSpell(1) > 0 Then
                    For Z = 1 To 10
                        If Player(Index).LearnFila(Z) = 0 Then
                            Player(Index).LearnFila(Z) = Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Habilidades(X).Spell
                            Exit For
                        End If
                    Next
                Else
                    For Z = 1 To MAX_POKE_SPELL
                        If GetPlayerEquipmentPokeInfoSpell(Index, weapon, Z) = 0 Then
                            Call SetPlayerEquipmentPokeInfoSpell(Index, Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Habilidades(X).Spell, weapon, Z)
                            Call SetPlayerSpell(Index, Z, Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Habilidades(X).Spell)
                            SendPlayerSpells Index
                            SlotVazio = True
                            Exit For
                        End If
                    Next
                    
                    If SlotVazio = False Then
                        Player(Index).LearnSpell(1) = 1
                        Player(Index).LearnSpell(2) = Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Habilidades(X).Spell
                        Player(Index).LearnSpell(3) = 0
                        SendAprenderSpell Index, 0
                    End If
                End If
            End If
        End If
        Next
        
        'setar Vitals...
        Call SetPlayerEquipmentPokeInfoMaxVital(Index, GetPlayerEquipmentPokeInfoMaxVital(Index, weapon, 1) + 5, weapon, 1) 'Hp
        Call SetPlayerEquipmentPokeInfoMaxVital(Index, GetPlayerEquipmentPokeInfoMaxVital(Index, weapon, 2) + 5, weapon, 2) 'Mp
        
        Call SetPlayerExp(Index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            PlayerMsg Index, Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " subiu " & level_count & " Level!", Yellow
        Else
            'plural
            PlayerMsg Index, Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " subiu " & level_count & " Levels!", Yellow
        End If

        SendAnimation GetPlayerMap(Index), 12, GetPlayerX(Index), GetPlayerY(Index)

        'Setar Hp Máximo
        SetPlayerVital Index, HP, GetPlayerMaxVital(Index, HP)
        SetPlayerVital Index, HP, GetPlayerMaxVital(Index, MP)

        For i = 1 To Vitals.Vital_Count - 1
            SendVital Index, i
        Next
        
        SendEXP Index
        SendPlayerData Index
        SendWornEquipment (Index)
        Call UpdateRankLevel(Index)
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
    Player(Index).MySprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If GetPlayerEquipment(Index, weapon) > 0 Then
    GetPlayerLevel = GetPlayerEquipmentPokeInfoLevel(Index, weapon)
    Else
    GetPlayerLevel = 1
    End If
End Function

Function SetPlayerLevel(ByVal Index As Long, ByVal Level As Long) As Boolean
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
    Call SetPlayerEquipmentPokeInfoLevel(Index, Level, weapon)
    SetPlayerLevel = True
    End If
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long
Dim CurrentLevel As Integer

    GetPlayerNextLevel = 100
    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) = 0 Then Exit Function
    
    CurrentLevel = GetPlayerLevel(Index)
    If CurrentLevel <= 1 Then CurrentLevel = 2 'Evitar Erro Level 1
    
    Select Case Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).ExpType
    Case 0 'Rápido
    GetPlayerNextLevel = 0.8 * (CurrentLevel) ^ 3
    Case 1 'Médio Rápido
    GetPlayerNextLevel = (CurrentLevel) ^ 3
    Case 2 'Médio Lento
    GetPlayerNextLevel = (CurrentLevel) ^ 3 + ((((CurrentLevel) ^ 3) * 5.9) / 100)
    Case 3 'Lento
    GetPlayerNextLevel = 1.25 * (CurrentLevel) ^ 3
    Case Else 'Rápido
    GetPlayerNextLevel = 0.8 * (CurrentLevel) ^ 3
    End Select
    
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    If GetPlayerEquipmentPokeInfoExp(Index, weapon) > 0 Then
    GetPlayerExp = GetPlayerEquipmentPokeInfoExp(Index, weapon)
    Else
    GetPlayerExp = 1
    End If
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
    If GetPlayerEquipmentPokeInfoLevel(Index, weapon) = MAX_LEVELS And GetPlayerEquipmentPokeInfoExp(Index, weapon) >= GetPlayerNextLevel(Index) Then
        Call SetPlayerEquipmentPokeInfoExp(Index, GetPlayerNextLevel(Index), weapon)
        Exit Sub
    End If
    
    Call SetPlayerEquipmentPokeInfoExp(Index, EXP, weapon)
    End If
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerHonra(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerHonra = Player(Index).Honra
End Function

Sub SetPlayerHonra(ByVal Index As Long, ByVal Honra As Long)
    Player(Index).Honra = Honra
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
Dim i As Long

    If Vital = 0 Then Exit Sub
    
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
        If Vital = HP Then
            Call SetPlayerEquipmentPokeInfoVital(Index, Player(Index).Vital(Vital), weapon, 1)
        End If
    
        If Vital = MP Then
            Call SetPlayerEquipmentPokeInfoVital(Index, Player(Index).Vital(Vital), weapon, 2)
        End If
    End If

    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Dim X As Long, i As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    X = Player(Index).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(Index).Equipment(i) > 0 Then
            If Item(Player(Index).Equipment(i)).Add_Stat(Stat) > 0 Then
                X = X + Item(Player(Index).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = X
End Function

Public Function GetPlayerRawStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(Index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Or Index = 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Map = MapNum
    End If

End Sub

Function GetPlayerX(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(Index).Inv(invslot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(invslot).Num = ItemNum
End Sub

Function GetPlayerVisuais(ByVal Index As Long, ByVal ViSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If ViSlot = 0 Then Exit Function
    
    GetPlayerVisuais = Player(Index).Visuais(ViSlot)
End Function

Sub SetPlayerVisuais(ByVal Index As Long, ByVal ViSlot As Long, ByVal ViNum As Long)
    Player(Index).Visuais(ViSlot) = ViNum
End Sub

Function GetPlayerTeleport(ByVal Index As Long, ByVal TpSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If TpSlot = 0 Then Exit Function
    
    GetPlayerTeleport = Player(Index).Teleport(TpSlot)
End Function

Sub SetPlayerTeleport(ByVal Index As Long, ByVal TpSlot As Long, ByVal TpNum As Long)
    Player(Index).Teleport(TpSlot) = TpNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(Index).Inv(invslot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(invslot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(Index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal SpellNum As Long)
    Player(Index).Spell(spellslot) = SpellNum
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot) = InvNum
End Sub

' ToDo
Sub OnDeath(ByVal Index As Long)
    Dim i As Long, X As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(Index, Vitals.HP, 0)
    
        ' Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                    If TempPlayer(i).target = Index Then
                        TempPlayer(i).target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next

    ' Drop all worn items
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(Index, i) > 0 Then
            PlayerMapDropItem Index, GetPlayerEquipment(Index, i), 0
        End If
    Next

    PlayerUnequipItem Index, weapon

    ' Warp player away
    Call SetPlayerDir(Index, DIR_DOWN)
    
    With Map(GetPlayerMap(Index))
    
        If Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_ARENA Then
            If .BootMap > 0 Then
                PlayerWarp Index, .BootMap, .BootX, .BootY
            Else
               Call PlayerWarp(Index, START_MAP, START_X, START_Y)
            End If
        End If
            
        If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_ARENA Then
            PlayerUnequipItem Index, weapon
            Call PlayerWarp(Index, Player(Index).MyMap(1), Player(Index).MyMap(2), Player(Index).MyMap(3))
        End If
        
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(Index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(Index).spellBuffer.Spell = 0
    TempPlayer(Index).spellBuffer.Timer = 0
    TempPlayer(Index).spellBuffer.target = 0
    TempPlayer(Index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(Index)
    
    TempPlayer(Index).InBank = False
    TempPlayer(Index).InShop = 0
    If TempPlayer(Index).InTrade > 0 Then
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
            
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).Num = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).Value = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.Pokemon = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.Pokeball = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.Level = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.EXP = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.Felicidade = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.Sexo = 0
            TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.Shiny = 0
            
            For X = 1 To Vitals.Vital_Count - 1
                    TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.Vital(X) = 0
                    TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.MaxVital(X) = 0
            Next
            
            For X = 1 To Stats.Stat_Count - 1
                    TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.Stat(X) = 0
            Next
            
            For X = 1 To MAX_POKE_SPELL
                    TempPlayer(TempPlayer(Index).InTrade).TradeOffer(i).PokeInfo.Spells(X) = 0
            Next
        Next

        TempPlayer(Index).InTrade = 0
        TempPlayer(TempPlayer(Index).InTrade).InTrade = 0
        
        SendCloseTrade Index
        SendCloseTrade TempPlayer(Index).InTrade
    End If
    
    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
    End If

End Sub

Sub CheckResource(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal SpellNum As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    
    ' Check attack timer
    If GetPlayerEquipment(Index, weapon) > 0 Then
        If GetTickCount < TempPlayer(Index).AttackTimer + Item(GetPlayerEquipment(Index, weapon)).Speed Then Exit Sub
    Else
        If GetTickCount < TempPlayer(Index).AttackTimer + 1000 Then Exit Sub
    End If
    
    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(Index)).Tile(X, Y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count

            If ResourceCache(GetPlayerMap(Index)).ResourceData(i).X = X Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(i).Y = Y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(Index, weapon) > 0 Then
                If Item(GetPlayerEquipment(Index, weapon)).Data3 = Resource(Resource_index).ToolRequired Then

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(Index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).X
                        rY = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).Y
                        
                        'Item Break
                        If Resource(Resource_index).Spell = 0 Then
                            Damage = Item(GetPlayerEquipment(Index, weapon)).Data2
                        Else
                        
                        'Spell Break
                        If SpellNum > 0 Then
                            If Resource(Resource_index).Spell <> SpellNum Then Exit Sub
                            Damage = Spell(SpellNum).Vital 'Dano da spell
                            End If
                        End If
                    
                        ' check if damage is more than health
                        If Damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                SendActionMsg GetPlayerMap(Index), "-" & ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                End If
                                ' carry on
                                GiveInvItem Index, Resource(Resource_index).ItemReward, 1
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage
                                SendActionMsg GetPlayerMap(Index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            End If
                            ' send the sound
                            SendMapSound Index, rX, rY, SoundEntity.seResource, Resource_index
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(Index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                        End If
                    End If
                    
                     ' Reset attack timer
                    TempPlayer(Index).AttackTimer = GetTickCount
                Else
                    PlayerMsg Index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg Index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Public Function CheckResourceStatCut(ByVal Index As Long, ByVal X As Long, ByVal Y As Long) As Boolean
Dim i As Long, Resource_num As Long

CheckResourceStatCut = False

    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count

            If ResourceCache(GetPlayerMap(Index)).ResourceData(i).X = X Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(i).Y = Y Then
                    Resource_num = i
                End If
            End If

        Next
    
    If Resource_num = 0 Then Exit Function
        
    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1 Then
    CheckResourceStatCut = True
    End If
    End If
    
End Function

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(Index).Item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Bank(Index).Item(BankSlot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(Index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(Index).Item(BankSlot).Value = ItemValue
End Sub

'#################################Poke Info#####################################

Function GetPlayerBankItemPokemon(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemPokemon = Bank(Index).Item(BankSlot).PokeInfo.Pokemon
End Function

Sub SetPlayerBankItemPokemon(ByVal Index As Long, ByVal BankSlot As Long, ByVal PokemonNum As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Pokemon = PokemonNum
End Sub

Function GetPlayerBankItemPokeball(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemPokeball = Bank(Index).Item(BankSlot).PokeInfo.Pokeball
End Function

Sub SetPlayerBankItemPokeball(ByVal Index As Long, ByVal BankSlot As Long, ByVal PokeballNum As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Pokeball = PokeballNum
End Sub

Function GetPlayerBankItemLevel(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemLevel = Bank(Index).Item(BankSlot).PokeInfo.Level
End Function

Sub SetPlayerBankItemLevel(ByVal Index As Long, ByVal BankSlot As Long, ByVal LevelValue As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Level = LevelValue
End Sub

Function GetPlayerBankItemExp(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemExp = Bank(Index).Item(BankSlot).PokeInfo.EXP
End Function

Sub SetPlayerBankItemExp(ByVal Index As Long, ByVal BankSlot As Long, ByVal ExpValue As Long)
    Bank(Index).Item(BankSlot).PokeInfo.EXP = ExpValue
End Sub

Function GetPlayerBankItemVital(ByVal Index As Long, ByVal BankSlot As Long, ByVal VitalType As Long) As Long
    GetPlayerBankItemVital = Bank(Index).Item(BankSlot).PokeInfo.Vital(VitalType)
End Function

Sub SetPlayerBankItemVital(ByVal Index As Long, ByVal BankSlot As Long, ByVal VitalValue As Long, ByVal VitalType As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Vital(VitalType) = VitalValue
End Sub

Function GetPlayerBankItemMaxVital(ByVal Index As Long, ByVal BankSlot As Long, ByVal VitalType As Long) As Long
    GetPlayerBankItemMaxVital = Bank(Index).Item(BankSlot).PokeInfo.MaxVital(VitalType)
End Function

Sub SetPlayerBankItemMaxVital(ByVal Index As Long, ByVal BankSlot As Long, ByVal VitalValue As Long, ByVal VitalType As Long)
    Bank(Index).Item(BankSlot).PokeInfo.MaxVital(VitalType) = VitalValue
End Sub

Function GetPlayerBankItemStat(ByVal Index As Long, ByVal BankSlot As Long, ByVal StatNum As Long) As Long
    GetPlayerBankItemStat = Bank(Index).Item(BankSlot).PokeInfo.Stat(StatNum)
End Function

Sub SetPlayerBankItemStat(ByVal Index As Long, ByVal BankSlot As Long, ByVal StatValue As Long, ByVal StatNum As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Stat(StatNum) = StatValue
End Sub

Function GetPlayerBankItemSpell(ByVal Index As Long, ByVal BankSlot As Long, ByVal Spell As Long) As Long
    GetPlayerBankItemSpell = Bank(Index).Item(BankSlot).PokeInfo.Spells(Spell)
End Function

Sub SetPlayerBankItemSpell(ByVal Index As Long, ByVal BankSlot As Long, ByVal SpellValue As Long, ByVal Spell As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Spells(Spell) = SpellValue
End Sub

Function GetPlayerBankItemNgt(ByVal Index As Long, ByVal BankSlot As Long, ByVal Ngt As Long) As Long
    GetPlayerBankItemNgt = Bank(Index).Item(BankSlot).PokeInfo.Negatives(Ngt)
End Function

Sub SetPlayerBankItemNgt(ByVal Index As Long, ByVal BankSlot As Long, ByVal NgtValue As Long, ByVal Ngt As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Negatives(Ngt) = NgtValue
End Sub

Function GetPlayerBankItemFelicidade(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemFelicidade = Bank(Index).Item(BankSlot).PokeInfo.Felicidade
End Function

Sub SetPlayerBankItemFelicidade(ByVal Index As Long, ByVal BankSlot As Long, ByVal FelicidadeValue As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Felicidade = FelicidadeValue
End Sub

Function GetPlayerBankItemSexo(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemSexo = Bank(Index).Item(BankSlot).PokeInfo.Sexo
End Function

Sub SetPlayerBankItemSexo(ByVal Index As Long, ByVal BankSlot As Long, ByVal SexoValue As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Sexo = SexoValue
End Sub

Function GetPlayerBankItemShiny(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemShiny = Bank(Index).Item(BankSlot).PokeInfo.Shiny
End Function

Sub SetPlayerBankItemShiny(ByVal Index As Long, ByVal BankSlot As Long, ByVal ShinyValue As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Shiny = ShinyValue
End Sub

'Berry
Function GetPlayerBankItemBerry(ByVal Index As Long, ByVal BankSlot As Long, ByVal BerryStat As Long) As Long
    GetPlayerBankItemBerry = Bank(Index).Item(BankSlot).PokeInfo.Berry(BerryStat)
End Function

Sub SetPlayerBankItemBerry(ByVal Index As Long, ByVal BankSlot As Long, ByVal BerryStat As Long, ByVal Valor As Long)
    Bank(Index).Item(BankSlot).PokeInfo.Berry(BerryStat) = Valor
End Sub

'123

'###############################################################################

Sub DirectBankItemPokemon(ByVal Index As Long, ByVal ItemNum As Long, Optional ByVal Pokemon As Long, Optional ByVal Pokeball As Long, Optional ByVal Level As Long, Optional ByVal EXP As Long, Optional ByVal VitalHP As Long, Optional ByVal VitalMP As Long, Optional ByVal MaxVitalHp As Long, Optional ByVal MaxVitalMp As Long, Optional ByVal StatStr As Long, Optional ByVal StatAgi As Long, Optional ByVal StatEnd As Long, Optional ByVal StatInt As Long, Optional ByVal StatWill As Long, Optional ByVal Spell1 As Long, Optional ByVal Spell2 As Long, Optional ByVal Spell3 As Long, Optional ByVal Spell4 As Long, _
Optional ByVal Ngt1 As Long, Optional ByVal Ngt2 As Long, _
Optional ByVal Ngt3 As Long, Optional ByVal Ngt4 As Long, _
Optional ByVal Ngt5 As Long, Optional ByVal Ngt6 As Long, _
Optional ByVal Ngt7 As Long, Optional ByVal Ngt8 As Long, _
Optional ByVal Ngt9 As Long, Optional ByVal Ngt10 As Long, _
Optional ByVal Ngt11 As Long, Optional ByVal Felicidade As Long, _
Optional ByVal Sexo As Byte, Optional ByVal Shiny As Long, _
Optional ByVal Bry1 As Long, Optional ByVal Bry2 As Long, Optional ByVal Bry3 As Long, _
Optional ByVal Bry4 As Long, Optional ByVal Bry5 As Long)

Dim BankSlot As Long, i As Long

    BankSlot = FindOpenBankSlotPokemon(Index)
    
    If BankSlot > 0 Then
    
            SetPlayerBankItemNum Index, BankSlot, ItemNum
            SetPlayerBankItemValue Index, BankSlot, 1
            SetPlayerBankItemPokemon Index, BankSlot, Pokemon
            SetPlayerBankItemPokeball Index, BankSlot, Pokeball
            SetPlayerBankItemLevel Index, BankSlot, Level
            SetPlayerBankItemExp Index, BankSlot, EXP
                
            SetPlayerBankItemVital Index, BankSlot, VitalHP, 1
            SetPlayerBankItemVital Index, BankSlot, VitalMP, 2
            SetPlayerBankItemMaxVital Index, BankSlot, MaxVitalHp, 1
            SetPlayerBankItemMaxVital Index, BankSlot, MaxVitalMp, 2
                
            SetPlayerBankItemStat Index, BankSlot, StatStr, 1
            SetPlayerBankItemStat Index, BankSlot, StatEnd, 2
            SetPlayerBankItemStat Index, BankSlot, StatInt, 3
            SetPlayerBankItemStat Index, BankSlot, StatAgi, 4
            SetPlayerBankItemStat Index, BankSlot, StatWill, 5
            
            SetPlayerBankItemSpell Index, BankSlot, Spell1, 1
            SetPlayerBankItemSpell Index, BankSlot, Spell2, 2
            SetPlayerBankItemSpell Index, BankSlot, Spell3, 3
            SetPlayerBankItemSpell Index, BankSlot, Spell4, 4
            
            SetPlayerBankItemNgt Index, BankSlot, Ngt1, 1
            SetPlayerBankItemNgt Index, BankSlot, Ngt2, 2
            SetPlayerBankItemNgt Index, BankSlot, Ngt3, 3
            SetPlayerBankItemNgt Index, BankSlot, Ngt4, 4
            SetPlayerBankItemNgt Index, BankSlot, Ngt5, 5
            SetPlayerBankItemNgt Index, BankSlot, Ngt6, 6
            SetPlayerBankItemNgt Index, BankSlot, Ngt7, 7
            SetPlayerBankItemNgt Index, BankSlot, Ngt8, 8
            SetPlayerBankItemNgt Index, BankSlot, Ngt9, 9
            SetPlayerBankItemNgt Index, BankSlot, Ngt10, 10
            SetPlayerBankItemNgt Index, BankSlot, Ngt11, 11
            
            SetPlayerBankItemBerry Index, BankSlot, 1, Bry1
            SetPlayerBankItemBerry Index, BankSlot, 2, Bry2
            SetPlayerBankItemBerry Index, BankSlot, 3, Bry3
            SetPlayerBankItemBerry Index, BankSlot, 4, Bry4
            SetPlayerBankItemBerry Index, BankSlot, 5, Bry5
            
            SetPlayerBankItemFelicidade Index, BankSlot, Felicidade
            SetPlayerBankItemSexo Index, BankSlot, Sexo
            SetPlayerBankItemShiny Index, BankSlot, Shiny

    Else
            PlayerMsg Index, "O Computador está cheio, a pokéball quebrou e o Pokémon fugiu!", BrightRed
    End If
    
            SaveBank Index
            SavePlayer Index
    
End Sub

Sub GiveBankItemPokemon(ByVal Index As Long, ByVal invslot As Long, ByVal Amount As Long)
Dim BankSlot As Long, i As Long
Dim ItemNum As Long, ItemPokemon As Long, ItemPokeball As Long, ItemLevel As Long
Dim ItemExp As Long, ItemVital(1 To Vitals.Vital_Count - 1) As Long
Dim ItemStat(1 To Stats.Stat_Count - 1) As Long, ItemSpell(1 To MAX_POKE_SPELL) As Long
Dim ItemMaxVital(1 To Vitals.Vital_Count - 1) As Long, Ngt(1 To MAX_NEGATIVES) As Long
Dim Felicidade As Long, Sexo As Byte, Shiny As Byte, Bry(1 To MAX_BERRYS) As Long

    ItemNum = GetPlayerInvItemNum(Index, invslot)
    ItemPokemon = GetPlayerInvItemPokeInfoPokemon(Index, invslot)
    ItemPokeball = GetPlayerInvItemPokeInfoPokeball(Index, invslot)
    ItemLevel = GetPlayerInvItemPokeInfoLevel(Index, invslot)
    ItemExp = GetPlayerInvItemPokeInfoExp(Index, invslot)
    Felicidade = GetPlayerInvItemFelicidade(Index, invslot)
    Sexo = GetPlayerInvItemSexo(Index, invslot)
    Shiny = GetPlayerInvItemShiny(Index, invslot)

    If Item(GetPlayerInvItemNum(Index, invslot)).NDeposit = True Then
        PlayerMsg Index, "Este item não pode ser depositado", BrightRed
        Exit Sub
    End If

    For i = 1 To Vitals.Vital_Count - 1
        ItemVital(i) = GetPlayerInvItemPokeInfoVital(Index, invslot, i)
        ItemMaxVital(i) = GetPlayerInvItemPokeInfoMaxVital(Index, invslot, i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        ItemStat(i) = GetPlayerInvItemPokeInfoStat(Index, invslot, i)
    Next
    
    For i = 1 To MAX_POKE_SPELL
        ItemSpell(i) = GetPlayerInvItemPokeInfoSpell(Index, invslot, i)
    Next
    
    For i = 1 To MAX_NEGATIVES
        Ngt(i) = GetPlayerInvItemNgt(Index, invslot, i)
    Next
    
    For i = 1 To MAX_BERRYS
        Bry(i) = GetPlayerInvItemBerry(Index, invslot, i)
    Next

    If invslot < 0 Or invslot > MAX_INV Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invslot) Then
        Exit Sub
    End If
    
    If Item(GetPlayerInvItemNum(Index, invslot)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Then Exit Sub
    End If
    
    BankSlot = FindOpenBankSlotPokemon(Index)
        
    If BankSlot > 0 Then
    
                'Capacidade de Pokémons
                If GetPlayerInvItemPokeInfoPokemon(Index, invslot) > 0 Then
                    If Not Player(Index).PokeQntia = 1 Or Player(Index).PokeQntia = 0 Then
                        Player(Index).PokeQntia = Player(Index).PokeQntia - 1
                    Else
                        PlayerMsg Index, "Você não pode ficar sem pokémons!", BrightRed
                        Exit Sub
                    End If
                End If
    
                SetPlayerBankItemNum Index, BankSlot, ItemNum
                SetPlayerBankItemValue Index, BankSlot, 1
                SetPlayerBankItemPokemon Index, BankSlot, ItemPokemon
                SetPlayerBankItemPokeball Index, BankSlot, ItemPokeball
                SetPlayerBankItemLevel Index, BankSlot, ItemLevel
                SetPlayerBankItemExp Index, BankSlot, ItemExp
                SetPlayerBankItemFelicidade Index, BankSlot, Felicidade
                SetPlayerBankItemSexo Index, BankSlot, Sexo
                SetPlayerBankItemShiny Index, BankSlot, Shiny
                
                For i = 1 To Vitals.Vital_Count - 1
                    SetPlayerBankItemVital Index, BankSlot, ItemVital(i), i
                    SetPlayerBankItemMaxVital Index, BankSlot, ItemMaxVital(i), i
                Next
                
                For i = 1 To Stats.Stat_Count - 1
                    SetPlayerBankItemStat Index, BankSlot, ItemStat(i), i
                Next
                
                For i = 1 To MAX_POKE_SPELL
                    SetPlayerBankItemSpell Index, BankSlot, ItemSpell(i), i
                Next
                
                For i = 1 To MAX_NEGATIVES
                    SetPlayerBankItemNgt Index, BankSlot, Ngt(i), i
                Next
                
                For i = 1 To MAX_BERRYS
                    SetPlayerBankItemBerry Index, BankSlot, i, Bry(i)
                Next
                
                'Pegar item do inventario
                Call TakeInvSlot(Index, invslot, 1)
        Else
                'Não depositar
                PlayerMsg Index, "O computador está Cheio!", BrightRed
                Exit Sub
        End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index
    SendInventory Index

End Sub

Sub TakeBankItemPokemon(ByVal Index As Long, ByVal BankSlot As Long)
Dim invslot, i As Long
Dim ItemNum As Long, ItemPokemon As Long, ItemPokeball As Long, ItemLevel As Long
Dim ItemExp As Long, ItemVital(1 To Vitals.Vital_Count - 1) As Long
Dim ItemStat(1 To Stats.Stat_Count - 1) As Long, ItemSpell(1 To MAX_POKE_SPELL) As Long
Dim ItemMaxVital(1 To Vitals.Vital_Count - 1) As Long, Ngt(1 To MAX_NEGATIVES) As Long
Dim Felicidade As Long, Sexo As Byte, Shiny As Byte, Bry(1 To MAX_BERRYS) As Long

    ItemNum = GetPlayerBankItemNum(Index, BankSlot)
    ItemPokemon = GetPlayerBankItemPokemon(Index, BankSlot)
    ItemPokeball = GetPlayerBankItemPokeball(Index, BankSlot)
    ItemLevel = GetPlayerBankItemLevel(Index, BankSlot)
    ItemExp = GetPlayerBankItemExp(Index, BankSlot)
    Felicidade = GetPlayerBankItemFelicidade(Index, BankSlot)
    Sexo = GetPlayerBankItemSexo(Index, BankSlot)
    Shiny = GetPlayerBankItemShiny(Index, BankSlot)
    
    For i = 1 To Vitals.Vital_Count - 1
        ItemVital(i) = GetPlayerBankItemVital(Index, BankSlot, i)
        ItemMaxVital(i) = GetPlayerBankItemMaxVital(Index, BankSlot, i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        ItemStat(i) = GetPlayerBankItemStat(Index, BankSlot, i)
    Next
    
    For i = 1 To MAX_POKE_SPELL
        ItemSpell(i) = GetPlayerBankItemSpell(Index, BankSlot, i)
    Next
    
    For i = 1 To MAX_NEGATIVES
        Ngt(i) = GetPlayerBankItemNgt(Index, BankSlot, i)
    Next
    
    For i = 1 To MAX_BERRYS
        Bry(i) = GetPlayerBankItemBerry(Index, BankSlot, i)
    Next

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    invslot = FindOpenInvSlot(Index, GetPlayerBankItemNum(Index, BankSlot))
        
    If invslot > 0 Then
    
    If Not Player(Index).PokeQntia >= 6 Then
        Player(Index).PokeQntia = Player(Index).PokeQntia + 1
    Else
        PlayerMsg Index, "Você já tem 6 pokémons!", BrightRed
    Exit Sub
    End If
        
        GiveInvItem Index, ItemNum, 1, False, ItemPokemon, _
        ItemPokeball, ItemLevel, ItemExp, ItemVital(1), _
        ItemVital(2), ItemMaxVital(1), ItemMaxVital(2), _
        ItemStat(1), ItemStat(4), ItemStat(2), ItemStat(3), _
        ItemStat(5), ItemSpell(1), ItemSpell(2), ItemSpell(3), ItemSpell(4), _
        Ngt(1), Ngt(2), Ngt(3), Ngt(4), Ngt(5), Ngt(6), Ngt(7), _
        Ngt(8), Ngt(9), Ngt(10), Ngt(11), Felicidade, Sexo, Shiny, Bry(1), Bry(2), Bry(3), Bry(4), Bry(5)
        
        SetPlayerBankItemNum Index, BankSlot, 0
        SetPlayerBankItemValue Index, BankSlot, 0
        SetPlayerBankItemPokemon Index, BankSlot, 0
        SetPlayerBankItemPokeball Index, BankSlot, 0
        SetPlayerBankItemLevel Index, BankSlot, 0
        SetPlayerBankItemExp Index, BankSlot, 0
        SetPlayerBankItemFelicidade Index, BankSlot, 0
        SetPlayerBankItemSexo Index, BankSlot, 0
        SetPlayerBankItemShiny Index, BankSlot, 0
        
        For i = 1 To Vitals.Vital_Count - 1
            SetPlayerBankItemVital Index, BankSlot, 0, i
            SetPlayerBankItemMaxVital Index, BankSlot, 0, i
        Next
        
        For i = 1 To Stats.Stat_Count - 1
            SetPlayerBankItemStat Index, BankSlot, 0, i
        Next
        
        For i = 1 To MAX_POKE_SPELL
            SetPlayerBankItemSpell Index, BankSlot, 0, i
        Next
        
        For i = 1 To MAX_NEGATIVES
            SetPlayerBankItemNgt Index, BankSlot, 0, i
        Next
        
        For i = 1 To MAX_BERRYS
            SetPlayerBankItemBerry Index, BankSlot, i, 0
        Next
        
    Else
        PlayerMsg Index, "O seu inventario está cheio!", White
        Exit Sub
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index
    SendInventory Index

End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal invslot As Long, ByVal Amount As Long)
Dim BankSlot

    If invslot < 0 Or invslot > MAX_INV Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invslot) Then
        Exit Sub
    End If
    
    If Item(GetPlayerInvItemNum(Index, invslot)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Then Exit Sub
    End If
    
    If Item(GetPlayerInvItemNum(Index, invslot)).NDeposit = True Then
        PlayerMsg Index, "Este item não pode ser depositado", BrightRed
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(Index, GetPlayerInvItemNum(Index, invslot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, invslot)).Type = ITEM_TYPE_CURRENCY Then
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invslot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invslot), Amount)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invslot))
                Call SetPlayerBankItemValue(Index, BankSlot, Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invslot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invslot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invslot), 0)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invslot))
                Call SetPlayerBankItemValue(Index, BankSlot, 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invslot), 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim invslot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerBankItemValue(Index, BankSlot) Then
        Exit Sub
    End If
    
    If Item(GetPlayerBankItemNum(Index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
        If Amount = 0 Then Exit Sub
    End If
    
    invslot = FindOpenInvSlot(Index, GetPlayerBankItemNum(Index, BankSlot))
        
    If invslot > 0 Then
        If Item(GetPlayerBankItemNum(Index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), Amount)
            Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - Amount)
            If GetPlayerBankItemValue(Index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(Index, BankSlot) > 1 Then
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - 1)
            Else
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

End Sub

Public Sub KillPlayer(ByVal Index As Long)
Dim EXP As Long

    ' Calculate exp to give attacker
    EXP = GetPlayerExp(Index) \ 3

    ' Make sure we dont get less then 0
    If EXP < 0 Then EXP = 0
    If EXP = 0 Then
        Call PlayerMsg(Index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(Index, GetPlayerExp(Index) - EXP)
        SendEXP Index
        Call PlayerMsg(Index, "You lost " & EXP & " exp.", BrightRed)
    End If
    
    Call OnDeath(Index)
End Sub

Public Sub UseItem(ByVal Index As Long, ByVal InvNum As Long)
Dim n As Long, i As Long, tempItem As Long, X As Long, Y As Long, ItemNum As Long
Dim Chance As Long, R As Long, InvVazio As Byte, BauItemQnt As Byte
Dim Buffer As clsBuffer

Dim TempItemPokemon As Long, TempItemPokeball As Long, TempItemLevel As Long, TempItemExp As Long
Dim TempItemVital(1 To Vitals.Vital_Count - 1) As Long, TempItemMaxVital(1 To Vitals.Vital_Count - 1) As Long
Dim TempItemStat(1 To Stats.Stat_Count - 1) As Long, TempItemSpell(1 To MAX_POKE_SPELL) As Long, TempItemNgt(1 To 11), TempItemFelicidade, TempItemSexo, TempItemShiny
Dim TempItemBry(1 To MAX_BERRYS) As Long

Dim ItemPokemon As Long, ItemPokeball As Long, ItemLevel As Long, ItemExp As Long
Dim ItemVital(1 To Vitals.Vital_Count - 1) As Long, ItemMaxVital(1 To Vitals.Vital_Count - 1) As Long
Dim ItemStat(1 To Stats.Stat_Count - 1) As Long, ItemSpell(1 To MAX_POKE_SPELL) As Long, ItemNgt(1 To 11), ItemFelicidade, ItemSexo, ItemShiny
Dim PokebolaUsada As Byte, ItemBry(1 To MAX_BERRYS) As Long

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    ' Prevent Using Item In Trade
    If TempPlayer(Index).InTrade > 0 Then Exit Sub
    
    ' Prevent Using Item In Fishing
    If Player(Index).InFishing > 0 Then
        Exit Sub
    End If
    
    ' Prevent Using Item Learning Pókemon Spell
    If Player(Index).LearnSpell(1) > 0 Then
        Exit Sub
    End If
    
    If TempPlayer(Index).ScanTime > 0 Then Exit Sub
    
    If Player(Index).EvolPermition = 1 Then Exit Sub

    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
        ItemNum = GetPlayerInvItemNum(Index, InvNum)
        
        ItemPokemon = GetPlayerInvItemPokeInfoPokemon(Index, InvNum)
        ItemPokeball = GetPlayerInvItemPokeInfoPokeball(Index, InvNum)
        ItemLevel = GetPlayerInvItemPokeInfoLevel(Index, InvNum)
        ItemExp = GetPlayerInvItemPokeInfoExp(Index, InvNum)
        ItemFelicidade = GetPlayerInvItemFelicidade(Index, InvNum)
        ItemSexo = GetPlayerInvItemSexo(Index, InvNum)
        ItemShiny = GetPlayerInvItemShiny(Index, InvNum)
        
        For i = 1 To Vitals.Vital_Count - 1
            ItemVital(i) = GetPlayerInvItemPokeInfoVital(Index, InvNum, i)
            ItemMaxVital(i) = GetPlayerInvItemPokeInfoMaxVital(Index, InvNum, i)
        Next
        
        For i = 1 To Stats.Stat_Count - 1
            ItemStat(i) = GetPlayerInvItemPokeInfoStat(Index, InvNum, i)
        Next
        
        For i = 1 To MAX_POKE_SPELL
            ItemSpell(i) = GetPlayerInvItemPokeInfoSpell(Index, InvNum, i)
        Next
        
        For i = 1 To MAX_NEGATIVES
            ItemNgt(i) = GetPlayerInvItemNgt(Index, InvNum, i)
        Next
        
        For i = 1 To MAX_BERRYS
            ItemBry(i) = GetPlayerInvItemBerry(Index, InvNum, i)
        Next
        
        ' Find out what kind of item it is
        Select Case Item(ItemNum).Type
            Case ITEM_TYPE_ROD
            
            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                PlayerMsg Index, "Chame o pokémon antes de pescar!", White
                Exit Sub
            End If
            
            If GetPlayerEquipment(Index, weapon) > 0 Then
                PlayerMsg Index, "Saia da bicicleta antes de pescar!", White
                Exit Sub
            End If
            
            With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index))
                If .Type = TILE_TYPE_FISHING Then
                    Call SetPlayerDir(Index, .Data1)
                    Player(Index).InFishing = 10000 + GetTickCount
                    SendPlayerXYToMap Index
                    Call SetPlayerInvItemPokeInfoLevel(Index, InvNum, GetPlayerInvItemPokeInfoLevel(Index, InvNum) + 1)
                    Player(Index).UltRodLevel = GetPlayerInvItemPokeInfoLevel(Index, InvNum)
                    Call SendInventory(Index)
                    Call SendAttack(Index)
                    Call SendInFishing(Index)
                End If
            End With
            Case ITEM_TYPE_STONE
            
            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                For i = 1 To 8
                
                If Item(ItemNum).Tipo = (Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Evolução(i).Pedra - 1) Then
                    If Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Evolução(i).Level <= GetPlayerEquipmentPokeInfoLevel(Index, weapon) Then
                        Player(Index).EvolStone = Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Evolução(i).Pokemon
                        Player(Index).EvolTimerStone = 9000 + GetTickCount
                        Player(Index).EvolPermition = 1
                        Player(Index).EvoId = i
                        SendPokeEvolution Index, 1
                        TakeInvItem Index, ItemNum, 1
                    Else
                        PlayerMsg Index, "Level Insuficiente para evoluir", White
                    End If
                End If
                
                Next
            Else
                PlayerMsg Index, "O pokémon tem que estar fora da pokéball", White
            End If
            
            Case ITEM_TYPE_CURRENCY
            
            If ItemNum = 255 Or ItemNum = 254 Or ItemNum = 253 Or ItemNum = 252 Or ItemNum = 251 Or ItemNum = 250 Then 'Pokéball
            If TempPlayer(Index).target > 0 And TempPlayer(Index).targetType = 2 Then
            
            If MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Desmaiado = True Then
            
            If Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Pokemon = 0 Then Exit Sub
            PokebolaUsada = 1
            
            Select Case ItemNum
            Case 255 'Safari Ball 40%
            
            Select Case GetPlayerMap(Index)
            Case 99, 100 'Mapas
                Chance = 40
                PokebolaUsada = 6
            Case Else
                PlayerMsg Index, "Uso exclusivo Saffari Zone!", BrightGreen
                Exit Sub
            End Select
            
            Case 254 'Master Ball 100%
                Chance = 100
                PokebolaUsada = 4
            Case 253 'Premier Ball
                Chance = 25
                PokebolaUsada = 5
            Case 252 'Super Ball 30%
                Chance = 15
                PokebolaUsada = 3
            Case 251 'Great Ball 20%
                Chance = 10
                PokebolaUsada = 2
            Case 250 'Pokéball 10%
                Chance = 5
                PokebolaUsada = 1
            End Select
            
            'Info Pokemon Catching
            TempItemPokemon = Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Pokemon 'aki é o numero .-.
            TempItemLevel = MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Level
            TempItemMaxVital(1) = Pokemon(Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Pokemon).Vital(1)
            TempItemMaxVital(1) = Random(TempItemMaxVital(1) * 75 / 100, TempItemMaxVital(1)) + (TempItemLevel + 5)
            TempItemMaxVital(2) = Pokemon(Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Pokemon).Vital(2)
            TempItemMaxVital(2) = Random(TempItemMaxVital(2) * 75 / 100, TempItemMaxVital(2)) + (TempItemLevel + 5)
            TempItemVital(1) = TempItemMaxVital(1) * 25 / 100
            TempItemVital(2) = TempItemMaxVital(2) * 25 / 100
            
            If MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Sexo = False Then
                TempItemSexo = 0
            Else
                TempItemSexo = 1
            End If
            
            If MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Shiny = False Then
                TempItemShiny = 0
            Else
                TempItemShiny = 1
            End If
            
            For i = 1 To Stats.Stat_Count - 1
                TempItemStat(i) = Pokemon(Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Pokemon).Add_Stat(i)
                TempItemStat(i) = Random(TempItemStat(i) * 75 / 100, TempItemStat(i)) + (TempItemLevel * 3)
            Next
            
            For i = 1 To MAX_POKE_SPELL
                TempItemSpell(i) = 0
            Next
            
            If Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Pokemon = 0 Then Exit Sub
            
            Chance = Chance + Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Chance
            
            If Chance >= 100 Then
            TakeInvItem Index, ItemNum, 1 'pegar pokebola vazia
            '
            If Player(Index).PokeQntia < 6 Then
                GiveInvItem Index, 3, 0, False, TempItemPokemon, PokebolaUsada, _
                TempItemLevel, 0, TempItemVital(1), TempItemVital(2), _
                TempItemMaxVital(1), TempItemMaxVital(2), TempItemStat(1), _
                TempItemStat(4), TempItemStat(2), TempItemStat(3), _
                TempItemStat(5), TempItemSpell(1), TempItemSpell(2), _
                TempItemSpell(3), TempItemSpell(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, TempItemSexo, TempItemShiny, TempItemBry(1), TempItemBry(2), TempItemBry(3), TempItemBry(4), TempItemBry(5)
                PlayerMsg Index, "Parábens! você capturou um " & Trim$(Pokemon(TempItemPokemon).Name) & "!", BrightGreen
                Player(Index).PokeQntia = Player(Index).PokeQntia + 1
            Else
                TakeInvItem Index, ItemNum, 1 'pegar pokebola vazia
                PlayerMsg Index, "Você já possui 6 pokémons, o pokémon " & Trim$(Pokemon(TempItemPokemon).Name) & " foi enviado para o computador!", White
                
                DirectBankItemPokemon Index, 3, TempItemPokemon, PokebolaUsada, _
                TempItemLevel, TempItemExp, TempItemVital(1), TempItemVital(2), _
                TempItemMaxVital(1), TempItemMaxVital(2), TempItemStat(1), _
                TempItemStat(4), TempItemStat(2), TempItemStat(3), TempItemStat(5), _
                TempItemSpell(1), TempItemSpell(2), TempItemSpell(3), _
                TempItemSpell(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, TempItemSexo, TempItemShiny, TempItemBry(1), TempItemBry(2), TempItemBry(3), TempItemBry(4), TempItemBry(5)
            End If
            
            If PokebolaUsada = 4 Then
                SendAnimation GetPlayerMap(Index), 34, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).X, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Y
            End If
            
            ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
            MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num = 0
            MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).SpawnWait = GetTickCount
            MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Vital(1) = 0
    
            ' send death to the map
            If MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Pescado = True Then
                RemoverNpcPescado MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num, GetPlayerMap(Index)
            End If
            
            SendNpcDesmaiado GetPlayerMap(Index), TempPlayer(Index).target, True
            Call SendMapNpcVitals(GetPlayerMap(Index), TempPlayer(Index).target)
            
            TempPlayer(Index).target = 0
            TempPlayer(Index).targetType = 0
            Call SendTarget(Index)
            Exit Sub
            End If
            
            R = Random(1, 100)
            If Chance >= R Then
            
            If Player(Index).PokeQntia < 6 Then
                TakeInvItem Index, ItemNum, 1 'pegar pokebola vazia
                GiveInvItem Index, 3, 0, False, TempItemPokemon, PokebolaUsada, _
                TempItemLevel, 0, TempItemVital(1), TempItemVital(2), _
                TempItemMaxVital(1), TempItemMaxVital(2), TempItemStat(1), _
                TempItemStat(4), TempItemStat(2), TempItemStat(3), _
                TempItemStat(5), TempItemSpell(1), TempItemSpell(2), _
                TempItemSpell(3), TempItemSpell(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, TempItemSexo, TempItemShiny
                PlayerMsg Index, "Parábens! você capturou um " & Trim$(Pokemon(TempItemPokemon).Name) & "!", BrightGreen
                Player(Index).PokeQntia = Player(Index).PokeQntia + 1
            Else
                TakeInvItem Index, ItemNum, 1 'pegar pokebola vazia
                PlayerMsg Index, "Você já possui 6 pokémons, o pokémon " & Trim$(Pokemon(TempItemPokemon).Name) & " foi enviado para o computador!", White
                
                DirectBankItemPokemon Index, 3, TempItemPokemon, PokebolaUsada, _
                TempItemLevel, TempItemExp, TempItemVital(1), TempItemVital(2), _
                TempItemMaxVital(1), TempItemMaxVital(2), TempItemStat(1), _
                TempItemStat(4), TempItemStat(2), TempItemStat(3), TempItemStat(5), _
                TempItemSpell(1), TempItemSpell(2), TempItemSpell(3), _
                TempItemSpell(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, TempItemSexo, TempItemShiny, TempItemBry(1), TempItemBry(2), TempItemBry(3), TempItemBry(4), TempItemBry(5)
            End If
            
            Select Case PokebolaUsada
                Case 2
                    SendAnimation GetPlayerMap(Index), 29, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).X, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Y
                Case 3
                    SendAnimation GetPlayerMap(Index), 31, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).X, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Y
                Case Else
                    SendAnimation GetPlayerMap(Index), 27, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).X, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Y
                End Select
            Else
                Select Case PokebolaUsada
                Case 2
                    SendAnimation GetPlayerMap(Index), 28, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).X, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Y
                Case 3
                    SendAnimation GetPlayerMap(Index), 30, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).X, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Y
                Case Else
                    SendAnimation GetPlayerMap(Index), 26, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).X, MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Y
                End Select
                PlayerMsg Index, "O " & Trim$(Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Name) & " Escapou!", BrightRed
                TakeInvItem Index, ItemNum, 1 'pegar pokebola vazia
            End If
            
            ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
            MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num = 0
            MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).SpawnWait = GetTickCount
            MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Vital(1) = 0
    
            ' send death to the map
            If MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Pescado = True Then
                RemoverNpcPescado TempPlayer(Index).target, GetPlayerMap(Index)
            End If
            
            SendNpcDesmaiado GetPlayerMap(Index), TempPlayer(Index).target, True
            Call SendMapNpcVitals(GetPlayerMap(Index), TempPlayer(Index).target)
            
            TempPlayer(Index).target = 0
            TempPlayer(Index).targetType = 0
            Call SendTarget(Index)
            Else
            PlayerMsg Index, "O Pokémon alvo tem que estar exausto!", BrightGreen
            End If
            End If
            End If
            
            Case ITEM_TYPE_ARMOR
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Armor)
                End If

                SetPlayerEquipment Index, ItemNum, Armor
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 0

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                'Call SendStats(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_WEAPON
            
                'Limpar Timer caso Chame o Pokémon
                If ItemNum = 3 Then 'Pokeball
                    If TempPlayer(Index).SwitPoke > 0 Then
                        TempPlayer(Index).SwitPoke = 0
                    End If
                End If
                
                If Player(Index).InSurf = 1 Then
                    PlayerMsg Index, "Saia d'Agua antes de chamar o pokémon!", BrightRed
                    Exit Sub
                End If
                
                If GetPlayerEquipment(Index, weapon) > 0 Then
                
                'Bloquear troca de Pokémon
                If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_ARENA Then
                    If TempPlayer(Index).Lutando > 0 Then
                        PlayerMsg Index, "Não pode trocar de Pokémon na batalha", BrightRed
                        Exit Sub
                    End If
                End If
                
                If Not ItemNum = 3 Then
                    PlayerMsg Index, "Chame o Pokémon antes de subir na Bicicleta!", BrightRed
                    Exit Sub
                End If
                
                If Not GetPlayerEquipment(Index, weapon) = 3 Then
                    PlayerMsg Index, "Saia da Bicicleta antes de Chamar o Pokémon!", BrightRed
                    Exit Sub
                End If
                
                End If
                
                If ItemPokemon > 0 Then
                    If ItemVital(1) <= 0 Then
                        PlayerMsg Index, "Este pokémon está ferido!", BrightRed
                        Exit Sub
                    End If
                End If
                
                SetPlayerFlying Index, 0

                If GetPlayerEquipment(Index, weapon) > 0 Then
                
                    tempItem = GetPlayerEquipment(Index, weapon)
                    TempItemPokemon = GetPlayerEquipmentPokeInfoPokemon(Index, weapon)
                    TempItemPokeball = GetPlayerEquipmentPokeInfoPokeball(Index, weapon)
                    TempItemLevel = GetPlayerEquipmentPokeInfoLevel(Index, weapon)
                    TempItemExp = GetPlayerEquipmentPokeInfoExp(Index, weapon)
                    TempItemFelicidade = GetPlayerEquipmentFelicidade(Index, weapon)
                    TempItemSexo = GetPlayerEquipmentSexo(Index, weapon)
                    TempItemShiny = GetPlayerEquipmentShiny(Index, weapon)
                    
                    For i = 1 To Vitals.Vital_Count - 1
                        TempItemVital(i) = GetPlayerEquipmentPokeInfoVital(Index, weapon, i)
                        TempItemMaxVital(i) = GetPlayerEquipmentPokeInfoMaxVital(Index, weapon, i)
                    Next
                    
                    For i = 1 To Stats.Stat_Count - 1
                        TempItemStat(i) = GetPlayerEquipmentPokeInfoStat(Index, weapon, i)
                    Next
                    
                    For i = 1 To MAX_POKE_SPELL
                        TempItemSpell(i) = GetPlayerEquipmentPokeInfoSpell(Index, weapon, i)
                    Next
                    
                    For i = 1 To MAX_NEGATIVES
                        TempItemNgt(i) = GetPlayerEquipmentNgt(Index, weapon, i)
                    Next
                    
                    For i = 1 To MAX_BERRYS
                        TempItemBry(i) = GetPlayerEquipmentBerry(Index, weapon, i)
                    Next
                
                End If

                '#####################Setar Informações#####################################
                Player(Index).VitalTemp = GetPlayerVital(Index, HP) 'Hp do Jogador
                SetPlayerEquipment Index, ItemNum, weapon
                SetPlayerEquipmentPokeInfoPokemon Index, ItemPokemon, weapon
                SetPlayerEquipmentPokeInfoPokeball Index, ItemPokeball, weapon
                SetPlayerEquipmentPokeInfoLevel Index, ItemLevel, weapon
                SetPlayerEquipmentPokeInfoExp Index, ItemExp, weapon
                
                For i = 1 To Vitals.Vital_Count - 1
                    SetPlayerEquipmentPokeInfoVital Index, ItemVital(i), weapon, i
                    SetPlayerEquipmentPokeInfoMaxVital Index, ItemMaxVital(i), weapon, i
                Next
                
                If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                    SetPlayerVital Index, HP, GetPlayerEquipmentPokeInfoVital(Index, weapon, 1)
                    SetPlayerVital Index, MP, GetPlayerEquipmentPokeInfoVital(Index, weapon, 2)
                End If
                
                For i = 1 To Stats.Stat_Count - 1
                    SetPlayerEquipmentPokeInfoStat Index, ItemStat(i), weapon, i
                Next
                
                For i = 1 To MAX_POKE_SPELL
                    SetPlayerEquipmentPokeInfoSpell Index, ItemSpell(i), weapon, i
                Next
                
                For i = 1 To MAX_NEGATIVES
                    SetPlayerEquipmentNgt Index, i, weapon, ItemNgt(i)
                Next
                
                For i = 1 To MAX_BERRYS
                    SetPlayerEquipmentBerry Index, ItemBry(i), weapon, i
                Next
                
                SetPlayerEquipmentFelicidade Index, weapon, ItemFelicidade
                SetPlayerEquipmentSexo Index, weapon, ItemSexo
                SetPlayerEquipmentShiny Index, weapon, ItemShiny
                
                TakeInvSlot Index, InvNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0, False, TempItemPokemon, _
                    TempItemPokeball, TempItemLevel, TempItemExp, _
                    TempItemVital(1), TempItemVital(2), TempItemMaxVital(1), _
                    TempItemMaxVital(2), TempItemStat(1), TempItemStat(4), _
                    TempItemStat(2), TempItemStat(3), TempItemStat(5), _
                    TempItemSpell(1), TempItemSpell(2), TempItemSpell(3), _
                    TempItemSpell(4), TempItemNgt(1), TempItemNgt(2), TempItemNgt(3), TempItemNgt(4), TempItemNgt(5), TempItemNgt(6), _
                    TempItemNgt(7), TempItemNgt(8), TempItemNgt(9), TempItemNgt(10), TempItemNgt(11), TempItemFelicidade, TempItemSexo, TempItemShiny, _
                    TempItemBry(1), TempItemBry(2), TempItemBry(3), TempItemBry(4), TempItemBry(5)
                End If
                
                If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                'Pokémon
                If Player(Index).Sprite = Player(Index).MySprite Then
                    
                    Player(Index).TPDir = GetPlayerDir(Index)
                    Player(Index).TPX = GetPlayerX(Index)
                    Player(Index).TPY = GetPlayerY(Index)
                    Player(Index).TPSprite = GetPlayerSprite(Index)
                    
                    SetPlayerPokePosition Index
                    
                    Select Case GetPlayerEquipmentPokeInfoPokeball(Index, weapon)
                    Case 2
                        SendAnimation GetPlayerMap(Index), 32, GetPlayerX(Index), GetPlayerY(Index)
                    Case 3
                        SendAnimation GetPlayerMap(Index), 37, GetPlayerX(Index), GetPlayerY(Index)
                    Case 4
                        SendAnimation GetPlayerMap(Index), 35, GetPlayerX(Index), GetPlayerY(Index)
                    Case Else
                        SendAnimation GetPlayerMap(Index), 7, GetPlayerX(Index), GetPlayerY(Index)
                    End Select
                    
                    SendActionMsg GetPlayerMap(Index), "Vai " & Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " eu escolho você!", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                Else
                    Player(Index).X = Player(Index).TPX
                    Player(Index).Y = Player(Index).TPY
                    
                    SetPlayerPokePosition Index
                    
                    SendPlayerXY Index
                    SendAnimation GetPlayerMap(Index), 7, GetPlayerX(Index), GetPlayerY(Index)
                    SendActionMsg GetPlayerMap(Index), "vai " & Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " eu escolho você!", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                End If
                    If GetPlayerEquipmentShiny(Index, weapon) = 0 Then
                        Player(Index).Sprite = Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Sprite
                    Else
                        Player(Index).Sprite = Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Sprite + 1
                    End If
                    Call SendPlayerData(Index)
                End If
                
                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                Call SendInventory(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                SendEXP Index
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
                
                'Limpar Spells
                For i = 1 To MAX_PLAYER_SPELLS
                    Call SetPlayerSpell(Index, i, 0)
                Next
                
                'Setar Spell do Pokémon no Inv Spell :)!
                For i = 1 To MAX_POKE_SPELL
                    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                        If GetPlayerEquipmentPokeInfoSpell(Index, weapon, i) > 0 Then
                        Call SetPlayerSpell(Index, i, GetPlayerEquipmentPokeInfoSpell(Index, weapon, i))
                        End If
                    End If
                Next
                SendSpells Index
                
                For i = 1 To MAX_NEGATIVES
                    If GetPlayerEquipment(Index, weapon) > 0 Then
                        If GetPlayerEquipmentNgt(Index, weapon, i) > 0 Then
                            TempPlayer(Index).NgtTick(i) = 1 + GetTickCount
                        End If
                    End If
                Next
                
                Case ITEM_TYPE_HELMET
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Helmet)
                End If

                SetPlayerEquipment Index, ItemNum, Helmet
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 1
                
                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_SHIELD
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Shield)
                End If

                SetPlayerEquipment Index, ItemNum, Shield
                PlayerMsg Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                'Call SendStats(Index)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(ItemNum).AddHP > 0 Then
                    Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + Item(ItemNum).AddHP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, HP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                
                ' add mp
                If Item(ItemNum).AddMP > 0 Then
                    Player(Index).Vital(Vitals.MP) = Player(Index).Vital(Vitals.MP) + Item(ItemNum).AddMP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, MP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                
                ' add exp
                If Item(ItemNum).AddEXP > 0 Then
                    SetPlayerExp Index, GetPlayerExp(Index) + Item(ItemNum).AddEXP
                    CheckPlayerLevelUp Index
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendEXP Index
                End If
                
                'Se tiver com Pokémon em Mãos...
                If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                    Dim StatName(1 To Stats.Stat_Count - 1) As String
                    Dim StatVazio As Byte
                    
                    StatName(1) = " Atq"
                    StatName(2) = " Def"
                    StatName(3) = " SP.ATQ"
                    StatName(4) = " Speed"
                    StatName(5) = " SP.DEF"
                    
                    Select Case Item(ItemNum).Berry
                        Case 1
                            For i = 1 To MAX_BERRYS
                                If Item(ItemNum).Add_Stat(i) = 0 Then
                                    StatVazio = StatVazio + 1
                                Else
                                    If Item(ItemNum).YesNo(i) = True Then
                                            Call SetPlayerEquipmentBerry(Index, GetPlayerEquipmentBerry(Index, weapon, i) + Item(ItemNum).Add_Stat(i), weapon, i)
                                            SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).Add_Stat(i) & StatName(i), BrightGreen, 0, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32) - ((i - StatVazio) * 15)
                                        Else
                                            Call SetPlayerEquipmentBerry(Index, GetPlayerEquipmentBerry(Index, weapon, i) - Item(ItemNum).Add_Stat(i), weapon, i)
                                            SendActionMsg GetPlayerMap(Index), "-" & Item(ItemNum).Add_Stat(i) & StatName(i), BrightGreen, 0, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32) - ((i - StatVazio) * 15)
                                    End If
                                End If
                            Next
                        Case 2
                            Call SetPlayerEquipmentFelicidade(Index, weapon, GetPlayerEquipmentFelicidade(Index, weapon) + Item(ItemNum).Add_Stat(1))
                            SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).Add_Stat(1) & "Felicidade", BrightGreen, 0, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32) - ((1 - StatVazio) * 15)
                        Case 3
                            
                            'TempPlayer(Index).NgtTick(5) = 1000 + GetTickCount
                            'SetPlayerEquipmentNgt Index, 5, weapon, 60
                            'SendMapEquipment Index
                            
                            TempPlayer(Index).NgtTick(6) = 1000 + GetTickCount
                            SetPlayerEquipmentNgt Index, 6, weapon, 60
                            SendMapEquipment Index
                        Case 4
                            SetPlayerHonra Index, GetPlayerHonra(Index) + 50
                            PlayerMsg Index, GetPlayerHonra(Index), White
                            SendPlayerData Index
                        Case 5
                        If GetPlayerEquipment(Index, weapon) > 0 Then
                            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                                GivePlayerEXP Index, GetPlayerInvItemPokeInfoExp(Index, InvNum)
                                PlayerMsg Index, "Seu pokémon ganhou " & GetPlayerInvItemPokeInfoExp(Index, InvNum) & " de EXP!", BrightGreen
                                Call TakeInvSlot(Index, InvNum, 0)
                                Call SendInventory(Index)
                    
                                ' send the sound
                                Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
                                Exit Sub
                            End If
                        End If
                    End Select
                Else
                    PlayerMsg Index, "Berrys só podem ser entregues ao pokémon fora da pokéball", BrightRed
                    Exit Sub
                End If
                
                Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call TakeInvItem(Index, Player(Index).Inv(InvNum).Num, 0)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_KEY
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(Index)
                    Case DIR_UP

                        If GetPlayerY(Index) > 0 Then
                            X = GetPlayerX(Index)
                            Y = GetPlayerY(Index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(Index) < Map(GetPlayerMap(Index)).MaxY Then
                            X = GetPlayerX(Index)
                            Y = GetPlayerY(Index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(Index) > 0 Then
                            X = GetPlayerX(Index) - 1
                            Y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(Index) < Map(GetPlayerMap(Index)).MaxX Then
                            X = GetPlayerX(Index) + 1
                            Y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If ItemNum = Map(GetPlayerMap(Index)).Tile(X, Y).Data1 Then
                        TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        SendMapKey Index, X, Y, 1
                        Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, X, Y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(Index)).Tile(X, Y).Data2 = 1 Then
                            Call TakeInvItem(Index, ItemNum, 0)
                            Call PlayerMsg(Index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            
            Case ITEM_TYPE_SPELL
                Dim SlotVazio As Boolean, SpellLearnTo As Byte
                ' Get the spell num
                n = Item(ItemNum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(Index) Then
                            i = FindOpenSpellSlot(Index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(Index, n) Then
                                    Call SetPlayerSpell(Index, i, n)
                                    Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                                    If GetPlayerEquipment(Index, weapon) > 0 Then
                                        If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                                        
                                            For SpellLearnTo = 1 To MAX_POKE_SPELL
                                                If GetPlayerEquipmentPokeInfoSpell(Index, weapon, SpellLearnTo) = 0 Then
                                                    Call SetPlayerEquipmentPokeInfoSpell(Index, n, weapon, SpellLearnTo)
                                                    SetPlayerSpell Index, SpellLearnTo, n
                                                    SendPlayerSpells Index
                                                    Call SendWornEquipment(Index)
                                                    Call SendMapEquipment(Index)
                                                    PlayerMsg Index, "O pokémon " & Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " aprendeu a habilidade " & Trim$(Spell(n).Name) & ".", BrightGreen
                                                    SlotVazio = True
                                                    Exit For
                                                End If
                                            Next
                                            
                                            If SlotVazio = False Then
                                                Player(Index).LearnSpell(1) = 1
                                                Player(Index).LearnSpell(2) = n
                                                Player(Index).LearnSpell(3) = ItemNum
                                                SendAprenderSpell Index, 0
                                            End If
                                            
                                        End If
                                    End If
                                Else
                                    Call PlayerMsg(Index, "O Pokémon já aprendeu esse Ataque.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(Index, "Pokémon não pode aprender mais ataques.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(Index, "Sem level necessario para aprender " & i & ".", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(Index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            
             Case ITEM_TYPE_UP
                
                If GetPlayerEquipment(Index, weapon) < 1 Then
                        PlayerMsg Index, "Precisa de um pokémon fora da pokéball!", White
                        Exit Sub
                    Else
                        If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                            SetPlayerExp Index, GetPlayerNextLevel(Index)
                            CheckPlayerLevelUp Index
                            Call SetPlayerEquipmentFelicidade(Index, weapon, GetPlayerEquipmentFelicidade(Index, weapon) + 3)
                            TakeInvItem Index, ItemNum, 1
                        End If
                End If
                
            Case ITEM_TYPE_POKEDEX
        
            If TempPlayer(Index).ScanTime > 0 Then Exit Sub

                    If TempPlayer(Index).target > 0 Then
                        If TempPlayer(Index).targetType = TARGET_TYPE_NPC Then
                            If Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Pokemon > 0 Then 'verifica se a variavel pokemon é maior q 0 do editor
                            
                            If Player(Index).Pokedex(Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Pokemon) = 0 Then 'essa porra é onde salva o pokemo
                                TempPlayer(Index).ScanPokemon = Npc(MapNpc(GetPlayerMap(Index)).Npc(TempPlayer(Index).target).Num).Pokemon
                                TempPlayer(Index).ScanTime = 2000 + GetTickCount
                                SendActionMsg GetPlayerMap(Index), "Scaneando Pokémon.", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                                SendInFishing Index
                            Else
                                PlayerMsg Index, "Você já possui a informações deste pokémon!", Red
                            End If
                
                            End If
                        Else
                    If GetPlayerEquipmentPokeInfoPokemon(TempPlayer(Index).target, weapon) > 0 Then
                            If Player(Index).Pokedex(GetPlayerEquipmentPokeInfoPokemon(TempPlayer(Index).target, weapon)) = 0 Then
                                TempPlayer(Index).ScanPokemon = GetPlayerEquipmentPokeInfoPokemon(TempPlayer(Index).target, weapon)
                                TempPlayer(Index).ScanTime = 2000 + GetTickCount
                                SendActionMsg GetPlayerMap(Index), "Scaneando Pokémon.", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                                SendInFishing Index
                            Else
                                PlayerMsg Index, "Você já possui a informações deste pokémon!", Red
                            End If
                        End If
                        End If
                    End If
                    
                Case ITEM_TYPE_BAU
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                'Calcular Espaço
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(Index, i) = 0 Then
                        InvVazio = InvVazio + 1
                    End If
                Next
                
                If Item(ItemNum).GiveAll = False Then
                    'Verificar Espaço na Mochila
                    If InvVazio = 0 Then
                        PlayerMsg Index, "Sem espaço suficiente para abrir o Báu/Sacola/Pacote!", BrightRed
                        Exit Sub
                    End If
                    
                    'Calcular QntiaDeItensNoBau
                    For i = 1 To MAX_BAU
                        If Item(ItemNum).BauItem(i) = 0 Then
                            Exit For
                        End If
                    Next
                    
                    'Entregar o Item!
                    If i - 1 > 0 Then
                        n = RAND(1, i - 1)
                        Call GiveInvItem(Index, Item(ItemNum).BauItem(n), Item(ItemNum).BauValue(n))
                        Call TakeInvItem(Index, ItemNum, 0)
                    End If
                Else
                    'Calcular QntiaDeItensNoBau
                    For i = 1 To MAX_BAU
                        If Item(ItemNum).BauItem(i) > 0 Then
                            BauItemQnt = BauItemQnt + 1
                        End If
                    Next
                
                    'Verificar Espaço na Mochila
                    If InvVazio < BauItemQnt Then
                        PlayerMsg Index, "Sem espaço suficiente para abrir o Báu/Sacola/Pacote!", BrightRed
                        Exit Sub
                    End If
                
                    'Entregar os Itens!
                    For i = 1 To MAX_BAU
                        If Item(ItemNum).BauItem(i) > 0 Then
                            Call GiveInvItem(Index, Item(ItemNum).BauItem(i), Item(ItemNum).BauValue(i))
                        End If
                    Next
                    
                    'Retirar o Item
                    Call TakeInvItem(Index, ItemNum, 0)
                End If
                
        End Select
    End If
End Sub

'######################## Pokémon Info ################################

'PokeInfo Pokemon
Function GetPlayerInvItemPokeInfoPokemon(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemPokeInfoPokemon = Player(Index).Inv(invslot).PokeInfo.Pokemon
End Function

Sub SetPlayerInvItemPokeInfoPokemon(ByVal Index As Long, ByVal invslot As Long, ByVal PokeNum As Long)
    Player(Index).Inv(invslot).PokeInfo.Pokemon = PokeNum
End Sub

'PokeInfo Pokéball
Function GetPlayerInvItemPokeInfoPokeball(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemPokeInfoPokeball = Player(Index).Inv(invslot).PokeInfo.Pokeball
End Function

Sub SetPlayerInvItemPokeInfoPokeball(ByVal Index As Long, ByVal invslot As Long, ByVal PokeballNum As Long)
    Player(Index).Inv(invslot).PokeInfo.Pokeball = PokeballNum
End Sub

'PokeInfo Level
Function GetPlayerInvItemPokeInfoLevel(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemPokeInfoLevel = Player(Index).Inv(invslot).PokeInfo.Level
End Function

Sub SetPlayerInvItemPokeInfoLevel(ByVal Index As Long, ByVal invslot As Long, ByVal Level As Long)
    Player(Index).Inv(invslot).PokeInfo.Level = Level
End Sub

'PokeInfo Exp
Function GetPlayerInvItemPokeInfoExp(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemPokeInfoExp = Player(Index).Inv(invslot).PokeInfo.EXP
End Function

Sub SetPlayerInvItemPokeInfoExp(ByVal Index As Long, ByVal invslot As Long, ByVal EXP As Long)
    Player(Index).Inv(invslot).PokeInfo.EXP = EXP
End Sub

'PokeInfo MaxVital
Function GetPlayerInvItemPokeInfoMaxVital(ByVal Index As Long, ByVal invslot As Long, ByVal VitalType As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemPokeInfoMaxVital = Player(Index).Inv(invslot).PokeInfo.MaxVital(VitalType)
End Function

Sub SetPlayerInvItemPokeInfoMaxVital(ByVal Index As Long, ByVal invslot As Long, ByVal MaxVital As Long, ByVal VitalType As Long)
    If VitalType = 0 Or VitalType > Vitals.Vital_Count - 1 Then Exit Sub
    Player(Index).Inv(invslot).PokeInfo.MaxVital(VitalType) = MaxVital
End Sub

'PokeInfo Vital
Function GetPlayerInvItemPokeInfoVital(ByVal Index As Long, ByVal invslot As Long, ByVal VitalType As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemPokeInfoVital = Player(Index).Inv(invslot).PokeInfo.Vital(VitalType)
End Function

Sub SetPlayerInvItemPokeInfoVital(ByVal Index As Long, ByVal invslot As Long, ByVal Vital As Long, ByVal VitalType As Long)
    If VitalType = 0 Or VitalType > Vitals.Vital_Count - 1 Then Exit Sub
    Player(Index).Inv(invslot).PokeInfo.Vital(VitalType) = Vital
End Sub

'PokeInfo Spells
Function GetPlayerInvItemPokeInfoSpell(ByVal Index As Long, ByVal invslot As Long, ByVal spellslot As Byte) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemPokeInfoSpell = Player(Index).Inv(invslot).PokeInfo.Spells(spellslot)
End Function

Sub SetPlayerInvItemPokeInfoSpell(ByVal Index As Long, ByVal invslot As Long, ByVal SpellNum As Long, ByVal spellslot As Long)
    If spellslot > 4 Then Exit Sub
    Player(Index).Inv(invslot).PokeInfo.Spells(spellslot) = SpellNum
End Sub

'PokeInfo Stat Point
Function GetPlayerInvItemPokeInfoStat(ByVal Index As Long, ByVal invslot As Long, ByVal StatNum As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemPokeInfoStat = Player(Index).Inv(invslot).PokeInfo.Stat(StatNum)
End Function

Sub SetPlayerInvItemPokeInfoStat(ByVal Index As Long, ByVal invslot As Long, ByVal StatNum As Long, ByVal StatValue As Long)
    If StatNum = 0 Or StatNum > Stats.Stat_Count - 1 Then Exit Sub
    Player(Index).Inv(invslot).PokeInfo.Stat(StatNum) = StatValue
End Sub

'PokeInfo Stats negativo
Function GetPlayerInvItemNgt(ByVal Index As Long, ByVal invslot As Long, ByVal NgtNum As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemNgt = Player(Index).Inv(invslot).PokeInfo.Negatives(NgtNum)
End Function

Sub SetPlayerInvItemNgt(ByVal Index As Long, ByVal invslot As Long, ByVal NgtNum As Long, ByVal NgtValue As Long)
    If invslot = 0 Or NgtNum = 0 Then Exit Sub
    Player(Index).Inv(invslot).PokeInfo.Negatives(NgtNum) = NgtValue
End Sub

'PokeInfo Felicidade
Function GetPlayerInvItemFelicidade(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemFelicidade = Player(Index).Inv(invslot).PokeInfo.Felicidade
End Function

Sub SetPlayerInvItemFelicidade(ByVal Index As Long, ByVal invslot As Long, ByVal Felicidade As Long)
    If invslot = 0 Then Exit Sub
    If Felicidade > 500 Then Felicidade = 500
    If Felicidade < 0 Then Felicidade = 0
    
    Player(Index).Inv(invslot).PokeInfo.Felicidade = Felicidade
End Sub

'PokeInfo Sexo
Function GetPlayerInvItemSexo(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemSexo = Player(Index).Inv(invslot).PokeInfo.Sexo
End Function

Sub SetPlayerInvItemSexo(ByVal Index As Long, ByVal invslot As Long, ByVal Sexo As Long)
    If invslot = 0 Then Exit Sub
    Player(Index).Inv(invslot).PokeInfo.Sexo = Sexo
End Sub

'PokeInfo Shiny
Function GetPlayerInvItemShiny(ByVal Index As Long, ByVal invslot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemShiny = Player(Index).Inv(invslot).PokeInfo.Shiny
End Function

Sub SetPlayerInvItemShiny(ByVal Index As Long, ByVal invslot As Long, ByVal Shiny As Long)
    If invslot = 0 Then Exit Sub
    Player(Index).Inv(invslot).PokeInfo.Shiny = Shiny
End Sub

'PokeInfo Berry
Function GetPlayerInvItemBerry(ByVal Index As Long, ByVal invslot As Long, ByVal BerryStat As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    
    GetPlayerInvItemBerry = Player(Index).Inv(invslot).PokeInfo.Berry(BerryStat)
End Function

Sub SetPlayerInvItemBerry(ByVal Index As Long, ByVal invslot As Long, ByVal BerryStat As Long, ByVal Valor As Long)
    If BerryStat = 0 Or BerryStat > MAX_BERRYS Then Exit Sub
    If Valor > 30 Then Valor = 30
    If Valor <= 0 Then Valor = 0
    Player(Index).Inv(invslot).PokeInfo.Berry(BerryStat) = Valor
End Sub

'############################Poke Info Equipment##################################################

Function GetPlayerEquipmentPokeInfoPokemon(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
    If Index > MAX_PLAYERS Or Index = 0 Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoPokemon = Player(Index).EquipPokeInfo(EquipmentSlot).Pokemon
End Function

Sub SetPlayerEquipmentPokeInfoPokemon(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).EquipPokeInfo(EquipmentSlot).Pokemon = InvNum
End Sub

'PokeInfo Pokeball
Function GetPlayerEquipmentPokeInfoPokeball(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoPokeball = Player(Index).EquipPokeInfo(EquipmentSlot).Pokeball
End Function

Sub SetPlayerEquipmentPokeInfoPokeball(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).EquipPokeInfo(EquipmentSlot).Pokeball = InvNum
End Sub

'PokeInfo Level
Function GetPlayerEquipmentPokeInfoLevel(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoLevel = Player(Index).EquipPokeInfo(EquipmentSlot).Level
End Function

Sub SetPlayerEquipmentPokeInfoLevel(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).EquipPokeInfo(EquipmentSlot).Level = InvNum
End Sub

'PokeInfo Exp
Function GetPlayerEquipmentPokeInfoExp(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoExp = Player(Index).EquipPokeInfo(EquipmentSlot).EXP
End Function

Sub SetPlayerEquipmentPokeInfoExp(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).EquipPokeInfo(EquipmentSlot).EXP = InvNum
End Sub

'PokeInfo Vitals
Function GetPlayerEquipmentPokeInfoVital(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal VitalType As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoVital = Player(Index).EquipPokeInfo(EquipmentSlot).Vital(VitalType)
End Function

Sub SetPlayerEquipmentPokeInfoVital(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment, ByVal VitalType As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).Vital(VitalType) = InvNum
End Sub

'PokeInfo MaxVital
Function GetPlayerEquipmentPokeInfoMaxVital(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal VitalType As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoMaxVital = Player(Index).EquipPokeInfo(EquipmentSlot).MaxVital(VitalType)
End Function

Sub SetPlayerEquipmentPokeInfoMaxVital(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment, ByVal VitalType As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).MaxVital(VitalType) = InvNum
End Sub

'PokeInfo Stat
Function GetPlayerEquipmentPokeInfoStat(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal StatNum As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoStat = Player(Index).EquipPokeInfo(EquipmentSlot).Stat(StatNum)
End Function

Sub SetPlayerEquipmentPokeInfoStat(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment, ByVal StatNum As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).Stat(StatNum) = InvNum
End Sub

'PokeInfo Spells
Function GetPlayerEquipmentPokeInfoSpell(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal SpellNum As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentPokeInfoSpell = Player(Index).EquipPokeInfo(EquipmentSlot).Spells(SpellNum)
End Function

Sub SetPlayerEquipmentPokeInfoSpell(ByVal Index As Long, ByVal SpellNum As Long, ByVal EquipmentSlot As Equipment, ByVal spellslot As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).Spells(spellslot) = SpellNum
End Sub

'PokeInfo Negatives
Function GetPlayerEquipmentNgt(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal NgtNum As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentNgt = Player(Index).EquipPokeInfo(EquipmentSlot).Negatives(NgtNum)
End Function

Sub SetPlayerEquipmentNgt(ByVal Index As Long, ByVal NgtNum As Long, ByVal EquipmentSlot As Equipment, ByVal NgtValue As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).Negatives(NgtNum) = NgtValue
End Sub

'PokeInfo Sexo
Function GetPlayerEquipmentSexo(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentSexo = Player(Index).EquipPokeInfo(EquipmentSlot).Sexo
End Function

Sub SetPlayerEquipmentSexo(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal Sexo As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).Sexo = Sexo
End Sub

'PokeInfo Felicidade
Function GetPlayerEquipmentFelicidade(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentFelicidade = Player(Index).EquipPokeInfo(EquipmentSlot).Felicidade
End Function

Sub SetPlayerEquipmentFelicidade(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal Felicidade As Long)
    If Felicidade > 500 Then Felicidade = 500
    If Felicidade < 0 Then Felicidade = 0
    Player(Index).EquipPokeInfo(EquipmentSlot).Felicidade = Felicidade
End Sub

'PokeInfo Shiny
Function GetPlayerEquipmentShiny(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentShiny = Player(Index).EquipPokeInfo(EquipmentSlot).Shiny
End Function

Sub SetPlayerEquipmentShiny(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal Shiny As Long)
    Player(Index).EquipPokeInfo(EquipmentSlot).Shiny = Shiny
End Sub

'PokeInfo Berry
Function GetPlayerEquipmentBerry(ByVal Index As Long, ByVal EquipmentSlot As Equipment, ByVal BerryNum As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentBerry = Player(Index).EquipPokeInfo(EquipmentSlot).Berry(BerryNum)
End Function

Sub SetPlayerEquipmentBerry(ByVal Index As Long, ByVal Value As Long, ByVal EquipmentSlot As Equipment, ByVal BerryNum As Long)
    If BerryNum = 0 Or BerryNum > MAX_BERRYS Then Exit Sub
    If Value > 30 Then Value = 30
    If Value <= 0 Then Value = 0
    Player(Index).EquipPokeInfo(EquipmentSlot).Berry(BerryNum) = Value
End Sub

Sub SetPlayerPokePosition(ByVal Index As Long)
Dim X As Long, Y As Long, MapNum As Long
Dim n As Long, N2 As Long, N3 As Long, N4 As Long

X = GetPlayerX(Index)
Y = GetPlayerY(Index)
MapNum = GetPlayerMap(Index)

Select Case GetPlayerDir(Index)
Case DIR_UP

If PokeTileIsOpen(MapNum, X, Y - 1) = True Then
SetPlayerY Index, Y - 1
ElseIf PokeTileIsOpen(MapNum, X, Y + 1) = True Then
SetPlayerY Index, Y + 1
ElseIf PokeTileIsOpen(MapNum, X + 1, Y) = True Then
SetPlayerX Index, X + 1
ElseIf PokeTileIsOpen(MapNum, X - 1, Y) = True Then
SetPlayerX Index, X - 1
End If

Case DIR_DOWN

If PokeTileIsOpen(MapNum, X, Y + 1) = True Then
SetPlayerY Index, Y + 1
ElseIf PokeTileIsOpen(MapNum, X, Y - 1) = True Then
SetPlayerY Index, Y - 1
ElseIf PokeTileIsOpen(MapNum, X + 1, Y) = True Then
SetPlayerX Index, X + 1
ElseIf PokeTileIsOpen(MapNum, X - 1, Y) = True Then
SetPlayerX Index, X - 1
End If

Case DIR_LEFT

If PokeTileIsOpen(MapNum, X - 1, Y) = True Then
SetPlayerX Index, X - 1
ElseIf PokeTileIsOpen(MapNum, X + 1, Y) = True Then
SetPlayerX Index, X + 1
ElseIf PokeTileIsOpen(MapNum, X, Y + 1) = True Then
SetPlayerY Index, Y + 1
ElseIf PokeTileIsOpen(MapNum, X, Y - 1) = True Then
SetPlayerY Index, Y - 1
End If

Case DIR_RIGHT

If PokeTileIsOpen(MapNum, X + 1, Y) = True Then
SetPlayerX Index, X + 1
ElseIf PokeTileIsOpen(MapNum, X - 1, Y) = True Then
SetPlayerX Index, X - 1
ElseIf PokeTileIsOpen(MapNum, X, Y + 1) = True Then
SetPlayerY Index, Y + 1
ElseIf PokeTileIsOpen(MapNum, X, Y - 1) = True Then
SetPlayerY Index, Y - 1
End If


End Select
End Sub

Sub BlockPlayer(ByVal Index As Long)
Dim Dir As Long
Dir = GetPlayerDir(Index)
Select Case Dir
Case 0
Call SetPlayerY(Index, GetPlayerY(Index) + 1)
Case 1
Call SetPlayerY(Index, GetPlayerY(Index) - 1)
Case 2
Call SetPlayerX(Index, GetPlayerX(Index) + 1)
Case 3
Call SetPlayerX(Index, GetPlayerX(Index) - 1)
End Select
Call SendPlayerXY(Index)
End Sub

Function GetPlayerFlying(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerFlying = Player(Index).Flying
End Function

Sub SetPlayerFlying(ByVal Index As Long, ByVal Flying As Long)
    Player(Index).Flying = Flying
End Sub

Function CanBlockedTile(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal FlyDown As Boolean = False) As Boolean
    ' Verifica sobrecarga
    If IsPlaying(Index) = False Then Exit Function
    
    CanBlockedTile = False

    ' Verifica se o jogador está voando
    If FlyDown = False Then
        If GetPlayerFlying(Index) Then
            CanBlockedTile = False
            Exit Function
        End If
    End If
        
    ' Check to make sure that the tile is walkable
    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then CanBlockedTile = True
    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then CanBlockedTile = True
    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY Then CanBlockedTile = True
    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES Then CanBlockedTile = True
End Function

Public Function PokDispBattle(ByVal Index, ByVal Numero As Long) As Boolean
Dim i As Long, QntiaVivo As Byte

For i = 1 To MAX_INV
    If GetPlayerInvItemPokeInfoPokemon(Index, i) > 0 Then
        If GetPlayerInvItemPokeInfoVital(Index, i, 1) > 0 Then
            QntiaVivo = QntiaVivo + 1
        End If
    End If
Next

If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
    QntiaVivo = QntiaVivo + 1
End If

If QntiaVivo >= Numero Then
    PokDispBattle = True
End If

End Function

Sub CheckAORGlevelUP(ByVal Index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    'Evitar OverFlow!
    If Player(Index).ORG = 0 Then Exit Sub
    
     Do While Organization(Player(Index).ORG).EXP >= GetONextLevel(Index)
    
    expRollover = Organization(Player(Index).ORG).EXP - GetONextLevel(Index)
        
    Organization(Player(Index).ORG).Level = Organization(Player(Index).ORG).Level + 1
    Organization(Player(Index).ORG).EXP = expRollover
    level_count = level_count + 1
    Loop
    
    For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
    If Player(i).ORG = Player(Index).ORG Then
    Call SendOrganização(i)
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            PlayerMsg Index, "Sua organização acabou de subir de Nível", BrightGreen
        Else
            'plural
            PlayerMsg Index, "Sua Organização acaba de evoluir " & level_count & " Níveis.", BrightGreen
        End If
    End If
  End If
End If
Next
Call SaveOrg(Player(Index).ORG)
End Sub

Function GetONextLevel(ByVal Index As Long, Optional ByVal OrgNum As Byte) As Long
    
    If OrgNum > 0 And OrgNum <= MAX_ORGS Then
        GetONextLevel = (30 / 4) * ((Organization(OrgNum).Level + 15) ^ 4 - (15 * (Organization(OrgNum).Level + 15) ^ 2) + 17 * (Organization(OrgNum).Level + 1) - 12) / 4
        Exit Function
    End If
    
    If Player(Index).ORG > 0 Then
        GetONextLevel = (30 / 4) * ((Organization(Player(Index).ORG).Level + 15) ^ 4 - (15 * (Organization(Player(Index).ORG).Level + 15) ^ 2) + 17 * (Organization(Player(Index).ORG).Level + 1) - 12) / 4
    End If
    
End Function
