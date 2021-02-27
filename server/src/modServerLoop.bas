Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim I As Long, X As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long, tmr10000 As Long, PendNum As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long
    Dim PokeEvolve As Long, n As Long
    Dim tmrSaveLeilao As Long, Desafiante As Long
    
    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr25 Then
            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(I).spellBuffer.Spell > 0 Then
                    If GetPlayerEquipment(I, weapon) > 0 Then
                        If GetTickCount > TempPlayer(I).spellBuffer.Timer + (Spell(Player(I).Spell(TempPlayer(I).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell I, TempPlayer(I).spellBuffer.Spell, TempPlayer(I).spellBuffer.target, TempPlayer(I).spellBuffer.tType
                            TempPlayer(I).spellBuffer.Spell = 0
                            TempPlayer(I).spellBuffer.Timer = 0
                            TempPlayer(I).spellBuffer.target = 0
                            TempPlayer(I).spellBuffer.tType = 0
                        End If
                            Else
                            TempPlayer(I).spellBuffer.Spell = 0
                            TempPlayer(I).spellBuffer.Timer = 0
                            TempPlayer(I).spellBuffer.target = 0
                            TempPlayer(I).spellBuffer.tType = 0
                        End If
                    End If
                    
                    ' check if need to turn off stunned
                    If TempPlayer(I).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(I).StunTimer + (TempPlayer(I).StunDuration * 1000) Then
                            TempPlayer(I).StunDuration = 0
                            TempPlayer(I).StunTimer = 0
                            SendStunned I
                        End If
                    End If
                    
                    ' check regen timer
                    If TempPlayer(I).stopRegen Then
                        If TempPlayer(I).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(I).stopRegen = False
                            TempPlayer(I).stopRegenTimer = 0
                        End If
                    End If
                    
                    ' HoT and DoT logic
                    For X = 1 To MAX_DOTS
                        HandleDoT_Player I, X
                        HandleHoT_Player I, X
                    Next
                End If
                
                ' Check muted timer
                If Player(I).MutedTime > 0 Then
                    If Player(I).MutedTime < GetTickCount Then
                        Player(I).MutedTime = 0
                        PlayerMsg I, "Você pode falar novamente!", BrightGreen
                    End If
                End If
                
                If TempPlayer(I).SwitPoke > 0 Then
                    If TempPlayer(I).SwitPoke < GetTickCount Then
                            Desafiante = TempPlayer(I).Lutando
                            TempPlayer(I).SwitPoke = 0
                            
                            If Desafiante > 0 Then
                                'Avisar o Cotoco que Perdeu e adicionar 1 de Derrota Pela patifaria...
                                PlayerMsg I, "Você não chamou nenhum pokémon em 30 seg e perdeu a batalha.", BrightRed
                                Player(I).Derrotas = Player(I).Derrotas + 1
                                
                                'Recompensa
                                If GetPlayerIP(I) = GetPlayerIP(Desafiante) Then
                                    PlayerMsg Desafiante, "Você venceu a batalha mas não ganhou Vitórias por ter o mesmo IP do Jogador: " & Trim$(GetPlayerName(I)), BrightRed
                                Else
                                    PlayerMsg Desafiante, "Você venceu a batalha!", BrightGreen
                                    Player(Desafiante).Vitorias = Player(Desafiante).Vitorias + 1
                                End If
                                
                                'Limpar dados da BATALHA
                                TempPlayer(I).Lutando = 0
                                TempPlayer(I).LutandoA = 0
                                TempPlayer(I).LutandoT = 0
                                TempPlayer(I).LutQntPoke = 0
                                
                                TempPlayer(Desafiante).Lutando = 0
                                TempPlayer(Desafiante).LutandoA = 0
                                TempPlayer(Desafiante).LutandoT = 0
                                TempPlayer(Desafiante).LutQntPoke = 0
                                
                                'Voltar Jogadores para local Salvo
                                PlayerUnequipItem I, weapon
                                PlayerWarp I, Player(I).MyMap(1), Player(I).MyMap(2), Player(I).MyMap(3)
                                
                                PlayerUnequipItem Desafiante, weapon
                                PlayerWarp Desafiante, Player(Desafiante).MyMap(1), Player(Desafiante).MyMap(2), Player(Desafiante).MyMap(3)
                                
                                'Informar1 Arena VAZIA
                                SendArenaStatus TempPlayer(I).LutandoA, 0
                            Else
                                'Informar Arena Vazia
                                SendArenaStatus TempPlayer(I).LutandoA, 0
                                
                                'Limpar dados da BATALHA
                                TempPlayer(I).Lutando = 0
                                TempPlayer(I).LutandoA = 0
                                TempPlayer(I).LutandoT = 0
                                TempPlayer(I).LutQntPoke = 0
                            End If
                    End If
                End If
                
            'Checar Status Negativo
            If GetPlayerEquipment(I, weapon) > 0 Then
                If GetPlayerEquipmentPokeInfoPokemon(I, weapon) > 0 Then
                    Call CheckNgtStatus(I)
                End If
            End If
                
                'Check Evolution Timer
                If TempPlayer(I).EvolTimer > 0 Then
                If TempPlayer(I).EvolTimer < GetTickCount Then
                If Player(I).EvolPermition = 1 Then
                    PokeEvolve = Pokemon(GetPlayerEquipmentPokeInfoPokemon(I, weapon)).Evolução(1).Pokemon
                    
                    Call SetPlayerEquipmentPokeInfoPokemon(I, PokeEvolve, weapon)
                    Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 1) + 10, weapon, 1) 'Str
                    Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 2) + 10, weapon, 2) 'End
                    Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 3) + 10, weapon, 3) 'Int
                    Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 4) + 10, weapon, 4) 'Agi
                    Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 5) + 2, weapon, 5) 'Res
                    
                If GetPlayerEquipmentShiny(I, weapon) = 0 Then
                    Player(I).Sprite = Pokemon(GetPlayerEquipmentPokeInfoPokemon(I, weapon)).Sprite
                Else
                    Player(I).Sprite = Pokemon(GetPlayerEquipmentPokeInfoPokemon(I, weapon)).Sprite + 1
                End If
                
                'Atualizar Infos
                Call SendPlayerData(I)
                Call SendWornEquipment(I)
                Call SendMapEquipment(I)
                SendAnimation GetPlayerMap(I), 9, GetPlayerX(I), GetPlayerY(I)
                End If
                TempPlayer(I).EvolTimer = 0
                Player(I).EvolTimerStone = 0
                Player(I).EvolPermition = 0
                Player(I).EvolStone = 0
                End If
                End If
                
                If Player(I).EvolTimerStone > 0 Then
                If Player(I).EvolTimerStone < GetTickCount Then
                
                If Player(I).EvolPermition = 1 Then
                PokeEvolve = Player(I).EvolStone
                
                Call SetPlayerEquipmentPokeInfoPokemon(I, PokeEvolve, weapon)
                Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 1) + 10, weapon, 1) 'Str
                Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 2) + 10, weapon, 2) 'End
                Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 3) + 10, weapon, 3) 'Int
                Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 4) + 10, weapon, 4) 'Agi
                Call SetPlayerEquipmentPokeInfoStat(I, GetPlayerEquipmentPokeInfoStat(I, weapon, 5) + 2, weapon, 5) 'Res
                
                If GetPlayerEquipmentShiny(I, weapon) = 0 Then
                    Player(I).Sprite = Pokemon(GetPlayerEquipmentPokeInfoPokemon(I, weapon)).Sprite
                Else
                    Player(I).Sprite = Pokemon(GetPlayerEquipmentPokeInfoPokemon(I, weapon)).Sprite + 1
                End If
                
                'Atualizar Infos
                Call SendPlayerData(I)
                Call SendWornEquipment(I)
                Call SendMapEquipment(I)
                SendAnimation GetPlayerMap(I), 9, GetPlayerX(I), GetPlayerY(I)
                End If
                
                TempPlayer(I).EvolTimer = 0
                Player(I).EvolTimerStone = 0
                Player(I).EvolPermition = 0
                Player(I).EvolStone = 0
                End If
                End If
                
                If Player(I).InFishing > 0 Then
                    If Player(I).InFishing < GetTickCount Then
                        Player(I).InFishing = 0
                        SendInFishing I
                        Call PescarPokemon(I)
                    End If
                End If
                
                If TempPlayer(I).ScanTime > 0 Then
                    If TempPlayer(I).ScanTime < GetTickCount Then
                        SendActionMsg GetPlayerMap(I), "Scaneamento Completo.", White, ACTIONMSG_SCROLL, GetPlayerX(I) * 32, GetPlayerY(I) * 32
                        Player(I).Pokedex(TempPlayer(I).ScanPokemon) = 1 '
                        ChecarTarefasAtuais I, QUEST_TYPE_POKEDEX, 0
                        SendPlayerPokedex I
                        TempPlayer(I).ScanTime = 0
                        TempPlayer(I).ScanPokemon = 0
                        SendInFishing I
                    End If
                End If
                
                If TempPlayer(I).GymTimer > 0 Then
                    If TempPlayer(I).GymTimer < GetTickCount Then
                        TempPlayer(I).GymTimer = 0
                        
                        Select Case TempPlayer(I).InBattleGym
                        Case 1
                        SendContagem I, 0
                        PlayerUnequipItem I, weapon
                        PlayerWarp I, 7, 12, 9
                        SpawnPokeGym 2, 7, 0, 0, 0, DIR_DOWN, False, 0
                        MapNpc(7).Npc(1).InBattle = False
                        PlayerMsg I, "[Brock]: O tempo acabou, infelizmente você perdeu volte quando estiver mais forte e irei batalhar com você novamente.", White
                        TempPlayer(I).InBattleGym = 0
                        TempPlayer(I).GymQntPoke = 0
                        End Select
                    End If
                End If
                
                If TempPlayer(I).GymLeaderPoke(2) > 0 Then
                    If TempPlayer(I).GymLeaderPoke(2) < GetTickCount Then
                    TempPlayer(I).GymLeaderPoke(2) = 0
                    Select Case TempPlayer(I).InBattleGym
                    Case 1 'Brock
                        If TempPlayer(I).GymLeaderPoke(1) = 0 Then
                            IniciarBatalharGym I, 1
                        ElseIf TempPlayer(I).GymLeaderPoke(1) = 1 Then
                            SpawnPokeGym 2, 8, 95, 12, 8, DIR_DOWN, False, 14
                            SendActionMsg GetPlayerMap(I), "Eu escolho você! Vai ONIX!", White, 0, MapNpc(GetPlayerMap(I)).Npc(1).X * 32, MapNpc(GetPlayerMap(I)).Npc(1).Y * 32 - 16
                            SendAnimation GetPlayerMap(I), 7, 12, 8
                        ElseIf TempPlayer(I).GymLeaderPoke(1) = 2 Then
                            PlayerMsg I, "[Aviso]: Você será enviado para o Lobby do Ginásio em 5 Segundos, Não deslogue.", Yellow
                            TempPlayer(I).GymLeaderPoke(1) = 3
                            TempPlayer(I).GymLeaderPoke(2) = 5000 + GetTickCount
                        ElseIf TempPlayer(I).GymLeaderPoke(1) = 3 Then
                            Call PlayerUnequipItem(I, weapon)
                            MapNpc(7).Npc(1).InBattle = False
                            PlayerWarp I, 7, 12, 8
                        End If
                    End Select
                    End If
                End If
            Next
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr25 = GetTickCount + 25
        End If
                    

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For I = 1 To MAX_PLAYERS
                If frmServer.Socket(I).State > sckConnected Then
                    Call CloseSocket(I)
                End If
            Next
            UpdateMapLogic
            tmr500 = GetTickCount + 500
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            tmr1000 = GetTickCount + 1000
        End If
        
        If tmr10000 < GetTickCount Then
        For I = 1 To MAX_LEILAO
            
        If Leilao(I).Vendedor <> vbNullString Then
            Leilao(I).Tempo = Leilao(I).Tempo - 1
            
            If Leilao(I).Tempo = 0 Then
            
                If IsPlaying(FindPlayer(Leilao(I).Vendedor)) = True Then
                    GiveInvItem FindPlayer(Leilao(I).Vendedor), Leilao(I).ItemNum, 1, True, _
                    Leilao(I).Poke.Pokemon, Leilao(I).Poke.Pokeball, _
                    Leilao(I).Poke.Level, Leilao(I).Poke.EXP, _
                    Leilao(I).Poke.Vital(1), Leilao(I).Poke.Vital(2), _
                    Leilao(I).Poke.MaxVital(1), Leilao(I).Poke.MaxVital(2), _
                    Leilao(I).Poke.Stat(1), Leilao(I).Poke.Stat(4), _
                    Leilao(I).Poke.Stat(2), Leilao(I).Poke.Stat(3), _
                    Leilao(I).Poke.Stat(5), Leilao(I).Poke.Spells(1), _
                    Leilao(I).Poke.Spells(2), Leilao(I).Poke.Spells(3), _
                    Leilao(I).Poke.Spells(4)
                    PlayerMsg FindPlayer(Leilao(I).Vendedor), "Leilão falhou, tempo expirado.", BrightRed
                Else
                    PendNum = FindPend
                    
                    Pendencia(PendNum).Vendedor = Leilao(I).Vendedor
                    Pendencia(PendNum).ItemNum = Leilao(I).ItemNum
                    Pendencia(PendNum).Tipo = 0
                    Pendencia(PendNum).Price = 1
                    
                    '#Pokemon#
                    Pendencia(PendNum).Poke.Pokemon = Leilao(I).Poke.Pokemon
                    Pendencia(PendNum).Poke.Pokeball = Leilao(I).Poke.Pokeball
                    Pendencia(PendNum).Poke.Level = Leilao(I).Poke.Level
                    Pendencia(PendNum).Poke.EXP = Leilao(I).Poke.EXP
                    Pendencia(PendNum).Poke.Felicidade = Leilao(I).Poke.Felicidade
                    Pendencia(PendNum).Poke.Sexo = Leilao(I).Poke.Sexo
                    Pendencia(PendNum).Poke.Shiny = Leilao(I).Poke.Shiny
                    
                    For X = 1 To Vitals.Vital_Count - 1
                        Pendencia(PendNum).Poke.Vital(X) = Leilao(I).Poke.Vital(X)
                        Pendencia(PendNum).Poke.MaxVital(X) = Leilao(I).Poke.MaxVital(X)
                    Next
                    
                    For X = 1 To Stats.Stat_Count - 1
                        Pendencia(PendNum).Poke.Stat(X) = Leilao(I).Poke.Stat(X)
                    Next
                    
                    For X = 1 To MAX_POKE_SPELL
                        Pendencia(PendNum).Poke.Spells(X) = Leilao(I).Poke.Spells(X)
                    Next
                    
                    SavePendencia PendNum
                End If
            
                LimparLeilaoSlot I '< sub que eu fiz pra limpar o leilão ._.
                ArrumaLeilao
                SendAttLeilao
            End If
        End If
        Next
        
            tmr10000 = GetTickCount + 1000
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 300000
        End If

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub UpdateMapSpawnItems()
    Dim X As Long
    Dim Y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For Y = 1 To MAX_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(Y) Then

            ' Clear out unnecessary junk
            For X = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(X, Y)
            Next

            ' Spawn the items
            Call SpawnMapItems(Y)
            Call SendMapItemsToAll(Y)
        End If

        DoEvents
    Next

End Sub

Private Sub UpdateMapLogic()
    Dim I As Long, X As Long, MapNum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long
    Dim target As Long, targetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean

    For MapNum = 1 To MAX_MAPS
        ' items appearing to everyone
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(MapNum, I).Num > 0 Then
                ' despawn item?
                If MapItem(MapNum, I).canDespawn Then
                    If MapItem(MapNum, I).despawnTimer < GetTickCount Then
                        ' despawn it
                        ClearMapItem I, MapNum
                        ' send updates to everyone
                        SendMapItemsToAll MapNum
                    End If
                End If
            End If
        Next

        
        '  Close the doors
        If TickCount > TempTile(MapNum).DoorTimer + 5000 Then
            For x1 = 0 To Map(MapNum).MaxX
                For y1 = 0 To Map(MapNum).MaxY
                    If Map(MapNum).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(MapNum).DoorOpen(x1, y1) = YES Then
                        TempTile(MapNum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap MapNum, x1, y1, 0
                    End If
                Next
            Next
        End If
        
        ' check for DoTs + hots
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(MapNum).Npc(I).Num > 0 Then
                For X = 1 To MAX_DOTS
                    HandleDoT_Npc MapNum, I, X
                    HandleHoT_Npc MapNum, I, X
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(MapNum).Resource_Count > 0 Then
            For I = 0 To ResourceCache(MapNum).Resource_Count
                Resource_index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(I).X, ResourceCache(MapNum).ResourceData(I).Y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(MapNum).ResourceData(I).ResourceState = 1 Or ResourceCache(MapNum).ResourceData(I).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(MapNum).ResourceData(I).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(MapNum).ResourceData(I).ResourceTimer = GetTickCount
                            ResourceCache(MapNum).ResourceData(I).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(MapNum).ResourceData(I).cur_health = Resource(Resource_index).health
                            SendResourceCacheToMap MapNum, I
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(MapNum) = YES Then
            TickCount = GetTickCount
            
            For X = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(MapNum).Npc(X).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(MapNum).Npc(X).StunDuration > 0 Then
    
                            For I = 1 To Player_HighIndex
                                If IsPlaying(I) Then
                                    If GetPlayerMap(I) = MapNum And MapNpc(MapNum).Npc(X).target = 0 And GetPlayerAccess(I) <= ADMIN_MONITOR Then
                                        n = Npc(NpcNum).Range
                                        DistanceX = MapNpc(MapNum).Npc(X).X - GetPlayerX(I)
                                        DistanceY = MapNpc(MapNum).Npc(X).Y - GetPlayerY(I)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(I) = YES Then
                                                If Len(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                                                    Call PlayerMsg(I, Trim$(Npc(NpcNum).Name) & " says: " & Trim$(Npc(NpcNum).AttackSay), SayColor)
                                                End If
                                                MapNpc(MapNum).Npc(X).targetType = 1 ' player
                                                MapNpc(MapNum).Npc(X).target = I
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then
                    If MapNpc(MapNum).Npc(X).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(MapNum).Npc(X).StunTimer + (MapNpc(MapNum).Npc(X).StunDuration * 1000) Then
                            MapNpc(MapNum).Npc(X).StunDuration = 0
                            MapNpc(MapNum).Npc(X).StunTimer = 0
                        End If
                    Else
                            
                        target = MapNpc(MapNum).Npc(X).target
                        targetType = MapNpc(MapNum).Npc(X).targetType
    
                        ' Check to see if its time for the npc to walk
                        If Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        
                            If targetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = GetPlayerY(target)
                                        TargetX = GetPlayerX(target)
                                    Else
                                        MapNpc(MapNum).Npc(X).targetType = 0 ' clear
                                        MapNpc(MapNum).Npc(X).target = 0
                                    End If
                                End If
                            
                            ElseIf targetType = 2 Then 'npc
                                
                                If target > 0 Then
                                    
                                    If MapNpc(MapNum).Npc(target).Num > 0 Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = MapNpc(MapNum).Npc(target).Y
                                        TargetX = MapNpc(MapNum).Npc(target).X
                                    Else
                                        MapNpc(MapNum).Npc(X).targetType = 0 ' clear
                                        MapNpc(MapNum).Npc(X).target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                
                                I = Int(Rnd * 5)
    
                                ' Lets move the npc
                                Select Case I
                                    Case 0
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_UP) Then
                                                Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 1
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_UP) Then
                                                Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 2
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_UP) Then
                                                Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 3
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_UP) Then
                                                Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                End Select
    
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(MapNum).Npc(X).X - 1 = TargetX And MapNpc(MapNum).Npc(X).Y = TargetY Then
                                        If MapNpc(MapNum).Npc(X).Dir <> DIR_LEFT Then
                                            Call NpcDir(MapNum, X, DIR_LEFT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(MapNum).Npc(X).X + 1 = TargetX And MapNpc(MapNum).Npc(X).Y = TargetY Then
                                        If MapNpc(MapNum).Npc(X).Dir <> DIR_RIGHT Then
                                            Call NpcDir(MapNum, X, DIR_RIGHT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(MapNum).Npc(X).X = TargetX And MapNpc(MapNum).Npc(X).Y - 1 = TargetY Then
                                        If MapNpc(MapNum).Npc(X).Dir <> DIR_UP Then
                                            Call NpcDir(MapNum, X, DIR_UP)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(MapNum).Npc(X).X = TargetX And MapNpc(MapNum).Npc(X).Y + 1 = TargetY Then
                                        If MapNpc(MapNum).Npc(X).Dir <> DIR_DOWN Then
                                            Call NpcDir(MapNum, X, DIR_DOWN)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    ' We could not move so Target must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        I = Int(Rnd * 2)
    
                                        If I = 1 Then
                                            I = Int(Rnd * 4)
    
                                            If CanNpcMove(MapNum, X, I) Then
                                                Call NpcMove(MapNum, X, I, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
    
                            Else
                                I = Int(Rnd * 4)
    
                                If I = 1 Then
                                    I = Int(Rnd * 4)
    
                                    If CanNpcMove(MapNum, X, I) Then
                                        Call NpcMove(MapNum, X, I, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then
                    target = MapNpc(MapNum).Npc(X).target
                    targetType = MapNpc(MapNum).Npc(X).targetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                    
                        If targetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                            
                            If MapNpc(MapNum).Npc(X).Desmaiado = False Then
                                TryNpcAttackPlayer X, target
                            End If
                            
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(MapNum).Npc(X).target = 0
                                MapNpc(MapNum).Npc(X).targetType = 0 ' clear
                            End If
                        Else
                            ' lol no npc combat :(
                        End If
                    End If
                End If
                
                'Correção NPC HP REGEN
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                Dim PI As Long
                
                    ' check regen timer
                    If MapNpc(MapNum).Npc(X).stopRegen Then
                        If MapNpc(MapNum).Npc(X).stopRegenTimer + 5000 < GetTickCount Then
                            MapNpc(MapNum).Npc(X).stopRegen = False
                            MapNpc(MapNum).Npc(X).stopRegenTimer = 0
                        End If
                    End If
                
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(MapNum).Npc(X).stopRegen Then
                    If MapNpc(MapNum).Npc(X).Desmaiado = False Then
                        If MapNpc(MapNum).Npc(X).Num > 0 And TickCount > GiveNPCHPTimer + 1000 Then
                            
                            If Npc(MapNpc(MapNum).Npc(X).Num).Pokemon = 0 Then
                                If MapNpc(MapNum).Npc(X).Vital(Vitals.HP) > 0 And MapNpc(MapNum).Npc(X).Vital(Vitals.HP) < GetNpcMaxVital(NpcNum, Vitals.HP) Then
                                    MapNpc(MapNum).Npc(X).Vital(Vitals.HP) = MapNpc(MapNum).Npc(X).Vital(Vitals.HP) + GetNpcVitalRegen(NpcNum, Vitals.HP)
                                    SendActionMsg MapNum, "+" & GetNpcVitalRegen(NpcNum, Vitals.HP), Green, 1, MapNpc(MapNum).Npc(X).X * 32, MapNpc(MapNum).Npc(X).Y * 32
                                      
                                    MapNpc(MapNum).Npc(X).stopRegen = True
                                    MapNpc(MapNum).Npc(X).stopRegenTimer = GetTickCount
            
                                    ' Check if they have more then they should and if so just set it to max
                                    If MapNpc(MapNum).Npc(X).Vital(Vitals.HP) > GetNpcMaxVital(NpcNum, Vitals.HP) Then
                                        MapNpc(MapNum).Npc(X).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
                                        SendMapNpcVitals MapNum, X
                                    End If
                                End If
                            Else
                                If MapNpc(MapNum).Npc(X).Vital(Vitals.HP) > 0 And MapNpc(MapNum).Npc(X).Vital(Vitals.HP) < GetPokemonMaxVital(NpcNum, Vitals.HP, MapNpc(MapNum).Npc(X).Level) Then
                                    MapNpc(MapNum).Npc(X).Vital(Vitals.HP) = MapNpc(MapNum).Npc(X).Vital(Vitals.HP) + GetPokemonVitalRegen(NpcNum, Vitals.HP)
                                    If GetPokemonVitalRegen(NpcNum, Vitals.HP) > 0 Then
                                        SendActionMsg MapNum, "+" & GetPokemonVitalRegen(NpcNum, Vitals.HP), Green, 1, MapNpc(MapNum).Npc(X).X * 32, MapNpc(MapNum).Npc(X).Y * 32
                                    End If
                                    
                                    MapNpc(MapNum).Npc(X).stopRegen = True
                                    MapNpc(MapNum).Npc(X).stopRegenTimer = GetTickCount
            
                                    ' Check if they have more then they should and if so just set it to max
                                    If MapNpc(MapNum).Npc(X).Vital(Vitals.HP) > GetPokemonMaxVital(NpcNum, Vitals.HP, MapNpc(MapNum).Npc(X).Level) Then
                                        MapNpc(MapNum).Npc(X).Vital(Vitals.HP) = GetPokemonMaxVital(NpcNum, Vitals.HP, MapNpc(MapNum).Npc(X).Level)
                                    End If
                                    
                                    'Enviar Hp do pokémon
                                    SendMapNpcVitals MapNum, X
                                    
                                    'Verificar se não tem jogadores com target nessa peste de BIXO!
                                    For PI = 1 To Player_HighIndex
                                        If IsPlaying(PI) = True Then
                                            If GetPlayerMap(PI) = MapNum Then
                                                If TempPlayer(PI).target = X Then
                                                    SendTarget PI, X
                                                End If
                                            End If
                                        End If
                                    Next
                                    
                                End If
                            End If

                        End If
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(MapNum).Npc(X).Desmaiado = True And Map(MapNum).Npc(X) > 0 Then
                    'If TickCount > MapNpc(MapNum).Npc(X).SpawnWait + (Npc(Map(MapNum).Npc(X)).SpawnSecs * 1000) Then
                    If TickCount > MapNpc(MapNum).Npc(X).SpawnWait + (30 * 1000) Then
                        If MapNpc(MapNum).Npc(X).Pescado = True Then
                            Call RemoverNpcPescado(X, MapNum)
                        Else
                            Call SpawnNpc(X, MapNum)
                        End If
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

End Sub

Private Sub UpdatePlayerVitals()
Dim I As Long
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If Not TempPlayer(I).stopRegen Then
                If GetPlayerVital(I, Vitals.HP) <> GetPlayerMaxVital(I, Vitals.HP) Then
                    Call SetPlayerVital(I, Vitals.HP, GetPlayerVital(I, Vitals.HP) + GetPlayerVitalRegen(I, Vitals.HP))
                    Call SendVital(I, Vitals.HP)
                    ' send vitals to party if in one
                    
                    If TempPlayer(I).inParty > 0 Then SendPartyVitals TempPlayer(I).inParty, I
                End If
    
                If GetPlayerVital(I, Vitals.MP) <> GetPlayerMaxVital(I, Vitals.MP) Then
                    Call SetPlayerVital(I, Vitals.MP, GetPlayerVital(I, Vitals.MP) + GetPlayerVitalRegen(I, Vitals.MP))
                    Call SendVital(I, Vitals.MP)
                    ' send vitals to party if in one
                    
                    If TempPlayer(I).inParty > 0 Then SendPartyVitals TempPlayer(I).inParty, I
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim I As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")

        For I = 1 To Player_HighIndex

            If IsPlaying(I) Then
                Call SavePlayer(I)
                Call SaveBank(I)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub

Public Sub CheckNgtStatus(ByVal I As Long)
Dim X As Long, Y As Long, MapNum As Long

X = GetPlayerX(I)
Y = GetPlayerY(I)
MapNum = GetPlayerMap(I)

    'Burn
        If TempPlayer(I).NgtTick(1) > 0 Then
            If TempPlayer(I).NgtTick(1) < GetTickCount Then
            
            If GetPlayerEquipmentNgt(I, weapon, 1) > 0 Then
                If GetPlayerVital(I, HP) - Player(I).NgtDamage(1) <= 0 Then
                    SetPlayerEquipmentNgt I, 1, weapon, 0
                    SetPlayerVital I, HP, 0
                    PlayerMsg I, "Seu Pokémon Morreu Por chamas!", White
                    PlayerUnequipItem I, weapon
                    TempPlayer(I).NgtTick(1) = 0
                Else
                    SetPlayerEquipmentNgt I, 1, weapon, GetPlayerEquipmentNgt(I, weapon, 1) - 1
                    SetPlayerVital I, HP, GetPlayerVital(I, HP) - Player(I).NgtDamage(1)
                    Call SendActionMsg(MapNum, "-" & Player(I).NgtDamage(1), BrightRed, 1, X * 32, Y * 32)
                    SendVital I, HP
                    SendMapEquipment I
                    TempPlayer(I).NgtTick(1) = 1000 + GetTickCount
                End If
            End If
            
        End If
    End If
    
    'Frozen
    If TempPlayer(I).NgtTick(2) > 0 Then
        If TempPlayer(I).NgtTick(2) < GetTickCount Then
            If GetPlayerEquipmentNgt(I, weapon, 2) > 0 Then
                    SetPlayerEquipmentNgt I, 2, weapon, GetPlayerEquipmentNgt(I, weapon, 2) - 1
                    SendVital I, HP
                    SendMapEquipment I
                    TempPlayer(I).NgtTick(2) = 1000 + GetTickCount
            End If
        End If
    End If

    'Poison
    If TempPlayer(I).NgtTick(3) > 0 Then
            If TempPlayer(I).NgtTick(3) < GetTickCount Then
            
            If GetPlayerEquipmentNgt(I, weapon, 3) > 0 Then
                If GetPlayerVital(I, HP) - Player(I).NgtDamage(3) <= 0 Then
                    SetPlayerEquipmentNgt I, 3, weapon, 0
                    SetPlayerVital I, HP, 0
                    PlayerMsg I, "Seu Pokémon Morreu Envenenado!", White
                    PlayerUnequipItem I, weapon
                    TempPlayer(I).NgtTick(3) = 0
                Else
                    SetPlayerEquipmentNgt I, 3, weapon, GetPlayerEquipmentNgt(I, weapon, 3) - 1
                    SetPlayerVital I, HP, GetPlayerVital(I, HP) - Player(I).NgtDamage(3)
                    Call SendActionMsg(MapNum, "-" & Player(I).NgtDamage(3), BrightRed, 1, X * 32, Y * 32)
                    SendVital I, HP
                    SendMapEquipment I
                    TempPlayer(I).NgtTick(3) = 1000 + GetTickCount
                End If
            End If
            
        End If
    End If
    
    'Atração
    If TempPlayer(I).NgtTick(4) > 0 Then
        If TempPlayer(I).NgtTick(4) < GetTickCount Then
            If GetPlayerEquipmentNgt(I, weapon, 4) > 0 Then
                    SetPlayerEquipmentNgt I, 4, weapon, GetPlayerEquipmentNgt(I, weapon, 4) - 1
                    SendVital I, HP
                    SendMapEquipment I
                    TempPlayer(I).NgtTick(4) = 1000 + GetTickCount
            End If
        End If
    End If
    
    'Confuso
    If TempPlayer(I).NgtTick(5) > 0 Then
        If TempPlayer(I).NgtTick(5) < GetTickCount Then
            If GetPlayerEquipmentNgt(I, weapon, 5) > 0 Then
                    SetPlayerEquipmentNgt I, 5, weapon, GetPlayerEquipmentNgt(I, weapon, 5) - 1
                    SendVital I, HP
                    SendMapEquipment I
                    TempPlayer(I).NgtTick(5) = 1000 + GetTickCount
            End If
        End If
    End If

    'Paralisia
    If TempPlayer(I).NgtTick(6) > 0 Then
        If TempPlayer(I).NgtTick(6) < GetTickCount Then
            If GetPlayerEquipmentNgt(I, weapon, 6) > 0 Then
                    SetPlayerEquipmentNgt I, 6, weapon, GetPlayerEquipmentNgt(I, weapon, 6) - 1
                    SendVital I, HP
                    SendMapEquipment I
                    TempPlayer(I).NgtTick(6) = 1000 + GetTickCount
            End If
        End If
    End If
    
End Sub
