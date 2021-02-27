Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(Index)
                Case Else ' Treinador
                
                If GetPlayerEquipment(Index, weapon) = 0 Then
                    GetPlayerMaxVital = 100
                Else
                    If GetPlayerEquipmentPokeInfoMaxVital(Index, weapon, 1) > 0 Then
                        GetPlayerMaxVital = GetPlayerEquipmentPokeInfoMaxVital(Index, weapon, 1)
                    End If
                End If
                
                End Select
        Case MP
            Select Case GetPlayerClass(Index)
                Case Else ' Treinador
                
                If GetPlayerEquipment(Index, weapon) = 0 Then
                    GetPlayerMaxVital = 1
                Else
                    If GetPlayerEquipmentPokeInfoMaxVital(Index, weapon, 2) > 0 Then
                    GetPlayerMaxVital = GetPlayerEquipmentPokeInfoMaxVital(Index, weapon, 2)
                    End If
                End If
                
                End Select
    End Select
End Function

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim I As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            I = (GetPlayerStat(Index, Stats.Willpower) * 0.8) + 6
            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                I = (GetPlayerEquipmentPokeInfoStat(Index, weapon, 5) * 0.5) + (Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Vital(1) * 25 / 100)
            End If
        Case MP
            I = (GetPlayerStat(Index, Stats.Willpower) / 4) + 12.5
    End Select

    If I < 2 Then I = 2
    GetPlayerVitalRegen = I
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(Index, weapon) > 0 Then
        weaponNum = GetPlayerEquipment(Index, weapon)
        GetPlayerDamage = GetPlayerEquipmentPokeInfoStat(Index, weapon, 1)
        
        'Felicidade
        Select Case GetPlayerEquipmentFelicidade(Index, weapon)
        Case 0 To 50 '-3% de Atk
            GetPlayerDamage = GetPlayerDamage - ((GetPlayerDamage * 3) / 100)
        Case 51 To 100 '0% de Atk
            GetPlayerDamage = GetPlayerDamage
        Case 101 To 200 '3% de +Atk
            GetPlayerDamage = GetPlayerDamage + ((GetPlayerDamage * 3) / 100)
        Case 201 To 300 '5% de +Atk
            GetPlayerDamage = GetPlayerDamage + ((GetPlayerDamage * 5) / 100)
        Case 301 To 400 '8% de +Atk
            GetPlayerDamage = GetPlayerDamage + ((GetPlayerDamage * 8) / 100)
        Case Else '10% de +Atk
            GetPlayerDamage = GetPlayerDamage + ((GetPlayerDamage * 10) / 100)
        End Select
        
        'Berry - Dano está fora da Porcentagem da Felicidade :)
        GetPlayerDamage = GetPlayerDamage + GetPlayerEquipmentBerry(Index, weapon, 1)
        
    Else
        GetPlayerDamage = 0
    End If

End Function

Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = Npc(NpcNum).HP
        Case MP
            GetNpcMaxVital = 30 + (Npc(NpcNum).Stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetPokemonMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals, ByVal Level As Byte) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetPokemonMaxVital = 0
        Exit Function
    End If
    
    GetPokemonMaxVital = Pokemon(Npc(NpcNum).Pokemon).Vital(Vitals.HP) + (Level * 5)

End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim I As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            I = (Npc(NpcNum).Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            I = (Npc(NpcNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = I

End Function

Function GetPokemonVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim I As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetPokemonVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            I = (Pokemon(Npc(NpcNum).Pokemon).Vital(1) * 10 / 100) + (Pokemon(Npc(NpcNum).Pokemon).Add_Stat(3) * 0.5)
        Case MP
            I = (Npc(NpcNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetPokemonVitalRegen = I

End Function

Function GetNpcDamage(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Long

    GetNpcDamage = MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Strength)
    
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    If GetPlayerEquipment(Index, weapon) = 0 Then Exit Function
    
    rate = GetPlayerEquipmentPokeInfoStat(Index, weapon, Stats.Agility) / 40
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False

    If GetPlayerEquipment(Index, weapon) = 0 Then Exit Function

    rate = GetPlayerEquipmentPokeInfoStat(Index, weapon, Stats.Agility) / 90
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False

    If GetPlayerEquipment(Index, weapon) = 0 Then Exit Function

    rate = GetPlayerEquipmentPokeInfoStat(Index, weapon, Stats.Strength) * 0.1
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcCrit(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Agility) / 40
    rndNum = RAND(1, 100)
    
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
    
End Function

Public Function CanNpcDodge(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Agility) / 90
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Strength) * 0.1
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal MapNpcNum As Long)
Dim NpcNum As Long, MapNum As Long, Damage As Long, Random50 As Long

    Damage = 0
    
    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, MapNpcNum) Then
    
        MapNum = GetPlayerMap(Index)
        NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num

        ' Esquiva Pokémon
        If CanNpcDodge(MapNum, MapNpcNum) Then
            SendActionMsg MapNum, "Esquivou!", Pink, 1, (MapNpc(MapNum).Npc(MapNpcNum).X * 32), (MapNpc(MapNum).Npc(MapNpcNum).Y * 32)
            Exit Sub
        End If
        
        ' Defesa Pokémon
        If CanNpcParry(MapNum, MapNpcNum) Then
            SendActionMsg MapNum, "Defendeu!", Pink, 1, (MapNpc(MapNum).Npc(MapNpcNum).X * 32), (MapNpc(MapNum).Npc(MapNpcNum).Y * 32)
            Exit Sub
        End If

        ' Dano do Pokémon do Jogador
        Damage = GetPlayerDamage(Index)
        
        ' Defesa do Pokémon
        Damage = Damage - RAND(1, (MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Endurance) * 1.5))
        
        ' 75% de Atk
        Random50 = Porcento(Damage, 75)
        
        ' Random 75%~100% Atk
        Damage = RAND(Random50, Damage)
        
        ' Dano * 2 Critico
        If CanPlayerCriticalHit(Index) Then
            Damage = Damage * 2
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
        
        'Confusão 35% - (Atq * 0.01)% = Dano / 3
        If GetPlayerEquipmentNgt(Index, weapon, 5) > 0 Then
            If CanPlayerConfusionHit(Index) = True Then
                Damage = Int(Damage / 3)
                SendActionMsg MapNum, "Confuso!", Yellow, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                Call NpcAttackPlayer(MapNpcNum, Index, Damage)
                Exit Sub
            End If
        End If
        
        ' Atração 45% de Chance de atacar com 10% de atk
        If GetPlayerEquipmentNgt(Index, weapon, 4) > 0 Then
            If Not MapNpc(MapNum).Npc(MapNpcNum).Sexo = GetPlayerEquipmentSexo(Index, weapon) Then
                If CanPlayerAtractHit(Index) = True Then
                    Damage = (Damage * 10) / 100
                End If
            End If
        End If
            
        ' Conclusão
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, MapNpcNum, Damage)
        Else
            SendActionMsg MapNum, "Fraco!", Black, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
   ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Or MapNpc(MapNum).Npc(MapNpcNum).Desmaiado Then
        If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
        ' exit out early
        If IsSpell Then
             If NpcNum > 0 Then
                If Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER And Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_QUEST Then
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(attacker, weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(attacker, weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If NpcNum > 0 And GetTickCount > TempPlayer(attacker).AttackTimer + attackspeed Then
            
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).X
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).Y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).X
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).Y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).X + 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).Y
                Case DIR_RIGHT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).X - 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).Y
            End Select

            If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Or Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_QUEST Then
                If GetPlayerEquipmentPokeInfoPokemon(attacker, weapon) > 0 Then
                    CanPlayerAttackNpc = False
                    Exit Function
                End If
            End If

            If NpcX = GetPlayerX(attacker) Then
                If NpcY = GetPlayerY(attacker) Then
                    If Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER And Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_QUEST Then
                        
                        If GetPlayerEquipment(attacker, weapon) < 1 Then
                            PlayerMsg attacker, "Só é possivel atacar usando um Pokémon!", White
                            Exit Function
                        Else
                        If GetPlayerEquipmentPokeInfoPokemon(attacker, weapon) = 0 Then
                                PlayerMsg attacker, "Só é possivel atacar usando um Pokémon!", White
                                Exit Function
                            End If
                        End If
                        
                        CanPlayerAttackNpc = True
                    Else
                        
                        If NpcNum = 1 Then 'Profº Oak
                            If Player(attacker).PokeInicial = 1 Then
                                SendEscolherPokeInicial attacker
                                Exit Function
                            End If
                        ElseIf NpcNum = 999 Then
                            If Player(attacker).Insignia(1) = 0 Then
                                SendComandoGym attacker, 1
                                Exit Function
                            End If
                        End If
                        
                        ' Check if the player completed a quest
                        ChecarTarefasAtuais attacker, QUEST_TYPE_TALKNPC, NpcNum
                            
                        ' Open the selector of quest
                        If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_QUEST Then
                            SendQuestCommand attacker, 1, NpcNum
                            TempPlayer(attacker).QuestSelect = NpcNum
                            Exit Function
                        End If
                        
                        If Len(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                            PlayerMsg attacker, Trim$(Npc(NpcNum).Name) & ": " & Trim$(Npc(NpcNum).AttackSay), White
                        End If
                        
                        ' Reset attack timer
                         TempPlayer(attacker).AttackTimer = GetTickCount
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim EXP As Long
    Dim n As Long
    Dim I As Long
    Dim STR As Long
    Dim DEF As Long
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    Name = Trim$(Npc(NpcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, weapon) > 0 Then
        n = GetPlayerEquipment(attacker, weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If Damage >= MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) Then
    
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).Npc(MapNpcNum).X * 32), (MapNpc(MapNum).Npc(MapNpcNum).Y * 32)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(MapNpcNum).X, MapNpc(MapNum).Npc(MapNpcNum).Y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, Pokemon(GetPlayerEquipmentPokeInfoPokemon(attacker, weapon)).AnimAttack, MapNpc(MapNum).Npc(MapNpcNum).X, MapNpc(MapNum).Npc(MapNpcNum).Y)
                If SpellNum > 0 Then Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, MapNpc(MapNum).Npc(MapNpcNum).X, MapNpc(MapNum).Npc(MapNpcNum).Y)
            End If
        End If

        ' Calculate exp to give attacker
        If Npc(NpcNum).Pokemon > 0 Then
            EXP = Pokemon(Npc(NpcNum).Pokemon).ExpBase * MapNpc(MapNum).Npc(MapNpcNum).Level
        Else
            EXP = Npc(NpcNum).EXP
        End If
        
        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 1
        End If

        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, EXP, attacker
        Else
            For I = 1 To Player_HighIndex
                If IsPlaying(I) = True Then
                    If Player(I).ORG > 0 Then
                        If GetPlayerMap(I) = GetPlayerMap(attacker) Then
                            GivePlayerEXP I, (EXP * 5) / 100
                        End If
                    End If
                End If
            Next
        End If
        
        If TempPlayer(attacker).inParty = 0 And Player(attacker).ORG = 0 Then
            GivePlayerEXP attacker, EXP
        End If
        
        If Player(attacker).ORG > 0 Then
            If Organization(Player(attacker).ORG).Level <= 9 Then
                Organization(Player(attacker).ORG).EXP = Organization(Player(attacker).ORG).EXP + (EXP * 2) / 100
                Call CheckAORGlevelUP(Player(attacker).ORG)
                Call SaveOrgExp(Player(attacker).ORG)
                        
                For I = 1 To Player_HighIndex
                    If Player(I).ORG = Player(attacker).ORG Then
                        Call SendOrganização(I)
                    End If
                Next
            Else
                If Organization(Player(attacker).ORG).EXP < GetONextLevel(attacker) Then
                    Organization(Player(attacker).ORG).EXP = GetONextLevel(attacker)
                End If
            End If
        End If
        
        'Drop the goods if they get it
        n = Int(Rnd * Npc(NpcNum).DropChance) + 1

        If n = 1 Then
            Call SpawnItem(Npc(NpcNum).DropItem, Npc(NpcNum).DropItemValue, MapNum, MapNpc(MapNum).Npc(MapNpcNum).X, MapNpc(MapNum).Npc(MapNpcNum).Y)
        End If

        'Setar Felicidade
        If GetPlayerEquipment(attacker, weapon) > 0 Then
            Call SetPlayerEquipmentFelicidade(attacker, weapon, GetPlayerEquipmentFelicidade(attacker, weapon) + 1)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = 0
        MapNpc(MapNum).Npc(MapNpcNum).Desmaiado = True
        MapNpc(MapNum).Npc(MapNpcNum).targetType = 0
        MapNpc(MapNum).Npc(MapNpcNum).target = 0
        
        'Checar Quest
        ChecarTarefasAtuais attacker, 1, MapNpc(MapNum).Npc(MapNpcNum).Num
        
        ' clear DoTs and HoTs
        For I = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(MapNpcNum).DoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).Npc(MapNpcNum).HoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Call SendNpcDesmaiado(MapNum, MapNpcNum, False)
        Call SendMapNpcVitals(MapNum, MapNpcNum)
        
        'Loop through entire map and purge NPC from targets
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And IsConnected(I) Then
                If Player(I).Map = MapNum Then
                    If TempPlayer(I).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(I).target = MapNpcNum Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendTarget I
                        End If
                    End If
                End If
            End If
        Next
        
        'Pokémons de Ginasios
        If GetPlayerMap(attacker) = 8 Then
            If MapNpcNum = 2 Then
                If MapNpc(MapNum).Npc(MapNpcNum).Num = 74 Then
                    SendActionMsg MapNum, "Volte Geodude! Você lutou Bem!", White, 0, MapNpc(MapNum).Npc(1).X * 32, MapNpc(MapNum).Npc(1).Y * 32 - 16
                    SendAnimation MapNum, 8, MapNpc(MapNum).Npc(2).X, MapNpc(MapNum).Npc(2).Y
                    SpawnPokeGym 2, 8, 0, 0, 0, DIR_DOWN, False, 0
                    TempPlayer(attacker).GymLeaderPoke(1) = 1
                    TempPlayer(attacker).GymLeaderPoke(2) = 3000 + GetTickCount
                    
                ElseIf MapNpc(MapNum).Npc(MapNpcNum).Num = 95 Then
                    SendActionMsg MapNum, "Volte Onix! Você lutou Bem!", White, 0, MapNpc(MapNum).Npc(1).X * 32, MapNpc(MapNum).Npc(1).Y * 32 - 16
                    SendAnimation MapNum, 8, MapNpc(MapNum).Npc(2).X, MapNpc(MapNum).Npc(2).Y
                    SpawnPokeGym 2, 8, 0, 0, 0, DIR_DOWN, False, 0
                    TempPlayer(attacker).GymTimer = 0
                    SendContagem attacker, 5
                    TempPlayer(attacker).GymLeaderPoke(1) = 2
                    TempPlayer(attacker).GymLeaderPoke(2) = 3000 + GetTickCount
                    Player(attacker).Insignia(1) = 1
                    Player(attacker).Insignia(2) = 1
                    Player(attacker).Insignia(3) = 1
                    Player(attacker).Insignia(4) = 1
                    Player(attacker).Insignia(5) = 1
                    Player(attacker).Insignia(6) = 1
                    Player(attacker).Insignia(7) = 1
                    Player(attacker).Insignia(8) = 1
                    PlayerMsg attacker, "[Brock]: Parábens! Eu Substmei você, aqui está a Insignia de Pedra e um presente para futuras batalhas a TM Rock Tomb.", BrightGreen
                End If
            End If
        End If
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).Npc(MapNpcNum).X * 32), (MapNpc(MapNum).Npc(MapNpcNum).Y * 32)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(MapNpcNum).X, MapNpc(MapNum).Npc(MapNpcNum).Y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, Pokemon(GetPlayerEquipmentPokeInfoPokemon(attacker, weapon)).AnimAttack, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(MapNpcNum).targetType = 1 ' player
        MapNpc(MapNum).Npc(MapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(I).Num = MapNpc(MapNum).Npc(MapNpcNum).Num Then
                    MapNpc(MapNum).Npc(I).target = attacker
                    MapNpc(MapNum).Npc(I).targetType = 1 ' player
                End If
            Next
        End If
        
        'Enviar Hp do Pokémon Alvo
        If TempPlayer(attacker).target = MapNpcNum Then
            SendTarget attacker, MapNpcNum
        End If
        
        ' set the regen timer
        MapNpc(MapNum).Npc(MapNpcNum).stopRegen = True
        MapNpc(MapNum).Npc(MapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC MapNpcNum, MapNum, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Npc MapNum, MapNpcNum, SpellNum, attacker
            End If
        End If
        
        SendMapNpcVitals MapNum, MapNpcNum
    End If

    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = GetTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long)
Dim MapNum As Long, NpcNum As Long, blockAmount As Long, Damage As Long
Dim Dano75 As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNpcNum, Index) Then
        MapNum = GetPlayerMap(Index)
        NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
        ' Esquiva
        If CanPlayerDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            Exit Sub
        End If
        
        ' Defesa
        If CanPlayerParry(Index) Then
            SendActionMsg MapNum, "Defendeu!", Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            Exit Sub
        End If

        ' Dano do Npc
        Damage = GetNpcDamage(MapNum, MapNpcNum)
        
        ' Dano - Defesa do Pokémon
        If GetPlayerEquipment(Index, weapon) > 0 Then
            Damage = Damage - RAND(1, (GetPlayerEquipmentPokeInfoStat(Index, weapon, Stats.Endurance) * 2))
        Else
            Damage = Damage * 0.1
        End If
        
        ' 75%~100% Dano
        Dano75 = Porcento(Damage, 75)
        Damage = Random(Dano75, Damage)
        
        ' Dano * 1.5 Critico
        If CanNpcCrit(MapNum, MapNpcNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(MapNpcNum).X * 32), (MapNpc(MapNum).Npc(MapNpcNum).Y * 32)
        End If

        'Concluir Ataque
        If Damage > 0 Then
            Call NpcAttackPlayer(MapNpcNum, Index, Damage)
        Else
            SendActionMsg MapNum, "Miss!", Cyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Index)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make surte then Player is evolving your pokémon!
    If Player(Index).EvolPermition = 1 Or TempPlayer(Index).EvolTimer > 0 Or Player(Index).EvolTimerStone > 0 Then
        MapNpc(MapNum).Npc(MapNpcNum).targetType = 0
        MapNpc(MapNum).Npc(MapNpcNum).target = 0
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum).Npc(MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(MapNum).Npc(MapNpcNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NpcNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNpcNum).Y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNpcNum).Y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).X) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim EXP As Long
    Dim MapNum As Long
    Dim I As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).Npc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(victim)
    Name = Trim$(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Name)
        
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(MapNum).Npc(MapNpcNum).stopRegen = True
    MapNpc(MapNum).Npc(MapNpcNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
    
        'Morte Sem Pokémons!
        If GetPlayerEquipment(victim, weapon) = 0 Then
            ' Say damage
            SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            
            ' send the sound
            SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNpcNum).Num
            
            ' kill player
            KillPlayer victim
            
            ' Set NPC target to 0
            MapNpc(MapNum).Npc(MapNpcNum).target = 0
            MapNpc(MapNum).Npc(MapNpcNum).targetType = 0
        Else
            If TempPlayer(victim).GymQntPoke > 0 Then
                If TempPlayer(victim).GymQntPoke = 1 Then
                    TempPlayer(victim).GymQntPoke = 0
                    Select Case TempPlayer(victim).InBattleGym
                    Case 1 'Brock
                        SendContagem victim, 0
                        PlayerWarp victim, 7, 12, 9
                        MapNpc(7).Npc(1).InBattle = False
                        PlayerMsg victim, "[Brock]: Infelizmente você perdeu a batalha fique mais forte e vamos batalhar novamente!", White
                    End Select
                Else
                        TempPlayer(victim).GymQntPoke = TempPlayer(victim).GymQntPoke - 1
                    End If
                End If
            End If
        
            If GetPlayerEquipmentPokeInfoVital(victim, weapon, 1) <= Damage Then
                Call SetPlayerEquipmentFelicidade(victim, weapon, GetPlayerEquipmentFelicidade(victim, weapon) - 1)
                Call SetPlayerEquipmentPokeInfoVital(victim, 0, weapon, 1)
                PlayerUnequipItem victim, weapon
            End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        'Call SetPlayerEquipmentPokeInfoVital(victim, GetPlayerEquipmentPokeInfoVital(victim, weapon, 1) - Damage, weapon, 1)
        
        Call SendVital(victim, Vitals.HP)
        Call SendAnimation(MapNum, Npc(MapNpc(GetPlayerMap(victim)).Npc(MapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNpcNum).Num
        
        ' Mandar o Dano
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' Parar Regeneração de HP
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
    End If
    
    ' Mandar Frame Attack para o Client
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SNpcAttack
    Buffer.WriteLong MapNpcNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim NpcNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        MapNum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(victim, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCriticalHit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        ' Conclusão da Formula
        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage)
        End If
        
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, weapon) > 0 Then
            If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If
    
    If Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_ARENA Then
        If GetPlayerEquipmentPokeInfoPokemon(victim, weapon) = 0 Then
            PlayerMsg attacker, "Você não pode atacar um treinador.", BrightRed
            Exit Function
        Else
            CanPlayerAttackPlayer = True
            Exit Function
        End If
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "Está área não é possivel Batalhar!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function
    
     If Player(attacker).ORG > 0 Then
        If Player(attacker).ORG = Player(victim).ORG Then
            Exit Function
        End If
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim EXP As Long
    Dim n As Long
    Dim I As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, weapon) > 0 Then
        n = GetPlayerEquipment(attacker, weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        ' purge target info of anyone who targetted dead guy
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And IsConnected(I) Then
                If Player(I).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(I).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(I).target = victim Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendTarget I
                        End If
                    End If
                End If
            End If
        Next
        
        ' Verifica se está em uma arena
        If Map(GetPlayerMap(victim)).Moral = MAP_MORAL_ARENA Then
        
        Select Case TempPlayer(victim).LutandoT
         Case 1
         
         If TempPlayer(victim).LutQntPoke = 0 Then
            SendArenaStatus TempPlayer(victim).LutandoA, 0
            
            If GetPlayerIP(victim) = GetPlayerIP(attacker) Then
                PlayerMsg attacker, "Você venceu a batalha mas não ganhou Vitórias por ter o mesmo IP do Jogador: " & Trim$(GetPlayerName(victim)), BrightRed
                PlayerMsg victim, "Você perdeu a Batalha mas não ganhou Derrotas por ter o mesmo IP do jogador: " & Trim$(GetPlayerName(attacker)), BrightRed
            Else
                Player(victim).Derrotas = Player(victim).Derrotas + 1
                Player(attacker).Vitorias = Player(attacker).Vitorias + 1
                SetPlayerEquipmentFelicidade attacker, weapon, GetPlayerEquipmentFelicidade(attacker, weapon) + 2
                SetPlayerEquipmentFelicidade victim, weapon, GetPlayerEquipmentFelicidade(victim, weapon) - 2
            End If
            
            TempPlayer(victim).Lutando = 0
            TempPlayer(victim).LutandoA = 0
            TempPlayer(victim).LutandoT = 0
            
            PlayerUnequipItem victim, weapon
            PlayerWarp victim, Player(victim).MyMap(1), Player(victim).MyMap(2), Player(victim).MyMap(3)
            
            TempPlayer(attacker).Lutando = 0
            TempPlayer(attacker).LutandoA = 0
            TempPlayer(attacker).LutandoT = 0
            
            PlayerUnequipItem attacker, weapon
            PlayerWarp attacker, Player(attacker).MyMap(1), Player(attacker).MyMap(2), Player(attacker).MyMap(3)
            
            Call GlobalMsg(GetPlayerName(attacker) & " venceu a luta contra: " & GetPlayerName(victim), White)
         Else
                'Setar Felicidade
                If Not GetPlayerIP(victim) = GetPlayerIP(attacker) Then
                    SetPlayerEquipmentFelicidade attacker, weapon, GetPlayerEquipmentFelicidade(attacker, weapon) + 2
                    SetPlayerEquipmentFelicidade victim, weapon, GetPlayerEquipmentFelicidade(victim, weapon) - 2
                End If
                
                TempPlayer(victim).LutQntPoke = TempPlayer(victim).LutQntPoke - 1
                PlayerMsg victim, "Seu " & Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(victim, weapon)).Name) & " foi derrotado, escolha outro pokémon em 30 segundos.", BrightGreen
                Call SetPlayerEquipmentPokeInfoVital(victim, 0, weapon, 1)
                PlayerUnequipItem victim, weapon
                TempPlayer(victim).SwitPoke = 30000 + GetTickCount
            End If
         
        Case Else
            SetPlayerVital victim, HP, 0
            PlayerMsg victim, "Seu " & Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(victim, weapon)).Name) & " foi derrotado.", BrightGreen
            PlayerUnequipItem victim, weapon
        End Select
        
        If TempPlayer(victim).inParty > 0 Then
          Select Case TempPlayer(Party(TempPlayer(victim).inParty).Leader).LutandoT
           Case 2
            SendArenaStatus TempPlayer(victim).LutandoA, 0
            PlayerWarp victim, 350, 11, 8
            If Party(TempPlayer(victim).inParty).PT > 0 Then
                Party(TempPlayer(victim).inParty).PT = Party(TempPlayer(victim).inParty).PT - 1
                PartyMsg TempPlayer(victim).inParty, "Seu grupo possui agora: " & Party(TempPlayer(victim).inParty).PT & " pontos no combate", BrightGreen
                PartyMsg TempPlayer(attacker).inParty, "Seu grupo dimnui mais um ponto do grupo inimigo", BrightGreen
            End If
            
            If Party(TempPlayer(victim).inParty).PT = 0 Then
            For I = 1 To Player_HighIndex
                If IsPlaying(I) = True Then
                    If TempPlayer(I).inParty = TempPlayer(victim).inParty Then
                        Player(I).Derrotas = Player(I).Derrotas + 1
                    End If
                End If
            Next
            
            For I = 1 To Player_HighIndex
                If IsPlaying(I) = True Then
                    If TempPlayer(I).inParty = TempPlayer(attacker).inParty Then
                        Player(I).Vitorias = Player(I).Vitorias + 1
                        PlayerWarp attacker, 350, 11, 8
                        Party(TempPlayer(attacker).inParty).PT = 0
                    End If
                End If
            Next
            GlobalMsg "O grupo do jogador: " & GetPlayerName(attacker) & " venceu um combate , contra o grupo do jogador : " & GetPlayerName(victim), BrightCyan
            End If
            Case Else
           End Select
          End If
          
        Else
            ChecarTarefasAtuais attacker, QUEST_TYPE_KILLPLAYER, victim
            Call OnDeath(victim)
        End If
        
    Else
    
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer victim, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Player victim, SpellNum, attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal Index As Long, ByVal spellslot As Long)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    SpellNum = GetPlayerSpell(Index, spellslot)
    MapNum = GetPlayerMap(Index)
    
    If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(Index).targetType
    target = TempPlayer(Index).target
    Range = Spell(SpellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg Index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(MapNum).Npc(target).X, MapNpc(MapNum).Npc(target).Y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index
       ' SendActionMsg MapNum, Trim$(Spell(spellnum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        TempPlayer(Index).spellBuffer.Spell = spellslot
        TempPlayer(Index).spellBuffer.Timer = GetTickCount
        TempPlayer(Index).spellBuffer.target = TempPlayer(Index).target
        TempPlayer(Index).spellBuffer.tType = TempPlayer(Index).targetType
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim I As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim X As Long, Y As Long
    
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    Dim linha As Long, linha2 As Long, AnimL As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    SpellNum = GetPlayerSpell(Index, spellslot)
    MapNum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    
    ' set the vital
        Vital = GetPlayerEquipmentPokeInfoStat(Index, weapon, 3) 'Pegar o Atq Especial do Pokémon :P
        AoE = Spell(SpellNum).AoE
        Range = Spell(SpellNum).Range
        
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, Index, Vital, SpellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, Index, Vital, SpellNum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    SetPlayerDir Index, Spell(SpellNum).Dir
                    PlayerWarp Index, Spell(SpellNum).Map, Spell(SpellNum).X, Spell(SpellNum).Y
                    SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    DidCast = True
                Case SPELL_TYPE_SCRIPT
                   SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                   ScriptedSpell Index, Spell(SpellNum).Script, SpellNum
                   DidCast = True
                 Case SPELL_TYPE_FLY
                    CanFly Index
                    If IsCanFly(Index) = True Then
                        DidCast = True
                    End If
                    
                    If Not Spell(SpellNum).Sound = vbNullString Then
                        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                    End If
            End Select
            
                    
        Case 1, 3 ' self-cast AOE & targetted AOE 'deu erro com certeza :3
            If SpellCastType = 1 Then
                X = GetPlayerX(Index)
                Y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(target)
                    Y = GetPlayerY(target)
                Else
                    X = MapNpc(MapNum).Npc(target).X
                    Y = MapNpc(MapNum).Npc(target).Y
                End If
                
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    SendClearSpellBuffer Index
                End If
            End If
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If I <> Index Then
                                If GetPlayerMap(I) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, X, Y, GetPlayerX(I), GetPlayerY(I)) Then
                                        If CanPlayerAttackPlayer(Index, I, True) Then
                                            SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, I
                                            
                                            If GetPlayerEquipmentPokeInfoPokemon(I, weapon) > 0 Then
                                            Select Case DanoElemental(Spell(SpellNum).Element, Pokemon(GetPlayerEquipmentPokeInfoPokemon(I, weapon)).Tipo(1))
                                            Case 0 'Normal
                                            Vital = Vital 'Dano Normal
                                            Case 1
                                            Vital = Vital / 2 'Metade do Dano
                                            Case 2
                                            Vital = Vital * 2 'Dobro do Dano
                                            Case 3
                                            Vital = 0 'Dano 0
                                            End Select
                                            End If
                                            
                                            PlayerAttackPlayer Index, I, Vital, SpellNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(I).Num > 0 Then
                            If MapNpc(MapNum).Npc(I).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, MapNpc(MapNum).Npc(I).X, MapNpc(MapNum).Npc(I).Y) Then
                                    If CanPlayerAttackNpc(Index, I, True) Then
                                        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, I
                                        
                                        If Npc(MapNpc(MapNum).Npc(I).Num).Pokemon > 0 Then
                                        Select Case DanoElemental(Spell(SpellNum).Element, Pokemon(Npc(MapNpc(MapNum).Npc(I).Num).Pokemon).Tipo(1))
                                        Case 0 'Normal
                                        Vital = Vital 'Dano Normal
                                        Case 1
                                        Vital = Vital / 1.5
                                        Case 2
                                        Vital = Vital * 1.5 'Dobro do Dano
                                        Case 3
                                        Vital = 0 'Dano 0
                                        End Select
                                        Else
                                            Vital = Vital
                                        End If
                                        
                                        PlayerAttackNpc Index, I, Vital, SpellNum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                    
                    DidCast = True
                     For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If I <> Index Then
                                If GetPlayerMap(I) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, X, Y, GetPlayerX(I), GetPlayerY(I)) Then
                                        SpellPlayer_Effect VitalType, increment, I, Vital, SpellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(I).Num > 0 Then
                            If MapNpc(MapNum).Npc(I).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, MapNpc(MapNum).Npc(I).X, MapNpc(MapNum).Npc(I).Y) Then
                                    SpellNpc_Effect VitalType, increment, I, Vital, SpellNum, MapNum
                                End If
                            End If
                        End If
                    Next
                    
                    Case SPELL_TYPE_LINEAR
                   '/// - MAGIA LINEAR AVANÇADA - ///
                  For linha = 1 To Spell(SpellNum).AoE
                   Select Case GetPlayerDir(Index)
                     Case DIR_UP
                      If Not GetPlayerY(Index) - linha < 0 Then
                        SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, GetPlayerX(Index), GetPlayerY(Index) - linha
                        UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index) - linha
                      End If
                     Case DIR_DOWN
                      If Not GetPlayerY(Index) + linha > Map(MapNum).MaxY Then
                        SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, GetPlayerX(Index), GetPlayerY(Index) + linha
                        UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index) + linha
                      End If
                     Case DIR_LEFT
                      If Not GetPlayerX(Index) - linha < 0 Then
                        SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, GetPlayerX(Index) - linha, GetPlayerY(Index)
                        UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) - linha, GetPlayerY(Index)
                      End If
                     Case DIR_RIGHT
                      If Not GetPlayerX(Index) + linha > Map(MapNum).MaxX Then
                        SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, GetPlayerX(Index) + linha, GetPlayerY(Index)
                        UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) + linha, GetPlayerY(Index)
                      End If
                   End Select

                   '/// - Animação Lateral - ///
                   If Spell(SpellNum).AnimL > 0 Then
                    AnimL = Spell(SpellNum).AnimL
                   Else
                    AnimL = Spell(SpellNum).SpellAnim
                   End If

                   '/// - Magia Lateral - ///
                   If Spell(SpellNum).Tamanho > 0 Then
                    If linha > 1 Then
                     For linha2 = 1 To Spell(SpellNum).Tamanho
                      Select Case GetPlayerDir(Index)
                        Case DIR_UP
                           If Not GetPlayerY(Index) - linha < 0 Then
                             If Not GetPlayerX(Index) - linha2 < 0 Then
                               SendAnimation GetPlayerMap(Index), AnimL, GetPlayerX(Index) - linha2, GetPlayerY(Index) - linha
                               UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) - linha2, GetPlayerY(Index) - linha
                             End If
                           
                             If Not GetPlayerX(Index) + linha2 > Map(MapNum).MaxX Then
                               SendAnimation GetPlayerMap(Index), AnimL, GetPlayerX(Index) + linha2, GetPlayerY(Index) - linha
                               UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) + linha2, GetPlayerY(Index) - linha
                             End If
                           End If
                        Case DIR_DOWN
                           If Not GetPlayerY(Index) + linha > Map(MapNum).MaxY Then
                             If Not GetPlayerX(Index) + linha2 > Map(MapNum).MaxX Then
                               SendAnimation GetPlayerMap(Index), AnimL, GetPlayerX(Index) + linha2, GetPlayerY(Index) + linha
                               UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) + linha2, GetPlayerY(Index) + linha
                             End If
                           
                              If Not GetPlayerX(Index) - linha2 < 0 Then
                               SendAnimation GetPlayerMap(Index), AnimL, GetPlayerX(Index) - linha2, GetPlayerY(Index) + linha
                               UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) - linha2, GetPlayerY(Index) + linha
                             End If
                           End If
                        Case DIR_LEFT
                           If Not GetPlayerX(Index) - linha < 0 Then
                             If Not GetPlayerY(Index) - linha2 < 0 Then
                               SendAnimation GetPlayerMap(Index), AnimL, GetPlayerX(Index) - linha, GetPlayerY(Index) - linha2
                               UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) - linha, GetPlayerY(Index) - linha2
                             End If
                           
                             If Not GetPlayerY(Index) + linha2 > Map(MapNum).MaxY Then
                               SendAnimation GetPlayerMap(Index), AnimL, GetPlayerX(Index) - linha, GetPlayerY(Index) + linha2
                               UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) - linha, GetPlayerY(Index) + linha2
                             End If
                           End If
                        Case DIR_RIGHT
                            If Not GetPlayerX(Index) + linha > Map(MapNum).MaxX Then
                              If Not GetPlayerY(Index) + linha2 > Map(MapNum).MaxY Then
                                SendAnimation GetPlayerMap(Index), AnimL, GetPlayerX(Index) + linha, GetPlayerY(Index) + linha2
                                UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) + linha, GetPlayerY(Index) + linha2
                              End If
                              
                              If Not GetPlayerY(Index) - linha2 < 0 Then
                                SendAnimation GetPlayerMap(Index), AnimL, GetPlayerX(Index) + linha, GetPlayerY(Index) - linha2
                                UsarMagiaLinear Index, SpellNum, Vital, GetPlayerMap(Index), GetPlayerX(Index) + linha, GetPlayerY(Index) - linha2
                              End If
                            End If
                      End Select
                     Next
                    End If
                   End If
                  Next
                  
                  DidCast = True
                        
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
            
            If targetType = TARGET_TYPE_PLAYER Then
                X = GetPlayerX(target)
                Y = GetPlayerY(target)
            Else
                X = MapNpc(MapNum).Npc(target).X
                Y = MapNpc(MapNum).Npc(target).Y
            End If
                
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                
                                If GetPlayerEquipmentPokeInfoPokemon(I, weapon) > 0 Then
                                Select Case DanoElemental(Spell(SpellNum).Element, Pokemon(GetPlayerEquipmentPokeInfoPokemon(target, weapon)).Tipo(1))
                                Case 0 'Normal
                                Vital = Vital 'Dano Normal
                                Case 1
                                Vital = Vital / 2 'Metade do Dano
                                Case 2
                                Vital = Vital * 2 'Dobro do Dano
                                Case 3
                                Vital = 0 'Dano 0
                                End Select
                                End If
                                
                                PlayerAttackPlayer Index, target, Vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                
                                Select Case DanoElemental(Spell(SpellNum).Element, Pokemon(Npc(MapNpc(MapNum).Npc(target).Num).Pokemon).Tipo(1))
                                Case 0 'Normal
                                Vital = Vital 'Dano Normal
                                Case 1
                                Vital = Vital / 1.5 'Metade do Dano
                                Case 2
                                Vital = Vital * 1.5 'Dobro do Dano
                                Case 3
                                Vital = 0 'Dano 0
                                End Select
                                
                                PlayerAttackNpc Index, target, Vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    End If
                    
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    End If
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(Index, target, True) Then
                                SpellPlayer_Effect VitalType, increment, target, Vital, SpellNum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, target, Vital, SpellNum
                        End If
                    Else
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(Index, target, True) Then
                                SpellNpc_Effect VitalType, increment, target, Vital, SpellNum, MapNum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, target, Vital, SpellNum, MapNum
                        End If
                    End If
            End Select
    End Select
    
    If DidCast Then
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPCost)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
        
        TempPlayer(Index).SpellCD(spellslot) = GetTickCount + (Spell(SpellNum).CDTime * 1000)
        Call SendCooldown(Index, spellslot)
        'SendActionMsg MapNum, Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        'Mandar msg
        If GetPlayerEquipment(Index, weapon) > 0 Then
            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                SendActionMsg MapNum, Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " use " & Trim$(Spell(SpellNum).Name) & "!", White, ACTIONMSG_STATIC, Player(Index).TPX * 32 + 8, Player(Index).TPY * 32 - 16
            End If
        End If
        
        Call SendAttack(Index)
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
        
        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Player Index, SpellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
        End If
        
        SendVital Index, Vital
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    End If
    
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal MapNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
        SendActionMsg MapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(Index).X * 32, MapNpc(MapNum).Npc(Index).Y * 32
        
        ' send the sound
        SendMapSound Index, MapNpc(MapNum).Npc(Index).X, MapNpc(MapNum).Npc(Index).Y, SoundEntity.seSpell, SpellNum
        
        If increment Then
            MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Npc MapNum, Index, SpellNum
            End If
        ElseIf Not increment Then
            MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) - Damage
        End If
        
          ' send update
        SendMapNpcVitals MapNum, Index
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(I)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal Index As Long, ByVal SpellNum As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(I)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(Index).DoT(I)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(Index).HoT(I)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)
With TempPlayer(Index).DoT(dotNum)
If .Used And .Spell > 0 Then
' time to tick?
If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
If CanPlayerAttackPlayer(.Caster, Index, True) Then
PlayerAttackPlayer .Caster, Index, RAND(1, GetSpellBaseStat(.Caster, .Spell))
End If
.Timer = GetTickCount
' check if DoT is still active - if player died it'll have been purged
If .Used And .Spell > 0 Then
' destroy DoT if finished
If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
.Used = False
.Spell = 0
.Timer = 0
.Caster = 0
.StartTime = 0
End If
End If
End If
End If
End With
End Sub
Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)
    With TempPlayer(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
               If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                    SendActionMsg Player(Index).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(Index).X * 32, Player(Index).Y * 32
                    SetPlayerVital Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + Spell(.Spell).Vital
                    Call SendVital(Index, Vitals.HP)
                Else
                    SendActionMsg Player(Index).Map, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, Player(Index).X * 32, Player(Index).Y * 32
                    SetPlayerVital Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + Spell(.Spell).Vital
                    Call SendVital(Index, Vitals.MP)
               End If
               
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal dotNum As Long)
With MapNpc(MapNum).Npc(Index).DoT(dotNum)
If .Used And .Spell > 0 Then
' time to tick?
If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
If CanPlayerAttackNpc(.Caster, Index, True) Then
PlayerAttackNpc .Caster, Index, RAND(1, GetSpellBaseStat(.Caster, .Spell)), , True
End If
.Timer = GetTickCount
' check if DoT is still active - if NPC died it'll have been purged
If .Used And .Spell > 0 Then
' destroy DoT if finished
If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
.Used = False
.Spell = 0
.Timer = 0
.Caster = 0
.StartTime = 0
End If
End If
End If
End If
End With
End Sub

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal hotNum As Long)
    With MapNpc(MapNum).Npc(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                    SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(Index).X * 32, MapNpc(MapNum).Npc(Index).Y * 32
                    MapNpc(MapNum).Npc(Index).Vital(Vitals.HP) = MapNpc(MapNum).Npc(Index).Vital(Vitals.HP) + Spell(.Spell).Vital
                    SendMapNpcVitals MapNum, MapNpc(MapNum).Npc(Index).Num
                Else
                    SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(Index).X * 32, MapNpc(MapNum).Npc(Index).Y * 32
                    MapNpc(MapNum).Npc(Index).Vital(Vitals.MP) = MapNpc(MapNum).Npc(Index).Vital(Vitals.MP) + Spell(.Spell).Vital
                    SendMapNpcVitals MapNum, MapNpc(MapNum).Npc(Index).Num
                End If
                
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(Index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal MapNum As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(MapNum).Npc(Index).StunDuration = Spell(SpellNum).StunDuration
        MapNpc(MapNum).Npc(Index).StunTimer = GetTickCount
    End If
End Sub

Function UsarMagiaLinear(ByVal Index As Integer, ByVal SpellNum As Integer, ByVal Vital As Long, ByVal Mapa As Long, ByVal X As Byte, ByVal Y As Byte)
Dim I As Long

    'Loop Global Npc
    For I = 1 To MAX_MAP_NPCS
      If MapNpc(Mapa).Npc(I).Num > 0 And MapNpc(Mapa).Npc(I).X = X And MapNpc(Mapa).Npc(I).Y = Y And MapNpc(Mapa).Npc(I).Vital(HP) > 0 Then
        If CanPlayerAttackNpc(Index, I, True) Then
        
        'Evitar OverFlow
        If Npc(MapNpc(Mapa).Npc(I).Num).Pokemon = 0 Then Exit Function
        
            Select Case DanoElemental(Spell(SpellNum).Element, Pokemon(Npc(MapNpc(Mapa).Npc(I).Num).Pokemon).Tipo(1))
            Case 0 'Normal
            Vital = Vital 'Dano Normal
            Case 1
            Vital = Vital / 2 'Metade do Dano
            Case 2
            Vital = Vital * 2 'Dobro do Dano
            Case 3
            Vital = 0 'Dano 0
            End Select
        
          PlayerAttackNpc Index, I, Vital, SpellNum
        End If
      End If
    Next
    
    'Loop Global Player
    For I = 1 To Player_HighIndex
      If IsPlaying(I) Then
        If GetPlayerMap(I) = Mapa And GetPlayerX(I) = X And GetPlayerY(I) = Y Then
          If CanPlayerAttackPlayer(Index, I, True) Then
          
            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
            Select Case DanoElemental(Spell(SpellNum).Element, Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Tipo(1))
            Case 0 'Normal
            Vital = Vital 'Dano Normal
            Case 1
            Vital = Vital / 2 'Metade do Dano
            Case 2
            Vital = Vital * 2 'Dobro do Dano
            Case 3
            Vital = 0 'Dano 0
            End Select
            End If
          
            PlayerAttackPlayer Index, I, Vital, SpellNum
          End If
        End If
      End If
    Next
    
    CheckResource Index, X, Y, SpellNum
End Function

Function DanoElemental(ByVal EAttacker As Long, ByVal EVictim As Long) As Byte
'Damage Description:
'0.Normal '1.Metade '2.Dobro '3.Nada

'Elemental Numbers:
'1.Fogo '2.Água '3.Grama '4.Elétrico '5.Terrestre '6.Normal '7.Pedra '8.Voador
'9.Venenoso '10.Inseto '11.Noturno '12.Fantasma '13.Psíquico '14.Dragão
'15.Metálico '16.Gelo '17.Lutador '18.Fada

If EAttacker = 1 Then 'Attacker Fire
Select Case EVictim
Case 1, 2, 7, 14 'Metade do Dano
DanoElemental = 1
Case 3, 16, 10, 15 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function 'Sair da Função
End If

If EAttacker = 2 Then 'Attacker Water
Select Case EVictim
Case 2, 3, 14 'Metade do Dano
DanoElemental = 1
Case 1, 5, 7 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 3 Then 'Attacker Grass
Select Case EVictim
Case 1, 3, 8, 9, 10, 14, 15 'Metade do Dano
DanoElemental = 1
Case 2, 5, 7 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 4 Then 'Attacker Eletric
Select Case EVictim
Case 3, 4, 14 'Metade do Dano
DanoElemental = 1
Case 2, 8 'Dobro do Dano
DanoElemental = 2
Case 5 'Dano 0
DanoElemental = 3
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 5 Then 'Attacker Ground
Select Case EVictim
Case 3, 10 'Metade do Dano
DanoElemental = 1
Case 1, 4, 7, 9, 15 'Dobro do Dano
DanoElemental = 2
Case 8 'Dano 0
DanoElemental = 3
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 6 Then 'Attacker Normal
Select Case EVictim
Case 7, 15 'Metade do Dano
DanoElemental = 1
Case 12 'Dano 0
DanoElemental = 3
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 7 Then 'Attacker Pedra
Select Case EVictim
Case 5, 17 'Metade do Dano
DanoElemental = 1
Case 1, 16, 8, 10 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 8 Then 'Attacker Fly
Select Case EVictim
Case 4, 7, 15 'Metade do Dano
DanoElemental = 1
Case 3, 10, 17 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 9 Then 'Attacker Poison
Select Case EVictim
Case 5, 7, 9, 12 'Metade do Dano
DanoElemental = 1
Case 3, 18 'Dobro do Dano
DanoElemental = 2
Case 15 'Dano 0
DanoElemental = 3
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 10 Then 'Attacker Bug
Select Case EVictim
Case 1, 8, 9, 12, 15, 17, 18 'Metade do Dano
DanoElemental = 1
Case 3, 11, 13 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 11 Then 'Attacker Darkness
Select Case EVictim
Case 17, 11, 18 'Metade do Dano
DanoElemental = 1
Case 12, 13 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 12 Then 'Attacker Ghost
Select Case EVictim
Case 14 'Metade do Dano
DanoElemental = 1
Case 8, 13 'Dobro do Dano
DanoElemental = 2
Case 6 'Dano 0
DanoElemental = 3
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 13 Then 'Attacker Psycho
Select Case EVictim
Case 13, 15 'Metade do Dano
DanoElemental = 1
Case 9, 17 'Dobro do Dano
DanoElemental = 2
Case 11 'Dano 0
DanoElemental = 3
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 14 Then 'Attacker Dragon
Select Case EVictim
Case 15 'Metade do Dano
DanoElemental = 1
Case 14 'Dobro do Dano
DanoElemental = 2
Case 18 'Dano 0
DanoElemental = 3
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 15 Then 'Attacker Steel
Select Case EVictim
Case 1, 2, 4, 15 'Metade do Dano
DanoElemental = 1
Case 7, 16, 18 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 16 Then 'Attacker Ice
Select Case EVictim
Case 1, 2, 15, 16 'Metade do Dano
DanoElemental = 1
Case 3, 5, 8, 14 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 17 Then 'Attacker Fighter
Select Case EVictim
Case 8, 9, 10, 13, 18 'Metade do Dano
DanoElemental = 1
Case 1, 7, 11, 15, 16 'Dobro do Dano
DanoElemental = 2
Case 12 'Dano 0
DanoElemental = 3
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

If EAttacker = 18 Then 'Attacker Fairy
Select Case EVictim
Case 1, 5, 15 'Metade do Dano
DanoElemental = 1 '
Case 9, 11, 14 'Dobro do Dano
DanoElemental = 2
Case Else
DanoElemental = 0 'Dano Normal
End Select
Exit Function
End If

End Function

Function IsCanFly(ByVal Index As Long) As Boolean
    If GetPlayerFlying(Index) = 1 Then ' Voando
    
        ' Verifica se pode descer
        If CanBlockedTile(Index, GetPlayerX(Index), GetPlayerY(Index), True) = True Then
            IsCanFly = False
            Exit Function
        End If
        
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_FLYAVOID Then
            IsCanFly = False
            Exit Function
        End If
        
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_WATER Then
            IsCanFly = False
            Exit Function
        End If
        
        IsCanFly = True
    Else
        IsCanFly = True
    End If

End Function

Public Sub CanFly(ByVal Index As Long)
    If GetPlayerFlying(Index) = 1 Then ' Voando
    
        ' Verifica se pode descer
        If CanBlockedTile(Index, GetPlayerX(Index), GetPlayerY(Index), True) = True Then
            Call PlayerMsg(Index, "Você não pode descer aqui!", Red)
            Exit Sub
        End If
        
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_FLYAVOID Then
            Call PlayerMsg(Index, "Você não pode descer aqui!", Red)
            Exit Sub
        End If
        
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_WATER Then
            Call PlayerMsg(Index, "Você não pode descer aqui!", Red)
            Exit Sub
        End If
            
        Call SetPlayerFlying(Index, 0)
    Else ' Normal
        Call SetPlayerFlying(Index, 1)
    End If

    Call SendPlayerData(Index)
End Sub


Public Function CanPlayerAtractHit(ByVal Index As Long) As Boolean
Dim Rnd As Byte, Extra As Byte
Rnd = Random(1, 99)

'Sair Caso não tenha pokémon! = 0
If GetPlayerEquipment(Index, weapon) = 0 Then Exit Function

'Sp Def
Extra = Int(GetPlayerEquipmentPokeInfoStat(Index, weapon, 5) * 0.05)

'Sp Berry Def
Extra = Extra + Int(GetPlayerEquipmentBerry(Index, weapon, 5) * 0.1)

'Se Chance for Menor que o valor Random o Ataque vai ser 10% do Ataque Total
If (Random(1, 44) + Extra) < Rnd Then
    CanPlayerAtractHit = True
End If

End Function

Public Function CanPlayerConfusionHit(ByVal Index As Long) As Boolean
Dim Rnd As Byte, Extra As Byte
Rnd = Random(1, 99)

'Sair Caso não tenha pokémon! = 0
If GetPlayerEquipment(Index, weapon) = 0 Then Exit Function

'Atq * 0.01
Extra = Int(GetPlayerEquipmentPokeInfoStat(Index, weapon, 1) * 0.01)

'Se Chance for Menor que o valor Random o
If (Random(1, 34) + Extra) < Rnd Then
    CanPlayerConfusionHit = True
End If

End Function

' /////////////////////
' // General Purpose //
' /////////////////////

Public Sub ChecarTarefasAtuais(ByVal Index As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim I As Long
    
    For I = 1 To MAX_QUESTS
        If Player(Index).Quests(I).Status > 0 And Player(Index).Quests(I).Status < 3 Then
            If Quest(I).Task(Player(Index).Quests(I).Part).Type = TaskType Then
                ChecarQuest Index, I, TaskType, TargetIndex
            End If
        End If
    Next
    
End Sub
