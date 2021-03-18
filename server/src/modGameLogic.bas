Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim I As Long
    FindOpenPlayerSlot = 0

    For I = 1 To MAX_PLAYERS

        If Not IsConnected(I) Then
            FindOpenPlayerSlot = I
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim I As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For I = 1 To MAX_MAP_ITEMS

        If MapItem(MapNum, I).Num = 0 Then
            FindOpenMapItemSlot = I
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim I As Long
    TotalOnlinePlayers = 0

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(I)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(I), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal Pokemon As Long, Optional ByVal Pokeball As Long, Optional ByVal Level As Long, Optional ByVal EXP As Long, Optional ByVal VitalHP As Long, Optional ByVal VitalMP As Long, Optional ByVal MaxVitalHp As Long, Optional ByVal MaxVitalMp As Long, Optional ByVal StatStr As Long, Optional ByVal StatAgi As Long, Optional ByVal StatEnd As Long, Optional ByVal StatInt As Long, Optional ByVal StatWill As Long, Optional ByVal Spell1 As Long, Optional ByVal Spell2 As Long, Optional ByVal Spell3 As Long, Optional ByVal Spell4 As Long, _
Optional ByVal Ngt1 As Long, Optional ByVal Ngt2 As Long, Optional ByVal Ngt3 As Long, Optional ByVal Ngt4 As Long, _
Optional ByVal Ngt5 As Long, Optional ByVal Ngt6 As Long, Optional ByVal Ngt7 As Long, Optional ByVal Ngt8 As Long, _
Optional ByVal Ngt9 As Long, Optional ByVal Ngt10 As Long, Optional ByVal Ngt11 As Long, Optional ByVal Felicidade As Long, _
Optional ByVal Sexo As Byte, Optional ByVal Shiny As Long, Optional ByVal Bry1 As Long, Optional ByVal Bry2 As Long, Optional ByVal Bry3 As Long, _
Optional ByVal Bry4 As Long, Optional ByVal Bry5 As Long)

    Dim I As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    I = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(I, ItemNum, ItemVal, MapNum, x, Y, playerName, False, Pokemon, Pokeball, Level, EXP, VitalHP, VitalMP, MaxVitalHp, MaxVitalMp, StatStr, StatAgi, StatEnd, StatInt, StatWill, Spell1, Spell2, Spell3, Spell4, Ngt1, Ngt2, Ngt3, Ngt4, Ngt5, Ngt6, Ngt7, Ngt8, Ngt9, Ngt10, Ngt11, Felicidade, Sexo, Shiny, Bry1, Bry2, Bry3, Bry4, Bry5)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True, Optional ByVal Pokemon As Long, Optional ByVal Pokeball As Long, Optional ByVal Level As Long, Optional ByVal EXP As Long, Optional ByVal VitalHP As Long, Optional ByVal VitalMP As Long, Optional ByVal MaxVitalHp As Long, Optional ByVal MaxVitalMp As Long, Optional ByVal StatStr As Long, Optional ByVal StatAgi As Long, Optional ByVal StatEnd As Long, Optional ByVal StatInt As Long, Optional ByVal StatWill As Long, Optional ByVal Spell1 As Long, Optional ByVal Spell2 As Long, Optional ByVal Spell3 As Long, Optional ByVal Spell4 As Long, _
Optional ByVal Ngt1 As Long, Optional ByVal Ngt2 As Long, _
Optional ByVal Ngt3 As Long, Optional ByVal Ngt4 As Long, _
Optional ByVal Ngt5 As Long, Optional ByVal Ngt6 As Long, _
Optional ByVal Ngt7 As Long, Optional ByVal Ngt8 As Long, _
Optional ByVal Ngt9 As Long, Optional ByVal Ngt10 As Long, _
Optional ByVal Ngt11 As Long, Optional ByVal Felicidade As Long, _
Optional ByVal Sexo As Byte, Optional ByVal Shiny As Long, _
Optional ByVal Bry1 As Long, Optional ByVal Bry2 As Long, Optional ByVal Bry3 As Long, Optional ByVal Bry4 As Long, Optional ByVal Bry5 As Long)
    
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    I = MapItemSlot

    If I <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then

            MapItem(MapNum, I).Num = ItemNum
            MapItem(MapNum, I).Value = ItemVal
            MapItem(MapNum, I).x = x
            MapItem(MapNum, I).Y = Y
            MapItem(MapNum, I).PokeInfo.Pokemon = Pokemon
            MapItem(MapNum, I).PokeInfo.Pokeball = Pokeball
            MapItem(MapNum, I).PokeInfo.Level = Level
            MapItem(MapNum, I).PokeInfo.EXP = EXP
            
            MapItem(MapNum, I).PokeInfo.Vital(1) = VitalHP
            MapItem(MapNum, I).PokeInfo.Vital(2) = VitalMP
            MapItem(MapNum, I).PokeInfo.MaxVital(1) = MaxVitalHp
            MapItem(MapNum, I).PokeInfo.MaxVital(2) = MaxVitalMp
            
            MapItem(MapNum, I).PokeInfo.Stat(1) = StatStr
            MapItem(MapNum, I).PokeInfo.Stat(2) = StatEnd
            MapItem(MapNum, I).PokeInfo.Stat(3) = StatInt
            MapItem(MapNum, I).PokeInfo.Stat(4) = StatAgi
            MapItem(MapNum, I).PokeInfo.Stat(5) = StatWill
            
            MapItem(MapNum, I).PokeInfo.Spells(1) = Spell1
            MapItem(MapNum, I).PokeInfo.Spells(2) = Spell2
            MapItem(MapNum, I).PokeInfo.Spells(3) = Spell3
            MapItem(MapNum, I).PokeInfo.Spells(4) = Spell4
            
            MapItem(MapNum, I).PokeInfo.Negatives(1) = Ngt1
            MapItem(MapNum, I).PokeInfo.Negatives(2) = Ngt2
            MapItem(MapNum, I).PokeInfo.Negatives(3) = Ngt3
            MapItem(MapNum, I).PokeInfo.Negatives(4) = Ngt4
            MapItem(MapNum, I).PokeInfo.Negatives(5) = Ngt5
            MapItem(MapNum, I).PokeInfo.Negatives(6) = Ngt6
            MapItem(MapNum, I).PokeInfo.Negatives(7) = Ngt7
            MapItem(MapNum, I).PokeInfo.Negatives(8) = Ngt8
            MapItem(MapNum, I).PokeInfo.Negatives(9) = Ngt9
            MapItem(MapNum, I).PokeInfo.Negatives(10) = Ngt10
            MapItem(MapNum, I).PokeInfo.Negatives(11) = Ngt11
            
            MapItem(MapNum, I).PokeInfo.Felicidade = Felicidade
            MapItem(MapNum, I).PokeInfo.Sexo = Sexo
            MapItem(MapNum, I).PokeInfo.Shiny = Shiny
            
            ' send to map
            SendSpawnItemToMap MapNum, I
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call SpawnMapItems(I)
    Next

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim x As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(x, Y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(x, Y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).Tile(x, Y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, Y).Data1, 1, MapNum, x, Y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, Y).Data1, Map(MapNum).Tile(x, Y).Data2, MapNum, x, Y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Function Porcento(ByVal ValorTotal As Long, ByVal Porcentagem As Long) As Long
    Porcento = (ValorTotal * Porcentagem) / 100
End Function

Public Sub SpawnPokeGym(ByVal MapNpcNum As Byte, ByVal MapNum As Integer, ByVal NpcNum As Byte, ByVal x As Integer, ByVal Y As Integer, ByVal DirNum As Byte, ByVal Sex As Boolean, ByVal Level As Integer)
Dim Buffer As clsBuffer, I As Long
Dim Sumir As Boolean

    If NpcNum > 0 Then
        Map(MapNum).Npc(2) = NpcNum
        
        With MapNpc(MapNum).Npc(MapNpcNum)
            .Num = NpcNum
            .x = x
            .Y = Y
            .Dir = DirNum
            .Sexo = Sex
            .Shiny = False
            .Level = Level
            .target = 0
            .targetType = 0
            .Desmaiado = False
            .Vital(1) = Pokemon(NpcNum).Vital(1) + (Level * 5)
            .Vital(2) = Pokemon(NpcNum).Vital(2) + (Level * 5)
            .Stat(1) = Pokemon(NpcNum).Add_Stat(1) + (Level * 3)
            .Stat(2) = Pokemon(NpcNum).Add_Stat(2) + (Level * 3)
            .Stat(3) = Pokemon(NpcNum).Add_Stat(3) + (Level * 3)
            .Stat(4) = Pokemon(NpcNum).Add_Stat(4) + (Level * 3)
            .Stat(5) = Pokemon(NpcNum).Add_Stat(5) + (Level * 3)
            .StunDuration = 0
            .SpawnWait = 0
            .StunTimer = 0
        End With
        Sumir = False
    Else
        With MapNpc(MapNum).Npc(MapNpcNum)
            .Num = NpcNum
            .x = x
            .Y = Y
            .Dir = DirNum
            .Sexo = Sex
            .Shiny = False
            .Level = Level
            .Desmaiado = False
            .Vital(1) = 0
            .Vital(2) = 0
            .Stat(1) = 0
            .Stat(2) = 0
            .Stat(3) = 0
            .Stat(4) = 0
            .Stat(5) = 0
        End With
        Sumir = True
    End If

    'Set Buffer = New clsBuffer
    '    Buffer.WriteLong SSpawnNpc
    '    Buffer.WriteLong MapNpcNum
    '    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Num
    '    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).X
    '    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Y
    '    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
    '    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Sexo
    '    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Shiny
    '    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Level
    '    SendDataToMap MapNum, Buffer.ToArray()
    'Set Buffer = Nothing
    
    'Enviar informações do mapa
    'For I = 1 To Player_HighIndex
    '    If IsPlaying(I) Then
    '        If GetPlayerMap(I) = MapNum Then
    '            SendMap I, MapNum
    '        End If
    '    End If
    'Next
    
    'Set Buffer = New clsBuffer
    '    Buffer.WriteLong SCheckForMap
    '    Buffer.WriteLong MapNum
    '    Buffer.WriteLong Map(MapNum).Revision
    '    SendDataToMap MapNum, Buffer.ToArray()
    'Set Buffer = Nothing
    
    'Criar o cache do mapa
    Call MapCache_Create(MapNum)
    
    'Mandar Desmaiado
    For I = 1 To Player_HighIndex
        If GetPlayerMap(I) = MapNum Then
            SendNpcDesmaiado MapNum, MapNpcNum, Sumir, I
        End If
    Next
    
    SendMapNpcsToMap MapNum
    SendMapNpcVitals MapNum, MapNpcNum
End Sub

Public Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, Optional ByVal SetX As Long, Optional ByVal SetY As Long)
    Dim Buffer As clsBuffer
    Dim NpcNum As Long, I As Long, x As Long, Y As Long
    Dim Spawned As Boolean, ShinyValue As Integer, ControlSex As Byte
    Dim Random50 As Long, Random100 As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    NpcNum = Map(MapNum).Npc(MapNpcNum)

    If NpcNum > 0 Then
    
        'Info Basica
        MapNpc(MapNum).Npc(MapNpcNum).Num = NpcNum
        MapNpc(MapNum).Npc(MapNpcNum).target = 0
        MapNpc(MapNum).Npc(MapNpcNum).targetType = 0 ' clear
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(NpcNum, Vitals.MP)
        MapNpc(MapNum).Npc(MapNpcNum).Dir = Int(Rnd * 4)
        
        'Controle Level do Pokémon
        If Map(MapNum).LevelPoke(1) > Map(MapNum).LevelPoke(2) Then
            MapNpc(MapNum).Npc(MapNpcNum).Level = Int(Random(Map(MapNum).LevelPoke(1), Map(MapNum).LevelPoke(2)))
        Else
            MapNpc(MapNum).Npc(MapNpcNum).Level = Int(Random(Map(MapNum).LevelPoke(2), Map(MapNum).LevelPoke(1)))
        End If
        
        'Vitals Pokémon
        If Npc(NpcNum).Pokemon > 0 Then
            MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = GetPokemonMaxVital(NpcNum, Vitals.HP, MapNpc(MapNum).Npc(MapNpcNum).Level)
        Else
            MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
        End If
        
        MapNpc(MapNum).Npc(MapNpcNum).Desmaiado = False
        SendNpcDesmaiado MapNum, MapNpcNum, False
        
        If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon > 0 Then
            'Ataque
            Random50 = Porcento(Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Strength), 50)
            Random100 = Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Strength)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Strength) = Random(Random50, Random100)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Strength) = MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Strength) + (MapNpc(MapNum).Npc(MapNpcNum).Level * 3)
            
            'Defesa
            Random50 = Porcento(Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Endurance), 50)
            Random100 = Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Endurance)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Endurance) = Random(Random50, Random100)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Endurance) = MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Endurance) + (MapNpc(MapNum).Npc(MapNpcNum).Level * 3)
            
            'Sp.Atq
            Random50 = Porcento(Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Intelligence), 50)
            Random100 = Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Intelligence)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Intelligence) = Random(Random50, Random100)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Intelligence) = MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Intelligence) + (MapNpc(MapNum).Npc(MapNpcNum).Level * 3)
            
            'Speed
            Random50 = Porcento(Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Agility), 50)
            Random100 = Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Agility)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Agility) = Random(Random50, Random100)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Agility) = MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Agility) + (MapNpc(MapNum).Npc(MapNpcNum).Level * 3)
            
            'Sp.Def
            Random50 = Porcento(Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Willpower), 50)
            Random100 = Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).Add_Stat(Stats.Willpower)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Willpower) = Random(Random50, Random100)
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Willpower) = MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Willpower) + (MapNpc(MapNum).Npc(MapNpcNum).Level * 3)
            
        Else
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Strength) = 0
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Endurance) = 0
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Intelligence) = 0
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Agility) = 0
            MapNpc(MapNum).Npc(MapNpcNum).Stat(Stats.Willpower) = 0
        End If
        
        'Controle Sexo Pokémon
        If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon > 0 Then
            ControlSex = Pokemon(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Pokemon).ControlSex
        
            If Int(Rnd * 100) <= ControlSex Then
                MapNpc(MapNum).Npc(MapNpcNum).Sexo = 0
            Else
                MapNpc(MapNum).Npc(MapNpcNum).Sexo = 1
            End If
        
        End If
        
        'Controle Pokémons Shinys
        ShinyValue = Int(Rnd * 100)
        
        If ShinyValue = 10 Then
            If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behaviour = 0 Or Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behaviour = 1 Then
                MapNpc(MapNum).Npc(MapNpcNum).Shiny = True
            End If
        Else
            MapNpc(MapNum).Npc(MapNpcNum).Shiny = False
        End If
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(MapNum).MaxX
            For Y = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(x, Y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(MapNum).Tile(x, Y).Data1 = MapNpcNum Then
                        MapNpc(MapNum).Npc(MapNpcNum).x = x
                        MapNpc(MapNum).Npc(MapNpcNum).Y = Y
                        MapNpc(MapNum).Npc(MapNpcNum).Dir = Map(MapNum).Tile(x, Y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next Y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For I = 1 To 100
                
                If SetX = 0 And SetY = 0 Then
                    x = Random(0, Map(MapNum).MaxX)
                    Y = Random(0, Map(MapNum).MaxY)
                Else
                    x = SetX
                    Y = SetY
                End If
    
                If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
                If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, x, Y) Then
                    MapNpc(MapNum).Npc(MapNpcNum).x = x
                    MapNpc(MapNum).Npc(MapNpcNum).Y = Y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(MapNum).MaxX
                For Y = 0 To Map(MapNum).MaxY

                    If NpcTileIsOpen(MapNum, x, Y) Then
                        MapNpc(MapNum).Npc(MapNpcNum).x = x
                        MapNpc(MapNum).Npc(MapNpcNum).Y = Y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Num
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Sexo
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Shiny
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Level
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
        
        SendMapNpcVitals MapNum, MapNpcNum
    End If

End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = Y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(MapNum).Npc(LoopI).x = x Then
                If MapNpc(MapNum).Npc(LoopI).Y = Y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_ITEM Then
                If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_GRASS Then
                NpcTileIsOpen = False
                End If
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim I As Long

    For I = 1 To MAX_MAP_NPCS
        Call SpawnNpc(I, MapNum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call SpawnMapNpcs(I)
    Next

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim I As Long
    Dim n As Long
    Dim x As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(MapNum).Npc(MapNpcNum).x
    Y = MapNpc(MapNum).Npc(MapNpcNum).Y
    CanNpcMove = True
    
    If MapNpc(MapNum).Npc(MapNpcNum).Desmaiado = True Then
        CanNpcMove = False
        Exit Function
    End If

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = Map(MapNum).Tile(x, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN And n <> TILE_TYPE_GRASS Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum).Npc(MapNpcNum).x) And (GetPlayerY(I) = MapNpc(MapNum).Npc(MapNpcNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If MapNpc(MapNum).Npc(I).Desmaiado = False Then
                    If (I <> MapNpcNum) And (MapNpc(MapNum).Npc(I).Num > 0) And (MapNpc(MapNum).Npc(I).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(I).Y = MapNpc(MapNum).Npc(MapNpcNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).Y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(x, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN And n <> TILE_TYPE_GRASS Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum).Npc(MapNpcNum).x) And (GetPlayerY(I) = MapNpc(MapNum).Npc(MapNpcNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If MapNpc(MapNum).Npc(I).Desmaiado = False Then
                    If (I <> MapNpcNum) And (MapNpc(MapNum).Npc(I).Num > 0) And (MapNpc(MapNum).Npc(I).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(I).Y = MapNpc(MapNum).Npc(MapNpcNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).Y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(MapNum).Tile(x - 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN And n <> TILE_TYPE_GRASS Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (GetPlayerY(I) = MapNpc(MapNum).Npc(MapNpcNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If MapNpc(MapNum).Npc(I).Desmaiado = False Then
                    If (I <> MapNpcNum) And (MapNpc(MapNum).Npc(I).Num > 0) And (MapNpc(MapNum).Npc(I).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (MapNpc(MapNum).Npc(I).Y = MapNpc(MapNum).Npc(MapNpcNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN And n <> TILE_TYPE_GRASS Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (GetPlayerY(I) = MapNpc(MapNum).Npc(MapNpcNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If MapNpc(MapNum).Npc(I).Desmaiado = False Then
                    If (I <> MapNpcNum) And (MapNpc(MapNum).Npc(I).Num > 0) And (MapNpc(MapNum).Npc(I).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (MapNpc(MapNum).Npc(I).Y = MapNpc(MapNum).Npc(MapNpcNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(MapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum).Npc(MapNpcNum).Y = MapNpc(MapNum).Npc(MapNpcNum).Y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(MapNum).Npc(MapNpcNum).Y = MapNpc(MapNum).Npc(MapNpcNum).Y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(MapNpcNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong MapNpcNum
    Buffer.WriteLong Dir
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim I As Long
    Dim n As Long
    n = 0

    For I = 1 To Player_HighIndex

        If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim I As Long

    For I = 1 To MAX_MAPS
        ClearTempTile I
    Next

End Sub

Sub ClearTempTile(ByVal MapNum As Long)
    Dim Y As Long
    Dim x As Long
    TempTile(MapNum).DoorTimer = 0
    ReDim TempTile(MapNum).DoorOpen(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For x = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            TempTile(MapNum).DoorOpen(x, Y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal MapNum As Long)
    Dim x As Long, Y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            If Map(MapNum).Tile(x, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).x = x
                ResourceCache(MapNum).ResourceData(Resource_Count).Y = Y
                ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).Tile(x, Y).Data1).health
            End If

        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim I As Long
    Dim OldNum As Long
    Dim OldValue As Long
    Dim OldPokemon As Long
    Dim OldPokeball As Long
    Dim OldLevel As Long
    Dim OldExp As Long
    Dim OldVital(1 To Vitals.Vital_Count - 1) As Long
    Dim OldMaxVital(1 To Vitals.Vital_Count - 1) As Long
    Dim OldStat(1 To Stats.Stat_Count - 1) As Long
    Dim OldSpell(1 To MAX_POKE_SPELL) As Long
    Dim OldNegatives(1 To 11) As Long
    Dim OldFelicidade As Long
    Dim OldSexo As Byte
    Dim OldShiny As Byte
    Dim OldBry(1 To MAX_BERRYS) As Long
    
    Dim NewNum As Long
    Dim NewValue As Long
    Dim NewPokemon As Long
    Dim NewPokeball As Long
    Dim NewLevel As Long
    Dim NewExp As Long
    Dim NewVital(1 To Vitals.Vital_Count - 1) As Long
    Dim NewMaxVital(1 To Vitals.Vital_Count - 1) As Long
    Dim NewStat(1 To Stats.Stat_Count - 1) As Long
    Dim NewSpell(1 To MAX_POKE_SPELL) As Long
    Dim NewNegatives(1 To 11) As Long
    Dim NewFelicidade As Long
    Dim NewSexo As Byte
    Dim NewShiny As Byte
    Dim NewBry(1 To MAX_BERRYS) As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(Index, oldSlot)
    OldValue = GetPlayerBankItemValue(Index, oldSlot)
    OldPokemon = GetPlayerBankItemPokemon(Index, oldSlot)
    OldPokeball = GetPlayerBankItemPokeball(Index, oldSlot)
    OldLevel = GetPlayerBankItemLevel(Index, oldSlot)
    OldExp = GetPlayerBankItemExp(Index, oldSlot)
    
    For I = 1 To Vitals.Vital_Count - 1
        OldVital(I) = GetPlayerBankItemVital(Index, oldSlot, I)
        OldMaxVital(I) = GetPlayerBankItemMaxVital(Index, oldSlot, I)
    Next
    
    For I = 1 To Stats.Stat_Count - 1
        OldStat(I) = GetPlayerBankItemStat(Index, oldSlot, I)
    Next
    
    For I = 1 To MAX_POKE_SPELL
        OldSpell(I) = GetPlayerBankItemSpell(Index, oldSlot, I)
    Next
    
    For I = 1 To MAX_NEGATIVES
        OldNegatives(I) = GetPlayerBankItemNgt(Index, oldSlot, I)
    Next
    
    For I = 1 To MAX_BERRYS
        OldBry(I) = GetPlayerBankItemBerry(Index, oldSlot, I)
    Next
    
    OldFelicidade = GetPlayerBankItemFelicidade(Index, oldSlot)
    OldSexo = GetPlayerBankItemSexo(Index, oldSlot)
    OldShiny = GetPlayerBankItemShiny(Index, oldSlot)
    
    NewNum = GetPlayerBankItemNum(Index, newSlot)
    NewValue = GetPlayerBankItemValue(Index, newSlot)
    NewPokemon = GetPlayerBankItemPokemon(Index, newSlot)
    NewPokeball = GetPlayerBankItemPokeball(Index, newSlot)
    NewLevel = GetPlayerBankItemLevel(Index, newSlot)
    NewExp = GetPlayerBankItemExp(Index, newSlot)
    
    For I = 1 To Vitals.Vital_Count - 1
        NewVital(I) = GetPlayerBankItemVital(Index, newSlot, I)
        NewMaxVital(I) = GetPlayerBankItemMaxVital(Index, newSlot, I)
    Next
    
    For I = 1 To Stats.Stat_Count - 1
        NewStat(I) = GetPlayerBankItemStat(Index, newSlot, I)
    Next
    
    For I = 1 To MAX_POKE_SPELL
        NewSpell(I) = GetPlayerBankItemSpell(Index, newSlot, I)
    Next
    
    For I = 1 To MAX_NEGATIVES
        NewNegatives(I) = GetPlayerBankItemNgt(Index, newSlot, I)
    Next
    
    For I = 1 To MAX_BERRYS
        NewBry(I) = GetPlayerBankItemBerry(Index, newSlot, I)
    Next
    
    NewFelicidade = GetPlayerBankItemFelicidade(Index, newSlot)
    NewSexo = GetPlayerBankItemSexo(Index, newSlot)
    NewShiny = GetPlayerBankItemShiny(Index, newSlot)

    SetPlayerBankItemNum Index, newSlot, OldNum
    SetPlayerBankItemValue Index, newSlot, OldValue
    SetPlayerBankItemPokemon Index, newSlot, OldPokemon
    SetPlayerBankItemPokeball Index, newSlot, OldPokeball
    SetPlayerBankItemLevel Index, newSlot, OldLevel
    SetPlayerBankItemExp Index, newSlot, OldExp
    
    For I = 1 To Vitals.Vital_Count - 1
        SetPlayerBankItemVital Index, newSlot, OldVital(I), I
        SetPlayerBankItemMaxVital Index, newSlot, OldVital(I), I
    Next
    
    For I = 1 To Stats.Stat_Count - 1
        SetPlayerBankItemStat Index, newSlot, OldStat(I), I
    Next
    
    For I = 1 To MAX_POKE_SPELL
        SetPlayerBankItemSpell Index, newSlot, OldSpell(I), I
    Next
    
    For I = 1 To MAX_NEGATIVES
        SetPlayerBankItemNgt Index, newSlot, OldNegatives(I), I
    Next
    
    For I = 1 To MAX_BERRYS
        SetPlayerBankItemBerry Index, newSlot, I, OldBry(I)
    Next
    
    SetPlayerBankItemFelicidade Index, newSlot, OldFelicidade
    SetPlayerBankItemSexo Index, newSlot, OldSexo
    SetPlayerBankItemShiny Index, newSlot, OldShiny
    
    '----------------------------------------------------------
    
    SetPlayerBankItemNum Index, oldSlot, NewNum
    SetPlayerBankItemValue Index, oldSlot, NewValue
    SetPlayerBankItemPokemon Index, oldSlot, NewPokemon
    SetPlayerBankItemPokeball Index, oldSlot, NewPokeball
    SetPlayerBankItemLevel Index, oldSlot, NewLevel
    SetPlayerBankItemExp Index, oldSlot, NewExp
    
    For I = 1 To Vitals.Vital_Count - 1
        SetPlayerBankItemVital Index, oldSlot, NewVital(I), I
        SetPlayerBankItemMaxVital Index, oldSlot, NewVital(I), I
    Next
    
    For I = 1 To Stats.Stat_Count - 1
        SetPlayerBankItemStat Index, oldSlot, NewStat(I), I
    Next
    
    For I = 1 To MAX_POKE_SPELL
        SetPlayerBankItemSpell Index, oldSlot, NewSpell(I), I
    Next
    
    For I = 1 To MAX_NEGATIVES
        SetPlayerBankItemNgt Index, oldSlot, NewNegatives(I), I
    Next
    
    For I = 1 To MAX_BERRYS
        SetPlayerBankItemBerry Index, oldSlot, I, NewBry(I)
    Next
    
    SetPlayerBankItemFelicidade Index, oldSlot, NewFelicidade
    SetPlayerBankItemSexo Index, oldSlot, NewSexo
    SetPlayerBankItemShiny Index, oldSlot, NewShiny

    SendBank Index
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim I As Long
    Dim OldNum As Long
    Dim OldValue As Long
    Dim OldPokemon As Long
    Dim OldPokeball As Long
    Dim OldLevel As Long
    Dim OldExp As Long
    Dim OldVital(1 To Vitals.Vital_Count - 1) As Long
    Dim OldMaxVital(1 To Vitals.Vital_Count - 1) As Long
    Dim OldStat(1 To Stats.Stat_Count - 1) As Long
    Dim OldSpell(1 To MAX_POKE_SPELL) As Long
    Dim OldNegatives(1 To MAX_NEGATIVES) As Long
    Dim OldFelicidade As Long
    Dim OldSexo As Byte
    Dim OldShiny As Byte
    Dim OldBerry(1 To MAX_BERRYS) As Long
    
    Dim NewNum As Long
    Dim NewValue As Long
    Dim NewPokemon As Long
    Dim NewPokeball As Long
    Dim NewLevel As Long
    Dim NewExp As Long
    Dim NewVital(1 To Vitals.Vital_Count - 1) As Long
    Dim NewMaxVital(1 To Vitals.Vital_Count - 1) As Long
    Dim NewStat(1 To Stats.Stat_Count - 1) As Long
    Dim NewSpell(1 To MAX_POKE_SPELL) As Long
    Dim NewNegatives(1 To MAX_NEGATIVES) As Long
    Dim NewFelicidade As Long
    Dim NewSexo As Byte
    Dim NewShiny As Byte
    Dim NewBerry(1 To MAX_BERRYS) As Long
    
    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    '##########################Old Numbers##################################

    OldNum = GetPlayerInvItemNum(Index, oldSlot)
    OldValue = GetPlayerInvItemValue(Index, oldSlot)
    OldPokemon = GetPlayerInvItemPokeInfoPokemon(Index, oldSlot)
    OldPokeball = GetPlayerInvItemPokeInfoPokeball(Index, oldSlot)
    OldLevel = GetPlayerInvItemPokeInfoLevel(Index, oldSlot)
    OldExp = GetPlayerInvItemPokeInfoExp(Index, oldSlot)
    
    For I = 1 To Vitals.Vital_Count - 1
        OldVital(I) = GetPlayerInvItemPokeInfoVital(Index, oldSlot, I)
        OldMaxVital(I) = GetPlayerInvItemPokeInfoMaxVital(Index, oldSlot, I)
    Next
    
    For I = 1 To Stats.Stat_Count - 1
        OldStat(I) = GetPlayerInvItemPokeInfoStat(Index, oldSlot, I)
    Next
    
    For I = 1 To MAX_POKE_SPELL
        OldSpell(I) = GetPlayerInvItemPokeInfoSpell(Index, oldSlot, I)
    Next
    
    For I = 1 To MAX_NEGATIVES
        OldNegatives(I) = GetPlayerInvItemNgt(Index, oldSlot, I)
    Next
    
    For I = 1 To MAX_BERRYS
        OldBerry(I) = GetPlayerInvItemBerry(Index, oldSlot, I)
    Next
    
    OldFelicidade = GetPlayerInvItemFelicidade(Index, oldSlot)
    OldSexo = GetPlayerInvItemSexo(Index, oldSlot)
    OldShiny = GetPlayerInvItemShiny(Index, oldSlot)
    
    '##########################New Numbers##################################
    
    NewNum = GetPlayerInvItemNum(Index, newSlot)
    NewValue = GetPlayerInvItemValue(Index, newSlot)
    NewPokemon = GetPlayerInvItemPokeInfoPokemon(Index, newSlot)
    NewPokeball = GetPlayerInvItemPokeInfoPokeball(Index, newSlot)
    NewLevel = GetPlayerInvItemPokeInfoLevel(Index, newSlot)
    NewExp = GetPlayerInvItemPokeInfoExp(Index, newSlot)
    
    For I = 1 To Vitals.Vital_Count - 1
        NewVital(I) = GetPlayerInvItemPokeInfoVital(Index, newSlot, I)
        NewMaxVital(I) = GetPlayerInvItemPokeInfoMaxVital(Index, newSlot, I)
    Next
    
    For I = 1 To Stats.Stat_Count - 1
        NewStat(I) = GetPlayerInvItemPokeInfoStat(Index, newSlot, I)
    Next
    
    For I = 1 To MAX_POKE_SPELL
        NewSpell(I) = GetPlayerInvItemPokeInfoSpell(Index, newSlot, I)
    Next
    
    For I = 1 To MAX_NEGATIVES
        NewNegatives(I) = GetPlayerInvItemNgt(Index, newSlot, I)
    Next
    
    For I = 1 To MAX_BERRYS
        NewBerry(I) = GetPlayerInvItemBerry(Index, newSlot, I)
    Next
    
    NewFelicidade = GetPlayerInvItemFelicidade(Index, newSlot)
    NewSexo = GetPlayerInvItemSexo(Index, newSlot)
    NewShiny = GetPlayerInvItemShiny(Index, newSlot)

    '#########################Old SetValues#################################
    
    SetPlayerInvItemNum Index, newSlot, OldNum
    SetPlayerInvItemValue Index, newSlot, OldValue
    SetPlayerInvItemPokeInfoPokemon Index, newSlot, OldPokemon
    SetPlayerInvItemPokeInfoPokeball Index, newSlot, OldPokeball
    SetPlayerInvItemPokeInfoLevel Index, newSlot, OldLevel
    SetPlayerInvItemPokeInfoExp Index, newSlot, OldExp
    
    For I = 1 To Vitals.Vital_Count - 1
        SetPlayerInvItemPokeInfoVital Index, newSlot, OldVital(I), I
        SetPlayerInvItemPokeInfoMaxVital Index, newSlot, OldMaxVital(I), I
    Next
    
    For I = 1 To Stats.Stat_Count - 1
        SetPlayerInvItemPokeInfoStat Index, newSlot, I, OldStat(I)
    Next
    
    For I = 1 To MAX_POKE_SPELL
        SetPlayerInvItemPokeInfoSpell Index, newSlot, OldSpell(I), I
    Next
    
    For I = 1 To MAX_NEGATIVES
        SetPlayerInvItemNgt Index, newSlot, I, OldNegatives(I)
    Next
    
    For I = 1 To MAX_BERRYS
        SetPlayerInvItemBerry Index, newSlot, I, OldBerry(I)
    Next
    
    SetPlayerInvItemFelicidade Index, newSlot, OldFelicidade
    SetPlayerInvItemSexo Index, newSlot, OldSexo
    SetPlayerInvItemShiny Index, newSlot, OldShiny
    
    '#########################New SetValues#################################
    
    SetPlayerInvItemNum Index, oldSlot, NewNum
    SetPlayerInvItemValue Index, oldSlot, NewValue
    SetPlayerInvItemPokeInfoPokemon Index, oldSlot, NewPokemon
    SetPlayerInvItemPokeInfoPokeball Index, oldSlot, NewPokeball
    SetPlayerInvItemPokeInfoLevel Index, oldSlot, NewLevel
    SetPlayerInvItemPokeInfoExp Index, oldSlot, NewExp
    
    For I = 1 To Vitals.Vital_Count - 1
        SetPlayerInvItemPokeInfoVital Index, oldSlot, NewVital(I), I
        SetPlayerInvItemPokeInfoMaxVital Index, oldSlot, NewMaxVital(I), I
    Next
    
    For I = 1 To Stats.Stat_Count - 1
        SetPlayerInvItemPokeInfoStat Index, oldSlot, I, NewStat(I)
    Next
    
    For I = 1 To MAX_POKE_SPELL
        SetPlayerInvItemPokeInfoSpell Index, oldSlot, NewSpell(I), I
    Next
    
    For I = 1 To MAX_NEGATIVES
        SetPlayerInvItemNgt Index, oldSlot, I, NewNegatives(I)
    Next
    
    For I = 1 To MAX_BERRYS
        SetPlayerInvItemBerry Index, oldSlot, I, NewBerry(I)
    Next
    
    SetPlayerInvItemFelicidade Index, oldSlot, NewFelicidade
    SetPlayerInvItemSexo Index, oldSlot, NewSexo
    SetPlayerInvItemShiny Index, oldSlot, NewShiny
    
    SendInventory Index
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(Index, oldSlot)
    NewNum = GetPlayerSpell(Index, newSlot)
    SetPlayerSpell Index, oldSlot, NewNum
    SetPlayerSpell Index, newSlot, OldNum
    SendPlayerSpells Index
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)
Dim I As Long

Dim ItemNum As Long, Pokemon As Long, Pokeball As Long, Level As Long
Dim EXP As Long, VitalHP As Long, VitalMP As Long, MVitalHp As Long, MVitalMP As Long
Dim Stat(1 To Stats.Stat_Count - 1) As Long, Spell(1 To MAX_POKE_SPELL) As Long, Ngt(1 To MAX_NEGATIVES) As Long
Dim Felicidade As Long, Sexo As Byte, Shiny As Byte, Bry(1 To MAX_BERRYS)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If TempPlayer(Index).EvolTimer > 0 Or Player(Index).EvolTimerStone > 0 Then Exit Sub
    If Player(Index).EvolPermition > 0 Then Exit Sub
    
    If FindOpenInvSlot(Index, GetPlayerEquipment(Index, EqSlot)) > 0 Then
        
    If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_ARENA Then
        If TempPlayer(Index).LutandoA > 0 Then
            If GetPlayerEquipmentPokeInfoVital(Index, weapon, 1) > 0 Then
                PlayerMsg Index, "Você não pode chamar de volta seu pokémon em batalha!", BrightRed
                Exit Sub
            End If
        End If
    End If

    If GetPlayerEquipmentPokeInfoPokemon(Index, EqSlot) > 0 Then
        Select Case GetPlayerEquipmentPokeInfoPokeball(Index, weapon)
        Case 2
            SendAnimation GetPlayerMap(Index), 33, GetPlayerX(Index), GetPlayerY(Index)
        Case 3
            SendAnimation GetPlayerMap(Index), 38, GetPlayerX(Index), GetPlayerY(Index)
        Case 4
            SendAnimation GetPlayerMap(Index), 36, GetPlayerX(Index), GetPlayerY(Index)
        Case Else
            SendAnimation GetPlayerMap(Index), 8, GetPlayerX(Index), GetPlayerY(Index)
        End Select
    End If
    
    If Item(GetPlayerEquipment(Index, weapon)).Type = ITEM_TYPE_WEAPON Then
    
    'Limpar Spells
    For I = 1 To MAX_PLAYER_SPELLS
        Call SetPlayerSpell(Index, I, 0)
            SendSpells Index
        Next
    End If
    
    ItemNum = GetPlayerEquipment(Index, EqSlot)
    Pokemon = GetPlayerEquipmentPokeInfoPokemon(Index, EqSlot)
    Pokeball = GetPlayerEquipmentPokeInfoPokeball(Index, EqSlot)
    Level = GetPlayerEquipmentPokeInfoLevel(Index, EqSlot)
    EXP = GetPlayerEquipmentPokeInfoExp(Index, EqSlot)
    VitalHP = GetPlayerEquipmentPokeInfoVital(Index, EqSlot, 1)
    VitalMP = GetPlayerEquipmentPokeInfoVital(Index, EqSlot, 2)
    MVitalHp = GetPlayerEquipmentPokeInfoMaxVital(Index, EqSlot, 1)
    MVitalMP = GetPlayerEquipmentPokeInfoMaxVital(Index, EqSlot, 2)
    Felicidade = GetPlayerEquipmentFelicidade(Index, EqSlot)
    Sexo = GetPlayerEquipmentSexo(Index, EqSlot)
    Shiny = GetPlayerEquipmentShiny(Index, EqSlot)
    
    For I = 1 To Stats.Stat_Count - 1
        Stat(I) = GetPlayerEquipmentPokeInfoStat(Index, EqSlot, I)
    Next
    
    For I = 1 To MAX_POKE_SPELL
        Spell(I) = GetPlayerEquipmentPokeInfoSpell(Index, EqSlot, I)
    Next
    
    For I = 1 To MAX_NEGATIVES
        Ngt(I) = GetPlayerEquipmentNgt(Index, EqSlot, I)
        
        If TempPlayer(Index).NgtTick(I) > 0 Then
            TempPlayer(Index).NgtTick(I) = 0
        End If
    Next
    
    For I = 1 To MAX_BERRYS
        Bry(I) = GetPlayerEquipmentBerry(Index, EqSlot, I)
    Next
    
        'GiveinvItem
        GiveInvItem Index, ItemNum, 1, False, Pokemon, Pokeball, Level, EXP, _
        VitalHP, VitalMP, MVitalHp, MVitalMP, Stat(1), Stat(4), Stat(2), Stat(3), _
        Stat(5), Spell(1), Spell(2), Spell(3), Spell(4), Ngt(1), Ngt(2), Ngt(3), Ngt(4), _
        Ngt(5), Ngt(6), Ngt(7), Ngt(8), Ngt(9), Ngt(10), Ngt(11), Felicidade, Sexo, Shiny, Bry(1), Bry(2), Bry(3), Bry(4), Bry(5)
        
        'Sair do Flying Mode
        If Player(Index).Flying = 1 Then
            SetPlayerFlying Index, 0
        End If

        'Retirar Pokémon
        If GetPlayerEquipmentPokeInfoPokemon(Index, EqSlot) > 0 Then
            Player(Index).x = Player(Index).TPX
            Player(Index).Y = Player(Index).TPY
            Player(Index).Dir = Player(Index).TPDir
            Player(Index).TPDir = 0
            Player(Index).TPX = 0
            Player(Index).TPY = 0
            Player(Index).TPSprite = 0
            SendPlayerXY Index
            Player(Index).Sprite = Player(Index).MySprite
        Else
            If Item(GetPlayerEquipment(Index, EqSlot)).Pokemon > 0 Then
                Player(Index).Sprite = Player(Index).MySprite
            End If
        End If
        
        ' send the sound
        SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, EqSlot)
        
        ' remove equipment
        SetPlayerExp Index, 0
        SendEXP Index
        SetPlayerEquipment Index, 0, EqSlot
        SetPlayerEquipmentPokeInfoPokemon Index, 0, EqSlot
        SetPlayerEquipmentPokeInfoLevel Index, 0, EqSlot
        SetPlayerEquipmentPokeInfoExp Index, 0, EqSlot
        
        For I = 1 To MAX_NEGATIVES
            SetPlayerEquipmentNgt Index, I, weapon, 0
        Next
        
        SendWornEquipment Index
        SendMapEquipment Index
        'SendStats Index
        SendPlayerData Index
        
        ' send vitals
        Call GetPlayerMaxVital(Index, HP)
        Call GetPlayerMaxVital(Index, MP)
        Call SetPlayerVital(Index, HP, Player(Index).VitalTemp)
        Player(Index).VitalTemp = 0
        Call SetPlayerVital(Index, MP, 1)
        Call SendVital(Index, Vitals.HP)
        Call SendVital(Index, Vitals.MP)
     
    'Else
    '    PlayerMsg Index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
Dim partynum As Long, I As Long

    partynum = TempPlayer(Index).inParty
    If partynum > 0 Then
        ' find out how many members we have
        Party_CountMembers partynum
        ' make sure there's more than 2 people
        If Party(partynum).MemberCount > 2 Then
            ' check if leader
            If Party(partynum).Leader = Index Then
                ' set next person down as leader
                For I = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(I) > 0 And Party(partynum).Member(I) <> Index Then
                        Party(partynum).Leader = Party(partynum).Member(I)
                        PartyMsg partynum, GetPlayerName(Party(partynum).Leader) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partynum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For I = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(I) = Index Then
                        Party(partynum).Member(I) = 0
                        Exit For
                    End If
                Next
                
                TempPlayer(Index).inParty = 0
                TempPlayer(Index).partyInvite = 0
                
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo Index
            Else
                ' not the leader, just leave
                PartyMsg partynum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For I = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(I) = Index Then
                        Party(partynum).Member(I) = 0
                        Exit For
                    End If
                Next
                
                TempPlayer(Index).inParty = 0
                TempPlayer(Index).partyInvite = 0
                
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo Index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partynum
            ' only 2 people, disband
            PartyMsg partynum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For I = 1 To MAX_PARTY_MEMBERS
                Index = Party(partynum).Member(I)
                ' player exist?
                If Index > 0 Then
                    ' remove them
                    TempPlayer(Index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo Index
                End If
            Next
            ' clear out the party itself
            ClearParty partynum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partynum As Long, I As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg Index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg Index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(Index).inParty > 0 Then
        partynum = TempPlayer(Index).inParty
        ' make sure we're the leader
        If Party(partynum).Leader = Index Then
            ' got a blank slot?
            For I = 1 To MAX_PARTY_MEMBERS
                If Party(partynum).Member(I) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, Index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = Index
                    ' let them know
                    PlayerMsg Index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg Index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg Index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, Index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = Index
        ' let them know
        PlayerMsg Index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partynum As Long, I As Long, x As Long

    ' check if already in a party
    If TempPlayer(Index).inParty > 0 Then
        ' get the partynumber
        partynum = TempPlayer(Index).inParty
        ' got a blank slot?
        For I = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(I) = 0 Then
                'add to the party
                Party(partynum).Member(I) = targetPlayer
                ' recount party
                Party_CountMembers partynum
                ' send update to all - including new player
                SendPartyUpdate partynum
                
                For x = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(x) > 0 Then
                        SendPartyVitals partynum, Party(partynum).Member(x)
                    End If
                Next
                ' let everyone know they've joined
                PartyMsg partynum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partynum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg Index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        TempPlayer(targetPlayer).partyInvite = 0
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For I = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(I).Leader > 0 Then
                partynum = I
                Exit For
            End If
        Next
        ' create the party
        Party(partynum).MemberCount = 2
        Party(partynum).Leader = Index
        Party(partynum).Member(1) = Index
        Party(partynum).Member(2) = targetPlayer
        SendPartyUpdate partynum
        SendPartyVitals partynum, Index
        SendPartyVitals partynum, targetPlayer
        ' let them know it's created
        PartyMsg partynum, "Party created.", BrightGreen
        PartyMsg partynum, GetPlayerName(Index) & " has joined the party.", Pink
        PartyMsg partynum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(Index).inParty = partynum
        TempPlayer(targetPlayer).inParty = partynum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal targetPlayer As Long)
    PlayerMsg Index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partynum As Long)
Dim I As Long, highIndex As Long, x As Long
    ' find the high index
    For I = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partynum).Member(I) > 0 Then
            highIndex = I
            Exit For
        End If
    Next
    ' count the members
    For I = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partynum).Member(I) = 0 Then
            ' is it lower than the high index?
            If I < highIndex Then
                ' move everyone down a slot
                For x = I To MAX_PARTY_MEMBERS - 1
                    Party(partynum).Member(x) = Party(partynum).Member(x + 1)
                    Party(partynum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partynum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If I = MAX_PARTY_MEMBERS Then
            If highIndex = I Then
                Party(partynum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partynum
End Sub

Public Sub Party_ShareExp(ByVal partynum As Long, ByVal EXP As Long, ByVal Index As Long)
Dim expShare As Long, leftOver As Long, I As Long, tmpIndex As Long

If Party(partynum).MemberCount <= 0 Then Exit Sub

    ' check if it's worth sharing
    If Not EXP >= Party(partynum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP Index, EXP
        Exit Sub
    End If
    
    ' find out the equal share
    expShare = EXP \ Party(partynum).MemberCount
    leftOver = EXP Mod Party(partynum).MemberCount
    
    ' loop through and give everyone exp
    For I = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partynum).Member(I)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partynum).Member(RAND(1, Party(partynum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal EXP As Long)
    ' give the exp
    Call SetPlayerExp(Index, GetPlayerExp(Index) + EXP)
    SendEXP Index
    SendWornEquipment Index
    SendActionMsg GetPlayerMap(Index), "+" & EXP & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp Index
End Sub

Public Sub AcabarTrans(ByVal Index As Long)
If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
TempPlayer(Index).Pokemon = 0
SendPlayerData Index
SavePlayer Index
'SendStats Index
End Sub

Public Function GetSpellBaseStat(ByVal Index As Long, ByVal SpellNum As Long) As Long

If SpellNum > 0 Then
Select Case Spell(SpellNum).BaseStat
Case 1
GetSpellBaseStat = GetPlayerStat(Index, Stats.Strength) * 1.5 + GetPlayerLevel(Index) / 6.2 + Spell(SpellNum).Vital
Case 2
GetSpellBaseStat = GetPlayerStat(Index, Stats.Intelligence) * 1.5 + GetPlayerLevel(Index) / 6.2 + Spell(SpellNum).Vital
Case 3
GetSpellBaseStat = GetPlayerStat(Index, Stats.Agility) * 1.5 + GetPlayerLevel(Index) / 6.2 + Spell(SpellNum).Vital
Case 4
GetSpellBaseStat = GetPlayerStat(Index, Stats.Endurance) * 1.5 + GetPlayerLevel(Index) / 6.2 + Spell(SpellNum).Vital
Case 5
GetSpellBaseStat = GetPlayerStat(Index, Stats.Willpower) * 1.5 + GetPlayerLevel(Index) / 6.2 + Spell(SpellNum).Vital

End Select
End If

End Function

Public Sub PescarPokemon(ByVal Index As Long)
Dim I As Long, x As Long
Dim ExtraChance As Integer, RandomPokeId As Integer
Dim LevelMin As Integer, PokeList(1 To 20) As Integer

If Map(GetPlayerMap(Index)).LevelPoke(1) <= Map(GetPlayerMap(Index)).LevelPoke(2) Then
    LevelMin = Map(GetPlayerMap(Index)).LevelPoke(1)
Else
    LevelMin = Map(GetPlayerMap(Index)).LevelPoke(2)
End If

'Select Case PokeList()

'End Select


ExtraChance = Player(Index).UltRodLevel * 0.08
I = Random(1, 100)

If I <= 25 + Int(ExtraChance) Then

    Select Case LevelMin
    Case 1 To 10
        SpawnPokePescado Index, 8
    Case 11 To 20
        SpawnPokePescado Index, 8
    Case 21 To 30
        SpawnPokePescado Index, 8
    Case 31 To 40
        SpawnPokePescado Index, 8
    Case 41 To 50
        SpawnPokePescado Index, 8
    Case 51 To 60
        SpawnPokePescado Index, 8
    Case 61 To 70
        SpawnPokePescado Index, 8
    Case 71 To 80
        SpawnPokePescado Index, 8
    Case 81 To 90
        SpawnPokePescado Index, 8
    Case 91 To 99
        SpawnPokePescado Index, 8
    End Select
    
    'Mandar Animação
    SendAnimation GetPlayerMap(Index), 13, GetPlayerX(Index), GetPlayerY(Index)
Else
    PlayerMsg Index, "Você não pescou nada, mais sorte na próxima!", White
End If

End Sub

Public Sub SpawnPokePescado(ByVal Index As Long, ByVal NpcNum As Long)
Dim I As Long, OpenSlot As Long, PlayerMap As Long
Dim Buffer As clsBuffer
Dim Random50 As Long, Random100 As Long, ShinyValue As Integer, ControlSex As Byte

PlayerMap = GetPlayerMap(Index)

        'Verificar Slot Vazio...
        For I = 1 To MAX_MAP_NPCS
            If Map(PlayerMap).Npc(I) = 0 Then
                OpenSlot = I
                Exit For
            End If
        Next
        
        'Sem Slot Aberto FODEU!
        If OpenSlot = 0 Then
            PlayerMsg Index, "Excesso de Pokémons no mapa! O pokémon voltou para Água!", BrightRed
            Exit Sub
        End If
        
        'Sem numero do pokémon não tem como ligar para ele vir né...
        If NpcNum = 0 Then Exit Sub
        
        'Criar o Npc no mapa
        Map(PlayerMap).Npc(OpenSlot) = NpcNum
        MapNpc(PlayerMap).Npc(OpenSlot).Num = NpcNum
        
        'Level do Pokémon!
        If Map(PlayerMap).LevelPoke(1) > Map(PlayerMap).LevelPoke(2) Then
            MapNpc(PlayerMap).Npc(OpenSlot).Level = Int(Random(Map(PlayerMap).LevelPoke(1), Map(PlayerMap).LevelPoke(2)))
        Else
            MapNpc(PlayerMap).Npc(OpenSlot).Level = Int(Random(Map(PlayerMap).LevelPoke(2), Map(PlayerMap).LevelPoke(1)))
        End If
        
        'Vitals Pokémon
        If Npc(NpcNum).Pokemon > 0 Then
            MapNpc(PlayerMap).Npc(OpenSlot).Vital(Vitals.HP) = GetPokemonMaxVital(NpcNum, Vitals.HP, MapNpc(PlayerMap).Npc(OpenSlot).Level)
        Else
            MapNpc(PlayerMap).Npc(OpenSlot).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
        End If
        
        If Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon > 0 Then
            'Ataque
            Random50 = Porcento(Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Strength), 50)
            Random100 = Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Strength)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Strength) = Random(Random50, Random100)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Strength) = MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Strength) + (MapNpc(PlayerMap).Npc(OpenSlot).Level * 3)
            
            'Defesa
            Random50 = Porcento(Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Endurance), 50)
            Random100 = Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Endurance)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Endurance) = Random(Random50, Random100)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Endurance) = MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Endurance) + (MapNpc(PlayerMap).Npc(OpenSlot).Level * 3)
            
            'Sp.Atq
            Random50 = Porcento(Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Intelligence), 50)
            Random100 = Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Intelligence)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Intelligence) = Random(Random50, Random100)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Intelligence) = MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Intelligence) + (MapNpc(PlayerMap).Npc(OpenSlot).Level * 3)
            
            'Speed
            Random50 = Porcento(Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Agility), 50)
            Random100 = Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Agility)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Agility) = Random(Random50, Random100)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Agility) = MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Agility) + (MapNpc(PlayerMap).Npc(OpenSlot).Level * 3)
            
            'Sp.Def
            Random50 = Porcento(Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Willpower), 50)
            Random100 = Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).Add_Stat(Stats.Willpower)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Willpower) = Random(Random50, Random100)
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Willpower) = MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Willpower) + (MapNpc(PlayerMap).Npc(OpenSlot).Level * 3)
            
        Else
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Strength) = 0
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Endurance) = 0
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Intelligence) = 0
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Agility) = 0
            MapNpc(PlayerMap).Npc(OpenSlot).Stat(Stats.Willpower) = 0
        End If
        
        'Controle Sexo Pokémon
        If Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon > 0 Then
            ControlSex = Pokemon(Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Pokemon).ControlSex
        
            If Int(Rnd * 100) <= ControlSex Then
                MapNpc(PlayerMap).Npc(OpenSlot).Sexo = 0
            Else
                MapNpc(PlayerMap).Npc(OpenSlot).Sexo = 1
            End If
        
        End If
        
        'Controle Pokémons Shinys
        ShinyValue = Int(Rnd * 100)
        
        If ShinyValue = 10 Then
            If Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Behaviour = 0 Or Npc(MapNpc(PlayerMap).Npc(OpenSlot).Num).Behaviour = 1 Then
                MapNpc(PlayerMap).Npc(OpenSlot).Shiny = True
            End If
        Else
            MapNpc(PlayerMap).Npc(OpenSlot).Shiny = False
        End If
        
        'Criar o cache do mapa
        Call MapCache_Create(PlayerMap)
        
        'Enviar informações do mapa
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If GetPlayerMap(I) = GetPlayerMap(Index) Then
                    SendMap I, PlayerMap
                End If
            End If
        Next
        
        'Setar Posição do Npc Pescado!
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                MapNpc(PlayerMap).Npc(OpenSlot).x = GetPlayerX(Index)
                MapNpc(PlayerMap).Npc(OpenSlot).Y = GetPlayerY(Index) + 1
                MapNpc(PlayerMap).Npc(OpenSlot).Dir = DIR_UP
            Case DIR_DOWN
                MapNpc(PlayerMap).Npc(OpenSlot).x = GetPlayerX(Index)
                MapNpc(PlayerMap).Npc(OpenSlot).Y = GetPlayerY(Index) - 1
                MapNpc(PlayerMap).Npc(OpenSlot).Dir = DIR_DOWN
            Case DIR_LEFT
                MapNpc(PlayerMap).Npc(OpenSlot).x = GetPlayerX(Index) + 1
                MapNpc(PlayerMap).Npc(OpenSlot).Y = GetPlayerY(Index)
                MapNpc(PlayerMap).Npc(OpenSlot).Dir = DIR_LEFT
            Case DIR_RIGHT
                MapNpc(PlayerMap).Npc(OpenSlot).x = GetPlayerX(Index) - 1
                MapNpc(PlayerMap).Npc(OpenSlot).Y = GetPlayerY(Index)
                MapNpc(PlayerMap).Npc(OpenSlot).Dir = DIR_RIGHT
            End Select
            
            MapNpc(PlayerMap).Npc(OpenSlot).Pescado = True
            
            Set Buffer = New clsBuffer
            Buffer.WriteLong SCheckForMap
            Buffer.WriteLong PlayerMap
            Buffer.WriteLong Map(PlayerMap).Revision
            SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
            Set Buffer = Nothing
            
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(PlayerMap).Npc(I).Num > 0 Then
                    If MapNpc(PlayerMap).Npc(I).Desmaiado = True And Map(PlayerMap).Npc(I) > 0 Then
                        SendNpcDesmaiado PlayerMap, I, False
                        SendMapNpcVitals PlayerMap, I
                    End If
                End If
            Next
        
End Sub

Public Sub RemoverNpcPescado(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim I As Long, Buffer As clsBuffer, PlayerMap As Long

PlayerMap = GetPlayerMap(MapNum)

If MapNpcNum = 0 Then Exit Sub
            'Criar o Npc no mapa
            Map(MapNum).Npc(MapNpcNum) = 0
            MapNpc(MapNum).Npc(MapNpcNum).Num = 0
            MapNpc(MapNum).Npc(MapNpcNum).Desmaiado = False
            MapNpc(MapNum).Npc(MapNpcNum).SpawnWait = 0
            
            'Criar o cache do mapa
            Call MapCache_Create(MapNum)
            
            'Enviar informações do mapa
            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    If GetPlayerMap(I) = MapNum Then
                        SendMap I, MapNum
                    End If
                End If
            Next
            
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(I).Num > 0 Then
                    If MapNpc(MapNum).Npc(I).Desmaiado = True And Map(MapNum).Npc(I) > 0 Then
                        SendNpcDesmaiado MapNum, I, False
                        SendMapNpcVitals MapNum, I
                    End If
                End If
            Next
            '
            SendMapNpcsToMap PlayerMap
End Sub

Public Function PokeTileIsOpen(ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long) As Boolean
    Dim LoopI As Long
    PokeTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = Y Then
                        PokeTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(MapNum).Npc(LoopI).x = x Then
                If MapNpc(MapNum).Npc(LoopI).Y = Y Then
                    PokeTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If x < 0 Then x = 0
    If Y < 0 Then Y = 0
    If x > Map(MapNum).MaxX Then x = 0
    If Y > Map(MapNum).MaxY Then Y = 0

    If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_ITEM Then
                If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_GRASS Then
                    If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_FISHING Then
                        If Map(MapNum).Tile(x, Y).Type <> TILE_TYPE_NPCAVOID Then
                            PokeTileIsOpen = False
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function FindRankLevel(Level As Long) As Byte
Dim I As Byte
    
    For I = 1 To MAX_RANKS
        If RankLevel(I).Level < Level Then
            FindRankLevel = I
            Exit Function
        End If
    Next
    
End Function

Public Sub UpdateRankLevel(Index As Long)
Dim Name As String, Level As Long, PokeNum As Long
Dim Name2 As String, Level2 As Long, PokeNum2 As Long
Dim Position As Long

Name = GetPlayerName(Index)
Level = GetPlayerLevel(Index)
PokeNum = GetPlayerEquipmentPokeInfoPokemon(Index, weapon)

Position = FindRankLevel(Level)

If GetPlayerAccess(Index) > 0 Then Exit Sub
If Position <= 0 Then Exit Sub

If Position + 1 <= MAX_RANKS And RankLevel(Position).Name <> GetPlayerName(Index) Then
    RankLevel(Position + 1).Name = RankLevel(Position).Name
    RankLevel(Position + 1).Level = RankLevel(Position).Level
    RankLevel(Position + 1).PokeNum = RankLevel(Position).PokeNum 'GetPlayerEquipmentPokeInfoPokemon(index, weapon)
End If

RankLevel(Position).Name = Name
RankLevel(Position).Level = Level
RankLevel(Position).PokeNum = PokeNum

SendRankLevel
SaveRankLevel
End Sub

Public Sub CheckVipDays(ByVal Index As Long, MsgDiasRest As Boolean)
Dim I As Long

    'Evitar OverFlow
    If Player(Index).MyVip > 0 And Player(Index).VipStart = "00/00/0000" Then Exit Sub

    If Player(Index).MyVip > 0 Then
            If DateDiff("d", Player(Index).VipStart, Date) < Player(Index).VipDays(Player(Index).MyVip) Then
                If MsgDiasRest = True Then
                    PlayerMsg Index, "Você possui " & Player(Index).VipDays(Player(Index).MyVip) - DateDiff("d", Player(Index).VipStart, Date) & " dias de #Vip " & Player(Index).MyVip & ", Tenha um bom jogo!", BrightCyan
                End If
            ElseIf DateDiff("d", Player(Index).VipStart, Date) >= Player(Index).VipDays(Player(Index).MyVip) Then
                If Player(Index).MyVip = 1 Then
                    Player(Index).MyVip = 0
                    Player(Index).VipDays(1) = 0
                    Player(Index).VipStart = "00/00/0000"
                    PlayerMsg Index, "Seu Vip 1 Expirou, Tenha um bom jogo!", BrightRed
                    SendVipPointsInfo Index
                ElseIf Player(Index).MyVip >= 2 Then
                    'Zerar Dias Vip Atual
                    Player(Index).VipDays(Player(Index).MyVip) = 0
                    
                    'Verificar se há Outros Vips com Dias Restantes!
                    For I = Player(Index).MyVip To 1 Step -1
                        If Player(Index).VipDays(I) > 0 Then
                            PlayerMsg Index, "Seu dias de #Vip " & Player(Index).MyVip & " acabou, e foi alterado para #Vip " & I & " por " & Player(Index).VipDays(I) & " Dias.", BrightCyan
                            Player(Index).MyVip = I
                            Player(Index).VipStart = Trim$(DateValue(Date))
                            SendVipPointsInfo Index
                            Exit Sub
                        End If
                    Next
                    
                    'Limpar Vip caso não tenha outros Vips para utilizar!
                    Player(Index).VipDays(Player(Index).MyVip) = 0
                    Player(Index).VipStart = "00/00/0000"
                    PlayerMsg Index, "Seus dias de Vip " & Player(Index).MyVip & " Expirou, Tenha um bom jogo!", BrightRed
                    Player(Index).MyVip = 0
                    SendVipPointsInfo Index
                End If
            End If
        Else
            'Verificar se Há algum Vip com dias Restantes!
            For I = 6 To 1 Step -1
                If Player(Index).VipDays(I) > 0 Then
                    PlayerMsg Index, "#Vip " & I & " foi ativado por: " & Player(Index).VipDays(I) & " Dias.", BrightCyan
                    Player(Index).MyVip = I
                    Player(Index).VipStart = Trim$(DateValue(Date))
                    Exit For
                End If
            Next
    End If
    
    SendVipPointsInfo Index
End Sub
