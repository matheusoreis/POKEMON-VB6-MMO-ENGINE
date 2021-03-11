Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop(ByVal InvNum As Long)
    Dim FrameTime As Long
    Dim Tick As Long
    Dim TickFPS As Long
    Dim FPS As Long
    Dim i As Long
    Dim WalkTimer As Long
    Dim tmr25 As Long
    Dim tmr100 As Long, Tmr1000 As Long
    Dim tmr10000 As Long
    Dim AnFrame500 As Long
    Dim AnFrame250 As Long
    Dim Index As Long
    Dim AnQuest As Long
    Dim Confusion120 As Long
    Dim X As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' *** Start GameLoop ***
    Do While InGame
        Tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < Tick Then

            ' characters
            If NumCharacters > 0 Then
                For i = 1 To NumCharacters    'Check to unload surfaces
                    If CharacterTimer(i) > 0 Then    'Only update surfaces in use
                        If CharacterTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i)))
                            Set DDS_Character(i) = Nothing
                            CharacterTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' Paperdolls
            If NumPaperdolls > 0 Then
                For i = 1 To NumPaperdolls    'Check to unload surfaces
                    If PaperdollTimer(i) > 0 Then    'Only update surfaces in use
                        If PaperdollTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i)))
                            Set DDS_Paperdoll(i) = Nothing
                            PaperdollTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' animations
            If NumAnimations > 0 Then
                For i = 1 To NumAnimations    'Check to unload surfaces
                    If AnimationTimer(i) > 0 Then    'Only update surfaces in use
                        If AnimationTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i)))
                            Set DDS_Animation(i) = Nothing
                            AnimationTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' Items
            If numitems > 0 Then
                For i = 1 To numitems    'Check to unload surfaces
                    If ItemTimer(i) > 0 Then    'Only update surfaces in use
                        If ItemTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i)))
                            Set DDS_Item(i) = Nothing
                            ItemTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' Resources
            If NumResources > 0 Then
                For i = 1 To NumResources    'Check to unload surfaces
                    If ResourceTimer(i) > 0 Then    'Only update surfaces in use
                        If ResourceTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i)))
                            Set DDS_Resource(i) = Nothing
                            ResourceTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' spell icons
            If NumSpellIcons > 0 Then
                For i = 1 To NumSpellIcons    'Check to unload surfaces
                    If SpellIconTimer(i) > 0 Then    'Only update surfaces in use
                        If SpellIconTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i)))
                            Set DDS_SpellIcon(i) = Nothing
                            SpellIconTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' faces
            If NumFaces > 0 Then
                For i = 1 To NumFaces    'Check to unload surfaces
                    If FaceTimer(i) > 0 Then    'Only update surfaces in use
                        If FaceTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Face(i)), LenB(DDSD_Face(i)))
                            Set DDS_Face(i) = Nothing
                            FaceTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
             ' facesshiny
            If NumFacesShiny > 0 Then
                For i = 1 To NumFacesShiny    'Check to unload surfaces
                    If FaceShinyTimer(i) > 0 Then    'Only update surfaces in use
                        If FaceShinyTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_FaceShiny(i)), LenB(DDSD_FaceShiny(i)))
                            Set DDS_FaceShiny(i) = Nothing
                            FaceShinyTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' PokeIcons
            If NumPokeIcons > 0 Then
                For i = 1 To NumPokeIcons   'Check to unload surfaces
                    If PokeIconTimer(i) > 0 Then    'Only update surfaces in use
                        If PokeIconTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_PokeIcons(i)), LenB(DDSD_PokeIcons(i)))
                            Set DDS_PokeIcons(i) = Nothing
                            PokeIconTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' PokeIcons Shiny
            If NumPokeIconShiny > 0 Then
                For i = 1 To NumPokeIconShiny   'Check to unload surfaces
                    If PokeIconShinyTimer(i) > 0 Then    'Only update surfaces in use
                        If PokeIconShinyTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_PokeIconShiny(i)), LenB(DDSD_PokeIconShiny(i)))
                            Set DDS_PokeIconShiny(i) = Nothing
                            PokeIconShinyTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' check ping
            Call GetPing
            Call DrawPing
            tmr10000 = Tick + 10000
        End If

        If tmr25 < Tick Then
            InGame = IsConnected
            Call CheckKeys    ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hwnd Then
                Call CheckInputKeys    ' Check which keys were pressed
            End If

            ' check if we need to end the CD icon
            If NumSpellIcons > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i) > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i)).CDTime * 1000) < Tick Then
                                SpellCD(i) = 0
                                BltPlayerSpells
                                blthotbar
                            End If
                        End If
                    End If
                Next
            End If

            'Check Ping loading
            If LoadingPing > 0 Then
                If LoadingPing < GetTickCount Then
                    LoadingPing = 0
                End If
            End If

            ' Contador Gym
            If ContagemGym > 0 Then
                If ContagemTick < GetTickCount Then
                    ContagemGym = ContagemGym - 1
                    ContagemTick = 1000 + GetTickCount
                End If
            End If

            'Confusion MeAnimation
            If Confusion120 < GetTickCount Then
                For i = 1 To Player_HighIndex
                    If GetPlayerEquipmentNgt(i, weapon, 5) > 0 Then
                        MeAnimation 17, GetPlayerX(i), GetPlayerY(i) - 1, 1, i
                    End If

                    If GetPlayerEquipmentNgt(i, weapon, 6) > 0 Then
                        MeAnimation 21, GetPlayerX(i), GetPlayerY(i) - 1, 1, i
                    End If
                Next
                Confusion120 = 350 + GetTickCount
            End If

            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If PlayerSpells(SpellBuffer) > 0 Then
                    If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < Tick Then
                        SpellBuffer = 0
                        SpellBufferTimer = 0
                    End If
                Else
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            ' Informações do Jogador
            For i = 1 To Player_HighIndex
                ' Pescando > 0 then
                If Player(i).InFishing > 0 Then
                    If Player(i).InFishing + 10000 < GetTickCount Then Player(i).InFishing = 0
                End If

                ' Scan > 0 then
                If Player(i).ScanTime > 0 Then
                    If Player(i).ScanTime + 2000 < GetTickCount Then Player(i).ScanTime = 0
                End If

                ' Parado
                If Player(i).Parado > 0 Then
                    If Player(i).Parado < GetTickCount Then
                        Player(i).Parado = 0
                    End If
                End If

                ' Pular
                If Player(i).PuloStatus > 0 Then
                    If Player(i).PuloSlide > 0 Then
                        Player(i).PuloSlide = Player(i).PuloSlide - 1
                    Else
                        Player(i).PuloStatus = 0
                    End If
                End If
            Next

            If AnQuest < GetTickCount Then
                If AnimQuest <= 4 And AnimQuest > 0 Then
                    AnimQuest = AnimQuest - 1
                Else
                    AnimQuest = 4
                End If
                AnQuest = 250 + GetTickCount
            End If

            'AnimFrame 1Seg
            If AnFrame250 < GetTickCount Then

                If InGame = True Then
                    For Index = 1 To Player_HighIndex
                        If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then

                            If Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).AnimFrame(1) > 0 Then
                                If Player(Index).Flying = 0 Then
                                    If Player(Index).AnimFrame = Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).AnimFrame(1) Then
                                        Player(Index).AnimFrame = 0
                                    Else
                                        Player(Index).AnimFrame = Player(Index).AnimFrame + 1
                                    End If
                                End If
                            End If

                            If Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).AnimFrame(2) > 0 Then
                                If Player(Index).Flying > 0 Then
                                    If Player(Index).AnimFrame = Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).AnimFrame(2) Then
                                        Player(Index).AnimFrame = 0
                                    Else
                                        Player(Index).AnimFrame = Player(Index).AnimFrame + 1
                                    End If
                                End If
                            End If

                        End If
                    Next
                End If

                AnFrame250 = 150 + GetTickCount
            End If

            If CanMoveNow Then
                Call CheckMovement    ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If

            ' Change map animation every 250 milliseconds
            If MapAnimTimer < Tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = Tick + 250
            End If

            ' Update inv animation
            If numitems > 0 Then
                If tmr100 < Tick Then
                    BltAnimatedInvItems
                    tmr100 = Tick + 100
                End If
            End If

            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next

            tmr25 = Tick + 25
        End If

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then

            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex
                If Map.Npc(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i

            WalkTimer = Tick + 30    ' edit this value to change WalkTimer
        End If
                    
        Dim visibleWarp As Boolean
        visibleWarp = False
        
        For i = 1 To MAX_INV ' Ta certo isso?
            If PlayerInv(i).PokeInfo.Pokemon > 0 Then ' isso diz se tem poke? s blz
                For X = 1 To UBound(PlayerInv(i).PokeInfo.Spells)
                    If PlayerInv(i).PokeInfo.Spells(X) = 12 Then
                        visibleWarp = True
                        Exit For
                    End If
                Next
            End If
            If visibleWarp Then Exit For
        Next
        frmMain.picTele.Visible = visibleWarp
    
        '123

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        DoEvents

        If Not GettingMap Then
            frmMain.PicLoading.ZOrder 0
            If frmMain.PicLoading.Visible = True Then frmMain.PicLoading.Visible = False
        Else
            frmMain.PicLoading.ZOrder 0
            If frmMain.PicLoading.Visible = False Then frmMain.PicLoading.Visible = True
        End If

        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + 15
                DoEvents
                Sleep 1
            Loop
        End If

        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If

        'loop mapmusic if needed and its a mp3 file
        LoopMp3

    Loop

    frmMain.Visible = False

    If isLogging Then
        isLogging = False
        frmMain.picScreen.Visible = False
        frmMenu.Visible = True
        GettingMap = True
        StopMusic
        PlayMusic Options.MenuMusic
    Else
        ' Shutdown the game
        frmLoad.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessMovement(ByVal Index As Long)
    Dim MovementSpeed As Long
    Dim itemNum As Long
    Dim VelocidadeItem As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerEquipment(Index, weapon) > 0 Then
        VelocidadeItem = Item(GetPlayerEquipment(Index, weapon)).vel
    Else
        If GetPlayerEquipment(Index, Shield) > 0 Then
            VelocidadeItem = Item(GetPlayerEquipment(Index, Shield)).vel
        End If
    End If

    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving
    Case MOVING_WALKING: MovementSpeed = ((ElapsedTime / 600) * ((WALK_SPEED + VelocidadeItem) * SIZE_X))
    Case MOVING_RUNNING: MovementSpeed = ((ElapsedTime / 600) * ((RUN_SPEED + VelocidadeItem) * SIZE_X))
    Case Else: Exit Sub
    End Select

    If Player(Index).Step = 0 Then Player(Index).Step = 1

    Select Case GetPlayerDir(Index)
    Case DIR_UP
        Player(Index).YOffset = Player(Index).YOffset - MovementSpeed
        If Player(Index).YOffset < 0 Then Player(Index).YOffset = 0
    Case DIR_DOWN
        Player(Index).YOffset = Player(Index).YOffset + MovementSpeed
        If Player(Index).YOffset > 0 Then Player(Index).YOffset = 0
    Case DIR_LEFT
        Player(Index).XOffset = Player(Index).XOffset - MovementSpeed
        If Player(Index).XOffset < 0 Then Player(Index).XOffset = 0
    Case DIR_RIGHT
        Player(Index).XOffset = Player(Index).XOffset + MovementSpeed
        If Player(Index).XOffset > 0 Then Player(Index).XOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).XOffset >= 0) And (Player(Index).YOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        Else
            If (Player(Index).XOffset <= 0) And (Player(Index).YOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)

' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then

        Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).YOffset < 0 Then MapNpc(MapNpcNum).YOffset = 0

        Case DIR_DOWN
            MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).YOffset > 0 Then MapNpc(MapNpcNum).YOffset = 0

        Case DIR_LEFT
            MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).XOffset < 0 Then MapNpc(MapNpcNum).XOffset = 0

        Case DIR_RIGHT
            MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
            If MapNpc(MapNpcNum).XOffset > 0 Then MapNpc(MapNpcNum).XOffset = 0

        End Select

        ' Check if completed walking over to the next tile
        If MapNpc(MapNpcNum).Moving > 0 Then
            If MapNpc(MapNpcNum).Dir = DIR_RIGHT Or MapNpc(MapNpcNum).Dir = DIR_DOWN Then
                If (MapNpc(MapNpcNum).XOffset >= 0) And (MapNpc(MapNpcNum).YOffset >= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            Else
                If (MapNpc(MapNpcNum).XOffset <= 0) And (MapNpc(MapNpcNum).YOffset <= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckMapGetItem()
    Dim Buffer As New clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer

    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = GetTickCount
            Buffer.WriteLong CMapGetItem
            SendData Buffer.ToArray()
        End If
    End If

    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()
    Dim Buffer As clsBuffer
    Dim attackspeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ControlDown Then

        If SpellBuffer > 0 Then Exit Sub    ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub    ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, weapon)).speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then

                Set Buffer = New clsBuffer
                Buffer.WriteLong CAttack
                SendData Buffer.ToArray()
                Set Buffer = Nothing
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsTryingToMove() As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanMove() As Boolean
    Dim d As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If

    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If

    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If

    ' not in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        frmMain.picCover.Visible = False
        frmMain.picBank.Visible = False
    End If

    If frmMain.PicPokeInicial.Visible = True Then
        CanMove = False
        Exit Function
    End If

    d = GetPlayerDir(MyIndex)

    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                If GetPlayerEquipment(MyIndex, weapon) = 0 Or GetPlayerEquipment(MyIndex, weapon) = 247 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                If GetPlayerEquipment(MyIndex, weapon) = 0 Or GetPlayerEquipment(MyIndex, weapon) = 247 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                If GetPlayerEquipment(MyIndex, weapon) = 0 Or GetPlayerEquipment(MyIndex, weapon) = 247 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                If GetPlayerEquipment(MyIndex, weapon) = 0 Or GetPlayerEquipment(MyIndex, weapon) = 247 Then
                    Call MapEditorLeaveMap
                    Call SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
            End If

            CanMove = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
    Dim X As Long, Y As Long
    Dim X2 As Long, Y2 As Long
    Dim i As Long, Number As Long
    Dim AntSlide As Byte, PokemonId As String, PokeRuido As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    CheckDirection = False

    ' check directional blocking
    If isDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, Direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case Direction
    Case DIR_UP
        X = GetPlayerX(MyIndex)
        Y = GetPlayerY(MyIndex) - 1
        X2 = GetPlayerX(MyIndex)
        Y2 = GetPlayerY(MyIndex)
    Case DIR_DOWN
        X = GetPlayerX(MyIndex)
        Y = GetPlayerY(MyIndex) + 1
        X2 = GetPlayerX(MyIndex)
        Y2 = GetPlayerY(MyIndex)
    Case DIR_LEFT
        X = GetPlayerX(MyIndex) - 1
        Y = GetPlayerY(MyIndex)
        X2 = GetPlayerX(MyIndex)
        Y2 = GetPlayerY(MyIndex)
    Case DIR_RIGHT
        X = GetPlayerX(MyIndex) + 1
        Y = GetPlayerY(MyIndex)
        X2 = GetPlayerX(MyIndex)
        Y2 = GetPlayerY(MyIndex)
    End Select

    'Limite Pokémon Walking
    If GetPlayerEquipment(MyIndex, weapon) > 0 And GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon) > 0 Then
        If Player(MyIndex).Flying = 0 Then
            If Not isInRange(15, X, Y, Player(MyIndex).TPX, Player(MyIndex).TPY) Then
                CheckDirection = True
                Exit Function
            End If
        Else
            If Not isInRange(20, X, Y, Player(MyIndex).TPX, Player(MyIndex).TPY) Then
                CheckDirection = True
                Exit Function
            End If
        End If
    End If

    If GetPlayerFlying(MyIndex) Then Exit Function

    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is water or not
    If Player(MyIndex).InSurf = 3 Or Player(MyIndex).Equipment(1) > 0 Then
        If Map.Tile(X, Y).Type = TILE_TYPE_WATER Then
            CheckDirection = True
            Exit Function
        End If
    End If

    ' Check Direction
    If Map.Tile(X, Y).Type = TILE_TYPE_SLIDE Then
        Select Case Direction
        Case DIR_UP
            AntSlide = DIR_DOWN
        Case DIR_LEFT
            AntSlide = DIR_RIGHT
        Case DIR_RIGHT
            AntSlide = DIR_LEFT
        Case DIR_DOWN
            AntSlide = DIR_UP
        End Select

        If Map.Tile(X, Y).Data1 = AntSlide Then
            CheckDirection = True
            Exit Function
        Else
            Player(MyIndex).PuloStatus = 1
            Player(MyIndex).PuloSlide = 15
        End If
    End If

    ' Check to see if the map tile is tree or not
    If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE And CheckResourceStatCut(MyIndex, X, Y) = False Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the key door is open or not
    If Map.Tile(X, Y).Type = TILE_TYPE_KEY Then

        ' This actually checks if its open or not
        If TempTile(X, Y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If

    ' Check to see if a player is already on that tile
    If Map.Moral = 0 Or Map.Moral = 2 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        If Not GetPlayerFlying(i) Then
                            CheckDirection = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next i
    End If

    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If MapNpc(i).Desmaiado = False Then
                If MapNpc(i).X = X Then
                    If MapNpc(i).Y = Y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next

    ' Check Point Trainer
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If Player(i).TPX = X Then
                If Player(i).TPY = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next

    'Check to see if the map tile is Grass or not
    If Map.Tile(X, Y).Type = TILE_TYPE_GRASS Then
        If Map.Tile(X2, Y2).Type = TILE_TYPE_GRASS Then
            MeAnimation 10, GetPlayerX(MyIndex), GetPlayerY(MyIndex)
        End If
    End If

    'Check player in surf
    If Player(MyIndex).InSurf = 1 Then
        MeAnimation 14, GetPlayerX(MyIndex), GetPlayerY(MyIndex)
    End If

    'Check Sound You Pokémon
    If GetPlayerEquipment(MyIndex, weapon) > 0 Then
        PokeRuido = 100 * Rnd
        If PokeRuido <= 5 Then
            If GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon) > 0 Then
                Select Case GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)
                Case 1 To 9
                    PokemonId = "00" & GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)
                Case 10 To 99
                    PokemonId = "0" & GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)
                Case Else
                    PokemonId = GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)
                End Select

                PlaySound "PokeSounds\" & PokemonId & ".mp3", -1, -1
            End If
        End If
    End If

    'Check PicChat Tirar quando tiver visivel e sair da Tile ...
    If frmMain.picPlaca.Visible = True Then frmMain.picPlaca.Visible = False

    'Check Sign Tile
    If Map.Tile(X, Y).Type = TILE_TYPE_SIGN Then
        If GetPlayerEquipment(MyIndex, weapon) = 0 Then
            Number = Map.Tile(X, Y).Data1
            frmMain.lblChat.Caption = GetVar(App.Path & "\Data Files\chat.ini", "CHAT", Val(Number))
            frmMain.picPlaca.Visible = True
            frmMain.picPlaca.top = (frmMain.ScaleHeight / 2) - (frmMain.picPlaca.Height / 2)
            frmMain.picPlaca.Left = (frmMain.ScaleWidth / 2) - (frmMain.picPlaca.Width / 2)
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "checkDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub CheckMovement()

' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Player(MyIndex).EvolPermition = 1 Then Exit Sub
    If Player(MyIndex).InFishing > 0 Then Exit Sub
    If Player(MyIndex).ScanTime > 0 Then Exit Sub

    If Player(MyIndex).Attacking = 1 Then Exit Sub

    If GetPlayerEquipment(MyIndex, weapon) > 0 Then
        If GetPlayerEquipmentNgt(MyIndex, weapon, 2) > 0 Then Exit Sub
    End If

    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MOVING_RUNNING
            Else
                Player(MyIndex).Moving = MOVING_WALKING
            End If

            Select Case GetPlayerDir(MyIndex)
            Case DIR_UP
                Call SendPlayerMove
                Player(MyIndex).YOffset = PIC_Y
                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
            Case DIR_DOWN
                Call SendPlayerMove
                Player(MyIndex).YOffset = PIC_Y * -1
                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
            Case DIR_LEFT
                Call SendPlayerMove
                Player(MyIndex).XOffset = PIC_X
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
            Case DIR_RIGHT
                Call SendPlayerMove
                Player(MyIndex).XOffset = PIC_X * -1
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select

            If Player(MyIndex).XOffset = 0 Then
                If Player(MyIndex).YOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub UpdateDrawMapName()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Map.LevelPoke(1) <= Map.LevelPoke(2) Then
        DrawMapName = Trim$(Map.Name) & " - Lvl " & Map.LevelPoke(1) & "~" & Map.LevelPoke(2) + 1
    Else
        DrawMapName = Trim$(Map.Name) & " - Lvl " & Map.LevelPoke(2) & "~" & Map.LevelPoke(1) + 1
    End If

    DrawMapNameX = Camera.Left + ((MAX_MAPX + 1) * PIC_X / 2) - getWidth(TexthDC, Trim$(DrawMapName))
    DrawMapNameY = Camera.top + 1

    Select Case Map.Moral
    Case MAP_MORAL_NONE
        DrawMapNameColor = QBColor(BrightRed)
    Case MAP_MORAL_SAFE
        DrawMapNameColor = QBColor(White)
    Case Else
        DrawMapNameColor = QBColor(White)
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateDrawMapName", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UseItem()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ForgetSpell(ByVal spellslot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    ' dont let them forget a spell which is in CD
    If SpellCD(spellslot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If

    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellslot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If

    If PlayerSpells(spellslot) > 0 Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CForgetSpell
        Buffer.WriteLong spellslot
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellslot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    If SpellCD(spellslot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If

    If PlayerSpells(spellslot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellslot)).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellslot)).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellslot) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong CCast
                Buffer.WriteLong spellslot
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                SpellBuffer = spellslot
                SpellBufferTimer = GetTickCount
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearTempTile()
    Dim X As Long
    Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim TempTile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            TempTile(X, Y).DoorOpen = NO
        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearTempTile", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DevMsg(ByVal text As String, ByVal color As Byte)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(text, color)
        End If
    End If

    Debug.Print text

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "TwipsToPixels", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) <= 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) <= 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawPing()
    Dim PingToDraw As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    PingToDraw = Ping

    Select Case Ping
    Case -1
        PingToDraw = "Syncing"
    Case 0 To 5
        PingToDraw = "Local"
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPing", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateDescWindow(ByVal itemNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Long
    Dim FirstLetter As String * 1
    Dim Name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FirstLetter = LCase$(Left$(Trim$(Item(itemNum).Name), 1))

    If FirstLetter = "$" Then
        Name = (Mid$(Trim$(Item(itemNum).Name), 2, Len(Trim$(Item(itemNum).Name)) - 1))
    Else
        Name = Trim$(Item(itemNum).Name)
    End If

    ' check for off-screen
    If Y + frmMain.picItemDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picItemDesc.Height
    End If

    ' set z-order
    frmMain.picItemDesc.ZOrder (0)

    With frmMain
        .picItemDesc.top = Y
        .picItemDesc.Left = X
        .picItemDesc.Visible = True

        If LastItemDesc = itemNum Then Exit Sub    ' exit out after setting x + y so we don't reset values

        ' set the name
        Select Case Item(itemNum).Rarity
        Case 0    ' white
            .lblItemName.ForeColor = RGB(255, 255, 255)
        Case 1    ' green
            .lblItemName.ForeColor = RGB(117, 198, 92)
        Case 2    ' blue
            .lblItemName.ForeColor = RGB(103, 140, 224)
        Case 3    ' maroon
            .lblItemName.ForeColor = RGB(205, 34, 0)
        Case 4    ' purple
            .lblItemName.ForeColor = RGB(193, 104, 204)
        Case 5    ' orange
            .lblItemName.ForeColor = RGB(217, 150, 64)
        End Select

        ' set captions
        .lblItemName.Caption = Name
        .lblItemDesc.Caption = Trim$(Item(itemNum).Desc)

        ' render the item
        BltItemDesc itemNum
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateDescWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CacheResources()
    Dim X As Long, Y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Resource_Count = 0

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateActionMsg(ByVal Message As String, ByVal color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .Message = Message
        .color = color
        .Type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .X = X
        .Y = Y
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If

    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ActionMsg(Index).Message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).X = 0
    ActionMsg(Index).Y = 0

    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
    Dim looptime As Long
    Dim Layer As Long
    Dim FrameCount As Long
    Dim lockindex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub

    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)

            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).FrameIndex(Layer) = 0 Then AnimInstance(Index).FrameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1

            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).timer(Layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(Index).FrameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).FrameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).FrameIndex(Layer) = AnimInstance(Index).FrameIndex(Layer) + 1
                End If
                AnimInstance(Index).timer(Layer) = GetTickCount
            End If
        End If
    Next

    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InShop = shopnum
    ShopAction = 0
    'frmMain.picCover.Visible = True
    frmMain.picShop.Visible = True
    frmMain.picShop.top = (frmMain.ScaleHeight / 2) - (frmMain.picShop.Height / 2)
    frmMain.picShop.Left = (frmMain.ScaleWidth / 2) - (frmMain.picShop.Width / 2)
    BltShop

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If

    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If

    GetBankItemNum = Bank.Item(bankslot).num

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemNum As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Bank.Item(bankslot).num = itemNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetBankItemValue = Bank.Item(bankslot).value

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Bank.Item(bankslot).value = ItemValue

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef Dir As Byte, ByVal block As Boolean)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If block Then
        blockvar = blockvar Or (2 ^ Dir)
    Else
        blockvar = blockvar And Not (2 ^ Dir)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single) As Long
    Dim top As Long, Left As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsHotbarSlot = 0

    For i = 1 To MAX_HOTBAR
        top = HotbarTop
        Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
        If X >= Left And X <= Left + PIC_X Then
            If Y >= top And Y <= top + PIC_Y Then
                IsHotbarSlot = i
                Exit Function
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim soundName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If entityNum <= 0 Then Exit Sub

    ' find the sound
    Select Case entityType
        ' animations
    Case SoundEntity.seAnimation
        If entityNum > MAX_ANIMATIONS Then Exit Sub
        soundName = Trim$(Animation(entityNum).sound)
        ' items
    Case SoundEntity.seItem
        If entityNum > MAX_ITEMS Then Exit Sub
        soundName = Trim$(Item(entityNum).sound)
        ' npcs
    Case SoundEntity.seNpc
        If entityNum > MAX_NPCS Then Exit Sub
        soundName = Trim$(Npc(entityNum).sound)
        ' resources
    Case SoundEntity.seResource
        If entityNum > MAX_RESOURCES Then Exit Sub
        soundName = Trim$(Resource(entityNum).sound)
        ' spells
    Case SoundEntity.seSpell
        If entityNum > MAX_SPELLS Then Exit Sub
        soundName = Trim$(Spell(entityNum).sound)
        ' other
    Case Else
        Exit Sub
    End Select

    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    If isInRange(5, X, Y, GetPlayerX(MyIndex), GetPlayerY(MyIndex)) Then
        ' play the sound
        PlaySound soundName, -1, -1
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub

    ' set global dialogue index
    dialogueIndex = diIndex

    ' set the global dialogue data
    dialogueData1 = Data1

    ' set the captions
    frmMain.lblDialogue_Title.Caption = diTitle
    frmMain.lblDialogue_Text.Caption = diText

    ' show/hide buttons
    If Not isYesNo Then
        frmMain.lblDialogueBtn(1).Visible = True    ' Okay button
        frmMain.lblDialogueBtn(2).Visible = False    ' Yes button
        frmMain.lblDialogueBtn(3).Visible = False    ' No button
    Else
        frmMain.lblDialogueBtn(1).Visible = False    ' Okay button
        frmMain.lblDialogueBtn(2).Visible = True    ' Yes button
        frmMain.lblDialogueBtn(3).Visible = True    ' No button
    End If

    ' show the dialogue box
    frmMain.picDialogue.Visible = True
    frmMain.picDialogue.top = (frmMain.ScaleHeight / 2) - (frmMain.picDialogue.Height / 2)
    frmMain.picDialogue.Left = (frmMain.ScaleWidth / 2) - (frmMain.picDialogue.Width / 2)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
' find out which button
    If Index = 1 Then    ' okay button
' dialogue index
        Select Case dialogueIndex

        End Select
    ElseIf Index = 2 Then    ' yes button
        ' dialogue index
        Select Case dialogueIndex
        Case DIALOGUE_TYPE_TRADE
            SendAcceptTradeRequest
        Case DIALOGUE_TYPE_FORGET
            ForgetSpell dialogueData1
        Case DIALOGUE_TYPE_PARTY
            SendAcceptParty
        Case DIALOGUE_TYPE_QUEST
            SendQuestCommand 1
        Case DIALOGUE_TYPE_PM
            SendChatComando 1, vbNullString
        Case DIALOGUE_TYPE_LT
            SendLutarComando 2, 0, 0, 0, vbNullString
        End Select
    ElseIf Index = 3 Then    ' no button
        ' dialogue index
        Select Case dialogueIndex
        Case DIALOGUE_TYPE_TRADE
            SendDeclineTradeRequest
        Case DIALOGUE_TYPE_PARTY
            SendDeclineParty
        Case DIALOGUE_TYPE_PM
            SendChatComando 2, vbNullString
        Case DIALOGUE_TYPE_LT
            SendLutarComando 3, 0, 0, 0, vbNullString
        End Select
    End If
End Sub

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - X2) ^ 2 + (y1 - Y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Sub UpdateSpellWindow(ByVal SpellNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' check for off-screen
    If Y + frmMain.picSpellDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picSpellDesc.Height
    End If

    With frmMain
        .picSpellDesc.top = Y
        .picSpellDesc.Left = X
        .picSpellDesc.Visible = True

        If LastSpellDesc = SpellNum Then Exit Sub

        .lblSpellName.Caption = Trim$(Spell(SpellNum).Name)
        .lblSpellDesc.Caption = Trim$(Spell(SpellNum).Desc)
        BltSpellDesc SpellNum
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdteSpellWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdatePokeWindow(ByVal InvNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Command As Long, Optional ByVal QuestNum As Integer)
    Dim i As Long, FemQntia As Long
    Dim Name As String
    Dim Felicity As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' check for off-screen
    If Y + frmMain.picPokeDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picPokeDesc.Height
    End If

    ' set z-order
    frmMain.picPokeDesc.ZOrder (0)

    With frmMain
        .picPokeDesc.top = Y
        .picPokeDesc.Left = X
        .picPokeDesc.Visible = True

        BltFacePokemon (InvNum)
        If LastItemPoke = InvNum Then Exit Sub    ' exit out after setting x + y so we don't reset values
        If Command = 0 Then    'Inventario        If GetPlayerInvItemPokeInfoPokemon(MyIndex, InvNum) = 0 Then Exit Sub
                If GetPlayerInvItemShiny(MyIndex, InvNum) = 0 Then
                    .lblPokeInfoDesc(0).Caption = Trim$(Pokemon(GetPlayerInvItemPokeInfoPokemon(MyIndex, InvNum)).Name)
                Else
                    .lblPokeInfoDesc(0).Caption = "Shiny " & Trim$(Pokemon(GetPlayerInvItemPokeInfoPokemon(MyIndex, InvNum)).Name)
                End If
                'If GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, i) >= 50 Then
                'If PlayerInv(InvNum).PokeInfo.Stat(1) + PlayerInv(InvNum).PokeInfo.Stat(2) + PlayerInv(InvNum).PokeInfo.Stat(3) + PlayerInv(InvNum).PokeInfo.Stat(4) + PlayerInv(InvNum).PokeInfo.Stat(5) > 50 Then
                 '   .lblPokeInfoDesc(0).ForeColor = QBColor(Red)
                'Else
                  '  .lblPokeInfoDesc(0).ForeColor = QBColor(White)
                  '  End If
'Mecher Futuramente

            .lblPokeInfoDesc(3).Caption = "Exp:" & GetPlayerInvItemPokeInfoExp(MyIndex, InvNum) & "/" & GetInvPokeNextLevel(InvNum, 0)
            .lblPokeInfoDesc(4).Caption = "Hp:" & GetPlayerInvItemPokeInfoVital(MyIndex, InvNum, 1) & "/" & GetPlayerInvItemPokeInfoMaxVital(MyIndex, InvNum, 1)
            .lblPokeInfoDesc(5).Caption = "PP:" & GetPlayerInvItemPokeInfoVital(MyIndex, InvNum, 2) & "/" & GetPlayerInvItemPokeInfoMaxVital(MyIndex, InvNum, 2)
            .lblPokeInfoDesc(6).Caption = "Lv: " & GetPlayerInvItemPokeInfoLevel(MyIndex, InvNum)

            If GetPlayerInvItemBerry(MyIndex, InvNum, 1) > 0 Then
                .lblPokeInfoDesc(8).Caption = "Atq:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 1) & " + " & GetPlayerInvItemBerry(MyIndex, InvNum, 1)
            Else
                .lblPokeInfoDesc(8).Caption = "Atq:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 1)
            End If

            If GetPlayerInvItemBerry(MyIndex, InvNum, 2) > 0 Then
                .lblPokeInfoDesc(9).Caption = "Def:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 2) & " + " & GetPlayerInvItemBerry(MyIndex, InvNum, 2)
            Else
                .lblPokeInfoDesc(9).Caption = "Def:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 2)
            End If

            If GetPlayerInvItemBerry(MyIndex, InvNum, 3) > 0 Then
                .lblPokeInfoDesc(10).Caption = "EAtq:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 3) & " + " & GetPlayerInvItemBerry(MyIndex, InvNum, 3)
            Else
                .lblPokeInfoDesc(10).Caption = "EAtq:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 3)
            End If

            If GetPlayerInvItemBerry(MyIndex, InvNum, 5) > 0 Then
                .lblPokeInfoDesc(11).Caption = "EDef:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 5) & " + " & GetPlayerInvItemBerry(MyIndex, InvNum, 5)
            Else
                .lblPokeInfoDesc(11).Caption = "EDef:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 5)
            End If

            If GetPlayerInvItemBerry(MyIndex, InvNum, 4) > 0 Then
                .lblPokeInfoDesc(12).Caption = "Vel:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 4) & " + " & GetPlayerInvItemBerry(MyIndex, InvNum, 4)
            Else
                .lblPokeInfoDesc(12).Caption = "Vel:" & GetPlayerInvItemPokeInfoStat(MyIndex, InvNum, 4)
            End If

            Select Case GetPlayerInvItemFelicidade(MyIndex, InvNum)
            Case 0 To 50
                Felicity = " Triste"
            Case 51 To 100
                Felicity = " Normal"
            Case 101 To 200
                Felicity = " Esperto"
            Case 201 To 300
                Felicity = " Alegre"
            Case 301 To 400
                Felicity = " Feliz"
            Case Else
                Felicity = " Muito Feliz"
            End Select

            .lblPokeInfoDesc(1).Caption = "Humor:" & Felicity

            If GetPlayerInvItemSexo(MyIndex, InvNum) = 0 Then
                .lblPokeInfoDesc(2).Caption = "Sexo: Masculino"
            Else
                .lblPokeInfoDesc(2).Caption = "Sexo: Feminino "
            End If

            For i = 1 To 4
                If GetPlayerInvItemPokeInfoSpell(MyIndex, InvNum, i) > 0 Then
                    .lblPokeInfoDesc(12 + i).Caption = Trim$(Spell(GetPlayerInvItemPokeInfoSpell(MyIndex, InvNum, i)).Name)
                Else
                    .lblPokeInfoDesc(12 + i).Caption = "SlotVazio"
                End If
            Next
        End If

        If Command = 1 Then
            If GetPlayerBankItemPokemon(InvNum) = 0 Then Exit Sub

            If GetPlayerBankItemShiny(MyIndex, InvNum) = 0 Then
                .lblPokeInfoDesc(0).Caption = Trim$(Pokemon(GetPlayerBankItemPokemon(InvNum)).Name) & " - Level:" & GetPlayerBankItemLevel(InvNum)
            Else
                .lblPokeInfoDesc(0).Caption = "Shiny" & Trim$(Pokemon(GetPlayerBankItemPokemon(InvNum)).Name) & " - Level:" & GetPlayerBankItemLevel(InvNum)
            End If

            .lblPokeInfoDesc(3).Caption = "Exp:" & GetPlayerBankItemExp(InvNum) & "/" & GetInvPokeNextLevel(InvNum, 1)
            .lblPokeInfoDesc(4).Caption = "Hp:" & GetPlayerBankItemVital(InvNum, 1) & "/" & GetPlayerBankItemMaxVital(InvNum, 1)
            .lblPokeInfoDesc(5).Caption = "PP:" & GetPlayerBankItemVital(InvNum, 2) & "/" & GetPlayerBankItemMaxVital(InvNum, 2)
            .lblPokeInfoDesc(8).Caption = "Atq:" & GetPlayerBankItemStat(InvNum, 1)
            .lblPokeInfoDesc(9).Caption = "Def:" & GetPlayerBankItemStat(InvNum, 2)
            .lblPokeInfoDesc(10).Caption = "EAtq:" & GetPlayerBankItemStat(InvNum, 3)
            .lblPokeInfoDesc(11).Caption = "EDef:" & GetPlayerBankItemStat(InvNum, 5)
            .lblPokeInfoDesc(12).Caption = "Vel:" & GetPlayerBankItemStat(InvNum, 4)
            .lblPokeInfoDesc(1).Caption = "Humor:" & GetPlayerBankItemFelicidade(MyIndex, InvNum)
            If GetPlayerBankItemSexo(MyIndex, InvNum) = 0 Then
                .lblPokeInfoDesc(2).Caption = "Sexo: Masculino"
            Else
                .lblPokeInfoDesc(2).Caption = "Sexo: Feminino"
            End If

            For i = 1 To 4
                If GetPlayerBankItemSpell(InvNum, i) > 0 Then
                    .lblPokeInfoDesc(12 + i).Caption = Trim$(Spell(GetPlayerBankItemSpell(InvNum, i)).Name)
                Else
                    .lblPokeInfoDesc(12 + i).Caption = "SlotVazio"
                End If
            Next
        End If

        If Command = 2 Then

            If Leilao(InvNum).Poke.Pokemon = 0 Then Exit Sub
            .lblPokeInfoDesc(0).Caption = Trim$(Pokemon(Leilao(InvNum).Poke.Pokemon).Name) & " - Level:" & Leilao(InvNum).Poke.Level
            .lblPokeInfoDesc(3).Caption = "Exp:" & Leilao(InvNum).Poke.Exp & "/" & GetInvPokeNextLevel(InvNum, 2)
            .lblPokeInfoDesc(4).Caption = "Hp:" & Leilao(InvNum).Poke.Vital(1) & "/" & Leilao(InvNum).Poke.MaxVital(1)
            .lblPokeInfoDesc(5).Caption = "PP:" & Leilao(InvNum).Poke.Vital(2) & "/" & Leilao(InvNum).Poke.MaxVital(2)
            .lblPokeInfoDesc(8).Caption = "Atq:" & Leilao(InvNum).Poke.Stat(1)
            .lblPokeInfoDesc(9).Caption = "Def:" & Leilao(InvNum).Poke.Stat(2)
            .lblPokeInfoDesc(10).Caption = "EAtq:" & Leilao(InvNum).Poke.Stat(3)
            .lblPokeInfoDesc(11).Caption = "EDef:" & Leilao(InvNum).Poke.Stat(5)
            .lblPokeInfoDesc(12).Caption = "Vel:" & Leilao(InvNum).Poke.Stat(4)
            .lblPokeInfoDesc(1).Caption = "Humor:" & Leilao(InvNum).Poke.Felicidade
            If Leilao(InvNum).Poke.Sexo = 0 Then
                .lblPokeInfoDesc(2).Caption = "Sexo: Masculino"
            Else
                .lblPokeInfoDesc(2).Caption = "Sexo: Feminino"
            End If


            For i = 1 To 4
                If Leilao(InvNum).Poke.Spells(i) > 0 Then
                    .lblPokeInfoDesc(12 + i).Caption = Trim$(Spell(Leilao(InvNum).Poke.Spells(i)).Name)
                Else
                    .lblPokeInfoDesc(12 + i).Caption = "SlotVazio"
                End If
            Next

        End If

        If Command = 3 Then
            If TradeTheirOffer(InvNum).PokeInfo.Pokemon = 0 Then Exit Sub

            If TradeTheirOffer(InvNum).PokeInfo.Shiny = 0 Then
                .lblPokeInfoDesc(0).Caption = Trim$(Pokemon(TradeTheirOffer(InvNum).PokeInfo.Pokemon).Name) & " - Level:" & TradeTheirOffer(InvNum).PokeInfo.Level
            Else
                .lblPokeInfoDesc(0).Caption = "Shiny" & Trim$(Pokemon(TradeTheirOffer(InvNum).PokeInfo.Pokemon).Name) & " - Level:" & TradeTheirOffer(InvNum).PokeInfo.Level
            End If

            .lblPokeInfoDesc(3).Caption = "Exp:" & TradeTheirOffer(InvNum).PokeInfo.Exp & "/" & GetInvPokeNextLevel(InvNum, 3)
            .lblPokeInfoDesc(4).Caption = "Hp:" & TradeTheirOffer(InvNum).PokeInfo.Vital(1) & "/" & TradeTheirOffer(InvNum).PokeInfo.MaxVital(1)
            .lblPokeInfoDesc(5).Caption = "PP:" & TradeTheirOffer(InvNum).PokeInfo.Vital(2) & "/" & TradeTheirOffer(InvNum).PokeInfo.MaxVital(2)
            .lblPokeInfoDesc(8).Caption = "Atq:" & TradeTheirOffer(InvNum).PokeInfo.Stat(1)
            .lblPokeInfoDesc(9).Caption = "Def:" & TradeTheirOffer(InvNum).PokeInfo.Stat(2)
            .lblPokeInfoDesc(10).Caption = "EAtq:" & TradeTheirOffer(InvNum).PokeInfo.Stat(3)
            .lblPokeInfoDesc(11).Caption = "EDef:" & TradeTheirOffer(InvNum).PokeInfo.Stat(5)
            .lblPokeInfoDesc(12).Caption = "Vel:" & TradeTheirOffer(InvNum).PokeInfo.Stat(4)
            .lblPokeInfoDesc(1).Caption = "Humor:" & TradeTheirOffer(InvNum).PokeInfo.Felicidade

            If TradeTheirOffer(InvNum).PokeInfo.Sexo = 0 Then
                .lblPokeInfoDesc(2).Caption = "Sexo: Masculino"
            Else
                .lblPokeInfoDesc(2).Caption = "Sexo: Feminino"
            End If

            For i = 1 To 4
                If GetPlayerBankItemSpell(InvNum, i) > 0 Then
                    .lblPokeInfoDesc(12 + i).Caption = Trim$(Spell(TradeTheirOffer(InvNum).PokeInfo.Spells(i)).Name)
                Else
                    .lblPokeInfoDesc(12 + i).Caption = "SlotVazio"
                End If
            Next

        End If

        If Command = 4 Then

            If Quest(QuestNum).PokeRew(InvNum) = 0 Then Exit Sub

            .lblPokeInfoDesc(0).Caption = Trim$(Pokemon(Quest(QuestNum).PokeRew(InvNum)).Name) & " - Level: " & Quest(QuestNum).ValueRew(InvNum)

            .lblPokeInfoDesc(3).Caption = "Exp: 0/" & GetInvPokeNextLevel(Quest(QuestNum).ValueRew(InvNum), 4, Quest(QuestNum).PokeRew(InvNum))
            .lblPokeInfoDesc(4).Caption = "Hp: " & Pokemon(Quest(QuestNum).PokeRew(InvNum)).Vital(1)
            .lblPokeInfoDesc(5).Caption = "PP:" & Pokemon(Quest(QuestNum).PokeRew(InvNum)).Vital(2)
            .lblPokeInfoDesc(8).Caption = "Atk: " & Pokemon(Quest(QuestNum).PokeRew(InvNum)).Add_Stat(1)
            .lblPokeInfoDesc(9).Caption = "Def: " & Pokemon(Quest(QuestNum).PokeRew(InvNum)).Add_Stat(2)
            .lblPokeInfoDesc(10).Caption = "EAtq: " & Pokemon(Quest(QuestNum).PokeRew(InvNum)).Add_Stat(3)
            .lblPokeInfoDesc(11).Caption = "EDef: " & Pokemon(Quest(QuestNum).PokeRew(InvNum)).Add_Stat(5)
            .lblPokeInfoDesc(12).Caption = "Vel: " & Pokemon(Quest(QuestNum).PokeRew(InvNum)).Add_Stat(4)
            .lblPokeInfoDesc(1).Caption = "Humor: Triste"

            If Pokemon(Quest(QuestNum).PokeRew(InvNum)).ControlSex = 100 Then
                .lblPokeInfoDesc(2).Caption = "Sexo: Masculino"
            ElseIf Pokemon(Quest(QuestNum).PokeRew(InvNum)).ControlSex = 0 Then
                .lblPokeInfoDesc(2).Caption = "Sexo: Feminino"
            Else
                FemQntia = Pokemon(Quest(QuestNum).PokeRew(InvNum)).ControlSex - 100
                If FemQntia < 0 Then FemQntia = FemQntia * -1
                .lblPokeInfoDesc(2).Caption = "Sexo: M(" & Pokemon(Quest(QuestNum).PokeRew(InvNum)).ControlSex & "%)" & " F(" & FemQntia & "%)"
            End If

            For i = 1 To 4
                If Pokemon(Quest(QuestNum).PokeRew(InvNum)).Habilidades(i).Spell > 0 Then
                    .lblPokeInfoDesc(12 + i).Caption = Trim$(Spell(Pokemon(Quest(QuestNum).PokeRew(InvNum)).Habilidades(i).Spell).Name)
                Else
                    .lblPokeInfoDesc(12 + i).Caption = "???"
                End If
            Next

        End If

        If Command = 5 Then

            If InvNum = 0 Then Exit Sub

            .lblPokeInfoDesc(0).Caption = Trim$(Pokemon(InvNum).Name) & " - Level: 5"

            .lblPokeInfoDesc(3).Caption = "Exp: 0/" & GetInvPokeNextLevel(5, 4, 5)
            .lblPokeInfoDesc(4).Caption = "Hp: " & Pokemon(InvNum).Vital(1) + 4 * 5
            .lblPokeInfoDesc(5).Caption = "PP:" & Pokemon(InvNum).Vital(2) + 4 * 5
            .lblPokeInfoDesc(8).Caption = "Atk: " & Pokemon(InvNum).Add_Stat(1) + 4 * 3
            .lblPokeInfoDesc(9).Caption = "Def: " & Pokemon(InvNum).Add_Stat(2) + 4 * 3
            .lblPokeInfoDesc(10).Caption = "EAtq: " & Pokemon(InvNum).Add_Stat(3) + 4 * 3
            .lblPokeInfoDesc(11).Caption = "EDef: " & Pokemon(InvNum).Add_Stat(5) + 4 * 3
            .lblPokeInfoDesc(12).Caption = "Vel: " & Pokemon(InvNum).Add_Stat(4) + 4 * 3
            .lblPokeInfoDesc(1).Caption = "Humor: Triste"

            If Pokemon(InvNum).ControlSex = 100 Then
                .lblPokeInfoDesc(2).Caption = "Sexo: Masculino"
            ElseIf Pokemon(InvNum).ControlSex = 0 Then
                .lblPokeInfoDesc(2).Caption = "Sexo: Feminino"
            Else
                FemQntia = Pokemon(InvNum).ControlSex - 100
                If FemQntia < 0 Then FemQntia = FemQntia * -1
                .lblPokeInfoDesc(2).Caption = "Sexo: M(" & Pokemon(InvNum).ControlSex & "%)" & " F(" & FemQntia & "%)"
            End If

            For i = 1 To 4
                If Pokemon(InvNum).Habilidades(i).Spell > 0 Then
                    .lblPokeInfoDesc(12 + i).Caption = Trim$(Spell(Pokemon(InvNum).Habilidades(i).Spell).Name)
                Else
                    .lblPokeInfoDesc(12 + i).Caption = "???"
                End If
            Next

        End If

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateDescWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetInvPokeNextLevel(ByVal InvNum As Long, ByVal Command As Byte, Optional ByVal IdPokemon As Integer) As Long
    Dim CurrentLevel As Long
    Dim PokeExpType As Byte

    Select Case Command
    Case 0
        CurrentLevel = GetPlayerInvItemPokeInfoLevel(MyIndex, InvNum)
        PokeExpType = Pokemon(GetPlayerInvItemPokeInfoPokemon(MyIndex, InvNum)).ExpType
    Case 1
        CurrentLevel = GetPlayerBankItemLevel(InvNum)
        PokeExpType = Pokemon(GetPlayerBankItemPokemon(InvNum)).ExpType
    Case 2
        CurrentLevel = Leilao(InvNum).Poke.Level
        PokeExpType = Pokemon(Leilao(InvNum).Poke.Pokemon).ExpType
    Case 3    'Trade Their
        CurrentLevel = TradeTheirOffer(InvNum).PokeInfo.Level
        PokeExpType = Pokemon(TradeTheirOffer(InvNum).PokeInfo.Pokemon).ExpType
    Case 4
        CurrentLevel = InvNum
        PokeExpType = Pokemon(IdPokemon).ExpType
    End Select

    If CurrentLevel <= 1 Then CurrentLevel = 2

    Select Case PokeExpType
    Case 0    'Rápido
        GetInvPokeNextLevel = 0.8 * (CurrentLevel) ^ 3
    Case 1    'Médio Rápido
        GetInvPokeNextLevel = (CurrentLevel) ^ 3
    Case 2    'Médio Lento
        GetInvPokeNextLevel = (CurrentLevel) ^ 3 + ((((CurrentLevel) ^ 3) * 5.9) / 100)
    Case 3    'Lento
        GetInvPokeNextLevel = 1.25 * (CurrentLevel) ^ 3
    Case Else    'Rápido
        GetInvPokeNextLevel = 0.8 * (CurrentLevel) ^ 3
    End Select

End Function

Sub MeAnimation(ByVal Animation As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal LockType As Long, Optional ByVal lockindex As Long)

    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1

    If LockType > 0 Then
        With AnimInstance(AnimationIndex)
            .Animation = Animation
            .X = X
            .Y = Y
            .LockType = LockType
            .lockindex = lockindex
            .Used(0) = True
            .Used(1) = True
        End With

        ' play the sound if we've got one
        PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

        Exit Sub
    End If

    With AnimInstance(AnimationIndex)
        .Animation = Animation
        .X = X
        .Y = Y
        .LockType = 0
        .lockindex = 0
        .Used(0) = True
        .Used(1) = True
    End With

    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

End Sub

'#################################Poke Info#####################################

Function GetPlayerBankItemPokemon(ByVal bankslot As Long) As Long
    GetPlayerBankItemPokemon = Bank.Item(bankslot).PokeInfo.Pokemon
End Function

Sub SetPlayerBankItemPokemon(ByVal bankslot As Long, ByVal PokemonNum As Long)
    Bank.Item(bankslot).PokeInfo.Pokemon = PokemonNum
End Sub

Function GetPlayerBankItemPokeball(ByVal bankslot As Long) As Long
    GetPlayerBankItemPokeball = Bank.Item(bankslot).PokeInfo.Pokeball
End Function

Sub SetPlayerBankItemPokeball(ByVal bankslot As Long, ByVal PokeballNum As Long)
    Bank.Item(bankslot).PokeInfo.Pokeball = PokeballNum
End Sub

Function GetPlayerBankItemLevel(ByVal bankslot As Long) As Long
    GetPlayerBankItemLevel = Bank.Item(bankslot).PokeInfo.Level
End Function

Sub SetPlayerBankItemLevel(ByVal bankslot As Long, ByVal LevelValue As Long)
    Bank.Item(bankslot).PokeInfo.Level = LevelValue
End Sub

Function GetPlayerBankItemExp(ByVal bankslot As Long) As Long
    GetPlayerBankItemExp = Bank.Item(bankslot).PokeInfo.Exp
End Function

Sub SetPlayerBankItemExp(ByVal bankslot As Long, ByVal ExpValue As Long)
    Bank.Item(bankslot).PokeInfo.Exp = ExpValue
End Sub

Function GetPlayerBankItemVital(ByVal bankslot As Long, ByVal VitalType As Long) As Long
    GetPlayerBankItemVital = Bank.Item(bankslot).PokeInfo.Vital(VitalType)
End Function

Sub SetPlayerBankItemVital(ByVal bankslot As Long, ByVal VitalValue As Long, ByVal VitalType As Long)
    Bank.Item(bankslot).PokeInfo.Vital(VitalType) = VitalValue
End Sub

Function GetPlayerBankItemMaxVital(ByVal bankslot As Long, ByVal VitalType As Long) As Long
    GetPlayerBankItemMaxVital = Bank.Item(bankslot).PokeInfo.MaxVital(VitalType)
End Function

Sub SetPlayerBankItemMaxVital(ByVal bankslot As Long, ByVal VitalValue As Long, ByVal VitalType As Long)
    Bank.Item(bankslot).PokeInfo.MaxVital(VitalType) = VitalValue
End Sub

Function GetPlayerBankItemStat(ByVal bankslot As Long, ByVal StatNum As Long) As Long
    GetPlayerBankItemStat = Bank.Item(bankslot).PokeInfo.Stat(StatNum)
End Function

Sub SetPlayerBankItemStat(ByVal bankslot As Long, ByVal VitalValue As Long, ByVal StatNum As Long)
    Bank.Item(bankslot).PokeInfo.Stat(StatNum) = VitalValue
End Sub

Function GetPlayerBankItemSpell(ByVal bankslot As Long, ByVal Spell As Long) As Long
    GetPlayerBankItemSpell = Bank.Item(bankslot).PokeInfo.Spells(Spell)
End Function

Sub SetPlayerBankItemSpell(ByVal bankslot As Long, ByVal SpellValue As Long, ByVal Spell As Long)
    Bank.Item(bankslot).PokeInfo.Spells(Spell) = SpellValue
End Sub

Function GetPlayerBankItemNgt(ByVal Index As Long, ByVal bankslot As Long, ByVal Ngt As Long) As Long
    GetPlayerBankItemNgt = Bank.Item(bankslot).PokeInfo.Negatives(Ngt)
End Function

Sub SetPlayerBankItemNgt(ByVal Index As Long, ByVal bankslot As Long, ByVal NgtValue As Long, ByVal Ngt As Long)
    Bank.Item(bankslot).PokeInfo.Negatives(Ngt) = NgtValue
End Sub

Function GetPlayerBankItemFelicidade(ByVal Index As Long, ByVal bankslot As Long) As Long
    GetPlayerBankItemFelicidade = Bank.Item(bankslot).PokeInfo.Felicidade
End Function

Sub SetPlayerBankItemFelicidade(ByVal Index As Long, ByVal bankslot As Long, ByVal FelicidadeValue As Long)
    Bank.Item(bankslot).PokeInfo.Felicidade = FelicidadeValue
End Sub

Function GetPlayerBankItemSexo(ByVal Index As Long, ByVal bankslot As Long) As Long
    GetPlayerBankItemSexo = Bank.Item(bankslot).PokeInfo.Sexo
End Function

Sub SetPlayerBankItemSexo(ByVal Index As Long, ByVal bankslot As Long, ByVal SexoValue As Long)
    Bank.Item(bankslot).PokeInfo.Sexo = SexoValue
End Sub

Function GetPlayerBankItemShiny(ByVal Index As Long, ByVal bankslot As Long) As Long
    GetPlayerBankItemShiny = Bank.Item(bankslot).PokeInfo.Shiny
End Function

Sub SetPlayerBankItemShiny(ByVal Index As Long, ByVal bankslot As Long, ByVal ShinyValue As Long)
    Bank.Item(bankslot).PokeInfo.Shiny = ShinyValue
End Sub

Function GetPlayerBankItemBerry(ByVal bankslot As Long, ByVal BerryStat As Long) As Long
    GetPlayerBankItemBerry = Bank.Item(bankslot).PokeInfo.Berry(BerryStat)
End Function

Sub SetPlayerBankItemBerry(ByVal bankslot As Long, ByVal BerryStat As Long, ByVal Valor As Long)
    Bank.Item(bankslot).PokeInfo.Stat(BerryStat) = Valor
End Sub

'###############################################################################

Public Function GetStat(ByVal Stat As Stats) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return with Stat name
    Select Case Stat
    Case Strength
        GetStat = "Strength"
    Case Endurance
        GetStat = "Endurance"
    Case Intelligence
        GetStat = "Intelligence"
    Case Agility
        GetStat = "Agility"
    Case Willpower
        GetStat = "Willpower"
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetStat", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetStatAbb(ByVal Stat As Stats) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return with Stat name
    Select Case Stat
    Case Strength
        GetStatAbb = "Str"
    Case Endurance
        GetStatAbb = "End"
    Case Intelligence
        GetStatAbb = "Int"
    Case Agility
        GetStatAbb = "Agi"
    Case Willpower
        GetStatAbb = "Will"
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetStat", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetVital(ByVal Vital As Vitals) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return with vital name
    Select Case Vital
    Case HP
        GetVital = "HP"
    Case MP
        GetVital = "MP"
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVital", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function SetCaption(ByVal Declaration As String, ByVal text As String) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SetCaption = Declaration & text & vbNewLine

    ' Error handler
    Exit Function
errorhandler:
    HandleError "SetCaption", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function SetCaptionBut(ByVal Declaration As String, ByVal text As String, ByVal value As Long) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If value <= 0 Then Exit Function
    SetCaptionBut = Declaration & text & value & vbNewLine

    ' Error handler
    Exit Function
errorhandler:
    HandleError "SetCaptionBut", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetYesNo(ByVal value As Long) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If value Then
        GetYesNo = "Yes"
    Else
        GetYesNo = "No"
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "SetCaption", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SendALeilao()
    Dim i As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With frmLeilao

        .cmbBolsa.Clear
        .cmbBolsa.AddItem "None"
        .cmbBolsa.ListIndex = 0

        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                If GetPlayerInvItemPokeInfoPokemon(MyIndex, i) > 0 Then
                    .cmbBolsa.AddItem i & "°: " & Trim$(Pokemon(GetPlayerInvItemPokeInfoPokemon(MyIndex, i)).Name) & "(" & GetPlayerInvItemPokeInfoLevel(MyIndex, i) & ")"
                Else
                    If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                        .cmbBolsa.AddItem i & "°: " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & "(" & GetPlayerInvItemValue(MyIndex, i) & ")"
                    Else
                        .cmbBolsa.AddItem i & "°: " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                    End If
                End If
            Else
                .cmbBolsa.AddItem i & "°: " & "Slot vazio!"
            End If
        Next

        .lstLeilao.Clear

        For i = 1 To MAX_LEILAO
            If Leilao(i).itemNum > 0 Then

                If Leilao(i).Poke.Pokemon > 0 Then
                    Select Case Leilao(i).Tipo
                    Case 1
                        .lstLeilao.AddItem i & "°: " & Trim$(Pokemon(Leilao(i).Poke.Pokemon).Name) & "[" & Leilao(i).Poke.Level & "]" & " Preço: " & Leilao(i).Price & " Zenys"
                    Case 2
                        .lstLeilao.AddItem i & "°: " & Trim$(Pokemon(Leilao(i).Poke.Pokemon).Name) & "[" & Leilao(i).Poke.Level & "]" & " Preço: " & Leilao(i).Price & " PokeCredits"
                    Case Else
                    End Select

                Else
                    Select Case Leilao(i).Tipo
                    Case 1
                        .lstLeilao.AddItem i & "°: " & Trim$(Item(Leilao(i).itemNum).Name) & " Preço: " & Leilao(i).Price & " Zenys"
                    Case 2
                        .lstLeilao.AddItem i & "°: " & Trim$(Item(Leilao(i).itemNum).Name) & " Preço: " & Leilao(i).Price & " PokeCredits"
                    Case Else
                    End Select

                End If
            Else
                .lstLeilao.AddItem i & "°: Não contém item!"
            End If
        Next

        .lstMyLeiloes.Clear
        .lstMyLeiloes.AddItem "None"

        For i = 1 To MAX_LEILAO

            If Leilao(i).Poke.Pokemon > 0 Then

                If Leilao(i).itemNum > 0 Then
                    If Leilao(i).Vendedor = GetPlayerName(MyIndex) Then
                        Select Case Leilao(i).Tipo
                        Case 1
                            Player(MyIndex).MyLeiloes(.lstMyLeiloes.ListCount) = i
                            .lstMyLeiloes.AddItem i & "°: " & Trim$(Pokemon(Leilao(i).Poke.Pokemon).Name) & "[" & Leilao(i).Poke.Level & "] - " & Leilao(i).Price & "Zenys"
                        Case 2
                            Player(MyIndex).MyLeiloes(.lstMyLeiloes.ListCount) = i
                            .lstMyLeiloes.AddItem i & "°: " & Trim$(Pokemon(Leilao(i).Poke.Pokemon).Name) & "[" & Leilao(i).Poke.Level & "] - " & Leilao(i).Price & "PokeCredits"
                        End Select
                    End If
                End If

            Else
                If Leilao(i).itemNum > 0 Then
                    If Leilao(i).Vendedor = GetPlayerName(MyIndex) Then
                        Player(MyIndex).MyLeiloes(.lstMyLeiloes.ListCount) = i
                        .lstMyLeiloes.AddItem i & "°: " & Trim$(Item(Leilao(i).itemNum).Name) & " Preço: " & Leilao(i).Price
                    End If
                End If
            End If

        Next

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendALeilao", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CarregarInsignia()
    Dim i As Long

    For i = 1 To 8    'MAX_INSIGNIAS
        If Player(MyIndex).Insignia(i) = 1 Then
            frmMain.PicInsignia(i).Visible = True
        Else
            frmMain.PicInsignia(i).Visible = False
        End If
    Next

End Sub

Public Sub UpdateRankLevel()
    Dim i As Byte

    frmMain.lstRank.Clear


    For i = 1 To MAX_RANKS
        If Len(Trim$(RankLevel(i).Name)) > 0 Then
            frmMain.lstRank.AddItem i & ": " & RankLevel(i).Name & " " & Trim$(Pokemon(RankLevel(i).PokeNum).Name) & " " & " Lvl: " & RankLevel(i).Level
            frmMain.lblRank(i).Caption = RankLevel(i).Name
            frmMain.lblP(i).Caption = Trim$(Pokemon(RankLevel(i).PokeNum).Name)
            frmMain.lblL(i).Caption = RankLevel(i).Level
        End If
    Next
End Sub

Public Function CheckResourceStatCut(ByVal Index As Long, ByVal X As Long, ByVal Y As Long) As Boolean

    Dim i As Long, Resource_num As Long

    CheckResourceStatCut = False

    If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0

        ' Get the cache number
        For i = 0 To Resource_Index

            If MapResource(i).X = X Then
                If MapResource(i).Y = Y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num = 0 Then Exit Function

        If MapResource(Resource_num).ResourceState = 1 Then
            CheckResourceStatCut = True
        End If
    End If

End Function

Public Function IsValidEmail(strEmail As String) As Boolean
    Dim names, Name, i, c
    IsValidEmail = True

    names = Split(strEmail, "@")

    If UBound(names) <> 1 Then
        IsValidEmail = False
        Exit Function
    End If

    For Each Name In names

        If Len(Name) <= 0 Then
            IsValidEmail = False
            Exit Function
        End If

        For i = 1 To Len(Name)
            c = LCase(Mid(Name, i, 1))

            If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
                IsValidEmail = False
                Exit Function
            End If
        Next

        If Left(Name, 1) = "." Or Right(Name, 1) = "." Then
            IsValidEmail = False
            Exit Function
        End If

    Next

    If InStr(names(1), ".") <= 0 Then
        IsValidEmail = False
        Exit Function
    End If

    i = Len(names(1)) - InStrRev(names(1), ".")

    If i <> 2 And i <> 3 Then
        IsValidEmail = False
        Exit Function
    End If

    If InStr(strEmail, "..") > 0 Then
        IsValidEmail = False
        Exit Function
    End If

End Function

Public Function GetOrgItemNum(ByVal OrgShopSlot As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If OrgShopSlot = 0 Then
        GetOrgItemNum = 0
        Exit Function
    End If

    If OrgShopSlot > MAX_ORG_SHOP Then
        GetOrgItemNum = 0
        Exit Function
    End If

    GetOrgItemNum = OrgShop(OrgShopSlot).Item

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetOrgItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub ScrollBarMembers()
    Dim i As Long, OrgP As Byte
    Dim Quantia As Byte, Height As Integer

    'Pagina OrgMembros
    OrgPagMem = 1

    'Sair caso não tenha Org
    If GetPlayerOrg(MyIndex) = 0 Then Exit Sub
    OrgP = GetPlayerOrg(MyIndex)

    'Quantia
    For i = 1 To MAX_ORG_MEMBERS
        If Organization(OrgP).OrgMember(i).Used = True Then
            QntOrgPag = QntOrgPag + 1
        End If
    Next

    Select Case QntOrgPag
    Case 1 To 9
        Height = 102
    Case 10 To 18
        Height = 102
    Case 19 To 27
        Height = 102
    Case 28 To 36
        Height = 102
    End Select

    'Cordenadas
    frmMain.ScrollBarFake(5).Height = Height
    frmMain.ScrollBarFake(5).top = 124
    frmMain.ScrollBarFake(5).Left = 219
End Sub

' Chat Bubble Mondo
Public Sub AddChatBubble(ByVal target As Long, ByVal targetType As Byte, ByVal Msg As String, ByVal colour As Long)
    Dim i As Long, Index As Long

    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1

    ' default to new bubble
    Index = chatBubbleIndex

    ' loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If chatBubble(i).targetType = targetType Then
            If chatBubble(i).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = i
                Exit For
            End If
        End If
    Next

    ' set the bubble up
    With chatBubble(Index)
        .target = target
        .targetType = targetType
        .Msg = Msg
        .colour = colour
        .timer = GetTickCount
        .active = True
    End With
End Sub
