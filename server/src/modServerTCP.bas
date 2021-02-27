Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next

End Sub

Function IsConnected(ByVal Index As Long) As Boolean

    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean

    If IsConnected(Index) Then
        If LenB(Trim$(Player(Index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsBanned(ByVal IP As String) As Boolean
    Dim filename As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long
    filename = App.Path & "\data\banlist.txt"

    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

Function IsInventoryFull(ByVal tradeTarget As Long, ByVal Index As Long) As Boolean
Dim InvEmpty As Long, TradeFull As Long, i As Long

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(tradeTarget, i) > 0 And GetPlayerInvItemNum(tradeTarget, i) <= MAX_ITEMS Then
            InvEmpty = InvEmpty + 1
        End If
    Next
        
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num > 0 And TempPlayer(Index).TradeOffer(i).Num <= MAX_ITEMS Then
            TradeFull = TradeFull + 1
        End If
    Next
        
    If TradeFull > (MAX_INV - InvEmpty) Then
        IsInventoryFull = True
        Exit Function
    End If
    
    IsInventoryFull = False

End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(Index).SendData Buffer.ToArray()
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> Index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next

End Sub

Sub SendDataToParty(ByVal partynum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To Party(partynum).MemberCount
        If Party(partynum).Member(i) > 0 Then
            Call SendDataTo(Party(partynum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, Buffer.ToArray
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToMap MapNum, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo Index, Buffer.ToArray
    DoEvents
    Call CloseSocket(Index)
    
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partynum As Long, ByVal Msg As String, ByVal color As Byte)
Dim i As Long
    ' send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partynum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partynum).Member(i)) And IsPlaying(Party(partynum).Member(i)) Then
                PlayerMsg Party(partynum).Member(i), Msg, color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)

    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(Index, "You have lost your connection with " & Options.Game_Name & ".")
    End If

End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (Index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal Index As Long)
Dim i As Long

    If Index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(Index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".")
        Else
            Call AlertMsg(Index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
    End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

 If GetPlayerAccess(Index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then
            If GetTickCount < TempPlayer(Index).DataTimer Then
                Exit Sub
            End If
        End If
    
        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 25 Then
            If GetTickCount < TempPlayer(Index).DataTimer Then
                Exit Sub
            End If
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = GetTickCount + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(Index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(Index).Buffer.Length >= 4 Then
        pLength = TempPlayer(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.Length - 4
        If pLength <= TempPlayer(Index).Buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).Buffer.ReadLong
            HandleData Index, TempPlayer(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(Index).Buffer.Length >= 4 Then
            pLength = TempPlayer(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(Index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal Index As Long)

    If Index > 0 Then
        Call LeftGame(Index)
        Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        frmServer.Socket(Index).Close
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If

End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong MapNum
    Buffer.WriteString Trim$(Map(MapNum).Name)
    Buffer.WriteString Trim$(Map(MapNum).Music)
    Buffer.WriteLong Map(MapNum).Revision
    Buffer.WriteByte Map(MapNum).Moral
    Buffer.WriteLong Map(MapNum).Up
    Buffer.WriteLong Map(MapNum).Down
    Buffer.WriteLong Map(MapNum).Left
    Buffer.WriteLong Map(MapNum).Right
    Buffer.WriteLong Map(MapNum).BootMap
    Buffer.WriteByte Map(MapNum).BootX
    Buffer.WriteByte Map(MapNum).BootY
    Buffer.WriteByte Map(MapNum).MaxX
    Buffer.WriteByte Map(MapNum).MaxY
    Buffer.WriteLong Map(MapNum).Weather
    Buffer.WriteLong Map(MapNum).Intensity
    
    For X = 1 To 2
        Buffer.WriteLong Map(MapNum).LevelPoke(X)
    Next

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            With Map(MapNum).Tile(X, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).X
                    Buffer.WriteLong .Layer(i).Y
                    Buffer.WriteLong .Layer(i).Tileset
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(MapNum).Npc(X)
    Next

    MapCache(MapNum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal Index As Long)
    Dim S As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                S = S & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        S = "There are no other players online."
    Else
        S = Mid$(S, 1, Len(S) - 2)
        S = "There are " & n & " other players online: " & S & "."
    End If

    Call PlayerMsg(Index, S, WhoColor)
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long

    If Index > MAX_PLAYERS Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteLong GetPlayerPOINTS(Index)
    Buffer.WriteLong GetPlayerSprite(Index)
    Buffer.WriteLong GetPlayerMap(Index)
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteLong GetPlayerFlying(Index)
    Buffer.WriteLong Player(Index).TPX
    Buffer.WriteLong Player(Index).TPY
    Buffer.WriteLong Player(Index).TPDir
    Buffer.WriteLong Player(Index).MySprite
    Buffer.WriteLong Player(Index).Vitorias
    Buffer.WriteLong Player(Index).Derrotas
    Buffer.WriteByte Player(Index).ORG
    Buffer.WriteLong Player(Index).Honra
    Buffer.WriteByte Player(Index).MyVip
    
    If Player(Index).VipInName = True Then
        Buffer.WriteByte 1
    Else
        Buffer.WriteByte 0
    End If
    
    If Player(Index).PokeLight = True Then
        Buffer.WriteByte 1
    Else
        Buffer.WriteByte 0
    End If
    
    For i = 1 To MAX_INSIGNIAS
        Buffer.WriteLong Player(Index).Insignia(i)
    Next
    
    For i = 1 To MAX_QUESTS
        Buffer.WriteByte Player(Index).Quests(i).Status
        Buffer.WriteByte Player(Index).Quests(i).Part
    Next
    
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    SendDataTo Index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
    
    Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong Index
    SendDataToMapBut Index, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).X
        Buffer.WriteLong MapItem(MapNum, i).Y
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).X
        Buffer.WriteLong MapItem(MapNum, i).Y
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal MapNum As Long, ByVal MapNpcNum As Long, Optional ByVal ToIndex As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteLong MapNpcNum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Vital(i)
    Next

    If ToIndex > 0 Then
        SendDataTo ToIndex, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).X
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
        
        If MapNpc(MapNum).Npc(i).Sexo = False Then
            Buffer.WriteLong 0
        Else
            Buffer.WriteLong 1
        End If
        
        If MapNpc(MapNum).Npc(i).Shiny = False Then
            Buffer.WriteLong 0
        Else
            Buffer.WriteLong 1
        End If
        
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Level
        
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).X
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
        
        If MapNpc(MapNum).Npc(i).Sexo = False Then
            Buffer.WriteLong 0
        Else
            Buffer.WriteLong 1
        End If
        
        If MapNpc(MapNum).Npc(i).Shiny = False Then
            Buffer.WriteLong 0
        Else
            Buffer.WriteLong 1
        End If
        
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Level
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If

    Next

End Sub

Sub SendAnimations(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(Index, i)
        End If

    Next

End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If

    Next

End Sub

Sub SendResources(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(Index, i)
        End If

    Next

End Sub

Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, i)
        Buffer.WriteLong GetPlayerInvItemValue(Index, i)
        Buffer.WriteLong GetPlayerInvItemPokeInfoPokemon(Index, i)
        Buffer.WriteLong GetPlayerInvItemPokeInfoPokeball(Index, i)
        Buffer.WriteLong GetPlayerInvItemPokeInfoLevel(Index, i)
        Buffer.WriteLong GetPlayerInvItemPokeInfoExp(Index, i)
        
        For X = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong GetPlayerInvItemPokeInfoVital(Index, i, X)
            Buffer.WriteLong GetPlayerInvItemPokeInfoMaxVital(Index, i, X)
        Next
        
        For X = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong GetPlayerInvItemPokeInfoStat(Index, i, X)
        Next
        
        For X = 1 To MAX_POKE_SPELL
            Buffer.WriteLong GetPlayerInvItemPokeInfoSpell(Index, i, X)
        Next
        
        For X = 1 To MAX_NEGATIVES
            Buffer.WriteLong GetPlayerInvItemNgt(Index, i, X)
        Next
        
        For X = 1 To MAX_BERRYS
            Buffer.WriteLong GetPlayerInvItemBerry(Index, i, X)
        Next
        
        Buffer.WriteLong GetPlayerInvItemFelicidade(Index, i)
        Buffer.WriteLong GetPlayerInvItemSexo(Index, i)
        Buffer.WriteLong GetPlayerInvItemShiny(Index, i)
                
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal invslot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim X As Long
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invslot
    Buffer.WriteLong GetPlayerInvItemNum(Index, invslot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, invslot)
    Buffer.WriteLong GetPlayerInvItemPokeInfoPokemon(Index, invslot)
    Buffer.WriteLong GetPlayerInvItemPokeInfoPokeball(Index, invslot)
    Buffer.WriteLong GetPlayerInvItemPokeInfoLevel(Index, invslot)
    Buffer.WriteLong GetPlayerInvItemPokeInfoExp(Index, invslot)
    
    For X = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerInvItemPokeInfoVital(Index, invslot, X)
        Buffer.WriteLong GetPlayerInvItemPokeInfoMaxVital(Index, invslot, X)
    Next
    
    For X = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerInvItemPokeInfoStat(Index, invslot, X)
    Next
    
    For X = 1 To MAX_POKE_SPELL
        Buffer.WriteLong GetPlayerInvItemPokeInfoSpell(Index, invslot, X)
    Next
    
    For X = 1 To MAX_NEGATIVES
        Buffer.WriteLong GetPlayerInvItemNgt(Index, invslot, X)
    Next
        
    For X = 1 To MAX_BERRYS
        Buffer.WriteLong GetPlayerInvItemBerry(Index, invslot, X)
    Next
    
    Buffer.WriteLong GetPlayerInvItemFelicidade(Index, invslot)
    Buffer.WriteLong GetPlayerInvItemSexo(Index, invslot)
    Buffer.WriteLong GetPlayerInvItemShiny(Index, invslot)

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Long
    
    Buffer.WriteLong SPlayerWornEq
    
    'Armor Infos
    Buffer.WriteLong GetPlayerEquipment(Index, Armor)
    
    'Weapon Infos
    Buffer.WriteLong GetPlayerEquipment(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokemon(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokeball(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoLevel(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoExp(Index, weapon)
    
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoVital(Index, weapon, i)
        Buffer.WriteLong GetPlayerEquipmentPokeInfoMaxVital(Index, weapon, i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoStat(Index, weapon, i)
    Next
    
    For i = 1 To MAX_POKE_SPELL
        Buffer.WriteLong GetPlayerEquipmentPokeInfoSpell(Index, weapon, i)
    Next

    For i = 1 To MAX_NEGATIVES
        Buffer.WriteLong GetPlayerEquipmentNgt(Index, weapon, i)
    Next
    
    For i = 1 To MAX_BERRYS
        Buffer.WriteLong GetPlayerEquipmentBerry(Index, weapon, i)
    Next
    
    Buffer.WriteLong GetPlayerEquipmentFelicidade(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentSexo(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentShiny(Index, weapon)

    'Helmet Infos
    Buffer.WriteLong GetPlayerEquipment(Index, Helmet)

    'Shield Infos
    Buffer.WriteLong GetPlayerEquipment(Index, Shield)

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Long
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong Index
    
    'Armor Infos
    Buffer.WriteLong GetPlayerEquipment(Index, Armor)
    
    'Weapon Infos
    Buffer.WriteLong GetPlayerEquipment(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokemon(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokeball(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoLevel(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoExp(Index, weapon)
    
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoVital(Index, weapon, i)
        Buffer.WriteLong GetPlayerEquipmentPokeInfoMaxVital(Index, weapon, i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoStat(Index, weapon, i)
    Next
    
    For i = 1 To MAX_POKE_SPELL
        Buffer.WriteLong GetPlayerEquipmentPokeInfoSpell(Index, weapon, i)
    Next
    
    For i = 1 To MAX_NEGATIVES
        Buffer.WriteLong GetPlayerEquipmentNgt(Index, weapon, i)
    Next
    
    For i = 1 To MAX_BERRYS
        Buffer.WriteLong GetPlayerEquipmentBerry(Index, weapon, i)
    Next
    
    Buffer.WriteLong GetPlayerEquipmentFelicidade(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentSexo(Index, weapon)
    Buffer.WriteLong GetPlayerEquipmentShiny(Index, weapon)

    'Helmet Infos
    Buffer.WriteLong GetPlayerEquipment(Index, Helmet)

    'Shield Infos
    Buffer.WriteLong GetPlayerEquipment(Index, Shield)
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Long
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    'Armor Infos
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    
    'Weapon Infos
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokemon(PlayerNum, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokeball(PlayerNum, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoLevel(PlayerNum, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoExp(PlayerNum, weapon)
    
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoVital(PlayerNum, weapon, i)
        Buffer.WriteLong GetPlayerEquipmentPokeInfoMaxVital(PlayerNum, weapon, i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoStat(PlayerNum, weapon, i)
    Next
    
    For i = 1 To MAX_POKE_SPELL
        Buffer.WriteLong GetPlayerEquipmentPokeInfoSpell(PlayerNum, weapon, i)
    Next
    
    For i = 1 To MAX_NEGATIVES
        Buffer.WriteLong GetPlayerEquipmentNgt(PlayerNum, weapon, i)
    Next
    
    For i = 1 To MAX_BERRYS
        Buffer.WriteLong GetPlayerEquipmentBerry(PlayerNum, weapon, i)
    Next
    
    Buffer.WriteLong GetPlayerEquipmentFelicidade(PlayerNum, weapon)
    Buffer.WriteLong GetPlayerEquipmentSexo(PlayerNum, weapon)
    Buffer.WriteLong GetPlayerEquipmentShiny(PlayerNum, weapon)

    'Helmet Infos
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Helmet)

    'Shield Infos
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendEXP(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(Index)
    Buffer.WriteLong GetPlayerNextLevel(Index)
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStats(ByVal Index As Long)
'Dim i As Long
'Dim packet As String
'Dim Buffer As clsBuffer
'
'    Set Buffer = New clsBuffer
'    Buffer.WriteLong SPlayerStats
'    For i = 1 To Stats.Stat_Count - 1
'        Buffer.WriteLong GetPlayerStat(Index, i)
'    Next
'    SendDataTo Index, Buffer.ToArray()
'    Set Buffer = Nothing
End Sub

Sub SendWelcome(ByVal Index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(Index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NpcNum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes NPCData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NpcNum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes NPCData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If

    Next

End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(Index, i)
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal Index As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).Resource_Count

    If ResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).X
            Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).Y
        Next

    End If

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(MapNum).Resource_Count

    If ResourceCache(MapNum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(MapNum).Resource_Count
            Buffer.WriteByte ResourceCache(MapNum).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).X
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).Y
        Next

    End If

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SDoorAnimation
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendActionMsg(ByVal MapNum As Long, ByVal message As String, ByVal color As Long, ByVal MsgType As Long, ByVal X As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong color
    Buffer.WriteLong MsgType
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendAnimation(ByVal MapNum As Long, ByVal Anim As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCooldown(ByVal Index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Long, ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, X As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Bank(Index).Item(i).Num
        Buffer.WriteLong Bank(Index).Item(i).Value
        Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Pokemon
        Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Pokeball
        Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Level
        Buffer.WriteLong Bank(Index).Item(i).PokeInfo.EXP
        
        For X = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Vital(X)
            Buffer.WriteLong Bank(Index).Item(i).PokeInfo.MaxVital(X)
        Next
        
        For X = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Stat(X)
        Next
        
        For X = 1 To MAX_POKE_SPELL
            Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Spells(X)
        Next
        
        For X = 1 To MAX_NEGATIVES
            Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Negatives(X)
        Next
        
        For X = 1 To MAX_BERRYS
            Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Berry(X)
        Next
        
        Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Felicidade
        Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Sexo
        Buffer.WriteLong Bank(Index).Item(i).PokeInfo.Shiny
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKey(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte Value
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte Value
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal Index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendTrade(ByVal Index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal dataType As Byte)
Dim Buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
Dim X As Long
    
    tradeTarget = TempPlayer(Index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
        
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Value
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Pokemon
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Pokeball
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Level
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.EXP
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Felicidade
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Sexo
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Shiny
            
            For X = 1 To Vitals.Vital_Count - 1
                Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Vital(X)
                Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.MaxVital(X)
            Next
            
            For X = 1 To Stats.Stat_Count - 1
                Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Stat(X)
            Next
            
            For X = 1 To MAX_POKE_SPELL
                Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Spells(X)
            Next
            
            For X = 1 To MAX_BERRYS
                Buffer.WriteLong TempPlayer(Index).TradeOffer(i).PokeInfo.Berry(X)
            Next
            
            ' add total worth
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(Index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).Price * TempPlayer(Index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
        
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            Buffer.WriteLong GetPlayerInvItemPokeInfoPokemon(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong GetPlayerInvItemPokeInfoPokeball(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong GetPlayerInvItemPokeInfoLevel(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong GetPlayerInvItemPokeInfoExp(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong GetPlayerInvItemFelicidade(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong GetPlayerInvItemSexo(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong GetPlayerInvItemShiny(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            
            For X = 1 To Vitals.Vital_Count - 1
                Buffer.WriteLong GetPlayerInvItemPokeInfoVital(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
                Buffer.WriteLong GetPlayerInvItemPokeInfoMaxVital(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
            Next
            
            For X = 1 To Stats.Stat_Count - 1
                Buffer.WriteLong GetPlayerInvItemPokeInfoStat(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
            Next
            
            For X = 1 To MAX_POKE_SPELL
                Buffer.WriteLong GetPlayerInvItemPokeInfoSpell(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
            Next
            
            For X = 1 To MAX_BERRYS
                Buffer.WriteLong GetPlayerInvItemBerry(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, X)
            Next
            
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTarget(ByVal Index As Long, Optional ByVal MapNpcNum As Integer)
Dim Buffer As clsBuffer
Dim MapNum As Integer
MapNum = GetPlayerMap(Index)

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(Index).target
    Buffer.WriteLong TempPlayer(Index).targetType
    Buffer.WriteInteger MapNpcNum
    If MapNpcNum > 0 Then
        Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Vital(1)
    Else
        Buffer.WriteLong 0
    End If
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(Index).Hotbar(i).Slot
        Buffer.WriteByte Player(Index).Hotbar(i).sType
        Buffer.WriteLong Player(Index).Hotbar(i).Pokemon
        Buffer.WriteLong Player(Index).Hotbar(i).Pokeball
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLoginOk(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong Index
    Buffer.WriteLong Player_HighIndex
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHighIndex()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong Player_HighIndex
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal Index As Long, ByVal targetPlayer As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(targetPlayer).Name)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partynum As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partynum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partynum).Member(i)
    Next
    Buffer.WriteLong Party(partynum).MemberCount
    SendDataToParty partynum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long, partynum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partynum = TempPlayer(Index).inParty
    If partynum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partynum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partynum).Member(i)
        Next
        Buffer.WriteLong Party(partynum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partynum As Long, ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong Index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(Index, i)
        Buffer.WriteLong Player(Index).Vital(i)
    Next
    SendDataToParty partynum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal MapNum As Long, ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong Index
    Buffer.WriteLong MapItem(MapNum, Index).Num
    Buffer.WriteLong MapItem(MapNum, Index).Value
    Buffer.WriteLong MapItem(MapNum, Index).X
    Buffer.WriteLong MapItem(MapNum, Index).Y
    
    Buffer.WriteLong MapItem(MapNum, Index).PokeInfo.Pokemon
    Buffer.WriteLong MapItem(MapNum, Index).PokeInfo.Pokeball
    Buffer.WriteLong MapItem(MapNum, Index).PokeInfo.Level
    Buffer.WriteLong MapItem(MapNum, Index).PokeInfo.EXP
    
    For i = 1 To Vitals.Vital_Count - 1
    Buffer.WriteLong MapItem(MapNum, Index).PokeInfo.Vital(i)
    Buffer.WriteLong MapItem(MapNum, Index).PokeInfo.MaxVital(i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
    Buffer.WriteLong MapItem(MapNum, Index).PokeInfo.Stat(i)
    Next
    
    For i = 1 To MAX_POKE_SPELL
    Buffer.WriteLong MapItem(MapNum, Index).PokeInfo.Spells(i)
    Next
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAttack(ByVal Index As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong ServerPackets.SAttack
Buffer.WriteLong Index

SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
Set Buffer = Nothing
End Sub

Sub SendNpcDesmaiado(ByVal MapNum, ByVal MapNpcNum, ByVal Sumir As Boolean, Optional ByVal ToIndex As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        If Sumir = True Then
        Buffer.WriteByte 1
        Else
        Buffer.WriteByte 0
        End If
        Buffer.WriteLong MapNpcNum
        
        If MapNpc(MapNum).Npc(MapNpcNum).Desmaiado = True Then
        Buffer.WriteLong 1
        Else
        Buffer.WriteLong 0
        End If
        
        If ToIndex > 0 Then
            SendDataTo ToIndex, Buffer.ToArray()
        Else
            SendDataToMap MapNum, Buffer.ToArray()
        End If
        Set Buffer = Nothing
End Sub

Sub SendPokeEvolution(ByVal Index As Long, ByVal Command As Byte)
Dim Buffer As clsBuffer
    
    If Player(Index).EvolPermition = 0 Then Exit Sub
    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPokeEvo
    Buffer.WriteByte Command
    Buffer.WriteByte Player(Index).EvolPermition
    Buffer.WriteInteger Player(Index).EvoId
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInFishing(ByVal Index As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong ServerPackets.SFishing
Buffer.WriteLong Index

If Player(Index).InFishing > 0 Then
Buffer.WriteByte 1
Else
Buffer.WriteByte 0
End If

If TempPlayer(Index).ScanTime > 0 Then
Buffer.WriteByte 1
Else
Buffer.WriteByte 0
End If

SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
Set Buffer = Nothing
End Sub

'#####################
'###### QUESTS #######
'#####################

Sub SendQuests(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(Index, i)
        End If
    Next
End Sub

Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    
    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong ServerPackets.SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong ServerPackets.SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendQuestCommand(ByVal Index As Integer, ByVal Command As Byte, Optional ByVal Value As Long, Optional ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim MapName As String

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SQuestCommand
    Buffer.WriteByte Command
    Buffer.WriteLong Value
    
    If Command = 2 Then
        Buffer.WriteLong QuestNum
        Buffer.WriteLong Player(Index).Quests(QuestNum).KillNpcs
        Buffer.WriteLong Player(Index).Quests(QuestNum).KillPlayers
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDialogue(ByVal Index As Integer, ByVal Title As String, ByVal Text As String, ByVal dType As Byte, ByVal YesNo As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SDialogue
    Buffer.WriteString Title
    Buffer.WriteString Text
    Buffer.WriteByte dType
    Buffer.WriteLong YesNo
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAttLeilao()
Dim Buffer As clsBuffer
Dim i As Long, X As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLeiloar
    
    For i = 1 To MAX_LEILAO
        Buffer.WriteString Leilao(i).Vendedor
        Buffer.WriteLong Leilao(i).ItemNum
        Buffer.WriteLong Leilao(i).Price
        Buffer.WriteLong Leilao(i).Tempo
        Buffer.WriteLong Leilao(i).Tipo
        'Pokemon
        Buffer.WriteLong Leilao(i).Poke.Pokemon
        Buffer.WriteLong Leilao(i).Poke.Pokeball
        Buffer.WriteLong Leilao(i).Poke.Level
        Buffer.WriteLong Leilao(i).Poke.EXP
        Buffer.WriteLong Leilao(i).Poke.Felicidade
        Buffer.WriteLong Leilao(i).Poke.Sexo
        Buffer.WriteLong Leilao(i).Poke.Shiny
        
        For X = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong Leilao(i).Poke.Vital(X)
            Buffer.WriteLong Leilao(i).Poke.MaxVital(X)
        Next
        
        For X = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Leilao(i).Poke.Stat(X)
        Next
        
        For X = 1 To MAX_POKE_SPELL
            Buffer.WriteLong Leilao(i).Poke.Spells(X)
        Next
        
        For X = 1 To MAX_NEGATIVES
            Buffer.WriteLong Leilao(i).Poke.Negatives(X)
        Next
        
        For X = 1 To MAX_BERRYS
            Buffer.WriteLong Leilao(i).Poke.Berry(X)
        Next
    Next
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAttLeilaoTo(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long, X As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLeiloar
    
    For i = 1 To MAX_LEILAO
        Buffer.WriteString Leilao(i).Vendedor
        Buffer.WriteLong Leilao(i).ItemNum
        Buffer.WriteLong Leilao(i).Price
        Buffer.WriteLong Leilao(i).Tempo
        Buffer.WriteLong Leilao(i).Tipo
        'Pokemon
        Buffer.WriteLong Leilao(i).Poke.Pokemon
        Buffer.WriteLong Leilao(i).Poke.Pokeball
        Buffer.WriteLong Leilao(i).Poke.Level
        Buffer.WriteLong Leilao(i).Poke.EXP
        Buffer.WriteLong Leilao(i).Poke.Felicidade
        Buffer.WriteLong Leilao(i).Poke.Sexo
        Buffer.WriteLong Leilao(i).Poke.Shiny
        
        For X = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong Leilao(i).Poke.Vital(X)
            Buffer.WriteLong Leilao(i).Poke.MaxVital(X)
        Next
        
        For X = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Leilao(i).Poke.Stat(X)
        Next
        
        For X = 1 To MAX_POKE_SPELL
            Buffer.WriteLong Leilao(i).Poke.Spells(X)
        Next
        
        For X = 1 To MAX_NEGATIVES
            Buffer.WriteLong Leilao(i).Poke.Negatives(X)
        Next
        
        For X = 1 To MAX_BERRYS
            Buffer.WriteLong Leilao(i).Poke.Berry(X)
        Next
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function FindLeilao() As Long
Dim i As Long

FindLeilao = 0

    For i = 1 To MAX_LEILAO
        If Leilao(i).Vendedor = vbNullString Then
            FindLeilao = i
            Exit Function
        End If
    Next
End Function

Function FindPend() As Long
Dim i As Long

FindPend = 0

    For i = 1 To MAX_LEILAO
        If Pendencia(i).Vendedor = vbNullString Then
            FindPend = i
            Exit Function
        End If
    Next
End Function

Sub SendChat(ByVal C As Long, ByVal Index As Long, ByVal Mandou As Long, ByVal T As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCChat
    Buffer.WriteLong C
    Buffer.WriteLong Mandou
    Buffer.WriteString T
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendEscolherPokeInicial(ByVal Index As Long)
Dim Buffer As clsBuffer

If Player(Index).PokeInicial = 0 Then Exit Sub
                                
Set Buffer = New clsBuffer
Buffer.WriteLong SPokeSelect
SendDataTo Index, Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendSurfInit(ByVal Index As Long, Optional ByVal ToTarget As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong SSurfInit
Buffer.WriteLong Index
Buffer.WriteByte Player(Index).InSurf

If ToTarget > 0 Then
    SendDataTo ToTarget, Buffer.ToArray()
Else
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
End If

Set Buffer = Nothing

End Sub

Sub SendRankLevel()
Dim Buffer As clsBuffer
Dim i As Long

Set Buffer = New clsBuffer
Buffer.WriteLong SUpdateRankLevel

For i = 1 To MAX_RANKS
    Buffer.WriteString RankLevel(i).Name
    Buffer.WriteLong RankLevel(i).Level
    Buffer.WriteLong RankLevel(i).PokeNum
Next

SendDataToAll Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendRankLevelTo(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long

Set Buffer = New clsBuffer
Buffer.WriteLong SUpdateRankLevel

For i = 1 To MAX_RANKS
    Buffer.WriteString RankLevel(i).Name
    Buffer.WriteLong RankLevel(i).Level
    Buffer.WriteLong RankLevel(i).PokeNum
Next

SendDataTo Index, Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendAprenderSpell(ByVal Index As Long, ByVal Command As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer

    Buffer.WriteLong SAprender
    Buffer.WriteByte Command
    Buffer.WriteInteger Player(Index).LearnSpell(1)
    Buffer.WriteInteger Player(Index).LearnSpell(2)
    
    SendDataTo Index, Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendLutarComando(C, T, Index, CC, A, Pok)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCLutar
    Buffer.WriteLong C
    Buffer.WriteLong T
    Buffer.WriteLong CC
    Buffer.WriteLong A
    Buffer.WriteLong Pok
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendArenaStatus(ByVal Arena As Long, ByVal Status As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SArenas
    Buffer.WriteLong Arena
    Buffer.WriteLong Status
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendNoticia(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SNoticia
    Buffer.WriteString Msg
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Sub SendOrganização(ByVal Index As Long, Optional ByVal Membros As Boolean)
Dim Buffer As clsBuffer
Dim i As Long

If Player(Index).ORG = 0 Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong SOrganização
    If Player(Index).ORG > 0 Then
        Buffer.WriteLong Organization(Player(Index).ORG).EXP
        Buffer.WriteLong Organization(Player(Index).ORG).Level
        Buffer.WriteLong GetONextLevel(Index)
    End If
    
    If Membros = True Then
        Buffer.WriteByte 1
        For i = 1 To MAX_ORG_MEMBERS
            Buffer.WriteString Trim$(Organization(Player(Index).ORG).OrgMember(i).User_Name)
            Buffer.WriteByte Organization(Player(Index).ORG).OrgMember(i).Online
            Buffer.WriteByte Organization(Player(Index).ORG).OrgMember(i).Used
        Next
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendOrganizaçãoToOrg(ByVal OrgNum As Byte)
Dim Buffer As clsBuffer
Dim i As Long

    'Evitar OverFlow
    If OrgNum = 0 Or OrgNum > MAX_ORGS Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong SOrganização
    
    Buffer.WriteLong Organization(OrgNum).EXP
    Buffer.WriteLong Organization(OrgNum).Level
    Buffer.WriteLong GetONextLevel(1, OrgNum)
    Buffer.WriteByte 1
    
    For i = 1 To MAX_ORG_MEMBERS
        Buffer.WriteString Trim$(Organization(OrgNum).OrgMember(i).User_Name)
        Buffer.WriteByte Organization(OrgNum).OrgMember(i).Online
        Buffer.WriteByte Organization(OrgNum).OrgMember(i).Used
    Next
    
    'Mandar Packet Só para membros da Organização
    For i = 1 To Player_HighIndex
        If Player(i).ORG = OrgNum Then
            SendDataTo i, Buffer.ToArray()
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Sub SendAbrir(ByVal Index As Long, ByVal Msg As String, ByVal Evento As Byte)
'Dim Buffer As clsBuffer
    
'    Set Buffer = New clsBuffer
'    Buffer.WriteLong SEvento
'    Buffer.WriteString Msg
'    Buffer.WriteByte Evento
'    SendDataTo Index, Buffer.ToArray()
'    Set Buffer = Nothing
End Sub

Sub SendOrgShop(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOrgShop
    For i = 1 To MAX_ORG_SHOP
        Buffer.WriteLong OrgShop(i).Item
        Buffer.WriteLong OrgShop(i).Quantia
        Buffer.WriteLong OrgShop(i).Valor
        Buffer.WriteLong OrgShop(i).Level
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendChatBubble(ByVal MapNum As Long, ByVal target As Long, ByVal targetType As Long, ByVal message As String, ByVal Colour As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatBubble
    Buffer.WriteLong target
    Buffer.WriteLong targetType
    Buffer.WriteString message
    Buffer.WriteLong Colour
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendVipPointsInfo(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SVipInfo
    Buffer.WriteLong Player(Index).VipPoints
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
    
    SendPlayerData Index
End Sub

Sub SendAparencia(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    'Buffer.WriteLong SAparencia
    Buffer.WriteLong Index
    Buffer.WriteByte Player(Index).Sex
    
    'Modelo
    Buffer.WriteInteger Player(Index).HairModel
    Buffer.WriteInteger Player(Index).ClothModel
    Buffer.WriteInteger Player(Index).LegsModel
    
    'Cor
    Buffer.WriteByte Player(Index).HairColor
    Buffer.WriteByte Player(Index).ClothColor
    Buffer.WriteByte Player(Index).LegsColor
    
    'Numero
    Buffer.WriteInteger Player(Index).HairNum
    Buffer.WriteInteger Player(Index).ClothNum
    Buffer.WriteInteger Player(Index).LegsNum
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerRun(ByVal Index As Long, Optional ByVal ToTarget As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SRunning
    Buffer.WriteLong Index
    
    If TempPlayer(Index).Running = True Then Buffer.WriteByte 1
    If TempPlayer(Index).Running = False Then Buffer.WriteByte 0
    
    If ToTarget = 0 Then SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    If ToTarget > 0 Then SendDataTo ToTarget, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendComandoGym(ByVal Index As Long, ByVal Comando As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SComandGym
    Buffer.WriteByte Comando
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendContagem(ByVal Index As Long, ByVal Tempo As Integer)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SContagem
    Buffer.WriteInteger Tempo
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
