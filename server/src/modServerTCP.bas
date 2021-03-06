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

Function IsConnected(ByVal index As Long) As Boolean

    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If TempPlayer(index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If LenB(Trim$(Player(index).Login)) > 0 Then
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

Function IsInventoryFull(ByVal tradeTarget As Long, ByVal index As Long) As Boolean
Dim InvEmpty As Long, TradeFull As Long, i As Long

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(tradeTarget, i) > 0 And GetPlayerInvItemNum(tradeTarget, i) <= MAX_ITEMS Then
            InvEmpty = InvEmpty + 1
        End If
    Next
        
    For i = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(i).Num > 0 And TempPlayer(index).TradeOffer(i).Num <= MAX_ITEMS Then
            TradeFull = TradeFull + 1
        End If
    Next
        
    If TradeFull > (MAX_INV - InvEmpty) Then
        IsInventoryFull = True
        Exit Function
    End If
    
    IsInventoryFull = False

End Function

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData Buffer.ToArray()
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

Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
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

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> index Then
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

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataTo index, Buffer.ToArray
    
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

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo index, Buffer.ToArray
    DoEvents
    Call CloseSocket(index)
    
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

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)

    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(index, "You have lost your connection with " & Options.Game_Name & ".")
    End If

End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal index As Long)
Dim i As Long

    If index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(index) & ".")
        Else
            Call AlertMsg(index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
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

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

 If GetPlayerAccess(index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1000 Then
            If GetTickCount < TempPlayer(index).DataTimer Then
                Exit Sub
            End If
        End If
    
        ' Check for packet flooding
        If TempPlayer(index).DataPackets > 25 Then
            If GetTickCount < TempPlayer(index).DataTimer Then
                Exit Sub
            End If
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(index).Buffer.Length >= 4 Then
        pLength = TempPlayer(index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).Buffer.Length - 4
        If pLength <= TempPlayer(index).Buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).Buffer.ReadLong
            HandleData index, TempPlayer(index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).Buffer.Length >= 4 Then
            pLength = TempPlayer(index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)

    If index > 0 Then
        Call LeftGame(index)
        Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        frmServer.Socket(index).Close
        Call UpdateCaption
        Call ClearPlayer(index)
    End If

End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim x As Long
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
    
    For x = 1 To 2
        Buffer.WriteLong Map(MapNum).LevelPoke(x)
    Next

    For x = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            With Map(MapNum).Tile(x, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).x
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

    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(MapNum).Npc(x)
    Next

    MapCache(MapNum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim S As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
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

    Call PlayerMsg(index, S, WhoColor)
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long

    If index > MAX_PLAYERS Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerLevel(index)
    Buffer.WriteLong GetPlayerPOINTS(index)
    Buffer.WriteLong GetPlayerSprite(index)
    Buffer.WriteLong GetPlayerMap(index)
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteLong GetPlayerFlying(index)
    Buffer.WriteLong Player(index).TPX
    Buffer.WriteLong Player(index).TPY
    Buffer.WriteLong Player(index).TPDir
    Buffer.WriteLong Player(index).MySprite
    Buffer.WriteLong Player(index).Vitorias
    Buffer.WriteLong Player(index).Derrotas
    Buffer.WriteByte Player(index).ORG
    Buffer.WriteLong Player(index).Honra
    Buffer.WriteByte Player(index).MyVip
    
    If Player(index).VipInName = True Then
        Buffer.WriteByte 1
    Else
        Buffer.WriteByte 0
    End If
    
    If Player(index).PokeLight = True Then
        Buffer.WriteByte 1
    Else
        Buffer.WriteByte 0
    End If
    
    For i = 1 To MAX_INSIGNIAS
        Buffer.WriteLong Player(index).Insignia(i)
    Next
    
    For i = 1 To MAX_QUESTS
        Buffer.WriteByte Player(index).Quests(i).Status
        Buffer.WriteByte Player(index).Quests(i).Part
    Next
    
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong index
    SendDataToMapBut index, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).x
        Buffer.WriteLong MapItem(MapNum, i).Y
    Next

    SendDataTo index, Buffer.ToArray()
    
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
        Buffer.WriteLong MapItem(MapNum, i).x
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

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
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

    SendDataTo index, Buffer.ToArray()
    
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
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
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

Sub SendItems(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If

    Next

End Sub

Sub SendAnimations(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(index, i)
        End If

    Next

End Sub

Sub SendNpcs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(index, i)
        End If

    Next

End Sub

Sub SendResources(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(index, i)
        End If

    Next

End Sub

Sub SendInventory(ByVal index As Long)
    Dim packet As String
    Dim i As Long, x As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(index, i)
        Buffer.WriteLong GetPlayerInvItemValue(index, i)
        Buffer.WriteLong GetPlayerInvItemPokeInfoPokemon(index, i)
        Buffer.WriteLong GetPlayerInvItemPokeInfoPokeball(index, i)
        Buffer.WriteLong GetPlayerInvItemPokeInfoLevel(index, i)
        Buffer.WriteLong GetPlayerInvItemPokeInfoExp(index, i)
        
        For x = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong GetPlayerInvItemPokeInfoVital(index, i, x)
            Buffer.WriteLong GetPlayerInvItemPokeInfoMaxVital(index, i, x)
        Next
        
        For x = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong GetPlayerInvItemPokeInfoStat(index, i, x)
        Next
        
        For x = 1 To MAX_POKE_SPELL
            Buffer.WriteLong GetPlayerInvItemPokeInfoSpell(index, i, x)
        Next
        
        For x = 1 To MAX_NEGATIVES
            Buffer.WriteLong GetPlayerInvItemNgt(index, i, x)
        Next
        
        For x = 1 To MAX_BERRYS
            Buffer.WriteLong GetPlayerInvItemBerry(index, i, x)
        Next
        
        Buffer.WriteLong GetPlayerInvItemFelicidade(index, i)
        Buffer.WriteLong GetPlayerInvItemSexo(index, i)
        Buffer.WriteLong GetPlayerInvItemShiny(index, i)
                
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal invslot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim x As Long
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invslot
    Buffer.WriteLong GetPlayerInvItemNum(index, invslot)
    Buffer.WriteLong GetPlayerInvItemValue(index, invslot)
    Buffer.WriteLong GetPlayerInvItemPokeInfoPokemon(index, invslot)
    Buffer.WriteLong GetPlayerInvItemPokeInfoPokeball(index, invslot)
    Buffer.WriteLong GetPlayerInvItemPokeInfoLevel(index, invslot)
    Buffer.WriteLong GetPlayerInvItemPokeInfoExp(index, invslot)
    
    For x = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerInvItemPokeInfoVital(index, invslot, x)
        Buffer.WriteLong GetPlayerInvItemPokeInfoMaxVital(index, invslot, x)
    Next
    
    For x = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerInvItemPokeInfoStat(index, invslot, x)
    Next
    
    For x = 1 To MAX_POKE_SPELL
        Buffer.WriteLong GetPlayerInvItemPokeInfoSpell(index, invslot, x)
    Next
    
    For x = 1 To MAX_NEGATIVES
        Buffer.WriteLong GetPlayerInvItemNgt(index, invslot, x)
    Next
        
    For x = 1 To MAX_BERRYS
        Buffer.WriteLong GetPlayerInvItemBerry(index, invslot, x)
    Next
    
    Buffer.WriteLong GetPlayerInvItemFelicidade(index, invslot)
    Buffer.WriteLong GetPlayerInvItemSexo(index, invslot)
    Buffer.WriteLong GetPlayerInvItemShiny(index, invslot)

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Long
    
    Buffer.WriteLong SPlayerWornEq
    
    'Armor Infos
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    
    'Weapon Infos
    Buffer.WriteLong GetPlayerEquipment(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokemon(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokeball(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoLevel(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoExp(index, weapon)
    
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoVital(index, weapon, i)
        Buffer.WriteLong GetPlayerEquipmentPokeInfoMaxVital(index, weapon, i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoStat(index, weapon, i)
    Next
    
    For i = 1 To MAX_POKE_SPELL
        Buffer.WriteLong GetPlayerEquipmentPokeInfoSpell(index, weapon, i)
    Next

    For i = 1 To MAX_NEGATIVES
        Buffer.WriteLong GetPlayerEquipmentNgt(index, weapon, i)
    Next
    
    For i = 1 To MAX_BERRYS
        Buffer.WriteLong GetPlayerEquipmentBerry(index, weapon, i)
    Next
    
    Buffer.WriteLong GetPlayerEquipmentFelicidade(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentSexo(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentShiny(index, weapon)

    'Helmet Infos
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)

    'Shield Infos
    Buffer.WriteLong GetPlayerEquipment(index, Shield)

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Long
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong index
    
    'Armor Infos
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    
    'Weapon Infos
    Buffer.WriteLong GetPlayerEquipment(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokemon(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoPokeball(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoLevel(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentPokeInfoExp(index, weapon)
    
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoVital(index, weapon, i)
        Buffer.WriteLong GetPlayerEquipmentPokeInfoMaxVital(index, weapon, i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerEquipmentPokeInfoStat(index, weapon, i)
    Next
    
    For i = 1 To MAX_POKE_SPELL
        Buffer.WriteLong GetPlayerEquipmentPokeInfoSpell(index, weapon, i)
    Next
    
    For i = 1 To MAX_NEGATIVES
        Buffer.WriteLong GetPlayerEquipmentNgt(index, weapon, i)
    Next
    
    For i = 1 To MAX_BERRYS
        Buffer.WriteLong GetPlayerEquipmentBerry(index, weapon, i)
    Next
    
    Buffer.WriteLong GetPlayerEquipmentFelicidade(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentSexo(index, weapon)
    Buffer.WriteLong GetPlayerEquipmentShiny(index, weapon)

    'Helmet Infos
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)

    'Shield Infos
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal index As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendEXP(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(index)
    Buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStats(ByVal index As Long)
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

Sub SendWelcome(ByVal index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
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

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal index As Long)
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

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
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

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
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
    SendDataTo index, Buffer.ToArray()
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

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
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
    SendDataTo index, Buffer.ToArray()
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

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NpcNum As Long)
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
    SendDataTo index, Buffer.ToArray()
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

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(index, i)
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

Sub SendUpdateShopTo(ByVal index As Long, ByVal shopNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpells(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(index, i)
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

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(index, i)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).x
            Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).Y
        Next

    End If

    SendDataTo index, Buffer.ToArray()
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
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).x
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).Y
        Next

    End If

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SDoorAnimation
    Buffer.WriteLong x
    Buffer.WriteLong Y
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendActionMsg(ByVal MapNum As Long, ByVal message As String, ByVal color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong color
    Buffer.WriteLong MsgType
    Buffer.WriteLong x
    Buffer.WriteLong Y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendAnimation(ByVal MapNum As Long, ByVal Anim As Long, ByVal x As Long, ByVal Y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong x
    Buffer.WriteLong Y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Gru(ByVal MapNum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteString message
    Buffer.WriteString "[Grupo] "
    Buffer.WriteLong saycolour
    
    SendDataTo MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, x As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Bank(index).Item(i).Num
        Buffer.WriteLong Bank(index).Item(i).Value
        Buffer.WriteLong Bank(index).Item(i).PokeInfo.Pokemon
        Buffer.WriteLong Bank(index).Item(i).PokeInfo.Pokeball
        Buffer.WriteLong Bank(index).Item(i).PokeInfo.Level
        Buffer.WriteLong Bank(index).Item(i).PokeInfo.EXP
        
        For x = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong Bank(index).Item(i).PokeInfo.Vital(x)
            Buffer.WriteLong Bank(index).Item(i).PokeInfo.MaxVital(x)
        Next
        
        For x = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Bank(index).Item(i).PokeInfo.Stat(x)
        Next
        
        For x = 1 To MAX_POKE_SPELL
            Buffer.WriteLong Bank(index).Item(i).PokeInfo.Spells(x)
        Next
        
        For x = 1 To MAX_NEGATIVES
            Buffer.WriteLong Bank(index).Item(i).PokeInfo.Negatives(x)
        Next
        
        For x = 1 To MAX_BERRYS
            Buffer.WriteLong Bank(index).Item(i).PokeInfo.Berry(x)
        Next
        
        Buffer.WriteLong Bank(index).Item(i).PokeInfo.Felicidade
        Buffer.WriteLong Bank(index).Item(i).PokeInfo.Sexo
        Buffer.WriteLong Bank(index).Item(i).PokeInfo.Shiny
    Next
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKey(ByVal index As Long, ByVal x As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong Y
    Buffer.WriteByte Value
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong Y
    Buffer.WriteByte Value
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendTrade(ByVal index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal dataType As Byte)
Dim Buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
Dim x As Long
    
    tradeTarget = TempPlayer(index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
        
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Value
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Pokemon
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Pokeball
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Level
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.EXP
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Felicidade
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Sexo
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Shiny
            
            For x = 1 To Vitals.Vital_Count - 1
                Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Vital(x)
                Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.MaxVital(x)
            Next
            
            For x = 1 To Stats.Stat_Count - 1
                Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Stat(x)
            Next
            
            For x = 1 To MAX_POKE_SPELL
                Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Spells(x)
            Next
            
            For x = 1 To MAX_BERRYS
                Buffer.WriteLong TempPlayer(index).TradeOffer(i).PokeInfo.Berry(x)
            Next
            
            ' add total worth
            If TempPlayer(index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price * TempPlayer(index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price
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
            
            For x = 1 To Vitals.Vital_Count - 1
                Buffer.WriteLong GetPlayerInvItemPokeInfoVital(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, x)
                Buffer.WriteLong GetPlayerInvItemPokeInfoMaxVital(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, x)
            Next
            
            For x = 1 To Stats.Stat_Count - 1
                Buffer.WriteLong GetPlayerInvItemPokeInfoStat(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, x)
            Next
            
            For x = 1 To MAX_POKE_SPELL
                Buffer.WriteLong GetPlayerInvItemPokeInfoSpell(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, x)
            Next
            
            For x = 1 To MAX_BERRYS
                Buffer.WriteLong GetPlayerInvItemBerry(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, x)
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
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTarget(ByVal index As Long, Optional ByVal MapNpcNum As Integer)
Dim Buffer As clsBuffer
Dim MapNum As Integer
MapNum = GetPlayerMap(index)

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(index).target
    Buffer.WriteLong TempPlayer(index).targetType
    Buffer.WriteInteger MapNpcNum
    If MapNpcNum > 0 Then
        Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Vital(1)
    Else
        Buffer.WriteLong 0
    End If
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(index).Hotbar(i).Slot
        Buffer.WriteByte Player(index).Hotbar(i).sType
        Buffer.WriteLong Player(index).Hotbar(i).Pokemon
        Buffer.WriteLong Player(index).Hotbar(i).Pokeball
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLoginOk(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong index
    Buffer.WriteLong Player_HighIndex
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    SendDataTo index, Buffer.ToArray()
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

Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal index As Long, ByVal x As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal targetPlayer As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(targetPlayer).Name)
    SendDataTo index, Buffer.ToArray()
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

Sub SendPartyUpdateTo(ByVal index As Long)
Dim Buffer As clsBuffer, i As Long, partynum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partynum = TempPlayer(index).inParty
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partynum As Long, ByVal index As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(index, i)
        Buffer.WriteLong Player(index).Vital(i)
    Next
    SendDataToParty partynum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal MapNum As Long, ByVal index As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong index
    Buffer.WriteLong MapItem(MapNum, index).Num
    Buffer.WriteLong MapItem(MapNum, index).Value
    Buffer.WriteLong MapItem(MapNum, index).x
    Buffer.WriteLong MapItem(MapNum, index).Y
    
    Buffer.WriteLong MapItem(MapNum, index).PokeInfo.Pokemon
    Buffer.WriteLong MapItem(MapNum, index).PokeInfo.Pokeball
    Buffer.WriteLong MapItem(MapNum, index).PokeInfo.Level
    Buffer.WriteLong MapItem(MapNum, index).PokeInfo.EXP
    
    For i = 1 To Vitals.Vital_Count - 1
    Buffer.WriteLong MapItem(MapNum, index).PokeInfo.Vital(i)
    Buffer.WriteLong MapItem(MapNum, index).PokeInfo.MaxVital(i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
    Buffer.WriteLong MapItem(MapNum, index).PokeInfo.Stat(i)
    Next
    
    For i = 1 To MAX_POKE_SPELL
    Buffer.WriteLong MapItem(MapNum, index).PokeInfo.Spells(i)
    Next
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAttack(ByVal index As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong ServerPackets.SAttack
Buffer.WriteLong index

SendDataToMap GetPlayerMap(index), Buffer.ToArray()
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

Sub SendPokeEvolution(ByVal index As Long, ByVal Command As Byte)
Dim Buffer As clsBuffer
    
    If Player(index).EvolPermition = 0 Then Exit Sub
    If GetPlayerEquipmentPokeInfoPokemon(index, weapon) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPokeEvo
    Buffer.WriteByte Command
    Buffer.WriteByte Player(index).EvolPermition
    Buffer.WriteInteger Player(index).EvoId
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInFishing(ByVal index As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong ServerPackets.SFishing
Buffer.WriteLong index

If Player(index).InFishing > 0 Then
Buffer.WriteByte 1
Else
Buffer.WriteByte 0
End If

If TempPlayer(index).ScanTime > 0 Then
Buffer.WriteByte 1
Else
Buffer.WriteByte 0
End If

SendDataToMap GetPlayerMap(index), Buffer.ToArray()
Set Buffer = Nothing
End Sub

'#####################
'###### QUESTS #######
'#####################

Sub SendQuests(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(index, i)
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

Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Long)
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
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendQuestCommand(ByVal index As Integer, ByVal Command As Byte, Optional ByVal Value As Long, Optional ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim MapName As String

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SQuestCommand
    Buffer.WriteByte Command
    Buffer.WriteLong Value
    
    If Command = 2 Then
        Buffer.WriteLong QuestNum
        Buffer.WriteLong Player(index).Quests(QuestNum).KillNpcs
        Buffer.WriteLong Player(index).Quests(QuestNum).KillPlayers
    End If
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDialogue(ByVal index As Integer, ByVal Title As String, ByVal Text As String, ByVal dType As Byte, ByVal YesNo As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SDialogue
    Buffer.WriteString Title
    Buffer.WriteString Text
    Buffer.WriteByte dType
    Buffer.WriteLong YesNo
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAttLeilao()
Dim Buffer As clsBuffer
Dim i As Long, x As Long

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
        
        For x = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong Leilao(i).Poke.Vital(x)
            Buffer.WriteLong Leilao(i).Poke.MaxVital(x)
        Next
        
        For x = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Leilao(i).Poke.Stat(x)
        Next
        
        For x = 1 To MAX_POKE_SPELL
            Buffer.WriteLong Leilao(i).Poke.Spells(x)
        Next
        
        For x = 1 To MAX_NEGATIVES
            Buffer.WriteLong Leilao(i).Poke.Negatives(x)
        Next
        
        For x = 1 To MAX_BERRYS
            Buffer.WriteLong Leilao(i).Poke.Berry(x)
        Next
    Next
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAttLeilaoTo(ByVal index As Long)
Dim Buffer As clsBuffer
Dim i As Long, x As Long

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
        
        For x = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong Leilao(i).Poke.Vital(x)
            Buffer.WriteLong Leilao(i).Poke.MaxVital(x)
        Next
        
        For x = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Leilao(i).Poke.Stat(x)
        Next
        
        For x = 1 To MAX_POKE_SPELL
            Buffer.WriteLong Leilao(i).Poke.Spells(x)
        Next
        
        For x = 1 To MAX_NEGATIVES
            Buffer.WriteLong Leilao(i).Poke.Negatives(x)
        Next
        
        For x = 1 To MAX_BERRYS
            Buffer.WriteLong Leilao(i).Poke.Berry(x)
        Next
    Next
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendChat(ByVal C As Long, ByVal index As Long, ByVal Mandou As Long, ByVal T As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCChat
    Buffer.WriteLong C
    Buffer.WriteLong Mandou
    Buffer.WriteString T
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendEscolherPokeInicial(ByVal index As Long)
Dim Buffer As clsBuffer

If Player(index).PokeInicial = 0 Then Exit Sub
                                
Set Buffer = New clsBuffer
Buffer.WriteLong SPokeSelect
SendDataTo index, Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendSurfInit(ByVal index As Long, Optional ByVal ToTarget As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer
Buffer.WriteLong SSurfInit
Buffer.WriteLong index
Buffer.WriteByte Player(index).InSurf

If ToTarget > 0 Then
    SendDataTo ToTarget, Buffer.ToArray()
Else
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
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

Sub SendRankLevelTo(ByVal index As Long)
Dim Buffer As clsBuffer
Dim i As Long

Set Buffer = New clsBuffer
Buffer.WriteLong SUpdateRankLevel

For i = 1 To MAX_RANKS
    Buffer.WriteString RankLevel(i).Name
    Buffer.WriteLong RankLevel(i).Level
    Buffer.WriteLong RankLevel(i).PokeNum
Next

SendDataTo index, Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendAprenderSpell(ByVal index As Long, ByVal Command As Long)
Dim Buffer As clsBuffer

Set Buffer = New clsBuffer

    Buffer.WriteLong SAprender
    Buffer.WriteByte Command
    Buffer.WriteInteger Player(index).LearnSpell(1)
    Buffer.WriteInteger Player(index).LearnSpell(2)
    
    SendDataTo index, Buffer.ToArray()
Set Buffer = Nothing

End Sub

Sub SendLutarComando(C, T, index, CC, A, Pok)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCLutar
    Buffer.WriteLong C
    Buffer.WriteLong T
    Buffer.WriteLong CC
    Buffer.WriteLong A
    Buffer.WriteLong Pok
    SendDataTo index, Buffer.ToArray()
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

Sub SendOrganização(ByVal index As Long, Optional ByVal Membros As Boolean)
Dim Buffer As clsBuffer
Dim i As Long

If Player(index).ORG = 0 Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong SOrganização
    If Player(index).ORG > 0 Then
        Buffer.WriteLong Organization(Player(index).ORG).EXP
        Buffer.WriteLong Organization(Player(index).ORG).Level
        Buffer.WriteLong GetONextLevel(index)
    End If
    
    If Membros = True Then
        Buffer.WriteByte 1
        For i = 1 To MAX_ORG_MEMBERS
            Buffer.WriteString Trim$(Organization(Player(index).ORG).OrgMember(i).User_Name)
            Buffer.WriteByte Organization(Player(index).ORG).OrgMember(i).Online
            Buffer.WriteByte Organization(Player(index).ORG).OrgMember(i).Used
        Next
    End If
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendAbrir(ByVal index As Long, ByVal Msg As String, ByVal Evento As Byte)
'Dim Buffer As clsBuffer
    
'    Set Buffer = New clsBuffer
'    Buffer.WriteLong SEvento
'    Buffer.WriteString Msg
'    Buffer.WriteByte Evento
'    SendDataTo Index, Buffer.ToArray()
'    Set Buffer = Nothing
End Sub

Sub SendOrgShop(ByVal index As Long)
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
    SendDataTo index, Buffer.ToArray()
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

Sub SendVipPointsInfo(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SVipInfo
    Buffer.WriteLong Player(index).VipPoints
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    
    SendPlayerData index
End Sub

Sub SendAparencia(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    'Buffer.WriteLong SAparencia
    Buffer.WriteLong index
    Buffer.WriteByte Player(index).Sex
    
    'Modelo
    Buffer.WriteInteger Player(index).HairModel
    Buffer.WriteInteger Player(index).ClothModel
    Buffer.WriteInteger Player(index).LegsModel
    
    'Cor
    Buffer.WriteByte Player(index).HairColor
    Buffer.WriteByte Player(index).ClothColor
    Buffer.WriteByte Player(index).LegsColor
    
    'Numero
    Buffer.WriteInteger Player(index).HairNum
    Buffer.WriteInteger Player(index).ClothNum
    Buffer.WriteInteger Player(index).LegsNum
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerRun(ByVal index As Long, Optional ByVal ToTarget As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SRunning
    Buffer.WriteLong index
    
    If TempPlayer(index).Running = True Then Buffer.WriteByte 1
    If TempPlayer(index).Running = False Then Buffer.WriteByte 0
    
    If ToTarget = 0 Then SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    If ToTarget > 0 Then SendDataTo ToTarget, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendComandoGym(ByVal index As Long, ByVal Comando As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SComandGym
    Buffer.WriteByte Comando
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendContagem(ByVal index As Long, ByVal Tempo As Integer)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SContagem
    Buffer.WriteInteger Tempo
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
