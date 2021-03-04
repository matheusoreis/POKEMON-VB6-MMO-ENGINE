Attribute VB_Name = "modText"
Option Explicit

' Text declares
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal U As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal size As Byte)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GameFont = CreateFont(size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetFont", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal text As String, color As Long)
' If debug mode, handle error then exit out
    Dim OldFont As Long    ' HFONT

    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SetFont(FONT_NAME, FONT_SIZE)
    OldFont = SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, X + 1, Y + 0, text, Len(text))
    Call TextOut(hDC, X + 0, Y + 1, text, Len(text))
    Call TextOut(hDC, X - 1, Y - 0, text, Len(text))
    Call TextOut(hDC, X - 0, Y - 1, text, Len(text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, Y, text, Len(text))
    Call SelectObject(hDC, OldFont)
    Call DeleteObject(GameFont)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTextVisitor(ByVal hDC As Long, ByVal X, ByVal Y, ByVal text As String, color As Long)
' If debug mode, handle error then exit out
    Dim OldFont As Long    ' HFONT

    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SetFont(FONT_NAME2, FONT_SIZE2)
    OldFont = SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, X + 1, Y + 0, text, Len(text))
    Call TextOut(hDC, X + 0, Y + 1, text, Len(text))
    Call TextOut(hDC, X - 1, Y - 0, text, Len(text))
    Call TextOut(hDC, X - 0, Y - 1, text, Len(text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, Y, text, Len(text))
    Call SelectObject(hDC, OldFont)
    Call DeleteObject(GameFont)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTextVisitor", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTextDamage(ByVal hDC As Long, ByVal X, ByVal Y, ByVal text As String, color As Long)
' If debug mode, handle error then exit out
    Dim OldFont As Long    ' HFONT

    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SetFont(FONT_NAME3, FONT_SIZE3)
    OldFont = SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, X + 1, Y + 0, text, Len(text))
    Call TextOut(hDC, X + 0, Y + 1, text, Len(text))
    Call TextOut(hDC, X - 1, Y - 0, text, Len(text))
    Call TextOut(hDC, X - 0, Y - 1, text, Len(text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, Y, text, Len(text))
    Call SelectObject(hDC, OldFont)
    Call DeleteObject(GameFont)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTextDamage", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' GDI text drawing onto buffer
Public Sub DrawTextOrgShop(ByVal hDC As Long, ByVal X, ByVal Y, ByVal text As String, color As Long)
' If debug mode, handle error then exit out
    Dim OldFont As Long    ' HFONT

    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SetFont(FONT_NAME, 12)
    OldFont = SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, 0)
    Call TextOut(hDC, X + 1, Y + 0, text, Len(text))
    Call TextOut(hDC, X + 0, Y + 1, text, Len(text))
    Call TextOut(hDC, X - 1, Y - 0, text, Len(text))
    Call TextOut(hDC, X - 0, Y - 1, text, Len(text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, Y, text, Len(text))
    Call SelectObject(hDC, OldFont)
    Call DeleteObject(GameFont)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayerOrg(ByVal Index As Long)
    If Player(Index).ORG < 0 Then Exit Sub

    Dim TextX As Long
    Dim TextY As Long
    Dim color As Long
    Dim Name As String

    Select Case Player(Index).ORG
    Case 0
        Name = vbNullString
    Case 1
        Name = "Equipe Rocket"
        color = QBColor(BrightRed)
    Case 2
        Name = "Team Magma"
        color = QBColor(BrightRed)
    Case 3
        Name = "Team Aqua"
        color = QBColor(BrightRed)
    Case Else
        Exit Sub
    End Select

    ' calc pos
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Name)))

    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) = 0 Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - (DDSD_Character(GetPlayerSprite(Index)).lHeight / 8) + 26 - Player(Index).PuloSlide
    Else
        TextX = ConvertMapX(Player(Index).TPX * PIC_X) + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Name)))
        TextY = ConvertMapY(Player(Index).TPY * PIC_Y) - (DDSD_Character(Player(Index).TPSprite).lHeight / 8) + 26
    End If

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Name, color)
End Sub


Public Sub DrawPlayerName(ByVal Index As Long)
    Dim TextX As Long, TextX2 As Long, TextXPoke As Long
    Dim TextY As Long, TextY2 As Long, TextYPoke As Long
    Dim color As Long, PokeName As String
    Dim Name As String, Name2 As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check access level
    If GetPlayerPK(Index) = NO Then

        Select Case GetPlayerAccess(Index)
        Case 0
            color = RGB(255, 96, 0)
        Case 1
            color = QBColor(DarkGrey)
        Case 2
            color = QBColor(Cyan)
        Case 3
            color = QBColor(BrightGreen)
        Case 4
            color = QBColor(Yellow)
        End Select

    Else
        color = QBColor(BrightRed)
    End If

    Name = Trim$(Player(Index).Name)

    If Player(Index).VipInName = True Then
        If Player(Index).MyVip > 0 Then
            Name = "[Vip " & Player(Index).MyVip & "] " & Trim$(Player(Index).Name)
        End If
    End If

    ' calc pos
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Name)))

    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - (DDSD_Character(GetPlayerSprite(Index)).lHeight / 8) + 42 - Player(Index).PuloSlide
    End If

    If Player(Index).TPX > 0 And GetPlayerEquipment(Index, weapon) > 0 Then
        Name = Trim$(Player(Index).Name)

        If Player(Index).VipInName = True Then
            If Player(Index).MyVip > 0 Then
                Name = "[Vip " & Player(Index).MyVip & "] " & Trim$(Player(Index).Name)
            End If
        End If

        Name2 = Trim$(Player(Index).Name)

        If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
            If GetPlayerEquipmentShiny(Index, weapon) = 0 Then
                PokeName = Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " [" & GetPlayerEquipmentPokeInfoLevel(Index, weapon) & "]"
            Else
                PokeName = "S." & Trim$(Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).Name) & " [" & GetPlayerEquipmentPokeInfoLevel(Index, weapon) & "]"
            End If

        Else
            PokeName = "Erro #5"
        End If

        If Player(Index).Sprite = 0 Then Player(Index).Sprite = 1
        If Player(Index).TPSprite = 0 Then Player(Index).TPSprite = 1

        TextX2 = ConvertMapX(Player(Index).TPX * PIC_X) + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Name)))
        TextY2 = ConvertMapY(Player(Index).TPY * PIC_Y) - (DDSD_Character(Player(Index).TPSprite).lHeight / 8) + 42
        TextXPoke = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(PokeName)))

        If Player(Index).Flying = 0 Then
            TextYPoke = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - (DDSD_Character(Player(Index).Sprite).lHeight / 8) + 42 - Player(Index).PuloSlide
        Else
            TextYPoke = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - (DDSD_Character(Player(Index).Sprite).lHeight / 8) + 21 - Player(Index).PuloSlide
        End If

        Select Case GetPlayerSprite(Index)
        Case 144 To 151
            TextYPoke = TextYPoke + 22 - Player(Index).PuloSlide
        Case 154, 155
            TextYPoke = TextYPoke + 47 - Player(Index).PuloSlide
        End Select

        Call DrawText(TexthDC, TextX2, TextY2, Name, color)
        Call DrawText(TexthDC, TextXPoke, TextYPoke, PokeName, color)
        Call DrawText(TexthDC, TextX, TextYPoke - 16, Name, color)

    Else
        Call DrawText(TexthDC, TextX, TextY, Name, color)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim color As Long
    Dim Name As String
    Dim NpcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NpcNum = MapNpc(Index).num

    Select Case Npc(NpcNum).Behaviour
    Case NPC_BEHAVIOUR_ATTACKONSIGHT
        color = QBColor(BrightRed)
    Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
        color = QBColor(Yellow)
    Case NPC_BEHAVIOUR_GUARD
        color = QBColor(Grey)
    Case Else
        color = QBColor(BrightGreen)
    End Select

    If MapNpc(Index).Desmaiado = False Then
        If MapNpc(Index).Shiny = True Then
            color = QBColor(BrightCyan)
            Name = "S." & Trim$(Npc(NpcNum).Name) & " [" & MapNpc(Index).Level & "]"
        Else
            If Npc(MapNpc(Index).num).Behaviour = 1 Or Npc(MapNpc(Index).num).Behaviour = 0 Then
                Name = Trim$(Npc(NpcNum).Name) & " [" & MapNpc(Index).Level & "]"
            Else
                Name = Trim$(Npc(NpcNum).Name)
            End If
        End If
    Else
        Name = vbNullString
    End If

    TextX = ConvertMapX(MapNpc(Index).X * PIC_X) + MapNpc(Index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Name)))

    If Npc(NpcNum).Sprite < 1 Or Npc(NpcNum).Sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).YOffset - 16
    Else
        If MapNpc(Index).Shiny = True Then
            ' Determine location for text
            TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).YOffset - (DDSD_Character(Npc(NpcNum).Sprite + 1).lHeight / 8) + 42
        Else
            ' Determine location for text
            TextY = ConvertMapY(MapNpc(Index).Y * PIC_Y) + MapNpc(Index).YOffset - (DDSD_Character(Npc(NpcNum).Sprite).lHeight / 8) + 42
        End If
    End If

    Select Case Npc(MapNpc(Index).num).Sprite
    Case 144 To 151
        TextY = TextY + 22
    Case 154, 155
        TextY = TextY + 47
    End Select

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Trim$(Name), color)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function BltMapAttributes()
    Dim X As Long
    Dim Y As Long
    Dim tx As Long
    Dim ty As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optAttribs.value Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    With Map.Tile(X, Y)
                        tx = ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                        Case TILE_TYPE_BLOCKED
                            DrawText TexthDC, tx, ty, "B", QBColor(BrightRed)
                        Case TILE_TYPE_WARP
                            DrawText TexthDC, tx, ty, "W", QBColor(BrightBlue)
                        Case TILE_TYPE_ITEM
                            DrawText TexthDC, tx, ty, "I", QBColor(White)
                        Case TILE_TYPE_NPCAVOID
                            DrawText TexthDC, tx, ty, "N", QBColor(White)
                        Case TILE_TYPE_KEY
                            DrawText TexthDC, tx, ty, "K", QBColor(White)
                        Case TILE_TYPE_KEYOPEN
                            DrawText TexthDC, tx, ty, "O", QBColor(White)
                        Case TILE_TYPE_RESOURCE
                            DrawText TexthDC, tx, ty, "O", QBColor(Green)
                        Case TILE_TYPE_DOOR
                            DrawText TexthDC, tx, ty, "D", QBColor(Brown)
                        Case TILE_TYPE_NPCSPAWN
                            DrawText TexthDC, tx, ty, "S", QBColor(Yellow)
                        Case TILE_TYPE_SHOP
                            DrawText TexthDC, tx, ty, "S", QBColor(BrightBlue)
                        Case TILE_TYPE_BANK
                            DrawText TexthDC, tx, ty, "B", QBColor(Blue)
                        Case TILE_TYPE_HEAL
                            DrawText TexthDC, tx, ty, "H", QBColor(BrightGreen)
                        Case TILE_TYPE_TRAP
                            DrawText TexthDC, tx, ty, "T", QBColor(BrightRed)
                        Case TILE_TYPE_SLIDE
                            DrawText TexthDC, tx, ty, "S", QBColor(BrightCyan)
                        Case TILE_TYPE_GRASS
                            DrawText TexthDC, tx, ty, "G", QBColor(BrightGreen)
                        Case TILE_TYPE_WATER
                            DrawText TexthDC, tx, ty, "W", QBColor(BrightCyan)
                        Case TILE_TYPE_FISHING
                            DrawText TexthDC, tx, ty, "P", QBColor(Yellow)
                        Case TILE_TYPE_SCRIPT
                            DrawText TexthDC, tx, ty, "S", QBColor(Cyan)
                        Case TILE_TYPE_FLYAVOID
                            DrawText TexthDC, tx, ty, "F", QBColor(Red)
                        Case TILE_TYPE_SIGN
                            DrawText TexthDC, tx, ty, "S", QBColor(White)
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "BltMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub BltActionMsg(ByVal Index As Long)
    Dim X As Long, Y As Long, i As Long, Time As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' does it exist
    If ActionMsg(Index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
    Case ACTIONMSG_STATIC
        Time = 1500

        If ActionMsg(Index).Y > 0 Then
            X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2
        Else
            X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18
        End If

    Case ACTIONMSG_SCROLL
        Time = 1500

        If ActionMsg(Index).Y > 0 Then
            X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
            ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
        Else
            X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.6)
            ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
        End If

    Case ACTIONMSG_SCREEN
        Time = 3000

        ' This will kill any action screen messages that there in the system
        For i = MAX_BYTE To 1 Step -1
            If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                If i <> Index Then
                    ClearActionMsg Index
                    Index = i
                End If
            End If
        Next
        X = (frmMain.picScreen.Width \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
        Y = 425

    Case ACTIONMSG_STATICLOCKED
        Time = 1500

        If ActionMsg(Index).Y > 0 Then
            X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2
        Else
            X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18
        End If

    End Select

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    If GetTickCount < ActionMsg(Index).Created + Time Then
        If IsNumeric(ActionMsg(Index).message) = True Then
            Call DrawTextDamage(TexthDC, X, Y, ActionMsg(Index).message, QBColor(ActionMsg(Index).color))
        Else
            Call DrawText(TexthDC, X, Y, ActionMsg(Index).message, QBColor(ActionMsg(Index).color))
        End If
    Else
        ClearActionMsg Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function getWidth(ByVal DC As Long, ByVal text As String) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    getWidth = frmMain.TextWidth(text) \ 2

    ' Error handler
    Exit Function
errorhandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub AddText(ByVal Msg As String, ByVal color As Integer)
Dim S As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    S = vbNewLine & Msg
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = QBColor(color)
    frmMain.txtChat.SelText = S
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function SplitString(ByVal str As String, ByVal numOfChar As Long) As String()
    Dim sArr() As String
    Dim nCount As Long
    Dim X As Long
    X = Len(str) \ numOfChar
    If X * numOfChar = Len(str) Then    ' evently divisible
        ReDim sArr(1 To X)
    Else
        ReDim sArr(1 To X + 1)
    End If
    For X = 1 To Len(str) Step numOfChar
        nCount = nCount + 1
        sArr(nCount) = Mid$(str, X, numOfChar)
    Next
    SplitString = sArr

End Function

Public Sub DrawPingText()
    Dim PingToDraw As String

    PingToDraw = Ping

    Select Case Ping
    Case -1
        PingToDraw = "Syncing"
    Case 0 To 15
        PingToDraw = "Local"
    End Select

    If LoadingPing > 0 Then
        PingToDraw = "Syncing"
    End If

    Call DrawText(TexthDC, Camera.Left + 10, Camera.top + 120, "Ping: " & PingToDraw, QBColor(White))
End Sub

Public Sub DrawNoticiaText()
    Dim X As Long, Y As Long, Noticia As String

    If Not NoticiaServ(1) = vbNullString Then
        Noticia = NoticiaServ(1)
    Else
        Noticia = vbNullString
    End If

    X = Camera.Left + ((MAX_MAPX) * 32)
    Y = Camera.top + 150
    DrawText TexthDC, X - NotX, Y, Noticia, QBColor(Yellow)
End Sub

Public Sub DrawQuestsInWindow()
    Dim i As Long, X As Long, Y As Long, QuestString(1 To 5) As String, QuestString2(1 To 5) As String
    Dim Ordem As Byte
    Dim QuestType As Byte, ObjetivoString As String, Task As Byte, QuestNum As Long
    Dim ColQnt As Integer

    For i = 1 To 5
        QuestNum = GetQuestNum(Trim$(QstWin(i)))

        If QstWin(1) = vbNullString Then Exit Sub

        If QuestNum > 0 Then
            Task = Player(MyIndex).Quests(QuestNum).Part
        Else
            Task = 0
        End If

        If Task > 0 Then
            QuestType = Quest(QuestNum).Task(Task).Type
        Else
            QuestType = 0
        End If

        If QuestType = 6 Then Exit Sub

        Select Case QuestType
        Case 0    'Nenhum
            ObjetivoString = "Não Configurado!"
        Case 1    'Derrotar NPCS
            If Quest(QuestNum).Task(Task).num > 0 Then
                ObjetivoString = "Derrotar " & Trim$(Npc(Quest(QuestNum).Task(Task).num).Name) & "[" & Player(MyIndex).Quests(QuestNum).KillNpcs & "/" & Quest(QuestNum).Task(Task).value & "]"
            Else
                ObjetivoString = "Derrotar ???"    '"Derrotar " & Trim$(Npc(Quest(QuestNum).Task(Task).Num).Name) & "[" & Player(MyIndex).Quests(QuestNum).KillNpcs & "/" & Quest(QuestNum).Task(Task).Value & "]"
            End If
        Case 2    'Ganhar De Jogadores
            ObjetivoString = "Ganhar de Jogadores " & "[" & Player(MyIndex).Quests(QuestNum).KillPlayers & "/" & Quest(QuestNum).Task(Task).value & "]"
        Case 3    'Ir Até Mapa
            ObjetivoString = "Vá até o mapa " & Quest(QuestNum).Task(Task).message(2)
        Case 4    'Falar Com Npc
            ObjetivoString = "Fale com " & Trim$(Npc(Quest(QuestNum).Task(Task).num).Name) & "."
        Case 5    'Coletar Itens
            ColQnt = HasItem(MyIndex, Quest(QuestNum).Task(Task).num)
            ObjetivoString = "Coletar " & Trim$(Item(Quest(QuestNum).Task(Task).num).Name) & "[" & ColQnt & "/" & Quest(QuestNum).Task(Task).value & "]"
        Case 7    'Completar Quest no Menu
            ObjetivoString = "[Completa]"
        End Select

        QuestString2(i) = ObjetivoString

        If i > 0 And QstWin(i) = vbNullString Then Ordem = Ordem + 1
        DrawText TexthDC, Camera.Left + 10, Camera.top + 130 + ((i - Ordem) * 40), "[Quest]: " & QstWin(i), QBColor(Yellow)

        If Not QstWin(i) = vbNullString Then
            DrawText TexthDC, Camera.Left + 10, Camera.top + 145 + ((i - Ordem) * 40), QuestString2(i), QBColor(White)
        End If
    Next

End Sub

Public Sub WordWrap_Array(ByVal text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
    Dim lineCount As Long, i As Long, size As Long, lastSpace As Long, B As Long

    'Too small of text
    If Len(text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = text
        Exit Sub
    End If

    ' default values
    B = 1
    lastSpace = 1
    size = 0

    For i = 1 To Len(text)
        ' if it's a space, store it
        Select Case Mid$(text, i, 1)
        Case " ": lastSpace = i
        Case "_": lastSpace = i
        Case "-": lastSpace = i
        End Select

        'Add up the size
        size = size + getWidth(TexthDC, Mid$(text, i, 1))

        'Check for too large of a size
        If size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, (i - 1) - B))
                B = i - 1
                size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, lastSpace - B))
                B = lastSpace + 1

                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                size = getWidth(TexthDC, Mid$(text, lastSpace, i - lastSpace))
            End If
        End If

        ' Remainder
        If i = Len(text) Then
            If B <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(text, B, i)
            End If
        End If
    Next
End Sub

Public Function WordWrap(ByVal text As String, ByVal MaxLineLen As Integer) As String
    Dim TempSplit() As String
    Dim TSLoop As Long
    Dim lastSpace As Long
    Dim size As Long
    Dim i As Long
    Dim B As Long

    'Too small of text
    If Len(text) < 2 Then
        WordWrap = text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbNewLine)

    For TSLoop = 0 To UBound(TempSplit)

        'Clear the values for the new line
        size = 0
        B = 1
        lastSpace = 1

        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine

        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then

            'Loop through all the characters
            For i = 1 To Len(TempSplit(TSLoop))

                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": lastSpace = i
                Case "_": lastSpace = i
                Case "-": lastSpace = i
                End Select

                'Add up the size
                size = size + getWidth(TexthDC, Mid$(TempSplit(TSLoop), i, 1))

                'Check for too large of a size
                If size > MaxLineLen Then
                    'Check if the last space was too far back
                    If i - lastSpace > 12 Then
                        'Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, (i - 1) - B)) & vbNewLine
                        B = i - 1
                        size = 0
                    Else
                        'Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)) & vbNewLine
                        B = lastSpace + 1

                        'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        size = getWidth(TexthDC, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                    End If
                End If

                'This handles the remainder
                If i = Len(TempSplit(TSLoop)) Then
                    If B <> i Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), B, i)
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function

' CHANGE FONT FIX
' PLEASE NOTE THIS WILL FAIL MISERABLY IF YOU DIDN'T APPLY THE FONT MEMORY LEAK FIX FIRST
' Chat Bubble Mondo
' I ONLY DID THIS COZ THE CHATBUBBLE TEXT LOOKS BETTER WITHOUT SHADOW OVER WHITE BUBBLES!
Public Sub DrawTextNoShadow(ByVal hDC As Long, ByVal X, ByVal Y, ByVal text As String, color As Long)
' If debug mode, handle error then exit out
    Dim OldFont As Long    ' HFONT

    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SetFont(FONT_NAME, FONT_SIZE)
    OldFont = SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, Y, text, Len(text))
    Call SelectObject(hDC, OldFont)
    Call DeleteObject(GameFont)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTextNoShadow", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapaItem(ByVal ItemNum As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim color As Long
    Dim Nome As String, PokeballName As String

    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Verificando se existem itens no mapa !!
    If ItemNum > 0 Then

        Select Case Item(MapItem(ItemNum).num).Rarity
        Case 0    'Sem raridade
            color = QBColor(White)
        Case 1
            color = RGB(117, 198, 92)
        Case 2
            color = RGB(103, 140, 224)
        Case 3
            color = RGB(205, 34, 0)
        Case 4
            color = RGB(193, 104, 204)
        Case 5
            color = RGB(217, 150, 64)
        End Select

    Else
        Exit Sub    ' Verificação sem sucesso!!
    End If

    If MapItem(ItemNum).value > 0 Then
        Nome = MapItem(ItemNum).value & " " & Trim$(Item(MapItem(ItemNum).num).Name)
    Else
        If MapItem(ItemNum).PokeInfo.Pokemon = 0 Then
            Nome = Trim$(Item(MapItem(ItemNum).num).Name)
        Else
            Select Case MapItem(ItemNum).PokeInfo.Pokeball
            Case 1
                PokeballName = "Pokéball"
            Case 2
                PokeballName = "GreatBall"
                color = RGB(103, 140, 224)
            Case 3
                PokeballName = "UltraBall"
                color = QBColor(DarkGrey)
            Case 4
                PokeballName = "MasterBall"
                color = RGB(193, 104, 204)
            Case 5
                PokeballName = "RapidBall"
                color = QBColor(Grey)
            Case 6
                PokeballName = "SafariBall"
                color = RGB(117, 198, 92)
            End Select

            Nome = PokeballName & " " & Trim$(Pokemon(MapItem(ItemNum).PokeInfo.Pokemon).Name) & " [" & MapItem(ItemNum).PokeInfo.Level & "]"
        End If
    End If

    ' Calcular coordenadas
    TextX = ConvertMapX(MapItem(ItemNum).X * PIC_X) + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(Nome)))
    If MapItem(ItemNum).num < 1 Or MapItem(ItemNum).num > numitems Then
        TextY = ConvertMapY(MapItem(ItemNum).Y * PIC_Y)
    Else
        ' Determinação do texto
        TextY = ConvertMapY(MapItem(ItemNum).Y * PIC_Y) - (DDSD_Item(MapItem(ItemNum).num).lHeight / 4) + 16
    End If

    ' Execução dos textos
    Call DrawText(TexthDC, TextX, TextY, Nome, color)

    ' Error handlerr
    Exit Sub
errorhandler:
    HandleError "DrawMapaItem", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawContagem()
    DrawText TexthDC, Camera.Left + 664, Camera.top + 560, "Tempo Limite: " & ContagemGym & "Seg", QBColor(White)
End Sub
