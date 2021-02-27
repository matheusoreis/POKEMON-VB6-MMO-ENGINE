Attribute VB_Name = "ModQuest"
'////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Editado: Orochi ///////////////////
'////////////////////////////////////////////////////////////////////

Option Explicit
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Quest Contants
Public Const QUEST_MAX_REWARDS As Byte = 10

Public Const QUEST_TYPE_NONE As Byte = 0
Public Const QUEST_TYPE_KILLNPC As Byte = 1
Public Const QUEST_TYPE_KILLPLAYER As Byte = 2
Public Const QUEST_TYPE_GOTOMAP As Byte = 3
Public Const QUEST_TYPE_TALKNPC As Byte = 4
Public Const QUEST_TYPE_COLLECTITEMS As Byte = 5
Public Const QUEST_TYPE_POKEDEX As Byte = 6

' Quest status constants
Public Const QUEST_STATUS_NONE As Byte = 0
Public Const QUEST_STATUS_STARTING As Byte = 1
Public Const QUEST_STATUS_COMPLETE As Byte = 2
Public Const QUEST_STATUS_END As Byte = 3

Public Type PlayerQuestRec
    Status As Byte '0 - None 1-Começou 2-Task Completa 3-Completa
    Part As Byte 'Tarefa Atual da Quest
    KillNpcs As Integer 'Quantia de Npcs da Task Mortos
    KillPlayers As Integer 'Quantia de Jogadores derrotados em Batalha Tarefa
    Diaria As Boolean 'Já Fez Quest Diaria?
End Type

Public Type QuestTaskRec
    Type As Byte
    message(1 To 3) As String * 255
    Instant As Boolean
    Num As Integer
    Value As Long
End Type

Public Type QuestRec
    ' Properties
    Name As String * NAME_LENGTH
    Description As String * 255
    Retry As Boolean
    Diaria As Boolean
    ' Requirements
    OrgLvlReq As Byte
    QuestReq As Byte
    InsiReq As Byte
    ItemReq As Integer
    ValueReq As Long
    RetItemReq As Boolean
    ' Tasks
    Task(1 To MAX_QUEST_TASKS) As QuestTaskRec
    ' Rewards
    Coin(1 To 3) As Long '1 - Dollar, 2 - Cash & 3 - Honra
    ItemRew(1 To QUEST_MAX_REWARDS) As Integer
    ValueRew(1 To QUEST_MAX_REWARDS) As Long
    PokeRew(1 To QUEST_MAX_REWARDS) As Integer
    OrgExpRew As Long
    ExpBallRew As Long
End Type

Sub ChecarQuest(ByVal Index As Long, ByVal QuestNum As Integer, ByVal TypeQuest As Byte, ByVal TargetIndex As Long)
Dim TarefaAtual As Byte, I As Long, X As Long
Dim Quantia As Long, InvVazio As Byte, IQuestQnt As Byte, PokeNum As Integer
Dim PokeRewVit(1 To Vitals.Vital_Count - 1) As Long, PokeRewStat(1 To Stats.Stat_Count - 1) As Long
Dim SpellRew(1 To 4), SexoRew As Byte
Dim Extra(1 To 3) As Byte '1 - Coin, 2- Coin, 3-Exp Ball

    'Evitar Overflow
    If QuestNum = 0 Then Exit Sub
    TarefaAtual = Player(Index).Quests(QuestNum).Part
        
    Select Case TypeQuest
        Case QUEST_TYPE_KILLNPC
        
            'Verificar se Alvo do jogador tem o mesmo numero que o da Quest
            If TargetIndex = Quest(QuestNum).Task(TarefaAtual).Num Then
            
                'Contador de Mortes do Npc Alvo
                If Player(Index).Quests(QuestNum).KillNpcs + 1 > Quest(QuestNum).Task(TarefaAtual).Value Then
                    Player(Index).Quests(QuestNum).KillNpcs = Quest(QuestNum).Task(TarefaAtual).Value
                Else
                    Player(Index).Quests(QuestNum).KillNpcs = Player(Index).Quests(QuestNum).KillNpcs + 1
                End If
                
                'Mandar Msg e Quantia de Npc Mortos
                SendQuestCommand Index, 2, 0, QuestNum
                PlayerMsg Index, "[" & Trim$(Quest(QuestNum).Name) & "]: " & "Derrotar " & Trim$(Npc(Quest(QuestNum).Task(TarefaAtual).Num).Name) & "[" & Player(Index).Quests(QuestNum).KillNpcs & "/" & Quest(QuestNum).Task(TarefaAtual).Value & "]", Yellow
                
                'Checar Quantia de Npcs Mortos
                If Player(Index).Quests(QuestNum).KillNpcs >= Quest(QuestNum).Task(TarefaAtual).Value Then
                    
                    'Quest Completa Só falar com Npc
                    Player(Index).Quests(QuestNum).Status = 2
                    
                    If Quest(QuestNum).Task(TarefaAtual).Instant = True Then
                            TerminarQuest Index, TarefaAtual, QuestNum
                        Else
                            Player(Index).Quests(QuestNum).KillNpcs = 0
                            Player(Index).Quests(QuestNum).KillPlayers = 0
                            
                            If TarefaAtual + 1 > QuestMaxTasks(QuestNum) Then TerminarQuest Index, TarefaAtual, QuestNum
                            Player(Index).Quests(QuestNum).Part = TarefaAtual + 1
                            If Quest(QuestNum).Task(TarefaAtual).Num > 0 Then TakeInvItem Index, Quest(QuestNum).Task(TarefaAtual).Num, Quest(QuestNum).Task(TarefaAtual).Value
                            
                            'Mandar Msg
                            PlayerMsg Index, "[ Quest: " & Trim$(Quest(QuestNum).Name) & "]:" & Trim$(Quest(QuestNum).Task(TarefaAtual).message(1)), White
                        
                            'Atualizar Jogador
                            SendPlayerData Index
                    End If
                    
                End If
            End If
            
        Case QUEST_TYPE_KILLPLAYER
        '   No Momento Não vou configurar isso
        
        Case QUEST_TYPE_GOTOMAP
            'Verificar se o jogador foi ao mapa Alvo
            If TargetIndex = Quest(QuestNum).Task(TarefaAtual).Num Then
            
                    'Quest Completa Só falar com Npc
                    Player(Index).Quests(QuestNum).Status = 2
            
                    If Quest(QuestNum).Task(TarefaAtual).Instant = True Then
                            TerminarQuest Index, TarefaAtual, QuestNum
                        Else
                            Player(Index).Quests(QuestNum).KillNpcs = 0
                            Player(Index).Quests(QuestNum).KillPlayers = 0
                            
                            If TarefaAtual + 1 > QuestMaxTasks(QuestNum) Then TerminarQuest Index, TarefaAtual, QuestNum
                            Player(Index).Quests(QuestNum).Part = TarefaAtual + 1
                            If Quest(QuestNum).Task(TarefaAtual).Num > 0 Then TakeInvItem Index, Quest(QuestNum).Task(TarefaAtual).Num, Quest(QuestNum).Task(TarefaAtual).Value
                            
                            'Mandar Msg
                            PlayerMsg Index, "[ Quest: " & Trim$(Quest(QuestNum).Name) & "]:" & Trim$(Quest(QuestNum).Task(TarefaAtual).message(1)), White
                        
                            'Atualizar Jogador
                            SendPlayerData Index
                    End If
            End If
            
        Case QUEST_TYPE_TALKNPC
            'Verificar se o Npc que o jogador falou é o Npc Alvo
            If Quest(QuestNum).Task(TarefaAtual).Num = TargetIndex Then
                
                If Quest(QuestNum).Task(TarefaAtual).Instant = True Then
                        TerminarQuest Index, TarefaAtual, QuestNum
                    Else
                        Player(Index).Quests(QuestNum).KillNpcs = 0
                        Player(Index).Quests(QuestNum).KillPlayers = 0
                    
                        If TarefaAtual + 1 > QuestMaxTasks(QuestNum) Then TerminarQuest Index, TarefaAtual, QuestNum
                        If Quest(QuestNum).Task(TarefaAtual).Num > 0 Then TakeInvItem Index, Quest(QuestNum).Task(TarefaAtual).Num, Quest(QuestNum).Task(TarefaAtual).Value
                        Player(Index).Quests(QuestNum).Part = TarefaAtual + 1
                            
                        'Mandar Msg
                        PlayerMsg Index, "[ Quest: " & Trim$(Quest(QuestNum).Name) & "]:" & Trim$(Quest(QuestNum).Task(TarefaAtual).message(1)), White
                        
                        'Atualizar Jogador
                        SendPlayerData Index
                End If
            End If
        
        Case QUEST_TYPE_COLLECTITEMS
            'Verificar se o Item Que o jogador pegou é o mesmo que o Item Alvo
            If Quest(QuestNum).Task(TarefaAtual).Num = TargetIndex Then
            
                'Verificar Quantia de Itens Coletados
                If Item(Quest(QuestNum).Task(TarefaAtual).Num).Type = ITEM_TYPE_CURRENCY Then
                    For I = 1 To MAX_INV
                        If GetPlayerInvItemNum(Index, I) = Quest(QuestNum).Task(TarefaAtual).Num Then
                            Quantia = Quantia + GetPlayerInvItemValue(Index, I)
                            Exit For
                        End If
                    Next
                Else
                    For I = 1 To MAX_INV
                        If GetPlayerInvItemNum(Index, I) = Quest(QuestNum).Task(TarefaAtual).Num Then
                            Quantia = Quantia + 1
                        End If
                    Next
                End If
            
                'Mandar Msg
                PlayerMsg Index, "[ Quest: " & Trim$(Quest(QuestNum).Name) & "]: " & "Coletar " & Trim$(Item(Quest(QuestNum).Task(TarefaAtual).Num).Name) & "[" & Quantia & "/" & Quest(QuestNum).Task(TarefaAtual).Value & "]", Yellow
                
                'Verificar Conclusão da Quest
                If Quantia >= Quest(QuestNum).Task(TarefaAtual).Value Then
                
                    'Quest Completa Só falar com Npc
                    Player(Index).Quests(QuestNum).Status = 2
                
                    If Quest(QuestNum).Task(TarefaAtual).Instant = True Then
                            TerminarQuest Index, TarefaAtual, QuestNum
                        Else
                            Player(Index).Quests(QuestNum).KillNpcs = 0
                            Player(Index).Quests(QuestNum).KillPlayers = 0
                            If Quest(QuestNum).Task(TarefaAtual).Num > 0 Then TakeInvItem Index, Quest(QuestNum).Task(TarefaAtual).Num, Quest(QuestNum).Task(TarefaAtual).Value
                            
                            'Mandar Msg
                            PlayerMsg Index, "[ Quest: " & Trim$(Quest(QuestNum).Name) & "]:" & Trim$(Quest(QuestNum).Task(TarefaAtual).message(1)), White
                            
                            If TarefaAtual + 1 > QuestMaxTasks(QuestNum) Then TerminarQuest Index, TarefaAtual, QuestNum
                            Player(Index).Quests(QuestNum).Part = TarefaAtual + 1
                            
                            'Atualizar Jogador
                            SendPlayerData Index
                    End If
                    
                End If
                
            End If
            
        Case QUEST_TYPE_POKEDEX
            For I = 1 To MAX_POKEMONS
                If Player(Index).Pokedex(I) = 1 Then
                    X = X + 1
                End If
            Next
            
            If Quest(QuestNum).Task(TarefaAtual).Value = X Then
                If TarefaAtual + 1 > QuestMaxTasks(QuestNum) Then TerminarQuest Index, TarefaAtual, QuestNum
                Player(Index).Quests(QuestNum).Part = TarefaAtual + 1
                
                'Atualizar Jogador
                SendPlayerData Index
            End If
        
    End Select
End Sub

Sub TerminarQuest(ByVal Index As Integer, ByVal TarefaAtual As Byte, ByVal QuestNum As Integer)
Dim I As Long, X As Long
Dim Quantia As Long, InvVazio As Byte, IQuestQnt As Byte, PokeNum As Integer
Dim PokeRewVit(1 To Vitals.Vital_Count - 1) As Long, PokeRewStat(1 To Stats.Stat_Count - 1) As Long
Dim SpellRew(1 To 4), SexoRew As Byte
Dim Extra(1 To 3) As Byte '1 - Coin, 2- Coin, 3-Exp Ball
        
        'Calcular Espaço
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, I) = 0 Then
                InvVazio = InvVazio + 1
            End If
        Next
        
        'Verificar se tem Coin 1 na Mochila
        If Quest(QuestNum).Coin(1) > 0 Then
            If HasItem(Index, 1) = 0 Then
                Extra(1) = 1
            End If
        End If
                        
        'Verificar se tem Coin 2 na Mochila
        If Quest(QuestNum).Coin(2) > 0 Then
            If HasItem(Index, 2) = 0 Then
                Extra(2) = 1
            End If
        End If
        
        'Adicionar +1 em Quantidade de Itens recebidos pela quest
        If Quest(QuestNum).ExpBallRew > 0 Then
            Extra(3) = 1
        End If
                    
        'Calcular Qntidade de Items da Quest
        For I = 1 To QUEST_MAX_REWARDS
            If Quest(QuestNum).ItemRew(I) > 0 Then
                IQuestQnt = IQuestQnt + 1 + Extra(1) + Extra(2) + Extra(3)
            End If
        Next
        
        'Deixar de Receber a recompensa e não terminar
        If IQuestQnt > InvVazio Then
            PlayerMsg Index, "Sem espaço suficiente no Inventario para completar a Quest!", BrightRed
            PlayerMsg Index, "Retire " & IQuestQnt & " itens do seu inventario para que seja entregue a recompensa!", BrightGreen
            Exit Sub
        End If
        
        'Recompensas
        For I = 1 To QUEST_MAX_REWARDS
        If Quest(QuestNum).ItemRew(I) > 0 Then
            If Quest(QuestNum).PokeRew(I) = 0 Then
                GiveInvItem Index, Quest(QuestNum).ItemRew(I), Quest(QuestNum).ValueRew(I)
            Else
                'Carregar Informações Pokémon!
                PokeNum = Quest(QuestNum).PokeRew(I)
                
                'Vitalidade Base + (level * 5)
                For X = 1 To Vitals.Vital_Count - 1
                PokeRewVit(X) = Pokemon(PokeNum).Vital(X) + (Quest(QuestNum).ValueRew(I) * 5)
            Next
            
            'Status = De Captura 75%~100% Base + Level * 3
            For X = 1 To Stats.Stat_Count - 1
                If Pokemon(PokeNum).Add_Stat(I) > 0 Then
                    PokeRewStat(X) = Random(Pokemon(PokeNum).Add_Stat(I) * 75 / 100, Pokemon(PokeNum).Add_Stat(I)) + (Quest(QuestNum).ValueRew(I) * 3)
                End If
            Next
            
            'Spell
            For X = 1 To 4
                If Pokemon(PokeNum).Habilidades(X).Spell > 0 Then
                    SpellRew(X) = Pokemon(PokeNum).Habilidades(X).Spell
                End If
            Next
                                            
            'Sexo
            If Int(Rnd * 100) <= Pokemon(PokeNum).ControlSex Then
                SexoRew = 1
            Else
                SexoRew = 0
            End If
                                            
            'Entregar Pokémon
            GiveInvItem Index, Quest(QuestNum).ItemRew(I), 1, False, Quest(QuestNum).PokeRew(I), 1, Quest(QuestNum).ValueRew(I), 0, PokeRewVit(1), PokeRewVit(2), PokeRewVit(1), PokeRewVit(2), PokeRewStat(1), PokeRewStat(4), PokeRewStat(2), PokeRewStat(3), PokeRewStat(5), SpellRew(1), SpellRew(2), SpellRew(3), SpellRew(4), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 50, SexoRew, 0, 0, 0, 0, 0, 0
            End If
        End If
    Next
    
    'Entregar Coin 1
    If Quest(QuestNum).Coin(1) Then
        GiveInvItem Index, 1, Quest(QuestNum).Coin(1)
    End If
    
    'Entregar Coin 2
    If Quest(QuestNum).Coin(2) > 0 Then
        GiveInvItem Index, 2, Quest(QuestNum).Coin(2)
    End If
    
    'Entregar Exp Ball
    If Quest(QuestNum).ExpBallRew > 0 Then
        GiveInvItem Index, 50, 1, False, 0, 0, 0, Quest(QuestNum).ExpBallRew
    End If
    
    'Adicionar Pontos de Honra
    If Quest(QuestNum).Coin(3) > 0 Then
        Call SetPlayerHonra(Index, GetPlayerHonra(Index) + Quest(QuestNum).Coin(3))
    End If
    
    'Adicionar Org Exp
    If Player(Index).ORG > 0 Then
        If Organization(Player(Index).ORG).Level <= 9 Then
            
            'Dar Exp e verificar level UP!
            Organization(Player(Index).ORG).EXP = Organization(Player(Index).ORG).EXP + Quest(QuestNum).OrgExpRew
            Call CheckAORGlevelUP(Player(Index).ORG)
            Call SaveOrgExp(Player(Index).ORG)
            Call SendOrganização(Index)
            
            For I = 1 To Player_HighIndex
                If Player(I).ORG = Player(Index).ORG Then
                    Call SendOrganização(I)
                End If
            Next
        Else
            If Organization(Player(Index).ORG).EXP < GetONextLevel(Index) Then
                Organization(Player(Index).ORG).EXP = GetONextLevel(Index)
            End If
        End If
    End If
    
    'Limpar Valores
    Player(Index).Quests(QuestNum).KillNpcs = 0
    Player(Index).Quests(QuestNum).KillPlayers = 0
    SendQuestCommand Index, 2, 0, QuestNum
    
    'Recolher Item da Ultima etapa Quest
    TakeInvItem Index, Quest(QuestNum).Task(TarefaAtual).Num, Quest(QuestNum).Task(TarefaAtual).Value
    
    If Quest(QuestNum).Retry = True Then
        Player(Index).Quests(QuestNum).Status = 4 'Repetivel
    Else
        Player(Index).Quests(QuestNum).Status = 3 'Finalizada
    End If
        Player(Index).Quests(QuestNum).Part = 0
    
    'Se a quest for Diaria
    If Quest(QuestNum).Diaria = True Then
        Player(Index).Quests(QuestNum).Diaria = True
    End If
    
    'Msg de termino
    PlayerMsg Index, Trim$(Quest(QuestNum).Task(TarefaAtual).message(3)), BrightGreen
    
    'Atualizar Jogador
    SendPlayerData Index
End Sub

'################
'#####QUESTS#####
'################

Public Function GetQuestTypeTwo(ByVal Quest_Type As Byte) As String
    ' Return with QuestType name
    Select Case Quest_Type
        Case QUEST_TYPE_NONE
            GetQuestTypeTwo = "None"
        Case QUEST_TYPE_KILLNPC
            GetQuestTypeTwo = "Derrotar "
        Case QUEST_TYPE_KILLPLAYER
            GetQuestTypeTwo = "Derrotar Jogador "
        Case QUEST_TYPE_GOTOMAP
            GetQuestTypeTwo = "Ir Até "
        Case QUEST_TYPE_TALKNPC
            GetQuestTypeTwo = "Falar com "
        Case QUEST_TYPE_COLLECTITEMS
            GetQuestTypeTwo = "Coletar "
    End Select
End Function

Public Function GetInsiTypeName(ByVal Numero As Byte) As String
    Select Case Numero
    Case 1 'Kanto
        GetInsiTypeName = "Rocha"
    Case 2
        GetInsiTypeName = "Cascata"
    Case 3
        GetInsiTypeName = "Trovão"
    Case 4
        GetInsiTypeName = "Arco Íris"
    Case 5
        GetInsiTypeName = "Alma"
    Case 6
        GetInsiTypeName = "Pântano"
    Case 7
        GetInsiTypeName = "Vulcão"
    Case 8
        GetInsiTypeName = "Terra"
    Case 9 'Johto
        GetInsiTypeName = "Zephir"
    Case 10
        GetInsiTypeName = "Colméia"
    Case 11
        GetInsiTypeName = "Planície"
    Case 12
        GetInsiTypeName = "Névoa"
    Case 13
        GetInsiTypeName = "Tempestade"
    Case 14
        GetInsiTypeName = "Mineral"
    Case 15
        GetInsiTypeName = "Geleira"
    Case 16
        GetInsiTypeName = "Nascente"
    Case 17 'Hoenn
        GetInsiTypeName = "Pedra"
    Case 18
        GetInsiTypeName = "Articulação"
    Case 19
        GetInsiTypeName = "Dínamo"
    Case 20
        GetInsiTypeName = "Calor"
    Case 21
        GetInsiTypeName = "Balança"
    Case 22
        GetInsiTypeName = "Pena"
    Case 23
        GetInsiTypeName = "Mente"
    Case 24
        GetInsiTypeName = "Chuva"
    Case 25 'Sinnoh
        GetInsiTypeName = "Carvão"
    Case 26
        GetInsiTypeName = "Floresta"
    Case 27
        GetInsiTypeName = "Pedregulho"
    Case 28
        GetInsiTypeName = "Brejo"
    Case 29
        GetInsiTypeName = "Relíquia"
    Case 30
        GetInsiTypeName = "Mina"
    Case 31
        GetInsiTypeName = "Sincelo"
    Case 32
        GetInsiTypeName = "Farol"
    Case Else
        GetInsiTypeName = "???"
    End Select
End Function

Sub AceitarQuest(ByVal Index As Integer)
Dim QuestNum As Integer, TarefaAtual As Byte
Dim I As Long

    'Evitar OverFlow
    If Not IsPlaying(Index) Then Exit Sub
    
    QuestNum = TempPlayer(Index).QuestInvite
    
    'Evitar OverFlow²
    If QuestNum = 0 Then Exit Sub
    
    'Verificar se é repetivel ou não!
    If Quest(QuestNum).Retry = False Then
        If Player(Index).Quests(QuestNum).Status <> 0 Then Exit Sub
    Else
        If Player(Index).Quests(QuestNum).Status < 3 Then
            If Player(Index).Quests(QuestNum).Status > 0 Then
                Exit Sub
            End If
        End If
    End If
    
    'Verificar se é diaria
    If Quest(QuestNum).Diaria = True Then
        If Player(Index).Quests(QuestNum).Diaria = True Then
            PlayerMsg Index, "Você já fez está quest hoje volte amanhã!", BrightRed
            Exit Sub
        End If
    End If
    
    'Item Requerido
    If Quest(QuestNum).ItemReq > 0 Then
        For I = 1 To MAX_INV
            If HasItem(Index, Quest(QuestNum).ItemReq) = 0 Then
                PlayerMsg Index, "Você não tem o item " & Trim$(Item(Quest(QuestNum).ItemReq).Name) & " para fazer a quest!", BrightRed
                Exit Sub
            Else
                If Quest(QuestNum).ValueReq > 1 Then
                    If HasItem(Index, Quest(QuestNum).ItemReq) < Quest(QuestNum).ValueReq Then
                        PlayerMsg Index, "Você não tem " & Quest(QuestNum).ValueReq & " " & Trim$(Item(Quest(QuestNum).ItemReq).Name) & " para fazer a quest!", BrightRed
                        Exit Sub
                    End If
                End If
            End If
        Next
    End If
    
    'Pegar Item Requerido
    If Quest(QuestNum).RetItemReq = True Then
        Call TakeInvItem(Index, Quest(QuestNum).ItemReq, Quest(QuestNum).ValueReq)
    End If
    
    'Atribuir Quest ao jogador
    Player(Index).Quests(QuestNum).Status = 1 'Começou a quest
    Player(Index).Quests(QuestNum).Part = 1 'Tarefa 1 da Quest!
    TarefaAtual = 1
    
    'Limpar Dados Npcs/Players
    Player(Index).Quests(QuestNum).KillNpcs = 0
    Player(Index).Quests(QuestNum).KillPlayers = 0
    
    ' Checar Quest's
    ChecarTarefasAtuais Index, QUEST_TYPE_GOTOMAP, GetPlayerMap(Index)
    ChecarTarefasAtuais Index, QUEST_TYPE_COLLECTITEMS, Quest(QuestNum).Task(TarefaAtual).Num

    ' Mandar Informações
    Call SendPlayerData(Index)
    TempPlayer(Index).QuestInvite = 0
End Sub

Sub ChecarReqQuest(ByVal Index As Integer, ByVal NpcNum As Integer, ByVal Slot As Byte)
Dim QuestNum As Integer, TarefaAtual As Byte, I As Long

    'Evitar OverFlow
    If Slot = 0 Then Exit Sub

    'Setar Valor de Questnum
    QuestNum = Npc(NpcNum).Quest(Slot)

    'Evitar OverFlow
    If QuestNum = 0 Then Exit Sub
    
    'Verificar se For Talk Npc
    If Player(Index).Quests(QuestNum).Status > 0 Then
        If Player(Index).Quests(QuestNum).Part > 0 Then
            If Quest(QuestNum).Task(Player(Index).Quests(QuestNum).Part).Type = QUEST_TYPE_TALKNPC Then
                ChecarQuest Index, QuestNum, QUEST_TYPE_TALKNPC, NpcNum
                Exit Sub
            End If
        End If
    End If
    
    'Verificar se é repetivel ou não!
    If Quest(QuestNum).Retry = False Then
        If Player(Index).Quests(QuestNum).Status <> 0 Then Exit Sub
    Else
        If Player(Index).Quests(QuestNum).Status < 3 Then
            If Player(Index).Quests(QuestNum).Status > 0 Then
                Exit Sub
            End If
        End If
    End If
    
    'Verificar se é diaria
    If Quest(QuestNum).Diaria = True Then
        If Player(Index).Quests(QuestNum).Diaria = True Then
            PlayerMsg Index, "Você já fez está quest hoje volte amanhã!", BrightRed
            Exit Sub
        End If
    End If
    
    'Verificar Requisitos - OrgLevel
    If Quest(QuestNum).OrgLvlReq > 0 Then
        If GetPlayerOrg(Index) = 0 Then
            PlayerMsg Index, "Quest só para membros de organizações!", BrightRed
            Exit Sub
        End If
        
        If Organization(GetPlayerOrg(Index)).Level < Quest(QuestNum).OrgLvlReq Then
            PlayerMsg Index, "Sua organização não tem level suficiente para adquirir está Quest!", BrightRed
            Exit Sub
        End If
    End If

    'Quest Requisito
    If Quest(QuestNum).QuestReq > 0 Then
        If Not Player(Index).Quests(Quest(QuestNum).QuestReq).Status = 2 Then
            PlayerMsg Index, "Você precisa fazer a quest: " & Trim$(Quest(Quest(QuestNum).QuestReq).Name) & " antes.", BrightRed
            Exit Sub
        End If
    End If
    
    'Insignia
    If Quest(QuestNum).InsiReq > 0 Then
        If Player(Index).Insignia(Quest(QuestNum).InsiReq) = False Then
            PlayerMsg Index, "Você não possui a Insígnia " & Trim$(GetInsiTypeName(Quest(QuestNum).InsiReq)) & ".", BrightRed
            Exit Sub
        End If
    End If
    
    'Item Requerido
    If Quest(QuestNum).ItemReq > 0 Then
        For I = 1 To MAX_INV
            If HasItem(Index, Quest(QuestNum).ItemReq) = 0 Then
                PlayerMsg Index, "Você não tem o item " & Trim$(Item(Quest(QuestNum).ItemReq).Name) & " para fazer a quest!", BrightRed
                Exit Sub
            Else
                If Quest(QuestNum).ValueReq > 1 Then
                    If HasItem(Index, Quest(QuestNum).ItemReq) < Quest(QuestNum).ValueReq Then
                        PlayerMsg Index, "Você não tem " & Quest(QuestNum).ValueReq & " " & Trim$(Item(Quest(QuestNum).ItemReq).Name) & " para fazer a quest!", BrightRed
                        Exit Sub
                    End If
                End If
            End If
        Next
    End If
    
    ' Caso esteja tudo nos conformes será enviado o dialogo de confirmação
    TempPlayer(Index).QuestInvite = QuestNum
    Call SendDialogue(Index, Trim$(Quest(QuestNum).Name), Trim$(Quest(QuestNum).Description), DIALOGUE_TYPE_QUEST, YES)
End Sub

Public Function QuestMaxTasks(ByVal QuestNum As Integer) As Byte
    Dim I As Byte

    For I = 1 To MAX_QUEST_TASKS
        If Quest(QuestNum).Task(I).Type = QUEST_TYPE_NONE Then
            QuestMaxTasks = I - 1
            Exit Function
        End If
    Next
End Function
