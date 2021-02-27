Attribute VB_Name = "ModQuest"
'////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Editado: Orochi ///////////////////
'////////////////////////////////////////////////////////////////////
Option Explicit

' Quest constants
Public Const QUEST_TYPE_NONE As Byte = 0
Public Const QUEST_TYPE_KILLNPC As Byte = 1
Public Const QUEST_TYPE_KILLPLAYER As Byte = 2
Public Const QUEST_TYPE_GOTOMAP As Byte = 3
Public Const QUEST_TYPE_TALKNPC As Byte = 4
Public Const QUEST_TYPE_COLLECTITEMS As Byte = 5
Public Const QUEST_TYPE_POKEDEX As Byte = 6

' Player quest constants
Public Const QUEST_STATUS_NONE As Byte = 0
Public Const QUEST_STATUS_STARTING As Byte = 1
Public Const QUEST_STATUS_COMPLETE As Byte = 2
Public Const QUEST_STATUS_END As Byte = 3

'Types
Public Type PlayerQuestRec
    status As Byte    '0 - None 1-Começou 2-Completa
    Part As Byte    'Tarefa Atual da Quest
    KillNpcs As Long
    KillPlayers As Long
End Type

Public Type QuestTaskRec
Type As Byte
    Message(1 To 3) As String * 255
    Instant As Boolean
    num As Integer
    value As Long
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
    Coin(1 To 3) As Long    '1 - Dollar, 2 - Cash & 3 - Honra
    ItemRew(1 To QUEST_MAX_REWARDS) As Integer
    ValueRew(1 To QUEST_MAX_REWARDS) As Long
    PokeRew(1 To QUEST_MAX_REWARDS) As Integer
    OrgExpRew As Long
    ExpBallRew As Long
End Type

Public Sub PlayerQuests()
    Dim i As Byte, QuestNum As Integer, QuestStatus As Byte, status As String
    Dim Ordem As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear list
    frmMain.lstQuests.Clear

    For i = 1 To MAX_QUESTS
        QuestNum = i
        QuestStatus = Player(MyIndex).Quests(QuestNum).status

        'Limpar QstWin
        If i > 0 And i < 6 Then
            QstWin(i) = vbNullString
        End If

        ' Add in list
        If Player(MyIndex).Quests(QuestNum).status > 0 And Player(MyIndex).Quests(QuestNum).status < 3 Then

            If i - Ordem < 5 And i - Ordem > 0 Then
                QstWin(i - Ordem) = Trim$(Quest(QuestNum).Name)
            End If
            frmMain.lstQuests.AddItem Trim$(Quest(QuestNum).Name)
        Else
            Ordem = Ordem + 1
        End If
    Next

    'Atualizar Caso A janela esteja aberta!
    If frmMain.picQuest.Visible = True Then
        If frmMain.lstQuests.ListCount > 0 Then
            UpdateQuestInfo GetQuestNum(Trim$(frmMain.lstQuests.text))
        End If
    End If

    'Limpar Quest Description
    If frmMain.lstQuests.ListCount = 0 Then
        Call UpdateQuestInfo(0)
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerQuests", "modGamelogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function QuestMaxTasks(ByVal QuestNum As Integer) As Byte
    Dim i As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_QUEST_TASKS
        If Quest(QuestNum).Task(i).Type = QUEST_TYPE_NONE Then
            QuestMaxTasks = i - 1
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "QuestMaxTasks", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetQuestType(ByVal Quest_Type As Byte) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return with QuestType name
    Select Case Quest_Type
    Case QUEST_TYPE_NONE
        GetQuestType = "Nenhum"
    Case QUEST_TYPE_KILLNPC
        GetQuestType = "Derrotar "
    Case QUEST_TYPE_KILLPLAYER
        GetQuestType = "Ganhar "
    Case QUEST_TYPE_GOTOMAP
        GetQuestType = "Ir até o Mapa "
    Case QUEST_TYPE_TALKNPC
        GetQuestType = "Falar com "
    Case QUEST_TYPE_COLLECTITEMS
        GetQuestType = "Coletar "
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetQuestType", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetQuestStatus(ByVal status As Byte) As String
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case status
    Case QUEST_STATUS_NONE
        GetQuestStatus = "[Disponivel]"
    Case QUEST_STATUS_STARTING
        GetQuestStatus = "[Ativa]"
    Case QUEST_STATUS_COMPLETE
        GetQuestStatus = "[Completa]"
    Case QUEST_STATUS_END
        GetQuestStatus = "[Finalizada]"
    Case 4
        GetQuestStatus = "[Repetivel]"
    End Select

    Exit Function
errorhandler:
    HandleError "GetQuestStatus", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetInsiTypeName(ByVal Numero As Byte) As String
    Select Case Numero
    Case 0
        GetInsiTypeName = "None"
    Case 1    'Kanto
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
    Case 9    'Johto
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
    Case 17    'Hoenn
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
    Case 25    'Sinnoh
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

Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim i As Long
    GetQuestNum = 0

    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) = Trim$(QuestName) Then
            GetQuestNum = i
            'AddText i & " - " & Trim$(Quest(i).Name) & "/" & QuestName, BrightRed
            Exit For
        End If
    Next
End Function

Public Sub UpdateQuestInfo(ByVal QuestNum As Long)
    Dim Task As Byte, MaxTask As Byte, QuestType As Byte
    Dim ObjetivoString As String, CHString As String
    Dim ColQnt As Long, X As Long, i As Long

    'Caso Quest For 0 Limpar Quadros
    If QuestNum <= 0 Then
        frmMain.lblQuestInfo(0).Caption = "[Nome]"
        frmMain.lblQuestInfo(1).Caption = "[Descrição]"
        frmMain.lblQuestInfo(2).Caption = "[Tarefa Atual]"
        frmMain.lblQuestInfo(3).Caption = "Tarefa Atual: 0/0"
        frmMain.lblQuestInfo(4).Caption = "???"
        frmMain.lblQuestInfo(5).Caption = "Exp: ??? OrgExp: ???"
        frmMain.lblQuestInfo(6).Caption = "Dollar: ???"
        frmMain.lblQuestInfo(7).Caption = "Honra: ???"
        Exit Sub
    End If

    Task = Player(MyIndex).Quests(QuestNum).Part
    If Task = 0 Then Exit Sub    'Evitar OverFlow

    MaxTask = QuestMaxTasks(QuestNum)
    QuestType = Quest(QuestNum).Task(Task).Type

    Select Case QuestType
    Case 0    'Nenhum
        ObjetivoString = "Não Configurado!"
    Case 1    'Derrotar NPCS
        ObjetivoString = "Derrotar " & Trim$(Npc(Quest(QuestNum).Task(Task).num).Name) & " [" & Player(MyIndex).Quests(QuestNum).KillNpcs & "/" & Quest(QuestNum).Task(Task).value & "]"
    Case 2    'Ganhar De Jogadores
        ObjetivoString = "Ganhar de Jogadores " & " [" & Player(MyIndex).Quests(QuestNum).KillPlayers & "/" & Quest(QuestNum).Task(Task).value & "]"
    Case 3    'Ir Até Mapa
        ObjetivoString = "Vá até o mapa " & Quest(QuestNum).Task(Task).Message(2)
    Case 4    'Falar Com Npc
        ObjetivoString = "Fale com " & Trim$(Npc(Quest(QuestNum).Task(Task).num).Name) & "."
    Case 5    'Coletar Itens
        ColQnt = HasItem(MyIndex, Quest(QuestNum).Task(Task).num)
        ObjetivoString = "Coletar " & Trim$(Item(Quest(QuestNum).Task(Task).num).Name) & "[" & ColQnt & "/" & Quest(QuestNum).Task(Task).value & "]"
    Case 6    'Pokédex
        For i = 1 To UNLOCKED_POKEMONS
            If Player(MyIndex).Pokedex(i) = 1 Then
                X = X + 1
            End If
        Next
        ObjetivoString = "Pokédex:" & "[" & X & "/" & Quest(QuestNum).Task(Task).value & "]"
    Case 7    'Completar Quest no Menu
        ObjetivoString = "Completa"
    End Select

    'Cash ou Honra
    If Quest(QuestNum).Coin(2) > 0 Then
        CHString = "Cash: " & Quest(QuestNum).Coin(2)
    Else
        If Quest(QuestNum).Coin(3) > 0 Then
            CHString = "Honra: " & Quest(QuestNum).Coin(3)
        Else
            CHString = "Honra: 0"
        End If
    End If

    'Setar Informações nas Labels
    frmMain.lblQuestInfo(0).Caption = Trim$(Quest(QuestNum).Name)
    frmMain.lblQuestInfo(1).Caption = Trim$(Quest(QuestNum).Description)
    frmMain.lblQuestInfo(2).Caption = Trim$(Quest(QuestNum).Task(Task).Message(1))
    frmMain.lblQuestInfo(3).Caption = "Tarefa Atual: " & Task & "/" & MaxTask
    frmMain.lblQuestInfo(4).Caption = ObjetivoString
    frmMain.lblQuestInfo(5).Caption = "Exp: " & Quest(QuestNum).ExpBallRew & " OrgExp: " & Quest(QuestNum).OrgExpRew
    frmMain.lblQuestInfo(6).Caption = "Dollar:" & Quest(QuestNum).Coin(1)
    frmMain.lblQuestInfo(7).Caption = CHString

End Sub

Public Sub UpdateSelectQuest(ByVal NpcNum As Integer)
    Dim i As Byte
    Dim NpcQuest As Integer, QuestSlot As Integer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear list
    frmMain.lstSelectQuest.Clear

    For i = 1 To MAX_NPC_QUESTS
        ' Declaration
        NpcQuest = Npc(NpcNum).Quest(i)
        QuestSlot = FindQuestSlot(MyIndex, NpcQuest)

        If NpcQuest > 0 Then
            If QuestSlot > 0 Then
                frmMain.lstSelectQuest.AddItem Trim$(Quest(NpcQuest).Name)    '& " - " & GetQuestStatus(GetPlayerQuestStatus(MyIndex, QuestSlot))
            Else
                frmMain.lstSelectQuest.AddItem Trim$(Quest(NpcQuest).Name)    '& " - " & GetQuestStatus(0)
            End If
        Else
            Exit Sub
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateSelectQuest", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
