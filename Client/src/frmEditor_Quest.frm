VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest Editor"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8205
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   547
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   615
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cmbOptions 
         Height          =   300
         ItemData        =   "frmEditor_Quest.frx":0000
         Left            =   120
         List            =   "frmEditor_Quest.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame fraList 
      Caption         =   "List"
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2895
      Begin VB.ListBox lstIndex 
         Height          =   4200
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame fraData 
      Caption         =   "Data"
      Height          =   3375
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CheckBox ChkBlank 
         Caption         =   "Não Completar"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   65
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox ChkBlank 
         Caption         =   "Não Excluir"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   64
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox ChkDiaria 
         Caption         =   "Diaria?"
         Height          =   180
         Left            =   1320
         TabIndex        =   62
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkRetry 
         Caption         =   "Repetitive?"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtDescription 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   930
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraRequirements 
      Caption         =   "Requirements"
      Height          =   3255
      Left            =   3120
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   4935
      Begin VB.HScrollBar scrlInsignia 
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1680
         Width           =   4695
      End
      Begin VB.HScrollBar scrlQuestReq 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   4695
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   4695
      End
      Begin VB.Frame frmBlank 
         Height          =   1215
         Index           =   2
         Left            =   120
         TabIndex        =   56
         Top             =   1920
         Width           =   2895
         Begin VB.HScrollBar scrlValueReq 
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   840
            Width           =   1095
         End
         Begin VB.HScrollBar scrlItemReq 
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox ChkRetItem 
            Caption         =   "RetirarOItem"
            Height          =   255
            Left            =   1440
            TabIndex        =   57
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblValueReq 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   61
            Top             =   600
            Width           =   645
         End
         Begin VB.Label lblItemReq 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Label lblInsignia 
         AutoSize        =   -1  'True
         Caption         =   "Insignia: None"
         Height          =   180
         Left            =   120
         TabIndex        =   55
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label lblQuestReq 
         AutoSize        =   -1  'True
         Caption         =   "Quest: None"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   960
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Org Level: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame fraTask 
      Caption         =   "Task - 1"
      Height          =   3375
      Left            =   3120
      TabIndex        =   18
      Top             =   840
      Width           =   4935
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Quest.frx":0037
         Left            =   120
         List            =   "frmEditor_Quest.frx":0050
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   840
         Width           =   4695
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   1440
         Value           =   1
         Width           =   2295
      End
      Begin VB.TextBox txtMessage 
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   4695
      End
      Begin VB.TextBox txtMessage 
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtMessage 
         Height          =   270
         Index           =   3
         Left            =   2520
         TabIndex        =   21
         Top             =   2640
         Width           =   2295
      End
      Begin VB.HScrollBar scrlTask 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   20
         Top             =   240
         Value           =   1
         Width           =   4695
      End
      Begin VB.CheckBox chkInstant 
         Caption         =   "Finish instantly?"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   435
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Value: 0"
         Height          =   180
         Left            =   2520
         TabIndex        =   29
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Message - At startup:"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1665
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Message - Not finished:"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Message - To finish:"
         Height          =   180
         Index           =   5
         Left            =   2520
         TabIndex        =   24
         Top             =   2400
         Width           =   1545
      End
   End
   Begin VB.Frame fraRewards 
      Caption         =   "Rewards"
      Height          =   3135
      Left            =   3120
      TabIndex        =   33
      Top             =   840
      Width           =   4935
      Begin VB.Frame frmBlank 
         Caption         =   "Exp Ball"
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   2160
         Width           =   2295
         Begin VB.HScrollBar ScrlExpBall 
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblExpBall 
            BackStyle       =   0  'Transparent
            Caption         =   "Value: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame frmBlank 
         Caption         =   "Org Exp Reward"
         Height          =   855
         Index           =   3
         Left            =   2520
         TabIndex        =   45
         Top             =   2160
         Width           =   2295
         Begin VB.HScrollBar scrlOrgExp 
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblOrgExp 
            BackStyle       =   0  'Transparent
            Caption         =   "Value: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame frmBlank 
         Caption         =   "Currency Coins"
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   2295
         Begin VB.HScrollBar ScrlCoin 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   2055
         End
         Begin VB.HScrollBar ScrlCoin 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   2055
         End
         Begin VB.HScrollBar ScrlCoin 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblCoin 
            BackStyle       =   0  'Transparent
            Caption         =   "Honra: 0"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label lblCoin 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash: 0"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label lblCoin 
            BackStyle       =   0  'Transparent
            Caption         =   "Dollar: 0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame frmItemRew 
         Caption         =   "Item - 1"
         Height          =   1935
         Left            =   2520
         TabIndex        =   34
         Top             =   240
         Width           =   2295
         Begin VB.HScrollBar scrlQItemRew 
            Height          =   255
            Left            =   1920
            Max             =   10
            TabIndex        =   63
            Top             =   0
            Width           =   375
         End
         Begin VB.HScrollBar ScrlValueRew 
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   960
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar ScrlPokeRew 
            Height          =   255
            Left            =   120
            Max             =   251
            TabIndex        =   51
            Top             =   1440
            Width           =   2055
         End
         Begin VB.HScrollBar ScrlItemRew 
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblPokeRew 
            BackStyle       =   0  'Transparent
            Caption         =   "Poké: #0 None"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label lblValueRew 
            BackStyle       =   0  'Transparent
            Caption         =   "Value: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label lblItemRew 
            BackStyle       =   0  'Transparent
            Caption         =   "Item: None"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   3135
         End
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public QItemIndex As Long

Private Sub ChkDiaria_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ChkDiaria.value = 0 Then
        Quest(EditorIndex).Diaria = False
    Else
        Quest(EditorIndex).Diaria = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDiaria_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkInstant_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkInstant.value = 0 Then
        Quest(EditorIndex).Task(QuestTask).Instant = False
    Else
        Quest(EditorIndex).Task(QuestTask).Instant = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkInstant_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ChkRetItem_Click()
    If ChkRetItem.value = 1 Then
        Quest(EditorIndex).RetItemReq = True
    Else
        Quest(EditorIndex).RetItemReq = False
    End If
End Sub

Private Sub chkRetry_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkRetry.value = 0 Then
        Quest(EditorIndex).Retry = False
    Else
        Quest(EditorIndex).Retry = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkRetry_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassRew_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'Quest(EditorIndex).ClassRew = cmbClassRew.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlClassRew_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbOptions_Click()
    Dim Index As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Index = cmbOptions.ListIndex

    If Index = 0 Then    ' Data
        fraData.Visible = True
    Else
        fraData.Visible = False
    End If

    If Index = 1 Then    ' Requirements
        fraRequirements.Visible = True
    Else
        fraRequirements.Visible = False
    End If

    If Index = 2 Then    ' Rewards
        fraRewards.Visible = True
    Else
        fraRewards.Visible = False
    End If

    If Index = 3 Then    ' Task
        fraTask.Visible = True
    Else
        fraTask.Visible = False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbOptions_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    QuestEditorTask
    Quest(EditorIndex).Task(QuestTask).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    QItemIndex = 1

    ' set max values for requeriments
    scrlLevelReq.max = MAX_LEVELS
    scrlQuestReq.max = MAX_QUESTS

    ' set max values for rewards
    scrlQItemRew.max = 10

    ' set max values for others
    scrlTask.max = MAX_QUEST_TASKS

    ' set values
    cmbOptions.ListIndex = 0
    ScrlItemRew.max = MAX_ITEMS
    ScrlPokeRew.max = UNLOCKED_POKEMONS

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    QuestEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlCoin_Change(Index As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case Index
    Case 0
        lblCoin(Index).Caption = "Dollar: " & ScrlCoin(Index).value
    Case 1
        lblCoin(Index).Caption = "Cash: " & ScrlCoin(Index).value
    Case 2
        lblCoin(Index).Caption = "Honra: " & ScrlCoin(Index).value
    End Select

    'Setar Valor Variavel
    Quest(EditorIndex).Coin(Index + 1) = ScrlCoin(Index).value
    ' Error handler
    Exit Sub

errorhandler:
    HandleError "ScrlCoin_Change_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlExpBall_Change()
    lblExpBall.Caption = "Value:" & ScrlExpBall.value
    Quest(EditorIndex).ExpBallRew = ScrlExpBall.value
End Sub

Private Sub scrlInsignia_Change()
    lblInsignia.Caption = "Insignia: " & GetInsiTypeName(scrlInsignia.value)
    Quest(EditorIndex).InsiReq = scrlInsignia.value
End Sub

Private Sub scrlItemReq_Change()
    If scrlItemReq.value > 0 Then
        lblItemReq.Caption = "Item: " & Trim$(Item(scrlItemReq.value).Name)
    Else
        lblItemReq.Caption = "Item: None"
    End If

    Quest(EditorIndex).ItemReq = scrlItemReq.value
End Sub

Private Sub ScrlItemRew_Change()

    If ScrlItemRew.value > 0 And ScrlItemRew.value <= MAX_ITEMS Then
        lblItemRew.Caption = "Item: " & Trim$(Item(ScrlItemRew.value).Name)
    Else
        lblItemRew.Caption = "Item: None"
    End If

    Quest(EditorIndex).ItemRew(QItemIndex) = ScrlItemRew.value
End Sub

Private Sub scrlLevelReq_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblLevelReq.Caption = "Org Level: " & scrlLevelReq.value
    Quest(EditorIndex).OrgLvlReq = scrlLevelReq.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlOrgExp_Change()
    lblOrgExp.Caption = "Value:" & scrlOrgExp.value
    Quest(EditorIndex).OrgExpRew = scrlOrgExp.value
End Sub

Private Sub ScrlPokeRew_Change()
    If ScrlPokeRew.value > 0 Then
        lblPokeRew.Caption = "Poké: #" & ScrlPokeRew.value & " " & Trim$(Pokemon(ScrlPokeRew.value).Name)
    Else
        lblPokeRew.Caption = "Poké: #0 None"
    End If

    Quest(EditorIndex).PokeRew(QItemIndex) = ScrlPokeRew.value
End Sub

Private Sub scrlQItemRew_Change()
    frmItemRew.Caption = "Item - " & scrlQItemRew.value
    QItemIndex = scrlQItemRew.value

    'Evitar OverFlow Desnecessario
    If QItemIndex <= 0 Or QItemIndex > 10 Then
        QItemIndex = 1
        scrlQItemRew.value = 1
    End If

    'Setar Valor nas Scrl
    ScrlItemRew.value = Quest(EditorIndex).ItemRew(QItemIndex)
    ScrlValueRew.value = Quest(EditorIndex).ValueRew(QItemIndex)
    ScrlPokeRew.value = Quest(EditorIndex).PokeRew(QItemIndex)
End Sub

Private Sub scrlTask_Change()
    Dim i As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set the label value
    fraTask.Caption = "Task - " & scrlTask.value

    QuestEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTask_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlNum.value > 0 Then
        If cmbType.ListIndex = QUEST_TYPE_COLLECTITEMS Then
            If Item(scrlNum.value).Type = ITEM_TYPE_CURRENCY Then
                scrlValue.Enabled = True
            Else
                scrlValue.Enabled = False
            End If
        End If
    End If

    lblNum.Caption = "Num: " & scrlNum.value
    Quest(EditorIndex).Task(QuestTask).num = scrlNum.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblValue.Caption = "Value: " & scrlValue.value
    Quest(EditorIndex).Task(QuestTask).value = scrlValue.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlQuestReq_Change()
    Dim sString As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlQuestReq.value = 0 Then sString = "None" Else sString = Trim$(Quest(scrlQuestReq.value).Name)
    lblQuestReq.Caption = "Quest: " & sString
    Quest(EditorIndex).QuestReq = scrlQuestReq.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlQuestReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValueReq_Change()
    If scrlValueReq.value > 0 Then
        lblValueReq.Caption = "Value: " & scrlValueReq.value
    Else
        lblValueReq.Caption = "Value: None"
    End If

    Quest(EditorIndex).ValueReq = scrlValueReq.value
End Sub

Private Sub ScrlValueRew_Change()
    lblValueRew.Caption = "Value: " & ScrlValueRew.value
    Quest(EditorIndex).ValueRew(QItemIndex) = ScrlValueRew.value
End Sub

Private Sub txtDescription_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Quest(EditorIndex).Description = txtDescription.text

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDescription_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage_Change(Index As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Quest(EditorIndex).Task(QuestTask).Message(Index) = txtMessage(Index).text

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    tmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    QuestEditorOk

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    QuestEditorCancel

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearQuest EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    QuestEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
