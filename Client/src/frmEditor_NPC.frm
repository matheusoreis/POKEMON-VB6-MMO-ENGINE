VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
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
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   772
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FrameBlank 
      Caption         =   "Data"
      Height          =   1455
      Left            =   8520
      TabIndex        =   57
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   120
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   60
         Top             =   240
         Width           =   960
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   58
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   1200
         TabIndex        =   59
         Top             =   720
         Width           =   660
      End
   End
   Begin VB.Frame fraQuest 
      Caption         =   "Quest - 1"
      Height          =   1335
      Left            =   8520
      TabIndex        =   53
      Top             =   3600
      Width           =   2895
      Begin VB.HScrollBar scrlQuestSlot 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   55
         Top             =   240
         Value           =   1
         Width           =   2655
      End
      Begin VB.HScrollBar scrlQuest 
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblQuest 
         Caption         =   "Quest: None"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Captura!"
      Height          =   1815
      Left            =   8520
      TabIndex        =   48
      Top             =   1680
      Width           =   2895
      Begin VB.HScrollBar scrlPokemon 
         Height          =   255
         Left            =   120
         Max             =   500
         TabIndex        =   51
         Top             =   1200
         Width           =   2655
      End
      Begin VB.HScrollBar ScrlChance 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   49
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblpokemon 
         AutoSize        =   -1  'True
         Caption         =   "Pokémon: None"
         Height          =   180
         Left            =   120
         TabIndex        =   52
         Top             =   960
         Width           =   1785
      End
      Begin VB.Label lblchance 
         AutoSize        =   -1  'True
         Caption         =   "Chance: 0%"
         Height          =   180
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   39
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   38
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   37
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "NPC Properties"
      Height          =   7095
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   3240
         TabIndex        =   43
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   960
         TabIndex        =   42
         Top             =   2520
         Width           =   1575
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2640
         TabIndex        =   41
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   30
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   1320
         List            =   "frmEditor_NPC.frx":334B
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1320
         Width           =   3615
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   28
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   600
         Width           =   3975
      End
      Begin VB.Frame fraDrop 
         Caption         =   "Drop"
         Height          =   2175
         Left            =   120
         TabIndex        =   17
         Top             =   4800
         Width           =   4815
         Begin VB.TextBox txtSpawnSecs 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   21
            Text            =   "0"
            Top             =   600
            Width           =   1815
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   20
            Top             =   1680
            Width           =   3495
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   19
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtChance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   18
            Text            =   "0"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Spawn Rate (in seconds)"
            Height          =   180
            Left            =   120
            TabIndex        =   26
            Top             =   600
            UseMnemonic     =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Chance 1 out of"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stats"
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   4815
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1680
            Max             =   255
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   3240
            Max             =   255
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   120
            Max             =   255
            TabIndex        =   8
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   1680
            Max             =   255
            TabIndex        =   7
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "For: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   450
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Def: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   15
            Top             =   480
            Width           =   465
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "SpAtk: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   14
            Top             =   480
            Width           =   675
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Speed: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "SpDef: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   12
            Top             =   1080
            Width           =   660
         End
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   2640
         TabIndex        =   44
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   2640
         TabIndex        =   32
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6540
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   7320
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbBehaviour_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Npc(EditorIndex).Behaviour = cmbBehaviour.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearNPC EditorIndex

    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex

    NpcEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    scrlSprite.max = NumCharacters
    scrlAnimation.max = MAX_ANIMATIONS
    scrlNum.max = MAX_ITEMS

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call NpcEditorOk

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call NpcEditorCancel

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NpcEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
    Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnimation.value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.value).Name)
    lblAnimation.Caption = "Anim: " & sString
    Npc(EditorIndex).Animation = scrlAnimation.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlChance_Change()

    If ScrlChance.value = 0 Then
        lblchance.Caption = "Chance: 0%"
    Else
        lblchance.Caption = "Chance: " & ScrlChance.value & "%"
    End If

    Npc(EditorIndex).Chance = ScrlChance.value

End Sub

Private Sub scrlPokemon_Change()

    If scrlPokemon.value = 0 Then
        lblpokemon.Caption = "Pokémon: None"
    Else
        lblpokemon.Caption = "#" & scrlPokemon.value & " " & Trim$(Pokemon(scrlPokemon.value).Name)
    End If

    Npc(EditorIndex).Pokemon = scrlPokemon.value

End Sub

Private Sub scrlQuest_Change()
    Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlQuest.value = 0 Then sString = "None" Else sString = Trim$(Quest(scrlQuest.value).Name)
    lblQuest.Caption = "Quest: " & sString
    Npc(EditorIndex).Quest(scrlQuestSlot.value) = scrlQuest.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlQuest_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlQuestSlot_Change()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' label
    fraQuest.Caption = "Quest: " & scrlQuestSlot.value

    ' values
    scrlQuest.value = Npc(EditorIndex).Quest(scrlQuestSlot.value)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlQuestSlot_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblSprite.Caption = "Sprite: " & scrlSprite.value
    Npc(EditorIndex).Sprite = scrlSprite.value
    EditorNpc_BltSprite

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblRange.Caption = "Range: " & scrlRange.value
    Npc(EditorIndex).Range = scrlRange.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblNum.Caption = "Num: " & scrlNum.value

    If scrlNum.value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.value).Name)
    End If

    Npc(EditorIndex).DropItem = scrlNum.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlStat_Change(Index As Integer)
    Dim prefix As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case Index
    Case 1
        prefix = "Str: "
    Case 2
        prefix = "End: "
    Case 3
        prefix = "Int: "
    Case 4
        prefix = "Agi: "
    Case 5
        prefix = "Will: "
    End Select
    lblStat(Index).Caption = prefix & ScrlStat(Index).value
    Npc(EditorIndex).Stat(Index) = ScrlStat(Index).value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblValue.Caption = "Value: " & scrlValue.value
    Npc(EditorIndex).DropItemValue = scrlValue.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAttackSay_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Npc(EditorIndex).AttackSay = txtAttackSay.text

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChance_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not Len(txtChance.text) > 0 Then Exit Sub
    If IsNumeric(txtChance.text) Then Npc(EditorIndex).DropChance = Val(txtChance.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChance_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDamage_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not Len(txtDamage.text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.text) Then Npc(EditorIndex).Damage = Val(txtDamage.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not Len(txtEXP.text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.text) Then Npc(EditorIndex).Exp = Val(txtEXP.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHP_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not Len(txtHP.text) > 0 Then Exit Sub
    If IsNumeric(txtHP.text) Then Npc(EditorIndex).HP = Val(txtHP.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevel_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not Len(txtLevel.text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.text) Then Npc(EditorIndex).Level = Val(txtLevel.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpawnSecs_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not Len(txtSpawnSecs.text) > 0 Then Exit Sub
    Npc(EditorIndex).SpawnSecs = Val(txtSpawnSecs.text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If cmbSound.ListIndex >= 0 Then
        Npc(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Npc(EditorIndex).sound = "None."
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
