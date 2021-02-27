VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14985
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   464
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   999
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraBau 
      Caption         =   "Báu Slot: 1"
      Height          =   2295
      Left            =   3360
      TabIndex        =   107
      Top             =   3600
      Width           =   3015
      Begin VB.CheckBox chkGiveAll 
         Caption         =   "Dar Todos!"
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   1920
         Width           =   1215
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   120
         Max             =   30000
         TabIndex        =   113
         Top             =   1560
         Width           =   1935
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   111
         Top             =   1080
         Width           =   1935
      End
      Begin VB.HScrollBar scrlBau 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   108
         Top             =   240
         Value           =   1
         Width           =   2655
      End
      Begin VB.Label lblValue 
         Caption         =   "Value: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lblNum 
         Caption         =   "Num: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblItemB 
         Caption         =   "Item: None"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   6135
      Left            =   11880
      TabIndex        =   88
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.HScrollBar scrlBr 
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   2280
         Width           =   2775
      End
      Begin VB.HScrollBar scrlEffect 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   99
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instant Cast?"
         Height          =   255
         Left            =   1440
         TabIndex        =   93
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   92
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   91
         Top             =   960
         Width           =   1815
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   90
         Top             =   600
         Width           =   1815
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   89
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblBr 
         Caption         =   "Berry: None"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblEffect 
         AutoSize        =   -1  'True
         Caption         =   "Efeito:"
         Height          =   180
         Left            =   120
         TabIndex        =   98
         Top             =   1680
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "C. Spell:"
         Height          =   180
         Left            =   120
         TabIndex        =   97
         Top             =   1320
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   96
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   95
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   94
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Restrições"
      Height          =   1935
      Left            =   9720
      TabIndex        =   84
      Top             =   4320
      Width           =   2055
      Begin VB.CheckBox ChkNDeposit 
         Caption         =   "Não Armazenar"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox ChkNDrop 
         Caption         =   "Não Dropar"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox ChkNTrade 
         Caption         =   "Não Negociar"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Stone Evolution"
      Height          =   855
      Left            =   9720
      TabIndex        =   82
      Top             =   3360
      Width           =   2055
      Begin VB.ComboBox cmbTipo 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   120
         List            =   "frmEditor_Item.frx":336C
         TabIndex        =   83
         Text            =   "Nenhum"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Bike"
      Height          =   975
      Left            =   9720
      TabIndex        =   79
      Top             =   2280
      Width           =   2055
      Begin VB.HScrollBar scrlVel 
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblVel 
         Caption         =   "Velocidade: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Spell"
      Height          =   375
      Left            =   14160
      TabIndex        =   69
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
      Begin VB.HScrollBar scrlspell4 
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1920
         Width           =   1815
      End
      Begin VB.HScrollBar scrlspell3 
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1440
         Width           =   1815
      End
      Begin VB.HScrollBar scrlspell2 
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   960
         Width           =   1815
      End
      Begin VB.HScrollBar scrlspell1 
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblspell4 
         Caption         =   "Spell-4:"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblspell3 
         Caption         =   "Spell-3:"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblspell2 
         Caption         =   "Spell-2:"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblspell1 
         Caption         =   "Spell-1:"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmPokemon 
      Caption         =   "Pokemon"
      Height          =   2055
      Left            =   9720
      TabIndex        =   66
      Top             =   120
      Width           =   2055
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   480
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   78
         Top             =   960
         Width           =   960
      End
      Begin VB.HScrollBar scrlPokemon 
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblPokemon 
         Caption         =   "Pokemon: Nenhum"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   6255
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   64
         Top             =   2760
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   62
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3405
         Left            =   4200
         List            =   "frmEditor_Item.frx":3412
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4200
         Max             =   30000
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":343B
         Left            =   120
         List            =   "frmEditor_Item.frx":3469
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   65
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   63
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   2880
         TabIndex        =   61
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   58
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   31
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   975
      Left            =   14880
      TabIndex        =   6
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   0
         TabIndex        =   12
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   5640
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   2655
      Left            =   3360
      TabIndex        =   32
      Top             =   3600
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox ChkYesNo 
         Caption         =   "Add"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   106
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox ChkYesNo 
         Caption         =   "Add"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   105
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox ChkYesNo 
         Caption         =   "Add"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   104
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox ChkYesNo 
         Caption         =   "Add"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   103
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox ChkYesNo 
         Caption         =   "Add"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   102
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   3960
         ScaleHeight     =   88
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   96
         TabIndex        =   55
         Top             =   1080
         Width           =   1440
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   3000
         TabIndex        =   54
         Top             =   2160
         Width           =   855
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   3000
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   40
         Top             =   1080
         Value           =   100
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   39
         Top             =   2280
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   38
         Top             =   1920
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   37
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   35
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":34D7
         Left            =   1320
         List            =   "frmEditor_Item.frx":34E7
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   33
         Top             =   840
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   2760
         X2              =   2760
         Y1              =   840
         Y2              =   2520
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   53
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   3000
         TabIndex        =   48
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   47
         Top             =   2280
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   43
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   49
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   50
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   51
         Top             =   720
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub chkGiveAll_Click()
    Item(EditorIndex).GiveAll = chkGiveAll.value
End Sub

Private Sub ChkNDeposit_Click()
    Item(EditorIndex).NDeposit = ChkNDeposit.value
End Sub

Private Sub ChkNDrop_Click()
    Item(EditorIndex).NDrop = ChkNDrop.value
End Sub

Private Sub ChkNTrade_Click()
    Item(EditorIndex).NTrade = ChkNTrade.value
End Sub

Private Sub ChkYesNo_Click(Index As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Item(EditorIndex).YesNo(Index) = ChkYesNo(Index).value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkYesNo_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbBind_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).sound = "None."
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTipo_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Tipo = cmbTipo.ListIndex
End Sub

Private Sub cmbTool_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ClearItem EditorIndex

    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex

    ItemEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    scrlPic.max = numitems
    scrlAnim.max = MAX_ANIMATIONS
    scrlPaperdoll.max = NumPaperdolls
    scrlBau.max = MAX_BAU

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ItemEditorOk

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ItemEditorCancel

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
        'scrlDamage_Change
    Else
        fraEquipment.Visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.Visible = True
        fraEquipment.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
        fraEquipment.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_BAU) Then
        fraBau.Visible = True
    Else
        fraBau.Visible = False
    End If



    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub lblStatBonus_Click(Index As Integer)
'
End Sub

Private Sub lstIndex_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccessReq_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.value
    Item(EditorIndex).AccessReq = scrlAccessReq.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAddHP.Caption = "Add HP: " & scrlAddHp.value
    Item(EditorIndex).AddHP = scrlAddHp.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAddMP.Caption = "Add MP: " & scrlAddMP.value
    Item(EditorIndex).AddMP = scrlAddMP.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAddExp.Caption = "Add Exp: " & scrlAddExp.value
    Item(EditorIndex).AddEXP = scrlAddExp.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
    Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.value).Name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBau_Change()
    fraBau.Caption = "Báu Slot:" & scrlBau.value
    scrlNum.value = Item(EditorIndex).BauItem(scrlBau.value)
    scrlValue.value = Item(EditorIndex).BauValue(scrlBau.value)
End Sub

Private Sub scrlBr_Change()
    Select Case scrlBr.value
    Case 1
        lblBr.Caption = "Berry: Status"
    Case 2
        lblBr.Caption = "Berry: Felicidade"
    Case 3
        lblBr.Caption = "Berry: Retirar Efeito"
    Case 4
        lblBr.Caption = "Berry: Buffs"
    Case 5
        lblBr.Caption = "ExpBall Script"
    Case Else
        lblBr.Caption = "Berry: None"
    End Select
    Item(EditorIndex).Berry = scrlBr.value
End Sub

Private Sub scrlDamage_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.value
    Item(EditorIndex).Data2 = scrlDamage.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = "Num: " & scrlNum.value

    If scrlNum.value > 0 Then
        lblItemB.Caption = "Item: " & Trim$(Item(scrlNum.value).Name)
    Else
        lblItemB.Caption = "Item: None"
    End If

    Item(EditorIndex).BauItem(scrlBau.value) = scrlNum.value
End Sub

Private Sub scrlPaperdoll_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.value
    Call EditorItem_BltPaperdoll

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.value
    Item(EditorIndex).Pic = scrlPic.value
    Call EditorItem_BltItem

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPokemon_Change()
    lblpokemon.Caption = "Pokemon: " & scrlPokemon.value
    Item(EditorIndex).Pokemon = scrlPokemon.value
    Call EditorItem_BltSprite
End Sub

Private Sub scrlPrice_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "Price: " & scrlPrice.value
    Item(EditorIndex).Price = scrlPrice.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.value
    Item(EditorIndex).Rarity = scrlRarity.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.value / 1000 & " sec"
    Item(EditorIndex).speed = scrlSpeed.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlspell1_Change()
    If scrlspell1.value = 0 Then
        lblspell1.Caption = "Spell-1: None"
    Else
        lblspell1.Caption = "Spell-1: " & Trim$(Spell(scrlspell1.value).Name)
    End If
    Item(EditorIndex).Spell(1) = scrlspell1.value
End Sub

Private Sub scrlspell2_Change()
    If scrlspell2.value = 0 Then
        lblspell2.Caption = "Spell-2: None"
    Else
        lblspell2.Caption = "Spell-2: " & Trim$(Spell(scrlspell2.value).Name)
    End If
    Item(EditorIndex).Spell(2) = scrlspell2.value
End Sub

Private Sub scrlspell3_Change()
    If scrlspell3.value = 0 Then
        lblspell3.Caption = "Spell-3: None"
    Else
        lblspell3.Caption = "Spell-3: " & Trim$(Spell(scrlspell3.value).Name)
    End If
    Item(EditorIndex).Spell(3) = scrlspell3.value
End Sub

Private Sub scrlspell4_Change()
    If scrlspell4.value = 0 Then
        lblspell4.Caption = "Spell-4: None"
    Else
        lblspell4.Caption = "Spell-4: " & Trim$(Spell(scrlspell4.value).Name)
    End If
    Item(EditorIndex).Spell(4) = scrlspell4.value
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
    Dim text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case Index
    Case 1
        text = "+ Str: "
    Case 2
        text = "+ End: "
    Case 3
        text = "+ Int: "
    Case 4
        text = "+ Agi: "
    Case 5
        text = "+ Will: "
    End Select

    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case Index
    Case 1
        text = "Str: "
    Case 2
        text = "End: "
    Case 3
        text = "Int: "
    Case 4
        text = "Agi: "
    Case 5
        text = "Will: "
    End Select

    lblStatReq(Index).Caption = text & scrlStatReq(Index).value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If Len(Trim$(Spell(scrlSpell.value).Name)) > 0 Then
        lblSpellName.Caption = "Name: " & Trim$(Spell(scrlSpell.value).Name)
    Else
        lblSpellName.Caption = "Name: None"
    End If

    lblSpell.Caption = "Spell: " & scrlSpell.value

    Item(EditorIndex).Data1 = scrlSpell.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = "Value: " & scrlValue.value
    Item(EditorIndex).BauValue(scrlBau.value) = scrlValue.value
End Sub

Private Sub scrlVel_Change()
    If scrlVel.value = 0 Then
        lblVel.Caption = "Velocidade: 0"
    Else
        lblVel.Caption = "Velocidade: " & scrlVel.value
    End If
    Item(EditorIndex).vel = scrlVel.value
End Sub

Private Sub txtDesc_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.text

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
