VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   19560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   39915
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   2661
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picPD 
      Height          =   3375
      Left            =   8640
      ScaleHeight     =   3315
      ScaleWidth      =   4035
      TabIndex        =   252
      Top             =   5280
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label lblPI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0CCA
         ForeColor       =   &H00000000&
         Height          =   975
         Index           =   1
         Left            =   0
         TabIndex        =   253
         Top             =   600
         Width           =   4215
      End
   End
   Begin VB.PictureBox picYourTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   3840
      Picture         =   "frmMain.frx":0D81
      ScaleHeight     =   174
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   248
      Top             =   2040
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox PicTradeOn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   150
         Picture         =   "frmMain.frx":BE00
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   249
         Top             =   2310
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblYourWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   251
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblTradeStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Esperando Confirmação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   250
         Top             =   2280
         Width           =   2415
      End
   End
   Begin VB.PictureBox picTheirTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   7320
      Picture         =   "frmMain.frx":DF91
      ScaleHeight     =   174
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   244
      Top             =   2040
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox PicTradeOn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   150
         Picture         =   "frmMain.frx":19010
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   245
         Top             =   2310
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblTheirWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   247
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblTradeStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Esperando Confirmação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   246
         Top             =   2280
         Width           =   2415
      End
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   5880
      Picture         =   "frmMain.frx":1B1A1
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   243
      Top             =   4800
      Visible         =   0   'False
      Width           =   2085
      Begin VB.Image imgAcceptTrade 
         Height          =   270
         Left            =   60
         Top             =   75
         Width           =   930
      End
      Begin VB.Image imgDeclineTrade 
         Height          =   270
         Left            =   1095
         Top             =   75
         Width           =   930
      End
   End
   Begin VB.PictureBox PicLoading 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9600
      Left            =   0
      Picture         =   "frmMain.frx":1F58B
      ScaleHeight     =   640
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1152
      TabIndex        =   242
      Top             =   0
      Width           =   17280
   End
   Begin VB.PictureBox PicSurf 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   30960
      Picture         =   "frmMain.frx":23B5CD
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   214
      TabIndex        =   217
      Top             =   12480
      Visible         =   0   'False
      Width           =   3210
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   12
         Left            =   750
         Picture         =   "frmMain.frx":24C845
         Top             =   990
         Width           =   1695
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   8
         Left            =   2865
         Picture         =   "frmMain.frx":24E715
         Top             =   15
         Width           =   330
      End
   End
   Begin VB.PictureBox picPokeDesc 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3030
      Left            =   26880
      Picture         =   "frmMain.frx":24EE41
      ScaleHeight     =   202
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   196
      Top             =   5640
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox PicFacePokemon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   330
         ScaleHeight     =   96.043
         ScaleMode       =   0  'User
         ScaleWidth      =   94
         TabIndex        =   197
         Top             =   285
         Width           =   1410
         Begin VB.Label lblPokeInfoDesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   213
            Top             =   29
            Width           =   1155
         End
      End
      Begin VB.Label lblPokeInfoDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   212
         Top             =   1545
         Width           =   510
      End
      Begin VB.Label lblPokeInfoDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Felicidade: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   211
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblPokeInfoDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2010
         TabIndex        =   210
         Top             =   300
         Width           =   1785
      End
      Begin VB.Label lblPokeInfoDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   209
         Top             =   1095
         Width           =   450
      End
      Begin VB.Label lblPokeInfoDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hp: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   2040
         TabIndex        =   208
         Top             =   645
         Width           =   360
      End
      Begin VB.Label lblPokeInfoDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mp: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   2040
         TabIndex        =   207
         Top             =   870
         Width           =   375
      End
      Begin VB.Label lblPokeInfoDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Atq: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   206
         Top             =   1770
         Width           =   1575
      End
      Begin VB.Label lblPokeInfoDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Def: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   2040
         TabIndex        =   205
         Top             =   2220
         Width           =   1455
      End
      Begin VB.Label lblPokeInfoDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "EAtq: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   204
         Top             =   2445
         Width           =   1335
      End
      Begin VB.Label lblPokeInfoDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EDef: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   2040
         TabIndex        =   203
         Top             =   2670
         Width           =   540
      End
      Begin VB.Label lblPokeInfoDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vel: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   2040
         TabIndex        =   202
         Top             =   1995
         Width           =   405
      End
      Begin VB.Label lblPokeInfoDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   225
         TabIndex        =   201
         Top             =   1875
         Width           =   1635
      End
      Begin VB.Label lblPokeInfoDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   225
         TabIndex        =   200
         Top             =   2100
         Width           =   1635
      End
      Begin VB.Label lblPokeInfoDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   15
         Left            =   225
         TabIndex        =   199
         Top             =   2325
         Width           =   1635
      End
      Begin VB.Label lblPokeInfoDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   16
         Left            =   225
         TabIndex        =   198
         Top             =   2550
         Width           =   1575
      End
   End
   Begin VB.PictureBox PicPokeSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   4
      Left            =   180
      Picture         =   "frmMain.frx":274FAD
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   192
      Top             =   3075
      Width           =   2670
      Begin VB.Label lblPokeSpell 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   105
         TabIndex        =   193
         Top             =   60
         Width           =   2445
      End
   End
   Begin VB.PictureBox PicPokeSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   180
      Picture         =   "frmMain.frx":278EC1
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   190
      Top             =   2550
      Width           =   2670
      Begin VB.Label lblPokeSpell 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   105
         TabIndex        =   191
         Top             =   60
         Width           =   2445
      End
   End
   Begin VB.PictureBox PicPokeSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   180
      Picture         =   "frmMain.frx":27CDD5
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   188
      Top             =   2025
      Width           =   2670
      Begin VB.Label lblPokeSpell 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   105
         TabIndex        =   189
         Top             =   60
         Width           =   2445
      End
   End
   Begin VB.PictureBox picAtalho 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   22440
      Picture         =   "frmMain.frx":280CE9
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   186
      Top             =   7080
      Visible         =   0   'False
      Width           =   2400
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   2
         Left            =   2055
         Picture         =   "frmMain.frx":2A7D8D
         Top             =   15
         Width           =   330
      End
      Begin VB.Image imgButton 
         Height          =   465
         Index           =   4
         Left            =   60
         Picture         =   "frmMain.frx":2A84B9
         Top             =   4410
         Width           =   2280
      End
      Begin VB.Image imgButton 
         Appearance      =   0  'Flat
         Height          =   465
         Index           =   2
         Left            =   60
         Picture         =   "frmMain.frx":2ABC35
         Top             =   3930
         Width           =   2280
      End
      Begin VB.Image imgButton 
         Height          =   465
         Index           =   9
         Left            =   60
         Picture         =   "frmMain.frx":2AF3B1
         Top             =   3435
         Width           =   2280
      End
      Begin VB.Image imgButton 
         Height          =   465
         Index           =   8
         Left            =   60
         Picture         =   "frmMain.frx":2B2B2D
         Top             =   2940
         Width           =   2280
      End
      Begin VB.Image imgButton 
         Height          =   465
         Index           =   7
         Left            =   60
         Picture         =   "frmMain.frx":2B62A9
         Top             =   2445
         Width           =   2280
      End
      Begin VB.Image imgButton 
         Height          =   465
         Index           =   3
         Left            =   60
         Picture         =   "frmMain.frx":2B9A25
         Top             =   1950
         Width           =   2280
      End
      Begin VB.Image imgButton 
         Height          =   465
         Index           =   10
         Left            =   60
         Picture         =   "frmMain.frx":2BD1A1
         Top             =   1455
         Width           =   2280
      End
      Begin VB.Image imgButton 
         Height          =   465
         Index           =   6
         Left            =   60
         Picture         =   "frmMain.frx":2C091D
         Top             =   960
         Width           =   2280
      End
      Begin VB.Image imgButton 
         Height          =   465
         Index           =   1
         Left            =   60
         Picture         =   "frmMain.frx":2C4099
         Top             =   465
         Width           =   2280
      End
   End
   Begin VB.PictureBox PicBlank 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2790
      Index           =   0
      Left            =   27600
      Picture         =   "frmMain.frx":2C7815
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   376
      TabIndex        =   180
      Top             =   120
      Visible         =   0   'False
      Width           =   5640
      Begin VB.PictureBox PicBlank 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   960
         Index           =   1
         Left            =   210
         Picture         =   "frmMain.frx":2FABE9
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   184
         Top             =   585
         Width           =   960
      End
      Begin VB.Label lblGym 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Batalhar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   105
         TabIndex        =   214
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   6
         Left            =   5295
         Picture         =   "frmMain.frx":2FFE4A
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblGym 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Desistir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   15
         TabIndex        =   183
         Top             =   2295
         Width           =   1335
      End
      Begin VB.Label lblGym 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Batalhar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   15
         TabIndex        =   182
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblGym 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1920
         Index           =   0
         Left            =   1515
         TabIndex        =   181
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   36240
      Picture         =   "frmMain.frx":300576
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   20
      Top             =   16080
      Visible         =   0   'False
      Width           =   2700
      Begin VB.OptionButton optWOff 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Off"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   236
         Top             =   2760
         Width           =   735
      End
      Begin VB.OptionButton optWOn 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   235
         Top             =   2520
         Width           =   735
      End
      Begin VB.PictureBox PicDesconect 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1890
         Left            =   30
         Picture         =   "frmMain.frx":31573A
         ScaleHeight     =   126
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   176
         TabIndex        =   220
         Top             =   435
         Visible         =   0   'False
         Width           =   2640
         Begin VB.Label lblDesco 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Desconectando em 5 segundos"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            TabIndex        =   221
            Top             =   405
            Width           =   2415
         End
         Begin VB.Image imgButton 
            Height          =   345
            Index           =   23
            Left            =   450
            Picture         =   "frmMain.frx":325B5E
            Top             =   1200
            Width           =   1695
         End
      End
      Begin VB.PictureBox picWasd 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   135
         Picture         =   "frmMain.frx":327A2E
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   219
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox PicSound 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   135
         Picture         =   "frmMain.frx":329916
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   106
         Top             =   870
         Width           =   240
      End
      Begin VB.PictureBox PicMusic 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   135
         Picture         =   "frmMain.frx":32B7FE
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   105
         Top             =   540
         Width           =   240
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   360
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   24
         Top             =   2400
         Width           =   735
         Begin VB.OptionButton optSOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optSOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   1560
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   21
         Top             =   2880
         Width           =   735
         Begin VB.OptionButton optMOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton optMOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   22
         Left            =   480
         Picture         =   "frmMain.frx":32D6E6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   13
         Left            =   2355
         Picture         =   "frmMain.frx":32F5B6
         Top             =   15
         Width           =   330
      End
   End
   Begin VB.PictureBox PicPokeInicial 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   33360
      Picture         =   "frmMain.frx":32FCE2
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   66
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   4
         Left            =   2070
         Picture         =   "frmMain.frx":34D3EA
         Top             =   15
         Width           =   330
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   11
         Left            =   360
         Picture         =   "frmMain.frx":34DB16
         Top             =   3105
         Width           =   1695
      End
      Begin VB.Image imgSelectPoke 
         Height          =   540
         Index           =   4
         Left            =   120
         Top             =   2280
         Width           =   540
      End
      Begin VB.Image imgSelectPoke 
         Height          =   240
         Index           =   0
         Left            =   1800
         Top             =   3840
         Width           =   900
      End
      Begin VB.Label lblPokeInicial 
         BackStyle       =   0  'Transparent
         Caption         =   "Pikachu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   750
         TabIndex        =   73
         Top             =   2475
         Width           =   2505
      End
      Begin VB.Label lblPokeInicial 
         BackStyle       =   0  'Transparent
         Caption         =   "Squirtle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   750
         TabIndex        =   72
         Top             =   1875
         Width           =   2505
      End
      Begin VB.Label lblPokeInicial 
         BackStyle       =   0  'Transparent
         Caption         =   "Chamander"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   750
         TabIndex        =   71
         Top             =   1275
         Width           =   2505
      End
      Begin VB.Label lblPokeInicial 
         BackStyle       =   0  'Transparent
         Caption         =   "Bulbasaur"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   750
         TabIndex        =   70
         Top             =   675
         Width           =   3105
      End
      Begin VB.Image imgSelectPoke 
         Height          =   540
         Index           =   3
         Left            =   120
         Top             =   1695
         Width           =   540
      End
      Begin VB.Image imgSelectPoke 
         Height          =   540
         Index           =   2
         Left            =   120
         Top             =   1110
         Width           =   540
      End
      Begin VB.Image imgSelectPoke 
         Height          =   540
         Index           =   1
         Left            =   120
         Top             =   525
         Width           =   540
      End
   End
   Begin VB.PictureBox picQuest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4215
      Left            =   30840
      Picture         =   "frmMain.frx":34F9E6
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   594
      TabIndex        =   160
      Top             =   3960
      Visible         =   0   'False
      Width           =   8910
      Begin VB.ListBox lstQuests 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3435
         IntegralHeight  =   0   'False
         Left            =   165
         TabIndex        =   161
         Top             =   570
         Width           =   2355
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   7
         Left            =   8565
         Picture         =   "frmMain.frx":3CA062
         Top             =   15
         Width           =   330
      End
      Begin VB.Image imgQuestC 
         Height          =   585
         Index           =   1
         Left            =   8655
         Top             =   3420
         Width           =   210
      End
      Begin VB.Image imgQuestC 
         Height          =   585
         Index           =   0
         Left            =   5460
         Top             =   3420
         Width           =   210
      End
      Begin VB.Label lblQuestInfo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash/Honra: 999.999"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   169
         Top             =   3795
         Width           =   2655
      End
      Begin VB.Label lblQuestInfo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dollar: 999.999.999.999"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   168
         Top             =   3585
         Width           =   2415
      End
      Begin VB.Label lblQuestInfo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Exp: 999.999.999.999"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   167
         Top             =   3375
         Width           =   2415
      End
      Begin VB.Label lblQuestInfo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tarefa Atual: 1/2"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Index           =   3
         Left            =   5760
         TabIndex        =   166
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label lblQuestInfo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Falar com Billy: 0/1"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   165
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label lblQuestInfo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":3CA78E
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Index           =   2
         Left            =   5760
         TabIndex        =   164
         Top             =   1080
         Width           =   2835
      End
      Begin VB.Label lblQuestInfo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":3CA821
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Index           =   1
         Left            =   2880
         TabIndex        =   163
         Top             =   1080
         Width           =   2505
      End
      Begin VB.Label lblQuestInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "[Missão Principal] Encontrando velho amigo do Prof.Oak"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   162
         Top             =   615
         Width           =   5895
      End
   End
   Begin VB.PictureBox PicTreinador 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   30840
      Picture         =   "frmMain.frx":3CA909
      ScaleHeight     =   152
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   258
      TabIndex        =   88
      Top             =   14640
      Visible         =   0   'False
      Width           =   3870
      Begin VB.PictureBox picFace 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1530
         Left            =   150
         ScaleHeight     =   102
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   102
         TabIndex        =   185
         Top             =   555
         Width           =   1530
      End
      Begin VB.PictureBox PicInsignia 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   8
         Left            =   3135
         Picture         =   "frmMain.frx":3E760D
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   96
         Top             =   1770
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox PicInsignia 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   2805
         Picture         =   "frmMain.frx":3E7A3F
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   95
         Top             =   1755
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox PicInsignia 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   2475
         Picture         =   "frmMain.frx":3E7E71
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   94
         Top             =   1770
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox PicInsignia 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   2145
         Picture         =   "frmMain.frx":3E82A3
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   93
         Top             =   1770
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox PicInsignia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   3135
         Picture         =   "frmMain.frx":3E86D5
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   92
         Top             =   1470
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox PicInsignia 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2805
         Picture         =   "frmMain.frx":3E8B07
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   91
         Top             =   1470
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox PicInsignia 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2475
         Picture         =   "frmMain.frx":3E8F39
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   90
         Top             =   1470
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox PicInsignia 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2145
         Picture         =   "frmMain.frx":3E936B
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   89
         Top             =   1470
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   3
         Left            =   3525
         Picture         =   "frmMain.frx":3E979D
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblTrainerCard 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pokédex: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   1800
         TabIndex        =   98
         Top             =   960
         Width           =   1965
      End
      Begin VB.Label lblTrainerCard 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Jogador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1740
         TabIndex        =   97
         Top             =   660
         Width           =   1995
      End
   End
   Begin VB.PictureBox PicOrg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4470
      Index           =   1
      Left            =   20640
      Picture         =   "frmMain.frx":3E9EC9
      ScaleHeight     =   298
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   150
      Top             =   2400
      Visible         =   0   'False
      Width           =   3375
      Begin VB.PictureBox PicOrg 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   150
         Picture         =   "frmMain.frx":41B1F5
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   205
         TabIndex        =   155
         Top             =   2865
         Width           =   3075
      End
      Begin VB.PictureBox ScrollBarFake 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   3000
         Picture         =   "frmMain.frx":422A87
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   154
         Top             =   2445
         Width           =   225
      End
      Begin VB.PictureBox ScrollBarFake 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   3000
         Picture         =   "frmMain.frx":422DC9
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   153
         Top             =   315
         Width           =   225
      End
      Begin VB.PictureBox ScrollBarFake 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   0
         Left            =   3030
         Picture         =   "frmMain.frx":42310B
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   152
         Top             =   555
         Width           =   165
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   26
         Left            =   840
         Picture         =   "frmMain.frx":4234F5
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label lblOrgInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Honra:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   151
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.PictureBox PicOrgs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   16920
      Picture         =   "frmMain.frx":4253C5
      ScaleHeight     =   363
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   149
      Top             =   2400
      Visible         =   0   'False
      Width           =   3675
      Begin VB.PictureBox faceOrg 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   210
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   226
         Top             =   615
         Width           =   480
      End
      Begin VB.PictureBox ScrollBarFake 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1890
         Index           =   5
         Left            =   3285
         Picture         =   "frmMain.frx":4667A9
         ScaleHeight     =   126
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   159
         Top             =   1860
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.PictureBox ScrollBarFake 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   3255
         Picture         =   "frmMain.frx":4679A3
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   158
         Top             =   1590
         Width           =   225
      End
      Begin VB.PictureBox ScrollBarFake 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   3255
         Picture         =   "frmMain.frx":467CE5
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   157
         Top             =   4200
         Width           =   225
      End
      Begin VB.ListBox OrgMembers 
         Appearance      =   0  'Flat
         BackColor       =   &H001D1D1D&
         ForeColor       =   &H00FFFFFF&
         Height          =   1590
         Left            =   3840
         TabIndex        =   156
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   15
         Left            =   3330
         Picture         =   "frmMain.frx":468027
         Top             =   15
         Width           =   330
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   25
         Left            =   990
         Picture         =   "frmMain.frx":468753
         Top             =   4875
         Width           =   1695
      End
      Begin VB.Label lblOrg 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   228
         Top             =   840
         Width           =   2385
      End
      Begin VB.Label lblOrg 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   227
         Top             =   600
         Width           =   2385
      End
      Begin VB.Image PicExp 
         Height          =   150
         Left            =   180
         Picture         =   "frmMain.frx":46A623
         Top             =   1275
         Width           =   3315
      End
   End
   Begin VB.PictureBox PicVip 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1395
      Index           =   2
      Left            =   34800
      Picture         =   "frmMain.frx":46C055
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   176
      Top             =   14640
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Label lblVip 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":481615
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   177
         Top             =   270
         Width           =   4455
      End
   End
   Begin VB.PictureBox PicVipPanel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3945
      Left            =   30840
      Picture         =   "frmMain.frx":4816D8
      ScaleHeight     =   263
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   488
      TabIndex        =   53
      Top             =   8280
      Visible         =   0   'False
      Width           =   7320
      Begin VB.PictureBox PicChkVipName 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   105
         Picture         =   "frmMain.frx":4DF724
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   175
         Top             =   2040
         Width           =   240
      End
      Begin VB.PictureBox PicVip 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1605
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":4DFA36
         ScaleHeight     =   107
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   440
         TabIndex        =   172
         Top             =   1020
         Visible         =   0   'False
         Width           =   6600
         Begin VB.Image VipConfCancel 
            Height          =   315
            Index           =   1
            Left            =   75
            Top             =   1125
            Width           =   1200
         End
         Begin VB.Image VipConfCancel 
            Height          =   315
            Index           =   0
            Left            =   120
            Top             =   750
            Width           =   1200
         End
         Begin VB.Label lblVip 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "VIP 6"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   174
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblVip 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ao obter o pacote de Vip 6 serão usados 1500 Pontos de Doação, Realmente deseja continuar?"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Index           =   11
            Left            =   1440
            TabIndex        =   173
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.PictureBox PicVipBar 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   825
         Picture         =   "frmMain.frx":502232
         ScaleHeight     =   150
         ScaleWidth      =   5670
         TabIndex        =   170
         Top             =   1020
         Width           =   5670
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   11
         Left            =   6975
         Picture         =   "frmMain.frx":504ED6
         Top             =   15
         Width           =   330
      End
      Begin VB.Image imgButton 
         Height          =   315
         Index           =   21
         Left            =   6165
         Picture         =   "frmMain.frx":505602
         Top             =   3420
         Width           =   990
      End
      Begin VB.Image imgButton 
         Height          =   315
         Index           =   20
         Left            =   4965
         Picture         =   "frmMain.frx":5066AE
         Top             =   3420
         Width           =   990
      End
      Begin VB.Image imgButton 
         Height          =   315
         Index           =   19
         Left            =   3765
         Picture         =   "frmMain.frx":50775A
         Top             =   3420
         Width           =   990
      End
      Begin VB.Image imgButton 
         Height          =   315
         Index           =   18
         Left            =   2565
         Picture         =   "frmMain.frx":508806
         Top             =   3420
         Width           =   990
      End
      Begin VB.Image imgButton 
         Height          =   315
         Index           =   17
         Left            =   1365
         Picture         =   "frmMain.frx":5098B2
         Top             =   3420
         Width           =   990
      End
      Begin VB.Image imgButton 
         Height          =   315
         Index           =   16
         Left            =   165
         Picture         =   "frmMain.frx":50A95E
         Top             =   3420
         Width           =   990
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   15
         Left            =   5475
         Picture         =   "frmMain.frx":50BA0A
         Top             =   1725
         Width           =   1695
      End
      Begin VB.Image cmdVip 
         Height          =   675
         Index           =   6
         Left            =   6300
         Top             =   2610
         Width           =   720
      End
      Begin VB.Image cmdVip 
         Height          =   675
         Index           =   5
         Left            =   5100
         Top             =   2610
         Width           =   720
      End
      Begin VB.Image cmdVip 
         Height          =   675
         Index           =   4
         Left            =   3900
         Top             =   2610
         Width           =   720
      End
      Begin VB.Image cmdVip 
         Height          =   675
         Index           =   3
         Left            =   2700
         Top             =   2610
         Width           =   720
      End
      Begin VB.Image cmdVip 
         Height          =   675
         Index           =   2
         Left            =   1500
         Top             =   2610
         Width           =   720
      End
      Begin VB.Image cmdVip 
         Height          =   675
         Index           =   1
         Left            =   300
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lblVip 
         BackStyle       =   0  'Transparent
         Caption         =   "99999"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   171
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.PictureBox picSSMap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   16920
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   37
      Top             =   10560
      Width           =   240
   End
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4260
      Left            =   24240
      Picture         =   "frmMain.frx":50D8DA
      ScaleHeight     =   284
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   214
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3210
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   5
         Left            =   2865
         Picture         =   "frmMain.frx":53A38E
         Top             =   15
         Width           =   330
      End
   End
   Begin VB.PictureBox PicTarget 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   3120
      Picture         =   "frmMain.frx":53AABA
      ScaleHeight     =   83
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   178
      Top             =   180
      Width           =   2670
      Begin VB.PictureBox ImgHpTarget 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   735
         Picture         =   "frmMain.frx":5458C6
         ScaleHeight     =   195
         ScaleWidth      =   1830
         TabIndex        =   194
         Top             =   585
         Width           =   1830
         Begin VB.Label lblTargetInfo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   195
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Image ElementTarget 
         Height          =   180
         Index           =   1
         Left            =   2085
         Top             =   885
         Width           =   480
      End
      Begin VB.Image ElementTarget 
         Height          =   180
         Index           =   0
         Left            =   1560
         Top             =   885
         Width           =   480
      End
      Begin VB.Label lblTargetInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   705
         TabIndex        =   179
         Top             =   285
         Width           =   1890
      End
   End
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   27600
      Picture         =   "frmMain.frx":546BBA
      ScaleHeight     =   164
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picItemDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   210
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   30
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   885
         TabIndex        =   6
         Top             =   480
         Width           =   1950
      End
      Begin VB.Label lblItemDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1110
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   2640
      End
   End
   Begin VB.PictureBox picPokedex 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5430
      Left            =   14040
      Picture         =   "frmMain.frx":5600DE
      ScaleHeight     =   362
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   370
      TabIndex        =   76
      Top             =   9960
      Visible         =   0   'False
      Width           =   5550
      Begin VB.PictureBox PicTipoDex 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   2760
         Picture         =   "frmMain.frx":5C2592
         ScaleHeight     =   180
         ScaleWidth      =   480
         TabIndex        =   216
         Top             =   1875
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox PicTipoDex 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3285
         Picture         =   "frmMain.frx":5C2A54
         ScaleHeight     =   180
         ScaleWidth      =   480
         TabIndex        =   215
         Top             =   1875
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.ListBox ListPokes 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         ForeColor       =   &H00FFFFFF&
         Height          =   4650
         IntegralHeight  =   0   'False
         Left            =   165
         TabIndex        =   78
         Top             =   570
         Width           =   2355
      End
      Begin VB.PictureBox picFD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1380
         Left            =   3345
         ScaleHeight     =   1380
         ScaleWidth      =   1410
         TabIndex        =   77
         Top             =   675
         Width           =   1410
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   12
         Left            =   5205
         Picture         =   "frmMain.frx":5C2F16
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblPI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   3750
         TabIndex        =   87
         Top             =   2325
         Width           =   1680
      End
      Begin VB.Label lblEv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evolui para:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6600
         TabIndex        =   86
         Top             =   3360
         Width           =   2715
      End
      Begin VB.Label lblPI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ataque: 0"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   3750
         TabIndex        =   85
         Top             =   3465
         Width           =   1680
      End
      Begin VB.Label lblPI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Defesa: 0"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   3750
         TabIndex        =   84
         Top             =   3840
         Width           =   1680
      End
      Begin VB.Label lblPI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EAtaque: 0"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   4
         Left            =   3750
         TabIndex        =   83
         Top             =   4200
         Width           =   1680
      End
      Begin VB.Label lblPI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EDefesa: 0"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   5
         Left            =   3720
         TabIndex        =   82
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblPI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidade: 0"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   81
         Top             =   4935
         Width           =   1695
      End
      Begin VB.Label lblPI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hp: 0"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   7
         Left            =   3750
         TabIndex        =   80
         Top             =   2715
         Width           =   1680
      End
      Begin VB.Label lblPI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PP: 0"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   8
         Left            =   3750
         TabIndex        =   79
         Top             =   3090
         Width           =   1680
      End
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6705
      Left            =   6840
      Picture         =   "frmMain.frx":5C3642
      ScaleHeight     =   447
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   540
      TabIndex        =   27
      Top             =   12480
      Visible         =   0   'False
      Width           =   8100
   End
   Begin VB.PictureBox picSelectQuest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   35880
      Picture         =   "frmMain.frx":674332
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   258
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   3870
      Begin VB.ListBox lstSelectQuest 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         ForeColor       =   &H00FFFFFF&
         Height          =   1980
         IntegralHeight  =   0   'False
         Left            =   225
         TabIndex        =   58
         Top             =   855
         Width           =   3420
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   13
         Left            =   1080
         Picture         =   "frmMain.frx":6A425E
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   1
         Left            =   3525
         Picture         =   "frmMain.frx":6A612E
         Top             =   15
         Width           =   330
      End
   End
   Begin VB.Timer tmrNoticia 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -240
      Top             =   4560
   End
   Begin VB.PictureBox picRank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5790
      Left            =   -7680
      Picture         =   "frmMain.frx":6A685A
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   608
      TabIndex        =   117
      Top             =   14640
      Visible         =   0   'False
      Width           =   9150
      Begin VB.ListBox lstRank 
         Height          =   645
         Left            =   8280
         TabIndex        =   118
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   6240
         TabIndex        =   148
         Top             =   4440
         Width           =   3195
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   6240
         TabIndex        =   147
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   6240
         TabIndex        =   146
         Top             =   1560
         Width           =   3195
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   6240
         TabIndex        =   145
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   6240
         TabIndex        =   144
         Top             =   2280
         Width           =   3195
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   6240
         TabIndex        =   143
         Top             =   2640
         Width           =   3195
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   6240
         TabIndex        =   142
         Top             =   3000
         Width           =   3195
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   6240
         TabIndex        =   141
         Top             =   3360
         Width           =   3195
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   6240
         TabIndex        =   140
         Top             =   3720
         Width           =   3195
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   6240
         TabIndex        =   139
         Top             =   4080
         Width           =   3195
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   138
         Top             =   1560
         Width           =   2355
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   137
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   136
         Top             =   1560
         Width           =   3195
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   3240
         TabIndex        =   135
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   3240
         TabIndex        =   134
         Top             =   2280
         Width           =   3195
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   3240
         TabIndex        =   133
         Top             =   2640
         Width           =   3195
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   3240
         TabIndex        =   132
         Top             =   3000
         Width           =   3195
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   3240
         TabIndex        =   131
         Top             =   3360
         Width           =   3195
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   3240
         TabIndex        =   130
         Top             =   3720
         Width           =   3195
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   3240
         TabIndex        =   129
         Top             =   4080
         Width           =   3195
      End
      Begin VB.Label lblP 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   3240
         TabIndex        =   128
         Top             =   4440
         Width           =   3195
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   127
         Top             =   4440
         Width           =   2355
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   126
         Top             =   4080
         Width           =   2355
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   125
         Top             =   3720
         Width           =   2355
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   124
         Top             =   3360
         Width           =   2355
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   123
         Top             =   3000
         Width           =   2235
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   122
         Top             =   2640
         Width           =   2355
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   121
         Top             =   2280
         Width           =   2355
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   120
         Top             =   1920
         Width           =   2355
      End
      Begin VB.Label lblRank 
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   119
         Top             =   1200
         Width           =   2355
      End
   End
   Begin VB.PictureBox PicBatalha 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   5670
      Left            =   18960
      Picture         =   "frmMain.frx":6C27ED
      ScaleHeight     =   378
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   412
      TabIndex        =   99
      Top             =   12720
      Visible         =   0   'False
      Width           =   6180
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmMain.frx":734939
         Left            =   2760
         List            =   "frmMain.frx":734943
         TabIndex        =   230
         Text            =   "Combo2"
         Top             =   6240
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMain.frx":734951
         Left            =   840
         List            =   "frmMain.frx":734958
         TabIndex        =   229
         Text            =   "Combo2"
         Top             =   6360
         Width           =   1215
      End
      Begin VB.PictureBox optEscolha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   3
         Left            =   3240
         Picture         =   "frmMain.frx":73496E
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   224
         Top             =   4395
         Width           =   240
      End
      Begin VB.PictureBox optEscolha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   3240
         Picture         =   "frmMain.frx":736856
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   223
         Top             =   4095
         Width           =   240
      End
      Begin VB.PictureBox optEscolha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   3240
         Picture         =   "frmMain.frx":73873E
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   222
         Top             =   3795
         Width           =   240
      End
      Begin VB.OptionButton optArena 
         BackColor       =   &H00808080&
         Caption         =   "6 Pokémons"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   116
         Top             =   5880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optArena 
         BackColor       =   &H00808080&
         Caption         =   "3 Pokémons"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   115
         Top             =   5880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optArena 
         BackColor       =   &H00808080&
         Caption         =   "1 Pokémon"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   114
         Top             =   5880
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox PicArena 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2520
         Index           =   1
         Left            =   240
         ScaleHeight     =   168
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   380
         TabIndex        =   102
         Top             =   645
         Width           =   5700
      End
      Begin VB.Label lblArena 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   225
         Top             =   3750
         Width           =   2925
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   24
         Left            =   2160
         Picture         =   "frmMain.frx":73A626
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   14
         Left            =   5835
         Picture         =   "frmMain.frx":73C4F6
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblArena 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   225
         TabIndex        =   113
         Top             =   4080
         Width           =   2925
      End
      Begin VB.Label lblArena 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   225
         TabIndex        =   104
         Top             =   4410
         Width           =   2925
      End
      Begin VB.Label lblArena 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Arena: 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   103
         Top             =   3315
         Width           =   5535
      End
      Begin VB.Label lblArena 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   101
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label lblArena 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   100
         Top             =   3240
         Width           =   495
      End
   End
   Begin VB.PictureBox picShop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   13800
      Picture         =   "frmMain.frx":73CC22
      ScaleHeight     =   164
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   17
      Top             =   12720
      Visible         =   0   'False
      Width           =   5100
      Begin VB.PictureBox picShopItems 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   30
         Picture         =   "frmMain.frx":7659D6
         ScaleHeight     =   130
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   210
         TabIndex        =   18
         Top             =   435
         Width           =   3150
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   18
         Left            =   4755
         Picture         =   "frmMain.frx":779B08
         Top             =   15
         Width           =   330
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   30
         Left            =   3240
         Picture         =   "frmMain.frx":77A234
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   29
         Left            =   3240
         Picture         =   "frmMain.frx":77C104
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblValor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3240
         TabIndex        =   241
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione o item no estoque!"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3240
         TabIndex        =   240
         Top             =   525
         Width           =   1695
      End
   End
   Begin VB.PictureBox picchat 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   25440
      Picture         =   "frmMain.frx":77DFD4
      ScaleHeight     =   1605
      ScaleWidth      =   5310
      TabIndex        =   74
      Top             =   8760
      Visible         =   0   'False
      Width           =   5310
      Begin VB.Label lblChat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "####"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         TabIndex        =   75
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Timer MyTargetDuel 
      Interval        =   500
      Left            =   -240
      Top             =   6960
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1800
      Left            =   -7080
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   3175
      _Version        =   393217
      BackColor       =   790032
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":799CD0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicHabilidade 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   25200
      Picture         =   "frmMain.frx":799D4C
      ScaleHeight     =   148
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   370
      TabIndex        =   107
      Top             =   10440
      Visible         =   0   'False
      Width           =   5550
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   10
         Left            =   5205
         Picture         =   "frmMain.frx":7C2070
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblDescHab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ataque 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   2760
         TabIndex        =   112
         Top             =   1770
         Width           =   2715
      End
      Begin VB.Label lblDescHab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ataque 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   111
         Top             =   1770
         Width           =   2475
      End
      Begin VB.Label lblDescHab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ataque 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   2835
         TabIndex        =   110
         Top             =   1335
         Width           =   2580
      End
      Begin VB.Label lblDescHab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ataque 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   109
         Top             =   1335
         Width           =   2475
      End
      Begin VB.Label lblDescHab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "asdasdasdasdasd"
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Index           =   0
         Left            =   270
         TabIndex        =   108
         Top             =   600
         Width           =   5025
      End
   End
   Begin VB.Timer TimerLogout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -240
      Top             =   6480
   End
   Begin VB.Timer TimerDestroyGame 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -240
      Top             =   5520
   End
   Begin VB.PictureBox picSpells 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   3240
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   13
      Top             =   13320
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.PictureBox PicMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   120
      Picture         =   "frmMain.frx":7C279C
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   67
      Top             =   8880
      Width           =   5400
      Begin VB.TextBox txtMyChat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   240
         MaxLength       =   80
         TabIndex        =   68
         Top             =   180
         Width           =   4920
      End
   End
   Begin VB.PictureBox picHotbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   13365
      Picture         =   "frmMain.frx":7CC830
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   248
      TabIndex        =   48
      Top             =   8670
      Width           =   3720
   End
   Begin VB.Timer TmrHotbar 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   -240
      Top             =   6000
   End
   Begin VB.PictureBox picUpDown 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   5250
      Picture         =   "frmMain.frx":7D5CAC
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   69
      Top             =   8520
      Width           =   270
   End
   Begin VB.PictureBox picLeilaoPainel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   5160
      Picture         =   "frmMain.frx":7D60DE
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   59
      Top             =   10080
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Timer LeilaoTempo 
         Interval        =   1000
         Left            =   3360
         Top             =   0
      End
      Begin VB.PictureBox picLeilao 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   3450
         Picture         =   "frmMain.frx":821C56
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   61
         Top             =   585
         Width           =   480
      End
      Begin VB.PictureBox picLeilao 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2625
         Index           =   2
         Left            =   90
         Picture         =   "frmMain.frx":822898
         ScaleHeight     =   175
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   208
         TabIndex        =   60
         Top             =   435
         Width           =   3120
      End
      Begin VB.Label lblLeilaoInfo 
         Caption         =   "Label1"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   254
         Top             =   120
         Width           =   2055
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   28
         Left            =   3960
         Picture         =   "frmMain.frx":83D36A
         Top             =   3090
         Width           =   1695
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   17
         Left            =   6030
         Picture         =   "frmMain.frx":83F23A
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblLeilaoInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#/#"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   1470
         TabIndex        =   234
         Top             =   3150
         Width           =   405
      End
      Begin VB.Label lblLeilaoInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   5
         Left            =   120
         TabIndex        =   233
         Top             =   3120
         Width           =   300
      End
      Begin VB.Label lblLeilaoInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   6
         Left            =   2880
         TabIndex        =   232
         Top             =   3120
         Width           =   300
      End
      Begin VB.Label lblLeilaoInfo 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   3525
         TabIndex        =   65
         Top             =   2715
         Width           =   2595
      End
      Begin VB.Label lblLeilaoInfo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   3480
         TabIndex        =   64
         Top             =   3060
         Width           =   1755
      End
      Begin VB.Label lblLeilaoInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   1170
         Index           =   2
         Left            =   3480
         TabIndex        =   63
         Top             =   1320
         Width           =   2715
      End
      Begin VB.Label lblLeilaoInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   62
         Top             =   720
         Width           =   2025
      End
   End
   Begin VB.PictureBox PicPokeSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   180
      Picture         =   "frmMain.frx":83F966
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   55
      Top             =   1500
      Width           =   2670
      Begin VB.Label lblPokeSpell 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   56
         Top             =   60
         Width           =   2445
      End
   End
   Begin VB.PictureBox picCharacter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   180
      Picture         =   "frmMain.frx":84387A
      ScaleHeight     =   83
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   4
      Top             =   180
      Width           =   2670
      Begin VB.PictureBox imgHPBar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   735
         Picture         =   "frmMain.frx":84E686
         ScaleHeight     =   195
         ScaleWidth      =   1830
         TabIndex        =   231
         Top             =   645
         Width           =   1830
         Begin VB.Label lblHP 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   255
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox imgMPBar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   735
         Picture         =   "frmMain.frx":84F97A
         ScaleHeight     =   105
         ScaleWidth      =   1830
         TabIndex        =   187
         Top             =   885
         Width           =   1830
         Begin VB.Label lblMp 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   150
            Left            =   600
            TabIndex        =   256
            Top             =   -20
            Width           =   525
         End
      End
      Begin VB.PictureBox PicPokeEquip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   54
         Top             =   270
         Width           =   480
      End
      Begin VB.Image imgEXPBar 
         Height          =   60
         Left            =   735
         Picture         =   "frmMain.frx":8503CE
         Top             =   1035
         Width           =   1830
      End
      Begin VB.Label lblPoints 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3240
         TabIndex        =   35
         Top             =   4080
         Width           =   120
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   3000
         TabIndex        =   12
         Top             =   3240
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   3240
         TabIndex        =   11
         Top             =   3120
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   3240
         TabIndex        =   10
         Top             =   3240
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   3240
         TabIndex        =   9
         Top             =   2520
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   2880
         Width           =   105
      End
      Begin VB.Label lblCharName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   615
         TabIndex        =   7
         Top             =   285
         Width           =   1950
      End
   End
   Begin VB.PictureBox PicEvolution 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   2160
      Left            =   34560
      Picture         =   "frmMain.frx":8509D2
      ScaleHeight     =   144
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   354
      TabIndex        =   50
      Top             =   12360
      Visible         =   0   'False
      Width           =   5310
      Begin VB.PictureBox PicPokeEvol 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   150
         ScaleHeight     =   96.043
         ScaleMode       =   0  'User
         ScaleWidth      =   94
         TabIndex        =   218
         Top             =   555
         Width           =   1410
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   14
         Left            =   3480
         Picture         =   "frmMain.frx":876096
         Top             =   1650
         Width           =   1695
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   9
         Left            =   4965
         Picture         =   "frmMain.frx":877F66
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblEvol 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Seu Bulbasaur está prestes a evoluir para Ivysaur deseja continuar?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   735
         Index           =   0
         Left            =   1800
         TabIndex        =   51
         Top             =   660
         Width           =   3315
      End
   End
   Begin VB.PictureBox picUpDown 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   5250
      Picture         =   "frmMain.frx":878692
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   52
      Top             =   6825
      Width           =   270
   End
   Begin VB.Timer EvolutionTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   -240
      Top             =   5040
   End
   Begin VB.PictureBox picSpellDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   3120
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   19
      Top             =   9600
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picSpellDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   34
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Elemento:"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   3240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lblSpellDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   240
         TabIndex        =   33
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblSpellName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   32
         Top             =   210
         Width           =   2805
      End
   End
   Begin VB.PictureBox picDialogue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   30840
      Picture         =   "frmMain.frx":878AC4
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   354
      TabIndex        =   38
      Top             =   17040
      Visible         =   0   'False
      Width           =   5310
      Begin VB.PictureBox lblDialogueBtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   1800
         Picture         =   "frmMain.frx":89F650
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   239
         Top             =   1680
         Width           =   1695
      End
      Begin VB.PictureBox lblDialogueBtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   840
         Picture         =   "frmMain.frx":8A1520
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   238
         Top             =   1680
         Width           =   1695
      End
      Begin VB.PictureBox lblDialogueBtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   2650
         Picture         =   "frmMain.frx":8A33F0
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   237
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblDialogue_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Robin has requested a trade. Would you like to accept?"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label lblDialogue_Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Request"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   440
         Width           =   4815
      End
   End
   Begin VB.PictureBox picCurrency 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   19920
      Picture         =   "frmMain.frx":8A52C0
      ScaleHeight     =   148
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   4125
      Begin VB.TextBox txtCurrency 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   240
         TabIndex        =   16
         Top             =   1305
         Width           =   3615
      End
      Begin VB.Image imgButton 
         Height          =   345
         Index           =   27
         Left            =   1200
         Picture         =   "frmMain.frx":8C31B4
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Image imgClose 
         Height          =   390
         Index           =   16
         Left            =   3780
         Picture         =   "frmMain.frx":8C5084
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblCurrency 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "How many do you want to drop?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.PictureBox picTempInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   13080
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   3
      Top             =   10560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   13680
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   28
      Top             =   10560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   14280
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   36
      Top             =   10560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picCover 
      Appearance      =   0  'Flat
      BackColor       =   &H00181C21&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   16920
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   31
      Top             =   10080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picParty 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   120
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   41
      Top             =   9600
      Visible         =   0   'False
      Width           =   2910
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   3075
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   2940
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2340
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2205
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1620
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1485
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   870
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   735
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Label lblPartyLeave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   47
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyInvite 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   45
         Top             =   2670
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   44
         Top             =   1935
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   42
         Top             =   465
         Width           =   2415
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00181C21&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Width           =   480
      Begin MSWinsockLib.Winsock Socket 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************

Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long
Public SegToQuit As Long
Private ObterVipNumber As Byte

Private Declare Function ShellExecute Lib "shell32.dll" _
                                      Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
                                                             ByVal lpFile As String, ByVal lpParameters As String, _
                                                             ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdVip_Click(Index As Integer)

    If PicVip(1).Visible = True Then Exit Sub
    PicVip(1).Visible = True
    Select Case Index
    Case 1
        lblVip(12).Caption = "VIP 1"
        lblVip(11).Caption = "Ao obter o pacote de Vip 1 serão usados 200 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 1
    Case 2
        lblVip(12).Caption = "VIP 2"
        lblVip(11).Caption = "Ao obter o pacote de Vip 2 serão usados 400 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 2
    Case 3
        lblVip(12).Caption = "VIP 3"
        lblVip(11).Caption = "Ao obter o pacote de Vip 3 serão usados 600 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 3
    Case 4
        lblVip(12).Caption = "VIP 4"
        lblVip(11).Caption = "Ao obter o pacote de Vip 4 serão usados 900 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 4
    Case 5
        lblVip(12).Caption = "VIP 5"
        lblVip(11).Caption = "Ao obter o pacote de Vip 5 serão usados 1200 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 5
    Case 6
        lblVip(12).Caption = "VIP 6"
        lblVip(11).Caption = "Ao obter o pacote de Vip 6 serão usados 1500 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 6
    End Select

End Sub

Private Sub cmdVip_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index > 0 And Index < 7 Then

        X = TwipsToPixels(X, 0)
        Y = TwipsToPixels(Y, 1)
        PicVip(2).ZOrder 0
        PicVip(2).Visible = True
        PicVip(2).Left = X + PicVipPanel.Left + cmdVip(Index).Left - (PicVip(2).Width / 2)
        PicVip(2).top = Y + PicVipPanel.top + cmdVip(Index).top - PicVip(2).Height - 2

        Select Case Index
        Case 1
            lblVip(0).Caption = "1 Mês #VIP1, 20 Pokéballs, 15 Greatballs, 10 Ultraballs, 5 Premierballs e Báu Surpresa. Vantagem: Áreas Vip."
        Case 2
            lblVip(0).Caption = "1 Mês #VIP2, 1 Ovo da Fortuna, 1 Báu Visual, 2 Báu Surpresa. Vantagem: Áreas Vip, Desconto PokeMart 15%, Acelerar incubadora em 30Min, 1.5x Exp."
        Case 3
            lblVip(0).Caption = "1 Mês #VIP3, 1 Ovo da Fortuna, 1 Báu Visual, 3 Báu Surpresa, Ticket Bicicleta. Vantagem: Áreas Vip, Desconto PokeMart 20%, Acelerar incubadora em 1Hr, Superfaturar 15%, 2x Exp."
        Case 4
            lblVip(0).Caption = "1 Mês #VIP4 2 Ovo da Fortuna, 1 Báus Visuais, 3 Báu Surpresa, Ticket Bicicleta. Vantagem: Áreas Vip, Desconto PokeMart 20%, Acelerar incubadora em 1h30m, Superfaturar 20%, 2x Exp."
        Case 5
            lblVip(0).Caption = "1 Mês #VIP5 2 Ovo da Fortuna, 1 Báus Visuais, 3 Báu Surpresa, Ticket Bicicleta, 1 MasterBall. Vantagem: Áreas Vip, Desconto PokeMart 25%, Acelerar incubadora em 2Hr, Superfaturar 25%, 2x Exp."
        Case 6
            lblVip(0).Caption = "1 Mês #VIP6 3 Ovo da Fortuna, 2 Báus Visuais,5 Báu Surpresa, Ticket Bicicleta, 2 MasterBall, Todas Vantagens Vip Anteriores."""
        End Select
    Else
        PicVip(2).Visible = False
    End If
End Sub

Private Sub imgQuestC_Click(Index As Integer)
    Dim QuestNum As Integer

    If lstQuests.ListCount = 0 Then
        Call UpdateQuestInfo(0)
        Exit Sub
    End If

    QuestNum = GetQuestNum(Trim$(lstQuests.text))

    If Quest(QuestNum).ItemRew(6) = 0 Then Exit Sub

    If RewardsPage = 1 Then
        RewardsPage = 0
    Else
        RewardsPage = 1
    End If

    BltQuestRewards
End Sub

Private Sub imgQuestC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.picItemDesc.Visible = False
    picPokeDesc.Visible = False
    LastBankDesc = 0
    LastItemPoke = 0
End Sub

Private Sub imgSelectPoke_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim X2 As Long, Y2 As Long

    X2 = TwipsToPixels(X, 0) + PicPokeInicial.Left + imgSelectPoke(Index).Left - picPokeDesc.Width - 25
    Y2 = TwipsToPixels(Y, 1) + PicPokeInicial.top + imgSelectPoke(Index).top - picPokeDesc.Height / 2
    Select Case Index
    Case 1
        UpdatePokeWindow 1, X2, Y2, 5
        LastItemDesc = 0
        LastItemPoke = 1
    Case 2
        UpdatePokeWindow 4, X2, Y2, 5
        LastItemDesc = 0
        LastItemPoke = 4
    Case 3
        UpdatePokeWindow 7, X2, Y2, 5
        LastItemDesc = 0
        LastItemPoke = 7
    Case 4
        UpdatePokeWindow 152, X2, Y2, 5
        LastItemDesc = 0
        LastItemPoke = 25
    End Select

    picItemDesc.Visible = False
    Exit Sub
End Sub


Private Sub lblDialogueBtn_Click(Index As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' call the handler
    dialogueHandler Index

    picDialogue.Visible = False
    dialogueIndex = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblDialogueBtn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblGym_Click(Index As Integer)
    If Index = 1 Then
        Select Case ChatGym
        Case 1    'Brock
            If ChatGymStep = 0 Then
                ChatGym = 1
                ChatGymStep = 1
                frmMain.lblGym(0) = "Isso é uma honra vindo do(a) treinador(a) que me obriga a desafia-lo(a). Muito bem então! Mostre-me o seu Melhor!"
                frmMain.lblGym(1) = "Eu Vou!"
                frmMain.lblGym(2) = "Cancelar"
            ElseIf ChatGymStep = 1 Then
                ChatGym = 0
                ChatGymStep = 0
                frmMain.PicBlank(0).Visible = False
                SendComandoGym 1
            End If
        End Select
    ElseIf Index = 2 Then
        ChatGym = 0
        ChatGymStep = 0
        frmMain.PicBlank(0).Visible = False
    End If
End Sub

Private Sub lblOrgInfo_Click(Index As Integer)
    If Index = 4 Then
        OrgPage = 1
        DragOrgShopNum = 0

        Select Case OrgPage
        Case 1: ScrollBarFake(0).top = 37
        Case 2: ScrollBarFake(0).top = 63
        Case 3: ScrollBarFake(0).top = 89
        Case 4: ScrollBarFake(0).top = 115
        Case 5: ScrollBarFake(0).top = 137
        End Select
    End If
End Sub

Private Sub lblPokeInicial_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picPokeDesc.Visible = False
    LastItemDesc = 0
    LastItemPoke = 0
End Sub

Private Sub lblVip_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PicVip(2).Visible = True Then PicVip(2).Visible = False
End Sub

Private Sub lstQuests_Click()
    Dim QuestNum As Long

    If lstQuests.ListCount = 0 Then
        Call UpdateQuestInfo(0)
        Exit Sub
    End If

    QuestNum = GetQuestNum(Trim$(lstQuests.text))
    Call UpdateQuestInfo(QuestNum)
    RewardsPage = 0
    BltQuestRewards
End Sub

Private Sub optEscolha_Click(Index As Integer)

    optArena(0).value = True

    Select Case Index

    Case 1
        optArena(0).value = True
        optEscolha(1).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")
        optEscolha(2).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
        optEscolha(3).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")

    Case 2
        optArena(1).value = True
        optEscolha(2).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")
        optEscolha(1).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
        optEscolha(3).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")

    Case 3
        optArena(2).value = True
        optEscolha(3).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")
        optEscolha(1).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
        optEscolha(2).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")

    End Select
End Sub

Private Sub PicChkVipName_Click()
    If Player(MyIndex).VipInName = True Then
        SendObterPacVip 0, 0, 1
    Else
        SendObterPacVip 0, 0, 2
    End If
End Sub

Private Sub PicMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PicVip(2).Visible = True Then PicVip(2).Visible = False
End Sub

Private Sub EvolutionTimer_Timer()
    EvolTick = EvolTick + 1

    If InGame = False Then
        EvolTick = 0
        EvolutionTimer.Enabled = False
        PicEvolution.Visible = False
        imgClose(9).Visible = True
        imgButton(14).Visible = True
        Exit Sub
    End If

    If EvolTick >= 18 Then
        EvolTick = 0
        EvolutionTimer.Enabled = False
        PicEvolution.Visible = False
        Player(MyIndex).EvolPermition = 0
        imgClose(9).Visible = True
        imgButton(14).Visible = True
    End If

    bltPokeEvolvePortrait

End Sub

Private Sub Form_Load()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Me.Width = PixelsToTwips(1152)
    Me.Height = PixelsToTwips(668)

    picScreen.Height = 640
    picScreen.Width = 1152

    picScreen.ScaleHeight = 640
    picScreen.ScaleWidth = 1152

    PicLoading.Height = 640
    PicLoading.Width = 1152

    picCover.top = picScreen.top - 1
    picCover.Left = picScreen.Left - 1
    picCover.Height = picScreen.Height + 2
    picCover.Width = picScreen.Width + 2
    PageLeilao = 1
    PageMaxLeilao = 1
    Arena = 1
    TmrHotbar = False
    OrgPage = 1
    OrgPagMem = 1

    PicBlank(0).Left = (TwipsToPixels(frmMain.Width, 0) / 2) - (PicBlank(0).Width / 2)
    PicBlank(0).top = (TwipsToPixels(frmMain.Height, 0) / 2) - (PicBlank(0).Height / 2)

    If Options.Music = 1 Then
        PicMusic.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\On.jpg")
    Else
        PicMusic.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
    End If

    If Options.sound = 1 Then
        PicSound.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\On.jpg")
    Else
        PicSound.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
    End If

    'PicTarget Bar
    ImgHpTarget.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\health.bmp")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function PixelsToTwips(pixels As Integer)
    PixelsToTwips = pixels * Screen.TwipsPerPixelX
End Function

Private Sub Form_Unload(Cancel As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Cancel = True
    logoutGame

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False

    ' reset all buttons
    resetButtons_Main
    resetButtons_Close

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgAcceptTrade_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    AcceptTrade

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgAcceptTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_Click(Index As Integer)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case Index

        ' Inventário
    Case 1
        picInventory.ZOrder 0
        picInventory.Visible = Not picInventory.Visible
        frmMain.picInventory.top = (frmMain.ScaleHeight / 2) - (frmMain.picInventory.Height / 2)
        frmMain.picInventory.Left = (frmMain.ScaleWidth / 2) - (frmMain.picInventory.Width / 2)
        BltInventory
        PlaySound Sound_ButtonClick, -1, -1
    Case 2
        If picLeilaoPainel.Visible = False Then
            picLeilaoPainel.ZOrder 0
            picLeilaoPainel.Visible = True
            frmMain.picLeilaoPainel.top = (frmMain.ScaleHeight / 2) - (frmMain.picLeilaoPainel.Height / 2)
            frmMain.picLeilaoPainel.Left = (frmMain.ScaleWidth / 2) - (frmMain.picLeilaoPainel.Width / 2)
            SendALeilao
            BltLeilao
        Else
            picLeilaoPainel.Visible = False
            SendALeilao
            BltLeilao
        End If

        ' play sound
        PlaySound Sound_ButtonClick, -1, -1

    Case 3
        If lstQuests.ListCount > 0 Then
            lstQuests.ListIndex = 0
        Else
            UpdateQuestInfo 0
        End If

        BltQuestRewards
        picQuest.ZOrder 0
        picQuest.ZOrder 0
        picQuest.Visible = Not picQuest.Visible
        frmMain.picQuest.top = (frmMain.ScaleHeight / 2) - (frmMain.picQuest.Height / 2)
        frmMain.picQuest.Left = (frmMain.ScaleWidth / 2) - (frmMain.picQuest.Width / 2)

        ' play sound
        PlaySound Sound_ButtonClick, -1, -1

    Case 4
        ' show the window
        picOptions.ZOrder 0
        picOptions.Visible = Not picOptions.Visible
        frmMain.picOptions.top = (frmMain.ScaleHeight / 2) - (frmMain.picOptions.Height / 2)
        frmMain.picOptions.Left = (frmMain.ScaleWidth / 2) - (frmMain.picOptions.Width / 2)

        ' play sound
        PlaySound Sound_ButtonClick, -1, -1

    Case 5
        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
            SendTradeRequest
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        Else
            AddText "Invalid trade target.", BrightRed
        End If

        ' play sound
        PlaySound Sound_ButtonClick, -1, -1
    Case 6
        If HasItem(MyIndex, 22) Then
            picPokedex.ZOrder 0
            picPokedex.Visible = Not picPokedex.Visible
            ListPokes.ListIndex = 0
            frmMain.picPokedex.top = (frmMain.ScaleHeight / 2) - (frmMain.picPokedex.Height / 2)
            frmMain.picPokedex.Left = (frmMain.ScaleWidth / 2) - (frmMain.picPokedex.Width / 2)

        Else
            AddText "Você ainda não tem uma Pokédex! Vá falar com o Prof.Oak!", BrightRed
        End If

        ' play sound
        PlaySound Sound_ButtonClick, -1, -1

    Case 7    'Private Msg
        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
            SendChatComando 3, GetPlayerName(myTarget)
            PlaySound Sound_ButtonClick, -1, -1
            AddText "Convite enviado com sucesso!", White
        Else
            AddText "Escolha um Jogador!", BrightRed
        End If

        ' play sound
        PlaySound Sound_ButtonClick, -1, -1

    Case 8    'Vip
        PicVipPanel.ZOrder 0
        PicVipPanel.Visible = Not PicVipPanel.Visible
        frmMain.PicVipPanel.top = (frmMain.ScaleHeight / 2) - (frmMain.PicVipPanel.Height / 2)
        frmMain.PicVipPanel.Left = (frmMain.ScaleWidth / 2) - (frmMain.PicVipPanel.Width / 2)
        PlaySound Sound_ButtonClick, -1, -1

    Case 9    'Organização
        If Player(MyIndex).ORG > 0 Then
            SendAbrir
            PicOrgs.ZOrder 0
            PicOrgs.Visible = Not PicOrgs.Visible
            PicOrg(1).Visible = False
            frmMain.PicOrgs.top = (frmMain.ScaleHeight / 2) - (frmMain.PicOrgs.Height / 2)
            frmMain.PicOrgs.Left = (frmMain.ScaleWidth / 2) - (frmMain.PicOrgs.Width / 2)
            ScrollBarMembers
            BltOrganização
        Else
            Call AddText("Você não está em uma organização!", BrightRed)
        End If

        ' play sound
        PlaySound Sound_ButtonClick, -1, -1

    Case 10
        PicTreinador.ZOrder 0
        PicTreinador.Visible = Not PicTreinador.Visible
        frmMain.PicTreinador.top = (frmMain.ScaleHeight / 2) - (frmMain.PicTreinador.Height / 2)
        frmMain.PicTreinador.Left = (frmMain.ScaleWidth / 2) - (frmMain.PicTreinador.Width / 2)
        CarregarTrainerCard
        BltFace
        PlaySound Sound_ButtonClick, -1, -1

    Case 11
        SendSelectPokeInicial
        PicPokeInicial.Visible = False

    Case 12
        SendSurfInit 1
        frmMain.PicSurf.Visible = True
        frmMain.PicSurf.top = (frmMain.ScaleHeight / 2) - (frmMain.PicSurf.Height / 2)
        frmMain.PicSurf.Left = (frmMain.ScaleWidth / 2) - (frmMain.PicSurf.Width / 2)

    Case 13
        SendQuestCommand 2, lstSelectQuest.ListIndex + 1
        picSelectQuest.Visible = False

    Case 14
        If Player(MyIndex).EvolPermition = 0 Then Exit Sub
        If frmMain.EvolutionTimer = False Then
            Call SendEvolCommand(0)
            imgClose(9).Visible = False
            imgButton(14).Visible = False
            frmMain.EvolutionTimer = True
            PlaySound Sound_Evolve, -1, -1
        End If

    Case 15
        Call ShellExecute(0, "open", "https://www.facebook.com/pokemonoriginseternal", 0, 0, 1)

    Case 16
        If PicVip(1).Visible = True Then Exit Sub
        PicVip(1).Visible = True
        lblVip(12).Caption = "VIP 1"
        lblVip(11).Caption = "Ao obter o pacote de Vip 1 serão usados 200 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 1

    Case 17
        If PicVip(1).Visible = True Then Exit Sub
        PicVip(1).Visible = True
        lblVip(12).Caption = "VIP 2"
        lblVip(11).Caption = "Ao obter o pacote de Vip 2 serão usados 400 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 2

    Case 18
        If PicVip(1).Visible = True Then Exit Sub
        PicVip(1).Visible = True
        lblVip(12).Caption = "VIP 3"
        lblVip(11).Caption = "Ao obter o pacote de Vip 3 serão usados 600 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 3

    Case 19
        If PicVip(1).Visible = True Then Exit Sub
        PicVip(1).Visible = True
        lblVip(12).Caption = "VIP 4"
        lblVip(11).Caption = "Ao obter o pacote de Vip 4 serão usados 900 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 4

    Case 20
        If PicVip(1).Visible = True Then Exit Sub
        PicVip(1).Visible = True
        lblVip(12).Caption = "VIP 5"
        lblVip(11).Caption = "Ao obter o pacote de Vip 5 serão usados 1200 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 5

    Case 21
        If PicVip(1).Visible = True Then Exit Sub
        PicVip(1).Visible = True
        lblVip(12).Caption = "VIP 6"
        lblVip(11).Caption = "Ao obter o pacote de Vip 6 serão usados 1500 Pontos de Doação, Realmente deseja continuar?"
        ObterVipNumber = 6

    Case 22
        If TimerDestroyGame.Enabled = True Then Exit Sub
        If TimerLogout.Enabled = True Then Exit Sub

        If InGame Then
            If SegToQuit = 0 Then
                SegToQuit = 5
            End If
            TimerDestroyGame.Enabled = True
            PicDesconect.Visible = True
            lblDesco.Caption = "Desconectando em 5 segundos"
            PicDesconect.ZOrder vbBringToFront
        End If

    Case 23
        If TimerDestroyGame.Enabled = True Then
            TimerDestroyGame.Enabled = False
            SegToQuit = 5
        End If
        If TimerLogout.Enabled = True Then
            TimerLogout.Enabled = False
            SegToQuit = 5
        End If

        PicDesconect.Visible = False

    Case 24
        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
            SendLutarComando 1, 0, Arena, PokQntiaArena, GetPlayerName(myTarget)
            PicBatalha.Visible = False
        Else
            AddText "Escolha um alvo valido", BrightRed
        End If

    Case 25
        BltOrgShop
        BltItemSelectOrgShop
        PicOrg(1).Visible = Not PicOrg(1).Visible

        PicOrg(1).Left = (PicOrgs.Width + PicOrgs.Left) + 10
        PicOrg(1).top = PicOrgs.top

        lblOrgInfo(3).Caption = "Honra: " & GetPlayerHonra(MyIndex)

    Case 26
        If DragOrgShopNum = 0 Then Exit Sub

        If Not Item(OrgShop(DragOrgShopNum).Item).Type = ITEM_TYPE_CURRENCY Then
            SendBuyOrgShop DragOrgShopNum, 1
        Else
            CurrencyMenu = 5    ' BuyOrgShop
            lblCurrency.Caption = "Qual a quantia que você deseja comprar?"
            txtCurrency.text = vbNullString
            picCurrency.Visible = True
            picCurrency.ZOrder 0
            frmMain.picCurrency.top = (frmMain.ScaleHeight / 2) - (frmMain.picCurrency.Height / 2)
            frmMain.picCurrency.Left = (frmMain.ScaleWidth / 2) - (frmMain.picCurrency.Width / 2)
            txtCurrency.SetFocus
        End If

    Case 27
        If IsNumeric(txtCurrency.text) Then
            Select Case CurrencyMenu
            Case 1    ' drop item
                If Val(txtCurrency.text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                SendDropItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 2    ' deposit item
                If Val(txtCurrency.text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                DepositItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 3    ' withdraw item
                WithdrawItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 4    ' offer trade item
                If Val(txtCurrency.text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                TradeItem tmpCurrencyItem, Val(txtCurrency.text)
            Case 5    'Comprar Quantia OrgShop
                If Val(txtCurrency.text) * OrgShop(DragOrgShopNum).Valor > GetPlayerHonra(MyIndex) Then
                    AddText "Você não possui honra o suficiente para comprar este item", BrightRed
                Else
                    SendBuyOrgShop DragOrgShopNum, Val(txtCurrency.text)
                End If
            End Select
        Else
            AddText "Por favor insira um valor válido.", BrightRed
            Exit Sub
        End If

        picCurrency.Visible = False
        tmpCurrencyItem = 0
        txtCurrency.text = vbNullString
        CurrencyMenu = 0    ' clear

    Case 28
        If LeilaoItemSelect = 0 Then Exit Sub

        If Leilao(LeilaoItemSelect).Vendedor <> vbNullString Then
            If Leilao(LeilaoItemSelect).ItemNum > 0 Then
                SendComprar LeilaoItemSelect
            End If
        End If

        LeilaoItemSelect = 0
        lblLeilaoInfo(1).Caption = vbNullString
        lblLeilaoInfo(2).Caption = vbNullString
        lblLeilaoInfo(3).Caption = vbNullString
        lblLeilaoInfo(4).Caption = vbNullString
        BltPokeLeilaoSelect

    Case 29
        If ShopAction = 1 Then Exit Sub
        ShopAction = 1    ' buying an item
        AddText "Agora clique no item para comprar.", White

    Case 30
        If ShopAction = 2 Then Exit Sub
        ShopAction = 2    ' selling an item
        AddText "Clique duas vezes no item para vender.", White

    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset other buttons
    resetButtons_Main Index

    ' change the button we're hovering on
    If Not MainButton(Index).state = 2 Then    ' make sure we're not clicking
        changeButtonState_Main Index, 1    ' hover
    End If

    ' play sound
    If Not LastButtonSound_Main = Index Then
        PlaySound Sound_ButtonHover, -1, -1
        LastButtonSound_Main = Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset all buttons
    resetButtons_Main -1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset other buttons
    resetButtons_Main Index

    ' change the button we're hovering on
    changeButtonState_Main Index, 2    ' clicked

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgQuest_Click(Index As Integer)
    Dim Slot As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Declarations
    Slot = lstQuests.ListIndex + 1

    ' Prevent subscript out range
    If Slot <= 0 Then Exit Sub
    If Player(Index).Quests(Slot).status = QUEST_STATUS_END Then Exit Sub

    ' play sound
    PlaySound Sound_ButtonClick, -1, -1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgQuest_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgSelectPoke_Click(Index As Integer)
    Dim i As Long

    Select Case Index
    Case 0
        SendSelectPokeInicial
        PicPokeInicial.Visible = False
    Case Else
        SelectPokeInicial = Index

        Select Case Index
        Case 1
            PlaySound "PokeSounds\001.mp3", -1, -1
        Case 2
            PlaySound "PokeSounds\004.mp3", -1, -1
        Case 3
            PlaySound "PokeSounds\007.mp3", -1, -1
        Case 4
            PlaySound "PokeSounds\025.mp3", -1, -1

        End Select

        lblPokeInicial(Index).ForeColor = &HFF00&

        For i = 1 To 4
            If i <> Index Then
                lblPokeInicial(i).ForeColor = &HFFFFFF
            End If
        Next

    End Select

End Sub

Private Sub lblArena_Click(Index As Integer)

    Select Case Index

    Case 1

    Case 4
        If Arena = 10 Then Arena = 1 Else Arena = Arena + 1
        If Player(MyIndex).Arena(Arena) = 1 Then
            lblArena(3).Caption = "Arena: " & Arena & " - (Ocupada)"
            lblArena(3).ForeColor = QBColor(BrightRed)
        Else
            lblArena(3).Caption = "Arena: " & Arena & " - (Livre)"
            lblArena(3).ForeColor = &HFF00&
        End If
        PicArena(1).Picture = LoadPicture(App.Path & "\data files\graphics\arenas\" & Arena & ".bmp")

    Case 5
        If Arena = 1 Then Arena = 10 Else Arena = Arena - 1
        If Arena <= 0 Then Arena = 1
        If Player(MyIndex).Arena(Arena) = 1 Then
            lblArena(3).Caption = "Arena: " & Arena & " - (Ocupada)"
            lblArena(3).ForeColor = QBColor(BrightRed)
        Else
            lblArena(3).Caption = "Arena: " & Arena & " - (Livre)"
            lblArena(3).ForeColor = &HFF00&
        End If
        PicArena(1).Picture = LoadPicture(App.Path & "\data files\graphics\arenas\" & Arena & ".bmp")


    End Select
End Sub

Private Sub imgDeclineTrade_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DeclineTrade

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgDeclineTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblDescHab_Click(Index As Integer)
    Select Case Index
    Case 1 To 4
        SendAprenderHab Index
        frmMain.PicHabilidade.Visible = False
    End Select
End Sub

Private Sub lblLeilaoInfo_Click(Index As Integer)
    Select Case Index
    Case 5
        If PageLeilao > 1 Then
            PageLeilao = PageLeilao - 1
            lblLeilaoInfo(0).Caption = PageLeilao & "/" & PageMaxLeilao
            BltLeilao
        End If
    Case 6
        If PageLeilao < PageMaxLeilao Then
            PageLeilao = PageLeilao + 1
            lblLeilaoInfo(0).Caption = PageLeilao & "/" & PageMaxLeilao
            BltLeilao
        End If
    End Select
End Sub

Private Sub lblPartyInvite_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
        SendPartyRequest
    Else
        AddText "Invalid invitation target.", BrightRed
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblPartyLeave_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Party.Leader > 0 Then
        SendPartyLeave
    Else
        AddText "You are not in a party.", BrightRed
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblPokeInicial_Click(Index As Integer)
    Dim i As Long

    SelectPokeInicial = Index

    Select Case Index
    Case 1
        PlaySound "PokeSounds\001.mp3", -1, -1
    Case 2
        PlaySound "PokeSounds\004.mp3", -1, -1
    Case 3
        PlaySound "PokeSounds\007.mp3", -1, -1
    Case 4
        PlaySound "PokeSounds\025.mp3", -1, -1
    End Select

    lblPokeInicial(Index).ForeColor = &HFF00&

    For i = 1 To 4
        If i <> Index Then
            lblPokeInicial(i).ForeColor = &HFFFFFF
        End If
    Next

End Sub

Private Sub lblTrainStat_Click(Index As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
    SendTrainStat Index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblTrainStat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub LeilaoTempo_Timer()
    Dim i As Long

    For i = 1 To MAX_LEILAO
        If Leilao(i).Tempo > 1 Then
            Leilao(i).Tempo = Leilao(i).Tempo - 1
        End If
    Next
    
    If LeilaoItemSelect > 0 Then
        If Leilao(LeilaoItemSelect).ItemNum = 0 Then
        frmMain.lblLeilaoInfo(7).Caption = vbNullString
    Else
        frmMain.lblLeilaoInfo(7).Caption = "V: " & Trim$(Leilao(LeilaoItemSelect).Vendedor) & " Tempo: " & Leilao(LeilaoItemSelect).Tempo
        End If
    End If

End Sub

Private Sub ListPokes_Click()
    Dim i As Long '123

    If Player(MyIndex).Pokedex(ListPokes.ListIndex + 1) = 1 Then
        lblPI(0).Caption = Trim$(Pokemon(ListPokes.ListIndex + 1).Name)
        lblPI(1).Caption = Trim$(Pokemon(ListPokes.ListIndex + 1).Desc)
        lblPI(2).Caption = Pokemon(ListPokes.ListIndex + 1).Add_Stat(1)
        lblPI(3).Caption = Pokemon(ListPokes.ListIndex + 1).Add_Stat(2)
        lblPI(4).Caption = Pokemon(ListPokes.ListIndex + 1).Add_Stat(3)
        lblPI(5).Caption = Pokemon(ListPokes.ListIndex + 1).Add_Stat(4)
        lblPI(6).Caption = Pokemon(ListPokes.ListIndex + 1).Add_Stat(5)
        lblPI(7).Caption = Pokemon(ListPokes.ListIndex + 1).Vital(1)
        lblPI(8).Caption = Pokemon(ListPokes.ListIndex + 1).Vital(2)

        picFD.Visible = True
        If Pokemon(ListPokes.ListIndex + 1).Evolução(1).Pokemon > 0 Then
            lblEv.Caption = "Evolui para: " & Trim$(Pokemon(Pokemon(ListPokes.ListIndex + 1).Evolução(1).Pokemon).Name)
        Else
            lblEv.Caption = "Não Evolui"
        End If

        BltPoke
    Else
        lblPI(0).Caption = "???"
        lblPI(1).Caption = "???"
        lblPI(2).Caption = "???"
        lblPI(3).Caption = "???"
        lblPI(4).Caption = "???"
        lblPI(5).Caption = "???"
        lblPI(6).Caption = "???"
        lblPI(7).Caption = "???"
        lblPI(8).Caption = "???"

        lblEv.Caption = "???"
        picFD.Visible = False
        BltPoke
    End If

End Sub

Private Sub MyTargetDuel_Timer()
    If myTarget = 0 Then
        lblArena(0).Caption = "Jogador"
        lblArena(6).Caption = "Vitorias: 0"
        lblArena(7).Caption = "Derrotas: 0"
    Else
        lblArena(0).Caption = GetPlayerName(myTarget)
        lblArena(6).Caption = "Vitorias: " & Player(myTarget).Vitorias
        lblArena(7).Caption = "Derrotas: " & Player(myTarget).Derrotas
    End If
End Sub

Private Sub optArena_Click(Index As Integer)
    PokQntiaArena = Index    'optArena
End Sub

Private Sub optMOff_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Music = 0
    ' stop music playing
    StopMusic
    ' save to config.ini
    SaveOptions

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMOn_Click()
    Dim MusicFile As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Music = 1
    ' start music playing
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        PlayMusic MusicFile
    Else
        StopMusic
    End If
    ' save to config.ini
    SaveOptions

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOff_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.sound = 0
    ' save to config.ini
    SaveOptions

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOn_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.sound = 1
    ' save to config.ini
    SaveOptions

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optWOff_click()

    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.wasd = 0
    SaveOptions

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optWOff_click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optWOn_click()

    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.wasd = 1
    SaveOptions

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optWOn_click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub picCover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False

    ' reset all buttons
    resetButtons_Main
    resetButtons_Close

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCover_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picHotbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(X, Y)

    If Button = 1 Then
        If SlotNum <> 0 Then
            SendHotbarUse SlotNum
        End If
    ElseIf Button = 2 Then
        If SlotNum <> 0 Then
            SendHotbarChange 0, 0, SlotNum
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picHotbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SlotNum As Long, InvNum As Long, QuestNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(X, Y)

    If SlotNum <> 0 Then
        If Hotbar(SlotNum).sType = 1 Then    ' item

            If Hotbar(SlotNum).Pokemon = 0 Then
                X = X + picHotbar.Left + 1
                Y = Y + picHotbar.top - picItemDesc.Height - 1
                UpdateDescWindow Hotbar(SlotNum).Slot, X, Y
                LastItemDesc = Hotbar(SlotNum).Slot    ' set it so you don't re-set values
            End If

            Exit Sub
        ElseIf Hotbar(SlotNum).sType = 2 Then    ' spell
            X = X + picHotbar.Left + 1
            Y = Y + picHotbar.top - picSpellDesc.Height - 1
            UpdateSpellWindow Hotbar(SlotNum).Slot, X, Y
            LastSpellDesc = Hotbar(SlotNum).Slot  ' set it so you don't re-set values
            Exit Sub
        End If
    End If


    picItemDesc.Visible = False
    LastItemDesc = 0    ' no item was last loaded
    picSpellDesc.Visible = False
    LastSpellDesc = 0    ' no spell was last loaded

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picLeilao_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LeilaoSlot As Long
    LeilaoSlot = IsLeilaoItem(X, Y)

    If Index = 2 Then
        If LeilaoSlot > 0 Then

            LeilaoItemSelect = LeilaoSlot
            BltPokeLeilaoSelect

            If Leilao(LeilaoSlot).ItemNum = 0 Then
                LeilaoItemSelect = 0
                lblLeilaoInfo(1).Caption = vbNullString
                lblLeilaoInfo(2).Caption = vbNullString
                lblLeilaoInfo(3).Caption = vbNullString
                lblLeilaoInfo(4).Caption = vbNullString
                BltPokeLeilaoSelect
                Exit Sub
            End If

            If Leilao(LeilaoSlot).Poke.Pokemon > 0 Then
                lblLeilaoInfo(1).Caption = Trim$(Leilao(LeilaoSlot).Vendedor) & " - " & Trim$(Pokemon(Leilao(LeilaoSlot).Poke.Pokemon).Name)
                lblLeilaoInfo(2).Caption = Trim$(Pokemon(Leilao(LeilaoSlot).Poke.Pokemon).Desc)
                If Leilao(LeilaoSlot).Tipo = 1 Then
                    lblLeilaoInfo(4).Caption = "Valor: " & Leilao(LeilaoSlot).Price & " Zeny(s)"
                Else
                    lblLeilaoInfo(4).Caption = "Valor: " & Leilao(LeilaoSlot).Price & " P.C"
                End If
                    lblLeilaoInfo(7).Caption = "V: " & Trim$(Leilao(LeilaoSlot).Vendedor) & " Tempo: " & Leilao(LeilaoSlot).Tempo
            Else

                lblLeilaoInfo(1).Caption = Trim$(Item(Leilao(LeilaoSlot).ItemNum).Name)
                lblLeilaoInfo(2).Caption = Trim$(Item(Leilao(LeilaoSlot).ItemNum).Desc)
                If Leilao(LeilaoSlot).Tipo = 1 Then
                    lblLeilaoInfo(4).Caption = "Valor: " & Leilao(LeilaoSlot).Price & " Zeny(s)"
                Else
                    lblLeilaoInfo(4).Caption = "Valor: " & Leilao(LeilaoSlot).Price & " P.C"
                End If
            End If

        End If
    End If

End Sub

Private Sub picLeilao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LeilaoSlot As Long
    LeilaoSlot = IsLeilaoItem(X, Y)

    If Index = 0 Then
        picPokeDesc.Visible = False
        picItemDesc.Visible = False
        LastItemDesc = 0
        LastItemPoke = 0
    End If

    If LeilaoSlot > 0 Then
        If Index = 2 Then
            If Leilao(LeilaoSlot).ItemNum > 0 Then
                If Leilao(LeilaoSlot).Poke.Pokemon > 0 Then
                    X = X + picLeilaoPainel.Left + picLeilao(2).Left - picPokeDesc.Width
                    Y = Y + picLeilaoPainel.top + picLeilao(2).top - picPokeDesc.Height
                    UpdatePokeWindow LeilaoSlot, X, Y, 2    'Inventario
                    LastItemDesc = 0
                    picItemDesc.Visible = False
                Else
                    X = X + picLeilaoPainel.Left + picLeilao(2).Left - picItemDesc.Width
                    Y = Y + picLeilaoPainel.top + picLeilao(2).top - picItemDesc.Height
                    UpdateDescWindow Leilao(LeilaoSlot).ItemNum, X, Y
                    LastItemPoke = 0
                    picPokeDesc.Visible = False
                End If

                LastItemDesc = Leilao(LeilaoSlot).ItemNum
                LastItemPoke = Leilao(LeilaoSlot).Poke.Pokemon
            Else
                picPokeDesc.Visible = False
                picItemDesc.Visible = False
                LastItemDesc = 0
                LastItemPoke = 0
            End If
        End If

        If Index = 3 Then
            If LeilaoItemSelect > 0 Then
                If Leilao(LeilaoItemSelect).ItemNum > 0 Then
                    If Leilao(LeilaoItemSelect).Poke.Pokemon > 0 Then
                        X = X + picLeilaoPainel.Left + picLeilao(3).Left - picPokeDesc.Width
                        Y = Y + picLeilaoPainel.top + picLeilao(3).top - picPokeDesc.Height
                        UpdatePokeWindow LeilaoItemSelect, X, Y, 2    'Inventario
                        LastItemDesc = 0
                        picItemDesc.Visible = False
                    Else
                        X = X + picLeilaoPainel.Left + picLeilao(3).Left - picItemDesc.Width
                        Y = Y + picLeilaoPainel.top + picLeilao(3).top - picItemDesc.Height
                        UpdateDescWindow Leilao(LeilaoItemSelect).ItemNum, X, Y
                        LastItemPoke = 0
                        picPokeDesc.Visible = False
                    End If

                    LastItemDesc = Leilao(LeilaoItemSelect).ItemNum
                    LastItemPoke = Leilao(LeilaoItemSelect).Poke.Pokemon
                Else
                    picPokeDesc.Visible = False
                    picItemDesc.Visible = False
                    LastItemDesc = 0
                    LastItemPoke = 0
                End If
            End If
        End If



    End If

End Sub

Private Sub PicMusic_Click()
    Dim MusicFile As String

    If optMOn.value = True Then
        optMOff.value = True
        Options.Music = 0
        StopMusic
        SaveOptions
        PicMusic.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
    Else
        optMOn.value = True
        Options.Music = 1
        ' start music playing
        MusicFile = Trim$(Map.Music)
        If Not MusicFile = "None." Then
            PlayMusic MusicFile
        Else
            StopMusic
        End If
        ' save to config.ini
        SaveOptions
        PicMusic.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")
    End If
End Sub

Private Sub picPokeDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picPokeDesc.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picPokeDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Janela de botão direito
    If myTargetType = TARGET_TYPE_PLAYER Or TARGET_TYPE_NPC Then
        If Button = vbLeftButton Then
            myTarget = 0
            PicTarget.Visible = False
        End If
    End If

    If InMapEditor Then
        Call MapEditorMouseDown(Button, X, Y, False)
    Else
        ' left click
        If Button = vbLeftButton Then

            ' targetting
            Call PlayerSearch(CurX, CurY)
            ' right click
        ElseIf Button = vbRightButton Then
            SendUnequip 1
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            End If
        End If
    End If

    Call SetFocusOnChat


    ' Atalhos '
    If Button = vbKeyMButton Then
        picAtalho.top = Y
        picAtalho.Left = X
        frmMain.picAtalho.Visible = Not frmMain.picAtalho.Visible


    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'Fecha Janelas Incovenientes
    If picPokeDesc.Visible = True Then picPokeDesc.Visible = False
    If PicVip(2).Visible = True Then PicVip(2).Visible = False
    If picPD.Visible = True Then picPD.Visible = False


    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.top + ((Y + Camera.top) \ PIC_Y)
    MouseX = X
    MouseY = Y

    If InMapEditor Then
        frmEditor_Map.shpLoc.Visible = False

        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, X, Y)
        End If
    End If

    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False

    ' reset all buttons
    resetButtons_Main
    resetButtons_Close

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For i = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            With tempRec
                .top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                .Bottom = .top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.top And Y <= tempRec.Bottom Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picShopItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim shopItem As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    shopItem = IsShopItem(X, Y)

    If shopItem > 0 Then
        Select Case ShopAction
        Case 0    ' no action, give cost
            With Shop(InShop).TradeItem(shopItem)
                frmMain.lblDesc.Caption = "Você pode comprar esse item por:"
                frmMain.lblValor.Caption = .CostValue & " " & Trim$(Item(.CostItem).Name)
            End With
        Case 1    ' buy item
            ' buy item code
            BuyItem shopItem
        End Select
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picShopItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim shopslot As Long
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    shopslot = IsShopItem(X, Y)

    If shopslot <> 0 Then
        X2 = X + picShop.Left + picShopItems.Left + 1
        Y2 = Y + picShop.top + picShopItems.top + 1
        UpdateDescWindow Shop(InShop).TradeItem(shopslot).Item, X2, Y2
        LastItemDesc = Shop(InShop).TradeItem(shopslot).Item
        Exit Sub
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub picWasd_Click()

' Se tiver marcado vai andar com wasd
    If optWOn.value = True Then
        optWOff.value = True
        Options.wasd = 0
        SaveOptions
        picWasd.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
    Else
        optWOn.value = True
        Options.wasd = 1
        SaveOptions
        picWasd.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")
    End If

End Sub

Private Sub PicSound_Click()
    If optSOn.value = True Then
        optSOff.value = True
        Options.sound = 0
        ' save to config.ini
        SaveOptions
        PicSound.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
    Else
        optSOn.value = True
        Options.sound = 1
        ' save to config.ini
        SaveOptions
        PicSound.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")
    End If
End Sub

Private Sub picSpellDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picSpellDesc.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpellDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_DblClick()
    Dim SpellNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellNum = IsPlayerSpell(SpellX, SpellY)

    If SpellNum <> 0 Then
        Call CastSpell(SpellNum)
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SpellNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellNum = IsPlayerSpell(SpellX, SpellY)
    If Button = 1 Then    ' left click
        If SpellNum <> 0 Then
            DragSpell = SpellNum
            Exit Sub
        End If
    ElseIf Button = 2 Then    ' right click
        If SpellNum <> 0 Then
            Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(SpellNum)).Name) & "?", DIALOGUE_TYPE_FORGET, True, SpellNum
            Exit Sub
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim spellslot As Long
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellX = X
    SpellY = Y

    spellslot = IsPlayerSpell(X, Y)

    If DragSpell > 0 Then
        Call BltDraggedSpell(X + picSpells.Left, Y + picSpells.top)
    Else
        If spellslot <> 0 Then
            X2 = X + picSpells.Left - picSpellDesc.Width - 1
            Y2 = Y + picSpells.top - picSpellDesc.Height - 1
            UpdateSpellWindow PlayerSpells(spellslot), X2, Y2
            LastSpellDesc = PlayerSpells(spellslot)
            Exit Sub
        End If
    End If

    picSpellDesc.Visible = False
    LastSpellDesc = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DragSpell > 0 Then
        ' drag + drop
        For i = 1 To MAX_PLAYER_SPELLS
            With rec_pos
                .top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.top And Y <= rec_pos.Bottom Then
                    If DragSpell <> i Then
                        SendChangeSpellSlots DragSpell, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .top = picHotbar.top - picSpells.top
                .Left = picHotbar.Left - picSpells.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.top - picSpells.top + 32
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 2, DragSpell, i
                    DragSpell = 0
                    picTempSpell.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragSpell = 0
    picTempSpell.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picUpDown_Click(Index As Integer)

    Select Case Index
    Case 0    'Down
        If TextStart > 1 Then
            TextStart = TextStart - 1
        End If
    Case 1    'Up
        If TextStart < 14 And Not Chat(TextStart + 6).text = vbNullString Then
            TextStart = TextStart + 1
        End If
    End Select

End Sub


Private Sub picYourTrade_DblClick()
    Dim TradeNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(TradeX, TradeY, True)

    If TradeNum <> 0 Then
        UntradeItem TradeNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picYourTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    Dim InvNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeX = X
    TradeY = Y

    TradeNum = IsTradeItem(X, Y, True)

    If TradeNum <> 0 Then

        InvNum = TradeYourOffer(TradeNum).num

        If GetPlayerInvItemPokeInfoPokemon(MyIndex, InvNum) > 0 Then
            X = X + picYourTrade.Left + 4    '- picPokeDesc.width - 1
            Y = Y + picYourTrade.top + 4    '- picPokeDesc.height - 1
            UpdatePokeWindow InvNum, X, Y, 0    'Inventario
            LastItemPoke = GetPlayerInvItemPokeInfoPokemon(MyIndex, InvNum)
            LastItemDesc = 0
            picItemDesc.Visible = False
            Exit Sub
        Else
            X = X + picYourTrade.Left + 4
            Y = Y + picYourTrade.top + 4
            UpdateDescWindow GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num), X, Y
            LastItemDesc = GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num)    ' set it so you don't re-set values
            picPokeDesc.Visible = False
            LastItemPoke = 0
            Exit Sub
        End If

    End If

    picPokeDesc.Visible = False
    LastItemPoke = 0
    picItemDesc.Visible = False
    LastItemDesc = 0    ' no item was last loaded

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTheirTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long, InvNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(X, Y, False)

    If TradeNum <> 0 Then

        InvNum = TradeNum

        If TradeTheirOffer(TradeNum).PokeInfo.Pokemon > 0 Then
            X = X + picTheirTrade.Left + 4    '- picPokeDesc.width - 1
            Y = Y + picTheirTrade.top + 4    '- picPokeDesc.height - 1
            UpdatePokeWindow TradeNum, X, Y, 3    'Inventario
            LastItemPoke = TradeTheirOffer(InvNum).PokeInfo.Pokemon
            LastItemDesc = 0
            picItemDesc.Visible = False
            Exit Sub
        Else

            X = X + picTheirTrade.Left + 4
            Y = Y + picTheirTrade.top + 4
            UpdateDescWindow TradeTheirOffer(TradeNum).num, X, Y
            LastItemDesc = TradeTheirOffer(TradeNum).num    ' set it so you don't re-set values
            Exit Sub
        End If
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0    ' no item was last loaded

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTheirTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrollBarFake_Click(Index As Integer)
    Select Case Index
    Case 1
        If OrgPage > 1 Then
            OrgPage = OrgPage - 1
            BltOrgShop
        End If
    Case 2
        If OrgPage < 5 Then
            OrgPage = OrgPage + 1
            BltOrgShop
        End If
    Case 4
        If OrgPagMem > 1 Then
            OrgPagMem = OrgPagMem - 1
            BltOrganização
        End If
    Case 3
        If OrgPagMem < 4 Then
            If OrgPagMem < (QntOrgPag - 1) Then
                OrgPagMem = OrgPagMem + 1
                BltOrganização
            End If
        End If
    End Select

    If Index = 1 Or Index = 2 Then
        Select Case OrgPage
        Case 1
            ScrollBarFake(0).top = 37
        Case 2
            ScrollBarFake(0).top = 63
        Case 3
            ScrollBarFake(0).top = 89
        Case 4
            ScrollBarFake(0).top = 115
        Case 5
            ScrollBarFake(0).top = 137
        End Select
    End If

    If Index = 3 Or Index = 4 Then
        If OrgPagMem = 4 Then
            ScrollBarFake(5).top = (148 + ((OrgPagMem - 1) * ScrollBarFake(5).Height) - 1)
        Else
            ScrollBarFake(5).top = 148 + ((OrgPagMem - 1) * ScrollBarFake(5).Height)
        End If
    End If

End Sub

Private Sub ScrollBarFake_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub ScrollBarFake_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then

        If frmMain.ScrollBarFake(0).top + Y - SOffsetY <= 36 Then GoTo Continue:
        If frmMain.ScrollBarFake(0).top + Y - SOffsetY >= 138 Then GoTo Continue:

        If Button = 1 Then
            frmMain.ScrollBarFake(0).top = frmMain.ScrollBarFake(0).top + Y - SOffsetY
        End If

Continue:
        If Button = 1 Then
            Select Case frmMain.ScrollBarFake(0).top
            Case 1 To 37
                If OrgPage <> 1 Then
                    OrgPage = 1
                    BltOrgShop
                End If
            Case 38 To 63
                If OrgPage <> 2 Then
                    OrgPage = 2
                    BltOrgShop
                End If
            Case 64 To 89
                If OrgPage <> 3 Then
                    OrgPage = 3
                    BltOrgShop
                End If
            Case 90 To 115
                If OrgPage <> 4 Then
                    OrgPage = 4
                    BltOrgShop
                End If
            Case 116 To 137
                If OrgPage <> 5 Then
                    OrgPage = 5
                    BltOrgShop
                End If
            End Select
        End If
    End If

End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picFD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
        X = TwipsToPixels(X, 0)
        Y = TwipsToPixels(Y, 1)
        picPD.ZOrder 0
        picPD.Visible = True

        picPD.Left = picFD.Left + 550
        picPD.top = picFD.top + (picFD.Height - 2)
'
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Shift = vbAltMask Then
        Select Case KeyCode
        Case vbKeyReturn
            If FullScreen = False Then
                ChangeToFullScreen
                FullScreen = True
            Else
                ChangeToWindowed
                FullScreen = False
            End If
            Exit Sub
        End Select
    End If

    If txtMyChat.Visible = False Then

        Select Case KeyCode
        Case vbKeyT
            If txtMyChat.Visible = False Then
                If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                    SendTradeRequest
                    ' play sound
                    PlaySound Sound_ButtonClick, -1, -1
                Else
                    AddText "Invalid trade target.", BrightRed
                End If
            End If

        Case vbKeyInsert
            If Player(MyIndex).Access > 0 Then
                frmPanel.Visible = Not frmPanel.Visible
            End If

        Case vbKeyE
            frmMain.picAtalho.Visible = Not frmMain.picAtalho.Visible
            frmMain.picAtalho.top = (frmMain.ScaleHeight / 2) - (frmMain.picAtalho.Height / 2)
            frmMain.picAtalho.Left = (frmMain.ScaleWidth / 2) - (frmMain.picAtalho.Width / 2)

        End Select

    End If

    ' hotbar
    If txtMyChat.Visible = False Then

        Select Case KeyCode
        Case vbKeyQ
            SendUnequip 1
        Case vbKeyZ
            If GetPlayerEquipmentPokeInfoSpell(MyIndex, weapon, 1) > 0 Then
                CastSpell 1
            End If
        Case vbKeyX
            If GetPlayerEquipmentPokeInfoSpell(MyIndex, weapon, 2) > 0 Then
                CastSpell 2
            End If
        Case vbKeyC
            If GetPlayerEquipmentPokeInfoSpell(MyIndex, weapon, 3) > 0 Then
                CastSpell 3
            End If
        Case vbKeyV
            If GetPlayerEquipmentPokeInfoSpell(MyIndex, weapon, 4) > 0 Then
                CastSpell 4
            End If

        End Select

        For i = 1 To 6
            If KeyCode = 48 + i Then
                If SpellBufferTimer > 0 Then Exit Sub
                SendHotbarUse i
            End If
        Next

    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub TimerDestroyGame_Timer()
    If InGame Then

        SegToQuit = SegToQuit - 1
        lblDesco.Caption = "Desconectando em " & SegToQuit & " segundos"
        PlaySound Sound_Desconect, -1, -1

        If SegToQuit = 0 Then
            DestroyGame
            TimerDestroyGame.Enabled = False
            PicDesconect.Visible = False
        End If

    End If
End Sub

Private Sub TimerLogout_Timer()
    If InGame Then

        SegToQuit = SegToQuit - 1
        lblDesco.Caption = "Desconectando em " & SegToQuit & " segundos"
        PlaySound Sound_Desconect, -1, -1

        If SegToQuit = 0 Then
            logoutGame
            TimerLogout.Enabled = False
            PicDesconect.Visible = False
        End If

    End If
End Sub


Private Sub tmrNoticia_Timer()
    Dim i As Long, Ordem As Byte

    ' Ações...
    If Len(NoticiaServ(1)) > 0 Then
        NotX = NotX + 5
        If Camera.Left + ((MAX_MAPX) * 32) - NotX <= 0 Then
            NoticiaServ(1) = Mid$(NoticiaServ(1), 4, Len(NoticiaServ(1)))
        End If
    Else

        'Organizar Mensagens
        For i = 1 To MAX_NOTICIAS
            If NoticiaServ(i) = vbNullString Then
                Ordem = Ordem + 1
            Else
                NoticiaServ(i - Ordem) = NoticiaServ(i)
                NoticiaServ(i) = vbNullString
            End If
        Next

        'Resetar Posições
        NotX = 0

        'Acabar as Msg .-.
        If NoticiaServ(1) = vbNullString Then
            tmrNoticia.Enabled = False
        End If

    End If

End Sub

Private Sub txtMyChat_Change()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    MyText = txtMyChat

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMyChat_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChat_GotFocus()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SetFocusOnChat

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChat_GotFocus", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsEqItem = 0

    For i = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, i) > 0 And GetPlayerEquipment(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .top = EqTop
                .Bottom = .top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.top And Y <= tempRec.Bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsEqItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsInvItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV

        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.top And Y <= tempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If PlayerSpells(i) > 0 And PlayerSpells(i) <= MAX_SPELLS Then

            With tempRec
                .top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.top And Y <= tempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean) As Long
    Dim tempRec As RECT
    Dim i As Long
    Dim ItemNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV

        If Yours Then
            ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        Else
            ItemNum = TradeTheirOffer(i).num
        End If

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then

            With tempRec
                .top = TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                .Bottom = .top + PIC_Y
                .Left = TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.top And Y <= tempRec.Bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTradeItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function


Private Sub picItemDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picItemDesc.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picItemDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************
' ** Char window **
' *****************

Private Sub picCharacter_Click()
    Dim EqNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqNum = IsEqItem(EqX, EqY)

    If EqNum <> 0 Then
        SendUnequip EqNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim EqNum As Long
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqX = X
    EqY = Y
    EqNum = IsEqItem(X, Y)

    If EqNum <> 0 Then
        Y2 = Y + picCharacter.top - frmMain.picItemDesc.Height - 1
        X2 = X + picCharacter.Left - frmMain.picItemDesc.Width - 1
        UpdateDescWindow GetPlayerEquipment(MyIndex, EqNum), X2, Y2
        LastItemDesc = GetPlayerEquipment(MyIndex, EqNum)    ' set it so you don't re-set values
        LastItemPoke = GetPlayerEquipmentPokeInfoPokemon(MyIndex, EqNum)
        Exit Sub
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0    ' no item was last loaded
    LastItemPoke = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' bank
Private Sub picBank_DblClick()
    Dim bankNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragBankSlotNum = 0

    bankNum = IsBankItem(BankX, BankY)
    If bankNum <> 0 Then
        If GetBankItemNum(bankNum) = ITEM_TYPE_NONE Then Exit Sub

        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_CURRENCY Then
            CurrencyMenu = 3    ' withdraw
            lblCurrency.Caption = "Qual a quantia que deseja retirar?"
            tmpCurrencyItem = bankNum
            txtCurrency.text = vbNullString
            picCurrency.Visible = True
            picCurrency.ZOrder 0
            frmMain.picCurrency.top = (frmMain.ScaleHeight / 2) - (frmMain.picCurrency.Height / 2)
            frmMain.picCurrency.Left = (frmMain.ScaleWidth / 2) - (frmMain.picCurrency.Width / 2)
            txtCurrency.SetFocus
            Exit Sub
        End If

        WithdrawItem bankNum, 0
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bankNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    bankNum = IsBankItem(X, Y)

    If bankNum <> 0 Then

        If Button = 1 Then
            DragBankSlotNum = bankNum
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' TODO : Add sub to change bankslots client side first so there's no delay in switching
    If DragBankSlotNum > 0 Then
        For i = 1 To MAX_BANK
            With rec_pos
                .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.top And Y <= rec_pos.Bottom Then
                    If DragBankSlotNum <> i Then
                        ChangeBankSlots DragBankSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    DragBankSlotNum = 0
    picTempBank.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bankNum As Long, ItemNum As Long, ItemType As Long
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    BankX = X
    BankY = Y

    If DragBankSlotNum > 0 Then
        Call BltBankItem(X + picBank.Left, Y + picBank.top)
    Else
        bankNum = IsBankItem(X, Y)

        If bankNum <> 0 Then

            If Bank.Item(bankNum).PokeInfo.Pokeball > 0 Then
                X = X + picBank.Left + 1
                Y = Y + picBank.top + 1
                UpdatePokeWindow bankNum, X, Y, 1    'Inventario
                LastItemPoke = GetPlayerBankItemPokemon(bankNum)
                LastItemDesc = 0
                picItemDesc.Visible = False
            End If

            If Bank.Item(bankNum).PokeInfo.Pokeball = 0 Then
                X2 = X + picBank.Left + 1
                Y2 = Y + picBank.top + 1
                UpdateDescWindow Bank.Item(bankNum).num, X2, Y2
                LastItemDesc = Bank.Item(bankNum).num
                LastItemPoke = 0
                picPokeDesc.Visible = False
            End If

            Exit Sub
        End If
    End If

    frmMain.picItemDesc.Visible = False
    picPokeDesc.Visible = False
    LastBankDesc = 0
    LastItemPoke = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsBankItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsBankItem = 0

    For i = 1 To MAX_BANK
        If GetBankItemNum(i) > 0 And GetBankItemNum(i) <= MAX_ITEMS Then

            With tempRec
                .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.top And Y <= tempRec.Bottom Then

                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub imgQuest_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset other buttons
    resetButtons_Quest Index

    ' change the button we're hovering on
    changeButtonState_Quest Index, 2    ' clicked

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgQuest_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset other buttons
    resetButtons_Quest Index

    ' change the button we're hovering on
    If Not QuestButton(Index).state = 2 Then    ' make sure we're not clicking
        changeButtonState_Quest Index, 1    ' hover
    End If

    ' play sound
    If Not LastButtonSound_Quest = Index Then
        PlaySound Sound_ButtonHover, -1, -1
        LastButtonSound_Quest = Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgQuest_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset all buttons
    resetButtons_Quest -1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsLeilaoItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    Dim AB As Long, BA As Long

    Select Case PageLeilao
    Case 1
        AB = 1
        BA = 20
    Case 2
        AB = 21
        BA = 40
    Case 3
        AB = 41
        BA = 60
    Case 4
        AB = 61
        BA = 80
    Case 5
        AB = 81
        BA = 100
    End Select

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsLeilaoItem = 0

    For i = AB To BA

        If AB >= 21 Then
            With tempRec
                .top = LeilaoTop + ((LeilaoOffsetY + 32) * ((i - AB) \ LeilaoColumns))
                .Bottom = .top + PIC_Y
                .Left = LeilaoLeft + ((LeilaoOffsetX + 32) * (((i - 1) Mod LeilaoColumns)))
                .Right = .Left + PIC_X
            End With
        Else
            With tempRec
                .top = LeilaoTop + ((LeilaoOffsetY + 32) * ((i - 1) \ LeilaoColumns))
                .Bottom = .top + PIC_Y
                .Left = LeilaoLeft + ((LeilaoOffsetX + 32) * (((i - 1) Mod LeilaoColumns)))
                .Right = .Left + PIC_X
            End With
        End If

        If X >= tempRec.Left And X <= tempRec.Right Then
            If Y >= tempRec.top And Y <= tempRec.Bottom Then
                IsLeilaoItem = i
                Exit Function
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsLeilaoItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub CarregarTrainerCard()
    Dim i As Long
    Dim QntiaPokeDex As Long

    lblTrainerCard(0).Caption = GetPlayerName(MyIndex)
    BltFace

    For i = 1 To MAX_POKEMONS
        If Player(MyIndex).Pokedex(i) = 1 Then
            QntiaPokeDex = QntiaPokeDex + 1
        End If
    Next

    lblTrainerCard(1).Caption = "Pokédex:" & QntiaPokeDex
End Sub

Private Function IsOrgShopItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long, a As Byte, B As Byte
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsOrgShopItem = 0

    Select Case OrgPage
    Case 1
        a = 1
        B = 4
    Case 2
        a = 5
        B = 8
    Case 3
        a = 9
        B = 12
    Case 4
        a = 13
        B = 16
    Case 5
        a = 17
        B = 20
    End Select

    For i = a To B
        If GetOrgItemNum(i) > 0 And GetOrgItemNum(i) <= MAX_ITEMS Then

            If a = 1 Then
                With tempRec
                    .top = OrgTop + ((OrgOffsetY + 32) * ((i - 1) \ OrgColumns))
                    .Bottom = .top + PIC_Y
                    .Left = OrgLeft + ((OrgOffsetX + 32) * (((i - 1) Mod OrgColumns)))
                    .Right = .Left + PIC_X
                End With
            Else
                With tempRec
                    .top = OrgTop + ((OrgOffsetY + 32) * ((i - a) \ OrgColumns))
                    .Bottom = .top + PIC_Y
                    .Left = OrgLeft + ((OrgOffsetX + 32) * (((i - a) Mod OrgColumns)))
                    .Right = .Left + PIC_X
                End With
            End If

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.top And Y <= tempRec.Bottom Then

                    IsOrgShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsQuestItemSlot(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long, a As Long, B As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsQuestItemSlot = 0

    Select Case RewardsPage
    Case 1
        a = 6
        B = 10
    Case Else
        a = 1
        B = 5
    End Select

    For i = a To B

        If i <= 5 Then
            With tempRec
                .top = 230
                .Bottom = .top + PIC_Y
                .Left = 380 + (i - 1 Mod 5) * 41
                .Right = .Left + PIC_X
            End With
        Else
            With tempRec
                .top = 230
                .Bottom = .top + PIC_Y
                .Left = 380 + (i - a Mod 10) * 41
                .Right = .Left + PIC_X
            End With
        End If

        If X >= tempRec.Left And X <= tempRec.Right Then
            If Y >= tempRec.top And Y <= tempRec.Bottom Then
                IsQuestItemSlot = i
                Exit Function
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub VipConfCancel_Click(Index As Integer)
    Select Case Index
    Case 0
        SendObterPacVip ObterVipNumber, 0
        ObterVipNumber = 0
        PicVip(1).Visible = False
    Case 1
        ObterVipNumber = 0
        PicVip(1).Visible = False
    End Select
End Sub

' BOTÕES COM ANIMAÇÃO

Private Sub imgClose_Click(Index As Integer)

' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case Index

        ' Lista de Quest
    Case 1
        frmMain.picSelectQuest.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Menu de Atalhos
    Case 2
        picAtalho.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Personagem
    Case 3
        PicTreinador.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Seleção de pokémon
    Case 4
        PicPokeInicial.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Inventário
    Case 5
        picInventory.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Líder
    Case 6
        ChatGym = 0
        ChatGymStep = 0
        frmMain.PicBlank(0).Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Quest
    Case 7
        picQuest.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Surf
    Case 8
        SendSurfInit 0
        frmMain.PicSurf.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Evoluir
    Case 9
        Call SendEvolCommand(1)
        EvolTick = 0
        EvolutionTimer.Enabled = False
        PicEvolution.Visible = False
        imgClose(9).Visible = False
        imgButton(14).Visible = False
        Player(MyIndex).EvolPermition = 0
        PlaySound Sound_ButtonClick, -1, -1

        ' Habilidades
    Case 10
        SendAprenderHab Index
        frmMain.PicHabilidade.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Premium
    Case 11
        frmMain.PicVipPanel.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Pokedex
    Case 12
        frmMain.picPokedex.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Opções
    Case 13
        frmMain.picOptions.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Batalha
    Case 14
        frmMain.PicBatalha.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Organização
    Case 15
        frmMain.PicOrgs.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Quantidade
    Case 16
        picCurrency.Visible = False
        txtCurrency.text = vbNullString
        tmpCurrencyItem = 0
        CurrencyMenu = 0    ' clear
        PlaySound Sound_ButtonClick, -1, -1

        ' Leilão
    Case 17
        picLeilaoPainel.Visible = False
        PlaySound Sound_ButtonClick, -1, -1

        ' Loja
    Case 18
        Dim Buffer As clsBuffer
        Set Buffer = New clsBuffer

        Buffer.WriteLong CCloseShop

        SendData Buffer.ToArray()

        Set Buffer = Nothing

        picCover.Visible = False
        picShop.Visible = False
        InShop = 0
        ShopAction = 0
        
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgClose_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgClose_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset other buttons
    resetButtons_Close Index

    ' change the button we're hovering on
    changeButtonState_Close Index, 2    ' clicked

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgClose_MouseDown", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset other buttons
    resetButtons_Close Index

    ' change the button we're hovering on
    If Not CloseButton(Index).state = 2 Then    ' make sure we're not clicking
        changeButtonState_Close Index, 1    ' hover
    End If

    ' play sound
    If Not LastButtonSound_Close = Index Then
        PlaySound Sound_ButtonHover, -1, -1
        LastButtonSound_Close = Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgClose_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgClose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset all buttons
    resetButtons_Close -1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgClose_MouseUp", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


'''''''''''''''
' CÓDIGO ORGS '
'''''''''''''''
Private Sub PicOrg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim OrgShopNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    OrgShopNum = IsOrgShopItem(X, Y)

    If OrgShopNum <> 0 Then
        DragOrgShopNum = OrgShopNum
    End If

    BltItemSelectOrgShop

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicOrg_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub PicOrgs_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicOrgs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.PicOrgs, Button, Shift, X, Y)
    Call MovePicture(frmMain.PicOrg(1), Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicOrg_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''
' CÓDIGO LOJA '
'''''''''''''''

Private Sub picShop_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picShop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.picShop, Button, Shift, X, Y)
    
    picItemDesc.Visible = False
    picSpellDesc.Visible = False

    ' reset all buttons
    resetButtons_Main
    resetButtons_Close
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picLeilaoPainel_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''''
' CÓDIGO LEILÃO '
'''''''''''''''''

Private Sub picLeilaoPainel_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picLeilaoPainel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.picLeilaoPainel, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picLeilaoPainel_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

''''''''''''''''''
' CÓDIGO BATALHA '
'''''''''''''''' '

Private Sub PicBatalha_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicBatalha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.PicBatalha, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picOptions_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''''
' CÓDIGO OPÇÕES '
'''''''''''''''''

Private Sub picOptions_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.picOptions, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picOptions_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

''''''''''''''''''
' CÓDIGO POKEDEX '
''''''''''''''''''

Private Sub picPokedex_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
    
End Sub

Private Sub picPokedex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.picPokedex, Button, Shift, X, Y)

    If PicVip(2).Visible = True Then PicVip(2).Visible = False
    If picPD.Visible = True Then picPD.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picPokedex_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''''''
' CÓDIGO CURRENCY '
'''''''''''''''''''

Private Sub picCurrency_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picCurrency_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.picCurrency, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCurrency_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

''''''''''''''
' CÓDIGO VIP '
''''''''''''''

Private Sub PicVipPanel_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicVipPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.PicVipPanel, Button, Shift, X, Y)

    If PicVip(2).Visible = True Then PicVip(2).Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicVipPanel_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

''''''''''''''''''''''''''
' CÓDIGO SELECIONAR QUEST'
''''''''''''''''''''''''''

Private Sub picSelectQuest_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picSelectQuest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.picSelectQuest, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSelectQuest_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''
' CÓDIGO SURF '
'''''''''''''''

Private Sub PicSurf_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicSurf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.PicSurf, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicSurf_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''''''
' CÓDIGO EVOLUCÃO '
'''''''''''''''''''

Private Sub PicEvolution_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicEvolution_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.PicEvolution, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicEvolution_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''''''''
' CÓDIGO PERSONAGEM '
'''''''''''''''''''''

Private Sub PicTreinador_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicTreinador_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.PicTreinador, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicTreinador_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

''''''''''''''''''
' CÓDIGO ATALHOS '
''''''''''''''''''

Private Sub picAtalho_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picAtalho_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMain.picAtalho, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picAtalho_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''''''
' ESCOLHA POKEMON '
'''''''''''''''''''

Private Sub PicPokeInicial_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicPokeInicial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picPokeDesc.Visible = False
    LastItemDesc = 0
    LastItemPoke = 0

    Call MovePicture(frmMain.PicPokeInicial, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicPokeInicial_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

''''''''''''''''''''''
' ESCOLHA HABILIDADE '
''''''''''''''''''''''

Private Sub PicHabilidade_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicHabilidade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picPokeDesc.Visible = False
    LastItemDesc = 0
    LastItemPoke = 0

    Call MovePicture(frmMain.PicHabilidade, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicHabilidade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***************
' ** Inventory **
' ***************

Private Sub picInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InvX = X
    InvY = Y

    If DragInvSlotNum > 0 Then
        If InTrade > 0 Then Exit Sub
        If InBank Or InShop Then Exit Sub
        Call BltInventoryItem(X + picInventory.Left, Y + picInventory.top)
    Else

        Call MovePicture(frmMain.picInventory, Button, Shift, X, Y)

        InvNum = IsInvItem(X, Y)

        If InvNum <> 0 Then
            ' exit out if we're offering that item
            If InTrade Then
                For i = 1 To MAX_INV
                    If TradeYourOffer(i).num = InvNum Then
                        ' is currency?
                        If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Type = ITEM_TYPE_CURRENCY Then
                            ' only exit out if we're offering all of it
                            If TradeYourOffer(i).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                                Exit Sub
                            End If
                        Else
                            Exit Sub
                        End If
                    End If
                Next
            End If

            If GetPlayerInvItemPokeInfoPokemon(MyIndex, InvNum) > 0 Then
                X = X + picInventory.Left - picPokeDesc.Width - 1
                Y = Y + picInventory.top - picPokeDesc.Height - 1
                UpdatePokeWindow InvNum, X, Y, 0    'Inventario
                LastItemDesc = 0
                picItemDesc.Visible = False
            End If

            If GetPlayerInvItemPokeInfoPokemon(MyIndex, InvNum) = 0 Then
                X = X + picInventory.Left - picItemDesc.Width - 1
                Y = Y + picInventory.top - picItemDesc.Height - 1
                UpdateDescWindow GetPlayerInvItemNum(MyIndex, InvNum), X, Y
                LastItemPoke = 0
                picPokeDesc.Visible = False
                If GetPlayerInvItemNum(MyIndex, InvNum) = 50 Then
                    lblItemDesc.Caption = "Entregue está ExpBall para um de Seus pokémons! Há " & PlayerInv(InvNum).PokeInfo.Exp & " Exp"
                End If
            End If

            LastItemDesc = GetPlayerInvItemNum(MyIndex, InvNum)    ' set it so you don't re-set values
            LastItemPoke = GetPlayerInvItemPokeInfoPokemon(MyIndex, InvNum)
            Exit Sub
        End If
    End If

    picItemDesc.Visible = False
    picPokeDesc.Visible = False
    LastItemDesc = 0    ' no item was last loaded
    LastItemPoke = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        ' drag + drop
        For i = 1 To MAX_INV
            With rec_pos
                .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.top And Y <= rec_pos.Bottom Then    '
                    If DragInvSlotNum <> i Then
                        SendChangeInvSlots DragInvSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .top = picHotbar.top - picInventory.top
                .Left = picHotbar.Left - picInventory.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.top - picInventory.top + 32
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 1, DragInvSlotNum, i
                    DragInvSlotNum = 0
                    picTempInv.Visible = False
                    blthotbar
                    Exit Sub
                End If
            End If
        Next
    End If

    DragInvSlotNum = 0
    picTempInv.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SOffsetX = X
    SOffsetY = Y

    InvNum = IsInvItem(X, Y)

    If Button = 1 Then
        If InvNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = InvNum
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If InvNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(MyIndex, InvNum) > 0 Then
                        CurrencyMenu = 1    ' drop
                        lblCurrency.Caption = "Quantos você quer soltar?"
                        tmpCurrencyItem = InvNum
                        txtCurrency.text = vbNullString
                        picCurrency.Visible = True
                        picCurrency.ZOrder 0
                        frmMain.picCurrency.top = (frmMain.ScaleHeight / 2) - (frmMain.picCurrency.Height / 2)
                        frmMain.picCurrency.Left = (frmMain.ScaleWidth / 2) - (frmMain.picCurrency.Width / 2)
                        txtCurrency.SetFocus
                    End If
                Else
                    Call SendDropItem(InvNum, 0)
                End If
            End If
        End If
    End If

    SetFocusOnChat

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_DblClick()
    Dim InvNum As Long
    Dim value As Long
    Dim multiplier As Double
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragInvSlotNum = 0
    InvNum = IsInvItem(InvX, InvY)

    picCurrency.ZOrder (0)

    If InvNum <> 0 Then

        ' are we in a shop?
        If InShop > 0 Then
            Select Case ShopAction
            Case 0    ' nothing, give value
                multiplier = Shop(InShop).BuyRate / 100
                value = Item(GetPlayerInvItemNum(MyIndex, InvNum)).Price * multiplier
                If value > 0 Then
                    AddText "You can sell this item for " & value & " gold.", White
                Else
                    AddText "The shop does not want this item.", BrightRed
                End If
            Case 2    ' 2 = sell
                SellItem InvNum
            End Select

            Exit Sub
        End If

        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 2    ' deposit
                lblCurrency.Caption = "Quantos você quer depositar?"
                tmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                picCurrency.ZOrder 0
                frmMain.picCurrency.top = (frmMain.ScaleHeight / 2) - (frmMain.picCurrency.Height / 2)
                frmMain.picCurrency.Left = (frmMain.ScaleWidth / 2) - (frmMain.picCurrency.Width / 2)
                txtCurrency.SetFocus
                Exit Sub
            End If

            Call DepositItem(InvNum, 0)
            Exit Sub
        End If

        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).num = InvNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Type = ITEM_TYPE_CURRENCY Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next

            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 4    ' offer in trade
                lblCurrency.Caption = "Quantos você quer trocar?"
                tmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                picCurrency.ZOrder 0
                frmMain.picCurrency.top = (frmMain.ScaleHeight / 2) - (frmMain.picCurrency.Height / 2)
                frmMain.picCurrency.Left = (frmMain.ScaleWidth / 2) - (frmMain.picCurrency.Width / 2)
                txtCurrency.SetFocus
                Exit Sub
            End If

            Call TradeItem(InvNum, 0)
            Exit Sub
        End If

        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(InvNum)
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

''''''''''''''''
' JANELA QUEST '
''''''''''''''''
Private Sub picQuest_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picQuest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ItemNum As Long, QuestSlot As Byte, X2 As Long, Y2 As Long
    Dim QuestNum As Long

    QuestSlot = IsQuestItemSlot(X, Y)
    QuestNum = GetQuestNum(Trim$(frmMain.lstQuests.text))

    Call MovePicture(frmMain.picQuest, Button, Shift, X, Y)

    If QuestNum = 0 Then GoTo Continue
    If QuestSlot = 0 Then GoTo Continue
    ItemNum = Quest(QuestNum).ItemRew(QuestSlot)

    If ItemNum = 0 Then GoTo Continue

    If Quest(QuestNum).PokeRew(QuestSlot) = 0 Then
        X2 = X + picQuest.Left - picItemDesc.Width - 1
        Y2 = Y + picQuest.top - picItemDesc.Height - 1
        UpdateDescWindow ItemNum, X2, Y2
        LastItemDesc = ItemNum
        LastItemPoke = 0
        picPokeDesc.Visible = False
        Exit Sub
    Else
        X2 = X + picQuest.Left - picPokeDesc.Width - 1
        Y2 = Y + picQuest.top - picPokeDesc.Height - 1
        UpdatePokeWindow QuestSlot, X2, Y2, 4, QuestNum
        LastItemDesc = 0
        LastItemPoke = Quest(QuestNum).PokeRew(QuestSlot)
        picItemDesc.Visible = False
        Exit Sub
    End If

Continue:
    frmMain.picItemDesc.Visible = False
    picPokeDesc.Visible = False
    LastBankDesc = 0
    LastItemPoke = 0
End Sub
