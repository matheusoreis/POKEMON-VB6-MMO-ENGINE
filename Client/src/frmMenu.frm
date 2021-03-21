VERSION 5.00
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17280
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMenu.frx":0CCA
   ScaleHeight     =   640
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1152
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   13920
      Top             =   7200
   End
   Begin VB.PictureBox picTermos 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7320
      Left            =   4080
      Picture         =   "frmMenu.frx":D4414
      ScaleHeight     =   488
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   554
      TabIndex        =   45
      Top             =   840
      Width           =   8310
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   11
         Left            =   7965
         Picture         =   "frmMenu.frx":19A858
         ScaleHeight     =   390
         ScaleWidth      =   330
         TabIndex        =   47
         Top             =   15
         Width           =   330
      End
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   12
         Left            =   3330
         Picture         =   "frmMenu.frx":19AF84
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   46
         Top             =   6675
         Width           =   1695
      End
   End
   Begin VB.PictureBox picCharacter 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4545
      Left            =   6525
      Picture         =   "frmMenu.frx":19CE54
      ScaleHeight     =   303
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   36
      Top             =   3570
      Visible         =   0   'False
      Width           =   4245
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   10
         Left            =   3900
         Picture         =   "frmMenu.frx":1DBF04
         ScaleHeight     =   390
         ScaleWidth      =   330
         TabIndex        =   0
         Top             =   15
         Width           =   330
      End
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   9
         Left            =   1275
         Picture         =   "frmMenu.frx":1DC630
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   1
         Top             =   3900
         Width           =   1695
      End
      Begin VB.OptionButton optFemale 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Female"
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
         Height          =   255
         Left            =   4440
         TabIndex        =   41
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton optMale 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Male"
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
         Height          =   255
         Left            =   4440
         TabIndex        =   40
         Top             =   1680
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.ComboBox cmbClass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1530
         Left            =   1365
         ScaleHeight     =   102
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   102
         TabIndex        =   38
         Top             =   1425
         Width           =   1530
         Begin VB.Label lblSprite 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trocar Sprite"
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
            Left            =   -15
            TabIndex        =   44
            Top             =   1200
            Visible         =   0   'False
            Width           =   1545
         End
      End
      Begin VB.TextBox txtCName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   360
         MaxLength       =   15
         TabIndex        =   37
         Top             =   870
         Width           =   3495
      End
      Begin VB.Label lblCAccept 
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
         Height          =   330
         Left            =   1560
         TabIndex        =   43
         Top             =   3360
         Width           =   1170
      End
      Begin VB.Image ImgCorCabelo 
         Height          =   255
         Index           =   1
         Left            =   2055
         Top             =   1515
         Width           =   255
      End
      Begin VB.Image ImgCorCabelo 
         Height          =   255
         Index           =   2
         Left            =   2325
         Top             =   1515
         Width           =   255
      End
      Begin VB.Image ImgCorCabelo 
         Height          =   255
         Index           =   3
         Left            =   2580
         Top             =   1515
         Width           =   255
      End
      Begin VB.Image imgSex 
         Height          =   375
         Index           =   1
         Left            =   2640
         Top             =   3000
         Width           =   255
      End
      Begin VB.Image imgSex 
         Height          =   375
         Index           =   0
         Left            =   1350
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label LblSex 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Masculino"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   42
         Top             =   3105
         Width           =   855
      End
   End
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      Picture         =   "frmMenu.frx":1DE500
      ScaleHeight     =   405
      ScaleWidth      =   17280
      TabIndex        =   32
      Top             =   9195
      Width           =   17280
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Offline"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   48
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblExtra 
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
         Height          =   405
         Index           =   0
         Left            =   12360
         TabIndex        =   35
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblExtra 
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
         Height          =   405
         Index           =   1
         Left            =   14280
         TabIndex        =   34
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblExtra 
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
         Height          =   405
         Index           =   2
         Left            =   1080
         TabIndex        =   33
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox picLogin 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   6525
      Picture         =   "frmMenu.frx":1F51C4
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   13
      Top             =   5760
      Width           =   4245
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   1
         Left            =   3900
         Picture         =   "frmMenu.frx":215C8C
         ScaleHeight     =   390
         ScaleWidth      =   330
         TabIndex        =   9
         Top             =   15
         Width           =   330
      End
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   1275
         Picture         =   "frmMenu.frx":2163B8
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   8
         Top             =   1785
         Width           =   1695
      End
      Begin VB.PictureBox PicPass 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   480
         Picture         =   "frmMenu.frx":218288
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   31
         Top             =   2520
         Width           =   240
      End
      Begin VB.CheckBox chkPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Save Password?"
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
         Left            =   3960
         TabIndex        =   26
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtLPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   375
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   25
         Top             =   1395
         Width           =   3495
      End
      Begin VB.TextBox txtLUser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   375
         MaxLength       =   12
         TabIndex        =   24
         Top             =   795
         Width           =   3495
      End
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3765
      Left            =   16680
      ScaleHeight     =   3765
      ScaleWidth      =   6630
      TabIndex        =   28
      Top             =   4080
      Visible         =   0   'False
      Width           =   6630
   End
   Begin VB.PictureBox picRegister 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   6525
      Picture         =   "frmMenu.frx":21A170
      ScaleHeight     =   284
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   27
      Top             =   3855
      Visible         =   0   'False
      Width           =   4245
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   4
         Left            =   3900
         Picture         =   "frmMenu.frx":268984
         ScaleHeight     =   390
         ScaleWidth      =   330
         TabIndex        =   11
         Top             =   15
         Width           =   330
      End
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   1275
         Picture         =   "frmMenu.frx":2690B0
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   7
         Top             =   3615
         Width           =   1695
      End
      Begin VB.TextBox txtEmail 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   375
         MaxLength       =   255
         TabIndex        =   20
         Top             =   1395
         Width           =   3495
      End
      Begin VB.TextBox txtRecovery 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   375
         MaxLength       =   20
         TabIndex        =   23
         Top             =   3195
         Width           =   3495
      End
      Begin VB.TextBox txtRPass2 
         Alignment       =   2  'Center
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
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   375
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   22
         Top             =   2595
         Width           =   3495
      End
      Begin VB.TextBox txtRPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   375
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   21
         Top             =   1995
         Width           =   3495
      End
      Begin VB.TextBox txtRUser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   225
         Left            =   375
         MaxLength       =   12
         TabIndex        =   19
         Top             =   795
         Width           =   3495
      End
   End
   Begin VB.PictureBox PicNewPass 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   4260
      Left            =   6525
      Picture         =   "frmMenu.frx":26AF80
      ScaleHeight     =   284
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   30
      Top             =   3855
      Visible         =   0   'False
      Width           =   4245
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   6
         Left            =   3900
         Picture         =   "frmMenu.frx":2A60F4
         ScaleHeight     =   390
         ScaleWidth      =   330
         TabIndex        =   10
         Top             =   15
         Width           =   330
      End
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   1275
         Picture         =   "frmMenu.frx":2A6820
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   6
         Top             =   3615
         Width           =   1695
      End
      Begin VB.TextBox txtNewPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   375
         MaxLength       =   255
         TabIndex        =   18
         Top             =   3195
         Width           =   3495
      End
      Begin VB.TextBox txtNewPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   375
         MaxLength       =   255
         TabIndex        =   17
         Top             =   2595
         Width           =   3495
      End
      Begin VB.TextBox txtNewPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   375
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1395
         Width           =   3495
      End
      Begin VB.TextBox txtNewPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   375
         MaxLength       =   12
         TabIndex        =   14
         Top             =   795
         Width           =   3495
      End
      Begin VB.TextBox txtNewPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   375
         MaxLength       =   255
         TabIndex        =   16
         Top             =   1995
         Width           =   3495
      End
   End
   Begin VB.PictureBox PicRecover 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   3075
      Left            =   6525
      Picture         =   "frmMenu.frx":2A86F0
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   4245
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   1275
         Picture         =   "frmMenu.frx":2D3178
         ScaleHeight     =   345
         ScaleWidth      =   1695
         TabIndex        =   2
         Top             =   2430
         Width           =   1695
      End
      Begin VB.PictureBox imgButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   8
         Left            =   3900
         Picture         =   "frmMenu.frx":2D5048
         ScaleHeight     =   390
         ScaleWidth      =   330
         TabIndex        =   12
         Top             =   15
         Width           =   330
      End
      Begin VB.TextBox txtRecover 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   375
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1995
         Width           =   3495
      End
      Begin VB.TextBox txtRecover 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   375
         MaxLength       =   12
         TabIndex        =   3
         Top             =   795
         Width           =   3495
      End
      Begin VB.TextBox txtRecover 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   375
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1395
         Width           =   3495
      End
   End
   Begin Project1.PictureG PictureG1 
      Height          =   10500
      Left            =   0
      Top             =   0
      Width           =   18000
      _ExtentX        =   31750
      _ExtentY        =   18521
      GIF             =   "frmMenu.frx":2D5774
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private VoltarBack As Boolean

Private Sub cmbClass_Click()
    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    NewCharacterBltSprite
End Sub

Private Sub Form_Activate()
tmrStatus.Enabled = True
If ConnectToServer(1) Then
SendRequestStatus
Else
lblStatus.Caption = "Offline"
lblStatus.ForeColor = QBColor(BrightRed)

End If
End Sub

Private Sub Form_Load()
    Dim tmpTxt As String, tmpArray() As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' general menu stuff
    Me.Caption = Options.Game_Name

    ' Load the username + pass
    txtLUser.text = Trim$(Options.Username)
    If Options.SavePass = 1 Then
        txtLPass.text = Trim$(Options.Password)
        chkPass.value = Options.SavePass
    End If

    If chkPass.value = False Then
        PicPass.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
    Else
        PicPass.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")
    End If
    
    If Options.Termos = 0 Then
        PicTermos.Visible = True
    Else
        PicTermos.Visible = False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    resetButtons_Menu

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not EnteringGame Then DestroyGame

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgSex_Click(Index As Integer)

    Select Case Index

    Case 0
        optMale.value = True
        LblSex.Caption = "Masculino"
    Case 1
        optFemale.value = True
        LblSex.Caption = "Feminino"
    End Select

End Sub

Private Sub lblExtra_Click(Index As Integer)
    Dim Name As String
    Dim RecoveryKey As String
    Dim Email As String
    Dim OldPassword As String, NewPassword As String, ReNewPassword As String

    Select Case Index

        ' Recuperar a senha
    Case 0
        ' Ativa a janela caso esteja no Login
        If picLogin.Visible Then
            frmMenu.picLogin.Visible = False
            frmMenu.PicNewPass.Visible = True
            PlaySound Sound_ButtonClick, -1, -1
        End If

        ' Esquecer a senha
    Case 1
        ' Ativa a janela caso esteja no Login
        If picLogin.Visible Then
            frmMenu.picLogin.Visible = False
            frmMenu.PicRecover.Visible = True
            PlaySound Sound_ButtonClick, -1, -1
        End If

        ' Criar uma conta
    Case 2
        ' Ativa a janela caso esteja no Login
        If picLogin.Visible Then
            DestroyTCP
            frmMenu.picLogin.Visible = False
            frmMenu.picRegister.Visible = True
            PlaySound Sound_ButtonClick, -1, -1
        End If

    End Select

End Sub

Private Sub lblSprite_Click()
    Dim spritecount As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If optMale.value Then
        spritecount = UBound(Class(cmbClass.ListIndex + 1).MaleSprite)
    Else
        spritecount = UBound(Class(cmbClass.ListIndex + 1).FemaleSprite)
    End If

    If newCharSprite >= spritecount Then
        newCharSprite = 0
    Else
        newCharSprite = newCharSprite + 2
    End If

    'newCharSprite = 2

    NewCharacterBltSprite

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblSprite_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optFemale_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    NewCharacterBltSprite

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optFemale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMale_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    NewCharacterBltSprite

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''''''''
' CÓDIGO PERSONAGEM '
'''''''''''''''''''''

Private Sub picCharacter_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMenu.picCharacter, Button, Shift, X, Y)
    resetButtons_Menu

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

''''''''''''''''
' CÓDIGO LOGIN '
''''''''''''''''

Private Sub picLogin_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Ativa o movimento da janela
    Call MovePicture(frmMenu.picLogin, Button, Shift, X, Y)

    resetButtons_Menu

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picLogin_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    resetButtons_Menu

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picMain_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub PicPass_Click()

    If chkPass.value = 0 Then
        chkPass.value = 1
    Else
        chkPass.value = 0
    End If

    If chkPass.value = False Then
        PicPass.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\off.jpg")
    Else
        PicPass.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\on.jpg")
    End If
End Sub

'''''''''''''''''''''''
' CÓDIGO TROCAR SENHA '
'''''''''''''''''''''''

Private Sub PicNewPass_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicNewPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMenu.PicNewPass, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicNewPass_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

''''''''''''''''''''''''''
' CÓDIGO RECUPERAR SENHA '
''''''''''''''''''''''''''

Private Sub PicRecover_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub PicRecover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMenu.PicRecover, Button, Shift, X, Y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PicRecover_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'''''''''''''''''''
' CÓDIGO REGISTRO '
'''''''''''''''''''

Private Sub picRegister_MouseDown(Button As Integer, Deslocamento As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picRegister_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call MovePicture(frmMenu.picRegister, Button, Shift, X, Y)

    resetButtons_Menu

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picRegister_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' BOTÕES COM ANIMAÇÃO

Private Sub imgButton_Click(Index As Integer)
    Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    Dim RecoveryKey As String
    Dim Email As String
    Dim OldPassword As String
    Dim NewPassword As String
    Dim ReNewPassword As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case Index

    Case 1
        DestroyTCP
        PlaySound Sound_ButtonClick, -1, -1
        Call DestroyGame

    Case 2
        DestroyTCP

        If isLoginLegal(txtLUser.text, txtLPass.text) Then
            Call MenuState(MENU_STATE_LOGIN)
        End If

        ' play sound
        PlaySound Sound_ButtonClick, -1, -1

    Case 3
        Name = Trim$(txtRUser.text)
        Password = Trim$(txtRPass.text)
        PasswordAgain = Trim$(txtRPass2.text)
        RecoveryKey = Trim$(txtRecovery.text)
        Email = Trim$(txtEmail.text)

        If isLoginLegal(Name, Password) Then
            If Password <> PasswordAgain Then
                Call MsgBox("Passwords don't match.")
                Exit Sub
            End If

            If Not isStringLegal(Name) Then
                Exit Sub
            End If

            'RecoveryKey
            If LenB(RecoveryKey) <= 5 Then
                Call MsgBox("Recovery Key curta demais.")
                Exit Sub
            End If

            'Email Válidos
            If Email = vbNullString Then
                Call MsgBox("Insira um Email Válido! A Recuperação de informações será a partir dele!")
                Exit Sub
            Else
                If IsValidEmail(Email) = False Then
                    Call MsgBox("Insira um Email Válido! A Recuperação de informações será a partir dele!")
                    Exit Sub
                End If
            End If

            Call MenuState(MENU_STATE_NEWACCOUNT)
        End If

    Case 4
        DestroyTCP
        PlaySound Sound_ButtonClick, -1, -1

        picLogin.Visible = True
        picRegister.Visible = False

    Case 5
        Name = Trim$(txtNewPass(0).text)
        OldPassword = Trim$(txtNewPass(1).text)
        NewPassword = Trim$(txtNewPass(2).text)
        ReNewPassword = Trim$(txtNewPass(3).text)
        Email = Trim$(txtNewPass(4).text)

        'Usuario
        If Not isStringLegal(Name) Then
            Exit Sub
        End If

        'Senha Antiga
        If LenB(OldPassword) < 3 Then
            Call MsgBox("Nova senha curta de mais!")
            Exit Sub
        End If

        'Verificar nova Senha
        If Not UCase$(Trim$(NewPassword)) = UCase$(Trim$(ReNewPassword)) Then
            Call MsgBox("As senhas não correspondem!")
            Exit Sub
        End If

        If NewPassword = vbNullString Then
            Call MsgBox("Nova senha em branco!")
            Exit Sub
        End If

        If LenB(NewPassword) < 3 Then
            Call MsgBox("Nova senha curta de mais!")
            Exit Sub
        End If

        'Email Válido
        If Email = vbNullString Then
            Call MsgBox("O Email da conta é obrigatório!")
            Exit Sub
        Else
            If IsValidEmail(Email) = False Then
                Call MsgBox("Insira um Email Válido!")
                Exit Sub
            End If
        End If

        'Enviar Verificação da Confirmação da Senha!
        If ConnectToServer(1) Then
            Call SendNewPassword(Name, OldPassword, NewPassword, Email)
        End If

    Case 6
        DestroyTCP
        PlaySound Sound_ButtonClick, -1, -1

        picLogin.Visible = True
        PicNewPass.Visible = False

    Case 7
        Name = Trim$(txtRecover(0).text)
        RecoveryKey = Trim$(txtRecover(1).text)
        Email = Trim$(txtRecover(2).text)

        If Not isStringLegal(Name) Then
            Exit Sub
        End If

        'RecoveryKey
        If LenB(RecoveryKey) <= 5 Then
            Call MsgBox("Recovery Key curta demais.")
            Exit Sub
        End If

        'Email Válidos
        If Email = vbNullString Then
            Call MsgBox("O Email da conta é obrigatório!")
            Exit Sub
        Else
            If IsValidEmail(Email) = False Then
                Call MsgBox("Insira um Email Válido!")
                Exit Sub
            End If
        End If

        'Enviar Verificação da Confirmação da Senha!
        If ConnectToServer(1) Then
            Call SendRecoverPassword(Name, RecoveryKey, Email)
        End If


    Case 8
        DestroyTCP
        PlaySound Sound_ButtonClick, -1, -1

        picLogin.Visible = True
        PicRecover.Visible = False

    Case 9
        PlaySound Sound_ButtonClick, -1, -1
        Call MenuState(MENU_STATE_ADDCHAR)

    Case 10
        DestroyTCP
        picCharacter.Visible = False
        picLogin.Visible = True
        PlaySound Sound_ButtonClick, -1, -1

    Case 11
        DestroyTCP
        DestroyGame

    Case 12
        PicTermos.Visible = False
        picLogin.Visible = True
        Options.Termos = 1
        SaveOptions
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset other buttons
    resetButtons_Menu Index

    ' change the button we're hovering on
    changeButtonState_Menu Index, 2    ' clicked

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset other buttons
    resetButtons_Menu Index

    ' change the button we're hovering on
    If Not MenuButton(Index).state = 2 Then    ' make sure we're not clicking
        changeButtonState_Menu Index, 1    ' hover
    End If

    ' play sound
    If Not LastButtonSound_Menu = Index Then
        PlaySound Sound_ButtonHover, -1, -1
        LastButtonSound_Menu = Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' reset all buttons
    resetButtons_Menu -1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub tmrStatus_Timer()
If ConnectToServer(1) Then
SendRequestStatus
Else
lblStatus.Caption = "Offline"
lblStatus.ForeColor = QBColor(BrightRed)
End If
End Sub

' Se pressionar enter

Private Sub txtLUser_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If isLoginLegal(txtLUser.text, txtLPass.text) Then
            Call MenuState(MENU_STATE_LOGIN)
        End If

        KeyAscii = 0
    End If

End Sub

Private Sub txtLPass_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If isLoginLegal(txtLUser.text, txtLPass.text) Then
            Call MenuState(MENU_STATE_LOGIN)
        End If

        KeyAscii = 0
    End If

End Sub

Private Sub txtCName_KeyPress(KeyAscii As Integer)
    Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    Dim RecoveryKey As String
    Dim Email As String

    If KeyAscii = vbKeyReturn Then

        PlaySound Sound_ButtonClick, -1, -1
        Call MenuState(MENU_STATE_ADDCHAR)

        KeyAscii = 0
    End If

End Sub

Private Sub txtRecovery_KeyPress(KeyAscii As Integer)

    Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    Dim RecoveryKey As String
    Dim Email As String
    
        Name = Trim$(txtRUser.text)
        Password = Trim$(txtRPass.text)
        PasswordAgain = Trim$(txtRPass2.text)
        RecoveryKey = Trim$(txtRecovery.text)
        Email = Trim$(txtEmail.text)
        
        If KeyAscii = vbKeyReturn Then
        If isLoginLegal(Name, Password) Then
            If Password <> PasswordAgain Then
                Call MsgBox("Passwords don't match.")
                Exit Sub
            End If

            If Not isStringLegal(Name) Then
                Exit Sub
            End If

            'RecoveryKey
            If LenB(RecoveryKey) <= 5 Then
                Call MsgBox("Recovery Key curta demais.")
                Exit Sub
            End If

            'Email Válidos
            If Email = vbNullString Then
                Call MsgBox("Insira um Email Válido! A Recuperação de informações será a partir dele!")
                Exit Sub
            Else
                If IsValidEmail(Email) = False Then
                    Call MsgBox("Insira um Email Válido! A Recuperação de informações será a partir dele!")
                    Exit Sub
                End If
            End If

            Call MenuState(MENU_STATE_NEWACCOUNT)
        End If
End If
End Sub
