VERSION 5.00
Begin VB.Form frmEditor_Chat 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Privado"
   ClientHeight    =   3855
   ClientLeft      =   3975
   ClientTop       =   3630
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4215
   Begin VB.TextBox txtEChat 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   3975
   End
   Begin VB.TextBox txtChat 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblJogador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "- Sem Jogador -"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmEditor_Chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label2_Click()
    Me.Hide
    txtEChat.text = vbNullString
End Sub

Private Sub txtEChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtEChat.text)) > 0 Then
            SendChatComando 4, Trim$(txtEChat.text)
        End If

        KeyAscii = 0
    End If

End Sub

Private Sub txtChat_Click()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    txtEChat.SetFocus

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChat_Click", "frmEditor_Chat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendChatComando 5, vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
