Attribute VB_Name = "modClientTCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set PlayerBuffer = New clsBuffer

    ' connect
    frmMain.Socket.RemoteHost = "127.0.0.1"
    frmMain.Socket.RemotePort = "7777"

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TcpInit", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DestroyTCP()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.Socket.Close

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyTCP", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.Socket.GetData Buffer, vbUnicode, DataLength

    PlayerBuffer.WriteBytes Buffer()

    If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.length - 4
        If pLength <= PlayerBuffer.length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "IncomingData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConnectToServer(ByVal i As Long) As Boolean
    Dim Wait As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If

    Wait = GetTickCount
    frmMain.Socket.Close
    frmMain.Socket.Connect

    SetStatus "Connecting to server..."

    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop

    ConnectToServer = IsConnected

    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConnectToServer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsConnected() As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsConnected", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlaying", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SendData(ByRef data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Set Buffer = New clsBuffer

        Buffer.WriteLong (UBound(data) - LBound(data)) + 1
        Buffer.WriteBytes data()
        frmMain.Socket.SendData Buffer.ToArray()
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal Name As String, ByVal Password As String, ByVal RecoveryKey As String, ByVal Email As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CNewAccount
    Buffer.WriteString Name
    Buffer.WriteString Password
    Buffer.WriteString RecoveryKey
    Buffer.WriteString Email
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendNewAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CDelAccount
    Buffer.WriteString Name
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDelAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CLogin
    Buffer.WriteString Name
    Buffer.WriteString Password
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendLogin", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAddChar
    Buffer.WriteString Name
    Buffer.WriteLong Sex
    Buffer.WriteLong ClassNum
    Buffer.WriteLong Sprite
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAddChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseChar
    Buffer.WriteLong CharSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUseChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SayMsg(ByVal text As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SayMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BroadcastMsg(ByVal text As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBroadcastMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BroadcastMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EmoteMsg(ByVal text As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CEmoteMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EmoteMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerMsg(ByVal text As String, ByVal MsgTo As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerMove()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong Player(MyIndex).Moving
    Buffer.WriteLong Player(MyIndex).X
    Buffer.WriteLong Player(MyIndex).Y
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerMove", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerDir()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerDir
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerDir", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerRequestNewMap()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNewMap
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerRequestNewMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMap()
    Dim packet As String
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    CanMoveNow = False

    With Map
        Buffer.WriteLong CMapData
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteString Trim$(.Music)
        Buffer.WriteByte .Moral
        Buffer.WriteLong .Up
        Buffer.WriteLong .Down
        Buffer.WriteLong .Left
        Buffer.WriteLong .Right
        Buffer.WriteLong .BootMap
        Buffer.WriteByte .BootX
        Buffer.WriteByte .BootY
        Buffer.WriteByte .MaxX
        Buffer.WriteByte .MaxY
        Buffer.WriteLong .Weather
        Buffer.WriteLong .Intensity
        For X = 1 To 2
            Buffer.WriteLong .LevelPoke(X)
        Next
    End With

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY

            With Map.Tile(X, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).X
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

    With Map

        For X = 1 To MAX_MAP_NPCS
            Buffer.WriteLong .Npc(X)
        Next

    End With

    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpMeTo(ByVal Name As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpMeTo
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpMeTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpToMe(ByVal Name As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpToMe
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarptoMe", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpTo
    Buffer.WriteLong MapNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetAccess
    Buffer.WriteString Name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetAccess", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetSprite(ByVal Name As String, ByVal SpriteNum As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetSprite
    Buffer.WriteString Trim$(Name)
    Buffer.WriteLong SpriteNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetSprite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetVisual(ByVal Name As String, ByVal VisualNum As Byte, ByVal VisualPart As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetVisual
    Buffer.WriteString Trim$(Name)
    Buffer.WriteLong VisualNum
    Buffer.WriteByte VisualPart
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetVisual", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendKick(ByVal Name As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CKickPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendKick", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBan(ByVal Name As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBan", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanList()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanList
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanList", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditItem()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditItem
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveItem(ByVal itemNum As Long)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    Buffer.WriteLong CSaveItem
    Buffer.WriteLong itemNum
    Buffer.WriteBytes ItemData
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditAnimation()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditAnimation
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    Buffer.WriteLong CSaveAnimation
    Buffer.WriteLong Animationnum
    Buffer.WriteBytes AnimationData
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditNpc()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNpc
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    Buffer.WriteLong CSaveNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes NpcData
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditResource()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditResource
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    Buffer.WriteLong CSaveResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapRespawn()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapRespawn
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapRespawn", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseItem(ByVal InvNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseItem
    Buffer.WriteLong InvNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUseItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InBank Or InShop Then Exit Sub

    ' do basic checks
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    If PlayerInv(InvNum).num < 1 Or PlayerInv(InvNum).num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Or Amount > PlayerInv(InvNum).value Then Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapDropItem
    Buffer.WriteLong InvNum
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDropItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendWhosOnline()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWhosOnline
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendWhosOnline", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetMotd
    Buffer.WriteString MOTD
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMOTDChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditShop()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveShop(ByVal shopnum As Long)
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    Buffer.WriteLong CSaveShop
    Buffer.WriteLong shopnum
    Buffer.WriteBytes ShopData
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditSpell()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditSpell
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveSpell(ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize

    Buffer.WriteLong CSaveSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    SendData Buffer.ToArray()

    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditMap()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanDestroy()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanDestroy
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanDestroy", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapInvSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapSpellSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub GetPing()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    PingStart = GetTickCount
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCheckPing
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GetPing", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUnequip(ByVal EqNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUnequip
    Buffer.WriteLong EqNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUnequip", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestPlayerData()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPlayerData
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestPlayerData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestItems()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestItems
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestItems", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestAnimations()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestAnimations
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestAnimations", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestNPCS()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNPCS
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestNPCS", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestResources()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestResources
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestResources", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestSpells()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestSpells
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestSpells", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestShops()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestShops
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestShops", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long, ByVal Command As Byte, Optional ByVal Pokemon As Long, Optional ByVal Pokeball As Long, Optional ByVal Level As Long, Optional ByVal Exp As Long, Optional ByVal VitalHp As Long, Optional ByVal VitalMp As Long, Optional ByVal MaxVitalHp As Long, Optional ByVal MaxVitalMp As Long, Optional ByVal StatStr As Long, Optional ByVal StatAgi As Long, Optional ByVal StatEnd As Long, Optional ByVal StatInt As Long, Optional ByVal StatWill As Long, Optional ByVal Spell1 As Long, Optional ByVal Spell2 As Long, Optional ByVal Spell3 As Long, Optional ByVal Spell4 As Long, Optional ByVal Felicidade As Long, Optional ByVal Sexo As Byte, Optional ByVal Shiny As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Command > 1 Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSpawnItem

    Buffer.WriteByte Command

    If Command = 0 Then
        Buffer.WriteLong tmpItem
        Buffer.WriteLong tmpAmount
    Else
        Buffer.WriteLong Pokemon
        Buffer.WriteLong Pokeball
        Buffer.WriteLong Level
        Buffer.WriteLong Exp

        Buffer.WriteLong VitalHp
        Buffer.WriteLong VitalMp
        Buffer.WriteLong MaxVitalHp
        Buffer.WriteLong MaxVitalMp

        Buffer.WriteLong StatStr
        Buffer.WriteLong StatEnd
        Buffer.WriteLong StatInt
        Buffer.WriteLong StatAgi
        Buffer.WriteLong StatWill

        Buffer.WriteLong Spell1
        Buffer.WriteLong Spell2
        Buffer.WriteLong Spell3
        Buffer.WriteLong Spell4

        Buffer.WriteLong Felicidade
        Buffer.WriteByte Sexo
        Buffer.WriteByte Shiny

    End If

    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSpawnItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTrainStat(ByVal StatNum As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseStatPoint
    Buffer.WriteByte StatNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestLevelUp()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestLevelUp
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestLevelUp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BuyItem(ByVal shopslot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyItem
    Buffer.WriteLong shopslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BuyItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SellItem(ByVal invslot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSellItem
    Buffer.WriteLong invslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SellItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DepositItem(ByVal invslot As Long, ByVal Amount As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CDepositItem
    Buffer.WriteLong invslot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DepositItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal Amount As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CWithdrawItem
    Buffer.WriteLong bankslot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WithdrawItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseBank()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseBank
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    InBank = False
    frmMain.picCover.Visible = False
    frmMain.picBank.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CloseBank", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChangeBankSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CChangeBankSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChangeBankSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AdminWarp(ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAdminWarp
    Buffer.WriteLong X
    Buffer.WriteLong Y
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AdminWarp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AcceptTrade()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AcceptTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DeclineTrade()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler


    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DeclineTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub TradeItem(ByVal invslot As Long, ByVal Amount As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeItem
    Buffer.WriteLong invslot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UntradeItem(ByVal invslot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CUntradeItem
    Buffer.WriteLong invslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UntradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarChange
    Buffer.WriteLong sType
    Buffer.WriteLong Slot
    Buffer.WriteLong hotbarNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
    Dim Buffer As clsBuffer
    Dim X As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' check if spell
    If Hotbar(Slot).sType = 2 Then    ' spell
        For X = 1 To MAX_PLAYER_SPELLS
            ' is the spell matching the hotbar?
            If PlayerSpells(X) = Hotbar(Slot).Slot Then
                ' found it, cast it
                CastSpell X
                Exit Sub
            End If
        Next
        ' can't find the spell, exit out
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarUse
    Buffer.WriteLong Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarUse", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapReport()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapReport
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapReport", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSearch(ByVal CurX As Long, ByVal CurY As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If isInBounds Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CSearch
        Buffer.WriteLong CurX
        Buffer.WriteLong CurY
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerSearch", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeRequest()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptTradeRequest()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineTradeRequest()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyLeave()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyLeave
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyLeave", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyRequest()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptParty()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineParty()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMutePlayer(ByVal Name As String, ByVal Tempo As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CMutePlayer

    Buffer.WriteString Name
    Buffer.WriteLong Tempo

    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendEvolCommand(ByVal Command As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CEvolCommand

    Buffer.WriteLong Command

    If Command = 1 Then
        Player(MyIndex).EvolPermition = 0
    End If

    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub


'#######################
'########QUESTS#########
'#######################

Sub SendRequestQuests()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ClientPackets.CRequestQuests
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestQuests", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendQuestCommand(ByVal Command As Byte, Optional value As Long = 0)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ClientPackets.CQuestCommand
    Buffer.WriteByte Command
    Buffer.WriteLong value
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendQuestCommand", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditQuest()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ClientPackets.CRequestEditQuest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditQuest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveQuest(ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong ClientPackets.CSaveQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveQuest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendComprar(ByVal LeilaoNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CComprar
    Buffer.WriteLong LeilaoNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendComprar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Sub SendRetirar(ByVal LeilaoNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRetirar
    Buffer.WriteLong LeilaoNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRetirar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendLeiloar(ByVal InvNum As Long, ByVal Price As Long, ByVal Tempo As Long, ByVal Tipo As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CLeiloar
    Buffer.WriteLong InvNum
    Buffer.WriteLong Price
    Buffer.WriteLong Tempo
    Buffer.WriteLong Tipo
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendLeiloar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChatComando(ByVal c As Long, S As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CChatComando
    Buffer.WriteLong c
    Buffer.WriteString S
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChatComando", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSurfInit(ByVal Command As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSendSurfInit
    Buffer.WriteLong Command
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSurfInit", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendLutarComando(c, T, a, Pok As Long, P As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CLutarComando
    Buffer.WriteLong c
    Buffer.WriteLong T
    Buffer.WriteLong a
    Buffer.WriteLong Pok
    Buffer.WriteString P
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendAprenderHab(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAprender
    Buffer.WriteByte Index
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub SendSetOrganization(ByVal Name As String, ByVal Access As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetOrg
    Buffer.WriteString Name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAbrir()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAbrir
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendBuyOrgShop(ByVal OrgShopSlot As Byte, ByVal Quantia As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyOrgShop
    Buffer.WriteByte OrgShopSlot
    Buffer.WriteLong Quantia
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRecoverPassword(ByVal Account As String, ByVal RecoveryKey As String, ByVal Email As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRecoverPass
    Buffer.WriteString Account
    Buffer.WriteString RecoveryKey
    Buffer.WriteString Email
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNewPassword(ByVal Account As String, ByVal OldPassword As String, ByVal NewPassword As String, ByVal Email As String)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CNewPass
    Buffer.WriteString Account
    Buffer.WriteString OldPassword
    Buffer.WriteString NewPassword
    Buffer.WriteString Email
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendObterPacVip(ByVal VipNum As Byte, ByVal VipView As Byte, Optional ByVal ViewVip As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CGiveVip
    Buffer.WriteByte VipNum
    Buffer.WriteByte VipView
    Buffer.WriteByte ViewVip
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerRun(ByVal Correr As Boolean)
    Dim Buffer As clsBuffer
    Player(MyIndex).Running = Correr

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerRun
    If Correr = True Then Buffer.WriteByte 1
    If Correr = False Then Buffer.WriteByte 0
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendComandoGym(ByVal Comando As Byte)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CComandGym
    Buffer.WriteByte Comando
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub GrupoMsg(ByVal text As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CGrupoMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "VilaMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
