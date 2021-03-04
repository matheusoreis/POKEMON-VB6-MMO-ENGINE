Attribute VB_Name = "modDirectDraw7"
Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectDraw7 Object
Public DD As DirectDraw7
' Clipper object
Public DD_Clip As DirectDrawClipper

' primary surface
Public DDS_Primary As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

' back buffer
Public DDS_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

' Used for pre-rendering
Public DDS_Map As DirectDrawSurface7
Public DDSD_Map As DDSURFACEDESC2

' Chat Bubble Mondo
Public DDS_ChatBubble As DirectDrawSurface7
Public DDSD_ChatBubble As DDSURFACEDESC2

' gfx buffers
Public DDS_Item() As DirectDrawSurface7    ' arrays
Public DDS_Character() As DirectDrawSurface7
Public DDS_Paperdoll() As DirectDrawSurface7
Public DDS_Tileset() As DirectDrawSurface7
Public DDS_Resource() As DirectDrawSurface7
Public DDS_Animation() As DirectDrawSurface7
Public DDS_SpellIcon() As DirectDrawSurface7
Public DDS_Face() As DirectDrawSurface7
Public DDS_Door As DirectDrawSurface7    ' singes
Public DDS_Misc As DirectDrawSurface7
Public DDS_Direction As DirectDrawSurface7
Public DDS_Target As DirectDrawSurface7
Public DDS_Bars As DirectDrawSurface7
Public DDS_PokeIcons() As DirectDrawSurface7
Public DDS_PokeIconShiny() As DirectDrawSurface7
Public DDS_SaN As DirectDrawSurface7
Public DDS_Weather As DirectDrawSurface7
Public DDS_Quest As DirectDrawSurface7
Public DDS_RockTunel As DirectDrawSurface7
Public DDS_MiniMap As DirectDrawSurface7
Public DDS_MapBorda As DirectDrawSurface7
Public DDS_OrgOn As DirectDrawSurface7
Public DDS_Hair() As DirectDrawSurface7

' descriptions
Public DDSD_Temp As DDSURFACEDESC2    ' arrays
Public DDSD_Item() As DDSURFACEDESC2
Public DDSD_Character() As DDSURFACEDESC2
Public DDSD_Paperdoll() As DDSURFACEDESC2
Public DDSD_Tileset() As DDSURFACEDESC2
Public DDSD_Resource() As DDSURFACEDESC2
Public DDSD_Animation() As DDSURFACEDESC2
Public DDSD_SpellIcon() As DDSURFACEDESC2
Public DDSD_Face() As DDSURFACEDESC2
Public DDSD_Door As DDSURFACEDESC2    ' singles
Public DDSD_Misc As DDSURFACEDESC2
Public DDSD_Direction As DDSURFACEDESC2
Public DDSD_Target As DDSURFACEDESC2
Public DDSD_Bars As DDSURFACEDESC2
Public DDSD_PokeIcons() As DDSURFACEDESC2
Public DDSD_PokeIconShiny() As DDSURFACEDESC2
Public DDSD_SaN As DDSURFACEDESC2
Public DDSD_Weather As DDSURFACEDESC2
Public DDSD_Quest As DDSURFACEDESC2
Public DDSD_RockTunel As DDSURFACEDESC2
Public DDSD_MiniMap As DDSURFACEDESC2
Public DDSD_MapBorda As DDSURFACEDESC2
Public DDSD_OrgOn As DDSURFACEDESC2
Public DDSD_Hair() As DDSURFACEDESC2

' timers
Public Const SurfaceTimerMax As Long = 10000
Public CharacterTimer() As Long
Public PaperdollTimer() As Long
Public ItemTimer() As Long
Public ResourceTimer() As Long
Public AnimationTimer() As Long
Public SpellIconTimer() As Long
Public FaceTimer() As Long
Public PokeIconTimer() As Long
Public PokeIconShinyTimer() As Long
Public HairTimer() As Long

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public numitems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumPokeIcons As Long
Public NumPokeIconShiny As Long
Public HairNum As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDirectDraw() As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear DD7
    Call DestroyDirectDraw

    ' Init Direct Draw
    Set DD = DX7.DirectDrawCreate(vbNullString)

    ' Windowed
    DD.SetCooperativeLevel frmMain.hwnd, DDSCL_NORMAL

    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        .lBackBufferCount = 1
    End With
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)

    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)

    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMain.picScreen.hwnd

    ' Have the blits to the screen clipped to the picture box
    DDS_Primary.SetClipper DD_Clip

    ' Initialise the surfaces
    InitSurfaces

    ' We're done
    InitDirectDraw = True

    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub InitSurfaces()
    Dim rec As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' DirectDraw Surface memory management setting
    DDSD_Temp.lFlags = DDSD_CAPS
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY

    ' clear out everything for re-init
    Set DDS_BackBuffer = Nothing

    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (MAX_MAPX + 3) * PIC_X
        .lHeight = (MAX_MAPY + 3) * PIC_Y
    End With
    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)

    ' load persistent surfaces
    If FileExist(App.Path & "\data files\graphics\door.bmp", True) Then Call InitDDSurf("door", DDSD_Door, DDS_Door)
    If FileExist(App.Path & "\data files\graphics\direction.bmp", True) Then Call InitDDSurf("direction", DDSD_Direction, DDS_Direction)
    If FileExist(App.Path & "\data files\graphics\target.bmp", True) Then Call InitDDSurf("target", DDSD_Target, DDS_Target)
    If FileExist(App.Path & "\data files\graphics\misc.bmp", True) Then Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    If FileExist(App.Path & "\data files\graphics\bars.bmp", True) Then Call InitDDSurf("bars", DDSD_Bars, DDS_Bars)
    If FileExist(App.Path & "\data files\graphics\SexAndNegatives.bmp", True) Then Call InitDDSurf("SexAndNegatives", DDSD_SaN, DDS_SaN)
    If FileExist(App.Path & "\data files\graphics\Weather.bmp", True) Then Call InitDDSurf("Weather", DDSD_Weather, DDS_Weather)
    If FileExist(App.Path & "\data files\graphics\Quest.bmp", True) Then Call InitDDSurf("Quest", DDSD_Quest, DDS_Quest)
    If FileExist(App.Path & "\data files\graphics\RockTunelEffect.bmp", True) Then Call InitDDSurf("RockTunelEffect", DDSD_RockTunel, DDS_RockTunel)
    If FileExist(App.Path & "\data files\graphics\minimap.bmp", True) Then Call InitDDSurf("minimap", DDSD_MiniMap, DDS_MiniMap)
    If FileExist(App.Path & "\data files\graphics\MapBorda.bmp", True) Then Call InitDDSurf("MapBorda", DDSD_MapBorda, DDS_MapBorda)
    If FileExist(App.Path & "\data files\graphics\StatusOrg.bmp", True) Then Call InitDDSurf("StatusOrg", DDSD_OrgOn, DDS_OrgOn)
    If FileExist(App.Path & "\data files\graphics\chatbubble.bmp", True) Then Call InitDDSurf("chatbubble", DDSD_ChatBubble, DDS_ChatBubble)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
    Dim TmpR As RECT
    Dim TmpDDSD As DDSURFACEDESC2
    Dim TmpColorKey As DDCOLORKEY

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With TmpR
        .Left = X
        .top = Y
        .Right = X
        .Bottom = Y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(X, Y)
        .High = .Low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    TheSurface.Unlock TmpR

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetMaskColorFromPixel", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(filename As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Set path
    filename = App.Path & GFX_PATH & filename & GFX_EXT

    ' Destroy surface if it exist
    If Not Surf Is Nothing Then
        Set Surf = Nothing
        Call ZeroMemory(ByVal VarPtr(SurfDesc), LenB(SurfDesc))
    End If

    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps

    ' init object
    Set Surf = DD.CreateSurfaceFromFile(filename, SurfDesc)

    ' Set mask
    Call SetMaskColorFromPixel(Surf, 0, 0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitDDSurf", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function CheckSurfaces() As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if we need to restore surfaces
    If Not DD.TestCooperativeLevel = DD_OK Then
        CheckSurfaces = False
    Else
        CheckSurfaces = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CheckSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function NeedToRestoreSurfaces() As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not DD.TestCooperativeLevel = DD_OK Then
        NeedToRestoreSurfaces = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "NeedToRestoreSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub ReInitDD()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call InitDirectDraw

    LoadTilesets

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ReInitDD", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyDirectDraw()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Unload DirectDraw
    Set DDS_Misc = Nothing

    For i = 1 To NumTileSets
        Set DDS_Tileset(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i))
    Next

    For i = 1 To numitems
        Set DDS_Item(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i))
    Next

    For i = 1 To NumCharacters
        Set DDS_Character(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i))
    Next

    For i = 1 To NumPaperdolls
        Set DDS_Paperdoll(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i))
    Next

    For i = 1 To NumResources
        Set DDS_Resource(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i))
    Next

    For i = 1 To NumAnimations
        Set DDS_Animation(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i))
    Next

    For i = 1 To NumSpellIcons
        Set DDS_SpellIcon(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i))
    Next

    For i = 1 To NumFaces
        Set DDS_Face(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Face(i)), LenB(DDSD_Face(i))
    Next

    For i = 1 To NumPokeIcons
        Set DDS_PokeIcons(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_PokeIcons(i)), LenB(DDSD_PokeIcons(i))
    Next

    For i = 1 To NumPokeIconShiny
        Set DDS_PokeIconShiny(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_PokeIconShiny(i)), LenB(DDSD_PokeIconShiny(i))
    Next

    For i = 1 To HairNum
        Set DDS_Hair(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Hair(i)), LenB(DDSD_Hair(i))
    Next

    Set DDS_Door = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Door), LenB(DDSD_Door)

    Set DDS_Direction = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Direction), LenB(DDSD_Direction)

    Set DDS_Target = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Target), LenB(DDSD_Target)

    Set DDS_SaN = Nothing
    ZeroMemory ByVal VarPtr(DDSD_SaN), LenB(DDSD_SaN)

    Set DDS_Weather = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Weather), LenB(DDSD_Weather)

    Set DDS_Quest = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Quest), LenB(DDSD_Quest)

    Set DDS_RockTunel = Nothing
    ZeroMemory ByVal VarPtr(DDSD_RockTunel), LenB(DDSD_RockTunel)

    Set DDS_MiniMap = Nothing
    ZeroMemory ByVal VarPtr(DDSD_MiniMap), LenB(DDSD_MiniMap)

    Set DDS_MapBorda = Nothing
    ZeroMemory ByVal VarPtr(DDSD_MapBorda), LenB(DDSD_MapBorda)

    Set DDS_OrgOn = Nothing
    ZeroMemory ByVal VarPtr(DDSD_OrgOn), LenB(DDSD_OrgOn)

    Set DDS_ChatBubble = Nothing
    ZeroMemory ByVal VarPtr(DDSD_ChatBubble), LenB(DDSD_ChatBubble)

    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    Set DD_Clip = Nothing
    Set DD = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Blitting **
' **************
Public Sub Engine_BltFast(ByVal dx As Long, ByVal dy As Long, ByRef ddS As DirectDrawSurface7, srcRECT As RECT, trans As CONST_DDBLTFASTFLAGS)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler


    If Not ddS Is Nothing Then
        Call DDS_BackBuffer.BltFast(dx, dy, ddS, srcRECT, trans)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Engine_BltFast", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Engine_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Clear Then
        picBox.Cls
    End If

    Call Surface.BltToDC(picBox.hDC, sRECT, dRECT)
    picBox.Refresh
    Engine_BltToDC = True

    ' Error handler
    Exit Function
errorhandler:
    HandleError "Engine_BltToDC", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub BltDirection(ByVal X As Long, ByVal Y As Long)
    Dim rec As DxVBLib.RECT
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' render grid
    rec.top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.Bottom = rec.top + 32
    Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' render dir blobs
    For i = 1 To 4
        rec.Left = (i - 1) * 8
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(i)) Then
            rec.top = 8
        Else
            rec.top = 16
        End If
        rec.Bottom = rec.top + 8
        'render!
        Call Engine_BltFast(ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDirection", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTarget(ByVal X As Long, ByVal Y As Long)
    Dim sRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub

    Width = DDSD_Target.lWidth / 2
    Height = DDSD_Target.lHeight

    With sRECT
        .top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With

    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    ' clipping
    If Y < 0 Then
        With sRECT
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping

    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTarget", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltHover(ByVal tType As Long, ByVal target As Long, ByVal X As Long, ByVal Y As Long)
    Dim sRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub

    Width = DDSD_Target.lWidth / 2
    Height = DDSD_Target.lHeight

    With sRECT
        .top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With

    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    ' clipping
    If Y < 0 Then
        With sRECT
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping

    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHover", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapTile(ByVal X As Long, ByVal Y As Long)
    Dim rec As DxVBLib.RECT
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile?
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                ' sort out rec
                rec.top = .Layer(i).Y * PIC_Y
                rec.Bottom = rec.top + PIC_Y
                rec.Left = .Layer(i).X * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next
    End With

    ' Error handler
    Exit Sub

errorhandler:
    HandleError "BltMapTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapFringeTile(ByVal X As Long, ByVal Y As Long)
    Dim rec As DxVBLib.RECT
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' skip tile if tileset isn't set
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                ' sort out rec
                rec.top = .Layer(i).Y * PIC_Y
                rec.Bottom = rec.top + PIC_Y
                rec.Left = .Layer(i).X * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapFringeTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltDoor(ByVal X As Long, ByVal Y As Long)
    Dim rec As DxVBLib.RECT
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' sort out animation
    With TempTile(X, Y)
        If .DoorAnimate = 1 Then    ' opening
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame < 4 Then
                    .DoorFrame = .DoorFrame + 1
                Else
                    .DoorAnimate = 2    ' set to closing
                End If
                .DoorTimer = GetTickCount
            End If
        ElseIf .DoorAnimate = 2 Then    ' closing
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame > 1 Then
                    .DoorFrame = .DoorFrame - 1
                Else
                    .DoorAnimate = 0    ' end animation
                End If
                .DoorTimer = GetTickCount
            End If
        End If

        If .DoorFrame = 0 Then .DoorFrame = 1
    End With

    With rec
        .top = 0
        .Bottom = DDSD_Door.lHeight
        .Left = ((TempTile(X, Y).DoorFrame - 1) * (DDSD_Door.lWidth / 4))
        .Right = .Left + (DDSD_Door.lWidth / 4)
    End With

    X2 = (X * PIC_X)
    Y2 = (Y * PIC_Y) - (DDSD_Door.lHeight / 2) + 4
    Call DDS_BackBuffer.BltFast(ConvertMapX(X2), ConvertMapY(Y2), DDS_Door, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDoor", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltAnimation(ByVal Index As Long, ByVal Layer As Long)
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim i As Long
    Dim Width As Long, Height As Long
    Dim looptime As Long
    Dim FrameCount As Long
    Dim X As Long, Y As Long
    Dim lockindex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If

    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)

    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub

    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)

    AnimationTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Animation(Sprite) Is Nothing Then
        Call InitDDSurf("animations\" & Sprite, DDSD_Animation(Sprite), DDS_Animation(Sprite))
    End If

    ' total width divided by frame count
    Width = DDSD_Animation(Sprite).lWidth / FrameCount
    Height = DDSD_Animation(Sprite).lHeight

    sRECT.top = 0
    sRECT.Bottom = Height
    sRECT.Left = (AnimInstance(Index).FrameIndex(Layer) - 1) * Width
    sRECT.Right = sRECT.Left + Width

    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then    ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).XOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).YOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).XOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).YOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    ' Clip to screen
    If Y < 0 Then

        With sRECT
            .top = .top - Y
        End With

        Y = 0
    End If

    If X < 0 Then

        With sRECT
            .Left = .Left - X
        End With

        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If

    Call Engine_BltFast(X, Y, DDS_Animation(Sprite), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimation", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItem(ByVal ItemNum As Long)
    Dim PicNum As Long
    Dim rec As DxVBLib.RECT
    Dim MaxFrames As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' get the picture
    PicNum = Item(MapItem(ItemNum).num).Pic

    If MapItem(ItemNum).num = 3 Then
        Select Case MapItem(ItemNum).PokeInfo.Pokeball
        Case 1
            PicNum = 2
        Case 2
            PicNum = 3
        Case 3
            PicNum = 4
        Case 4
            PicNum = 5
        Case 5
            PicNum = 6
        Case 6
            PicNum = 7
        End Select
    End If

    If PicNum < 1 Or PicNum > numitems Then Exit Sub
    ItemTimer(PicNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(PicNum) Is Nothing Then
        Call InitDDSurf("items\" & PicNum, DDSD_Item(PicNum), DDS_Item(PicNum))
    End If

    If DDSD_Item(PicNum).lWidth > 64 Then    ' has more than 1 frame
        With rec
            .top = 0
            .Bottom = 32
            .Left = (MapItem(ItemNum).Frame * 32)
            .Right = .Left + 32
        End With
    Else
        With rec
            .top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If

    Call Engine_BltFast(ConvertMapX(MapItem(ItemNum).X * PIC_X), ConvertMapY(MapItem(ItemNum).Y * PIC_Y), DDS_Item(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ScreenshotMap()
    Dim X As Long, Y As Long, i As Long, rec As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' clear the surface
    Set DDS_Map = Nothing

    ' Initialize it
    With DDSD_Map
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (Map.MaxX + 1) * 32
        .lHeight = (Map.MaxY + 1) * 32
    End With
    Set DDS_Map = DD.CreateSurface(DDSD_Map)

    ' render the tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For i = MapLayer.Ground To MapLayer.Mask2
                    ' skip tile?
                    If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                        ' sort out rec
                        rec.top = .Layer(i).Y * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        rec.Left = .Layer(i).X * PIC_X
                        rec.Right = rec.Left + PIC_X
                        ' render
                        DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Next
            End With
        Next
    Next

    ' render the resources
    For Y = 0 To Map.MaxY
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
                            Call BltMapResource(i, True)
                        End If
                    Next
                End If
            End If
        End If
    Next

    ' render the tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For i = MapLayer.Fringe To MapLayer.Fringe2
                    ' skip tile?
                    If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                        ' sort out rec
                        rec.top = .Layer(i).Y * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        rec.Left = .Layer(i).X * PIC_X
                        rec.Right = rec.Left + PIC_X
                        ' render
                        DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Next
            End With
        Next
    Next

    ' dump and save
    frmMain.picSSMap.Width = DDSD_Map.lWidth
    frmMain.picSSMap.Height = DDSD_Map.lHeight
    rec.top = 0
    rec.Left = 0
    rec.Bottom = DDSD_Map.lHeight
    rec.Right = DDSD_Map.lWidth
    Engine_BltToDC DDS_Map, rec, rec, frmMain.picSSMap
    SavePicture frmMain.picSSMap.Image, App.Path & "\map" & GetPlayerMap(MyIndex) & ".jpg"

    ' let them know we did it
    AddText "Screenshot of map #" & GetPlayerMap(MyIndex) & " saved.", BrightGreen

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotMap", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapResource(ByVal Resource_num As Long, Optional ByVal screenShot As Boolean = False)
    Dim Resource_master As Long
    Dim Resource_state As Long
    Dim Resource_sprite As Long
    Dim rec As DxVBLib.RECT
    Dim X As Long, Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub

    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).Data1

    If Resource_master = 0 Then Exit Sub
    '    If Resource_master > 255 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then    ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then    ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' Load early
    If DDS_Resource(Resource_sprite) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource_sprite, DDSD_Resource(Resource_sprite), DDS_Resource(Resource_sprite))
    End If

    ' src rect
    With rec
        .top = 0
        .Bottom = DDSD_Resource(Resource_sprite).lHeight
        .Left = 0
        .Right = DDSD_Resource(Resource_sprite).lWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (DDSD_Resource(Resource_sprite).lWidth / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - DDSD_Resource(Resource_sprite).lHeight + 32

    ' render it
    If Not screenShot Then
        Call BltResource(Resource_sprite, X, Y, rec)
    Else
        Call ScreenshotResource(Resource_sprite, X, Y, rec)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltResource(ByVal Resource As Long, ByVal dx As Long, dy As Long, rec As DxVBLib.RECT)
    Dim X As Long
    Dim Y As Long
    Dim Width As Long
    Dim Height As Long
    Dim destRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    ResourceTimer(Resource) = GetTickCount + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If

    X = ConvertMapX(dx)
    Y = ConvertMapY(dy)

    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.top)

    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If

    ' End clipping
    Call Engine_BltFast(X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScreenshotResource(ByVal Resource As Long, ByVal X As Long, Y As Long, rec As DxVBLib.RECT)
    Dim Width As Long
    Dim Height As Long
    Dim destRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    ResourceTimer(Resource) = GetTickCount + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If

    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.top)

    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_Map.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_Map.lHeight)
    End If

    If X + Width > DDSD_Map.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_Map.lWidth)
    End If

    ' End clipping
    'Call Engine_BltFast(x, y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    DDS_Map.BltFast X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltBars()
    Dim tmpY As Long, tmpX As Long
    Dim sWidth As Long, sHeight As Long
    Dim sRECT As RECT
    Dim barWidth As Long
    Dim i As Long, NpcNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' dynamic bar calculations
    sWidth = DDSD_Bars.lWidth
    sHeight = DDSD_Bars.lHeight / 4

    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        NpcNum = MapNpc(i).num
        ' exists?
        If NpcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < GetPokemonMaxVital(MapNpc(i).num, MapNpc(i).Level) Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).XOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).YOffset + 32

                ' calculate the width to fill
                barWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (GetPokemonMaxVital(MapNpc(i).num, MapNpc(i).Level) / sWidth)) * sWidth

                ' draw bar background
                With sRECT
                    .top = sHeight * 1    ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

                ' draw the bar proper
                With sRECT
                    .top = 0    ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If PlayerSpells(SpellBuffer) = 0 Then
            SpellBufferTimer = 0
            PlayerSpells(SpellBuffer) = 0
            Exit Sub
        End If
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 30 + sHeight + 1

            ' calculate the width to fill
            barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth

            ' draw bar background
            With sRECT
                .top = sHeight * 3    ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

            ' draw the bar proper
            With sRECT
                .top = sHeight * 2    ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
    End If

    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).YOffset + 32

        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth

        ' draw bar background
        With sRECT
            .top = sHeight * 1    ' HP bar background
            .Left = 0
            .Right = .Left + sWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

        ' draw the bar proper
        With sRECT
            .top = 0    ' HP bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If

    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).XOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).YOffset + 35

                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth

                    ' draw bar background
                    With sRECT
                        .top = sHeight * 1    ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

                    ' draw the bar proper
                    With sRECT
                        .top = 0    ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            End If
        Next
    End If

    ' Check Fishing Bar
    If Player(MyIndex).InFishing > 0 Then

        ' lock to player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 25 + sHeight + 1

        ' calculate the width to fill
        barWidth = (GetTickCount - Player(MyIndex).InFishing) / (10000) * sWidth

        ' draw bar background
        With sRECT
            .top = sHeight * 3    ' cooldown bar background
            .Left = 0
            .Right = sWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

        ' draw the bar proper
        With sRECT
            .top = sHeight * 2    ' cooldown bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If

    ' Check Scan Bar
    If Player(MyIndex).ScanTime > 0 Then

        ' lock to player
        If GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon) = 0 Then
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 25 + sHeight + 1
        Else
            tmpX = Player(MyIndex).TPX * PIC_X + 16 - (sWidth / 2)
            tmpY = Player(MyIndex).TPY * PIC_Y + 25 + sHeight + 1
        End If

        ' calculate the width to fill
        barWidth = (GetTickCount - Player(MyIndex).ScanTime) / (2000) * sWidth

        ' draw bar background
        With sRECT
            .top = sHeight * 3    ' cooldown bar background
            .Left = 0
            .Right = sWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

        ' draw the bar proper
        With sRECT
            .top = sHeight * 2    ' cooldown bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If


    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBars", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub blthotbar()
    Dim sRECT As RECT, dRECT As RECT, i As Long, num As String, n As Long
    Dim PokeX As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.picHotbar.Cls
    For i = 1 To MAX_HOTBAR

        With dRECT
            .top = HotbarTop + 2
            .Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR))) + 2
            .Bottom = .top + 32
            .Right = .Left + 32
        End With

        With sRECT
            .top = 0
            .Left = 32
            .Bottom = 32
            .Right = 64
        End With

        Select Case Hotbar(i).sType
        Case 1    ' inventory
            If Len(Item(Hotbar(i).Slot).Name) > 0 Then

                If Item(Hotbar(i).Slot).Pic > 0 Then

                    If Hotbar(i).Pokemon = 0 Then
                        If DDS_Item(Item(Hotbar(i).Slot).Pic) Is Nothing Then
                            Call InitDDSurf("Items\" & Item(Hotbar(i).Slot).Pic, DDSD_Item(Item(Hotbar(i).Slot).Pic), DDS_Item(Item(Hotbar(i).Slot).Pic))
                        End If
                        Engine_BltToDC DDS_Item(Item(Hotbar(i).Slot).Pic), sRECT, dRECT, frmMain.picHotbar, False
                    Else

                        If Hotbar(i).Pokemon < UNLOCKED_POKEMONS Then
                            If DDS_PokeIcons(Hotbar(i).Pokemon) Is Nothing Then
                                Call InitDDSurf("PokeIcon\" & Hotbar(i).Pokemon, DDSD_PokeIcons(Hotbar(i).Pokemon), DDS_PokeIcons(Hotbar(i).Pokemon))
                            End If

                            Engine_BltToDC DDS_PokeIcons(Hotbar(i).Pokemon), sRECT, dRECT, frmMain.picHotbar, False
                        Else

                            If DDS_PokeIcons(1) Is Nothing Then
                                Call InitDDSurf("PokeIcon\" & 1, DDSD_PokeIcons(1), DDS_PokeIcons(1))
                            End If

                            Engine_BltToDC DDS_PokeIcons(1), sRECT, dRECT, frmMain.picHotbar, False
                        End If
                    End If
                End If
            End If

        Case 2    ' spell

            With sRECT
                .top = 0
                .Left = 0
                .Bottom = 32
                .Right = 32
            End With

            If Len(Spell(Hotbar(i).Slot).Name) > 0 Then
                If Spell(Hotbar(i).Slot).Icon > 0 Then

                    If DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon) Is Nothing Then
                        Call InitDDSurf("Spellicons\" & Spell(Hotbar(i).Slot).Icon, DDSD_SpellIcon(Spell(Hotbar(i).Slot).Icon), DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon))
                    End If

                    ' check for cooldown
                    For n = 1 To MAX_PLAYER_SPELLS
                        If Hotbar(i).Slot = PlayerSpells(n) Then
                            ' has spell
                            If Not SpellCD(n) = 0 Then
                                sRECT.Left = 32
                                sRECT.Right = 64
                            End If
                        End If
                    Next

                    Engine_BltToDC DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon), sRECT, dRECT, frmMain.picHotbar, False
                End If
            End If
        Case 3    ' pokemon
            ' aki e para renderizar o pokemon em vez de renderizar a pokebola
        End Select

        'If i = 4 Then num = "V"
        DrawText frmMain.picHotbar.hDC, dRECT.Left + 2, dRECT.top + 16, i, QBColor(White)
    Next
    frmMain.picHotbar.Refresh

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHotbar", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPlayer(ByVal Index As Long)
    Dim Anim As Byte, i As Long, X As Long, Y As Long
    Dim Sprite As Long, SpriteTop As Long
    Dim rec As DxVBLib.RECT
    Dim attackspeed As Long
    Dim X2 As Long, Y2 As Long
    Dim Cabelo As Long, AnimStep(1 To 4) As Byte
    Dim Coluna As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    ' speed from weapon
    If GetPlayerEquipment(Index, weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, weapon)).speed
    Else
        attackspeed = 1000
    End If

    ' Configuração Ler Coluna 1 ou 2
    If Player(Index).InSurf = 1 Or Player(Index).InFishing Then Coluna = 1
    If GetPlayerEquipment(Index, weapon) Then
        If Item(GetPlayerEquipment(Index, weapon)).vel > 0 Then Coluna = 1
    End If

    ' Configuração Correr
    If GetPlayerEquipment(Index, weapon) = 0 Then
        If Player(Index).Running = True Then
            AnimStep(1) = 6
            AnimStep(2) = 6
            AnimStep(3) = 5
            AnimStep(4) = 7
        Else
            AnimStep(1) = 0
            AnimStep(2) = 2
            AnimStep(3) = 1
            AnimStep(4) = 3
        End If
    Else
        AnimStep(1) = 0
        AnimStep(2) = 2
        AnimStep(3) = 1
        AnimStep(4) = 3
    End If

    ' Reset frame
    If Player(Index).Step = 3 Then
        Anim = AnimStep(1)
    ElseIf Player(Index).Step = 1 Then
        Anim = AnimStep(2)
    End If

    'Bike Stand
    If Player(Index).PuloStatus = 0 Then
        If Player(Index).Parado = 0 Then
            If Player(Index).InSurf = 0 Then
                Anim = 4
            ElseIf Player(Index).InSurf = 1 Then
                Anim = 5
            End If
        End If
    End If

    'Fishing anim
    If Player(Index).InFishing > 0 Then
        If Player(Index).InFishing + 500 > GetTickCount Then
            Anim = 5
        Else
            Anim = 6
        End If
    End If

    'Surf
    If Player(Index).InSurf = 0 Then

        Select Case GetPlayerDir(Index)
        Case DIR_UP
            If (Player(Index).YOffset > 8) Then
                If Player(Index).Step = 1 Then Anim = AnimStep(3)
                If Player(Index).Step = 3 Then Anim = AnimStep(4)
                Player(Index).Parado = 200 + GetTickCount
            End If
        Case DIR_DOWN
            If (Player(Index).YOffset < -8) Then
                If Player(Index).Step = 1 Then Anim = AnimStep(3)
                If Player(Index).Step = 3 Then Anim = AnimStep(4)
                Player(Index).Parado = 200 + GetTickCount
            End If
        Case DIR_LEFT
            If (Player(Index).XOffset > 8) Then
                If Player(Index).Step = 1 Then Anim = AnimStep(3)
                If Player(Index).Step = 3 Then Anim = AnimStep(4)
                Player(Index).Parado = 200 + GetTickCount
            End If
        Case DIR_RIGHT
            If (Player(Index).XOffset < -8) Then
                If Player(Index).Step = 1 Then Anim = AnimStep(3)
                If Player(Index).Step = 3 Then Anim = AnimStep(4)
                Player(Index).Parado = 200 + GetTickCount
            End If
        End Select

    ElseIf Player(Index).InSurf = 1 Then
        Anim = 7
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    'Ler Coluna
    If Coluna = 0 Then
        Select Case GetPlayerDir(Index)
        Case DIR_UP: SpriteTop = 3
        Case DIR_RIGHT: SpriteTop = 2
        Case DIR_DOWN: SpriteTop = 0
        Case DIR_LEFT: SpriteTop = 1
        End Select
    Else
        Select Case GetPlayerDir(Index)
        Case DIR_UP: SpriteTop = 7
        Case DIR_RIGHT: SpriteTop = 6
        Case DIR_DOWN: SpriteTop = 4
        Case DIR_LEFT: SpriteTop = 5
        End Select
    End If

    'Animação de Pulo
    If Player(Index).PuloStatus > 0 Then
        Anim = 1
    End If

    'Recorte
    With rec
        .top = SpriteTop * (DDSD_Character(Sprite).lHeight / 8)
        .Bottom = .top + (DDSD_Character(Sprite).lHeight / 8)
        .Left = Anim * (DDSD_Character(Sprite).lWidth / 8)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 8)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((DDSD_Character(Sprite).lWidth / 8 - 32) / 2)

    ' Create a 32 pixel offset for larger sprites
    If (DDSD_Character(Sprite).lHeight) > 32 Then
        If Player(Index).Flying = 1 Then
            Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((DDSD_Character(Sprite).lHeight / 8))
        Else
            Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((DDSD_Character(Sprite).lHeight / 8) - 32) - Player(Index).PuloSlide
        End If
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
    End If

    If Player(Index).TPX > 0 Then
        Call BltTrainerPointMap(Index)
    End If

    CharacterTimer(15) = GetTickCount + SurfaceTimerMax
    If DDS_Character(15) Is Nothing Then
        Call InitDDSurf("characters\" & 15, DDSD_Character(15), DDS_Character(15))
    End If

    If Player(Index).InSurf = 1 Then
        If Not GetPlayerDir(Index) = DIR_DOWN Then
            Call BltSprite(15, X, Y, rec)
        End If
    End If

    ' Renderizar a Base
    Call BltSprite(Sprite, X, Y, rec)

    ' Renderizar Cabelo
    If Player(Index).HairNum > 0 Then
        Call BltCabelo(X, Y, Player(Index).HairNum, Anim, SpriteTop)
    End If

    ' check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, SpriteTop)
            End If
        End If
    Next

    If Player(Index).InSurf = 1 Then
        If Not GetPlayerDir(Index) <> DIR_DOWN Then
            Call BltSprite(15, X, Y, rec)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayer", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPlayerPokemon(ByVal Index As Long)
    Dim Anim As Byte, i As Long, X As Long, Y As Long
    Dim Sprite As Long, SpriteTop As Long
    Dim rec As DxVBLib.RECT
    Dim attackspeed As Long
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    ' speed from weapon
    If GetPlayerEquipment(Index, weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, weapon)).speed
    Else
        attackspeed = 1000
    End If

    ' Reset frame
    If Player(Index).Step = 3 Then
        Anim = 0
    ElseIf Player(Index).Step = 1 Then
        Anim = 2
    End If

    'Stand Pokémon
    If Player(Index).PuloStatus = 0 Then
        If Player(Index).Parado = 0 Then
            Anim = 4
        End If
    End If

    'Anim Pokémon
    If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
        If Player(Index).AttackTimer = 0 Then
            If Player(Index).Flying = 0 Then
                If Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).AnimFrame(1) > 0 Then
                    Anim = Player(Index).AnimFrame
                End If
            Else
                If Pokemon(GetPlayerEquipmentPokeInfoPokemon(Index, weapon)).AnimFrame(2) > 0 Then
                    Anim = Player(Index).AnimFrame
                End If
            End If
        End If
    End If

    ' Attack Pokémon
    If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If Player(Index).Attacking = 1 Then
            If GetPlayerEquipmentPokeInfoPokemon(Index, weapon) > 0 Then
                If Player(Index).AttackTimer + 100 > GetTickCount Then    '"Aqui está o segredo" quando o AttackTimer + 100 > gettickcount ira usar a animação 4
                    Anim = 5
                ElseIf Player(Index).AttackTimer + 300 > GetTickCount Then    'Quando o AttackTimer + 300 > gettickcount ira usar a animação 5
                    Anim = 6
                End If
            End If
        End If

    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(Index)
        Case DIR_UP
            If (Player(Index).YOffset > 8) Then
                Anim = Player(Index).Step
                Player(Index).Parado = 200 + GetTickCount
            End If
        Case DIR_DOWN
            If (Player(Index).YOffset < -8) Then
                Anim = Player(Index).Step
                Player(Index).Parado = 200 + GetTickCount
            End If
        Case DIR_LEFT
            If (Player(Index).XOffset > 8) Then
                Anim = Player(Index).Step
                Player(Index).Parado = 200 + GetTickCount
            End If
        Case DIR_RIGHT
            If (Player(Index).XOffset < -8) Then
                Anim = Player(Index).Step
                Player(Index).Parado = 200 + GetTickCount
            End If
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    If Player(Index).Flying = 0 Then
        Select Case GetPlayerDir(Index)
        Case DIR_UP
            SpriteTop = 3
        Case DIR_RIGHT
            SpriteTop = 2
        Case DIR_DOWN
            SpriteTop = 0
        Case DIR_LEFT
            SpriteTop = 1
        End Select
    Else
        Select Case GetPlayerDir(Index)
        Case DIR_UP
            SpriteTop = 7
        Case DIR_RIGHT
            SpriteTop = 6
        Case DIR_DOWN
            SpriteTop = 4
        Case DIR_LEFT
            SpriteTop = 5
        End Select
    End If

    If Player(Index).PuloStatus > 0 Then
        Anim = 1
    End If

    With rec
        .top = SpriteTop * (DDSD_Character(Sprite).lHeight / 8)
        .Bottom = .top + (DDSD_Character(Sprite).lHeight / 8)
        .Left = Anim * (DDSD_Character(Sprite).lWidth / 8)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 8)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((DDSD_Character(Sprite).lWidth / 8 - 32) / 2)

    ' Create a 32 pixel offset for larger sprites
    If (DDSD_Character(Sprite).lHeight) > 32 Then
        If Player(Index).Flying = 1 Then
            Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((DDSD_Character(Sprite).lHeight / 8))
        Else
            Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((DDSD_Character(Sprite).lHeight / 8) - 32) - Player(Index).PuloSlide
        End If
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
    End If

    If Player(Index).TPX > 0 Then
        Call BltTrainerPointMap(Index)
    End If

    Select Case Sprite
    Case 144 To 151
        Y = Y + 22
    Case 154, 155
        Y = Y + 25
    End Select

    ' render the actual sprite
    Call BltSprite(Sprite, X, Y, rec)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayerPokemon", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Sub BltNpc(ByVal MapNpcNum As Long)
    Dim Anim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, SpriteTop As Long
    Dim rec As DxVBLib.RECT
    Dim attackspeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).num = 0 Then Exit Sub    ' no npc set

    Sprite = Npc(MapNpc(MapNpcNum).num).Sprite

    If MapNpc(MapNpcNum).Shiny = True Then
        Sprite = Sprite + 1
    End If

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    attackspeed = 1000

    ' Reset frame
    Anim = 0

    ' Npc Desmaiado
    If MapNpc(MapNpcNum).Desmaiado = True Then
        Anim = 7
    End If

    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If MapNpc(MapNpcNum).Attacking = 1 Then

            If MapNpc(MapNpcNum).AttackTimer + 100 > GetTickCount Then    '"Aqui está o segredo" quando o AttackTimer + 100 > gettickcount ira usar a animação 4
                Anim = 5
            ElseIf MapNpc(MapNpcNum).AttackTimer + 300 > GetTickCount Then    'Quando o AttackTimer + 300 > gettickcount ira usar a animação 5
                Anim = 6
            End If
        End If
    End If


    ' If not attacking, walk normally
    Select Case MapNpc(MapNpcNum).Dir
    Case DIR_UP
        If (MapNpc(MapNpcNum).YOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
    Case DIR_DOWN
        If (MapNpc(MapNpcNum).YOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
    Case DIR_LEFT
        If (MapNpc(MapNpcNum).XOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
    Case DIR_RIGHT
        If (MapNpc(MapNpcNum).XOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
    End Select


    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
    Case DIR_UP
        SpriteTop = 3
    Case DIR_RIGHT
        SpriteTop = 2
    Case DIR_DOWN
        SpriteTop = 0
    Case DIR_LEFT
        SpriteTop = 1
    End Select

    With rec
        .top = (DDSD_Character(Sprite).lHeight / 8) * SpriteTop
        .Bottom = .top + DDSD_Character(Sprite).lHeight / 8
        .Left = Anim * (DDSD_Character(Sprite).lWidth / 8)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 8)
    End With

    ' Calculate the X
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset - ((DDSD_Character(Sprite).lWidth / 8 - 32) / 2)

    ' Is the player's height more than 32..?
    If (DDSD_Character(Sprite).lHeight / 8) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset - ((DDSD_Character(Sprite).lHeight / 8) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
    End If

    Select Case Sprite
    Case 144 To 151
        Y = Y + 22
    Case 154, 155
        Y = Y + 25
    End Select

    Call BltSprite(Sprite, X, Y, rec)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltNpc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal Sprite As Long, ByVal Anim As Long, ByVal SpriteTop As Long)
    Dim rec As DxVBLib.RECT
    Dim X As Long, Y As Long
    Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub

    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("Paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If

    With rec
        .top = SpriteTop * (DDSD_Paperdoll(Sprite).lHeight / 8)
        .Bottom = .top + (DDSD_Paperdoll(Sprite).lHeight / 8)
        .Left = Anim * (DDSD_Paperdoll(Sprite).lWidth / 8)
        .Right = .Left + (DDSD_Paperdoll(Sprite).lWidth / 8)
    End With

    ' clipping
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping

    Call Engine_BltFast(X, Y, DDS_Paperdoll(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltCabelo(ByVal X2 As Long, ByVal Y2 As Long, ByVal Cabelo As Long, ByVal Anim As Long, ByVal SpriteTop As Long)
    Dim rec As DxVBLib.RECT
    Dim X As Long, Y As Long
    Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Cabelo < 1 Or Cabelo > HairNum Then Exit Sub

    HairTimer(Cabelo) = GetTickCount + SurfaceTimerMax
    If DDS_Hair(Cabelo) Is Nothing Then
        Call InitDDSurf("characters\Cabelos\" & Cabelo, DDSD_Hair(Cabelo), DDS_Hair(Cabelo))
    End If

    With rec
        .top = SpriteTop * (DDSD_Hair(Cabelo).lHeight / 8)
        .Bottom = .top + (DDSD_Hair(Cabelo).lHeight / 8)
        .Left = Anim * (DDSD_Hair(Cabelo).lWidth / 8)
        .Right = .Left + (DDSD_Hair(Cabelo).lWidth / 8)
    End With

    ' clipping
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If

    ' /clipping
    Call Engine_BltFast(X, Y, DDS_Hair(Cabelo), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltCabelo", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltSprite(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, rec As DxVBLib.RECT)
    Dim X As Long
    Dim Y As Long
    Dim Width As Long
    Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.top)

    ' clipping
    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Height - DDSD_BackBuffer.lHeight)
    End If

    If X + Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping

    Call Engine_BltFast(X, Y, DDS_Character(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltAnimatedInvItems()
    Dim i As Long
    Dim ItemNum As Long, itempic As Long
    Dim X As Long, Y As Long
    Dim MaxFrames As Byte
    Dim Amount As Long
    Dim rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub

    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS

        If MapItem(i).num > 0 Then
            itempic = Item(MapItem(i).num).Pic

            If itempic < 1 Or itempic > numitems Then Exit Sub
            MaxFrames = (DDSD_Item(itempic).lWidth / 2) / 32    ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < MaxFrames - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 0
            End If
        End If

    Next

    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, i)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic

            If itempic > 0 And itempic <= numitems Then
                If DDSD_Item(itempic).lWidth > 64 Then
                    MaxFrames = (DDSD_Item(itempic).lWidth / 2) / 32    ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(i) < MaxFrames - 1 Then
                        InvItemFrame(i) = InvItemFrame(i) + 1
                    Else
                        InvItemFrame(i) = 0
                    End If

                    With rec
                        .top = 0
                        .Bottom = 32
                        .Left = (DDSD_Item(itempic).lWidth / 2) + (InvItemFrame(i) * 32)    ' middle to get the start of inv gfx, then +32 for each frame
                        .Right = .Left + 32
                    End With

                    With rec_pos
                        .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    ' We'll now re-blt the item, and place the currency value over it again :P
                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.top + 22
                        X = rec_pos.Left - 4
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        DrawText frmMain.picInventory.hDC, X, Y, ConvertCurrency(Amount), QBColor(Yellow)
                    End If
                End If
            End If
        End If

    Next

    frmMain.picInventory.Refresh

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimatedInvItems", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltFacePokemon()
    Dim rec As RECT, rec_pos As RECT, faceNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub

    frmMain.PicFacePokemon.Cls

    faceNum = LastItemPoke    'GetPlayerSprite(MyIndex)

    If faceNum <= 0 Or faceNum > NumFaces Then Exit Sub

    With rec
        .top = 0
        .Bottom = 94
        .Left = 0
        .Right = 94
    End With

    With rec_pos
        .top = 0
        .Bottom = 94
        .Left = 0
        .Right = 94
    End With

    ' Load face if not loaded, and reset timer
    FaceTimer(faceNum) = GetTickCount + SurfaceTimerMax

    If DDS_Face(faceNum) Is Nothing Then
        Call InitDDSurf("Faces\" & faceNum, DDSD_Face(faceNum), DDS_Face(faceNum))
    End If

    Engine_BltToDC DDS_Face(faceNum), rec, rec_pos, frmMain.PicFacePokemon, False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltFacePokemon", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltFace()
    Dim faceNum As Long
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumCharacters = 0 Then Exit Sub

    frmMain.picFace.Cls

    faceNum = GetPlayerSprite(MyIndex)

    If faceNum <= 0 Or faceNum > NumCharacters Then Exit Sub

    ' Load face if not loaded, and reset timer
    FaceTimer(faceNum) = GetTickCount + SurfaceTimerMax

    If DDS_Character(faceNum) Is Nothing Then
        Call InitDDSurf("characters\" & faceNum, DDSD_Character(faceNum), DDS_Character(faceNum))
    End If


    Width = DDSD_Character(faceNum).lWidth / 10
    Height = DDSD_Character(faceNum).lHeight / 10

    sRECT.top = 25
    sRECT.Bottom = sRECT.top + Height
    sRECT.Left = 13
    sRECT.Right = sRECT.Left + Width

    dRECT.top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width

    Engine_BltToDC DDS_Character(faceNum), sRECT, dRECT, frmMain.picFace, False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltFace", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltEquipment()
    Dim i As Long, ItemNum As Long, itempic As Long
    Dim rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If numitems = 0 Then Exit Sub

    frmMain.picCharacter.Cls

    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(MyIndex, i)

        If ItemNum > 0 Then
            itempic = Item(ItemNum).Pic

            If ItemNum = 3 Then
                Select Case GetPlayerEquipmentPokeInfoPokeball(MyIndex, i)
                Case 1
                    itempic = 2    'Pokeball
                Case 2
                    itempic = 3    'Super Ball
                Case 3
                    itempic = 4    'Great Ball
                Case 4
                    itempic = 5    'Master Ball
                Case 5
                    itempic = 6    'Wat ball
                Case 6
                    itempic = 7    'Safari ball
                End Select
            End If

            With rec
                .top = 0
                .Bottom = 32
                .Left = 32
                .Right = 64
            End With

            With rec_pos
                .top = EqTop
                .Bottom = .top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            ' Load item if not loaded, and reset timer
            ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

            If DDS_Item(itempic) Is Nothing Then
                Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
            End If

            Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picCharacter, False
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltEquipment", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltInventory()
    Dim i As Long, X As Long, Y As Long, ItemNum As Long, itempic As Long
    Dim Amount As Long
    Dim rec As RECT, rec_pos As RECT
    Dim colour As Long
    Dim tmpItem As Long, amountModifier As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub

    frmMain.picInventory.Cls

    For i = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, i)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic

            If GetPlayerInvItemPokeInfoPokemon(MyIndex, i) > 0 Then
                itempic = GetPlayerInvItemPokeInfoPokemon(MyIndex, i)
                If itempic >= UNLOCKED_POKEMONS Then itempic = 1
            End If

            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                    If TradeYourOffer(X).num = i Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(X).value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(X).value
                            End If
                        End If
                    End If
                Next
            End If

            With rec
                .top = 0
                .Bottom = 32
                .Left = 32
                .Right = 64
            End With

            If GetPlayerInvItemPokeInfoPokemon(MyIndex, i) > 0 Then
                If GetPlayerInvItemPokeInfoVital(MyIndex, i, 1) = 0 Then

                    With rec
                        .top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With

                End If
            End If

            With rec_pos
                .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If GetPlayerInvItemPokeInfoPokemon(MyIndex, i) > 0 Then

                If GetPlayerInvItemShiny(MyIndex, i) > 0 Then
                    'Timer
                    PokeIconShinyTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_PokeIconShiny(itempic) Is Nothing Then
                        Call InitDDSurf("PokeIcon\Shiny\" & itempic, DDSD_PokeIconShiny(itempic), DDS_PokeIconShiny(itempic))
                    End If

                    Engine_BltToDC DDS_PokeIconShiny(itempic), rec, rec_pos, frmMain.picInventory, False

                Else
                    'Timer
                    PokeIconTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_PokeIcons(itempic) Is Nothing Then
                        Call InitDDSurf("PokeIcon\" & itempic, DDSD_PokeIcons(itempic), DDS_PokeIcons(itempic))
                    End If

                    Engine_BltToDC DDS_PokeIcons(itempic), rec, rec_pos, frmMain.picInventory, False
                End If

            Else

                If itempic = 0 Then Exit Sub
                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If

                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picInventory, False

            End If

            ' If item is a stack - draw the amount you have

            If GetPlayerInvItemValue(MyIndex, i) > 1 Or GetPlayerInvItemPokeInfoLevel(MyIndex, i) > 0 Then
                Y = rec_pos.top + 22
                X = rec_pos.Left - 4

                If GetPlayerInvItemPokeInfoLevel(MyIndex, i) > 0 Then
                    Amount = GetPlayerInvItemPokeInfoLevel(MyIndex, i)
                    DrawText frmMain.picInventory.hDC, X + 4, Y - 2, Amount, QBColor(BrightGreen)

                    If Not Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_ROD Then
                        If GetPlayerInvItemSexo(MyIndex, i) = 0 Then
                            DrawText frmMain.picInventory.hDC, X + 24, Y - 22, "M", QBColor(BrightCyan)
                        Else
                            DrawText frmMain.picInventory.hDC, X + 24, Y - 22, "F", QBColor(BrightRed)
                        End If
                    End If

                Else
                    Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                End If

                If GetPlayerInvItemPokeInfoLevel(MyIndex, i) = 0 Then
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = QBColor(White)
                    ElseIf Amount > 10000000 Then
                        colour = QBColor(White)
                    End If

                    DrawText frmMain.picInventory.hDC, X, Y, Format$(ConvertCurrency(str(Amount)), "#,###,###,###"), colour
                End If

            End If
        End If
NextLoop:
    Next

    frmMain.picInventory.Refresh
    'update animated items
    BltAnimatedInvItems

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventory", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltInventoryPokemon()
    Dim i As Long, X As Long, Y As Long, ItemNum As Long, itempic As Long
    Dim Amount As Long
    Dim rec As RECT, rec_pos As RECT
    Dim colour As Long
    Dim tmpItem As Long, amountModifier As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub

    frmMain.picInventory.Cls

    For i = 1 To MAX_INV

        ItemNum = GetPlayerInvItemPokeInfoPokemon(MyIndex, i)

        If ItemNum > 0 And ItemNum <= MAX_POKEMONS Then
            itempic = GetPlayerInvItemPokeInfoPokemon(MyIndex, i)

            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                    If TradeYourOffer(X).num = i Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(X).value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(X).value
                            End If
                        End If
                    End If
                Next
            End If

            If itempic > 0 And itempic <= NumPokeIcons Then
                If DDSD_PokeIcons(itempic).lWidth <= 64 Then    ' more than 1 frame is handled by anim sub

                    With rec
                        .top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With

                    With rec_pos
                        .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    PokeIconTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_PokeIcons(itempic) Is Nothing Then
                        Call InitDDSurf("PokeIcon\" & itempic, DDSD_PokeIcons(itempic), DDS_PokeIcons(itempic))
                    End If

                    Engine_BltToDC DDS_PokeIcons(itempic), rec, rec_pos, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.top + 22
                        X = rec_pos.Left - 4

                        Amount = GetPlayerInvItemPokeInfoLevel(MyIndex, i)

                        DrawText frmMain.picInventory.hDC, X, Y, Amount, BrightGreen
                    End If
                End If
            End If
        End If
NextLoop:
    Next

    frmMain.picInventory.Refresh
    'update animated items
    BltAnimatedInvItems

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventory", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltTrade()
    Dim i As Long, X As Long, Y As Long, ItemNum As Long, itempic As Long
    Dim Amount As Long
    Dim rec As RECT, rec_pos As RECT
    Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picYourTrade.Cls
    frmMain.picTheirTrade.Cls

    For i = 1 To MAX_INV
        ' blt your own offer
        ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic

            If GetPlayerInvItemPokeInfoPokemon(MyIndex, TradeYourOffer(i).num) > 0 Then
                itempic = GetPlayerInvItemPokeInfoPokemon(MyIndex, TradeYourOffer(i).num)
            End If

            If itempic > 0 And itempic <= numitems Then

                If GetPlayerInvItemPokeInfoPokemon(MyIndex, TradeYourOffer(i).num) > 0 Then
                    If GetPlayerInvItemPokeInfoVital(MyIndex, TradeYourOffer(i).num, 1) > 0 Then
                        With rec
                            .top = 0
                            .Bottom = 32
                            .Left = 32
                            .Right = 64
                        End With
                    Else
                        With rec
                            .top = 0
                            .Bottom = 32
                            .Left = 0
                            .Right = 32
                        End With
                    End If

                Else
                    With rec
                        .top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With
                End If

                With rec_pos
                    .top = TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                    .Bottom = .top + PIC_Y
                    .Left = TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))
                    .Right = .Left + PIC_X
                End With

                If GetPlayerInvItemPokeInfoPokemon(MyIndex, TradeYourOffer(i).num) > 0 Then

                    ' Load item if not loaded, and reset timer
                    PokeIconTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_PokeIcons(itempic) Is Nothing Then
                        Call InitDDSurf("PokeIcon\" & itempic, DDSD_PokeIcons(itempic), DDS_PokeIcons(itempic))
                    End If

                    Engine_BltToDC DDS_PokeIcons(itempic), rec, rec_pos, frmMain.picYourTrade, False

                Else

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picYourTrade, False

                End If

                'Set Y,X
                Y = rec_pos.top + 22
                X = rec_pos.Left - 4

                If GetPlayerInvItemPokeInfoLevel(MyIndex, TradeYourOffer(i).num) > 0 Then
                    Amount = GetPlayerInvItemPokeInfoLevel(MyIndex, TradeYourOffer(i).num)
                    DrawText frmMain.picYourTrade.hDC, X + 4, Y - 2, Amount, QBColor(BrightGreen)

                    If Not Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Type = ITEM_TYPE_ROD Then
                        If GetPlayerInvItemSexo(MyIndex, TradeYourOffer(i).num) = 0 Then
                            DrawText frmMain.picYourTrade.hDC, X + 24, Y - 22, "M", QBColor(BrightCyan)
                        Else
                            DrawText frmMain.picYourTrade.hDC, X + 24, Y - 22, "F", QBColor(BrightRed)
                        End If
                    End If
                End If

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).value > 1 Then

                    Amount = TradeYourOffer(i).value

                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If

                    DrawText frmMain.picYourTrade.hDC, X, Y, ConvertCurrency(str(Amount)), colour
                End If
            End If
        End If

        ' blt their offer
        ItemNum = TradeTheirOffer(i).num

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic

            If TradeTheirOffer(i).PokeInfo.Pokemon > 0 Then
                itempic = TradeTheirOffer(i).PokeInfo.Pokemon
            End If

            If itempic > 0 And itempic <= numitems Then

                If TradeTheirOffer(i).PokeInfo.Pokemon > 0 Then

                    If TradeTheirOffer(i).PokeInfo.Vital(1) > 0 Then
                        With rec
                            .top = 0
                            .Bottom = 32
                            .Left = 32
                            .Right = 64
                        End With
                    Else
                        With rec
                            .top = 0
                            .Bottom = 32
                            .Left = 0
                            .Right = 32
                        End With
                    End If

                Else
                    With rec
                        .top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With
                End If

                With rec_pos
                    .top = TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
                    .Bottom = .top + PIC_Y
                    .Left = TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))
                    .Right = .Left + PIC_X
                End With

                If TradeTheirOffer(i).PokeInfo.Pokemon > 0 Then

                    ' Load item if not loaded, and reset timer
                    PokeIconTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_PokeIcons(itempic) Is Nothing Then
                        Call InitDDSurf("PokeIcon\" & itempic, DDSD_PokeIcons(itempic), DDS_PokeIcons(itempic))
                    End If

                    Engine_BltToDC DDS_PokeIcons(itempic), rec, rec_pos, frmMain.picTheirTrade, False

                Else
                    ' Load item if not loaded, and reset timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picTheirTrade, False
                End If
                'editar aqui amanhã
                'Set Y,X
                Y = rec_pos.top + 22
                X = rec_pos.Left - 4

                If TradeTheirOffer(i).PokeInfo.Level > 0 Then
                    Amount = TradeTheirOffer(i).PokeInfo.Level
                    DrawText frmMain.picTheirTrade.hDC, X + 4, Y - 2, Amount, QBColor(BrightGreen)

                    If Not Item(TradeTheirOffer(i).num).Type = ITEM_TYPE_ROD Then
                        If TradeTheirOffer(i).PokeInfo.Sexo = 0 Then
                            DrawText frmMain.picTheirTrade.hDC, X + 24, Y - 22, "M", QBColor(BrightCyan)
                        Else
                            DrawText frmMain.picTheirTrade.hDC, X + 24, Y - 22, "F", QBColor(BrightRed)
                        End If
                    End If
                End If

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).value > 1 Then

                    Amount = TradeTheirOffer(i).value

                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If

                    DrawText frmMain.picTheirTrade.hDC, X, Y, ConvertCurrency(str(Amount)), colour
                End If
            End If
        End If
    Next

    frmMain.picYourTrade.Refresh
    frmMain.picTheirTrade.Refresh

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTrade", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltPlayerSpells()
    Dim i As Long, X As Long, Y As Long, SpellNum As Long, spellicon As Long
    Dim Amount As String
    Dim rec As RECT, rec_pos As RECT
    Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picSpells.Cls

    For i = 1 To MAX_PLAYER_SPELLS
        SpellNum = PlayerSpells(i)

        If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
            spellicon = Spell(SpellNum).Icon

            If spellicon > 0 And spellicon <= NumSpellIcons Then

                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 0
                    .Right = 32
                End With

                If Not SpellCD(i) = 0 Then
                    rec.Left = 32
                    rec.Right = 64
                End If

                With rec_pos
                    .top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                    .Bottom = .top + PIC_Y
                    .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load spellicon if not loaded, and reset timer
                SpellIconTimer(spellicon) = GetTickCount + SurfaceTimerMax

                If DDS_SpellIcon(spellicon) Is Nothing Then
                    Call InitDDSurf("SpellIcons\" & spellicon, DDSD_SpellIcon(spellicon), DDS_SpellIcon(spellicon))
                End If

                Engine_BltToDC DDS_SpellIcon(spellicon), rec, rec_pos, frmMain.picSpells, False
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayerSpells", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltShop()
    Dim i As Long, X As Long, Y As Long, ItemNum As Long, itempic As Long
    Dim Amount As String
    Dim rec As RECT, rec_pos As RECT
    Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub

    frmMain.picShopItems.Cls

    For i = 1 To MAX_TRADES
        ItemNum = Shop(InShop).TradeItem(i).Item    'GetPlayerInvItemNum(MyIndex, i)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            itempic = Item(ItemNum).Pic
            If itempic > 0 And itempic <= numitems Then

                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With rec_pos
                    .top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                    .Bottom = .top + PIC_Y
                    .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If

                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picShopItems, False

                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    Y = rec_pos.top + 22
                    X = rec_pos.Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)

                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = QBColor(White)
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If

                    DrawText frmMain.picShopItems.hDC, X, Y, ConvertCurrency(Amount), colour
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltShop", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltInventoryItem(ByVal X As Long, ByVal Y As Long)
    Dim rec As RECT, rec_pos As RECT
    Dim ItemNum As Long, itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)

    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        itempic = Item(ItemNum).Pic

        If ItemNum = 3 Then
            Select Case GetPlayerInvItemPokeInfoPokeball(MyIndex, DragInvSlotNum)
            Case 1
                itempic = 2
            Case 2
                itempic = 3
            Case 3
                itempic = 4
            Case 4
                itempic = 5
            Case 5
                itempic = 6
            Case 6
                itempic = 7
            End Select
        End If

        If itempic = 0 Then Exit Sub

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = DDSD_Item(itempic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .top = 2
            .Bottom = .top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

        If DDS_Item(itempic) Is Nothing Then
            Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
        End If

        Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picTempInv, False

        With frmMain.picTempInv
            .top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltDraggedSpell(ByVal X As Long, ByVal Y As Long)
    Dim rec As RECT, rec_pos As RECT
    Dim SpellNum As Long, spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellNum = PlayerSpells(DragSpell)

    If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
        spellpic = Spell(SpellNum).Icon

        If spellpic = 0 Then Exit Sub

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .top = 2
            .Bottom = .top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = GetTickCount + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("Spellicons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        Engine_BltToDC DDS_SpellIcon(spellpic), rec, rec_pos, frmMain.picTempSpell, False

        With frmMain.picTempSpell
            .top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItemDesc(ByVal ItemNum As Long)
    Dim rec As RECT, rec_pos As RECT
    Dim itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.picItemDescPic.Cls

    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        itempic = Item(ItemNum).Pic

        If itempic = 0 Then Exit Sub

        ' Load item if not loaded, and reset timer
        ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

        If DDS_Item(itempic) Is Nothing Then
            Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
        End If

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = DDSD_Item(itempic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .top = 0
            .Bottom = 32
            .Left = 0
            .Right = 32
        End With
        Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picItemDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItemDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltSpellDesc(ByVal SpellNum As Long)
    Dim rec As RECT, rec_pos As RECT
    Dim spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.picSpellDescPic.Cls

    If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
        spellpic = Spell(SpellNum).Icon

        If spellpic <= 0 Or spellpic > NumSpellIcons Then Exit Sub

        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = GetTickCount + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("SpellIcons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        Engine_BltToDC DDS_SpellIcon(spellpic), rec, rec_pos, frmMain.picSpellDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSpellDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_BltTileset()
    Dim Height As Long
    Dim Width As Long
    Dim Tileset As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.value

    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub

    ' make sure it's loaded
    If DDS_Tileset(Tileset) Is Nothing Then
        Call InitDDSurf("tilesets\" & Tileset, DDSD_Tileset(Tileset), DDS_Tileset(Tileset))
    End If

    Height = DDSD_Tileset(Tileset).lHeight
    Width = DDSD_Tileset(Tileset).lWidth

    dRECT.top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width

    frmEditor_Map.picBackSelect.Height = Height
    frmEditor_Map.picBackSelect.Width = Width

    Call Engine_BltToDC(DDS_Tileset(Tileset), sRECT, dRECT, frmEditor_Map.picBackSelect)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltTileset", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTileOutline()
    Dim rec As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optBlock.value Then Exit Sub

    With rec
        .top = 0
        .Bottom = .top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    Call Engine_BltFast(ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), DDS_Misc, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTileOutline", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NewCharacterBltSprite()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub

    If frmMenu.optMale.value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("Characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    Width = DDSD_Character(Sprite).lWidth / 10
    Height = DDSD_Character(Sprite).lHeight / 10

    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height

    sRECT.top = 25
    sRECT.Bottom = sRECT.top + Height
    sRECT.Left = 13
    sRECT.Right = sRECT.Left + Width

    dRECT.top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width

    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmMenu.picSprite)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterBltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltMapItem()
    Dim ItemNum As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = Item(frmEditor_Map.scrlMapItem.value).Pic

    If ItemNum < 1 Or ItemNum > numitems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    ItemTimer(ItemNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(ItemNum) Is Nothing Then
        Call InitDDSurf("Items\" & ItemNum, DDSD_Item(ItemNum), DDS_Item(ItemNum))
    End If

    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(ItemNum), sRECT, dRECT, frmEditor_Map.picMapItem)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltMapItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltKey()
    Dim ItemNum As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = Item(frmEditor_Map.scrlMapKey.value).Pic

    If ItemNum < 1 Or ItemNum > numitems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If

    ItemTimer(ItemNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(ItemNum) Is Nothing Then
        Call InitDDSurf("Items\" & ItemNum, DDSD_Item(ItemNum), DDS_Item(ItemNum))
    End If

    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(ItemNum), sRECT, dRECT, frmEditor_Map.picMapKey)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltKey", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltItem()
    Dim ItemNum As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = frmEditor_Item.scrlPic.value

    If ItemNum < 1 Or ItemNum > numitems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ItemTimer(ItemNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(ItemNum) Is Nothing Then
        Call InitDDSurf("Items\" & ItemNum, DDSD_Item(ItemNum), DDS_Item(ItemNum))
    End If

    ' rect for source
    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X

    ' same for destination as source
    dRECT = sRECT
    Call Engine_BltToDC(DDS_Item(ItemNum), sRECT, dRECT, frmEditor_Item.picItem)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltPaperdoll()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Item.picPaperdoll.Cls

    Sprite = frmEditor_Item.scrlPaperdoll.value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    PaperdollTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If

    ' rect for source
    sRECT.top = 0
    sRECT.Bottom = DDSD_Paperdoll(Sprite).lHeight
    sRECT.Left = 0
    sRECT.Right = DDSD_Paperdoll(Sprite).lWidth
    ' same for destination as source
    dRECT = sRECT

    Call Engine_BltToDC(DDS_Paperdoll(Sprite), sRECT, dRECT, frmEditor_Item.picPaperdoll)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_BltIcon()
    Dim iconnum As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    iconnum = frmEditor_Spell.scrlIcon.value

    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If

    SpellIconTimer(iconnum) = GetTickCount + SurfaceTimerMax

    If DDS_SpellIcon(iconnum) Is Nothing Then
        Call InitDDSurf("SpellIcons\" & iconnum, DDSD_SpellIcon(iconnum), DDS_SpellIcon(iconnum))
    End If

    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call Engine_BltToDC(DDS_SpellIcon(iconnum), sRECT, dRECT, frmEditor_Spell.picSprite)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_BltIcon", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_BltAnim()
    Dim Animationnum As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim i As Long
    Dim Width As Long, Height As Long
    Dim looptime As Long
    Dim FrameCount As Long
    Dim ShouldRender As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).value

        If Animationnum < 1 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)

            ShouldRender = False

            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If

            If ShouldRender Then
                frmEditor_Animation.picSprite(i).Cls

                AnimationTimer(Animationnum) = GetTickCount + SurfaceTimerMax

                If DDS_Animation(Animationnum) Is Nothing Then
                    Call InitDDSurf("animations\" & Animationnum, DDSD_Animation(Animationnum), DDS_Animation(Animationnum))
                End If

                If frmEditor_Animation.scrlFrameCount(i).value > 0 Then
                    ' total width divided by frame count
                    Width = DDSD_Animation(Animationnum).lWidth / frmEditor_Animation.scrlFrameCount(i).value
                    Height = DDSD_Animation(Animationnum).lHeight

                    sRECT.top = 0
                    sRECT.Bottom = Height
                    sRECT.Left = (AnimEditorFrame(i) - 1) * Width
                    sRECT.Right = sRECT.Left + Width

                    dRECT.top = 0
                    dRECT.Bottom = Height
                    dRECT.Left = 0
                    dRECT.Right = Width

                    Call Engine_BltToDC(DDS_Animation(Animationnum), sRECT, dRECT, frmEditor_Animation.picSprite(i))
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_BltAnim", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_BltSprite()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_NPC.scrlSprite.value
    frmEditor_NPC.picSprite.Cls

    If Sprite < 1 Or Sprite > NumCharacters Then
        Exit Sub
    End If

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("Characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    Width = DDSD_Character(Sprite).lWidth / 8
    Height = DDSD_Character(Sprite).lHeight / 8

    sRECT.top = 0
    sRECT.Bottom = sRECT.top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width

    dRECT.top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width

    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmEditor_NPC.picSprite)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_BltSprite()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        ResourceTimer(Sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picNormalPic)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        ResourceTimer(Sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picExhaustedPic)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Render_Graphics()
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim rec As DxVBLib.RECT
    Dim rec_pos As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'Para de desenhar para mudar Resolução
    If ChangingResolution Then Exit Sub

    ' check if automation is screwed
    If Not CheckSurfaces Then
        ' exit out and let them know we need to re-init
        ReInitSurfaces = True
        Exit Sub
    Else
        ' if we need to fix the surfaces then do so
        If ReInitSurfaces Then
            ReInitSurfaces = False
            ReInitDD
        End If
    End If

    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub

    ' update the viewpoint
    UpdateCamera

    ' update animation editor
    If Editor = EDITOR_ANIMATION Then
        EditorAnim_BltAnim
    End If

    ' fill it with black
    DDS_BackBuffer.BltColorFill rec_pos, 0

    ' blit lower tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapTile(X, Y)
                End If
            Next
        Next
    End If

    ' Blit out the items
    If numitems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call BltItem(i)
            End If
        Next
    End If

    ' draw animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                BltAnimation i, 0
            End If
        Next
    End If

    ' Render the bars
    BltBars

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For Y = 0 To Map.MaxY
        If NumCharacters > 0 Then

            ' Npcs Morto
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Y = Y Then
                    If MapNpc(i).Desmaiado = True Then
                        Call BltNpc(i)
                    End If
                End If
            Next

            ' Players
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).Y = Y Then

                        If Player(i).Flying = 1 Or Player(i).PuloStatus > 0 Then
                            bltShadowFly i, Player(i).X, Player(i).Y
                            If GetPlayerEquipmentPokeInfoPokemon(i, weapon) = 0 Then Call BltPlayer(i)
                        Else
                            If GetPlayerEquipment(i, weapon) > 0 Then
                                If GetPlayerEquipmentPokeInfoPokemon(i, weapon) > 0 Then
                                    Call BltPlayerPokemon(i)
                                    bltSex i, Player(i).X, Player(i).Y
                                Else
                                    Call BltPlayer(i)
                                End If
                            Else
                                If GetPlayerEquipmentPokeInfoPokemon(i, weapon) = 0 Then Call BltPlayer(i)
                            End If
                        End If

                    End If
                End If
            Next

            ' Npcs Vivo
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Y = Y Then
                    If MapNpc(i).Desmaiado = False Then
                        Call BltNpc(i)
                        Call bltQuest(i, MapNpc(i).X, MapNpc(i).Y)
                        If MapNpc(i).num > 0 Then
                            If Npc(MapNpc(i).num).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(MapNpc(i).num).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
                                bltNpcSex i, MapNpc(i).X, MapNpc(i).Y
                            End If
                        End If
                    End If
                End If
            Next
        End If


        ' Resources
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
                            Call BltMapResource(i)
                        End If
                    Next
                End If
            End If
        End If
    Next

    ' animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                BltAnimation i, 1
            End If
        Next
    End If

    ' blit out upper tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapFringeTile(X, Y)
                End If
            Next
        Next
    End If

    ' blit out a square at mouse cursor
    If InMapEditor Then
        If frmEditor_Map.optBlock.value = True Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.top To TileView.Bottom
                    If IsValidMapPoint(X, Y) Then
                        Call BltDirection(X, Y)
                    End If
                Next
            Next
        End If
        Call BltTileOutline
    End If

    ' Blt the target icon
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            BltTarget (Player(myTarget).X * 32) + Player(myTarget).XOffset, (Player(myTarget).Y * 32) + Player(myTarget).YOffset
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            BltTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).XOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).YOffset
        End If
    End If

    ' blt the hover icon
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Map = Player(MyIndex).Map Then
                If CurX = Player(i).X And CurY = Player(i).Y Then
                    If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                        ' dont render lol
                    Else
                        BltHover TARGET_TYPE_PLAYER, i, (Player(i).X * 32) + Player(i).XOffset, (Player(i).Y * 32) + Player(i).YOffset
                    End If
                End If
            End If
        End If
    Next

    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then
                If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                    ' dont render lol
                Else
                    BltHover TARGET_TYPE_NPC, i, (MapNpc(i).X * 32) + MapNpc(i).XOffset, (MapNpc(i).Y * 32) + MapNpc(i).YOffset
                End If
            End If
        End If
    Next


    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For Y = 0 To Map.MaxY
        If NumCharacters > 0 Then
            ' Players
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).Y = Y Then
                        If Player(i).Flying = 1 Or Player(i).PuloStatus > 0 Then
                            If GetPlayerEquipment(i, weapon) > 0 Then
                                If GetPlayerEquipmentPokeInfoPokemon(i, weapon) > 0 Then
                                    Call BltPlayerPokemon(i)
                                    bltSex i, Player(i).X, Player(i).Y
                                Else
                                    Call BltPlayer(i)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next

    ' Status Negativo
    For i = 1 To Player_HighIndex
        If GetPlayerEquipment(i, weapon) > 0 Then
            If GetPlayerEquipmentPokeInfoPokemon(i, weapon) > 0 Then
                bltNgtStat i, GetPlayerX(i), GetPlayerY(i)
            End If
        End If
    Next

    ' weather
    BltWeather

    ' minimap
    If Options.MiniMap = 1 Then
        BltMiniMap
    End If

    ' Lock the backbuffer so we can draw text and names
    TexthDC = DDS_BackBuffer.GetDC

    'Ping - Ativar Depois
    'DrawPingText

    'Noticia
    DrawNoticiaText

    'Quests Na tela
    If Options.Quest = 1 Then
        DrawQuestsInWindow
    End If

    If Not Map.Moral = 4 Or Player(MyIndex).PokeLight = True Then
        If BFPS Then
            Call DrawText(TexthDC, Camera.Right - (Len("FPS: " & GameFPS) * 8), Camera.top + 1, Trim$("FPS: " & GameFPS), QBColor(Yellow))
        End If
    End If

    ' draw cursor, player X and Y locations
    If BLoc Then
        Call DrawText(TexthDC, Camera.Left, Camera.top + 100, Trim$("cur x: " & CurX & " y: " & CurY), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.top + 115, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.top + 127, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), QBColor(Yellow))
    End If

    ' draw player names
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(i)
            Call DrawPlayerOrg(i)
        End If
    Next

    ' draw npc names
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            Call DrawNpcName(i)
        End If
    Next

    'Draw Mapitem
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(i).num > 0 Then
            If CurX = MapItem(i).X And CurY = MapItem(i).Y Then
                Call DrawMapaItem(i)
            End If
        End If
    Next

    For i = 1 To Action_HighIndex
        Call BltActionMsg(i)
    Next i

    ' Blit out map attributes
    If InMapEditor Then
        Call BltMapAttributes
    End If

    'Contador
    If ContagemGym > 0 Then
        DrawContagem
    End If

    ' Draw map name
    Call DrawText(TexthDC, DrawMapNameX, DrawMapNameY, DrawMapName, DrawMapNameColor)

    ' Release DC
    DDS_BackBuffer.ReleaseDC TexthDC

    ' draw the messages at the very top!
    For i = 1 To MAX_BYTE
        If chatBubble(i).active Then
            DrawChatBubble i
        End If
    Next

    If Map.Moral = 4 And Player(MyIndex).PokeLight = False Then
        BltRockTunel

        ' Lock the backbuffer so we can draw text and names
        TexthDC = DDS_BackBuffer.GetDC

        ' Draw map name
        Call DrawText(TexthDC, DrawMapNameX, DrawMapNameY, DrawMapName, DrawMapNameColor)

        ' draw FPS
        If BFPS Then
            Call DrawText(TexthDC, Camera.Right - (Len("FPS: " & GameFPS) * 8), Camera.top + 1, Trim$("FPS: " & GameFPS), QBColor(Yellow))
        End If

        ' Release DC
        DDS_BackBuffer.ReleaseDC TexthDC

    End If

    ' Get rec
    With rec
        .top = Camera.top
        .Bottom = .top + ScreenY
        .Left = Camera.Left
        .Right = .Left + ScreenX
    End With

    ' rec_pos
    With rec_pos
        .Bottom = ((MAX_MAPY + 1) * PIC_Y)
        .Right = ((MAX_MAPX + 1) * PIC_X)
    End With

    ' Flip and render
    DX7.GetWindowRect frmMain.picScreen.hwnd, rec_pos
    DDS_Primary.Blt rec_pos, DDS_BackBuffer, rec, DDBLT_WAIT

    ' Error handler
    Exit Sub

errorhandler:
    HandleError "Render_Graphics", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateCamera()
    Dim offsetX As Long
    Dim offsetY As Long
    Dim StartX As Long
    Dim StartY As Long
    Dim EndX As Long
    Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    offsetX = Player(MyIndex).XOffset + PIC_X
    offsetY = Player(MyIndex).YOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).XOffset > 0 Then
                offsetX = Player(MyIndex).XOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).YOffset > 0 Then
                offsetY = Player(MyIndex).YOffset
            End If
        End If
        StartY = 0
    End If

    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        offsetX = 32
        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).XOffset < 0 Then
                offsetX = Player(MyIndex).XOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        offsetY = 32
        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).YOffset < 0 Then
                offsetY = Player(MyIndex).YOffset + PIC_Y
            End If
        End If
        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .top = offsetY
        .Bottom = .top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With

    UpdateDrawMapName

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = X - (TileView.Left * PIC_X)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = Y - (TileView.top * PIC_Y)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True

    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim tilesetInUse() As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(X, Y).Layer(i).Tileset > 0 And Map.Tile(X, Y).Layer(i).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(X, Y).Layer(i).Tileset) = True
                End If
            Next
        Next
    Next

    For i = 1 To NumTileSets
        If tilesetInUse(i) Then
            ' load tileset
            If DDS_Tileset(i) Is Nothing Then
                Call InitDDSurf("tilesets\" & i, DDSD_Tileset(i), DDS_Tileset(i))
            End If
        Else
            ' unload tileset
            Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            Set DDS_Tileset(i) = Nothing
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltBank()
    Dim i As Long, X As Long, Y As Long, ItemNum As Long
    Dim Amount As String
    Dim sRECT As RECT, dRECT As RECT
    Dim Sprite As Long, colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.picBank.Visible = True Then
        frmMain.picBank.Cls

        For i = 1 To MAX_BANK
            ItemNum = GetBankItemNum(i)
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then

                Sprite = Item(ItemNum).Pic

                'Icone Pokemon
                If GetPlayerBankItemPokemon(i) > 0 Then
                    Sprite = GetPlayerBankItemPokemon(i)
                    If Sprite >= UNLOCKED_POKEMONS Then Sprite = 1
                End If

                If GetPlayerBankItemPokemon(i) > 0 Then

                    If GetPlayerBankItemShiny(MyIndex, i) > 0 Then
                        If DDS_PokeIconShiny(Sprite) Is Nothing Then
                            Call InitDDSurf("PokeIcon\Shiny\" & Sprite, DDSD_PokeIconShiny(Sprite), DDS_PokeIconShiny(Sprite))
                        End If
                    Else
                        If DDS_PokeIcons(Sprite) Is Nothing Then
                            Call InitDDSurf("PokeIcon\" & Sprite, DDSD_PokeIcons(Sprite), DDS_PokeIcons(Sprite))
                        End If
                    End If

                Else

                    If DDS_Item(Sprite) Is Nothing Then
                        Call InitDDSurf("Items\" & Sprite, DDSD_Item(Sprite), DDS_Item(Sprite))
                    End If

                End If

                With sRECT
                    .top = 0
                    .Bottom = .top + PIC_Y
                    .Left = 32
                    .Right = .Left + PIC_X
                End With

                'Caso pokémon esteja morto
                If GetPlayerBankItemPokemon(i) > 0 Then
                    If GetPlayerBankItemVital(i, 1) = 0 Then

                        With sRECT
                            .top = 0
                            .Bottom = 32
                            .Left = 0
                            .Right = 32
                        End With


                    End If
                End If

                With dRECT
                    .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    .Bottom = .top + PIC_Y
                    .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                    .Right = .Left + PIC_X
                End With

                If GetPlayerBankItemPokemon(i) > 0 Then
                    If GetPlayerBankItemShiny(MyIndex, i) > 0 Then
                        Engine_BltToDC DDS_PokeIconShiny(Sprite), sRECT, dRECT, frmMain.picBank, False
                    Else
                        Engine_BltToDC DDS_PokeIcons(Sprite), sRECT, dRECT, frmMain.picBank, False
                    End If
                Else
                    Engine_BltToDC DDS_Item(Sprite), sRECT, dRECT, frmMain.picBank, False
                End If

                Y = dRECT.top + 22
                X = dRECT.Left - 4

                If Item(Bank.Item(i).num).Type = ITEM_TYPE_ROD Then
                    DrawText frmMain.picBank.hDC, X + 4, Y - 2, GetPlayerBankItemLevel(i), QBColor(BrightGreen)
                End If

                If GetPlayerBankItemPokemon(i) > 0 Then
                    DrawText frmMain.picBank.hDC, X + 4, Y - 2, GetPlayerBankItemLevel(i), QBColor(BrightGreen)

                    If Not Item(GetBankItemNum(i)).Type = ITEM_TYPE_ROD Then
                        If GetPlayerBankItemSexo(MyIndex, i) = 0 Then
                            DrawText frmMain.picBank.hDC, X + 24, Y - 22, "M", QBColor(BrightCyan)
                        Else
                            DrawText frmMain.picBank.hDC, X + 24, Y - 22, "F", QBColor(BrightRed)
                        End If
                    End If

                End If

                ' If item is a stack - draw the amount you have
                If GetBankItemValue(i) > 1 Then

                    Amount = CStr(GetBankItemValue(i))

                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = QBColor(White)
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    DrawText frmMain.picBank.hDC, X, Y, ConvertCurrency(Amount), colour
                End If
            End If
        Next

        frmMain.picBank.Refresh
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBank", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltBankItem(ByVal X As Long, ByVal Y As Long)
    Dim sRECT As RECT, dRECT As RECT
    Dim ItemNum As Long
    Dim Sprite As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = GetBankItemNum(DragBankSlotNum)
    Sprite = Item(GetBankItemNum(DragBankSlotNum)).Pic

    If ItemNum = 3 Then
        Select Case GetPlayerBankItemPokeball(DragBankSlotNum)
        Case 1
            Sprite = 2
        Case 2
            Sprite = 3
        Case 3
            Sprite = 4
        Case 4
            Sprite = 5
        Case 5
            Sprite = 6
        Case 6
            Sprite = 7
        End Select
    End If

    If DDS_Item(Sprite) Is Nothing Then
        Call InitDDSurf("Items\" & Sprite, DDSD_Item(Sprite), DDS_Item(Sprite))
    End If

    If ItemNum > 0 Then
        If ItemNum <= MAX_ITEMS Then
            With sRECT
                .top = 0
                .Bottom = .top + PIC_Y
                .Left = DDSD_Item(Sprite).lWidth / 2
                .Right = .Left + PIC_X
            End With
        End If
    End If

    With dRECT
        .top = 2
        .Bottom = .top + PIC_Y
        .Left = 2
        .Right = .Left + PIC_X
    End With

    Engine_BltToDC DDS_Item(Sprite), sRECT, dRECT, frmMain.picTempBank

    With frmMain.picTempBank
        .top = Y
        .Left = X
        .Visible = True
        .ZOrder (0)
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBankItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltSprite()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_Item.scrlPokemon.value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_Item.picSprite.Cls
        Exit Sub
    End If

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("Characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    Width = DDSD_Character(Sprite).lWidth / 4
    Height = DDSD_Character(Sprite).lHeight / 4

    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height

    sRECT.top = 0
    sRECT.Bottom = sRECT.top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width

    dRECT.top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width

    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmEditor_Item.picSprite)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltTrainerPointMap(ByVal Index As Long)
    Dim Sprite As Long, SpriteTop As Long
    Dim X As Long, Y As Long, Anim As Byte
    Dim rec As DxVBLib.RECT

    Sprite = Player(Index).TPSprite
    Anim = 2

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    ' Setar a direção
    Select Case Player(Index).TPDir
    Case DIR_UP
        SpriteTop = 3
    Case DIR_RIGHT
        SpriteTop = 2
    Case DIR_DOWN
        SpriteTop = 0
    Case DIR_LEFT
        SpriteTop = 1
    End Select

    'Setar recorte da Direção
    With rec
        .top = SpriteTop * (DDSD_Character(Sprite).lHeight / 8)
        .Bottom = .top + (DDSD_Character(Sprite).lHeight / 8)
        .Left = Anim * (DDSD_Character(Sprite).lWidth / 8)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 8)
    End With

    ' Calcular X e Y
    X = Player(Index).TPX * PIC_X - ((DDSD_Character(Sprite).lWidth / 8 - 32) / 2)
    Y = Player(Index).TPY * PIC_Y - ((DDSD_Character(Sprite).lHeight / 8) - 32)

    ' Renderizar Sprite do Jogador
    Call BltSprite(Sprite, X, Y, rec)

    ' Renderizar o Cabelo do Jogador
    If Player(Index).HairNum > 0 Then
        Call BltCabelo(X, Y, Player(Index).HairNum, 2, SpriteTop)
    End If

End Sub

Public Sub EditorPokemon_BltSprite()
    Dim Sprite As Long
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_Pokemon.scrlPokemon.value
    frmEditor_Pokemon.picSprite.Cls

    If Sprite < 1 Or Sprite > NumCharacters Then
        Exit Sub
    End If

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("Characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    Width = DDSD_Character(Sprite).lWidth / 8
    Height = DDSD_Character(Sprite).lHeight / 8

    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height

    sRECT.top = 0
    sRECT.Bottom = sRECT.top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width

    dRECT.top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width

    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmEditor_Pokemon.picSprite)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub bltPokeEvolvePortrait()
    Dim rec As RECT, rec_pos As RECT, faceNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub

    frmMain.PicPokeEvol.Cls

    Select Case EvolTick
    Case 0, 2, 4, 6, 8, 10, 12, 14, 16, 18
        faceNum = GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)
    Case 1, 3, 5, 7, 9, 11, 13, 15, 17
        faceNum = Pokemon(GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)).Evolução(Player(MyIndex).EvoId).Pokemon
    End Select

    If faceNum <= 0 Or faceNum > NumFaces Then Exit Sub

    With rec
        .top = 0
        .Bottom = 94
        .Left = 0
        .Right = 94
    End With

    With rec_pos
        .top = 0
        .Bottom = 94
        .Left = 0
        .Right = 94
    End With

    ' Load face if not loaded, and reset timer
    FaceTimer(faceNum) = GetTickCount + SurfaceTimerMax

    If DDS_Face(faceNum) Is Nothing Then
        Call InitDDSurf("Faces\" & faceNum, DDSD_Face(faceNum), DDS_Face(faceNum))
    End If

    Engine_BltToDC DDS_Face(faceNum), rec, rec_pos, frmMain.PicPokeEvol, False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "bltPokeEvolvePortrait", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltPoke()
    Dim rec As RECT, rec_pos As RECT, faceNum As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub

    frmMain.picFD.Cls

    faceNum = frmMain.ListPokes.ListIndex + 1

    If faceNum <= 0 Or faceNum > NumFaces Then Exit Sub

    With rec
        .top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    With rec_pos
        .top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    ' Load face if not loaded, and reset timer
    FaceTimer(faceNum) = GetTickCount + SurfaceTimerMax

    If DDS_Face(faceNum) Is Nothing Then
        Call InitDDSurf("Faces\" & faceNum, DDSD_Face(faceNum), DDS_Face(faceNum))
    End If

    Engine_BltToDC DDS_Face(faceNum), rec, rec_pos, frmMain.picFD, False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPoke", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltPokeEquip()
    Dim rec As RECT, rec_pos As RECT, faceNum As Long
    Dim i As Long
    Dim X As Long, Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub

    frmMain.PicPokeEquip.Cls

    faceNum = GetPlayerEquipmentPokeInfoPokemon(MyIndex, weapon)

    If faceNum <= 0 Or faceNum > NumFaces Then
        frmMain.PicPokeEquip.Visible = False
        Exit Sub
    Else
        If frmMain.PicPokeEquip.Visible = False Then
            frmMain.PicPokeEquip.Visible = True
        End If
    End If

    With rec
        .top = 0
        .Bottom = 32
        .Left = 32
        .Right = 64
    End With

    With rec_pos
        .top = 0
        .Bottom = .top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    If faceNum > NumPokeIcons Then Exit Sub

    If GetPlayerEquipmentShiny(MyIndex, weapon) > 0 Then
        ' Load face if not loaded, and reset timer
        PokeIconShinyTimer(faceNum) = GetTickCount + SurfaceTimerMax

        If DDS_PokeIconShiny(faceNum) Is Nothing Then
            Call InitDDSurf("PokeIcon\Shiny\" & faceNum, DDSD_PokeIconShiny(faceNum), DDS_PokeIconShiny(faceNum))
        End If

        Engine_BltToDC DDS_PokeIconShiny(faceNum), rec, rec_pos, frmMain.PicPokeEquip, False
    Else
        ' Load face if not loaded, and reset timer
        PokeIconTimer(faceNum) = GetTickCount + SurfaceTimerMax

        If DDS_PokeIcons(faceNum) Is Nothing Then
            Call InitDDSurf("PokeIcon\" & faceNum, DDSD_PokeIcons(faceNum), DDS_PokeIcons(faceNum))
        End If

        Engine_BltToDC DDS_PokeIcons(faceNum), rec, rec_pos, frmMain.PicPokeEquip, False
    End If

    Y = rec_pos.top + 18
    X = rec_pos.Left + 1
    DrawText frmMain.PicPokeEquip.hDC, X, Y, GetPlayerEquipmentPokeInfoLevel(MyIndex, weapon), QBColor(BrightGreen)

    frmMain.PicPokeEquip.Refresh
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPokeEquip", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResetDDS_Primary()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Apagamos nosso desenhista
    Set DDS_Primary = Nothing
    ' Definimos nosso desenhista
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    ' Falamos onde ele deve desenhar
    DDS_Primary.SetClipper DD_Clip

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResetDDS_Primary", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltLeilao()
    Dim i As Long, X As Long, Y As Long, ItemNum As Long
    Dim Amount As String
    Dim sRECT As RECT, dRECT As RECT
    Dim Sprite As Long, colour As Long
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

    If frmMain.picLeilao(2).Visible = True Then
        frmMain.picLeilao(2).Cls

        For i = AB To BA
            ItemNum = Leilao(i).ItemNum
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then

                Sprite = Item(ItemNum).Pic

                'Icone Pokemon
                If Leilao(i).Poke.Pokemon > 0 Then
                    Sprite = Leilao(i).Poke.Pokemon
                End If

                If Leilao(i).Poke.Pokemon > 0 Then

                    If Leilao(i).Poke.Shiny > 0 Then
                        If DDS_PokeIconShiny(Sprite) Is Nothing Then
                            Call InitDDSurf("PokeIcon\Shiny\" & Sprite, DDSD_PokeIconShiny(Sprite), DDS_PokeIconShiny(Sprite))
                        End If
                    Else
                        If DDS_PokeIcons(Sprite) Is Nothing Then
                            Call InitDDSurf("PokeIcon\" & Sprite, DDSD_PokeIcons(Sprite), DDS_PokeIcons(Sprite))
                        End If
                    End If
                Else
                    If DDS_Item(Sprite) Is Nothing Then
                        Call InitDDSurf("Items\" & Sprite, DDSD_Item(Sprite), DDS_Item(Sprite))
                    End If
                End If

                With sRECT
                    .top = 0
                    .Bottom = .top + PIC_Y
                    .Left = 32
                    .Right = .Left + PIC_X
                End With

                If AB >= 21 Then
                    With dRECT
                        .top = LeilaoTop + ((LeilaoOffsetY + 32) * ((i - AB) \ LeilaoColumns))
                        .Bottom = .top + PIC_Y
                        .Left = LeilaoLeft + ((LeilaoOffsetX + 32) * (((i - 1) Mod LeilaoColumns)))
                        .Right = .Left + PIC_X
                    End With
                Else
                    With dRECT
                        .top = LeilaoTop + ((LeilaoOffsetY + 32) * ((i - 1) \ LeilaoColumns))
                        .Bottom = .top + PIC_Y
                        .Left = LeilaoLeft + ((LeilaoOffsetX + 32) * (((i - 1) Mod LeilaoColumns)))
                        .Right = .Left + PIC_X
                    End With
                End If

                If Leilao(i).Poke.Pokemon > 0 Then
                    If Leilao(i).Poke.Shiny > 0 Then
                        Engine_BltToDC DDS_PokeIconShiny(Sprite), sRECT, dRECT, frmMain.picLeilao(2), False
                    Else
                        Engine_BltToDC DDS_PokeIcons(Sprite), sRECT, dRECT, frmMain.picLeilao(2), False
                    End If
                Else
                    Engine_BltToDC DDS_Item(Sprite), sRECT, dRECT, frmMain.picLeilao(2), False
                End If

                Y = dRECT.top + 22
                X = dRECT.Left - 4

                If Item(Leilao(i).ItemNum).Type = ITEM_TYPE_ROD Then
                    DrawText frmMain.picLeilao(2).hDC, X + 4, Y - 2, Leilao(i).Poke.Level, QBColor(BrightGreen)
                End If

                If Leilao(i).Poke.Pokemon > 0 Then
                    DrawText frmMain.picLeilao(2).hDC, X + 4, Y - 2, Leilao(i).Poke.Level, QBColor(BrightGreen)

                    If Not Item(Leilao(i).ItemNum).Type = ITEM_TYPE_ROD Then
                        If Leilao(i).Poke.Sexo = 0 Then
                            DrawText frmMain.picLeilao(2).hDC, X + 24, Y - 22, "M", QBColor(BrightCyan)
                        Else
                            DrawText frmMain.picLeilao(2).hDC, X + 24, Y - 22, "F", QBColor(BrightRed)
                        End If
                    End If

                End If
            End If
        Next

        frmMain.picLeilao(2).Refresh
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBank", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltPokeLeilaoSelect()
    Dim rec As RECT, rec_pos As RECT, faceNum As Long
    Dim i As Long
    Dim X As Long, Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'ItemSelecionado = 0 sair
    If LeilaoItemSelect = 0 Then
        frmMain.picLeilao(3).Visible = False
        Exit Sub
    Else
        frmMain.picLeilao(3).Visible = True
    End If

    'Leilao item = 0 sair
    If Leilao(LeilaoItemSelect).ItemNum = 0 Then
        frmMain.picLeilao(3).Visible = False
        Exit Sub
    Else
        frmMain.picLeilao(3).Visible = True
    End If

    frmMain.picLeilao(3).Cls

    faceNum = Item(Leilao(LeilaoItemSelect).ItemNum).Pic

    If Leilao(LeilaoItemSelect).Poke.Pokemon > 0 Then
        faceNum = Leilao(LeilaoItemSelect).Poke.Pokemon
    End If

    With rec
        .top = 0
        .Bottom = 32
        .Left = 32
        .Right = 64
    End With

    With rec_pos
        .top = 0
        .Bottom = .top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    If Leilao(LeilaoItemSelect).Poke.Pokemon > 0 Then

        If Leilao(LeilaoItemSelect).Poke.Shiny > 0 Then
            ' Load face if not loaded, and reset timer
            PokeIconShinyTimer(faceNum) = GetTickCount + SurfaceTimerMax

            If DDS_PokeIconShiny(faceNum) Is Nothing Then
                Call InitDDSurf("PokeIcon\Shiny\" & faceNum, DDSD_PokeIconShiny(faceNum), DDS_PokeIconShiny(faceNum))
            End If

            Engine_BltToDC DDS_PokeIconShiny(faceNum), rec, rec_pos, frmMain.picLeilao(3), False
        Else
            ' Load face if not loaded, and reset timer
            PokeIconTimer(faceNum) = GetTickCount + SurfaceTimerMax

            If DDS_PokeIcons(faceNum) Is Nothing Then
                Call InitDDSurf("PokeIcon\" & faceNum, DDSD_PokeIcons(faceNum), DDS_PokeIcons(faceNum))
            End If

            Engine_BltToDC DDS_PokeIcons(faceNum), rec, rec_pos, frmMain.picLeilao(3), False
        End If

    Else

        If DDS_Item(faceNum) Is Nothing Then
            Call InitDDSurf("items\" & faceNum, DDSD_Item(faceNum), DDS_Item(faceNum))
        End If

        Engine_BltToDC DDS_Item(faceNum), rec, rec_pos, frmMain.picLeilao(3), False

    End If

    Y = rec_pos.top + 18
    X = rec_pos.Left + 1

    If Leilao(LeilaoItemSelect).Poke.Level > 0 Then
        DrawText frmMain.picLeilao(3).hDC, X, Y, Leilao(LeilaoItemSelect).Poke.Level, QBColor(BrightGreen)
    End If

    frmMain.picLeilao(3).Refresh
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPokeEquip", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub bltSex(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim sRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long
    Dim Play As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Width = DDSD_SaN.lWidth
    Height = DDSD_SaN.lHeight

    With sRECT
        .top = 0 * (Height / 2)
        .Bottom = Height / 2

        If GetPlayerEquipmentSexo(Index, weapon) = 0 Then
            .Left = 0 * (Width / 7)
        Else
            .Left = 1 * (Width / 7)
        End If

        .Right = .Left + (Width / 7)
    End With

    X = ConvertMapX(X * 32)
    Y = ConvertMapY(Y * 32)

    ' /clipping
    If Player(Index).Flying = 1 Then
        Call Engine_BltFast(X + Player(Index).XOffset + 18, Y + Player(Index).YOffset - 37, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        Call Engine_BltFast(X + Player(Index).XOffset + 18, Y + Player(Index).YOffset - 15, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSex", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub bltNpcSex(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim sRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long
    Dim Play As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Width = DDSD_SaN.lWidth
    Height = DDSD_SaN.lHeight

    With sRECT
        .top = 0 * (Height / 2)
        .Bottom = Height / 2

        If MapNpc(Index).Sexo = 0 Then
            .Left = 0 * (Width / 7)
        Else
            .Left = 1 * (Width / 7)
        End If

        .Right = .Left + (Width / 7)
    End With

    X = ConvertMapX(X * 32)
    Y = ConvertMapY(Y * 32)

    ' /clipping
    Call Engine_BltFast(X + MapNpc(Index).XOffset + 18, Y + MapNpc(Index).YOffset - 15, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltNpcSex", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub bltShadowFly(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim sRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long
    Dim Play As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Width = DDSD_SaN.lWidth
    Height = DDSD_SaN.lHeight

    With sRECT
        .top = 1 * (Height / 2)
        .Bottom = .top + Height / 2
        .Left = 6 * (Width / 7)
        .Right = .Left + (Width / 7)
    End With

    If Not Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED And Not Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        X = ConvertMapX(X * 32)
        Y = ConvertMapY(Y * 32)

        Call Engine_BltFast(X + Player(Index).XOffset, Y + Player(Index).YOffset, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltNpcSex", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltWeather()
    Dim i As Long, sRECT As RECT

    ' rain
    If Map.Weather = WEATHER_RAINING Then
        Call DDS_BackBuffer.SetForeColor(RGB(255, 255, 255))
        For i = 1 To Map.Intensity
            With DropRain(i)
                If .Init = True Then
                    ' move o snow
                    .Y = .Y + .ySpeed + 10
                    .X = .X + .ySpeed + 10
                    ' checar a screen
                    If .Y > 600 + 64 Then
                        .Y = Rand(0, 100) - 30
                        .X = Rand(0, 1200 + 64)
                        .ySpeed = Rand(1, 4)
                        .xSpeed = Rand(0, 4)
                    End If
                    ' draw rain
                    With sRECT
                        .top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With
                    Engine_BltFast .X + Camera.Left - 500, .Y + Camera.top, DDS_Weather, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    .Y = Rand(0, 600)
                    .X = Rand(0, 1200 + 64)
                    .ySpeed = Rand(1, 4)
                    .xSpeed = Rand(0, 4) - 2
                    .Init = True
                End If
            End With
        Next
    End If

    ' snow
    If Map.Weather = WEATHER_SNOWING Then
        Call DDS_BackBuffer.SetForeColor(RGB(255, 255, 255))
        For i = 1 To Map.Intensity
            With DropSnow(i)
                If .Init = True Then
                    ' Move o snow
                    .Y = .Y + .ySpeed
                    .X = .X + .xSpeed
                    ' checar screen
                    If .Y > 600 + 64 Then
                        .Y = Rand(0, 100) - 100
                        .X = Rand(0, 800 + 64)
                        .ySpeed = Rand(1, 4)
                        .xSpeed = Rand(0, 4) - 2
                    End If

                    ' draw Snow
                    With sRECT
                        .top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With

                    Engine_BltFast .X + Camera.Left, .Y + Camera.top, DDS_Weather, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    .Y = Rand(0, 600)
                    .X = Rand(0, 800 + 64)
                    .ySpeed = Rand(1, 4)
                    .xSpeed = Rand(0, 4) - 2
                    .Init = True
                End If
            End With
        Next
    End If

    If Map.Weather = WEATHER_BIRD Then

        For i = 1 To Map.Intensity
            With DropBird(i)
                If .Init = True Then
                    ' move o Sand
                    .Y = .Y + .ySpeed
                    .X = .X + .xSpeed
                    ' checar a screen
                    If .Y > 600 + 64 Then
                        .Y = Rand(0, 100) - 100
                        .X = Rand(0, 800 + 64)
                        .ySpeed = Rand(1, 4)
                        .xSpeed = Rand(0, 4) - 2
                    End If
                    ' draw rain
                    With sRECT
                        .top = 0
                        .Bottom = 32
                        .Left = 96
                        .Right = 128
                    End With
                    Engine_BltFast .X + Camera.Left, .Y + Camera.top, DDS_Weather, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    .Y = Rand(0, 600)
                    .X = Rand(0, 800 + 64)
                    .ySpeed = Rand(1, 4)
                    .xSpeed = Rand(0, 4) - 2
                    .Init = True
                End If
            End With
        Next
    End If

    If Map.Weather = WEATHER_SAND Then

        For i = 1 To Map.Intensity
            With DropSand(i)
                If .Init = True Then
                    ' move o Sand
                    .Y = .Y + .ySpeed
                    .X = .X + .xSpeed
                    ' checkar a screen
                    If .Y > 600 + 64 Then
                        .Y = Rand(0, 100) - 100
                        .X = Rand(0, 800 + 64)
                        .ySpeed = Rand(1, 4)
                        .xSpeed = Rand(0, 4) - 2
                    End If
                    ' draw rain
                    With sRECT
                        .top = 0
                        .Bottom = 32
                        .Left = 64
                        .Right = 96
                    End With
                    Engine_BltFast .X + Camera.Left, .Y + Camera.top, DDS_Weather, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                Else
                    .Y = Rand(0, 600)
                    .X = Rand(0, 800 + 64)
                    .ySpeed = Rand(1, 4)
                    .xSpeed = Rand(0, 4) - 2
                    .Init = True
                End If
            End With
        Next
    End If
End Sub

Public Sub bltQuest(ByVal NpcNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim sRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long
    Dim Play As Long, QuestStatus As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(NpcNum).num = 0 Then Exit Sub
    If Npc(MapNpc(NpcNum).num).Behaviour <> NPC_BEHAVIOUR_QUEST Then Exit Sub
    If Npc(MapNpc(NpcNum).num).Quest(1) = 0 Then Exit Sub

    Select Case Player(MyIndex).Quests(Npc(MapNpc(NpcNum).num).Quest(1)).status
    Case 0, 1, 2
        QuestStatus = Player(MyIndex).Quests(Npc(MapNpc(NpcNum).num).Quest(1)).status
    Case 4
        QuestStatus = Player(MyIndex).Quests(Npc(MapNpc(NpcNum).num).Quest(1)).status - 1
    Case 3    'Finalizada
        Exit Sub
    End Select

    Width = DDSD_Quest.lWidth
    Height = DDSD_Quest.lHeight

    With sRECT
        .top = QuestStatus * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = AnimQuest * (Width / 5)
        .Right = .Left + (Width / 5)
    End With

    X = ConvertMapX(X * 32)
    Y = ConvertMapY(Y * 32)

    ' /clipping
    Call Engine_BltFast(X + MapNpc(NpcNum).XOffset, Y + MapNpc(NpcNum).YOffset - 50, DDS_Quest, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltNpcSex", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltRockTunel()
    Dim sRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long
    Dim Play As Long
    Dim Index As Long, X As Long, Y As Long

    Index = MyIndex

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Width = DDSD_RockTunel.lWidth
    Height = DDSD_RockTunel.lHeight

    With sRECT
        .top = 1 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 0 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    For X = -1 To Map.MaxX + 1
        For Y = -1 To Map.MaxY + 1
            If Not isInRange(2, X, Y, GetPlayerX(Index), GetPlayerY(Index)) Then
                Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next
    Next

    'Parte1
    With sRECT
        .top = 0 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 2 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index)
    Y = GetPlayerY(Index) - 2

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    'Parte2
    With sRECT
        .top = 0 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 1 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index) - 1
    Y = GetPlayerY(Index) - 2

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    'Parte 3
    With sRECT
        .top = 0 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 3 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index) + 1
    Y = GetPlayerY(Index) - 2

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    'Parte 4
    With sRECT
        .top = 1 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 1 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index) - 1
    Y = GetPlayerY(Index) - 1

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    'Parte 5
    With sRECT
        .top = 1 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 3 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index) + 1
    Y = GetPlayerY(Index) - 1

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)


    'Parte 6
    With sRECT
        .top = 2 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 1 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index) - 1
    Y = GetPlayerY(Index)

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)


    'Parte 7
    With sRECT
        .top = 2 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 3 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index) + 1
    Y = GetPlayerY(Index)

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)


    'Parte 8
    With sRECT
        .top = 3 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 1 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index) - 1
    Y = GetPlayerY(Index) + 1

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    'Parte 9
    With sRECT
        .top = 3 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 2 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index)
    Y = GetPlayerY(Index) + 1

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)


    'Parte 10
    With sRECT
        .top = 3 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 3 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index) + 1
    Y = GetPlayerY(Index) + 1

    Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)


    'Fechar Parte Restantes
    With sRECT
        .top = 1 * (Height / 4)
        .Bottom = .top + Height / 4
        .Left = 0 * (Width / 4)
        .Right = .Left + (Width / 4)
    End With

    X = GetPlayerX(Index) - 2
    For Y = GetPlayerY(Index) - 1 To GetPlayerY(Index) + 1
        Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Next

    X = GetPlayerX(Index) + 2
    For Y = GetPlayerY(Index) - 1 To GetPlayerY(Index) + 1
        Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Next

    X = GetPlayerY(Index) + 2
    For X = GetPlayerX(Index) - 1 To GetPlayerX(Index) + 1
        Call Engine_BltFast((X + 1) * 32 + Player(Index).XOffset, (Y + 1) * 32 + Player(Index).YOffset, DDS_RockTunel, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltRockTunel", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub bltNgtStat(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim sRECT As DxVBLib.RECT
    Dim Width As Long, Height As Long
    Dim Play As Long, Ordem As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerEquipment(Index, weapon) = 0 Then Exit Sub

    Width = DDSD_SaN.lWidth
    Height = DDSD_SaN.lHeight

    'Cordenadas
    X = ConvertMapX(X * 32)
    Y = ConvertMapY(Y * 32)

    'Burn
    If GetPlayerEquipmentNgt(Index, weapon, 1) > 0 Then
        With sRECT
            .top = 0 * (Height / 2)
            .Bottom = .top + Height / 2
            .Left = 2 * (Width / 7)
            .Right = .Left + (Width / 7)
        End With

        Select Case Ordem
        Case 0
            Call Engine_BltFast(X - 20 + Player(Index).XOffset, Y + 7 + Player(Index).YOffset, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Case Else
            Call Engine_BltFast((X - 20) + (Ordem * 13) + Player(Index).XOffset, Y + 7 + Player(Index).YOffset, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End Select

        Ordem = Ordem + 1
    End If

    'Frozen
    If GetPlayerEquipmentNgt(Index, weapon, 2) > 0 Then
        With sRECT
            .top = 0 * (Height / 2)
            .Bottom = .top + Height / 2
            .Left = 3 * (Width / 7)
            .Right = .Left + (Width / 7)
        End With

        Select Case Ordem
        Case 0
            Call Engine_BltFast(X - 20 + Player(Index).XOffset, Y + 7 + Player(Index).YOffset, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Case Else
            Call Engine_BltFast((X - 20) + (Ordem * 13) + Player(Index).XOffset, Y + 7 + Player(Index).YOffset, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End Select

        Ordem = Ordem + 1
    End If

    'Poison
    If GetPlayerEquipmentNgt(Index, weapon, 3) > 0 Then
        With sRECT
            .top = 0 * (Height / 2)
            .Bottom = .top + Height / 2
            .Left = 4 * (Width / 7)
            .Right = .Left + (Width / 7)
        End With

        Select Case Ordem
        Case 0
            Call Engine_BltFast(X - 20 + Player(Index).XOffset, Y + 7 + Player(Index).YOffset, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Case Else
            Call Engine_BltFast((X - 20) + (Ordem * 13) + Player(Index).XOffset, Y + 7 + Player(Index).YOffset, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End Select

        Ordem = Ordem + 1
    End If

    'Atract
    If GetPlayerEquipmentNgt(Index, weapon, 4) > 0 Then
        With sRECT
            .top = 0 * (Height / 2)
            .Bottom = .top + Height / 2
            .Left = 5 * (Width / 7)
            .Right = .Left + (Width / 7)
        End With

        Select Case Ordem
        Case 0
            Call Engine_BltFast(X - 20 + Player(Index).XOffset, Y + 7 + Player(Index).YOffset, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Case Else
            Call Engine_BltFast((X - 20) + (Ordem * 13) + Player(Index).XOffset, Y + 7 + Player(Index).YOffset, DDS_SaN, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End Select
        Ordem = Ordem + 1
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "bltNgtStat", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function IsInTileView(ByVal TileX As Long, ByVal TileY As Long)
    Dim X As Long, Y As Long

    For X = TileView.Left To TileView.Right
        For Y = TileView.top To TileView.Bottom
            If TileX = X Then
                If TileY = Y Then
                    IsInTileView = True
                End If
            End If
        Next
    Next

End Function

Sub BltMiniMap()
    Dim i As Long
    Dim X As Integer, Y As Integer
    Dim Direction As Byte
    Dim CameraX As Long, CameraY As Long
    Dim BlockRect As RECT, WarpRect As RECT, ItemRect As RECT, ShopRect As RECT, NpcOtherRect As RECT, PlayerRect As RECT, PlayerPkRect As RECT, NpcAttackerRect As RECT, NpcShopRect As RECT, NadaRect As RECT
    Dim MapX As Long, MapY As Long
    Dim LocMapX As Long, LocMapY As Long
    Dim BordaRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    MapX = 24
    MapY = 17
    LocMapX = 674
    LocMapY = 14

    ' ************
    ' *** Nada ***
    ' ************
    With NadaRect
        .top = 4
        .Bottom = .top + 4
        .Left = 0
        .Right = .Left + 4
    End With

    ' Defini-lo no minimap
    For X = TileView.Left To TileView.Right
        For Y = TileView.top To TileView.Bottom
            CameraX = Camera.Left + LocMapX + (X * 4) - (TileView.Left * 4)
            CameraY = Camera.top + LocMapY + (Y * 4) - (TileView.top * 4)
            Engine_BltFast CameraX, CameraY, DDS_MiniMap, NadaRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        Next Y
    Next X

    ' *****************
    ' *** Atributos ***
    ' *****************

    ' Bloqueio
    With BlockRect
        .top = 4
        .Bottom = .top + 4
        .Left = 4
        .Right = .Left + 4
    End With

    ' Warp
    With WarpRect
        .top = 4
        .Bottom = .top + 4
        .Left = 8
        .Right = .Left + 4
    End With

    ' Item
    With ItemRect
        .top = 4
        .Bottom = .top + 4
        .Left = 12
        .Right = .Left + 4
    End With

    ' Shop
    With ShopRect
        .top = 4
        .Bottom = .top + 4
        .Left = 16
        .Right = .Left + 4
    End With

    ' Defini-los no minimap
    For X = TileView.Left To TileView.Right
        For Y = TileView.top To TileView.Bottom
            If X >= 0 And X <= Map.MaxX Then
                If Y >= 0 And Y <= Map.MaxY Then

                    Select Case Map.Tile(X, Y).Type
                    Case TILE_TYPE_BLOCKED
                        CameraX = Camera.Left + LocMapX + (X * 4) - (TileView.Left * 4)
                        CameraY = Camera.top + LocMapY + (Y * 4) - (TileView.top * 4)
                        Engine_BltFast CameraX, CameraY, DDS_MiniMap, BlockRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    Case TILE_TYPE_WARP
                        CameraX = Camera.Left + LocMapX + (X * 4) - (TileView.Left * 4)
                        CameraY = Camera.top + LocMapY + (Y * 4) - (TileView.top * 4)
                        Engine_BltFast CameraX, CameraY, DDS_MiniMap, WarpRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    Case TILE_TYPE_ITEM
                        CameraX = Camera.Left + LocMapX + (X * 4) - (TileView.Left * 4)
                        CameraY = Camera.top + LocMapY + (Y * 4) - (TileView.top * 4)
                        Engine_BltFast CameraX, CameraY, DDS_MiniMap, ItemRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    Case TILE_TYPE_SHOP
                        CameraX = Camera.Left + LocMapX + (X * 4) - (TileView.Left * 4)
                        CameraY = Camera.top + LocMapY + (Y * 4) - (TileView.top * 4)
                        Engine_BltFast CameraX, CameraY, DDS_MiniMap, ShopRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End Select

                End If
            End If
        Next Y
    Next X

    ' **************
    ' *** Player ***
    ' **************

    ' Normal
    With PlayerRect
        .top = 0
        .Bottom = .top + 4
        .Left = 4
        .Right = .Left + 4
    End With

    ' Pk
    With PlayerPkRect
        .top = 0
        .Bottom = .top + 4
        .Left = 8
        .Right = .Left + 4
    End With

    ' Defini-los no minimap
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If IsInTileView(Player(i).X, Player(i).Y) = True Then
                X = Player(i).X
                Y = Player(i).Y
                CameraX = Camera.Left + LocMapX + (X * 4) - (TileView.Left * 4)
                CameraY = Camera.top + LocMapY + (Y * 4) - (TileView.top * 4)
                Call DDS_BackBuffer.BltFast(CameraX, CameraY, DDS_MiniMap, PlayerRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Next i

    ' ***********
    ' *** NPC ***
    ' ***********

    ' Atacar ao ser atacado e quando for atacado
    With NpcAttackerRect
        .top = 0
        .Bottom = .top + 4
        .Left = 12
        .Right = .Left + 4
    End With

    ' Vendendor
    With NpcShopRect
        .top = 0
        .Bottom = .top + 4
        .Left = 16
        .Right = .Left + 4
    End With

    ' Outros
    With NpcOtherRect
        .top = 0
        .Bottom = .top + 4
        .Left = 20
        .Right = .Left + 4
    End With

    ' Defini-lo no minimap
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If IsInTileView(MapNpc(i).X, MapNpc(i).Y) = True Then
                Select Case Npc(MapNpc(i).num).Behaviour
                Case NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC_BEHAVIOUR_ATTACKWHENATTACKED
                    X = MapNpc(i).X
                    Y = MapNpc(i).Y
                    CameraX = Camera.Left + LocMapX + (X * 4) - (TileView.Left * 4)
                    CameraY = Camera.top + LocMapY + (Y * 4) - (TileView.top * 4)
                    Call DDS_BackBuffer.BltFast(CameraX, CameraY, DDS_MiniMap, NpcAttackerRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Case NPC_BEHAVIOUR_SHOPKEEPER
                    X = MapNpc(i).X
                    Y = MapNpc(i).Y
                    CameraX = Camera.Left + LocMapX + (X * 4) - (TileView.Left * 4)
                    CameraY = Camera.top + LocMapY + (Y * 4) - (TileView.top * 4)
                    Call DDS_BackBuffer.BltFast(CameraX, CameraY, DDS_MiniMap, NpcShopRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Case Else
                    X = MapNpc(i).X
                    Y = MapNpc(i).Y
                    CameraX = Camera.Left + LocMapX + (X * 4) - (TileView.Left * 4)
                    CameraY = Camera.top + LocMapY + (Y * 4) - (TileView.top * 4)
                    Call DDS_BackBuffer.BltFast(CameraX, CameraY, DDS_MiniMap, NpcOtherRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End Select
            End If
        End If
    Next i

    ' ************
    ' *** Borda ***
    ' ************
    With BordaRect
        .top = 0
        .Bottom = 0
        .Left = 0
        .Right = 0
    End With

    ' Defini-lo no minimap
    CameraX = Camera.Left + 669
    CameraY = Camera.top + 9
    Engine_BltFast CameraX, CameraY, DDS_MapBorda, BordaRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY


    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMiniMap", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltOrgShop()
    Dim rec As RECT, rec_pos As RECT
    Dim i As Long, itempic As Long
    Dim a As Byte, B As Byte, X As Integer, Y As Integer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.PicOrg(1).Cls

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

        If OrgShop(i).Item = 0 Then GoTo Atualizar
        itempic = Item(OrgShop(i).Item).Pic

        If itempic = 0 Then Exit Sub

        With rec
            .top = 0
            .Bottom = PIC_Y
            .Left = 32
            .Right = PIC_X + 32
        End With

        If a = 1 Then
            With rec_pos
                .top = OrgTop + ((OrgOffsetY + 32) * ((i - 1) \ OrgColumns))
                .Bottom = .top + PIC_Y
                .Left = OrgLeft + ((OrgOffsetX + 32) * (((i - 1) Mod OrgColumns)))
                .Right = .Left + PIC_X
            End With
        Else
            With rec_pos
                .top = OrgTop + ((OrgOffsetY + 32) * ((i - a) \ OrgColumns))
                .Bottom = .top + PIC_Y
                .Left = OrgLeft + ((OrgOffsetX + 32) * (((i - a) Mod OrgColumns)))
                .Right = .Left + PIC_X
            End With
        End If

        ' Load item if not loaded, and reset timer
        ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

        If DDS_Item(itempic) Is Nothing Then
            Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
        End If

        Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.PicOrg(1), False

        X = rec_pos.Left + 40
        Y = rec_pos.top

        If Item(OrgShop(i).Item).Type = ITEM_TYPE_CURRENCY Then
            DrawText frmMain.PicOrg(1).hDC, X, Y, OrgShop(i).Quantia & " " & Trim$(Item(OrgShop(i).Item).Name), QBColor(White)
        Else
            DrawText frmMain.PicOrg(1).hDC, X, Y, Trim$(Item(OrgShop(i).Item).Name), QBColor(White)
        End If

        Y = rec_pos.top + 15
        DrawText frmMain.PicOrg(1).hDC, X, Y, "Honra: " & OrgShop(i).Valor & " Org Lvl: " & OrgShop(i).Level, QBColor(White)
    Next

Atualizar:
    X = 11
    Y = 242
    DrawText frmMain.PicOrg(1).hDC, X, Y, "Honra: " & GetPlayerHonra(MyIndex), QBColor(BrightGreen)
    frmMain.PicOrg(1).Refresh

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltOrgShop", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltItemSelectOrgShop()
    Dim rec As RECT, rec_pos As RECT
    Dim i As Long, itempic As Long
    Dim a As Byte, B As Byte, X As Integer, Y As Integer

    frmMain.PicOrg(2).Cls

    'Desenhar Item Selecionado
    If DragOrgShopNum = 0 Then GoTo Atualizar:

    itempic = Item(OrgShop(DragOrgShopNum).Item).Pic

    If itempic = 0 Then GoTo Atualizar:

    With rec
        .top = 0
        .Bottom = PIC_Y
        .Left = 32
        .Right = PIC_X + 32
    End With

    With rec_pos
        .top = 11
        .Bottom = .top + PIC_Y
        .Left = 3
        .Right = .Left + PIC_X
    End With

    ' Carregar Timer
    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itempic) Is Nothing Then
        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
    End If

    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.PicOrg(2), False

    'Renderizar nome Item Selecionado
    i = DragOrgShopNum
    X = rec_pos.Left + 40
    Y = rec_pos.top

    If Item(OrgShop(i).Item).Type = ITEM_TYPE_CURRENCY Then
        DrawText frmMain.PicOrg(2).hDC, X, Y, OrgShop(i).Quantia & " " & Trim$(Item(OrgShop(i).Item).Name), QBColor(White)
    Else
        DrawText frmMain.PicOrg(2).hDC, X, Y, Trim$(Item(OrgShop(i).Item).Name), QBColor(White)
    End If

    Y = rec_pos.top + 15
    DrawText frmMain.PicOrg(2).hDC, X, Y, "Honra: " & OrgShop(i).Valor & " Org Lvl: " & OrgShop(i).Level, QBColor(White)

Atualizar:
    frmMain.PicOrg(2).Refresh
End Sub

Sub BltOrganização()
    Dim rec As RECT, rec_pos As RECT
    Dim X As Integer, Y As Integer, i As Byte
    Dim Width As Long, Height As Long
    Dim a As Byte, B As Byte, c As Byte, U As String
    Dim Formula As String

    'Evitar OverFlow
    If GetPlayerOrg(MyIndex) = 0 Then Exit Sub

    Width = DDSD_OrgOn.lWidth / 2
    Height = DDSD_OrgOn.lHeight
    '
    frmMain.PicOrgs.Cls

    'Organizar
    Select Case OrgPagMem
    Case 1
        a = 1
        B = 9
        c = 1
    Case 2
        a = 10
        B = 18
        c = 10
    Case 3
        a = 19
        B = 27
        c = 19
    Case 4
        a = 28
        B = 36
        c = 28
    End Select

    'Org Names
    Select Case Player(MyIndex).ORG

    Case 1
        U = "Equipe Rocket"

        ' Carrega a Face da organização
        frmMain.faceOrg.Picture = LoadPicture(App.Path & "\data files\graphics\orgs\1.bmp")
    Case 2
        U = "Team Magma"

        ' Carrega a Face da organização
        frmMain.faceOrg.Picture = LoadPicture(App.Path & "\data files\graphics\orgs\2.bmp")
    Case 3
        U = "Team Aqua"

        ' Carrega a Face da organização
        frmMain.faceOrg.Picture = LoadPicture(App.Path & "\data files\graphics\orgs\3.bmp")

    Case Else

        Exit Sub

    End Select

    If MaxExpOrg > 0 Then
        Formula = Format(((Organization(Player(MyIndex).ORG).Exp / ORG) / (MaxExpOrg / ORG)), "0.0000") * 100 & "%"
    Else
        Formula = 100
    End If

    frmMain.lblOrg(1).Caption = U
    frmMain.lblOrg(2).Caption = "Level: " & Organization(GetPlayerOrg(MyIndex)).Level
    
    'Loc Text Lista Membros
    X = 35
    Y = 110

    For i = a To B
        If Organization(GetPlayerOrg(MyIndex)).OrgMember(i).Used = True Then
            If i = 1 Then
                DrawText frmMain.PicOrgs.hDC, X, Y + ((i - c) * 17), "Líder: " & Trim$(Organization(GetPlayerOrg(MyIndex)).OrgMember(i).User_Name), QBColor(White)
            Else
                DrawText frmMain.PicOrgs.hDC, X, Y + ((i - c) * 17), i - 1 & ": " & Trim$(Organization(GetPlayerOrg(MyIndex)).OrgMember(i).User_Name), QBColor(White)
            End If

            If Organization(GetPlayerOrg(MyIndex)).OrgMember(i).Online = True Then

                With rec
                    .top = 0
                    .Bottom = Height
                    .Left = 12
                    .Right = .Left + Width
                End With

                With rec_pos
                    .top = 112 + ((i - c) * 17)
                    .Bottom = .top + Height
                    .Left = 12
                    .Right = .Left + Width
                End With

                Engine_BltToDC DDS_OrgOn, rec, rec_pos, frmMain.PicOrgs, False
            Else

                With rec
                    .top = 0
                    .Bottom = .top + Height
                    .Left = 0
                    .Right = .Left + Width
                End With

                With rec_pos
                    .top = 112 + ((i - c) * 17)
                    .Bottom = .top + Height
                    .Left = 12
                    .Right = .Left + Width
                End With

                Engine_BltToDC DDS_OrgOn, rec, rec_pos, frmMain.PicOrgs, False
            End If
        End If
    Next

Atualizar:
    frmMain.PicOrgs.Refresh
End Sub

Public Sub BltQuestRewards()
    Dim i As Long, X As Long, Y As Long, ItemNum As Long, itempic As Long, QuestNum As Long
    Dim Amount As String
    Dim rec As RECT, rec_pos As RECT
    Dim colour As Long
    Dim a As Byte, B As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.picQuest.Cls

    'Verificar Numero da Quest Selecionada
    QuestNum = GetQuestNum(Trim$(frmMain.lstQuests.text))

    'Checar Page
    Select Case RewardsPage
    Case 1
        a = 6
        B = 10
    Case Else
        a = 1
        B = 5
    End Select

    For i = a To B
        If QuestNum > 0 Then ItemNum = Quest(QuestNum).ItemRew(i)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then

            'Numero do Icone
            If Quest(QuestNum).PokeRew(i) = 0 Then
                itempic = Item(ItemNum).Pic
            Else
                itempic = Quest(QuestNum).PokeRew(i)
            End If

            If itempic > 0 And itempic <= numitems Then

                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                If i <= 5 Then
                    With rec_pos
                        .top = 231
                        .Bottom = .top + PIC_Y
                        .Left = 381 + (i - 1 Mod 5) * 40
                        .Right = .Left + PIC_X
                        X = .Left
                    End With
                Else
                    With rec_pos
                        .top = 231
                        .Bottom = .top + PIC_Y
                        .Left = 381 + (i - a Mod 10) * 40
                        .Right = .Left + PIC_X
                        X = .Left
                    End With
                End If

                If Quest(QuestNum).PokeRew(i) = 0 Then
                    'Timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    'Carregar Se não estiver carregado
                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    'Renderizar
                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picQuest, False

                Else
                    'Timer
                    PokeIconTimer(itempic) = GetTickCount + SurfaceTimerMax

                    'Carregar Se não estiver carreg
                    If DDS_PokeIcons(itempic) Is Nothing Then
                        Call InitDDSurf("PokeIcon\" & itempic, DDSD_PokeIcons(itempic), DDS_PokeIcons(itempic))
                    End If

                    'Renderizar
                    Engine_BltToDC DDS_PokeIcons(itempic), rec, rec_pos, frmMain.picQuest, False
                End If

            End If
        End If

        If QuestNum > 0 Then
            If Quest(QuestNum).PokeRew(i) > 0 Then
                DrawText frmMain.picQuest.hDC, X + 3, 250, Quest(QuestNum).ValueRew(i), QBColor(BrightGreen)
            Else
                If Quest(QuestNum).ValueRew(i) > 1 Then
                    DrawText frmMain.picQuest.hDC, X + 3, 250, Quest(QuestNum).ValueRew(i), QBColor(White)
                End If
            End If
        End If

        frmMain.picQuest.Refresh
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltShop", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
    Dim theArray() As String, X As Long, Y As Long, i As Long, MaxWidth As Long, xwidth As Long, yheight As Long, colour As Long, x3 As Long, y3 As Long

    Dim MMx As Long
    Dim MMy As Long

    Dim TOPLEFTrect As RECT
    Dim TOPCENTERrect As RECT
    Dim TOPRIGHTrect As RECT
    Dim MIDDLELEFTrect As RECT
    Dim MIDDLECENTERrect As RECT
    Dim MIDDLERIGHTrect As RECT
    Dim BOTTOMLEFTrect As RECT
    Dim BOTTOMCENTERrect As RECT
    Dim BOTTOMRIGHTrect As RECT
    Dim TIPrect As RECT

    ' DESIGNATE CHATBUBBLE SECTIONS FROM CHATBUBBLE IMAGE
    With TOPRIGHTrect
        .top = 0
        .Bottom = .top + 4
        .Left = 0
        .Right = .Left + 4
    End With

    With TOPCENTERrect
        .top = 0
        .Bottom = .top + 4
        .Left = 4
        .Right = .Left + 4
    End With

    With TOPLEFTrect
        .top = 0
        .Bottom = .top + 4
        .Left = 8
        .Right = .Left + 4
    End With

    With MIDDLERIGHTrect
        .top = 4
        .Bottom = .top + 4
        .Left = 0
        .Right = .Left + 4
    End With

    With MIDDLECENTERrect
        .top = 4
        .Bottom = .top + 4
        .Left = 4
        .Right = .Left + 4
    End With

    With MIDDLELEFTrect
        .top = 4
        .Bottom = .top + 4
        .Left = 8
        .Right = .Left + 4
    End With

    With BOTTOMRIGHTrect
        .top = 8
        .Bottom = .top + 4
        .Left = 0
        .Right = .Left + 4
    End With

    With BOTTOMCENTERrect
        .top = 8
        .Bottom = .top + 4
        .Left = 4
        .Right = .Left + 4
    End With

    With BOTTOMLEFTrect
        .top = 8
        .Bottom = .top + 4
        .Left = 8
        .Right = .Left + 4
    End With

    With TIPrect
        .top = 12
        .Bottom = .top + 4
        .Left = 0
        .Right = .Left + 4
    End With

    Call DDS_BackBuffer.SetForeColor(RGB(255, 255, 255))

    With chatBubble(Index)
        If .targetType = TARGET_TYPE_PLAYER Then
            ' it's a player
            If GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then
                ' change the colour depending on access
                colour = QBColor(Yellow)

                If Player(.target).TPX = 0 Then
                    X = ConvertMapX((Player(.target).X * 32) + Player(.target).XOffset) + 12
                    Y = ConvertMapY((Player(.target).Y * 32) + Player(.target).YOffset) - 21
                Else
                    X = ConvertMapX(Player(.target).TPX * 32) + 12
                    Y = ConvertMapY(Player(.target).TPY * 32) - 21
                End If

                ' word wrap the text
                WordWrap_Array .Msg, 80, theArray

                ' find max width
                For i = 1 To UBound(theArray)
                    If getWidth(TexthDC, theArray(i)) > MaxWidth Then MaxWidth = getWidth(TexthDC, theArray(i))
                Next

                ' calculate the new position xwidth relative to DDS_ChatBubble and yheight relative to DDS_ChatBubble
                xwidth = 10 + MaxWidth    ' the first five is just air.
                yheight = 3 + (UBound(theArray) * 7)    ' the first three are just air.

                ' Compensate the yheight drift
                Y = Y - yheight

                ' top left
                Call Engine_BltFast(X + (xwidth + 4), Y - (yheight - 4), DDS_ChatBubble, TOPLEFTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                ' top center
                For x3 = X - (xwidth - 8) To X + (xwidth)
                    Call Engine_BltFast(x3, Y - (yheight - 4), DDS_ChatBubble, TOPCENTERrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Next x3

                ' top right
                Call Engine_BltFast(X - (xwidth - 4), Y - (yheight - 4), DDS_ChatBubble, TOPRIGHTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                ' middle left
                For y3 = Y - (yheight - 8) To Y + (yheight)
                    Call Engine_BltFast(X + (xwidth + 4), y3, DDS_ChatBubble, MIDDLELEFTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Next y3

                ' middle center
                For y3 = Y - (yheight - 8) To Y + (yheight)
                    For x3 = X - (xwidth - 8) To X + (xwidth)
                        Call Engine_BltFast(x3, y3, DDS_ChatBubble, MIDDLECENTERrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Next x3
                Next y3

                ' middle right
                For y3 = Y - (yheight - 8) To Y + (yheight)
                    Call Engine_BltFast(X - (xwidth - 4), y3, DDS_ChatBubble, MIDDLERIGHTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Next y3

                ' bottom left
                Call Engine_BltFast(X + (xwidth + 4), Y + (yheight + 4), DDS_ChatBubble, BOTTOMLEFTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                ' bottom center
                For x3 = X - (xwidth - 8) To X + (xwidth)
                    Call Engine_BltFast(x3, Y + (yheight + 4), DDS_ChatBubble, BOTTOMCENTERrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Next x3

                ' bottom right
                ' RenderTexture Tex_GUI(37), xwidth + MaxWidth, yheight, 119, 6, 9, (UBound(theArray) * 12), 9, 1
                Call Engine_BltFast(X - (xwidth - 4), Y + (yheight + 4), DDS_ChatBubble, BOTTOMRIGHTrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                ' little pointy bit
                Call Engine_BltFast(X, Y + (yheight + 8), DDS_ChatBubble, TIPrect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                ' Lock the backbuffer so we can draw text and names
                TexthDC = DDS_BackBuffer.GetDC

                ' render each line centralised
                Y = Y - (yheight - 5)
                For i = 1 To UBound(theArray)
                    DrawTextNoShadow TexthDC, X - (getWidth(TexthDC, theArray(i)) - 5), Y + 3, theArray(i), QBColor(Black)    ' .colour
                    Y = Y + 12
                Next

                ' Release DC
                DDS_BackBuffer.ReleaseDC TexthDC

            End If
        End If

        ' check if it's timed out - close it if so
        If .timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With
End Sub

Function GetPokemonMaxVital(ByVal NpcNum As Long, ByVal Level As Byte) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetPokemonMaxVital = 0
        Exit Function
    End If

    If Npc(NpcNum).Pokemon > 0 Then
        GetPokemonMaxVital = Pokemon(Npc(NpcNum).Pokemon).Vital(Vitals.HP) + (Level * 5)
    Else
        GetPokemonMaxVital = Npc(NpcNum).HP
    End If

End Function

Public Sub bltTargetHp(ByVal MapNpcNum As Integer, ByVal TargetHp As Long)
    Dim Resultado As Long

    If Npc(MapNpc(MapNpcNum).num).Pokemon > 0 Then
        Resultado = (TargetHp / 122) / (GetPokemonMaxVital(MapNpc(MapNpcNum).num, MapNpc(MapNpcNum).Level) / 122) * 122
    Else
        Resultado = (TargetHp / 122) / (Npc(MapNpc(MapNpcNum).num).HP / 122) * 122
    End If

    frmMain.ImgHpTarget.Width = Resultado
End Sub

Public Sub bltPokemonTarget(ByVal MapNpcNum As Integer)
    Dim rec As RECT, rec_pos As RECT, faceNum As Long
    Dim i As Long, NpcNum As Integer
    Dim X As Long, Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    frmMain.PicTarget.Cls

    If NumFaces = 0 Then Exit Sub
    NpcNum = MapNpc(MapNpcNum).num
    If NpcNum = 0 Then Exit Sub

    faceNum = Npc(NpcNum).Pokemon

    If faceNum = 0 Then Exit Sub

    With rec
        .top = 0
        .Bottom = 32
        .Left = 32
        .Right = 64
    End With

    With rec_pos
        .top = 18
        .Bottom = .top + PIC_Y
        .Left = 8
        .Right = .Left + PIC_X
    End With

    If faceNum > NumPokeIcons Then Exit Sub

    If MapNpc(MapNpcNum).Shiny = True Then
        ' Load face if not loaded, and reset timer
        PokeIconShinyTimer(faceNum) = GetTickCount + SurfaceTimerMax

        If DDS_PokeIconShiny(faceNum) Is Nothing Then
            Call InitDDSurf("PokeIcon\Shiny\" & faceNum, DDSD_PokeIconShiny(faceNum), DDS_PokeIconShiny(faceNum))
        End If

        Engine_BltToDC DDS_PokeIconShiny(faceNum), rec, rec_pos, frmMain.PicTarget, False
    Else
        ' Load face if not loaded, and reset timer
        PokeIconTimer(faceNum) = GetTickCount + SurfaceTimerMax

        If DDS_PokeIcons(faceNum) Is Nothing Then
            Call InitDDSurf("PokeIcon\" & faceNum, DDSD_PokeIcons(faceNum), DDS_PokeIcons(faceNum))
        End If

        Engine_BltToDC DDS_PokeIcons(faceNum), rec, rec_pos, frmMain.PicTarget, False
    End If

    Y = rec_pos.top + 18
    X = rec_pos.Left + 1
    DrawText frmMain.PicTarget.hDC, X, Y, MapNpc(MapNpcNum).Level, QBColor(BrightGreen)

    frmMain.PicTarget.Refresh
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPokeEquip", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
