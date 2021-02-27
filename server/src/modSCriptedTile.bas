Attribute VB_Name = "modSCriptedTile"
Option Explicit

Sub ScriptedTile(ByVal Index As Long, ByVal Script As Long)

Select Case Script
        Case 1
            If Player(Index).InSurf = 0 Then
                Player(Index).InSurf = 3
                SendSurfInit Index, Index
                TempPlayer(Index).SurfSlideTo = GetPlayerDir(Index)
                BlockPlayer Index
            Else
                Player(Index).InSurf = 0
                SendSurfInit Index
                BlockPlayer Index
        End If
    
End Select

End Sub
