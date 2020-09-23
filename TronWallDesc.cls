VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TronWallDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WallCnt As Long, WallUsed As Long
Private pWallDesc() As TronWall

Public Sub Initialize(ByVal NrStaticWalls As Long)
    WallCnt = NrStaticWalls
    WallUsed = 0
    ReDim pWallDesc(WallCnt - 1)
End Sub

Public Function AddWall(From As Long) As Long
    Dim i As Long, Index As Long

    Index = -1
    For i = 0 To WallUsed - 1
        If Not pWallDesc(i).Used Then
            Index = i
            Exit For
        End If
    Next i
    If Index = -1 Then
        If WallUsed = WallCnt Then
            'Create space for new walls.
            WallCnt = WallCnt * 2
            ReDim Preserve pWallDesc(WallCnt - 1)
        End If
        Index = WallUsed
        WallUsed = WallUsed + 1
    End If
    pWallDesc(Index).Used = True
    pWallDesc(Index).From = From
    AddWall = Index
End Function

Public Sub UpdateWall(Pos1 As D3DVECTOR2, Pos2 As D3DVECTOR2, ByVal WallIndex As Long)
    'Update wall desc
    pWallDesc(WallIndex).Pos1 = Pos1
    pWallDesc(WallIndex).Pos2 = Pos2
End Sub

Public Sub DeleteWall(ByVal WallIndex As Long)
    pWallDesc(WallIndex).Used = False
End Sub

Public Function CanPlace(Pos1 As D3DVECTOR2, Pos2 As D3DVECTOR2, ByVal OwnWallIndex As Long, ByVal xWall As Long, ByRef pWIndex As Long) As Boolean
    Dim i As Long
    Dim xW As Boolean, xD As Boolean, Hit As Boolean
    Dim NVec As D3DVECTOR2

    If Pos1.y = Pos2.y Then
        xD = True
    Else
        xD = False
    End If
    CanPlace = False
    For i = 0 To WallUsed - 1
        If pWallDesc(i).Used And Not i = OwnWallIndex Then
            If pWallDesc(i).Pos1.y = pWallDesc(i).Pos2.y Then
                xW = True
            Else
                xW = False
            End If
            'Check if the wall itself is crossed.
            If Not xD = xW Or Not xWall = 0 Then
                'This wall is normal to the check direction, so check it.
                Hit = False
                If xW Then
                    'x-Wall, check y first.
                    If InArea(Pos1.y, Pos2.y, pWallDesc(i).Pos1.y) Then
                        If InArea(pWallDesc(i).Pos1.x, pWallDesc(i).Pos2.x, Pos1.x) Then Hit = True
                    End If
                Else
                    'y-Wall, check x first.
                    If InArea(Pos1.x, Pos2.x, pWallDesc(i).Pos1.x) Then
                        If InArea(pWallDesc(i).Pos1.y, pWallDesc(i).Pos2.y, Pos1.y) Then Hit = True
                    End If
                End If
                If Hit Then
                    'crash
                    If Not xWall = 0 Then
                        CopyMemory ByVal xWall, xW, Len(xW)     ' *xWall = xW;
                    End If
                    pWIndex = i
                    Exit Function
                End If
            End If
            'Check if the bike comes too close to the wall with its side.
            If xD = xW And xWall = 0 Then
                'This wall is parallel to the check direction, so check it.
                NVec = NormalVector2dMake(pWallDesc(i).Pos1, pWallDesc(i).Pos2)
                D3DXVec2Scale NVec, NVec, 0.5
                If xW Then
                    'x-Wall, check y first.
                    If InArea(pWallDesc(i).Pos1.y - NVec.y, pWallDesc(i).Pos1.y + NVec.y, Pos1.y) Then
                        If InArea(Pos1.x, Pos2.x, pWallDesc(i).Pos1.x) Then Exit Function
                        If InArea(Pos1.x, Pos2.x, pWallDesc(i).Pos2.x) Then Exit Function
                    End If
                Else
                    'y-Wall, check x first.
                    If InArea(pWallDesc(i).Pos1.x - NVec.x, pWallDesc(i).Pos1.x + NVec.x, Pos1.x) Then
                        If InArea(Pos1.y, Pos2.y, pWallDesc(i).Pos1.y) Then Exit Function
                        If InArea(Pos1.y, Pos2.y, pWallDesc(i).Pos2.y) Then Exit Function
                    End If
                End If
            End If
        End If
    Next i
    CanPlace = True
End Function

Public Sub Delete()
    WallUsed = 0
End Sub

Public Function GetWallUsed() As Long
    GetWallUsed = WallUsed
End Function

'Returns a pointer.
Public Function GetWallDescPtr() As Long
    GetWallDescPtr = VarPtr(pWallDesc(0))
End Function

Public Sub DeletePlayer(ByVal Player As Long)
    Dim i As Long

    For i = 0 To WallUsed - 1
        If pWallDesc(i).From = Player Then pWallDesc(i).Used = False
    Next i
End Sub

Public Sub DestroyWall(ByVal WallIndex As Long, ByVal xWall As Boolean, ByRef pPos As D3DVECTOR2)
    On Local Error GoTo Failed
    
    Dim i As Long, WIndex As Long
    Dim WallPos2 As D3DVECTOR2, WallDir As D3DVECTOR2
    Dim WallCut(1) As Single, WallCoord As Single
    Dim CutValid(1) As Boolean
    Dim pCP As TronPlayer

    If pWallDesc(WallIndex).From = -1 Then Exit Sub
    Set pCP = gPlayer(pWallDesc(WallIndex).From)
    pCP.WallCrash
    'Compute wall direction
    D3DXVec2Subtract WallDir, pWallDesc(WallIndex).Pos2, pWallDesc(WallIndex).Pos1
    D3DXVec2Normalize WallDir, WallDir
    'Re-create wall description
    If xWall Then
        WallCut(0) = pPos.x - 5 * WallDir.x
        WallCut(1) = pPos.x + 5 * WallDir.x
        WallCoord = pWallDesc(WallIndex).Pos1.y
        For i = 0 To 1
            If InArea(pWallDesc(WallIndex).Pos1.x, pWallDesc(WallIndex).Pos2.x, WallCut(i)) Then
                CutValid(i) = True
            Else
                CutValid(i) = False
            End If
        Next i
        If CutValid(0) And CutValid(1) Then
            'General case: wall is split in the middle and 2 new walls are created.
            WallPos2 = pWallDesc(WallIndex).Pos2
            UpdateWall pWallDesc(WallIndex).Pos1, Vector2dMake(WallCut(0), WallCoord), WallIndex
            WIndex = AddWall(pWallDesc(WallIndex).From)
            UpdateWall Vector2dMake(WallCut(1), WallCoord), WallPos2, WIndex
        ElseIf Not CutValid(0) And CutValid(1) Then
            UpdateWall Vector2dMake(WallCut(1), WallCoord), pWallDesc(WallIndex).Pos2, WallIndex
        ElseIf CutValid(0) And Not CutValid(1) Then
            UpdateWall pWallDesc(WallIndex).Pos1, Vector2dMake(WallCut(0), WallCoord), WallIndex
        Else
            DeleteWall WallIndex
        End If
    Else
        WallCut(0) = pPos.y - 5 * WallDir.y
        WallCut(1) = pPos.y + 5 * WallDir.y
        WallCoord = pWallDesc(WallIndex).Pos1.x
        For i = 0 To 1
            If InArea(pWallDesc(WallIndex).Pos1.y, pWallDesc(WallIndex).Pos2.y, WallCut(i)) Then
                CutValid(i) = True
            Else
                CutValid(i) = False
            End If
        Next i
        If CutValid(0) And CutValid(1) Then
            'General case: wall is split in the middle and 2 new walls are created.
            WallPos2 = pWallDesc(WallIndex).Pos2
            UpdateWall pWallDesc(WallIndex).Pos1, Vector2dMake(WallCoord, WallCut(0)), WallIndex
            WIndex = AddWall(pWallDesc(WallIndex).From)
            UpdateWall Vector2dMake(WallCoord, WallCut(1)), WallPos2, WIndex
        ElseIf Not CutValid(0) And CutValid(1) Then
            UpdateWall Vector2dMake(WallCoord, WallCut(1)), pWallDesc(WallIndex).Pos2, WallIndex
        ElseIf CutValid(0) And Not CutValid(1) Then
            UpdateWall pWallDesc(WallIndex).Pos1, Vector2dMake(WallCoord, WallCut(0)), WallIndex
        Else
            DeleteWall WallIndex
        End If
    End If
    'Optical re-creation of the walls.
    pCP.WallProp.Rebuild Me, pWallDesc(WallIndex).From
    Set pCP = Nothing
Failed:
End Sub

'Determines the shortest distance from this point in this direction to a wall in the arena.
Public Function GetWP(pPos As D3DVECTOR2, pDir As D3DVECTOR2) As Single
    Dim i As Long
    Dim ActDist As Single, MinDist As Single
    Dim LPos As D3DVECTOR2
    Dim First As Boolean

    First = True
    LPos = pPos
    If LPos.x < 0 Then LPos.x = 0
    If LPos.x > gOptions.LandSize Then LPos.x = gOptions.LandSize
    If LPos.y < 0 Then LPos.y = 0
    If LPos.y > gOptions.LandSize Then LPos.y = gOptions.LandSize
    
    'Check walls of the bikes.
    For i = 0 To WallUsed - 1
        If pWallDesc(i).Used Then
            If Not pDir.y = 0 Then
                If (pWallDesc(i).Pos1.y - LPos.y) * pDir.y > 0 And InAreaDelta(pWallDesc(i).Pos1.x, pWallDesc(i).Pos2.x, LPos.x, 2.5 - gOptions.OpponentSkill) Then
                    ActDist = Abs(pWallDesc(i).Pos1.y - LPos.y)
                    If First Or ActDist < MinDist Then
                        First = False
                        MinDist = ActDist
                    End If
                End If
            ElseIf Not pDir.x = 0 Then
                If (pWallDesc(i).Pos1.x - LPos.x) * pDir.x > 0 And InAreaDelta(pWallDesc(i).Pos1.y, pWallDesc(i).Pos2.y, LPos.y, 2.5 - gOptions.OpponentSkill) Then
                    ActDist = Abs(pWallDesc(i).Pos1.x - LPos.x)
                    If First Or ActDist < MinDist Then
                        First = False
                        MinDist = ActDist
                    End If
                End If
            End If
        End If
    Next i
    'Check walls of the arena box.
    If pDir.x = 1 And (First Or MinDist > gOptions.LandSize - LPos.x) Then
        MinDist = gOptions.LandSize - LPos.x
        First = False
    End If
    If pDir.x = -1 And (First Or MinDist > LPos.x) Then
        MinDist = LPos.x
        First = False
    End If
    If pDir.y = 1 And (First Or MinDist > gOptions.LandSize - LPos.y) Then
        MinDist = gOptions.LandSize - LPos.y
        First = False
    End If
    If pDir.y = -1 And (First Or MinDist > LPos.y) Then MinDist = LPos.y
    GetWP = MinDist
End Function
