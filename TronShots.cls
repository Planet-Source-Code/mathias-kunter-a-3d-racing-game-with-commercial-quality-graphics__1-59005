VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TronShots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Shot As cls3dObject
Private ShotLight As cls3dPolygons
Private LightMap As cls2dTexture
Private pShotDesc() As TronShotDesc
Private ShotCnt As Long

Public Sub Initialize(ByVal NrStaticShots As Long)
    Dim x As Long, y As Long
    Dim TransErg As Single
    Dim vData(3) As Vertex
    Dim UseMat As D3DMATERIAL8

    'Create objects.
    Set Shot = New cls3dObject
    Set ShotLight = New cls3dPolygons
    Set LightMap = New cls2dTexture
    
    'Load shot object.
    Shot.LoadFromFile "objects\shot.x"
    'Load texture.
    If Not LightMap.LoadForEdit(64, 64) Then Exit Sub
    If LightMap.EditStart Then
        For x = 0 To 63
            For y = 0 To 63
                TransErg = 255 - 8 * Sqr((32 - x) * (32 - x) + (32 - y) * (32 - y))
                If TransErg > 0 Then
                    LightMap.Edit Vector2dMake(x, y), TransErg, 255, 255, 255
                Else
                    LightMap.Edit Vector2dMake(x, y), 0, 255, 255, 255
                End If
            Next y
        Next x
        LightMap.EditEnd
    End If
    'Create light source.
    ShotLight.Initialize False
    vData(0) = VertexMake(-5, 0, 5, 0, 1, 0, 0, 0)
    vData(1) = VertexMake(5, 0, 5, 0, 1, 0, 1, 0)
    vData(2) = VertexMake(-5, 0, -5, 0, 1, 0, 0, 1)
    vData(3) = VertexMake(5, 0, -5, 0, 1, 0, 1, 1)
    ShotLight.SetVertexData VarPtr(vData(0)), 0, 4
    ShotLight.SetTexture LightMap
    UseMat = MaterialMake(255, 255, 255, 0, 0, 0, 0, 100)
    ShotLight.SetMaterial UseMat
    ShotLight.SetPolyFormat TRIANGLESTRIP
    ShotLight.SetVisibility BOTHSIDES

    ShotCnt = NrStaticShots
    ReDim pShotDesc(ShotCnt - 1)
    Delete
End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

Public Sub Cleanup()
    Set Shot = Nothing
    Set ShotLight = Nothing
    Set LightMap = Nothing
End Sub

Public Sub Delete()
    Dim i As Long

    For i = 0 To ShotCnt - 1
        pShotDesc(i).Used = False
    Next i
End Sub

Public Sub FireShot(Pos As D3DVECTOR2, Dir As D3DVECTOR2, ByVal ShotType As Long, ByVal pFrom As Long)
    Dim i As Long, Index As Long
    Dim SPos As D3DVECTOR2, SDir As D3DVECTOR2

    Index = -1
    For i = 0 To ShotCnt - 1
        If Not pShotDesc(i).Used Then
            Index = i
            Exit For
        End If
    Next i
    If Index = -1 Then
        'Create new space.
        Index = ShotCnt
        ShotCnt = ShotCnt * 2
        ReDim Preserve pShotDesc(ShotCnt - 1)
    End If
    'Add the new shot into the local array.
    SPos.x = Pos.x + Dir.x * 2
    SPos.y = Pos.y + Dir.y * 2
    If SPos.x < 0 Then
        SPos.x = 0
    ElseIf SPos.x > gOptions.LandSize Then
        SPos.x = gOptions.LandSize
    End If
    If SPos.y < 0 Then
        SPos.y = 0
    ElseIf SPos.y > gOptions.LandSize Then
        SPos.y = gOptions.LandSize
    End If
    SDir.x = Dir.x + GetRandomVal(-0.1, 0.1)
    SDir.y = Dir.y + GetRandomVal(-0.1, 0.1)
    D3DXVec2Normalize SDir, SDir
    D3DXVec2Scale SDir, SDir, gOptions.MopedSpeed * 2
    With pShotDesc(Index)
        .Pos = SPos
        .Dir = SDir
        .NoDanger = 1
        .ShotType = ShotType
        If ShotType = 0 Then
            .TTL = 25                       'Default shot
        ElseIf ShotType = 1 Then
            .TTL = 20                       'Target-finding shot
        Else
            .TTL = 40                       'Super shot
        End If
        .pFrom = pFrom
        .Used = True
    End With
End Sub

Public Function GetDest(ByVal Index As Long) As D3DVECTOR2
    Dim i As Long, MinIndex As Long
    Dim ActDist As Single, MinDist As Single
    Dim First As Boolean
    Dim ActVec As D3DVECTOR2

    MinIndex = -1
    First = True
    'Get a target for the target-finding shot.
    For i = 0 To gOptions.OpponentCnt
        If gPlayer(i).Alive And Not pShotDesc(Index).pFrom = i Then
            'This bike could be targeted.
            D3DXVec2Subtract ActVec, gPlayer(i).PosProp, pShotDesc(Index).Pos
            ActDist = D3DXVec2Length(ActVec)
            If First Or ActDist < MinDist Then
                First = False
                MinDist = ActDist
                MinIndex = i
            End If
        End If
    Next i
    If Not MinIndex = -1 Then
        ActVec = gPlayer(MinIndex).PosProp
    Else
        'No target could be found. Use a random target.
        ActVec = Vector2dMake(GetRandomVal(0, gOptions.LandSize), GetRandomVal(0, gOptions.LandSize))
    End If
    GetDest = ActVec
End Function

Public Sub DoEventsX(ByVal RelTimeDiff As Single)
    Dim i As Long, j As Long, WIndex As Long
    Dim bx As Boolean, by As Boolean, xBounce As Boolean, wBounce As Boolean
    Dim PrevPos As D3DVECTOR2, ActDir As D3DVECTOR2, PrevDir As D3DVECTOR2

    For i = 0 To ShotCnt - 1
        With pShotDesc(i)
            If .Used Then
                'This shot is used and active now.
                If Not .ShotType = 0 Then
                    'Get the target.
                    D3DXVec2Subtract ActDir, GetDest(i), .Pos
                    D3DXVec2Normalize ActDir, ActDir
                    D3DXVec2Normalize PrevDir, .Dir
                    If .ShotType = 1 Then
                        D3DXVec2Lerp ActDir, PrevDir, ActDir, RelTimeDiff * 3                   '= delay of 1/3 second
                        D3DXVec2Scale .Dir, ActDir, gOptions.MopedSpeed * 1.1
                    Else
                        D3DXVec2Lerp ActDir, PrevDir, ActDir, RelTimeDiff * 5                   '= delay of 1/5 second
                        D3DXVec2Scale .Dir, ActDir, gOptions.MopedSpeed * 2
                    End If
                End If
                PrevPos = .Pos
                .Pos.x = .Pos.x + .Dir.x * RelTimeDiff
                .Pos.y = .Pos.y + .Dir.y * RelTimeDiff
                .NoDanger = .NoDanger - RelTimeDiff
                If .NoDanger < 0 Then .NoDanger = 0
                bx = False
                by = False
                wBounce = False
    
                'Check the walls of the bikes.
                If Not gWallDesc.CanPlace(PrevPos, .Pos, -1, VarPtr(xBounce), WIndex) Then
                    If xBounce Then
                        bx = True
                    Else
                        by = True
                    End If
                'Check the walls of the arena.
                ElseIf .Pos.x < 0 Or .Pos.x > gOptions.LandSize Then
                    by = True
                    wBounce = True
                ElseIf .Pos.y < 0 Or .Pos.y > gOptions.LandSize Then
                    bx = True
                    wBounce = True
                End If
                If bx Or by Then
                    If wBounce Then
                        'Bounce from arena wall. This do all shots.
                        If by Then
                            .Dir.x = .Dir.x * -1
                        Else
                            .Dir.y = .Dir.y * -1
                        End If
                        .Pos = PrevPos
                        .Pos.x = .Pos.x + .Dir.x * RelTimeDiff
                        .Pos.y = .Pos.y + .Dir.y * RelTimeDiff
                        'Decrease TTL, but don't explode on an arena wall.
                        .TTL = .TTL - 1
                        If .TTL = 0 Then .TTL = 1
                        'I hate that VB doesn't implement any continue command like in C++.
                        'So, use a label instead of messing around with complicated if's.
                        GoTo NextLoop
                    End If
                    If Not .ShotType = 0 Or .TTL = 1 Then
                        gWallDesc.DestroyWall WIndex, bx, .Pos
                        gExplosions.CreateExplosion .Pos, True
                    Else
                        'Normal shot: bounce from a bike wall.
                        If by Then
                            .Dir.x = .Dir.x * -1
                        Else
                            .Dir.y = .Dir.y * -1
                        End If
                        .Pos = PrevPos
                        .Pos.x = .Pos.x + .Dir.x * RelTimeDiff
                        .Pos.y = .Pos.y + .Dir.y * RelTimeDiff
                    End If
                    .TTL = .TTL - 1
                    If .TTL = 0 Then
                        'Delete this shot.
                        .Used = False
                    End If
                End If
                
                'Check if the shot hits a bike.
                For j = 0 To gOptions.OpponentCnt
                    If gPlayer(j).Alive Then
                        'VB, clever as it is, processes the whole if-instruction when using logical AND's...
                        'So, I'm using 2 ifs to increase performance.
                        If Abs(gPlayer(j).PosProp.x - .Pos.x) < 2 And Abs(gPlayer(j).PosProp.y - .Pos.y) < 2 Then
                            If Not (.NoDanger > 0 And .pFrom = j) Then
                                If gPlayer(j).ShieldTime > 0 Then
                                    'This bike is using a shield, bounce from it.
                                    .Dir.x = .Dir.x * -1
                                    .Dir.y = .Dir.y * -1
                                    .Pos = PrevPos
                                    .Pos.x = .Pos.x + .Dir.x * RelTimeDiff
                                    .Pos.y = .Pos.y + .Dir.y * RelTimeDiff
                                Else
                                    'Crash this bike.
                                    gPlayer(j).WillCrash = True
                                End If
                            End If
                        End If
                    End If
                Next j
            End If
        End With
NextLoop:
    Next i
End Sub

Public Sub Render(ByVal RenderLight As Boolean)
    Dim i As Long

    If RenderLight Then
        'Render the lights of the shots.
        gEngine.SetZState False, False
        For i = 0 To ShotCnt - 1
            If pShotDesc(i).Used Then
                ShotLight.Position = Vector3dMake(pShotDesc(i).Pos.x, 0, pShotDesc(i).Pos.y)
                If pShotDesc(i).ShotType = 0 Then
                    gEngine.SetGlobalLight ColorMake(0, 255, 0)
                ElseIf pShotDesc(i).ShotType = 1 Then
                    gEngine.SetGlobalLight ColorMake(255, 0, 0)
                Else
                    gEngine.SetGlobalLight ColorMake(0, 0, 0)
                End If
                ShotLight.Render 0, 2
            End If
        Next i
        gEngine.SetZState True, True
    End If
    'Render the shots itself.
    For i = 0 To ShotCnt - 1
        If pShotDesc(i).Used Then
            Shot.Position = Vector3dMake(pShotDesc(i).Pos.x, 0, pShotDesc(i).Pos.y)
            If pShotDesc(i).ShotType = 0 Then
                gEngine.SetGlobalLight ColorMake(0, 255, 0)
            ElseIf pShotDesc(i).ShotType = 1 Then
                gEngine.SetGlobalLight ColorMake(255, 0, 0)
            Else
                gEngine.SetGlobalLight ColorMake(0, 0, 0)
            End If
            Shot.Render
        End If
    Next i
End Sub
