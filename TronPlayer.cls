VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TronPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Moped As cls3dObject, Wheel As cls3dObject, SWheel As cls3dObject
Private Wall As TronWalls
Private Pos As D3DVECTOR2, PrevPos As D3DVECTOR2, LastWallPos As D3DVECTOR2
Private ActWallPos As D3DVECTOR2, LastOWPos As D3DVECTOR2, Dir As D3DVECTOR2
Private BackRPos As D3DVECTOR, FrontRPos As D3DVECTOR
Private YAngle As Single, RotationLeft As Single, ItemTime As Single
Private SpeedTime As Single
Private AIReact As Single, AILastCh As Single, AILastHook As Single
Private WallDescIndex As Long
Public Alive As Boolean, ShieldTime As Single
Public Frags As Long, ItemAviable As Long, Rank As Long
Public WillCrash As Boolean, Name As String, Ident As Long


Public Property Get LastWallPosProp() As D3DVECTOR2
    LastWallPosProp = LastWallPos
End Property

Public Property Get PosProp() As D3DVECTOR2
    PosProp = Pos
End Property

Public Property Get DirProp() As D3DVECTOR2
    DirProp = Dir
End Property

Public Property Get WallProp() As TronWalls
    Set WallProp = Wall
End Property




Private Sub Class_Terminate()
    Set Moped = Nothing
    Set Wheel = Nothing
    Set SWheel = Nothing
    Set Wall = Nothing
End Sub


Public Function Initialize(ByVal MopedNr As Long, ByVal nIdent As Long, pWallTex As cls2dTexture) As Boolean
    Dim DirVal As Long

    'Create objects.
    Set Moped = New cls3dObject
    Set Wheel = New cls3dObject
    Set SWheel = New cls3dObject
    Set Wall = New TronWalls
    

    'Load 3d data
    If Not Moped.LoadFromFile("objects\moped" & MopedNr & ".x") Then Exit Function
    If Not Wheel.LoadFromFile("objects\wheel.x") Then Exit Function
    If Not SWheel.LoadFromFile("objects\swheel.x") Then Exit Function
    Moped.SetSpecPower 10
    Wheel.SetSpecPower 10
    DirVal = CInt(GetRandomVal(0, 4))
    Dir = Vector2dMake(0, 0)
    If DirVal = 0 Then
        Dir.x = 1
        YAngle = PI * 1.5
    ElseIf DirVal = 1 Then
        Dir.x = -1
        YAngle = PI / 2
    ElseIf DirVal = 2 Then
        Dir.y = 1
        YAngle = PI
    Else
        Dir.y = -1
        YAngle = 0
    End If
    RotationLeft = 0
    Moped.Rotate Vector3dMake(0, YAngle, 0), Vector3dMake(0, 1, 2), Vector3dMake(0, 0, 0)
    Wheel.Rotate Vector3dMake(0, YAngle, 0), Vector3dMake(0, 1, 2), Vector3dMake(0, 0, 0)
    'Walls of the bike.
    If MopedNr = 0 Then
        Wall.Initialize 200, 0, 0, 200, pWallTex
    ElseIf MopedNr = 1 Then
        Wall.Initialize 200, 200, 0, 0, pWallTex
    ElseIf MopedNr = 2 Then
        Wall.Initialize 200, 50, 50, 50, pWallTex
    Else
        Wall.Initialize 200, 200, 200, 0, pWallTex
    End If
    'Name
    Ident = nIdent
    If Ident = 0 Then
        Name = Trim$(gOptions.PlayerName)
    ElseIf Ident = 1 Then
        Name = "Gates"
    ElseIf Ident = 2 Then
        Name = "Elvis"
    ElseIf Ident = 3 Then
        Name = "Bush"
    ElseIf Ident = 4 Then
        Name = "Britney"
    ElseIf Ident = 5 Then
        Name = "Zidane"
    End If
    Frags = 0
    Reset
    Initialize = True
End Function

Public Sub Reset()
    'Items
    ItemAviable = -1
    ItemTime = GetRandomVal(10, 15)

    'Walls
    Wall.AddWall
    WallDescIndex = gWallDesc.AddWall(Ident)

    'General
    GeneratePos
    Wall.UpdateWall Pos, Pos
    gWallDesc.UpdateWall Pos, Pos, WallDescIndex
    Alive = True
    WillCrash = False
    SpeedTime = 0
    ShieldTime = 0
    Rank = 1
    AIReact = 0
    AILastCh = 1
    AILastHook = 1
End Sub

Public Sub GeneratePos()
    Pos.x = GetRandomVal(20, gOptions.LandSize - 20)
    Pos.y = GetRandomVal(20, gOptions.LandSize - 20)
    PrevPos = Pos
    ActWallPos.x = Pos.x - Dir.x * 1.6
    ActWallPos.y = Pos.y - Dir.y * 1.6
    LastWallPos = ActWallPos
    LastOWPos = ActWallPos
End Sub

Public Sub Crash()
    If Not Alive Then Exit Sub
    DoEventsX 30, 0
    gWallDesc.DeletePlayer Ident
    Alive = False
End Sub

Public Sub Delete()
    Wall.Delete
End Sub

Public Function GetChangeDir(ByVal TurnDir As Long) As D3DVECTOR2
    If TurnDir = 0 Then
        'Left
        If Not Dir.y = 0 Then
            GetChangeDir = Vector2dMake(-Dir.y, 0)
        Else
            GetChangeDir = Vector2dMake(0, Dir.x)
        End If
    Else
        'Right
        If Not Dir.y = 0 Then
            GetChangeDir = Vector2dMake(Dir.y, 0)
        Else
            GetChangeDir = Vector2dMake(0, -Dir.x)
        End If
    End If
End Function

Public Sub ChangeDir(ByVal TurnDir As Long)
    Dim NVec As D3DVECTOR

    If Not Alive Then Exit Sub
    If TurnDir = 0 Then
        'Left
        If Not Dir.y = 0 Then
            Dir.x = -Dir.y
            Dir.y = 0
        Else
            Dir.y = Dir.x
            Dir.x = 0
        End If
        RotationLeft = RotationLeft - PI / 2#
        If RotationLeft < 2 * -PI Then RotationLeft = RotationLeft + 2 * PI
    Else
        'Right
        If Not Dir.y = 0 Then
            Dir.x = Dir.y
            Dir.y = 0
        Else
            Dir.y = -Dir.x
            Dir.x = 0
        End If
        RotationLeft = RotationLeft + PI / 2
        If RotationLeft > 2 * PI Then RotationLeft = RotationLeft - 2 * PI
    End If
    'Set last wall until the bike.
    Wall.UpdateWall LastOWPos, Pos
    gWallDesc.UpdateWall LastWallPos, Pos, WallDescIndex
    'Create a new wall.
    Wall.AddWall
    WallDescIndex = gWallDesc.AddWall(Ident)
    Wall.UpdateWall Pos, Pos
    gWallDesc.UpdateWall Pos, Pos, WallDescIndex
    'Update variables
    LastWallPos = Pos
    LastOWPos = Pos
    ActWallPos.x = LastWallPos.x - Dir.x * 1.6
    ActWallPos.y = LastWallPos.y - Dir.y * 1.6
End Sub

'*********************************************
'TRON AI III
'*********************************************
Public Sub AI(ByVal RelTimeDiff As Single)
    Dim i As Long, NDir As Long, WIndex As Long
    Dim AIRTime As Single, WayLen As Single, MopedSpeed As Single
    Dim DirVec As D3DVECTOR2

    If Not Alive Or AIReact > 0 Then Exit Sub
    MopedSpeed = gOptions.MopedSpeed
    If SpeedTime > 0 Then
        MopedSpeed = MopedSpeed * 1.5
    ElseIf ShieldTime > 0 Then
        MopedSpeed = MopedSpeed * 1.25
    End If
    'Fire target-finding shot or use a shield.
    If gOptions.EnableAction And ItemAviable >= 2 Then FireItem
    'Fire default shot, if aviable.
    WayLen = gWallDesc.GetWP(Pos, Dir)
    If gOptions.EnableAction And ItemAviable = 1 And WayLen > 3 * MopedSpeed Then FireItem
    
    AIRTime = 0.25 - gOptions.OpponentSkill * 0.1
    If AIChDir(NDir, RelTimeDiff) Then
        'It's recommended to change the direction.
        If Not NDir = -1 Then
            'The recommended changing direction is already known. Check now if this change makes sense.
            DirVec = NormalVector2dMake(Vector2dMake(0, 0), Dir)            'Right
            If NDir = 0 Then D3DXVec2Scale DirVec, DirVec, -1               'Left
            If gWallDesc.GetWP(Pos, DirVec) - MopedSpeed * (AIRTime + 0.05) > 0 Then
                ChangeDir NDir
                AIReact = AIRTime
                'Ensure that this action isn't made undone by changing the direction again too early.
                AILastCh = 2
            End If
        Else
            'The direction has still to be detected. It's also possible to drive straight ahead.
            For i = 0 To gOptions.OpponentCnt
                WIndex = gWallDesc.AddWall(-1)
                gWallDesc.UpdateWall gPlayer(i).LastWallPosProp, gPlayer(i).PosProp, WIndex
            Next i
            WayLen = AIGetWayLen(Pos, Dir, gOptions.OpponentSkill + 3, NDir)
            'Use speedup, if it makes sense.
            If gOptions.EnableAction And ItemAviable = 0 And WayLen > 5 * MopedSpeed Then FireItem
            gWallDesc.DeletePlayer -1
            If Not NDir = -1 Then
                ChangeDir NDir
                AIReact = AIRTime
            End If
        End If
    ElseIf AILastCh = 0 Then
        'Change the direction periodly.
        For i = 0 To gOptions.OpponentCnt
            WIndex = gWallDesc.AddWall(-1)
            gWallDesc.UpdateWall gPlayer(i).LastWallPosProp, gPlayer(i).PosProp, WIndex
        Next i
        WayLen = AIGetWayLen(Pos, Dir, gOptions.OpponentSkill + 3, NDir)
        'Use speedup, if it makes sense.
        If gOptions.EnableAction And ItemAviable = 0 And WayLen > 5 * MopedSpeed Then FireItem
        gWallDesc.DeletePlayer -1
        If Not NDir = -1 Then
            ChangeDir NDir
            AIReact = AIRTime
            AILastCh = 0.5
        End If
    End If
End Sub

'*********************************************
'TRON AI III
'*********************************************
Private Function AIGetWayLen(Pos As D3DVECTOR2, LastDir As D3DVECTOR2, ByVal d As Long, ByRef pDir As Long) As Single
    Dim i As Long, WIndex As Long
    Dim WayLen(2) As Single, AIRTime As Single, MopedSpeed As Single
    Dim DirVec(2) As D3DVECTOR2, AddVec As D3DVECTOR2, NPos As D3DVECTOR2

    MopedSpeed = gOptions.MopedSpeed
    If d = 0 Then Exit Function
    DirVec(0) = LastDir                                                     'Straight ahead
    DirVec(1) = NormalVector2dMake(Vector2dMake(0, 0), LastDir)             'Right
    D3DXVec2Scale DirVec(2), DirVec(1), -1                                  'Left
    AIRTime = 0.3 - gOptions.OpponentSkill * 0.1
    If SpeedTime > 0 Then
        MopedSpeed = MopedSpeed * 1.5
    ElseIf ShieldTime > 0 Then
        MopedSpeed = MopedSpeed * 1.25
    End If
    For i = 0 To 2
        'Check the three possibilites (straight ahead, left, right) for the longest possible way.
        WayLen(i) = gWallDesc.GetWP(Pos, DirVec(i)) - MopedSpeed * AIRTime
        If WayLen(i) > 0 Then
            D3DXVec2Scale AddVec, DirVec(i), WayLen(i)
            D3DXVec2Add NPos, Pos, AddVec
            'Create a temp wall
            WIndex = gWallDesc.AddWall(-1)
            gWallDesc.UpdateWall Pos, NPos, WIndex
            'Calculate recursively from the view of the next point.
            WayLen(i) = WayLen(i) + AIGetWayLen(NPos, DirVec(i), d - 1, pDir)
            'Delete temp wall
            gWallDesc.DeleteWall WIndex
        Else
            WayLen(i) = 0
        End If
    Next i
    'Find out the longest way.
    If WayLen(0) = 0 And WayLen(1) = 0 And WayLen(2) = 0 Then
        'It looks pretty bad. No way can be choosen.
        pDir = -1
        AIGetWayLen = 0
        Exit Function
    End If
    If WayLen(0) > WayLen(1) And WayLen(0) > WayLen(2) Then
        'Straight ahead
        pDir = -1
        AIGetWayLen = WayLen(0)
        Exit Function
    ElseIf WayLen(1) > WayLen(0) And WayLen(1) > WayLen(2) Then
        'Right
        pDir = 1
        AIGetWayLen = WayLen(1)
        Exit Function
    Else
        'Left
        pDir = 0
        AIGetWayLen = WayLen(2)
        Exit Function
    End If
End Function

'*********************************************
'TRON AI III
'*********************************************
Private Function AIChDir(ByRef pDir As Long, ByVal RelTimeDiff As Single) As Boolean
    Dim i As Long
    Dim Result As Single, dx As Single, dy As Single, MopedSpeed  As Single
    Dim TestPos As D3DVECTOR2

    MopedSpeed = gOptions.MopedSpeed
    If SpeedTime > 0 Then
        MopedSpeed = MopedSpeed * 1.5
    ElseIf ShieldTime > 0 Then
        MopedSpeed = MopedSpeed * 1.25
    End If
    'Cases where changing the direction is really nescessary.
    pDir = -1
    '1. when a wall would be hit.
    TestPos.x = Pos.x + Dir.x * (2 * MopedSpeed * RelTimeDiff + 5 - gOptions.OpponentSkill * 2)
    TestPos.y = Pos.y + Dir.y * (2 * MopedSpeed * RelTimeDiff + 5 - gOptions.OpponentSkill * 2)
    If TestPos.x < 0 Or TestPos.x > gOptions.LandSize Or TestPos.y < 0 Or TestPos.y > gOptions.LandSize Then
        AIChDir = True
        Exit Function
    End If
    If Not ShieldTime = 0 Then
        AIChDir = False
        Exit Function
    End If
    If Not gWallDesc.CanPlace(PrevPos, TestPos, WallDescIndex, 0, 0) Then
        AIChDir = True
        Exit Function
    End If
    For i = 0 To gOptions.OpponentCnt
        If Not gPlayer(i).Ident = Ident And Not CanMopedPlace(TestPos, i) Then
            AIChDir = True
            Exit Function
        End If
    Next i
    '2. when another bike comes in our way.
    For i = 0 To gOptions.OpponentCnt
        If Not gPlayer(i).Ident = Ident And gPlayer(i).Alive And gPlayer(i).DirProp.x = -Dir.x And gPlayer(i).DirProp.y = -Dir.y Then
            If Not Dir.x = 0 And Abs(gPlayer(i).PosProp.x - Pos.x) < 20 And Abs(gPlayer(i).PosProp.y - Pos.y) < 5 Then
                AIChDir = True
                Exit Function
            ElseIf Not Dir.y = 0 And Abs(gPlayer(i).PosProp.x - Pos.x) < 5 And Abs(gPlayer(i).PosProp.y - Pos.y) < 20 Then
                AIChDir = True
                Exit Function
            End If
        End If
    Next i
    '3. if an enemy is behind us, make a hook to cut off its way.
    If gOptions.OpponentSkill > 0 Then
        For i = 0 To gOptions.OpponentCnt
            If Not gPlayer(i).Ident = Ident And gPlayer(i).Alive And AILastHook = 0 Then
                If (gPlayer(i).DirProp.x = Dir.x And gPlayer(i).DirProp.y = Dir.y) Or (gPlayer(i).DirProp.x = -Dir.x And gPlayer(i).DirProp.y = -Dir.y) Then
                    'This bike drives either in the same direction or comes against us.
                    '0 = left, 1 = right
                    AILastHook = 1
                    dx = Abs(Pos.x - gPlayer(i).PosProp.x)
                    dy = Abs(Pos.y - gPlayer(i).PosProp.y)
                    If Not Dir.x = 0 And dx < 75 And dx > dy Then
                        Result = Dir.x * (gPlayer(i).PosProp.y - Pos.y)
                        If Result > 0 Then
                            pDir = 0
                        Else
                            pDir = 1
                        End If
                        AIChDir = True
                        Exit Function
                    ElseIf Not Dir.y = 0 And dy < 75 And dy > dx Then
                        Result = Dir.y * (gPlayer(i).PosProp.x - Pos.x)
                        If Result > 0 Then
                            pDir = 1
                        Else
                            pDir = 0
                        End If
                        AIChDir = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    AIChDir = False
End Function

Public Sub DoEventsX(ByVal RelTimeDiff As Single, ByVal MopedSpeed As Single)
    Dim AddDir As D3DVECTOR2, WheelRot As D3DVECTOR2, CalcVec As D3DVECTOR2
    Dim sinAngle As Single, cosAngle As Single, WLen As Single

    If Not Alive Then
        Wall.TransIncr RelTimeDiff
        Exit Sub
    End If
    
    '*************Items********************
    If gOptions.EnableAction Then
        SpeedTime = SpeedTime - RelTimeDiff
        If SpeedTime > 0 Then
            MopedSpeed = MopedSpeed * 1.5
        Else
            SpeedTime = 0
        End If
        ShieldTime = ShieldTime - RelTimeDiff
        If ShieldTime > 0 Then
            MopedSpeed = MopedSpeed * 1.25
        Else
            ShieldTime = 0
        End If
        If ItemAviable = -1 Then
            ItemTime = ItemTime - RelTimeDiff
            If ItemTime < 0 Then
                ItemTime = GetRandomVal(10, 15)
                ItemAviable = GetRandomVal(0, 100) Mod 100
                If ItemAviable < 25 Then
                    ItemAviable = 0             'Speedup: 25 %
                ElseIf ItemAviable < 50 Then
                    ItemAviable = 1             'Normal shot: 25 %
                ElseIf ItemAviable < 80 Then
                    ItemAviable = 2             'Target-finding shot: 30 %
                ElseIf ItemAviable < 82 Then
                    ItemAviable = 3             'Super shot: 2 %
                Else
                    ItemAviable = 4             'Shield: 18 %
                End If
            End If
        End If
    End If
    
    '*************Move the bike**************
    AddDir.x = Dir.x * MopedSpeed * RelTimeDiff
    AddDir.y = Dir.y * MopedSpeed * RelTimeDiff
    PrevPos = Pos
    D3DXVec2Add Pos, Pos, AddDir
    D3DXVec2Add ActWallPos, ActWallPos, AddDir
    Moped.Position = Vector3dMake(Pos.x, 0, Pos.y)
    
    '*******************AI***************
    If MopedSpeed > 3 Then
        'Game already runs
        AIReact = AIReact - RelTimeDiff
        If AIReact < 0 Then AIReact = 0
        AILastCh = AILastCh - RelTimeDiff
        If AILastCh < 0 Then AILastCh = 0
        AILastHook = AILastHook - RelTimeDiff
        If AILastHook < 0 Then AILastHook = 0
    End If
    '*************Rotate bike, if nescessary**************
    If RotationLeft > 0 Then
        'Right-rotation
        Moped.Rotate Vector3dMake(0, RelTimeDiff * 10, 0), Vector3dMake(0, 1, 2), Vector3dMake(Pos.x, 0, Pos.y)
        Wheel.Rotate Vector3dMake(0, RelTimeDiff * 10, 0), Vector3dMake(0, 1, 2), Wheel.Position
        RotationLeft = RotationLeft - RelTimeDiff * 10
        YAngle = YAngle + RelTimeDiff * 10
        If RotationLeft < 0 Then
            Moped.Rotate Vector3dMake(0, RotationLeft, 0), Vector3dMake(0, 1, 2), Vector3dMake(Pos.x, 0, Pos.y)
            Wheel.Rotate Vector3dMake(0, RotationLeft, 0), Vector3dMake(0, 1, 2), Wheel.Position
            YAngle = YAngle + RotationLeft
            RotationLeft = 0
        End If
    ElseIf RotationLeft < 0 Then
        'Left-rotation
        Moped.Rotate Vector3dMake(0, -RelTimeDiff * 10, 0), Vector3dMake(0, 1, 2), Vector3dMake(Pos.x, 0, Pos.y)
        Wheel.Rotate Vector3dMake(0, -RelTimeDiff * 10, 0), Vector3dMake(0, 1, 2), Wheel.Position
        RotationLeft = RotationLeft + RelTimeDiff * 10
        YAngle = YAngle - RelTimeDiff * 10
        If RotationLeft > 0 Then
            Moped.Rotate Vector3dMake(0, RotationLeft, 0), Vector3dMake(0, 1, 2), Vector3dMake(Pos.x, 0, Pos.y)
            Wheel.Rotate Vector3dMake(0, RotationLeft, 0), Vector3dMake(0, 1, 2), Wheel.Position
            YAngle = YAngle + RotationLeft
            RotationLeft = 0
        End If
    Else
        '*************Rotate the wheels**************
        WheelRot = Vector2dMake(0, 0)
        If Dir.x = 1 Then
            WheelRot.y = -MopedSpeed * RelTimeDiff * 2
        ElseIf Dir.x = -1 Then
            WheelRot.y = MopedSpeed * RelTimeDiff * 2
        ElseIf Dir.y = 1 Then
            WheelRot.x = MopedSpeed * RelTimeDiff * 2
        Else
            WheelRot.x = -MopedSpeed * RelTimeDiff * 2
        End If
        Wheel.Rotate Vector3dMake(WheelRot.x, 0, WheelRot.y), Vector3dMake(0, 1, 2), Wheel.Position
    End If
    '*************Compute wheel positions**************
    If Not RotationLeft = 0 Then
        sinAngle = Sin(YAngle)
        cosAngle = Cos(YAngle)
        'Back wheel
        BackRPos = Vector3dMake(Pos.x + 1.35 * sinAngle, 0.45, Pos.y + 1.35 * cosAngle)
        'Front wheel
        FrontRPos = Vector3dMake(Pos.x - 1.21 * sinAngle, 0.5, Pos.y - 1.21 * cosAngle)
    Else
        'Back wheel
        BackRPos = Vector3dMake(Pos.x - Dir.x * 1.35, 0.45, Pos.y - Dir.y * 1.35)
        'Front wheel
        FrontRPos = Vector3dMake(Pos.x + Dir.x * 1.21, 0.5, Pos.y + Dir.y * 1.21)
    End If
    'Update walls
    If ActWallPos.x * Dir.x > LastWallPos.x * Dir.x Or ActWallPos.y * Dir.y > LastWallPos.y * Dir.y Then
        D3DXVec2Subtract CalcVec, LastOWPos, ActWallPos
        WLen = D3DXVec2Length(CalcVec)
        If WLen > 15 Then
            'Optical split of the wall if the len exceeds 15. Done because of 3d lightning reasons.
            D3DXVec2Scale CalcVec, Dir, 15
            D3DXVec2Add CalcVec, LastOWPos, CalcVec
            Wall.UpdateWall LastOWPos, CalcVec
            Wall.AddWall
            LastOWPos = CalcVec
        End If
        Wall.UpdateWall LastOWPos, ActWallPos
        gWallDesc.UpdateWall LastWallPos, ActWallPos, WallDescIndex
    Else
        Wall.UpdateWall LastOWPos, LastOWPos
        gWallDesc.UpdateWall LastWallPos, LastWallPos, WallDescIndex
    End If
End Sub

Public Sub HandleShield()
    If ShieldTime = 0 Or Not Alive Then Exit Sub

    Dim CrashPos As D3DVECTOR2
    Dim WIndex As Long, bx As Boolean

    'This bike is using a shield at the moment and is therefore able to destroy walls.
    D3DXVec2Add CrashPos, Pos, Dir
    If Not gWallDesc.CanPlace(PrevPos, CrashPos, -1, VarPtr(bx), WIndex) Then
        gWallDesc.DestroyWall WIndex, bx, Pos
        gExplosions.CreateExplosion Pos, True
    End If
End Sub

Public Function CanPlace() As Boolean
    Dim i As Long
    Dim CrashPos As D3DVECTOR2

    If Not Alive Then
        CanPlace = True
        Exit Function
    End If
    D3DXVec2Add CrashPos, Pos, Dir
    'Check for walls of other bikes.
    If ShieldTime = 0 And Not gWallDesc.CanPlace(PrevPos, CrashPos, WallDescIndex, 0, 0) Then Exit Function
    'Check for walls of the arena.
    If CrashPos.x < 0 Or CrashPos.x > gOptions.LandSize Or CrashPos.y < 0 Or CrashPos.y > gOptions.LandSize Then Exit Function
    'Check for direct colissions with other bikes.
    For i = 0 To gOptions.OpponentCnt
        If Not gPlayer(i).Ident = Ident And Not CanMopedPlace(CrashPos, i) Then Exit Function
    Next i
    CanPlace = True
End Function

Private Function CanMopedPlace(CrashPos As D3DVECTOR2, ByVal pOpp As Long) As Boolean
    If Not Alive Or gPlayer(pOpp).Alive Then
        CanMopedPlace = True
        Exit Function
    End If

    Dim MopedPos(1) As D3DVECTOR2

    With gPlayer(pOpp)
        MopedPos(0) = Vector2dMake(.PosProp.x - .DirProp.x * 1.85, .PosProp.y - .DirProp.y * 1.85)
        MopedPos(1) = Vector2dMake(.PosProp.x + .DirProp.x * 1.71, .PosProp.y + .DirProp.y * 1.71)
    End With
    If InLine(PrevPos.x, CrashPos.x, MopedPos(0).x, MopedPos(1).x) Then
        'x area is overlaps
        If InLine(PrevPos.y, CrashPos.y, MopedPos(0).y, MopedPos(1).y) Then
            'x and y area overlap, crash
            'The other bike explodes, except he's using a shield and we currently don't have one.
            If gPlayer(pOpp).ShieldTime = 0 Or ShieldTime > 0 Then gPlayer(pOpp).WillCrash = True
            'We explode, except we've got a shield and he don't.
            If gPlayer(pOpp).ShieldTime > 0 Or ShieldTime = 0 Then Exit Function
        End If
    End If
    CanMopedPlace = True
End Function

Public Sub WallCrash()
    'Move last wall until the bike.
    Wall.UpdateWall LastOWPos, Pos
    gWallDesc.UpdateWall LastWallPos, Pos, WallDescIndex
    'Generate a new wall desc
    WallDescIndex = gWallDesc.AddWall(Ident)
    gWallDesc.UpdateWall Pos, Pos, WallDescIndex
    LastWallPos = Pos
    LastOWPos = Pos
    ActWallPos.x = LastWallPos.x - Dir.x * 1.6
    ActWallPos.y = LastWallPos.y - Dir.y * 1.6
End Sub

Public Sub FireItem()
    If ItemAviable = 0 Then
        SpeedTime = 5
    ElseIf ItemAviable = 1 Then
        gShots.FireShot Pos, Dir, 0, Ident
    ElseIf ItemAviable = 2 Then
        gShots.FireShot Pos, Dir, 1, Ident
    ElseIf ItemAviable = 3 Then
        gShots.FireShot Pos, Dir, 2, Ident
    ElseIf ItemAviable = 4 Then
        ShieldTime = 5
    End If
    ItemAviable = -1
End Sub

Public Sub Render(ByVal Shadow As Boolean, pShield As cls3dObject)
    If Not Alive Then Exit Sub

    Dim i As Long
    Dim RenderShield As Boolean

    'Render bike
    Moped.Render
    If Not Shadow Then
        'Render wheels
        Wheel.Position = BackRPos
        Wheel.Render
        Wheel.Position = FrontRPos
        Wheel.Render
        'Render shield, if nescessary
        If ShieldTime > 0 Then
            RenderShield = True
            If ShieldTime < 1.5 Then
                'Let the shield blink
                For i = 1 To 13 Step 2
                    If ShieldTime > i * 0.1 And ShieldTime < (i + 1) * 0.1 Then
                        RenderShield = False
                        Exit For
                    End If
                Next i
            End If
            If RenderShield Then
                pShield.Position = Vector3dMake(Pos.x, 0, Pos.y)
                pShield.Render
            End If
        End If
    Else
        'Render shadow wheels.
        SWheel.RotationMatrix = Wheel.RotationMatrix
        SWheel.Position = BackRPos
        SWheel.Render
        SWheel.Position = FrontRPos
        SWheel.Render
    End If
End Sub

Public Sub RenderWalls(ByVal Mirrored As Boolean)
    If Not Alive Then
        Wall.RenderTrans
        Exit Sub
    End If
    If Not gOptions.UseTransWalls Or Mirrored Then
        Wall.Render
    Else
        Wall.RenderTrans
    End If
End Sub
