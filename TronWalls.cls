VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TronWalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Walls As cls3dPolygons, WallsTop As cls3dPolygons
Private WallCnt As Long, WallUsed As Long
Private r As Byte, g As Byte, b As Byte
Private TransMat As D3DMATERIAL8, SolidMat As D3DMATERIAL8

Public Trans As Single

Public Sub Initialize(ByVal NrStaticWalls As Long, ByVal nr As Byte, ByVal ng As Byte, ByVal nb As Byte, pTex As cls2dTexture)
    WallCnt = NrStaticWalls
    WallUsed = 0
    Set Walls = New cls3dPolygons
    Walls.Initialize True
    Walls.SetVisibility BOTHSIDES
    Walls.SetPolyFormat TRIANGLELIST
    Walls.SetTexture pTex
    Set WallsTop = New cls3dPolygons
    WallsTop.Initialize True
    WallsTop.SetVisibility BOTHSIDES
    WallsTop.SetPolyFormat TRIANGLELIST
    SolidMat = MaterialMake(nr, ng, nb, 0, 0, 0, 0, 0)
    TransMat = MaterialMake(nr, ng, nb, 0, 0, 0, 0, 100)
    r = nr
    g = ng
    b = nb
    If gOptions.UseTransWalls Then
        Trans = 100
    Else
        Trans = 0
    End If
End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

Public Sub Cleanup()
    Set Walls = Nothing
    Set WallsTop = Nothing
End Sub

Public Sub Delete()
    WallUsed = 0
    TransMat = MaterialMake(r, g, b, 0, 0, 0, 0, 100)
    If gOptions.UseTransWalls Then
        Trans = 100
    Else
        Trans = 0
    End If
End Sub

Public Sub Rebuild(pWDesc As TronWallDesc, ByVal Player As Long)
    Dim i As Long, DescUsed As Long
    Dim pWalls As Long, curWall As TronWall
    Dim WallPos As D3DVECTOR2, WallVec As D3DVECTOR2, NWallVec As D3DVECTOR2

    DescUsed = pWDesc.GetWallUsed
    pWalls = pWDesc.GetWallDescPtr
    WallUsed = 0
    For i = 0 To DescUsed - 1
        CopyMemory curWall, ByVal pWalls + i * LenB(curWall), Len(curWall)
        If curWall.Used And curWall.From = Player Then
            'Add this wall
            WallPos = curWall.Pos1
            D3DXVec2Subtract WallVec, curWall.Pos2, WallPos
            D3DXVec2Normalize NWallVec, WallVec
            Do While D3DXVec2Length(WallVec) > 15
                AddWall
                UpdateWall WallPos, Vector2dMake(WallPos.x + NWallVec.x * 15, WallPos.y + NWallVec.y * 15)
                WallPos = Vector2dMake(WallPos.x + NWallVec.x * 15, WallPos.y + NWallVec.y * 15)
                D3DXVec2Subtract WallVec, curWall.Pos2, WallPos
            Loop
            'Add the last wall part of this straight wall.
            AddWall
            UpdateWall WallPos, curWall.Pos2
        End If
    Next i
    'Re-create the current wall, which is constantly updated.
    AddWall
End Sub

Public Sub AddWall()
    If WallUsed = WallCnt Then
        'Create space for new walls.
        WallCnt = WallCnt * 2
        Walls.ChangeBufferSize WallCnt * 6, 0
        WallsTop.ChangeBufferSize WallCnt * 6, 0
    End If
    WallUsed = WallUsed + 1
End Sub

Public Sub UpdateWall(Pos1 As D3DVECTOR2, Pos2 As D3DVECTOR2)
    Dim ActVertex(5) As Vertex, TopVertex(5) As Vertex
    Dim CalcVec As D3DVECTOR2, NVec As D3DVECTOR2
    Dim TexRep As Single
    Dim WallNr As Long
    
    If WallUsed = 0 Then Exit Sub
    WallNr = WallUsed - 1

    D3DXVec2Subtract CalcVec, Pos1, Pos2
    TexRep = D3DXVec2Length(CalcVec) / 15
    'Create vertex data
    ActVertex(0) = VertexMake(Pos1.x, 0, Pos1.y, 0, 1, 0, 0, 1)
    ActVertex(1) = VertexMake(Pos2.x, 2, Pos2.y, 0, 1, 0, TexRep, 0)
    ActVertex(2) = VertexMake(Pos2.x, 0, Pos2.y, 0, 1, 0, TexRep, 1)
    ActVertex(3) = ActVertex(0)
    ActVertex(4) = VertexMake(Pos1.x, 2, Pos1.y, 0, 1, 0, 0, 0)
    ActVertex(5) = ActVertex(1)
    'Create vertex top data
    NVec = NormalVector2dMake(Pos1, Pos2)
    D3DXVec2Scale NVec, NVec, 0.1
    TopVertex(0) = VertexMake(Pos1.x - NVec.x, 2, Pos1.y - NVec.y, 0, 1, 0, 0, 0)
    TopVertex(1) = VertexMake(Pos1.x + NVec.x, 2, Pos1.y + NVec.y, 0, 1, 0, 0, 0)
    TopVertex(2) = VertexMake(Pos2.x + NVec.x, 2, Pos2.y + NVec.y, 0, 1, 0, 0, 0)
    TopVertex(3) = TopVertex(0)
    TopVertex(4) = VertexMake(Pos2.x - NVec.x, 2, Pos2.y - NVec.y, 0, 1, 0, 0, 0)
    TopVertex(5) = TopVertex(2)
    'Assign vertex data to the walls.
    Walls.SetVertexData VarPtr(ActVertex(0)), WallNr * 6, 6
    WallsTop.SetVertexData VarPtr(TopVertex(0)), WallNr * 6, 6
End Sub

Public Sub TransIncr(ByVal RelTimeDiff As Single)
    If Trans = 255 Then Exit Sub
    Trans = Trans + RelTimeDiff * 200
    If Trans > 255 Then Trans = 255
    TransMat = MaterialMake(r, g, b, 0, 0, 0, 0, Trans)
End Sub

Public Sub Render()
    Walls.SetMaterial SolidMat
    Walls.Render 0, WallUsed * 2
    If Not gOptions.Cam = 6 Then
        WallsTop.SetMaterial SolidMat
        WallsTop.Render 0, WallUsed * 2
    End If
End Sub

Public Sub RenderTrans()
    If Trans = 255 Then Exit Sub
    Walls.SetMaterial TransMat
    Walls.Render 0, WallUsed * 2
    If Not gOptions.Cam = 6 Then
        WallsTop.SetMaterial TransMat
        WallsTop.Render 0, WallUsed * 2
    End If
End Sub

Public Function GetTransMat() As D3DMATERIAL8
    GetTransMat = TransMat
End Function

Public Function GetSolidMat() As D3DMATERIAL8
    GetSolidMat = SolidMat
End Function
