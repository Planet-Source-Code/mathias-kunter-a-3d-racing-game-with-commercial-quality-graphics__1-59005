VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TronExplosion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Explosion As cls3dPolygons
Private ExplCnt As Long
Private pExplDesc() As TronExplDesc

Public Sub Initialize(ByVal NrStaticExplosion As Long, pTex As cls2dTexture)
    Dim i As Long

    ExplCnt = NrStaticExplosion
    ReDim pExplDesc(ExplCnt - 1)
    Set Explosion = New cls3dPolygons
    Explosion.Initialize True
    Explosion.SetPolyFormat TRIANGLELIST
    Explosion.SetVisibility BOTHSIDES
    Explosion.SetTexture pTex
    Explosion.ChangeBufferSize 180, 0
End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

Public Sub Cleanup()
    Set Explosion = Nothing
End Sub

Public Sub CreateExplosion(Pos As D3DVECTOR2, ByVal Small As Boolean)
    Dim i As Long
    Dim ActExpl As Long

    For i = 0 To ExplCnt - 1
        If Not pExplDesc(i).Used Then Exit For
    Next i
    If i = ExplCnt Then
        'Create space for new explosions.
        ExplCnt = ExplCnt + 10
        ReDim Preserve pExplDesc(ExplCnt - 1)
    End If
    ActExpl = i
    With pExplDesc(ActExpl)
        .Used = True
        .Trans = 0.2
        If Small Then
            .FlameCnt = 10
        Else
            .FlameCnt = 30
        End If
        For i = 0 To .FlameCnt - 1
            'Create flames.
            .Flames(i).Pos = Vector3dMake(Pos.x + GetRandomVal(-2, 2), 0, Pos.y + GetRandomVal(-2, 2))
            .Flames(i).Dir = Vector3dMake(GetRandomVal(-5, 5), GetRandomVal(5, 20), GetRandomVal(-5, 5))
        Next i
    End With
End Sub

Public Sub DoEventsX(ByVal RelTimeDiff As Single)
    Dim i As Long, j As Long
    Dim AddVec As D3DVECTOR

    For i = 0 To ExplCnt - 1
        With pExplDesc(i)
            If .Used Then
                .Trans = .Trans + RelTimeDiff / 5
                If .Trans >= 1 Then
                    .Used = False
                Else
                    For j = 0 To .FlameCnt - 1
                        If .Flames(j).Pos.y >= 0 Then
                            .Flames(j).Dir.y = .Flames(j).Dir.y - 10 * RelTimeDiff
                            D3DXVec3Scale AddVec, .Flames(j).Dir, RelTimeDiff
                            D3DXVec3Add .Flames(j).Pos, .Flames(j).Pos, AddVec
                            If .Flames(j).Pos.x < 0 Then
                                .Flames(j).Pos.x = 0
                                .Flames(j).Dir.x = .Flames(j).Dir.x * -1
                            End If
                            If .Flames(j).Pos.x > gOptions.LandSize Then
                                .Flames(j).Pos.x = gOptions.LandSize
                                .Flames(j).Dir.x = .Flames(j).Dir.x * -1
                            End If
                            If .Flames(j).Pos.z < 0 Then
                                .Flames(j).Pos.z = 0
                                .Flames(j).Dir.z = .Flames(j).Dir.z * -1
                            End If
                            If .Flames(j).Pos.z > gOptions.LandSize Then
                                .Flames(j).Pos.z = gOptions.LandSize
                                .Flames(j).Dir.z = .Flames(j).Dir.z * -1
                            End If
                        End If
                    Next j
                End If
            End If
        End With
    Next i
End Sub

Public Sub Render(CamPos As D3DVECTOR)
    Dim i As Long, j As Long, ActIndex As Long
    Dim vData(179) As Vertex
    Dim Mat As D3DMATERIAL8

    Mat = gEngine.GetDefaultMaterial
    gEngine.SetZState True, False
    For i = 0 To ExplCnt - 1
        If pExplDesc(i).Used Then
            ActIndex = 0
            For j = 0 To pExplDesc(i).FlameCnt - 1
                If pExplDesc(i).Flames(j).Pos.y >= 0 Then
                    BillboardVertexMake pExplDesc(i).Flames(j).Pos, 3, 3, CamPos, True, VarPtr(vData(ActIndex))
                    'Convert format to triangle list
                    vData(ActIndex + 5) = vData(ActIndex + 3)
                    vData(ActIndex + 4) = vData(ActIndex + 1)
                    vData(ActIndex + 3) = vData(ActIndex + 2)
                    ActIndex = ActIndex + 6
                End If
            Next j
            If Not ActIndex = 0 Then
                Explosion.SetVertexData VarPtr(vData(0)), 0, ActIndex
                Mat.diffuse.a = 1 - pExplDesc(i).Trans
                Explosion.SetMaterial Mat
                Explosion.Render 0, ActIndex / 3
            End If
        End If
    Next i
    gEngine.SetZState True, True
End Sub
