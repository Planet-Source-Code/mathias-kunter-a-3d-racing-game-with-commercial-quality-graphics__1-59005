VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TronCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CamPos As D3DVECTOR, BlendPos As D3DVECTOR
Private CamDest As D3DVECTOR, BlendDest As D3DVECTOR
Private CamAng As Single, BlendAng As Single
Private PosTime As Single, DestTime As Single, AngleTime As Single

Public Property Let Pos(Pos As D3DVECTOR)
    CamPos = Pos
End Property
Public Property Get Pos() As D3DVECTOR
    Pos = CamPos
End Property

Public Property Let Dest(Dest As D3DVECTOR)
    CamDest = Dest
End Property
Public Property Get Dest() As D3DVECTOR
    Dest = CamDest
End Property
    
Public Function GetAngle() As Single
    GetAngle = CamAng
End Function

Public Function IsAngleBlended() As Boolean
    If Not AngleTime = 0 Then IsAngleBlended = True
End Function

Public Sub BlendPosition(NewPos As D3DVECTOR, ByVal TimeX As Single)
    D3DXVec3Subtract BlendPos, NewPos, CamPos
    If Not TimeX = 0 Then D3DXVec3Scale BlendPos, BlendPos, 1 / TimeX
    PosTime = TimeX
End Sub

Public Sub BlendDestination(NewDest As D3DVECTOR, ByVal TimeX As Single)
    D3DXVec3Subtract BlendDest, NewDest, CamDest
    D3DXVec3Scale BlendDest, BlendDest, 1 / TimeX
    DestTime = TimeX
End Sub

Public Sub BlendAngle(ByVal ActAngle As Single, ByVal NewAngle As Single, ByVal TimeX As Single)
    CamAng = ActAngle
    BlendAng = (NewAngle - ActAngle) / TimeX
    AngleTime = TimeX
End Sub

Public Sub DoEventsX(ByVal RelTimeDiff As Single)
    Dim BlendTime As Single, AddVec As D3DVECTOR
    
    If PosTime > 0 Then
        'Blend camera position
        If RelTimeDiff > PosTime Then
            BlendTime = PosTime
        Else
            BlendTime = RelTimeDiff
        End If
        D3DXVec3Scale AddVec, BlendPos, BlendTime
        D3DXVec3Add CamPos, CamPos, AddVec
        PosTime = PosTime - BlendTime
    End If
    If DestTime > 0 Then
        'Blend camera destination
        If RelTimeDiff > DestTime Then
            BlendTime = DestTime
        Else
            BlendTime = RelTimeDiff
        End If
        D3DXVec3Scale AddVec, BlendDest, BlendTime
        D3DXVec3Add CamDest, CamDest, AddVec
        DestTime = DestTime - BlendTime
    End If
    If AngleTime > 0 Then
        'Blend camera angle
        If RelTimeDiff > AngleTime Then
            BlendTime = AngleTime
        Else
            BlendTime = RelTimeDiff
        End If
        CamAng = CamAng + BlendAng * BlendTime
        AngleTime = AngleTime - BlendTime
    End If
End Sub

Public Sub SetCamera()
    Dim SetPos As D3DVECTOR
    
    SetPos = CamPos
    If SetPos.x < 1 Then SetPos.x = 1
    If SetPos.x > gOptions.LandSize - 1 Then SetPos.x = gOptions.LandSize - 1
    If SetPos.z < 1 Then SetPos.z = 1
    If SetPos.z > gOptions.LandSize - 1 Then SetPos.z = gOptions.LandSize - 1
    gEngine.SetCamera SetPos, CamDest, 0
End Sub

