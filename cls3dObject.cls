VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls3dObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pMesh As D3DXMesh
Private pMaterials() As D3DMATERIAL8
Private pTextures() As cls2dTexture
Private MaterialCnt As Long
Private mPosition As D3DVECTOR
Private matTrans As D3DMATRIX, matRot As D3DMATRIX
Private FVF As Long, VertexSize As Long, VertexCnt As Long, FaceCnt As Long
Private pVData() As Byte, pIData() As Byte, pAData() As Byte


Private Sub Class_Initialize()
    D3DXMatrixIdentity matTrans
    D3DXMatrixIdentity matRot
    gEngine.RegisterObject Me, OBJ_3DOBJ
End Sub

Private Sub Class_Terminate()
    Cleanup
    gEngine.UnregisterObject Me
End Sub

'//-----------------------------------------------------------------------------
'// Funktion: Cleanup
'// Desc: Terminates the object.
'//-----------------------------------------------------------------------------
Public Sub Cleanup()
    Dim i As Long
    
    For i = 0 To MaterialCnt - 1
        Set pTextures(i) = Nothing
    Next i
    Set pMesh = Nothing
End Sub

'Cleanup of the mesh only.
Friend Sub CleanupMesh()
    Set pMesh = Nothing
End Sub

'//-----------------------------------------------------------------------------
'// Function: LoadFromFile
'// Desc: Loads a 3d object from a x file.
'// Param: File, MipLevels (how many mip mapping levels of every texture of the
'// 3d object should be created, pass 0 to create complete mip maps automatically)
'//-----------------------------------------------------------------------------
Public Function LoadFromFile(ByVal File As String, Optional ByVal MipLevels As Long = 1) As Boolean
    On Local Error GoTo Failed
    
    Dim i As Long, Folder As String
    Dim pAdj() As Long
    Dim pMatBuffer As D3DXBuffer, pRdMat As Long, dRdMat As D3DXMATERIAL
    Dim pAdjBuffer As D3DXBuffer

    Cleanup
    Set pMesh = gD3DX.LoadMeshFromX(File, D3DXMESH_MANAGED, gD3DDevice, pAdjBuffer, pMatBuffer, MaterialCnt)
    Set pAdjBuffer = Nothing
    
    If MaterialCnt = 0 Then Exit Function
    pRdMat = pMatBuffer.GetBufferPointer
    ReDim pMaterials(MaterialCnt - 1)
    ReDim pTextures(MaterialCnt - 1)
    Folder = GetFolder(File)

    For i = 0 To MaterialCnt - 1
        'Create materials
        CopyMemory dRdMat, ByVal pRdMat + i * Len(dRdMat), Len(dRdMat)
        pMaterials(i) = dRdMat.MatD3D
        pMaterials(i).Ambient = pMaterials(i).diffuse
        pMaterials(i).Specular.a = 1
        pMaterials(i).Specular.r = 1
        pMaterials(i).Specular.g = 1
        pMaterials(i).Specular.b = 1
        'If nescessarry, create texture.
        If Not dRdMat.TextureFilename = "" Then
            pTextures(i).LoadFromFile Folder & dRdMat.TextureFilename, 0, MipLevels
        End If
    Next i

    Set pMatBuffer = Nothing
    'Optimize object
    ReDim pAdj(3 * pMesh.GetNumFaces - 1)
    pMesh.GenerateAdjacency 0, pAdj(0)
    pMesh.OptimizeInplace D3DXMESHOPT_ATTRSORT Or D3DXMESHOPT_COMPACT Or D3DXMESHOPT_VERTEXCACHE, pAdj(0), ByVal 0, ByVal 0, pAdjBuffer
    Set pAdjBuffer = Nothing
    'Save properties local.
    FVF = pMesh.GetFVF
    VertexSize = gD3DX.GetFVFVertexSize(FVF)
    VertexCnt = pMesh.GetNumVertices
    FaceCnt = pMesh.GetNumFaces
    'Compute normals.
    gD3DX.ComputeNormals pMesh

    SetTransparency 0
    SetSpecPower 0

    LoadFromFile = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Property: Position
'// Desc: Gets and sets the position of this object in world space.
'// Param: Pos (the position)
'//-----------------------------------------------------------------------------
Public Property Let Position(Pos As D3DVECTOR)
    mPosition = Pos
    D3DXMatrixTranslation matTrans, mPosition.x, mPosition.y, mPosition.z
End Property

Public Property Get Position() As D3DVECTOR
    Position = mPosition
End Property

'//-----------------------------------------------------------------------------
'// Function: Move
'// Desc: Moves the 3d object in world space.
'// Param: MoveDelta (the relative movement)
'//-----------------------------------------------------------------------------
Public Sub Move(MoveDelta As D3DVECTOR)
    D3DXVec3Add mPosition, mPosition, MoveDelta
    D3DXMatrixTranslation matTrans, mPosition.x, mPosition.y, mPosition.z
End Sub

'//-----------------------------------------------------------------------------
'// Property: RotationMatrix
'// Desc: Gets and sets the rotation of this object according to the given matrix.
'// Param: Matrix (the rotation matrix)
'//-----------------------------------------------------------------------------
Public Property Let RotationMatrix(Matrix As D3DMATRIX)
    matRot = Matrix
End Property

Public Property Get RotationMatrix() As D3DMATRIX
    RotationMatrix = matRot
End Property


'//-----------------------------------------------------------------------------
'// Function: SetRotation
'// Desc: Sets the rotation of this object new. Previous rotations are discarded.
'// Param: Angle (angle for x, y and z rotation), Order (the rotating order, pass a
'// vector which contains numbers from 0 to 2, eg. x = 1, y = 0, z = 2), Origin
'//-----------------------------------------------------------------------------
Public Sub SetRotation(Angle As D3DVECTOR, Order As D3DVECTOR, Origin As D3DVECTOR)
    Dim matTmp1 As D3DMATRIX, matTmp2 As D3DMATRIX

    D3DXMatrixTranslation matTmp1, mPosition.x - Origin.x, mPosition.y - Origin.y, mPosition.z - Origin.z
    matRot = RotationMatrixMake(Angle, Order)
    D3DXMatrixMultiply matTmp1, matTmp1, matRot
    D3DXMatrixTranslation matTmp2, Origin.x - mPosition.x, Origin.y - mPosition.y, Origin.z - mPosition.z
    D3DXMatrixMultiply matRot, matTmp1, matTmp2
End Sub

'//-----------------------------------------------------------------------------
'// Function: Rotate
'// Desc: Rotates the object. Previous rotations are considered.
'// Param: Angle (angle for x, y and z rotation), Order (the rotating order, pass a
'// vector which contains numbers from 0 to 2, eg. x = 1, y = 0, z = 2), Origin
'//-----------------------------------------------------------------------------
Public Sub Rotate(Angle As D3DVECTOR, Order As D3DVECTOR, Origin As D3DVECTOR)
    Dim matTmp1 As D3DMATRIX, matTmp2 As D3DMATRIX

    D3DXMatrixTranslation matTmp1, mPosition.x - Origin.x, mPosition.y - Origin.y, mPosition.z - Origin.z
    matTmp2 = RotationMatrixMake(Angle, Order)
    D3DXMatrixMultiply matTmp1, matTmp1, matTmp2
    D3DXMatrixTranslation matTmp2, Origin.x - mPosition.x, Origin.y - mPosition.y, Origin.z - mPosition.z
    D3DXMatrixMultiply matTmp1, matTmp1, matTmp2
    D3DXMatrixMultiply matRot, matRot, matTmp1
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetTransparency
'// Desc: Sets the transparency for this 3d object.
'//-----------------------------------------------------------------------------
Public Sub SetTransparency(ByVal Transparency As Byte)
    On Local Error GoTo Failed
    
    Dim i As Long

    For i = 0 To MaterialCnt - 1
        pMaterials(i).diffuse.a = 1 - Transparency / 255
    Next i
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Funktion: SetSpecPower
'// Desc: Sets the power of specular light reflections on this 3d object.
'//-----------------------------------------------------------------------------
Public Sub SetSpecPower(ByVal SpecPower As Single)
    On Local Error GoTo Failed
    
    Dim i As Long

    For i = 0 To MaterialCnt - 1
        pMaterials(i).power = SpecPower
    Next i
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Funktion: GetLowPoint
'// Desc: Returns the lower, left, front point of the object.
'//-----------------------------------------------------------------------------
Public Function GetLowPoint() As D3DVECTOR
    GetLowPoint = GetExtremePoint(0)
End Function

'//-----------------------------------------------------------------------------
'// Funktion: GetHighPoint
'// Desc: Returns the upper, right, back point of the object.
'//-----------------------------------------------------------------------------
Public Function GetHighPoint() As D3DVECTOR
    GetHighPoint = GetExtremePoint(1)
End Function

'//-----------------------------------------------------------------------------
'// Funktion: GetMidPoint
'// Desc: Returns the center point of the object.
'//-----------------------------------------------------------------------------
Public Function GetMidPoint() As D3DVECTOR
    GetMidPoint = GetExtremePoint(2)
End Function

'//-----------------------------------------------------------------------------
'// Function: Render
'// Desc: Renders the 3d object.
'//-----------------------------------------------------------------------------
Public Sub Render()
    On Local Error GoTo Failed
    
    Dim i As Long
    Dim myCurTex As cls2dTexture

    gD3DDevice.SetVertexShader VERTEX_FVF
    'Check for shadows
    If Not gEngine.GetUseShadow Then
        gEngine.SetTransform matTrans, matRot
        gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
    Else
        gEngine.SetTransformMultiply matTrans, matRot, gEngine.GetMatShadow
    End If
    'Check for mirroring
    If Not gEngine.GetUseMirror Then
        gD3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
    Else
        gD3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
    End If
    'Render object
    gD3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    For i = 0 To MaterialCnt - 1
        If Not gEngine.GetUseShadow Then
            gD3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, IIf(pMaterials(i).diffuse.a = 1, False, True)
            If Not pTextures(i) Is Nothing Then
                gD3DDevice.SetTexture 0, pTextures(i)
            Else
                gD3DDevice.SetTexture 0, Nothing
            End If
            gD3DDevice.SetMaterial pMaterials(i)
        End If
        pMesh.DrawSubset i
    Next i
    If Not gEngine.GetUseShadow Then gD3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
Failed:
End Sub




Private Function GetExtremePoint(ByVal Index As Long) As D3DVECTOR
    On Local Error GoTo Failed
    
    Dim i As Long
    Dim lx As Single, ly As Single, lz As Single, hx As Single, hy As Single, hz As Single
    Dim pVert As Long
    Dim RdVec As D3DVECTOR, retVec As D3DVECTOR

    pVert = pMesh.LockVertexBuffer(D3DLOCK_READONLY)
    For i = 0 To VertexCnt - 1
        CopyMemory RdVec, ByVal pVert + i * VertexSize, Len(RdVec)
        If i = 0 Or RdVec.x < lx Then lx = RdVec.x
        If i = 0 Or RdVec.x > hx Then hx = RdVec.x
        If i = 0 Or RdVec.y < ly Then ly = RdVec.y
        If i = 0 Or RdVec.y > hy Then hy = RdVec.y
        If i = 0 Or RdVec.z < lz Then lz = RdVec.z
        If i = 0 Or RdVec.z > hz Then hz = RdVec.z
    Next i
    pMesh.UnlockVertexBuffer
    If Index = 0 Then
        'Minima
        retVec = Vector3dMake(lx, ly, lz)
    ElseIf Index = 1 Then
        'Maxima
        retVec = Vector3dMake(hx, hy, hz)
    Else
        'Center
        retVec = Vector3dMake((lx + hx) / 2, (ly + hy) / 2, (lz + hz) / 2)
    End If
    D3DXVec3Add retVec, retVec, mPosition
    GetExtremePoint = retVec
Failed:
End Function

Friend Property Get getpMesh() As D3DXMesh
    Set getpMesh = pMesh
End Property

Public Sub StoreData()
    On Local Error GoTo Failed
    
    Dim pLockData As Long

    'Store vertex data
    ReDim pVData(VertexCnt * VertexSize - 1)
    pLockData = pMesh.LockVertexBuffer(D3DLOCK_READONLY)
    CopyMemory pVData(0), ByVal pLockData, VertexCnt * VertexSize
    pMesh.UnlockVertexBuffer
    'Store index data
    ReDim pIData(FaceCnt * 6 - 1)
    pLockData = pMesh.LockIndexBuffer(D3DLOCK_READONLY)
    CopyMemory pIData(0), ByVal pLockData, FaceCnt * 6
    pMesh.UnlockIndexBuffer
    'Store attribute data
    ReDim pAData(FaceCnt * 4 - 1)
    pLockData = pMesh.LockAttributeBuffer(D3DLOCK_READONLY)
    CopyMemory pAData(0), ByVal pLockData, FaceCnt * 4
    pMesh.UnlockAttributeBuffer
Failed:
End Sub

Public Sub ReStoreData()
    On Local Error GoTo Failed
    
    Dim pLockData As Long

    'Recreate mesh
    Set pMesh = gD3DX.CreateMeshFVF(FaceCnt, VertexCnt, D3DXMESH_MANAGED, FVF, gD3DDevice)
    'Restore vertex data
    pLockData = pMesh.LockVertexBuffer(0)
    CopyMemory ByVal pLockData, pVData(0), VertexCnt * VertexSize
    pMesh.UnlockVertexBuffer
    'Restore index data
    pLockData = pMesh.LockIndexBuffer(0)
    CopyMemory ByVal pLockData, pIData(0), FaceCnt * 6
    pMesh.UnlockIndexBuffer
    'Restore attribute data
    pLockData = pMesh.LockAttributeBuffer(0)
    CopyMemory ByVal pLockData, pAData(0), FaceCnt * 4
    pMesh.UnlockAttributeBuffer
Failed:
End Sub
