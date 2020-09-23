VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls3dPolygons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PolyFormat As PolyType
Private pTex As cls2dTexture
Private Mat As D3DMATERIAL8
Private Visibility As PolyVisibility
Private Indexed As Boolean, Dynamic As Boolean
Private VertexMaxCnt As Long, IndexMaxCnt As Long
Private pVData() As Vertex, vLen As Vertex, pIData() As Integer, iLen As Integer, iStored As Boolean
Private mPosition As D3DVECTOR
Private matTrans As D3DMATRIX, matRot As D3DMATRIX
Private pVBuf As Direct3DVertexBuffer8
Private pIBuf As Direct3DIndexBuffer8


'//-----------------------------------------------------------------------------
'// Function: Initialize
'// Desc: Initializes the 3d-Polygons.
'// Param: DynamicUsage (if the 3d polygons should be optimized for dynamic usage.
'// Set this parameter to true if you're going to change the 3d data often, e.g.
'// every frame. However, you can also change a non-dynamic buffer every frame,
'// but this will be slower)
'//-----------------------------------------------------------------------------
Public Sub Initialize(ByVal DynamicUsage As Boolean)
    D3DXMatrixIdentity matTrans
    D3DXMatrixIdentity matRot
    PolyFormat = TRIANGLELIST
    Mat = gEngine.GetDefaultMaterial
    Dynamic = DynamicUsage
    Visibility = FRONTSIDE
End Sub

Private Sub Class_Initialize()
    gEngine.RegisterObject Me, OBJ_3DPOLYS
End Sub

Private Sub Class_Terminate()
    Cleanup
    gEngine.UnregisterObject Me
End Sub

'//-----------------------------------------------------------------------------
'// Function: Cleanup
'// Desc: Terminates the 3d-Polygons.
'//-----------------------------------------------------------------------------
Public Sub Cleanup()
    Set pVBuf = Nothing
    Set pIBuf = Nothing
    Set pTex = Nothing
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetVertexData
'// Desc: Sets the vertex data for the polygons which should be rendered.
'// If nescessary, the size of the buffer will be adjusted.
'// Param: vData (pointer to a vertex array), Offset (start index in the
'// destination buffer, in vertices, beginning with 0), cnt (number of vertices
'// which should be copied)
'//-----------------------------------------------------------------------------
Public Sub SetVertexData(ByVal vData As Long, ByVal Offset As Long, ByVal cnt As Long)
    On Local Error GoTo Failed
    
    Dim pVertices As Long

    If Offset + cnt > VertexMaxCnt Then
        'Buffer size must be changed.
        If Not ChangeBufferSize(Offset + cnt, 0) Then Exit Sub
    End If
    'Copy the vertices into the buffer.
    If Dynamic Then
        CopyMemory pVData(Offset), ByVal vData, cnt * Len(vLen)
    Else
        pVBuf.Lock Offset * Len(vLen), cnt * Len(vLen), pVertices, 0
        CopyMemory ByVal pVertices, ByVal vData, cnt * Len(vLen)
        pVBuf.Unlock
    End If
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: GetVertexData
'// Desc: Returns the vertex data.
'// Param: vData (out pointer to a vertex array), Offset (start index in the
'// destination array, in vertices, beginning with 0), cnt (number of vertices
'// which should be copied)
'//-----------------------------------------------------------------------------
Public Sub GetVertexData(ByVal vData As Long, ByVal Offset As Long, ByVal cnt As Long)
    On Local Error GoTo Failed

    Dim pVertices As Long

    If Offset + cnt > VertexMaxCnt Then Exit Sub
    'Read vertices
    If Dynamic Then
        CopyMemory ByVal vData, pVData(Offset), cnt * Len(vLen)
    Else
        pVBuf.Lock Offset * Len(vLen), cnt * Len(vLen), pVertices, D3DLOCK_READONLY
        CopyMemory ByVal vData, ByVal pVertices, cnt * Len(vLen)
        pVBuf.Unlock
    End If
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetIndexData
'// Desc: Sets the index data for the polygons which should be rendered.
'// If nescessary, the size of the buffer will be adjusted.
'// Param: iData (pointer to an integer array), Offset (start index in the
'// destination buffer, in indices, beginning with 0), cnt (number of indices
'// which should be copied)
'//-----------------------------------------------------------------------------
Public Sub SetIndexData(ByVal iData As Long, ByVal Offset As Long, ByVal cnt As Long)
    On Local Error GoTo Failed
    
    Dim pIndices As Long

    If Offset + cnt > IndexMaxCnt Then
        'Buffer size must be changed.
        If Not ChangeBufferSize(0, Offset + cnt) Then Exit Sub
    End If
    'Copy the indices into the buffer.
    If Dynamic Then
        CopyMemory pIData(Offset), ByVal iData, cnt * Len(iLen)
    Else
        pIBuf.Lock Offset * Len(iLen), cnt * Len(iLen), pIndices, 0
        CopyMemory ByVal pIndices, ByVal iData, cnt * Len(iLen)
        pIBuf.Unlock
    End If
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: GetIndexData
'// Desc: Returns the index data.
'// Param: iData (out pointer to an integer array), Offset (start index in the
'// destination array, in indices, beginning with 0), cnt (number of indices
'// which should be copied)
'//-----------------------------------------------------------------------------
Public Sub GetIndexData(ByVal iData As Long, ByVal Offset As Long, ByVal cnt As Long)
    On Local Error GoTo Failed

    Dim pIndices As Long

    If Offset + cnt > IndexMaxCnt Then Exit Sub
    'Read indices
    If Dynamic Then
        CopyMemory ByVal iData, pIData(Offset), cnt * Len(iLen)
    Else
        pIBuf.Lock Offset * Len(iLen), cnt * Len(iLen), pIndices, D3DLOCK_READONLY
        CopyMemory ByVal iData, ByVal pIndices, cnt * Len(iLen)
        pIBuf.Unlock
    End If
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetMaterial
'// Desc: Sets the material of the polygons.
'//-----------------------------------------------------------------------------
Public Sub SetMaterial(nMat As D3DMATERIAL8)
    Mat = nMat
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetPolyFormat
'// Desc: Sets the format of the polygons.
'//-----------------------------------------------------------------------------
Public Sub SetPolyFormat(ByVal nPolyFormat As PolyType)
    PolyFormat = nPolyFormat
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetTexture
'// Desc: Sets the texture of the polygons.
'//-----------------------------------------------------------------------------
Public Sub SetTexture(ByRef nTex As cls2dTexture)
    Set pTex = nTex
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetVisibility
'// Desc: Sets the visibility of the polygons.
'//-----------------------------------------------------------------------------
Public Sub SetVisibility(ByVal nVisibility As PolyVisibility)
    Visibility = nVisibility
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetIndexed
'// Desc: Sets if the polygons should be rendered indexed or not.
'//-----------------------------------------------------------------------------
Public Sub SetIndexed(ByVal nIndexed As Boolean)
    Indexed = nIndexed
End Sub

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
'// Desc: Moves the polygons in world space.
'// Param: MoveDelta (the relative movement)
'//-----------------------------------------------------------------------------
Public Sub Move(MoveDelta As D3DVECTOR)
    D3DXVec3Add mPosition, mPosition, MoveDelta
    D3DXMatrixTranslation matTrans, mPosition.x, mPosition.y, mPosition.z
End Sub

'//-----------------------------------------------------------------------------
'// Property: RotationMatrix
'// Desc: Gets and sets the rotation of this polygons according to the given matrix.
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
'// Desc: Sets the rotation of this polygons new. Previous rotations are discarded.
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
'// Desc: Rotates the polygons. Previous rotations are considered.
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
'// Function: ChangeBufferSize
'// Desc: Changes the size of the vertex or index buffer.
'// Param: NewStaticVCnt (the new size of the vertex buffer, pass 0 to leave it unchanged),
'// NewStaticICnt (the new size of the index buffer, pass 0 to leave it unchanged)
'//-----------------------------------------------------------------------------
Public Function ChangeBufferSize(ByVal NewStaticVCnt As Long, ByVal NewStaticICnt As Long) As Boolean
    On Local Error GoTo Failed
    
    Dim pVertices As Long, pIndices As Long, vSave() As Vertex, iSave() As Integer, bRestore As Boolean
    Dim SaveVertexCnt As Long, SaveIndexCnt As Long

    'Change vertex buffer
    If NewStaticVCnt > 0 Then
        If Dynamic Then
            If Not NewStaticVCnt = VertexMaxCnt Then
                VertexMaxCnt = NewStaticVCnt
                ReDim Preserve pVData(VertexMaxCnt - 1)
            End If
        Else
            If Not NewStaticVCnt = VertexMaxCnt Or pVBuf Is Nothing Then
                bRestore = False
                If Not pVBuf Is Nothing Then
                    'Save data from the existing vertex buffer.
                    If NewStaticVCnt > VertexMaxCnt Then
                        SaveVertexCnt = VertexMaxCnt
                    Else
                        SaveVertexCnt = NewStaticVCnt
                    End If
                    ReDim vSave(SaveVertexCnt - 1)
                    pVBuf.Lock 0, SaveVertexCnt * Len(vSave(0)), pVertices, D3DLOCK_READONLY
                    CopyMemory vSave(0), ByVal pVertices, SaveVertexCnt * Len(vSave(0))
                    pVBuf.Unlock
                    Set pVBuf = Nothing
                    bRestore = True
                End If
                VertexMaxCnt = NewStaticVCnt
                'Create a new buffer
                Set pVBuf = gD3DDevice.CreateVertexBuffer(VertexMaxCnt * Len(vSave(0)), 0, VERTEX_FVF, D3DPOOL_MANAGED)
                'Restore saved data.
                If bRestore Then
                    pVBuf.Lock 0, SaveVertexCnt * Len(vSave(0)), pVertices, 0
                    CopyMemory ByVal pVertices, vSave(0), SaveVertexCnt * Len(vSave(0))
                    pVBuf.Unlock
                End If
            End If
        End If
    End If

    'Change index buffer
    If NewStaticICnt > 0 Then
        If Dynamic Then
            If Not NewStaticICnt = IndexMaxCnt Then
                IndexMaxCnt = NewStaticICnt
                ReDim Preserve pIData(IndexMaxCnt - 1)
            End If
        Else
            If Not NewStaticICnt = IndexMaxCnt Or pIBuf Is Nothing Then
                bRestore = False
                If Not pIBuf Is Nothing Then
                    'Save data from the existing index buffer.
                    If NewStaticICnt > IndexMaxCnt Then
                        SaveIndexCnt = IndexMaxCnt
                    Else
                        SaveIndexCnt = NewStaticICnt
                    End If
                    ReDim iSave(SaveIndexCnt - 1)
                    pIBuf.Lock 0, SaveIndexCnt * Len(iSave(0)), pIndices, D3DLOCK_READONLY
                    CopyMemory iSave(0), ByVal pIndices, SaveIndexCnt * Len(iSave(0))
                    pIBuf.Unlock
                    Set pIBuf = Nothing
                    bRestore = True
                End If
                IndexMaxCnt = NewStaticICnt
                'Create a new buffer
                Set pIBuf = gD3DDevice.CreateIndexBuffer(IndexMaxCnt * Len(iSave(0)), 0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
                'Restore saved data.
                If bRestore Then
                    pIBuf.Lock 0, SaveIndexCnt * Len(iSave(0)), pIndices, 0
                    CopyMemory ByVal pIndices, iSave(0), SaveIndexCnt * Len(iSave(0))
                    pIBuf.Unlock
                End If
            End If
        End If
    End If
    ChangeBufferSize = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: Render
'// Desc: Renders the 3d polygons.
'// Param: Offset (the index of the first vertex or index which should be rendered,
'// beginning with 0), PolyCnt (the number of polygons to render)
'//-----------------------------------------------------------------------------
Public Sub Render(ByVal Offset As Long, ByVal PolyCnt As Long)
    On Local Error GoTo Failed
    
    Dim UseTexture As Direct3DTexture8, tmpVertex As Vertex, tmpIndex As Integer
    Dim bBlend As Boolean
        
    gD3DDevice.SetVertexShader VERTEX_FVF
    'Check for shadows
    If Not gEngine.GetUseShadow Then
        gEngine.SetTransform matTrans, matRot
        If Not pTex Is Nothing Then
            Set UseTexture = pTex.getpTexture
            If Not UseTexture Is Nothing Then
                If Not UseTexture.GetLevelCount = 1 Then
                    gD3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
                Else
                    gD3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
                End If
                If Not pTex.getColKey = 0 Or pTex.getUseTransMap Then bBlend = True
            End If
        End If
        gD3DDevice.SetTexture 0, UseTexture
        gD3DDevice.SetMaterial Mat
        If Not Mat.diffuse.a = 1 Or bBlend Then
            'A blending operation must be performed. Enable alpha blending.
            gD3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
            If Not Mat.diffuse.a = 1 And bBlend Then
                gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
            ElseIf Not Mat.diffuse.a = 1 Then
                gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
            Else
                gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2
            End If
        End If
    Else
        gEngine.SetTransformMultiply matTrans, matRot, gEngine.GetMatShadow
    End If
    If Not Visibility = NOCHANGE Then
        'Check for mirroring
        If Not gEngine.GetUseMirror Then
            gD3DDevice.SetRenderState D3DRS_CULLMODE, Visibility
        Else
            gD3DDevice.SetRenderState D3DRS_CULLMODE, InvertCullMode(Visibility)
        End If
    End If
    'Render the polygons
    If Dynamic Then
        If Indexed Then
            gD3DDevice.DrawIndexedPrimitiveUP PolyFormat, 0, VertexMaxCnt, PolyCnt, pIData(0), D3DFMT_INDEX16, pVData(0), Len(vLen)
        Else
            gD3DDevice.DrawPrimitiveUP PolyFormat, PolyCnt, pVData(0), Len(vLen)
        End If
    Else
        gD3DDevice.SetStreamSource 0, pVBuf, Len(tmpVertex)
        If Indexed Then
            gD3DDevice.SetIndices pIBuf, 0
            gD3DDevice.DrawIndexedPrimitive PolyFormat, 0, VertexMaxCnt, Offset, PolyCnt
        Else
            gD3DDevice.DrawPrimitive PolyFormat, Offset, PolyCnt
        End If
    End If
    If Not gEngine.GetUseShadow Then gD3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
Failed:
End Sub

Private Function InvertCullMode(ByVal Mode As CONST_D3DCULL) As CONST_D3DCULL
    If Mode = D3DCULL_CW Then
        InvertCullMode = D3DCULL_CCW
    ElseIf Mode = D3DCULL_CCW Then
        InvertCullMode = D3DCULL_CW
    Else
        InvertCullMode = D3DCULL_NONE
    End If
End Function

Friend Property Get getDynamic() As Boolean
    getDynamic = Dynamic
End Property

Public Sub StoreData()
    If Dynamic Then Exit Sub
    On Local Error GoTo Failed

    Dim pLockData As Long

    'Store vertex data
    ReDim pVData(VertexMaxCnt * Len(vLen) - 1)
    pVBuf.Lock 0, 0, pLockData, D3DLOCK_READONLY
    CopyMemory pVData(0), ByVal pLockData, VertexMaxCnt * Len(vLen)
    pVBuf.Unlock
    'Store index data, if nescessary.
    If Not (pIBuf Is Nothing) Then
        ReDim pIData(IndexMaxCnt * Len(iLen) - 1)
        pIBuf.Lock 0, 0, pLockData, D3DLOCK_READONLY
        CopyMemory pIData(0), ByVal pLockData, IndexMaxCnt * Len(iLen)
        pIBuf.Unlock
        iStored = True
    End If
    iStored = False
Failed:
End Sub

Public Sub ReStoreData()
    If Dynamic Then Exit Sub
    On Local Error GoTo Failed

    Dim pLockData As Long

    'Recreate buffer.
    If Not ChangeBufferSize(VertexMaxCnt, IndexMaxCnt) Then Exit Sub
    'Restore vertex data.
    pVBuf.Lock 0, 0, pLockData, 0
    CopyMemory ByVal pLockData, pVData(0), VertexMaxCnt * Len(vLen)
    pVBuf.Unlock
    'Restore index data, if nescessary.
    If iStored Then
        pIBuf.Lock 0, 0, pLockData, 0
        CopyMemory ByVal pLockData, pIData(0), IndexMaxCnt * Len(iLen)
        pIBuf.Unlock
    End If
Failed:
End Sub
