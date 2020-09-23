VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2dTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colKey As Long
Private UseTransMap As Boolean
Private MipLevels As Long
Private eolOffset As Long
Private LockRect As D3DLOCKED_RECT
Private TexDesc As D3DSURFACE_DESC
Private pTexture As Direct3DTexture8
Private pData() As StoreTexLevel

Private Type StoreTexLevel
    TexData() As Byte
End Type



Private Sub Class_Initialize()
    gEngine.RegisterObject Me, OBJ_2DTEXTURE
End Sub

Private Sub Class_Terminate()
    Cleanup
    gEngine.UnregisterObject Me
End Sub

'//-----------------------------------------------------------------------------
'// Function: Cleanup
'// Desc: Terminates the class.
'//-----------------------------------------------------------------------------
Public Sub Cleanup()
    EditEnd
    Set pTexture = Nothing
End Sub

'//-----------------------------------------------------------------------------
'// Function: LoadFromFile
'// Desc: Loads a bitmap from a bmp file into the texture.
'// Param: File (path of the bitmap file), ColorKey (transparent color),
'// MipLevels (how many mip mapping levels should be created, pass 0 to create
'// a complete mip map automatically), EnableEdit (if the texture should be
'// enabled to perform per texel editing, only possible if MipLevels = 1)
'// Return: succeeded or not.
'//-----------------------------------------------------------------------------
Public Function LoadFromFile(ByVal File As String, ByVal ColorKey As Long, Optional ByVal nMipLevels As Long = 1, Optional ByVal EnableEdit As Boolean = False) As Boolean
    On Local Error GoTo Failed
    Cleanup
    Set pTexture = TextureMake(File, ColorKey, nMipLevels, EnableEdit)
    pTexture.GetLevelDesc 0, TexDesc
    colKey = ColorKey
    MipLevels = nMipLevels
    UseTransMap = False
    LoadFromFile = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: LoadForEdit
'// Desc: Creates an empty texture for per texel editing.
'// Param: x (width of texture), y (height of texture)
'// Note: x and y should both be a power of 2 (2 ^ n)
'//-----------------------------------------------------------------------------
Public Function LoadForEdit(ByVal x As Long, ByVal y As Long) As Boolean
    On Local Error GoTo Failed
    Cleanup
    Set pTexture = gD3DX.CreateTexture(gD3DDevice, x, y, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED)
    pTexture.GetLevelDesc 0, TexDesc
    colKey = 0
    MipLevels = 1
    UseTransMap = False
    LoadForEdit = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: EditStart
'// Desc: Must be called at the beginning of a per texel editing operation.
'//-----------------------------------------------------------------------------
Public Function EditStart() As Boolean
    On Local Error GoTo Failed
    If Not MipLevels = 1 Or Not LockRect.pBits = 0 Then Exit Function
    'check pixel format (D3DFMT_A8R8G8B8 required)
    If Not TexDesc.Format = D3DFMT_A8R8G8B8 Then Exit Function
    'Lock entire texture
    LockRect.pBits = 0
    pTexture.LockRect 0, LockRect, ByVal 0, 0
    'Compute offset
    eolOffset = LockRect.Pitch - TexDesc.Width * 4
    EditStart = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: Edit
'// Desc: Manipulates a single texel on the texture.
'// Param: Pos (the x and y position on the texture, ranging from [0, 0] to
'// [x - 1, y - 1]), a (the transparency value), r, g, b (the color values).
'// Pass -1 to leave a value unchanged.
'//-----------------------------------------------------------------------------
Public Sub Edit(Pos As D3DVECTOR2, ByVal a As Long, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    On Local Error GoTo Failed

    Dim pTexData As Long, Data As Byte

    If LockRect.pBits = 0 Then Exit Sub
    If Pos.x < 0 Or Pos.y < 0 Or Pos.x >= TexDesc.Width Or Pos.y >= TexDesc.Height Then Exit Sub
    pTexData = LockRect.pBits + Pos.y * LockRect.Pitch + Pos.x * 4
    '1. byte: blue
    If Not b = -1 Then
        Data = b
        CopyMemory ByVal pTexData, Data, 1
    End If
    pTexData = pTexData + 1
    '2. Byte: green
    If Not g = -1 Then
        Data = g
        CopyMemory ByVal pTexData, Data, 1
    End If
    pTexData = pTexData + 1
    '3. Byte: red
    If Not r = -1 Then
        Data = r
        CopyMemory ByVal pTexData, Data, 1
    End If
    pTexData = pTexData + 1
    '4. Byte: alpha
    If Not a = -1 Then
        Data = a
        UseTransMap = True
        CopyMemory ByVal pTexData, Data, 1
    End If
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: EditEnd
'// Desc: Must be called at the end of a per texel editing operation.
'//-----------------------------------------------------------------------------
Public Sub EditEnd()
    On Local Error Resume Next
    If pTexture Is Nothing Then Exit Sub
    pTexture.UnlockRect 0
    LockRect.pBits = 0
End Sub


Friend Property Get getColKey() As Long
    getColKey = colKey
End Property

Friend Property Get getpTexture() As Direct3DTexture8
    Set getpTexture = pTexture
End Property

Friend Property Get getUseTransMap() As Boolean
    getUseTransMap = UseTransMap
End Property

Public Sub StoreData()
    On Local Error GoTo Failed
    
    Dim i As Long, MipMapCnt As Long
    Dim MipMapDesc As D3DSURFACE_DESC
    Dim LRect As D3DLOCKED_RECT

    MipMapCnt = pTexture.GetLevelCount
    EditEnd
    ReDim pData(MipMapCnt - 1)
    For i = 0 To MipMapCnt - 1
        'Save texture in system memory.
        pTexture.GetLevelDesc i, MipMapDesc
        ReDim pData(i).TexData(MipMapDesc.Size - 1)
        pTexture.LockRect i, LRect, ByVal 0, D3DLOCK_READONLY
        CopyMemory pData(i).TexData(0), ByVal LRect.pBits, MipMapDesc.Size
        pTexture.UnlockRect i
    Next i
Failed:
End Sub

Public Sub ReStoreData()
    On Local Error GoTo Failed
    
    Dim i As Long, MipMapCnt As Long
    Dim MipMapDesc As D3DSURFACE_DESC
    Dim LRect As D3DLOCKED_RECT

    Set pTexture = gD3DX.CreateTexture(gD3DDevice, TexDesc.Width, TexDesc.Height, MipLevels, 0, TexDesc.Format, D3DPOOL_MANAGED)
    MipMapCnt = pTexture.GetLevelCount
    'Restore texture from system memory.
    For i = 0 To MipMapCnt - 1
        pTexture.GetLevelDesc i, MipMapDesc
        pTexture.LockRect i, LRect, ByVal 0, 0
        CopyMemory ByVal LRect.pBits, pData(i).TexData(0), MipMapDesc.Size
        pTexture.UnlockRect i
    Next i
Failed:
End Sub

