VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2dPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Vert(3) As TLVertex
Private NoRotVert(3) As TLVertex
Private RotAngle As Single
Private MP As D3DVECTOR2
Private alpha As Byte, red As Byte, green As Byte, blue As Byte


Private Sub Class_Initialize()
    Dim Col As Long

    alpha = 255
    red = 255
    green = 255
    blue = 255
    Col = ColorMake(255, 255, 255)
    Vert(0) = TLVertexMake(0, 0, Col, 0, 0)
    Vert(1) = TLVertexMake(0, 0, Col, 1, 0)
    Vert(2) = TLVertexMake(0, 0, Col, 0, 1)
    Vert(3) = TLVertexMake(0, 0, Col, 1, 1)
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetPosition
'// Desc: Sets the position of the picture on the screen.
'// Param: x1, y1 (the upper left corner), x2, y2 (the lower right corner)
'// Note: This function makes a rotation undone.
'//-----------------------------------------------------------------------------
Public Sub SetPosition(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    Vert(0).x = x1
    Vert(0).y = y1
    Vert(1).x = x2
    Vert(1).y = y1
    Vert(2).x = x1
    Vert(2).y = y2
    Vert(3).x = x2
    Vert(3).y = y2
    CopyMemory NoRotVert(0), Vert(0), 4 * Len(Vert(0))          'NoRotVert = Vert
    MP = Vector2dMake((x1 + x2) / 2, (y1 + y2) / 2)
End Sub

'//-----------------------------------------------------------------------------
'// Function: Move
'// Desc: Moves the picture on the screen.
'// Param: MoveX, MoveY (the relative movement)
'//-----------------------------------------------------------------------------
Public Sub Move(ByVal MoveX As Long, ByVal MoveY As Long)
    Dim i As Long

    For i = 0 To 3
        Vert(i).x = Vert(i).x + MoveX
        Vert(i).y = Vert(i).y + MoveY
    Next i
End Sub

'//-----------------------------------------------------------------------------
'// Function: Rotate
'// Desc: Rotates the picture clockwise around its centre.
'//-----------------------------------------------------------------------------
Public Sub Rotate(ByVal Angle As Single)
    Dim i As Long
    Dim cosPhi As Single, sinPhi As Single

    RotAngle = RotAngle + Angle
    cosPhi = Cos(RotAngle)
    sinPhi = Sin(RotAngle)
    For i = 0 To 3
        Vert(i).x = (NoRotVert(i).x - MP.x) * cosPhi - (NoRotVert(i).y - MP.y) * sinPhi + MP.x
        Vert(i).y = (NoRotVert(i).x - MP.x) * sinPhi + (NoRotVert(i).y - MP.y) * cosPhi + MP.y
    Next i
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetPictureRange
'// Desc: Defines the shown range of a texture on this picture.
'// Param: x1, y1 (the upper left corner), x2, y2 (the lower right corner)
'// Note: Pass [0, 0] and [1, 1] to show the entire texture.
'//-----------------------------------------------------------------------------
Public Sub SetPictureRange(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    Vert(0).tu = x1
    Vert(0).tv = y1
    Vert(1).tu = x2
    Vert(1).tv = y1
    Vert(2).tu = x1
    Vert(2).tv = y2
    Vert(3).tu = x2
    Vert(3).tv = y2
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetBackcolor
'// Desc: Sets the background color for this picture.
'// Param: r, g, b
'//-----------------------------------------------------------------------------
Public Sub SetBackcolor(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
    Dim i As Long, Col As Long

    red = r
    green = g
    blue = b
    Col = ColorAlphaMake(alpha, red, green, blue)
    For i = 0 To 3
        Vert(i).Color = Col
    Next i
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetTransparency
'// Desc: Sets the transparency for this picture.
'// Param: Transparency (from 0 to 255)
'//-----------------------------------------------------------------------------
Public Sub SetTransparency(ByVal Transparency As Byte)
    Dim i As Long, Col As Long

    alpha = 255 - Transparency
    Col = ColorAlphaMake(alpha, red, green, blue)
    For i = 0 To 3
        Vert(i).Color = Col
    Next i
End Sub

'//-----------------------------------------------------------------------------
'// Function: Render
'// Desc: Renders the 2d picture.
'// Param: Tex (the texture which should be used)
'//-----------------------------------------------------------------------------
Public Sub Render(Tex As cls2dTexture)
    On Local Error GoTo Failed
    
    Dim UseTrans As Boolean, UseKey As Boolean

    UseTrans = IIf(alpha = 255, False, True)
    UseKey = IIf(Tex.getColKey = 0, False, True)
    gD3DDevice.SetVertexShader TL_FVF
    If UseTrans Or UseKey Then
        'A blending operation is required. So, enable alpha blending.
        gD3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
        If UseTrans And UseKey Then
            gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        ElseIf UseTrans Then
            gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        Else
            gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2
        End If
    End If
    gD3DDevice.SetTexture 0, Tex.getpTexture
    gD3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vert(0), Len(Vert(0))
    gD3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
Failed:
End Sub
