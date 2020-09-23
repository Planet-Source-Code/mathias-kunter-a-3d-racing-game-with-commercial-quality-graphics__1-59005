Attribute VB_Name = "mdlMain"
Option Explicit


'API declarations

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Const DEFAULT_CHARSET As Long = 1
Public Const OUT_DEFAULT_PRECIS As Long = 0
Public Const CLIP_DEFAULT_PRECIS As Long = 0
Public Const PROOF_QUALITY As Long = 2
Public Const DEFAULT_PITCH As Long = 0

'Revo 3d engine relevant declarations

Public gEngine As cls3dEngine                               'Pointer to the main 3d engine object
Public gDX As DirectX8                                      'Pointer to DirectX8
Public gD3DX As D3DX8                                       'Pointer to the Direct3d helper class
Public gD3D As Direct3D8                                    'Pointer to Direct3d
Public gD3DDevice As Direct3DDevice8                        'Pointer to the rendering device

Public Const PI As Single = 3.1415927


'DirectX relevant declarations

Public Const TL_FVF As Long = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
Public Const VERTEX_FVF As Long = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1


Public Type TLVertex
    x As Single         'x screen pos
    y As Single         'y screen pos
    z As Single         'z buffer depth
    rhw As Single
    Color As Long
    tu As Single        'x texture coordinate
    tv As Single        'y texture coordinate
End Type

Public Type Vertex
    x As Single
    y As Single
    z As Single
    nx As Single
    ny As Single
    nz As Single
    tu As Single
    tv As Single
End Type

Public Type Revo3dRes
    ResX As Long
    ResY As Long
    dxFormat As CONST_D3DFORMAT
End Type

Public Type Revo3dCaps
    MaxAntiAlias As Long
    MaxTextureSize As D3DVECTOR2
    AnisotropicAviable As Boolean
    FogAviable As Boolean
    ShadowAviable As Boolean
    ShadowNoTransparencyAviable As Boolean
    SpecularAviable As Boolean
End Type

Public Type Revo3dPlane
    Points(2) As D3DVECTOR
    NormalVec As D3DVECTOR
    Plane As D3DPLANE
End Type

Public Type RegObj
    pObj As Object
    ObjType As ObjType
End Type

Public Enum PolyType
    POINTLIST = 1
    LINELIST = 2
    LINESTRIP = 3
    TRIANGLELIST = 4
    TRIANGLESTRIP = 5
    TRIANGLEFAN = 6
End Enum

Public Enum PolyVisibility
    NOCHANGE = 0
    BOTHSIDES = 1
    BACKSIDE = 2
    FRONTSIDE = 3
End Enum

Public Enum LightType
    PointLight = 1
    SPOTLIGHT = 2
    DIRECTIONALLIGHT = 3
End Enum

Public Enum ObjType
    OBJ_2DTEXTURE = 0
    OBJ_2DTEXT = 1
    OBJ_3DOBJ = 2
    OBJ_3DPOLYS = 3
End Enum





'Revo 3d engine functions.
Public Function RotationMatrixMake(Angle As D3DVECTOR, Order As D3DVECTOR) As D3DMATRIX
    Dim i As Long, j As Long
    Dim RMat As D3DMATRIX, Rotation As D3DMATRIX

    D3DXMatrixIdentity Rotation
    For i = 0 To 2
        If i = Order.x Then
            j = 0
        ElseIf i = Order.y Then
            j = 1
        ElseIf i = Order.z Then
            j = 2
        End If
        'j is now the axis which should be rotated.
        If j = 0 And Not Angle.x = 0 Then
            D3DXMatrixRotationX RMat, Angle.x
            D3DXMatrixMultiply Rotation, Rotation, RMat
        ElseIf j = 1 And Not Angle.y = 0 Then
            D3DXMatrixRotationY RMat, Angle.y
            D3DXMatrixMultiply Rotation, Rotation, RMat
        ElseIf j = 2 And Not Angle.z = 0 Then
            D3DXMatrixRotationZ RMat, Angle.z
            D3DXMatrixMultiply Rotation, Rotation, RMat
        End If
    Next i
    RotationMatrixMake = Rotation
End Function

Public Function MirrorPoint(ByRef pPlane As Revo3dPlane, ByRef pPoint As D3DVECTOR) As D3DVECTOR
    Dim OpVec As D3DVECTOR, SpPlane As D3DPLANE, SpVec As D3DVECTOR

    'Mirrors a given point to the other side of a plane.
    D3DXVec3Add OpVec, pPoint, pPlane.NormalVec
    D3DXPlaneIntersectLine SpPlane, pPlane.Plane, pPoint, OpVec
    SpVec.x = SpPlane.a
    SpVec.y = SpPlane.b
    SpVec.z = SpPlane.C
    D3DXVec3Subtract OpVec, SpVec, pPoint
    D3DXVec3Add OpVec, SpVec, OpVec
    MirrorPoint = OpVec
End Function

Public Function GetFolder(ByVal File As String) As String
    Dim i As Long, l As Long

    l = Len(File)
    For i = l To 1 Step -1
        If Mid$(File, i, 1) = "\" Then
            GetFolder = Left$(File, i)
            Exit Function
        End If
    Next i
End Function

'Not implemented in VB.

'GUID GuidMake (unsigned int Identifier)
'{
'    int i;
'    GUID retGuid;
'
'    retGuid.Data1 = (DWORD) Identifier;
'    retGuid.Data2 = (WORD) Identifier;
'    retGuid.Data3 = (WORD) Identifier;
'    for (i = 0; i < 8; i++)
'        retGuid.Data4[i] = (BYTE) Identifier;
'    return (retGuid);
'}
'
'
'
'WCHAR *UnicodeStringMake (char *String)
'{
'    WCHAR *wStr;
'
'    wStr = new WCHAR[strlen (String) + 1];
'    MultiByteToWideChar (CP_ACP, 0, String, -1, wStr, strlen (String) + 1);
'    return (wStr);
'}
'
'char *UnicodeStringTranslate (WCHAR *UnicodeString)
'{
'    char *Str;
'
'    Str = new char[wcslen (UnicodeString) + 1];
'    WideCharToMultiByte (CP_ACP, 0, UnicodeString, -1, Str, wcslen (UnicodeString) + 1, NULL, NULL);
'    return (Str);
'}

'//-----------------------------------------------------------------------------
'// Function: Vector2dMake
'// Desc: Creates a 2d vector.
'// Param: x, y
'//-----------------------------------------------------------------------------
Public Function Vector2dMake(ByVal x As Single, ByVal y As Single) As D3DVECTOR2
    Vector2dMake.x = x
    Vector2dMake.y = y
End Function

'//-----------------------------------------------------------------------------
'// Function: NormalVector2dMake
'// Desc: Creates a 2d normal vector from the v1->v2 vector.
'// Param: v1, v2
'//-----------------------------------------------------------------------------
Public Function NormalVector2dMake(ByRef v1 As D3DVECTOR2, ByRef v2 As D3DVECTOR2) As D3DVECTOR2
    Dim ErgVec As D3DVECTOR2, retVec As D3DVECTOR2

    D3DXVec2Subtract ErgVec, v2, v1
    retVec.x = ErgVec.y
    retVec.y = -ErgVec.x
    D3DXVec2Normalize retVec, retVec
    NormalVector2dMake = retVec
End Function

'//-----------------------------------------------------------------------------
'// Function: Vector2dRotate
'// Desc: Rotates a 2d vector.
'// Param: Vector (the vector which should be rotated), Angle (the angle for the
'// rotation), Origin (the origin of the rotation)
'//-----------------------------------------------------------------------------
Public Function Vector2dRotate(ByRef Vector As D3DVECTOR2, ByVal Angle As Single, ByRef Origin As D3DVECTOR2) As D3DVECTOR2
    Dim cosPhi As Single, sinPhi As Single

    cosPhi = Cos(Angle)
    sinPhi = Sin(Angle)
    Vector2dRotate.x = (Vector.x - Origin.x) * cosPhi - (Vector.y - Origin.y) * sinPhi + Origin.x
    Vector2dRotate.y = (Vector.x - Origin.x) * sinPhi + (Vector.y - Origin.y) * cosPhi + Origin.y
End Function

'//-----------------------------------------------------------------------------
'// Funkcion: Vector3dMake
'// Desc: Creates a 3d vector.
'// Parameter: x, y, z
'//-----------------------------------------------------------------------------
Public Function Vector3dMake(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
    With Vector3dMake
        .x = x
        .y = y
        .z = z
    End With
End Function

'//-----------------------------------------------------------------------------
'// Function: NormalVector3dMake
'// Desc: Creates a 3d normal vector from the given points v1, v2, v3
'// Param: v1, v2, v3
'//-----------------------------------------------------------------------------
Public Function NormalVector3dMake(ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR, ByRef v3 As D3DVECTOR) As D3DVECTOR
    Dim retVec As D3DVECTOR, VecA As D3DVECTOR, VecB As D3DVECTOR

    D3DXVec3Subtract VecA, v2, v1
    D3DXVec3Subtract VecB, v3, v2
    D3DXVec3Cross retVec, VecA, VecB
    D3DXVec3Normalize retVec, retVec
    NormalVector3dMake = retVec
End Function

'//-----------------------------------------------------------------------------
'// Function: Rotate
'// Desc: Rotates a 3d vector.
'// Param: Vector (the vector which should be rotated), Angle (angle for x, y and z rotation),
'// Order (the rotating order, pass a vector which contains numbers from 0 to 2,
'// eg. x = 1, y = 0, z = 2), Origin
'// Note: This function is slow. Don't use it to rotate many vectors per frame. Use the
'// rotating functions of 3d polygons and objects instead.
'//-----------------------------------------------------------------------------
Public Function Vector3dRotate(ByRef Vector As D3DVECTOR, ByRef Angle As D3DVECTOR, ByRef Order As D3DVECTOR, ByRef Origin As D3DVECTOR) As D3DVECTOR
    Dim RMat As D3DMATRIX, PMat As D3DMATRIX, DMat As D3DMATRIX
    Dim retVec As D3DVECTOR

    'Create a rotation matrix.
    RMat = RotationMatrixMake(Angle, Order)

    'Rotate the vector.
    D3DXMatrixIdentity PMat
    PMat.m41 = Vector.x - Origin.x
    PMat.m42 = Vector.y - Origin.y
    PMat.m43 = Vector.z - Origin.z
    D3DXMatrixMultiply DMat, PMat, RMat
    retVec.x = DMat.m41 + Origin.x
    retVec.y = DMat.m42 + Origin.y
    retVec.z = DMat.m43 + Origin.z

    Vector3dRotate = retVec
End Function

'//-----------------------------------------------------------------------------
'// Function: PlaneMake
'// Desc: Creates a 3d plane from the 3 given points.
'// Param: p1, p2, p3
'//-----------------------------------------------------------------------------
Public Function PlaneMake(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByRef p3 As D3DVECTOR) As Revo3dPlane
    PlaneMake.Points(0) = p1
    PlaneMake.Points(1) = p2
    PlaneMake.Points(2) = p3
    D3DXPlaneFromPoints PlaneMake.Plane, p1, p2, p3
    PlaneMake.NormalVec = NormalVector3dMake(p1, p2, p3)
End Function

'//-----------------------------------------------------------------------------
'// Function: VertexMake
'// Desc: Creates a vertex.
'// Param: x, y, z, nx, ny, nz, tu, tv
'//-----------------------------------------------------------------------------
Public Function VertexMake(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal nx As Single, ByVal ny As Single, ByVal nz As Single, ByVal tu As Single, ByVal tv As Single) As Vertex
    With VertexMake
        .x = x
        .y = y
        .z = z
        .nx = nx
        .ny = ny
        .nz = nz
        .tu = tu
        .tv = tv
    End With
End Function

'//-----------------------------------------------------------------------------
'// Function: TLVertexMake
'// Desc: Creates a TL-Vertex (transformed & lit vertex)
'// Parameter: x, y, tu, tv
'//-----------------------------------------------------------------------------
Public Function TLVertexMake(ByVal x As Single, ByVal y As Single, ByVal Color As Long, ByVal tu As Single, ByVal tv As Single) As TLVertex
    With TLVertexMake
        .x = x
        .y = y
        .z = 0
        .rhw = 1
        .Color = Color
        .tu = tu
        .tv = tv
    End With
End Function

'//-----------------------------------------------------------------------------
'// Function: BillboardVertexMake
'// Desc: Creates vertices that are billboarded to the camera position.
'// Param: MidPos (the center of the billboard, around which it is rotated),
'// Width (the width of the billboard), Height (the height of the billboard), CamPos (the
'// camera position), AllowXRotation (If the billboard is also allowed to rotate around the x axis.
'// If not, the billboard only rotates around the y axis. This can be used for trees, for example.),
'// outV (pointer to a field of at least 4 vertices).
'// Note: The returned vertices have to be rendered as triangle strip with 2 polygons.
'//-----------------------------------------------------------------------------

'//*****************************************************************************
'//*****************************************************************************
'// This function took me hours and even days to write. I think it's the fastest
'// billboarding algorithm possible.
'//
'// Only 3 SQR calls, NOT A SINGLE CALL of the slow sin or cos functions!
'// This is WORST case, so go and try to write a faster code :)
'//*****************************************************************************
'//*****************************************************************************

Public Sub BillboardVertexMake(ByRef MidPos As D3DVECTOR, ByVal Width As Single, ByVal Height As Single, ByRef CamPos As D3DVECTOR, ByVal AllowXRotation As Boolean, ByVal outV As Long)
    On Local Error GoTo Failed
    
    If MidPos.x = CamPos.x And MidPos.y = CamPos.y And MidPos.z = CamPos.z Then Exit Sub

    Dim i As Long
    Dim Delta As D3DVECTOR, DeltaCube As D3DVECTOR, NVecW As D3DVECTOR, NVecH As D3DVECTOR
    Dim Corner(3) As D3DVECTOR, tmpVertex(3) As Vertex
    Dim DeltaCubeSum As Single, HypSum As Single, sinx As Single, cosx As Single, siny As Single, cosy As Single

    'Compute delta.
    D3DXVec3Subtract Delta, MidPos, CamPos
    DeltaCube.x = Delta.x * Delta.x
    DeltaCube.y = Delta.y * Delta.y
    DeltaCube.z = Delta.z * Delta.z
    DeltaCubeSum = DeltaCube.x + DeltaCube.y + DeltaCube.z
    
    'Compute sin and cos of both rotating angles. The trick is, do this
    'without using the slow sin and cos functions!
    If AllowXRotation Then
        sinx = Delta.y / Sqr(DeltaCubeSum)
        cosx = Sqr(1 - DeltaCube.y / DeltaCubeSum)
    End If
    If Not DeltaCube.x + DeltaCube.z = 0 Then
        HypSum = Sqr(DeltaCube.x + DeltaCube.z)
        siny = Delta.x / HypSum
        cosy = Delta.z / HypSum
    Else
        siny = 0
        cosy = 1
    End If
    
    'create normal vector width (NVecW)
    NVecW = Vector3dMake(cosy, 0, -siny)
    D3DXVec3Scale NVecW, NVecW, Width / 2
    'create normal vector height (NVecH)
    If AllowXRotation Then
        NVecH = Vector3dMake(-sinx * siny, cosx, -sinx * cosy)
        D3DXVec3Scale NVecH, NVecH, Height / 2
    Else
        NVecH = Vector3dMake(0, Height / 2, 0)
    End If
    
    'Compute final points.
    D3DXVec3Subtract Corner(0), MidPos, NVecW
    D3DXVec3Add Corner(0), Corner(0), NVecH
    D3DXVec3Add Corner(1), MidPos, NVecW
    D3DXVec3Add Corner(1), Corner(1), NVecH
    D3DXVec3Subtract Corner(2), MidPos, NVecW
    D3DXVec3Subtract Corner(2), Corner(2), NVecH
    D3DXVec3Add Corner(3), MidPos, NVecW
    D3DXVec3Subtract Corner(3), Corner(3), NVecH
    
    D3DXVec3Subtract Delta, CamPos, MidPos
    D3DXVec3Normalize Delta, Delta
    
    'Write them to the vertices.
    For i = 0 To 3
        tmpVertex(i).x = Corner(i).x
        tmpVertex(i).y = Corner(i).y
        tmpVertex(i).z = Corner(i).z
        tmpVertex(i).nx = Delta.x
        tmpVertex(i).ny = Delta.y
        tmpVertex(i).nz = Delta.z
    Next i
    tmpVertex(1).tu = 1
    tmpVertex(2).tv = 1
    tmpVertex(3).tu = 1
    tmpVertex(3).tv = 1
    
    'Copy the results back.
    CopyMemory ByVal outV, tmpVertex(0), Len(tmpVertex(0)) * 4
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: ColorMake
'// Desc: Creates a color without transparency.
'// Param: r, g, b
'//-----------------------------------------------------------------------------
Public Function ColorMake(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte) As Long
    ColorMake = D3DColorXRGB(r, g, b)
End Function

'//-----------------------------------------------------------------------------
'// Function: ColorAlphaMake
'// Desc: Creates a color with transparency.
'// Param: a (the transparency), r, g, b
'//-----------------------------------------------------------------------------
Public Function ColorAlphaMake(ByVal a As Byte, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte) As Long
    ColorAlphaMake = D3DColorARGB(a, r, g, b)
End Function

'//-----------------------------------------------------------------------------
'// Funkcion: TextureMake
'// Desc: Creates a texture from a bitmap file.
'// Param: see cls2dTexture::LoadFromFile
'//-----------------------------------------------------------------------------
Public Function TextureMake(File As String, ByVal ColorKey As Long, ByVal MipLevels As Long, ByVal EnableEdit As Boolean) As Direct3DTexture8
    On Local Error GoTo Failed
    
    Dim TexFormat As CONST_D3DFORMAT

    If EnableEdit Then
        TexFormat = D3DFMT_A8R8G8B8
    ElseIf ColorKey Then
        TexFormat = D3DFMT_A1R5G5B5
    Else
        TexFormat = D3DFMT_R5G6B5
    End If
    Set TextureMake = gD3DX.CreateTextureFromFileEx(gD3DDevice, File, D3DX_DEFAULT, D3DX_DEFAULT, MipLevels, 0, TexFormat, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKey, ByVal 0, ByVal 0)
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: MaterialMake
'// Desc: Creates a material.
'// Param: rReflect, gReflect, bReflect (the rgb values of the reflective color),
'// rOwn, gOwn, bOwn (the rgb values of the material itself), SpecPower (the power
'// of specular lights on this material), Transparency
'// Note: All parameters are from 0 to 255.
'//-----------------------------------------------------------------------------
Public Function MaterialMake(ByVal rReflect As Byte, ByVal gReflect As Byte, ByVal bReflect As Byte, ByVal rOwn As Byte, ByVal gOwn As Byte, ByVal bOwn As Byte, ByVal SpecPower As Single, ByVal Transparency As Byte) As D3DMATERIAL8
    Dim retMat As D3DMATERIAL8
    Dim Reflect As D3DCOLORVALUE, Own As D3DCOLORVALUE

    Reflect.a = 1 - Transparency / 255
    Reflect.r = rReflect / 255
    Reflect.g = gReflect / 255
    Reflect.b = bReflect / 255
    Own.a = 1
    Own.r = rOwn / 255
    Own.g = gOwn / 255
    Own.b = bOwn / 255

    retMat.Ambient = Reflect
    retMat.Ambient.a = 1
    retMat.diffuse = Reflect
    retMat.emissive = Own
    retMat.Specular.a = 1
    retMat.Specular.r = 1
    retMat.Specular.g = 1
    retMat.Specular.b = 1
    retMat.power = SpecPower

    MaterialMake = retMat
End Function

'//-----------------------------------------------------------------------------
'// Function: LightMake
'// Desc: Creates a light source.
'// Param: lType (the light type), rEmit, gEmit, bEmit (the color of the light),
'// rSpecular, gSpecular, bSpecular (the color of the specular light), Position
'// (the position of the light source), Direction (the direction of the emitted light),
'// Range (the maximum range of the light), Attenuation (the reduction of light with
'// increasing distance, pass 1 for default value), InnerCore, OuterCore (the two angles
'// for the inner and outer core of the beam).
'//-----------------------------------------------------------------------------
Public Function LightMake(ByVal lType As CONST_D3DLIGHTTYPE, ByVal rEmit As Byte, ByVal gEmit As Byte, ByVal bEmit As Byte, ByVal rSpecular As Byte, ByVal gSpecular As Byte, ByVal bSpecular As Byte, ByRef Position As D3DVECTOR, ByRef Direction As D3DVECTOR, ByVal Range As Single, ByVal Attenuation As Single, ByVal InnerCore As Single, ByVal OuterCore As Single) As D3DLIGHT8
    Dim retLight As D3DLIGHT8
    Dim Emit As D3DCOLORVALUE, Spec As D3DCOLORVALUE

    Emit.a = 1
    Emit.r = rEmit
    Emit.g = gEmit
    Emit.b = bEmit
    Spec.a = 1
    Spec.r = rSpecular
    Spec.g = gSpecular
    Spec.b = bSpecular

    retLight.Type = lType
    retLight.Ambient = Emit
    retLight.diffuse = Emit
    retLight.Specular = Spec
    retLight.Position = Position
    retLight.Direction = Direction
    retLight.Range = Range
    retLight.Falloff = 1
    retLight.Attenuation1 = Attenuation
    retLight.Theta = InnerCore
    retLight.Phi = OuterCore
    
    LightMake = retLight
End Function

'//-----------------------------------------------------------------------------
'// Function: PointLightMake
'// Desc: Creates a point light which lights in every direction.
'// Param: see LightMake
'//-----------------------------------------------------------------------------
Public Function PointLightMake(ByVal rEmit As Byte, ByVal gEmit As Byte, ByVal bEmit As Byte, ByVal Specular As Byte, ByRef Position As D3DVECTOR, ByVal Range As Single) As D3DLIGHT8
    PointLightMake = LightMake(D3DLIGHT_POINT, rEmit, gEmit, bEmit, Specular, Specular, Specular, Position, Vector3dMake(0, 0, 0), Range, 1, 0, 0)
End Function

'//-----------------------------------------------------------------------------
'// Function: SpotLightMake
'// Desc: Creates a light which lights in a specified direction with a beam.
'// Param: see LightMake
'//-----------------------------------------------------------------------------
Public Function SpotLightMake(ByVal rEmit As Byte, ByVal gEmit As Byte, ByVal bEmit As Byte, ByVal Specular As Byte, ByRef Position As D3DVECTOR, ByRef Direction As D3DVECTOR, ByVal Range As Single, ByVal InnerCore As Single, ByVal OuterCore As Single) As D3DLIGHT8
    SpotLightMake = LightMake(D3DLIGHT_SPOT, rEmit, gEmit, bEmit, Specular, Specular, Specular, Position, Direction, Range, 1, InnerCore, OuterCore)
End Function

'//-----------------------------------------------------------------------------
'// Function: DirectionalLightMake
'// Desc: Creates a light which lights from a given direction to the whole 3d world. You can
'// compare this to sunlight, for example.
'// Param: see LightMake
'//-----------------------------------------------------------------------------
Public Function DirectionalLightMake(ByVal rEmit As Byte, ByVal gEmit As Byte, ByVal bEmit As Byte, ByVal Specular As Byte, ByRef Direction As D3DVECTOR) As D3DLIGHT8
    DirectionalLightMake = LightMake(D3DLIGHT_DIRECTIONAL, rEmit, gEmit, bEmit, Specular, Specular, Specular, Vector3dMake(0, 0, 0), Direction, 0, 0, 0, 0)
End Function

'//-----------------------------------------------------------------------------
'// Function: ResolutionMake
'// Desc: Creates a Revo 3d resolution.
'// Param: x, y (the x and y resolution), bpp (either 16 or 32)
'//-----------------------------------------------------------------------------
Public Function ResolutionMake(ByVal x As Long, ByVal y As Long, ByVal bpp As Long) As Revo3dRes
    ResolutionMake.ResX = x
    ResolutionMake.ResY = y
    If Not bpp = 32 Then
        ResolutionMake.dxFormat = D3DFMT_R5G6B5
    Else
        ResolutionMake.dxFormat = D3DFMT_X8R8G8B8
    End If
End Function
