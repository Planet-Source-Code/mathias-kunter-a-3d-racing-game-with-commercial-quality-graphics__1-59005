VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls3dEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'*******************************************************************************
'*******************************************************************************
'
'REVO 3d ENGINE
'
'Here we go again...
'
'I decided to port my 3d engine from C++ to Visual Basic. Enjoy!
'This is for all the people who still believe a 3d engine in VB isn't possible.
'It uses DirectX, of course, but every 3d game does.
'
'You might use this 3d engine for your projects, but please give me credits if
'you do so.
'
'Copyright by Mathias Kunter. You can mail me at mathiaskunter@yahoo.de
'
'*******************************************************************************
'*******************************************************************************



'Variables
Private hWnd As Long
Private TimerFrq As LARGE_INTEGER, TimerBase As LARGE_INTEGER
Private TimerLast As Single
Private ObjCleaned As Boolean
Private RegObjCnt As Long
Private RegisterObj() As RegObj

'Direct3D variables
Private DefaultMaterial As D3DMATERIAL8

Private AdapterUse As Long                                  'Number of the used display adapter (3d graphics card)
Private Res As Revo3dRes                                    'Resolution
Private AntiAlias As Long                                   'Anti-aliasing mode
Private Windowed As Boolean
Private MaxLightIndex As Long                               'Highest used lightning index
Private DevType As CONST_D3DDEVTYPE                         'Hardware or software rendering
Private TLType As CONST_D3DCREATEFLAGS                      'Hardware or software transforming and lightning
Private RdCaps As D3DCAPS8                                  'Stores the capabilities of the graphics card
Private PParam As D3DPRESENT_PARAMETERS
Private LT As D3DVECTOR2                                    'Left top corner of the viewport
Private RB As D3DVECTOR2                                    'Right bottom corner of the viewport
Private CamPos As D3DVECTOR

Private ShadowAble As Boolean                               'If shadows can be used
Private ShadowNoBlendAble As Boolean                        'If non-blended shadows can be used
Private UseShadow As Boolean
Private matShadow As D3DMATRIX

Private UseMirror As Boolean
Private matViewSave As D3DMATRIX

Private lpRes() As Revo3dRes






Private Sub Class_Initialize()
    Set gEngine = Me
End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

'//-----------------------------------------------------------------------------
'// Function: Initialize
'// Desc: Initializes the Revo3d engine.
'// Param: phWnd (defines the window which should be used)
'//-----------------------------------------------------------------------------
Public Function Initialize(ByVal phWnd As Long) As Boolean
    On Local Error GoTo Failed
    
    Dim i As Long, cntAdapter As Long
    Dim HWTL As Boolean
        
    'Init variables
    Cleanup
    AdapterUse = -1
    HWTL = False
    ShadowAble = False
    ShadowNoBlendAble = False
    UseShadow = False
    UseMirror = False
    DefaultMaterial = MaterialMake(255, 255, 255, 0, 0, 0, 0, 0)
    ObjCleaned = False
    MaxLightIndex = -1

    hWnd = phWnd

    '//********************DIRECT X********************
    Set gDX = New DirectX8
    Set gD3DX = New D3DX8
    Set gD3D = gDX.Direct3DCreate

    cntAdapter = gD3D.GetAdapterCount
    For i = 0 To cntAdapter - 1
        'Get best possible rendering settings (hardware rendering, hardware T&L)
        gD3D.GetDeviceCaps i, D3DDEVTYPE_HAL, RdCaps
        If (RdCaps.PresentationIntervals And D3DPRESENT_INTERVAL_IMMEDIATE) Or (RdCaps.PresentationIntervals And D3DPRESENT_INTERVAL_ONE) Then
            If Not HWTL Or (RdCaps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT) Then
                'This graphics card can be used.
                If RdCaps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT Then HWTL = True
                AdapterUse = i
            End If
        End If
    Next i
    If AdapterUse = -1 Then GoTo Failed

    'Set hardware T&L, if possible
    If HWTL Then
        TLType = D3DCREATE_HARDWARE_VERTEXPROCESSING
    Else
        TLType = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    End If
    
    Initialize = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: Cleanup
'// Desc: Terminates the Revo3dEngine.
'//-----------------------------------------------------------------------------
Public Sub Cleanup()
    'CleanupObjects
    Set gD3DDevice = Nothing
    Set gD3D = Nothing
    Set gD3DX = Nothing
    Set gDX = Nothing
End Sub

'Returns the default material.
Public Function GetDefaultMaterial() As D3DMATERIAL8
    GetDefaultMaterial = DefaultMaterial
End Function

'//-----------------------------------------------------------------------------
'// Function: GetPossibleRes
'// Desc: Gets all possible resolutions.
'// Param: AllowBelow640 (if resolutions < 640x480 are allowed),
'// AllowUnusualRes (if not usual resolutions are allowed),
'// Allow32bpp (if 32 bit color depth is allowed), pOutRes (returned pointer to a Revo3dRes field)
'// Returns: Number of matching resolutions
'//-----------------------------------------------------------------------------
Public Function GetPossibleRes(ByVal AllowBelow640 As Boolean, ByVal AllowUnusualRes As Boolean, ByVal Allow32bpp As Boolean, ByRef pOutRes As Long) As Long
    On Local Error GoTo Failed
    
    Dim i As Long, j As Long, cntMode As Long, cntAct As Long
    Dim ExAlr As Boolean, RdMode As D3DDISPLAYMODE

    cntMode = gD3D.GetAdapterModeCount(AdapterUse)
    For i = 0 To cntMode - 1
        'Loop through all possible resolutions.
        gD3D.EnumAdapterModes AdapterUse, i, RdMode
        If AllowBelow640 Or Not (RdMode.Width < 640 Or RdMode.Height < 480) Then
            If AllowUnusualRes Or IsUsualRes(RdMode.Width, RdMode.Height) Then
                If RdMode.Format = D3DFMT_R5G6B5 Or RdMode.Format = D3DFMT_X8R8G8B8 Then
                    'Check if this resolution exists already.
                    ExAlr = False
                    For j = 0 To cntAct - 1
                        If lpRes(j).ResX = RdMode.Width And lpRes(j).ResY = RdMode.Height And lpRes(j).dxFormat = RdMode.Format Then
                            ExAlr = True
                            Exit For
                        End If
                    Next j
                    'If the resolution doesn't exist so far, add it.
                    If Not ExAlr Then
                        ReDim Preserve lpRes(cntAct)
                        With lpRes(cntAct)
                            .ResX = RdMode.Width
                            .ResY = RdMode.Height
                            .dxFormat = RdMode.Format
                        End With
                        cntAct = cntAct + 1
                    End If
                End If
            End If
        End If
    Next i
    pOutRes = VarPtr(lpRes(0))
    GetPossibleRes = cntAct
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: GetPossibleCaps
'// Desc: Returns hardware graphics card capabilities.
'// Param: dxFormat (pixel format), dxWindowed, pOutCaps (pointer to a Revo3dCaps struct)
'// Returns: if failed or not.
'//-----------------------------------------------------------------------------
Public Function GetPossibleCaps(ByVal dxFormat As CONST_D3DFORMAT, ByVal dxWindowed As Boolean, ByVal pOutCaps As Long) As Boolean
    On Local Error GoTo Failed
    
    Dim i As Long
    Dim RdMode As D3DDISPLAYMODE
    Dim BackBufFormat As CONST_D3DFORMAT
    Dim retCaps As Revo3dCaps

    'Detect backbuffer mode
    If dxWindowed Then
        gD3D.GetAdapterDisplayMode AdapterUse, RdMode
        BackBufFormat = RdMode.Format
    Else
        BackBufFormat = dxFormat
    End If
    'Fill in caps struct
    'Anisotropic filtering
    If RdCaps.MaxAnisotropy > 1 Then retCaps.AnisotropicAviable = True
    'fog
    If RdCaps.RasterCaps And D3DPRASTERCAPS_ZFOG Then retCaps.FogAviable = True
    'Anti-aliasing
    i = 2
    Do
        If Not gD3D.CheckDeviceMultiSampleType(AdapterUse, D3DDEVTYPE_HAL, BackBufFormat, dxWindowed, i) = D3D_OK Then
            Exit Do
        End If
        i = i * 2
    Loop While i <= 8
    retCaps.MaxAntiAlias = i / 2
    If retCaps.MaxAntiAlias = 1 Then retCaps.MaxAntiAlias = 0
    'Texture size
    retCaps.MaxTextureSize.x = RdCaps.MaxTextureWidth
    retCaps.MaxTextureSize.y = RdCaps.MaxTextureHeight
    'Shadows
    If RdCaps.RasterCaps And D3DPRASTERCAPS_ZBIAS Then
        If gD3D.CheckDepthStencilMatch(AdapterUse, D3DDEVTYPE_HAL, BackBufFormat, BackBufFormat, D3DFMT_D15S1) = D3D_OK Then retCaps.ShadowAviable = True
        If gD3D.CheckDepthStencilMatch(AdapterUse, D3DDEVTYPE_HAL, BackBufFormat, BackBufFormat, D3DFMT_D24S8) = D3D_OK Then retCaps.ShadowAviable = True
        If gD3D.CheckDepthStencilMatch(AdapterUse, D3DDEVTYPE_HAL, BackBufFormat, BackBufFormat, D3DFMT_D24X4S4) = D3D_OK Then retCaps.ShadowAviable = True
        retCaps.ShadowAviable = retCaps.ShadowAviable And (RdCaps.StencilCaps And D3DSTENCILCAPS_DECR)
        retCaps.ShadowNoTransparencyAviable = True
    End If
    'Specular lights
    If RdCaps.ShadeCaps And D3DPSHADECAPS_SPECULARGOURAUDRGB Then retCaps.SpecularAviable = True
    CopyMemory ByVal pOutCaps, retCaps, Len(retCaps)
    GetPossibleCaps = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: SetRes
'// Desc: Switches to a selected resolution.
'// Param: Res (the resolution), AntiAlias (anti-aliasing), Windowed (if the app runs
'// windowed or fullscreen. The bpp param is ignored in windowed mode.)
'// Note: Sets as default values the viewing distance to 1000 m, the camera pos to
'// (0, 0, 0), the camera destination to (0, 0, 1), the camera roll angle to 0
'// and the texture filter to linear.
'//-----------------------------------------------------------------------------
Public Function SetRes(ByVal pRes As Long, ByVal pAntiAlias As Long, ByVal pWindowed As Boolean) As Boolean
    On Local Error GoTo Failed

    Dim i As Long
    Dim RdRes As Revo3dRes
    Dim PresentParam As D3DPRESENT_PARAMETERS, RdMode As D3DDISPLAYMODE
    Dim BackBufFormat As CONST_D3DFORMAT, StencilFormat As CONST_D3DFORMAT
    Dim RdMatrix(1) As D3DMATRIX
    Dim pRdLight() As D3DLIGHT8, pRdLightEnable() As Long
    Dim RdState(254) As Long, TexState As Long
    Dim Failed As Boolean, PrevSA As Boolean, PrevSBA As Boolean, PrevWindowed As Boolean
    Dim PrevRes As Revo3dRes
    Dim PrevAntiAlias As Long

    CopyMemory RdRes, ByVal pRes, Len(RdRes)
    
    'Initialize present parameters
    PresentParam.BackBufferWidth = RdRes.ResX
    PresentParam.BackBufferHeight = RdRes.ResY
    If pWindowed Then
        PresentParam.Windowed = True
        gD3D.GetAdapterDisplayMode AdapterUse, RdMode
        BackBufFormat = RdMode.Format
    Else
        BackBufFormat = RdRes.dxFormat
        If RdCaps.PresentationIntervals And D3DPRESENT_INTERVAL_IMMEDIATE Then
            PresentParam.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
        ElseIf RdCaps.PresentationIntervals And D3DPRESENT_INTERVAL_ONE Then
            PresentParam.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_ONE
        End If
    End If
    PresentParam.BackBufferFormat = BackBufFormat
    PresentParam.BackBufferCount = 1
    PresentParam.EnableAutoDepthStencil = True
    
    'Test if a stencil buffer is available. Required to render shadows.
    If Not gD3D.CheckDepthStencilMatch(AdapterUse, D3DDEVTYPE_HAL, BackBufFormat, BackBufFormat, D3DFMT_D15S1) = D3D_OK Then
        If Not gD3D.CheckDepthStencilMatch(AdapterUse, D3DDEVTYPE_HAL, BackBufFormat, BackBufFormat, D3DFMT_D24S8) = D3D_OK Then
            If Not gD3D.CheckDepthStencilMatch(AdapterUse, D3DDEVTYPE_HAL, BackBufFormat, BackBufFormat, D3DFMT_D24X4S4) = D3D_OK Then
                'Stencil buffer is not supported.
                StencilFormat = D3DFMT_D16
            Else
                StencilFormat = D3DFMT_D24X4S4
            End If
        Else
            StencilFormat = D3DFMT_D24S8
        End If
    Else
        StencilFormat = D3DFMT_D15S1
    End If
    PrevSA = ShadowAble
    PrevSBA = ShadowNoBlendAble
    If RdCaps.RasterCaps And D3DPRASTERCAPS_ZBIAS Then
        ShadowNoBlendAble = True
        If Not StencilFormat = D3DFMT_D16 Then
            ShadowAble = True
        Else
            ShadowAble = False
        End If
    End If
    PresentParam.AutoDepthStencilFormat = StencilFormat
    PresentParam.SwapEffect = D3DSWAPEFFECT_DISCARD
    PresentParam.MultiSampleType = pAntiAlias
    PresentParam.hDeviceWindow = hWnd
    
    PrevRes = Res
    PrevAntiAlias = AntiAlias
    PrevWindowed = Windowed
    Res = RdRes
    AntiAlias = pAntiAlias
    Windowed = pWindowed

    If gD3DDevice Is Nothing Then
        'Create the device new.
        Set gD3DDevice = gD3D.CreateDevice(AdapterUse, D3DDEVTYPE_HAL, hWnd, TLType, PresentParam)
        PParam = PresentParam
        SetLookDistance 1000
        SetCamera Vector3dMake(0, 0, 0), Vector3dMake(0, 0, 1), 0
        SetAnisotropicState False
        'Set viewport
        LT.x = 0
        LT.y = 0
        RB.x = Res.ResX
        RB.y = Res.ResY

        'Set render states
        gD3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        gD3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        gD3DDevice.SetRenderState D3DRS_FOGCOLOR, ColorMake(150, 150, 150)
        gD3DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_EXP2
        gD3DDevice.SetRenderState D3DRS_STENCILFUNC, D3DCMP_LESS
        gD3DDevice.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_DECR
    Else
        'Reset the device.
        'Store current rendering settings.
        gD3DDevice.GetTransform D3DTS_VIEW, RdMatrix(0)
        gD3DDevice.GetTransform D3DTS_PROJECTION, RdMatrix(1)
        If Not MaxLightIndex = -1 Then
            ReDim pRdLight(MaxLightIndex)
            ReDim pRdLightEnable(MaxLightIndex)
            For i = 0 To MaxLightIndex
                gD3DDevice.GetLight i, pRdLight(i)
                pRdLightEnable(i) = gD3DDevice.GetLightEnable(i)
            Next i
        End If
        For i = 0 To 254
            If IsRenderState(i) Then RdState(i) = gD3DDevice.GetRenderState(i)
        Next i
        TexState = gD3DDevice.GetTextureStageState(0, D3DTSS_MAGFILTER)

        'Cleanup everything, reset, re-create everything.
        gD3DDevice.SetTexture 0, Nothing
        CleanupObjects
        
        On Local Error Resume Next
        Err.Number = 0
        gD3DDevice.Reset PresentParam
        If Not Err.Number = 0 Then
            'Reset failed. Try to use the previous device.
            Err.Number = 0
            gD3DDevice.Reset PParam
            If Not Err.Number = 0 Then Exit Function
            'Restore data from the old device.
            ShadowAble = PrevSA
            ShadowNoBlendAble = PrevSBA
            Res = PrevRes
            AntiAlias = PrevAntiAlias
            Windowed = PrevWindowed
            Failed = True
        Else
            PParam = PresentParam
        End If
        On Local Error GoTo Failed

        RestoreObjects
        'Restore render settings from the previous device.
        gD3DDevice.SetTransform D3DTS_VIEW, RdMatrix(0)
        gD3DDevice.SetTransform D3DTS_PROJECTION, RdMatrix(1)
        If Not MaxLightIndex = -1 Then
            For i = 0 To MaxLightIndex
                gD3DDevice.SetLight i, pRdLight(i)
                gD3DDevice.LightEnable i, pRdLightEnable(i)
            Next i
        End If
        For i = 0 To 254
            If IsRenderState(i) Then gD3DDevice.SetRenderState i, RdState(i)
        Next i
        gD3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, TexState
        gD3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, TexState
    End If
    
    gD3DDevice.ShowCursor False

    gD3DDevice.SetTextureStageState 0, D3DTSS_MAXANISOTROPY, RdCaps.MaxAnisotropy
    gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
    gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_TEXTURE
    
    SetRes = Not Failed
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: GetFormatBPP
'// Desc: Returns the bits per pixel for a given pixel format.
'// Param: Format (the pixel format)
'// Returns: The bits per pixel, or 0 if failed.
'//-----------------------------------------------------------------------------
Public Function GetFormatBPP(ByVal Format As CONST_D3DFORMAT) As Long
    If Format = D3DFMT_R3G3B2 Or Format = D3DFMT_A8 Then
        GetFormatBPP = 8
    ElseIf Format = D3DFMT_R5G6B5 Or Format = D3DFMT_X1R5G5B5 Or Format = D3DFMT_A1R5G5B5 Or Format = D3DFMT_A4R4G4B4 Or Format = D3DFMT_A8R3G3B2 Or Format = D3DFMT_X4R4G4B4 Then
        GetFormatBPP = 16
    ElseIf Format = D3DFMT_R8G8B8 Then
        GetFormatBPP = 24
    ElseIf Format = D3DFMT_A8R8G8B8 Or Format = D3DFMT_X8R8G8B8 Then
        GetFormatBPP = 32
    End If
End Function



'//********** RENDERING **********

'//-----------------------------------------------------------------------------
'// Function: RenderClear
'// Desc: Clears the rendering device.
'// Param: ClearScreen (if the screen should be cleared),
'// ClearColor (clearing color),
'// WillRender3d (if 3d objects are going to be rendered),
'// WillRenderBlendedShadow (if blended shadows are going to be rendered)
'//-----------------------------------------------------------------------------
Public Sub RenderClear(ByVal ClearScreen As Boolean, ByVal ClearColor As Long, ByVal WillRender3d As Boolean, ByVal WillRenderBlendedShadow As Boolean)
    On Local Error GoTo Failed
    
    Dim Flags As CONST_D3DCLEARFLAGS

    If ClearScreen Then Flags = D3DCLEAR_TARGET
    If WillRender3d Then Flags = Flags Or D3DCLEAR_ZBUFFER
    If ShadowAble And WillRenderBlendedShadow Then Flags = Flags Or D3DCLEAR_STENCIL
    gD3DDevice.Clear 0, ByVal 0, Flags, ClearColor, 1, 1
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: RenderStart
'// Desc: Must be called at the beginning of the rendering of a frame.
'// Param: xy1 (the upper left corner of the viewport)
'// xy2 (the lower right corner of the viewport. if [0, 0] is passed, then the entire
'// screen is used)
'//-----------------------------------------------------------------------------
Public Sub RenderStart(ByRef xy1 As D3DVECTOR2, ByRef xy2 As D3DVECTOR2)
    On Local Error GoTo Failed
    
    Dim hr As Long
    Dim vPort As D3DVIEWPORT8, cmpVec(1) As D3DVECTOR2

    Do
        'Reset device, if nescessary
        DoEvents
        hr = gD3DDevice.TestCooperativeLevel
        If hr = D3DERR_DEVICENOTRESET Then
            If SetRes(VarPtr(Res), AntiAlias, Windowed) Then hr = D3D_OK
        End If
    Loop While hr = D3DERR_DEVICENOTRESET

    gD3DDevice.BeginScene

    cmpVec(0) = xy1
    cmpVec(1) = xy2
    If cmpVec(1).x = 0 Then cmpVec(1).x = Res.ResX
    If cmpVec(1).y = 0 Then cmpVec(1).y = Res.ResY
    If Not cmpVec(0).x = LT.x Or Not cmpVec(0).y = LT.y Or Not cmpVec(1).x = RB.x Or Not cmpVec(1).y = RB.y Then
        'Set new viewport
        vPort.x = cmpVec(0).x
        vPort.y = cmpVec(0).y
        vPort.Width = cmpVec(1).x - cmpVec(0).x
        vPort.Height = cmpVec(1).y - cmpVec(0).y
        vPort.MinZ = 0
        vPort.MaxZ = 1
        gD3DDevice.SetViewport vPort
        LT = cmpVec(0)
        RB = cmpVec(1)
    End If
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: RenderEnd
'// Desc: Must be called when rendering of a frame is finished.
'//-----------------------------------------------------------------------------
Public Sub RenderEnd()
    On Local Error Resume Next
    gD3DDevice.EndScene
End Sub

'//-----------------------------------------------------------------------------
'// Funktion: RenderShow
'// Desc: Shows the rendered frame on the screen. Call this function after RenderEnd.
'//-----------------------------------------------------------------------------
Public Sub RenderShow()
    On Local Error Resume Next
    gD3DDevice.Present ByVal 0, ByVal 0, hWnd, ByVal 0
End Sub

'//-----------------------------------------------------------------------------
'// Function: Render2dLine
'// Desc: Renders a 2d line.
'// Param: x1, y1 (first point), x2, y2 (second point), Color (the color)
'//-----------------------------------------------------------------------------
Public Sub Render2dLine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Color As Long)
    On Local Error GoTo Failed
    
    Dim v(1) As TLVertex

    v(0) = TLVertexMake(x1, y1, Color, 0, 0)
    v(1) = TLVertexMake(x2, y2, Color, 0, 0)
    gD3DDevice.SetVertexShader TL_FVF
    gD3DDevice.SetTexture 0, Nothing
    gD3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, v(0), Len(v(0))
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: Render2dBox
'// Desc: Renders a 2d box.
'// Param: x1, y1 (upper left corner), x2, y2 (lower right point),
'// Filled (if filled or not), Color (the color)
'//-----------------------------------------------------------------------------
Public Sub Render2dBox(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Filled As Boolean, ByVal Color As Long)
    On Local Error GoTo Failed
    
    Dim v(4) As TLVertex

    gD3DDevice.SetVertexShader TL_FVF
    gD3DDevice.SetTexture 0, Nothing
    v(0) = TLVertexMake(x1, y1, Color, 0, 0)
    v(1) = TLVertexMake(x2, y1, Color, 0, 0)
    If Filled Then
        v(2) = TLVertexMake(x1, y2, Color, 0, 0)
        v(3) = TLVertexMake(x2, y2, Color, 0, 0)
        gD3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, v(0), Len(v(0))
    Else
        v(2) = TLVertexMake(x2, y2, Color, 0, 0)
        v(3) = TLVertexMake(x1, y2, Color, 0, 0)
        v(4) = v(0)
        gD3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 4, v(0), Len(v(0))
    End If
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetShadowState
'// Desc: Defines if shadows are rendered or not.
'// Param: pPlaneTo (pointer to a Revo3dPlane which is the destination plane for the shadow),
'// LightPos (position of the light which throws the shadow), Intensity (from 0 to 255),
'// Type (the type of the light source), ClipEdges (if the shadows are clipped at the edges
'// of the plane)
'// Returns: Failed or not.
'// Note: Shadows should be rendered least in a frame.
'//-----------------------------------------------------------------------------
Public Function SetShadowState(ByVal pPlaneTo As Long, ByRef LightPos As D3DVECTOR, ByVal Intensity As Long, ByVal lType As LightType, ByVal ClipEdges As Boolean) As Boolean
    On Local Error GoTo Failed

    Dim PlaneTo As Revo3dPlane
    Dim ShadowMat As D3DMATERIAL8
    Dim Light As D3DVECTOR4, ClipPlane As D3DPLANE

    If pPlaneTo = 0 Or Intensity = 0 Or Not ShadowNoBlendAble Or (Not Intensity = 255 And Not ShadowAble) Then
        'Disable shadows.
        UseShadow = False
        gD3DDevice.SetRenderState D3DRS_STENCILENABLE, False
        gD3DDevice.SetRenderState D3DRS_CLIPPLANEENABLE, False
        gD3DDevice.SetRenderState D3DRS_ZBIAS, 0
        gD3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        Exit Function
    End If

    'Enable shadows.
    CopyMemory PlaneTo, ByVal pPlaneTo, Len(PlaneTo)
    UseShadow = True
    ShadowMat = MaterialMake(0, 0, 0, 0, 0, 0, 0, 255 - Intensity)
    'Texture and material
    gD3DDevice.SetTexture 0, Nothing
    gD3DDevice.SetMaterial ShadowMat
    If Not Intensity = 255 Then
        'Blended shadow.
        gD3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        gD3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
        gD3DDevice.SetRenderState D3DRS_STENCILENABLE, True
    Else
        'Non-blended shadow.
        gD3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        gD3DDevice.SetRenderState D3DRS_STENCILENABLE, False
    End If
    'Set light source
    CopyMemory Light, LightPos, Len(LightPos)
    Light.W = IIf(lType = DIRECTIONALLIGHT, 0, 1)
    If ClipEdges Then
        gD3DDevice.SetRenderState D3DRS_CLIPPLANEENABLE, D3DCLIPPLANE0 Or D3DCLIPPLANE1 Or D3DCLIPPLANE2
        D3DXPlaneFromPoints ClipPlane, PlaneTo.Points(1), PlaneTo.Points(0), CamPos
        gD3DDevice.SetClipPlane 0, ClipPlane
        D3DXPlaneFromPoints ClipPlane, PlaneTo.Points(2), PlaneTo.Points(1), CamPos
        gD3DDevice.SetClipPlane 1, ClipPlane
        D3DXPlaneFromPoints ClipPlane, PlaneTo.Points(0), PlaneTo.Points(2), CamPos
        gD3DDevice.SetClipPlane 2, ClipPlane
    Else
        gD3DDevice.SetRenderState D3DRS_CLIPPLANEENABLE, False
    End If
    gD3DDevice.SetRenderState D3DRS_ZBIAS, 5
    D3DXMatrixShadow matShadow, Light, PlaneTo.Plane
    SetShadowState = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: SetReflectionState
'// Desc: Defines if reflexions are rendered or not.
'// Parameter: pPlaneTo (a pointer to a plane where the reflexion occurs),
'// ClipEdges (if the reflexion should be clipped at the edges of this plane)
'//-----------------------------------------------------------------------------
Public Sub SetReflectionState(ByVal pPlaneTo As Long, ByVal ClipEdges As Boolean)
    On Local Error GoTo Failed
    
    Dim PlaneTo As Revo3dPlane, ClipPlane As D3DPLANE
    Dim matReflect As D3DMATRIX, MirrorCamPos As D3DVECTOR
    
    If pPlaneTo = 0 Then
        'Disable reflexions.
        gD3DDevice.SetTransform D3DTS_VIEW, matViewSave
        gD3DDevice.SetRenderState D3DRS_CLIPPLANEENABLE, False
        UseMirror = False
        Exit Sub
    End If
    
    CopyMemory PlaneTo, ByVal pPlaneTo, Len(PlaneTo)
    If Not UseMirror Then
        UseMirror = True
        'Save current view matrix.
        gD3DDevice.GetTransform D3DTS_VIEW, matViewSave
    End If
    D3DXMatrixReflect matReflect, PlaneTo.Plane
    D3DXMatrixMultiply matReflect, matReflect, matViewSave
    gD3DDevice.SetTransform D3DTS_VIEW, matReflect
    If ClipEdges Then
        gD3DDevice.SetRenderState D3DRS_CLIPPLANEENABLE, D3DCLIPPLANE0 Or D3DCLIPPLANE1 Or D3DCLIPPLANE2
        MirrorCamPos = MirrorPoint(PlaneTo, CamPos)
        D3DXPlaneFromPoints ClipPlane, PlaneTo.Points(1), MirrorCamPos, PlaneTo.Points(0)
        gD3DDevice.SetClipPlane 0, ClipPlane
        D3DXPlaneFromPoints ClipPlane, PlaneTo.Points(2), MirrorCamPos, PlaneTo.Points(1)
        gD3DDevice.SetClipPlane 1, ClipPlane
        D3DXPlaneFromPoints ClipPlane, PlaneTo.Points(0), MirrorCamPos, PlaneTo.Points(2)
        gD3DDevice.SetClipPlane 2, ClipPlane
    End If
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetLookDistance
'// Desc: Sets the maximum look distance.
'// Param: MaxView (the maximum look distance in meters), MaxAngle (the viewing angle of the camera)
'//-----------------------------------------------------------------------------
Public Sub SetLookDistance(ByVal MaxView As Single, Optional ByVal MaxAngle As Single = PI / 4)
    On Local Error GoTo Failed
    
    Dim matProj As D3DMATRIX

    D3DXMatrixPerspectiveFovLH matProj, MaxAngle, 1, 1, MaxView
    gD3DDevice.SetTransform D3DTS_PROJECTION, matProj
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetCamera
'// Desc: Sets camera position and orientation.
'// Param: CameraPos, CameraLookAt, RollAngle (rotation of camera)
'//-----------------------------------------------------------------------------
Public Sub SetCamera(ByRef CameraPos As D3DVECTOR, ByRef CameraLookAt As D3DVECTOR, ByVal RollAngle As Single)
    On Local Error GoTo Failed

    Dim matView As D3DMATRIX, matCalc As D3DMATRIX

    CamPos = CameraPos
    D3DXMatrixLookAtLH matView, CameraPos, CameraLookAt, Vector3dMake(0, 1, 0)
    If Not RollAngle = 0 Then
        D3DXMatrixRotationZ matCalc, RollAngle
        D3DXMatrixMultiply matView, matView, matCalc
    End If
    gD3DDevice.SetTransform D3DTS_VIEW, matView
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetLight
'// Desc: Adds a new light source or replaces an old one.
'// Param: Index (the index of the light source), Light (the light source)
'//-----------------------------------------------------------------------------
Public Function SetLight(ByVal Index As Long, ByRef Light As D3DLIGHT8) As Boolean
    On Local Error GoTo Failed
    gD3DDevice.SetLight Index, Light
    gD3DDevice.LightEnable Index, True
    If Index > MaxLightIndex Then MaxLightIndex = Index
    SetLight = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: SetLightState
'// Desc: Switches a light source on or off.
'// Param: Index (the index of the light source), State (the light state)
'//-----------------------------------------------------------------------------
Public Sub SetLightState(ByVal Index As Long, ByVal State As Boolean)
    On Local Error Resume Next
    gD3DDevice.LightEnable Index, State
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetSpecularState
'// Desc: Defines if specular lightning is enabled or disabled. Default: disabled.
'// Param: State
'//-----------------------------------------------------------------------------
Public Sub SetSpecularState(ByVal State As Boolean)
    On Local Error Resume Next
    gD3DDevice.SetRenderState D3DRS_SPECULARENABLE, State
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetZState
'// Desc: Defines how the z buffer is used.
'// Param: ReadEnable (if the z buffer can be read), WriteEnable (if to the z
'// buffer can be written)
'// Note: Setting ReadEnable to false forces all rendered 3d objects to be
'// drawn in foreground. Setting WriteEnable to false causes a rendered object
'// not to be written to the z buffer.
'//-----------------------------------------------------------------------------
Public Sub SetZState(ByVal ReadEnable As Boolean, ByVal WriteEnable As Boolean)
    On Local Error Resume Next
    gD3DDevice.SetRenderState D3DRS_ZENABLE, ReadEnable
    If ReadEnable Then gD3DDevice.SetRenderState D3DRS_ZWRITEENABLE, WriteEnable
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetAnisotropicState
'// Desc: Enables or disables the usage of the anisotropic texture filter.
'// Default: disabled (false).
'// Parameter: State
'//-----------------------------------------------------------------------------
Public Sub SetAnisotropicState(ByVal State As Boolean)
    On Local Error GoTo Failed
    If State Then
        gD3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
        gD3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    Else
        gD3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        gD3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    End If
Failed:
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetGlobalLight
'// Desc: Sets a minimum light.
'// Param: Color (the color of the light)
'//-----------------------------------------------------------------------------
Public Sub SetGlobalLight(ByVal Color As Long)
    On Local Error Resume Next
    gD3DDevice.SetRenderState D3DRS_AMBIENT, Color
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetFog
'// Desc: Defines the density of the fog.
'// Param: Density (from 0 to 255)
'//-----------------------------------------------------------------------------
Public Sub SetFog(ByVal Density As Long)
    On Local Error Resume Next
    
    Dim d As Single

    If Not Density = 0 Then
        gD3DDevice.SetRenderState D3DRS_FOGENABLE, True
        d = Density / 2550
        gD3DDevice.SetRenderState D3DRS_FOGDENSITY, d
    Else
        gD3DDevice.SetRenderState D3DRS_FOGENABLE, False
    End If
End Sub



'//********** GENERAL **********

'//-----------------------------------------------------------------------------
'// Function: StartTimer
'// Desc: Starts the timer.
'//-----------------------------------------------------------------------------
Public Sub StartTimer()
    QueryPerformanceFrequency TimerFrq
    QueryPerformanceCounter TimerBase
    TimerLast = 0
End Sub

'//-----------------------------------------------------------------------------
'// Function: GetRelTimeDiff
'// Desc: Returns the time in seconds since the last call of this function or of
'// the StartTimer function.
'//-----------------------------------------------------------------------------
Public Function GetRelTimeDiff() As Single
    Dim TimerCnt As LARGE_INTEGER
    Dim TimerTotal As Single, TimerDiff As Single

    QueryPerformanceCounter TimerCnt
    TimerTotal = (LargeInt2Currency(TimerCnt) - LargeInt2Currency(TimerBase)) / LargeInt2Currency(TimerFrq)
    TimerDiff = TimerTotal - TimerLast
    TimerLast = TimerTotal
    GetRelTimeDiff = TimerDiff
End Function

'//-----------------------------------------------------------------------------
'// Function: GetAbsTimeDiff
'// Desc: Returns the time in seconds since the last call of the StartTimer function.
'//-----------------------------------------------------------------------------
Public Function GetAbsTimeDiff() As Single
    Dim TimerCnt As LARGE_INTEGER

    QueryPerformanceCounter TimerCnt
    GetAbsTimeDiff = (LargeInt2Currency(TimerCnt) - LargeInt2Currency(TimerBase)) / LargeInt2Currency(TimerFrq)
End Function



'//******************** PRIVATE FUNKTIONEN ********************

Private Function IsUsualRes(ByVal ResX As Long, ByVal ResY As Long) As Boolean
    If ResX = 320 And ResY = 240 Then
        IsUsualRes = True
    ElseIf ResX = 640 And ResY = 480 Then
        IsUsualRes = True
    ElseIf ResX = 800 And ResY = 600 Then
        IsUsualRes = True
    ElseIf ResX = 1024 And ResY = 768 Then
        IsUsualRes = True
    ElseIf ResX = 1280 And ResY = 1024 Then
        IsUsualRes = True
    ElseIf ResX = 1600 And ResY = 1200 Then
        IsUsualRes = True
    ElseIf ResX = 1920 And ResY = 1440 Then
        IsUsualRes = True
    ElseIf ResX = 2048 And ResY = 1536 Then
        IsUsualRes = True
    End If
End Function

Private Function IsRenderState(ByVal State As Long) As Boolean
    If State >= 7 And State <= 10 Then IsRenderState = True
    If State >= 14 And State <= 16 Then IsRenderState = True
    If State = 19 Or State = 20 Or State = 40 Or State = 47 Or State = 48 Then IsRenderState = True
    If State >= 22 And State <= 30 Then IsRenderState = True
    If State >= 34 And State <= 38 Then IsRenderState = True
    If State = 40 Then IsRenderState = True
    If State >= 52 And State <= 60 Then IsRenderState = True
    If State >= 128 And State <= 137 Then IsRenderState = True
    If State >= 139 And State <= 143 Then IsRenderState = True
    If State >= 145 And State <= 148 Then IsRenderState = True
    If State >= 151 And State <= 168 Then IsRenderState = True
    If State = 170 Or State = 171 Then IsRenderState = True
End Function

Private Function LargeInt2Currency(LargeInt As LARGE_INTEGER) As Currency
    CopyMemory LargeInt2Currency, LargeInt, LenB(LargeInt)
    LargeInt2Currency = LargeInt2Currency * 10000
End Function

Friend Sub SetTransform(matTrans As D3DMATRIX, matRot As D3DMATRIX)
    Dim matWorld As D3DMATRIX

    D3DXMatrixMultiply matWorld, matRot, matTrans
    gD3DDevice.SetTransform D3DTS_WORLD, matWorld
End Sub

Friend Sub SetTransformMultiply(matTrans As D3DMATRIX, matRot As D3DMATRIX, matMultiply As D3DMATRIX)
    Dim matWorld As D3DMATRIX

    D3DXMatrixMultiply matWorld, matRot, matTrans
    D3DXMatrixMultiply matWorld, matWorld, matMultiply
    gD3DDevice.SetTransform D3DTS_WORLD, matWorld
End Sub

Friend Function GetUseShadow() As Boolean
    GetUseShadow = UseShadow
End Function

Friend Function GetUseMirror() As Boolean
    GetUseMirror = UseMirror
End Function

Friend Function GetMatShadow() As D3DMATRIX
    GetMatShadow = matShadow
End Function

Public Sub RegisterObject(pObj As Object, ByVal pObjType As ObjType)
    ReDim Preserve RegisterObj(RegObjCnt)
    Set RegisterObj(RegObjCnt).pObj = pObj
    RegisterObj(RegObjCnt).ObjType = pObjType
    RegObjCnt = RegObjCnt + 1
End Sub

Public Sub UnregisterObject(pObj As Object)
    Dim i As Long, j As Long

    For i = 0 To RegObjCnt - 1
        If RegisterObj(i).pObj Is pObj Then
            If Not (gD3DDevice Is Nothing) Then
                If RegisterObj(i).ObjType = OBJ_2DTEXTURE Or RegisterObj(i).ObjType = OBJ_3DOBJ Then
                    gD3DDevice.SetTexture 0, Nothing
                End If
            End If
            Set RegisterObj(i).pObj = Nothing
            For j = i To RegObjCnt - 2
                RegisterObj(j) = RegisterObj(j + 1)
            Next j
            RegObjCnt = RegObjCnt - 1
            Exit Sub
        End If
    Next i
End Sub

Public Sub CleanupObjects()
    If ObjCleaned Then Exit Sub

    Dim i As Long
    Dim pTex As cls2dTexture
    Dim pText As cls2dText
    Dim pObj As cls3dObject
    Dim pPolys As cls3dPolygons

    'Cleanup all created objects.
    For i = 0 To RegObjCnt - 1
        If RegisterObj(i).ObjType = OBJ_2DTEXTURE Then
            Set pTex = RegisterObj(i).pObj
            pTex.StoreData
            pTex.Cleanup
        ElseIf RegisterObj(i).ObjType = OBJ_2DTEXT Then
            Set pText = RegisterObj(i).pObj
            pText.Cleanup
        ElseIf RegisterObj(i).ObjType = OBJ_3DOBJ Then
            Set pObj = RegisterObj(i).pObj
            pObj.StoreData
            'Don't use the cleanup function, it would also delete textures and materials.
            'Just cleanup the mesh itself.
            pObj.CleanupMesh
        ElseIf RegisterObj(i).ObjType = OBJ_3DPOLYS Then
            Set pPolys = RegisterObj(i).pObj
            pPolys.StoreData
            If Not pPolys.getDynamic Then pPolys.Cleanup        'Clean only not dynamic polygons.
        End If
        Set pTex = Nothing
        Set pText = Nothing
        Set pObj = Nothing
        Set pPolys = Nothing
    Next i
    ObjCleaned = True
End Sub

Public Sub RestoreObjects()
    If Not ObjCleaned Then Exit Sub
    
    Dim i As Long
    Dim pTex As cls2dTexture
    Dim pText As cls2dText
    Dim pObj As cls3dObject
    Dim pPolys As cls3dPolygons

    'Restore all cleaned objects.
    For i = 0 To RegObjCnt - 1
        If RegisterObj(i).ObjType = OBJ_2DTEXTURE Then
            Set pTex = RegisterObj(i).pObj
            pTex.ReStoreData
        ElseIf RegisterObj(i).ObjType = OBJ_2DTEXT Then
            Set pText = RegisterObj(i).pObj
            pText.ReSetFont
        ElseIf RegisterObj(i).ObjType = OBJ_3DOBJ Then
            Set pObj = RegisterObj(i).pObj
            pObj.ReStoreData
        ElseIf RegisterObj(i).ObjType = OBJ_3DPOLYS Then
            Set pPolys = RegisterObj(i).pObj
            pPolys.ReStoreData
        End If
        Set pTex = Nothing
        Set pText = Nothing
        Set pObj = Nothing
        Set pPolys = Nothing
    Next i
    ObjCleaned = False
End Sub
