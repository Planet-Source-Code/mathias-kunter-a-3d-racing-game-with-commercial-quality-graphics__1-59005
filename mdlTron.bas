Attribute VB_Name = "mdlTron"
Option Explicit


'Constants

Private Const TRON_MVERS As Long = 1                'Main version number
Private Const TRON_SVERS As Long = 4                'Sub version number

Private Const MAX_MENUITEM As Long = 16
Private Const MAX_MENUTEXT As Long = 64
Private Const MENUITEM_WIDTH  As Long = 550
Private Const MENUITEM_HEIGHT As Long = 40

Private Const MENU_MAIN As Long = 0
'#define MENU_MULTIPLAYER 1                         'Sorry, not implemented in the VB version.
'#define MENU_MULTIPLAYER_SERVER 2
'#define MENU_MULTIPLAYER_CLIENT 3
'#define MENU_MULTIPLAYER_CLIENTWAIT 4
Private Const MENU_OPTIONS As Long = 5
Private Const MENU_OPTIONS_GRAFICS As Long = 6
'#define MENU_goptions_SOUNDS 7
Private Const MENU_OPTIONS_DETAILS As Long = 8
Private Const MENU_OPTIONS_GAMEPLAY As Long = 9


'Types

Public Type TronSettings
    'Graphics
    Res As Revo3dRes
    UseFSAA As Long
    UseReflection As Boolean
    UseShadow As Long
    UseSpecular As Boolean
    UseAnisotropic As Boolean
    UseTransWalls As Boolean
    'Gameplay
    LandSize As Single
    MopedSpeed As Single
    FragLimit As Long
    OpponentCnt As Long
    OpponentSkill As Long
    Cam As Long
    EnableAction As Boolean
    PlayerName As String * 64
End Type

Public Type TronWall
    Pos1 As D3DVECTOR2
    Pos2 As D3DVECTOR2
    From As Long
    Used As Boolean
End Type

Public Type TronFire
    Pos As D3DVECTOR
    Dir As D3DVECTOR
End Type

Public Type TronExplDesc
    Flames(29) As TronFire
    FlameCnt As Long
    Trans As Single
    Used As Boolean
End Type

Public Type TronShotDesc
    Pos As D3DVECTOR2
    Dir As D3DVECTOR2
    NoDanger As Single
    TTL As Long
    ShotType As Long
    Used As Boolean
    pFrom As Long
End Type

Private Type TronSwap
    Ident As Long
    Frags As Long
End Type


'API
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long









'3d Engine
Private Eng3d As cls3dEngine
Private RdCaps As Revo3dCaps
Private ResCnt As Long, ActRes As Long
Private pRes() As Revo3dRes
Private Shield As cls3dObject
Private Inp As clsInput

'TRON game variables
Private PIdent As Long, MopedAlive As Long, LeaderScore As Long
Private CamMode As Long
Private NotRunTime As Single, NotAliveTime As Single, BlendTime As Single, TimeFac As Single
Private GameRuns As Boolean, GamePassed As Boolean

'Public variables
Public gOptions As TronSettings
Public gExplosions As TronExplosion
Public gPlayer() As TronPlayer
Public gShots As TronShots
Public gWallDesc As TronWallDesc





'Sorry for the function name. *g*
'This is the C++ WinMain representative.

Public Sub Run_The_Whole_Damn_Show()
    Dim i As Long
    Dim CursorTex As cls2dTexture, BusyTex As cls2dTexture, ButtonTex As cls2dTexture, BackgrTex As cls2dTexture
    Dim MenuText As cls2dText, FileData() As Byte, InitResult As Long

    On Local Error GoTo Failed

    ReDim FileData(Len(gOptions) - 1)

'    'Create default settings data file.
'    With gOptions
'        .Res = ResolutionMake(640, 480, 32)
'        .UseFSAA = 0
'        .UseReflection = True
'        .UseShadow = 2
'        .UseSpecular = True
'        .UseAnisotropic = False
'        .UseTransWalls = True
'        .LandSize = 200
'        .MopedSpeed = 28
'        .FragLimit = 20
'        .OpponentCnt = 3
'        .OpponentSkill = 2
'        .Cam = 5
'        .EnableAction = True
'        .PlayerName = "Player"
'    End With
'    Open "settings\default.dat" For Binary Access Write As #1
'    CopyMemory FileData(0), gOptions, Len(gOptions)
'    Put #1, , FileData
'    Close #1
'    Exit Sub
    
    'Load settings from file.
    If PathFileExists("settings\config.dat") = 0 Then
        If PathFileExists("settings\default.dat") = 0 Then
            MsgBox "Couldn't load the game. Revo Tron will exit now.", vbCritical, "Revo Tron"
            Exit Sub
        End If
        Open "settings\default.dat" For Binary Access Read As #1
    Else
        Open "settings\config.dat" For Binary Access Read As #1
    End If
    Get #1, , FileData
    CopyMemory gOptions, FileData(0), Len(gOptions)
    Close #1
    
    For i = 1 To Len(gOptions.PlayerName)
        If Mid$(gOptions.PlayerName, i, 1) = " " Then
            gOptions.PlayerName = Left$(gOptions.PlayerName, i)
            Exit For
        End If
    Next i
    
    'Create objects.
    Set Eng3d = New cls3dEngine
    Set Inp = New clsInput
    Set gExplosions = New TronExplosion
    Set gShots = New TronShots
    Set gWallDesc = New TronWallDesc
    Set Shield = New cls3dObject
    
    Set CursorTex = New cls2dTexture
    Set BusyTex = New cls2dTexture
    Set ButtonTex = New cls2dTexture
    Set BackgrTex = New cls2dTexture
    Set MenuText = New cls2dText
    
    gEngine.StartTimer
    Do
    Loop While gEngine.GetAbsTimeDiff < 1
    
    Randomize Timer
    InitResult = OneTimeDXInit(gOptions.Res.ResX, gOptions.Res.ResY, gOptions.Res.dxFormat, gOptions.UseFSAA)
    If InitResult = 0 Then
        MsgBox "Couldn't find an usable 3d graphics card. Revo Tron will exit now.", vbCritical, "Revo Tron"
    ElseIf InitResult = 1 Then
        MsgBox "Couldn't get access to the keyboard. Revo Tron will exit now.", vbCritical, "Revo Tron"
    ElseIf InitResult = 2 Then
        MsgBox "Couldn't change the screen resolution. Revo Tron will exit now.", vbCritical, "Revo Tron"
    End If
    If Not InitResult = 3 Then Exit Sub

    'Load menu objects.
    If Not CursorTex.LoadFromFile("textures\cursor.bmp", ColorMake(0, 0, 0)) Then Exit Sub
    If Not BusyTex.LoadFromFile("textures\busy.bmp", ColorMake(0, 0, 0)) Then Exit Sub
    If Not ButtonTex.LoadFromFile("textures\button.bmp", 0) Then Exit Sub
    BackgrTex.LoadFromFile "textures\background.bmp", 0
    MenuText.SetFormat 1, 1
    MenuText.SetFont "Times New Roman", 20, False, False, False

    ShowMenu MENU_MAIN, Vector2dMake(gOptions.Res.ResX / 2, gOptions.Res.ResY / 2), CursorTex, BusyTex, ButtonTex, MenuText, BackgrTex

    'Destroy objects.
    Set Eng3d = Nothing
    Set Inp = Nothing
    Set gExplosions = Nothing
    Set gShots = Nothing
    Set gWallDesc = Nothing
    Set Shield = Nothing
    
    Set CursorTex = Nothing
    Set BusyTex = Nothing
    Set ButtonTex = Nothing
    Set BackgrTex = Nothing
    Set MenuText = Nothing
    
    'Save settings back to file.
    Open "settings\config.dat" For Binary Access Write As #1
    CopyMemory FileData(0), gOptions, Len(gOptions)
    Put #1, , FileData
    Close #1
Failed:
End Sub


Private Function OneTimeDXInit(ByVal ResX As Long, ByVal ResY As Long, Format As CONST_D3DFORMAT, ByVal AntiAlias As Long) As Long
    Dim i As Long, FoundRes As Boolean, ptrRes As Long

    On Local Error GoTo Failed

    '*****DIRECT 3D*****
    
    If Not Eng3d.Initialize(frmMain.hWnd) Then Exit Function
    
    If Not Eng3d.GetPossibleCaps(Format, False, VarPtr(RdCaps)) Then Exit Function
    If RdCaps.MaxTextureSize.x < 256 Or RdCaps.MaxTextureSize.y < 256 Then Exit Function

    ResCnt = Eng3d.GetPossibleRes(False, False, True, ptrRes)
    If ResCnt = 0 Then Exit Function
    ReDim pRes(ResCnt - 1)
    CopyMemory pRes(0), ByVal ptrRes, LenB(pRes(0)) * ResCnt
    
    'Check if the requested resolution is actually available.
    FoundRes = False
    For i = 0 To ResCnt - 1
        If ResX = pRes(i).ResX And ResY = pRes(i).ResY And Format = pRes(i).dxFormat And AntiAlias <= RdCaps.MaxAntiAlias Then
            FoundRes = True
            Exit For
        End If
    Next i
    If FoundRes Then
        gOptions.Res = pRes(i)
        gOptions.UseFSAA = AntiAlias
        ActRes = i
    Else
        gOptions.Res = pRes(0)
        gOptions.UseFSAA = 0
        ActRes = 0
    End If

    '*****DIRECT INPUT*****
    If Not Inp.Initialize(frmMain.hWnd, True, True) Then
        OneTimeDXInit = 1
        Exit Function
    End If


    'Set the resolution
    'todo debug: no fullscreen
    If Not Eng3d.SetRes(VarPtr(gOptions.Res), gOptions.UseFSAA, False) Then
        Inp.Cleanup
        OneTimeDXInit = 2
        Exit Function
    End If
    
    'Everything OK
    OneTimeDXInit = 3
Failed:
End Function

Private Sub GeometryInit(ByVal SizeX As Single, ByVal SizeZ As Single, ByVal WallHeight As Single, outV() As Vertex)
    Dim tuX As Single, tuZ As Single
    tuX = SizeX \ 75
    tuZ = SizeZ \ 75

    If tuX = 0 Then tuX = 1
    If tuZ = 0 Then tuZ = 1
    'Floor
    outV(0) = VertexMake(0, 0, SizeZ, 0, 1, 0, 0, 0)
    outV(1) = VertexMake(SizeX, 0, SizeZ, 0, 1, 0, SizeX / 5, 0)
    outV(2) = VertexMake(0, 0, 0, 0, 1, 0, 0, SizeZ / 5)
    outV(3) = VertexMake(SizeX, 0, 0, 0, 1, 0, SizeX / 5, SizeZ / 5)
    'Front wall
    outV(4) = VertexMake(SizeX, WallHeight, 0, 0, 0, 1, 0, 0)
    outV(5) = VertexMake(0, WallHeight, 0, 0, 0, 1, tuX, 0)
    outV(6) = VertexMake(SizeX, 0, 0, 0, 0, 1, 0, 0.25)
    outV(7) = VertexMake(0, 0, 0, 0, 0, 1, tuX, 0.25)
    'Back wall
    outV(8) = VertexMake(0, WallHeight, SizeZ, 0, 0, -1, 0, 0.25)
    outV(9) = VertexMake(SizeX, WallHeight, SizeZ, 0, 0, -1, tuX, 0.25)
    outV(10) = VertexMake(0, 0, SizeZ, 0, 0, -1, 0, 0.5)
    outV(11) = VertexMake(SizeX, 0, SizeZ, 0, 0, -1, tuX, 0.5)
    'Left wall
    outV(12) = VertexMake(0, WallHeight, 0, 1, 0, 0, 0, 0.5)
    outV(13) = VertexMake(0, WallHeight, SizeZ, 1, 0, 0, tuZ, 0.5)
    outV(14) = VertexMake(0, 0, 0, 1, 0, 0, 0, 0.75)
    outV(15) = VertexMake(0, 0, SizeZ, 1, 0, 0, tuZ, 0.75)
    'Right wall
    outV(16) = VertexMake(SizeX, WallHeight, SizeZ, -1, 0, 0, 0, 0.5)
    outV(17) = VertexMake(SizeX, WallHeight, 0, -1, 0, 0, tuZ, 0.5)
    outV(18) = VertexMake(SizeX, 0, SizeZ, -1, 0, 0, 0, 0.75)
    outV(19) = VertexMake(SizeX, 0, 0, -1, 0, 0, tuZ, 0.75)
End Sub

Private Sub ShowMenu(ByVal MenuID As Long, CursorPos As D3DVECTOR2, pCurs As cls2dTexture, pBusy As cls2dTexture, pButton As cls2dTexture, pText As cls2dText, pBackgr As cls2dTexture)
    Dim i As Long
    Dim LDown As Boolean, RDown As Boolean, ESCDown As Boolean, RetDown As Boolean
    Dim Busy As Boolean, EditName As Boolean
    Dim ButtonPosL As D3DVECTOR2, ButtonPosR As D3DVECTOR2
    Dim RdMouse As DIMOUSESTATE
    Dim RdKeybPtr As Long, RdKeyb(255) As Byte
    Dim MenuCnt As Long, MenuHl As Long, FadeIn As Long
    Dim MenuStr(MAX_MENUITEM) As String
    Dim MenuCursor As cls2dPicture, MenuButton As cls2dPicture, MenuBackgr As cls2dPicture
    Dim PrevRes As Long, ActFSAA As Long, InpStr As String

    'float WaitTime;
    'TronSettings SaveSettings;

    PrevRes = ActRes
    ActFSAA = gOptions.UseFSAA
    If Trim$(gOptions.PlayerName) = "" Then gOptions.PlayerName = "Player"
    MenuCnt = CreateMenuText(MenuID, MenuStr, pRes(ActRes), ActFSAA, EditName)
    
    Set MenuCursor = New cls2dPicture
    Set MenuButton = New cls2dPicture
    Set MenuBackgr = New cls2dPicture

    MenuButton.SetTransparency 125

    Do While True
        'Update mouse cursor
        RdMouse = Inp.ReadMouse
        If EditName Then
            InpStr = Trim$(gOptions.PlayerName)
            RdKeybPtr = Inp.ReadKeyboard(InpStr)
            If Len(InpStr) > 16 Then InpStr = Left$(InpStr, 16)
            gOptions.PlayerName = InpStr
            CreateMenuText MenuID, MenuStr, pRes(ActRes), ActFSAA, True
        Else
            RdKeybPtr = Inp.ReadKeyboard("")
        End If
        If RdKeybPtr = 0 Then Exit Do
        CopyMemory RdKeyb(0), ByVal RdKeybPtr, 256
        If RdMouse.lx Or RdMouse.ly Then
            LDown = False
            RDown = False
            RetDown = False
        End If
        CursorPos.x = CursorPos.x + RdMouse.lx * 1.5
        CursorPos.y = CursorPos.y + RdMouse.ly * 1.5
        If CursorPos.x < 0 Then CursorPos.x = 0
        If CursorPos.x > gOptions.Res.ResX Then CursorPos.x = gOptions.Res.ResX
        If CursorPos.y < 0 Then CursorPos.y = 0
        If CursorPos.y > gOptions.Res.ResY Then CursorPos.y = gOptions.Res.ResY
        MenuCursor.SetPosition CursorPos.x, CursorPos.y, CursorPos.x + 20, CursorPos.y + 20

        'Determine menu highlighting
        MenuHl = -1
        For i = 0 To MenuCnt - 1
            If CursorPos.x >= gOptions.Res.ResX / 2 - MENUITEM_WIDTH / 2 And CursorPos.x <= gOptions.Res.ResX / 2 + MENUITEM_WIDTH / 2 Then
                If CursorPos.y >= gOptions.Res.ResY / 2 - (MenuCnt / 2 - i) * MENUITEM_HEIGHT And CursorPos.y <= gOptions.Res.ResY / 2 - (MenuCnt / 2 - i - 1) * MENUITEM_HEIGHT Then
                    MenuHl = i
                End If
            End If
        Next i
        'Check if we need to display the sleep cursor...
        If RdKeyb(DIK_ESCAPE) And Not MenuID = MENU_MAIN Then ESCDown = True
        If RdKeyb(DIK_RETURN) Or RdKeyb(DIK_NUMPADENTER) Then RetDown = True
        If RdMouse.Buttons(0) Then LDown = True
        If RdMouse.Buttons(1) Then RDown = True
        If ((RdMouse.Buttons(0) = 0 And LDown) Or (RdMouse.Buttons(1) = 0 And RDown) Or (RdKeyb(DIK_RETURN) = 0 And RdKeyb(DIK_NUMPADENTER) = 0 And RetDown)) And Not MenuHl = -1 Then
            If MenuID = MENU_MAIN And MenuHl = MenuCnt - 1 Then
                Busy = True
            ElseIf MenuID = MENU_MAIN And MenuHl = 0 Then
                Busy = True
            ElseIf MenuID = MENU_OPTIONS_GRAFICS And MenuHl = 3 Then
                Busy = True
            End If
        Else
            Busy = False
        End If

        MenuBackgr.SetPosition 0, 0, gOptions.Res.ResX, gOptions.Res.ResY

        'Render
        Eng3d.RenderStart Vector2dMake(0, 0), Vector2dMake(0, 0)
        Eng3d.RenderClear True, ColorMake(0, 0, 255), False, False
        MenuBackgr.Render pBackgr
        For i = 0 To MenuCnt - 1
            If i = MenuHl Then
                FadeIn = 16
            Else
                FadeIn = 32
            End If
            ButtonPosL.x = gOptions.Res.ResX / 2 - MENUITEM_WIDTH / 2
            ButtonPosL.y = gOptions.Res.ResY / 2 - (MenuCnt / 2 - i) * MENUITEM_HEIGHT
            ButtonPosR.x = ButtonPosL.x + MENUITEM_WIDTH
            ButtonPosR.y = ButtonPosL.y + MENUITEM_HEIGHT

            'Left
            MenuButton.SetPosition ButtonPosL.x, ButtonPosL.y, ButtonPosL.x + FadeIn, ButtonPosR.y
            MenuButton.SetPictureRange 0.01, 0.51, 0.49, 0.99
            MenuButton.Render pButton
            'Center left
            MenuButton.SetPosition ButtonPosL.x + FadeIn, ButtonPosL.y, ButtonPosL.x + FadeIn + 32, ButtonPosR.y
            MenuButton.SetPictureRange 0.01, 0.01, 0.49, 0.49
            MenuButton.Render pButton
            'Center
            MenuButton.SetPosition ButtonPosL.x + FadeIn + 32, ButtonPosL.y, ButtonPosR.x - FadeIn - 32, ButtonPosR.y
            MenuButton.SetPictureRange 0.51, 0.51, 0.99, 0.99
            MenuButton.Render pButton
            'Center right
            MenuButton.SetPosition ButtonPosR.x - FadeIn - 32, ButtonPosL.y, ButtonPosR.x - FadeIn, ButtonPosR.y
            MenuButton.SetPictureRange 0.51, 0.01, 0.99, 0.49
            MenuButton.Render pButton
            'Right
            MenuButton.SetPosition ButtonPosR.x - FadeIn, ButtonPosL.y, ButtonPosR.x, ButtonPosR.y
            MenuButton.SetPictureRange 0.01, 0.51, 0.49, 0.99
            MenuButton.Render pButton
            'Text
            pText.SetPosition ButtonPosL.x, ButtonPosL.y, ButtonPosR.x, ButtonPosR.y
            If i = MenuHl Then
                pText.SetColor ColorMake(255, 255, 0)
            Else
                pText.SetColor ColorMake(255, 255, 255)
            End If
            pText.Render MenuStr(i)
        Next i
        If Not Busy Then
            MenuCursor.Render pCurs
        Else
            MenuCursor.Render pBusy
        End If
        Eng3d.RenderEnd
        Eng3d.RenderShow

        If ((RdMouse.Buttons(0) = 0 And LDown) Or (RdMouse.Buttons(1) = 0 And RDown) Or (RdKeyb(DIK_RETURN) = 0 And RdKeyb(DIK_NUMPADENTER) = 0 And RetDown)) And Not MenuHl = -1 Then
            If MenuHl = MenuCnt - 1 And Not MenuID = MENU_OPTIONS_GRAFICS Then
                'Back or exit.
                Exit Do
            End If
            If RetDown Then LDown = True
            If MenuID = MENU_MAIN Then
                'Main menu
                If MenuHl = 0 Then
                    'Start a single player game.
                    PIdent = 0
                    RunGame
                ElseIf MenuHl = 2 Then
                    'Options
                    ShowMenu MENU_OPTIONS, CursorPos, pCurs, pBusy, pButton, pText, pBackgr
                    If Trim$(gOptions.PlayerName) = "" Then gOptions.PlayerName = "Player"
                ElseIf MenuHl = 3 Then
                    'Game informations
                    ShowInfo pBackgr, False
                ElseIf MenuHl = 4 Then
                    'Credits
                    ShowInfo pBackgr, True
                End If
            ElseIf MenuID = MENU_OPTIONS Then
                'Options menu
                If MenuHl = 0 Then
                    ShowMenu MENU_OPTIONS_GRAFICS, CursorPos, pCurs, pBusy, pButton, pText, pBackgr
                ElseIf MenuHl = 2 Then
                    ShowMenu MENU_OPTIONS_DETAILS, CursorPos, pCurs, pBusy, pButton, pText, pBackgr
                ElseIf MenuHl = 3 Then
                    ShowMenu MENU_OPTIONS_GAMEPLAY, CursorPos, pCurs, pBusy, pButton, pText, pBackgr
                End If
            ElseIf MenuID = MENU_OPTIONS_GRAFICS Then
                'Graphics menu
                If MenuHl = 0 Then
                    'Change resolution. Bit depth keeps the same.
                    i = ActRes
                    Do
                        If LDown Then
                            i = i + 1
                        Else
                            i = i - 1
                        End If
                        If i < 0 Then i = ResCnt - 1
                        i = i Mod ResCnt
                    Loop While Not pRes(i).dxFormat = pRes(ActRes).dxFormat
                    ActRes = i
                ElseIf MenuHl = 1 Then
                    'Change bit depth. Resolution keeps the same.
                    For i = 0 To ResCnt - 1
                        If pRes(i).ResX = pRes(ActRes).ResX And pRes(i).ResY = pRes(ActRes).ResY And Not pRes(i).dxFormat = pRes(ActRes).dxFormat Then
                            ActRes = i
                            Eng3d.GetPossibleCaps pRes(ActRes).dxFormat, False, VarPtr(RdCaps)
                            Exit For
                        End If
                    Next i
                ElseIf MenuHl = 2 Then
                    'Change anti-aliasing
                    If LDown Then
                        If ActFSAA = 0 Then ActFSAA = 1
                        ActFSAA = ActFSAA * 2
                        If ActFSAA > RdCaps.MaxAntiAlias Then ActFSAA = 0
                    Else
                        ActFSAA = ActFSAA / 2
                        If ActFSAA = 1 Then
                            ActFSAA = 0
                        ElseIf ActFSAA = 0 Then
                            ActFSAA = RdCaps.MaxAntiAlias
                        End If
                    End If
                ElseIf MenuHl = 3 Then
                    'Accept
                    If Eng3d.SetRes(VarPtr(pRes(ActRes)), ActFSAA, False) Then
                        gOptions.Res = pRes(ActRes)
                        gOptions.UseFSAA = ActFSAA
                        Exit Do
                    Else
                        ActRes = PrevRes
                        ActFSAA = gOptions.UseFSAA
                    End If
                ElseIf MenuHl = 4 Then
                    'Back
                    ActRes = PrevRes
                    Eng3d.GetPossibleCaps pRes(ActRes).dxFormat, False, VarPtr(RdCaps)
                    Exit Do
                End If
                CreateMenuText MENU_OPTIONS_GRAFICS, MenuStr, pRes(ActRes), ActFSAA, False
            ElseIf MenuID = MENU_OPTIONS_DETAILS Then
                If MenuHl = 0 Then
                    'Mirror effects
                    gOptions.UseReflection = Not gOptions.UseReflection
                ElseIf MenuHl = 1 Then
                    'Shadows
                    If LDown Then
                        gOptions.UseShadow = gOptions.UseShadow + 1
                        gOptions.UseShadow = gOptions.UseShadow Mod 3
                    Else
                        gOptions.UseShadow = gOptions.UseShadow - 1
                        If gOptions.UseShadow < 0 Then gOptions.UseShadow = 2
                    End If
                ElseIf MenuHl = 2 Then
                    'Transparent walls
                    gOptions.UseTransWalls = Not gOptions.UseTransWalls
                ElseIf MenuHl = 3 Then
                    'Specular lights
                    gOptions.UseSpecular = Not gOptions.UseSpecular
                ElseIf MenuHl = 4 Then
                    'Anisotropic texture filtering
                    gOptions.UseAnisotropic = Not gOptions.UseAnisotropic
                End If
                CreateMenuText MENU_OPTIONS_DETAILS, MenuStr, pRes(ActRes), ActFSAA, False
            ElseIf MenuID = MENU_OPTIONS_GAMEPLAY Then
                If Not MenuHl = 7 Then EditName = False
                If MenuHl = 0 Then
                    'Action mode
                    gOptions.EnableAction = Not gOptions.EnableAction
                ElseIf MenuHl = 1 Then
                    'Camera
                    If LDown Then
                        gOptions.Cam = gOptions.Cam + 1
                        gOptions.Cam = gOptions.Cam Mod 8
                    Else
                        gOptions.Cam = gOptions.Cam - 1
                        If gOptions.Cam < 0 Then gOptions.Cam = 7
                    End If
                ElseIf MenuHl = 2 Then
                    'Arena size
                    If LDown Then
                        gOptions.LandSize = gOptions.LandSize + 100
                        If gOptions.LandSize > 500 Then
                            gOptions.LandSize = 100
                            If gOptions.OpponentCnt > 4 Then gOptions.OpponentCnt = 4
                        End If
                    Else
                        gOptions.LandSize = gOptions.LandSize - 100
                        If gOptions.LandSize = 100 And gOptions.OpponentCnt > 4 Then
                            gOptions.OpponentCnt = 4
                        ElseIf gOptions.LandSize < 100 Then
                            gOptions.LandSize = 500
                        End If
                    End If
                ElseIf MenuHl = 3 Then
                    'Bike speed
                    If LDown Then
                        gOptions.MopedSpeed = gOptions.MopedSpeed + 14
                        If gOptions.MopedSpeed = 126 Then gOptions.MopedSpeed = 14
                    Else
                        gOptions.MopedSpeed = gOptions.MopedSpeed - 14
                        If gOptions.MopedSpeed = 0 Then gOptions.MopedSpeed = 112
                    End If
                ElseIf MenuHl = 4 Then
                    'Bot count
                    If LDown Then
                        gOptions.OpponentCnt = gOptions.OpponentCnt + 1
                        If gOptions.OpponentCnt > 4 And gOptions.LandSize = 100 Then
                            gOptions.OpponentCnt = 1
                        ElseIf gOptions.OpponentCnt > 5 Then
                            gOptions.OpponentCnt = 1
                        End If
                    Else
                        gOptions.OpponentCnt = gOptions.OpponentCnt - 1
                        If gOptions.OpponentCnt = 0 Then
                            If gOptions.LandSize = 100 Then
                                gOptions.OpponentCnt = 4
                            Else
                                gOptions.OpponentCnt = 5
                            End If
                        End If
                    End If
                ElseIf MenuHl = 5 Then
                    'Bot skill
                    If LDown Then
                        gOptions.OpponentSkill = gOptions.OpponentSkill + 1
                        gOptions.OpponentSkill = gOptions.OpponentSkill Mod 3
                    Else
                        gOptions.OpponentSkill = gOptions.OpponentSkill - 1
                        If gOptions.OpponentSkill < 0 Then gOptions.OpponentSkill = 2
                    End If
                ElseIf MenuHl = 6 Then
                    'Point limit
                    If LDown Then
                        gOptions.FragLimit = gOptions.FragLimit + 5
                        If gOptions.FragLimit > 100 Then gOptions.FragLimit = 5
                    Else
                        gOptions.FragLimit = gOptions.FragLimit - 5
                        If gOptions.FragLimit < 5 Then gOptions.FragLimit = 100
                    End If
                ElseIf MenuHl = 7 Then
                    'Player's name
                    EditName = Not EditName
                End If
                CreateMenuText MENU_OPTIONS_GAMEPLAY, MenuStr, pRes(ActRes), ActFSAA, EditName
            End If
            LDown = False
            RDown = False
            RetDown = False
        End If
        If RdKeyb(DIK_ESCAPE) = 0 And ESCDown Then
            If MenuID = MENU_OPTIONS_GRAFICS Then
                ActRes = PrevRes
                Eng3d.GetPossibleCaps pRes(ActRes).dxFormat, False, VarPtr(RdCaps)
            End If
            Exit Do
        End If
    Loop
    
    Set MenuCursor = Nothing
    Set MenuButton = Nothing
    Set MenuBackgr = Nothing
End Sub

Private Sub ShowInfo(pBackgr As cls2dTexture, ByVal Credits As Boolean)
    Dim i As Long
    Dim Text As New cls2dText
    Dim MenuBackgr As New cls2dPicture
    Dim RdKeybPtr As Long, RdKeyb(255) As Byte
    Dim KeyDown As Boolean

    MenuBackgr.SetPosition 0, 0, gOptions.Res.ResX, gOptions.Res.ResY
    Text.SetFont "Arial", 40, True, False, False
    Text.SetColor ColorMake(0, 255, 0)
    Text.SetFormat 1, 1
    'Background picture
    Eng3d.RenderStart Vector2dMake(0, 0), Vector2dMake(0, 0)
    Eng3d.RenderClear True, ColorMake(0, 0, 255), False, False
    MenuBackgr.Render pBackgr
    'Heading
    Text.SetPosition 0, 0, gOptions.Res.ResX, 40
    If Credits Then
        Text.Render "CREDITS"
    Else
        Text.Render "REVO TRON v " & TRON_MVERS & "." & TRON_SVERS
    End If
    
    'Text
    Text.SetFont "Arial", 25, True, False, False
    Text.SetColor ColorMake(255, 255, 255)
    For i = 0 To 11
        Text.SetPosition 0, 60 + i * 30, gOptions.Res.ResX, 60 + (i + 1) * 30
        If Credits Then
            If i = 0 Then
                Text.Render "Completly programming work:"
            ElseIf i = 1 Then
                Text.Render "Mathias Kunter (mathiaskunter@yahoo.de)"
            ElseIf i = 3 Then
                Text.Render "Designer of the 3d objects:"
            ElseIf i = 4 Then
                Text.Render "Wolfgang Gelbmann (wolf_gelb@hotmail.com)"
            ElseIf i = 6 Then
                Text.Render "Designer of the floor and wall textures:"
            ElseIf i = 7 Then
                Text.Render "Tyler Esselstrom (hazard369@aol.com)"
            ElseIf i = 9 Then
                Text.Render "Beta testers:"
            ElseIf i = 10 Then
                Text.Render "Timon Kunter, True, Wolfgang Unger"
            End If
        Else
            If i = 0 Then
                Text.Render "Website: http://revotron.tripod.com"
            ElseIf i = 2 Then
                Text.Render "Key configuration:"
            ElseIf i = 3 Then
                Text.Render "Left and right: Changes the direction"
            ElseIf i = 4 Then
                Text.Render "Space: In action mode, use an item"
            ElseIf i = 5 Then
                Text.Render "C: Changes the camera view ingame"
            ElseIf i = 7 Then
                Text.Render "Action mode: This is a new feature which was invented"
            ElseIf i = 8 Then
                Text.Render "by Revo Tron. You'll get an item every 10 to 15 seconds."
            ElseIf i = 9 Then
                Text.Render "You can see this on the top-right corner of your screen."
            ElseIf i = 10 Then
                Text.Render "If you have an item, you can use it with the space key."
            End If
        End If
    Next i
    Eng3d.RenderEnd
    Eng3d.RenderShow
    'Wait for key press
    KeyDown = False
    Do While Not KeyDown
        RdKeybPtr = Inp.ReadKeyboard("")
        CopyMemory RdKeyb(0), ByVal RdKeybPtr, 256
        For i = 0 To 255
            If Not RdKeyb(i) = 0 Then
                KeyDown = True
                Exit For
            End If
        Next i
    Loop
    KeyDown = True
    Do While (KeyDown)
        KeyDown = False
        RdKeybPtr = Inp.ReadKeyboard("")
        CopyMemory RdKeyb(0), ByVal RdKeybPtr, 256
        For i = 0 To 255
            If Not RdKeyb(i) = 0 Then
                KeyDown = True
                Exit For
            End If
        Next i
    Loop
    
    Set Text = Nothing
    Set MenuBackgr = Nothing
End Sub

Private Function CreateMenuText(ByVal MenuID As Long, ByRef MenuStr() As String, ByRef UseRes As Revo3dRes, ByVal AntiAlias As Long, ByVal EditName As Boolean) As Long
    If MenuID = MENU_MAIN Then
        MenuStr(0) = "Singleplayer game"
        MenuStr(1) = "Multiplayer game - N/A"
        MenuStr(2) = "Options"
        MenuStr(3) = "Game information"
        MenuStr(4) = "Credits"
        MenuStr(5) = "Exit"
        CreateMenuText = 6
    ElseIf MenuID = MENU_OPTIONS Then
        MenuStr(0) = "Graphic options"
        MenuStr(1) = "Sound options - N/A"
        MenuStr(2) = "Detail options"
        MenuStr(3) = "Game settings"
        MenuStr(4) = "Back"
        CreateMenuText = 5
    ElseIf MenuID = MENU_OPTIONS_GRAFICS Then
        MenuStr(0) = "Resolution: " & UseRes.ResX & " x " & UseRes.ResY
        MenuStr(1) = "Color depth: " & Eng3d.GetFormatBPP(UseRes.dxFormat) & " bit"
        If Not AntiAlias = 0 Then
            MenuStr(2) = "Anti-Aliasing: " & AntiAlias & " x"
        Else
            MenuStr(2) = "Anti-Aliasing: off"
        End If
        MenuStr(3) = "Accept"
        MenuStr(4) = "Back"
        CreateMenuText = 5
    ElseIf MenuID = MENU_OPTIONS_DETAILS Then
        If gOptions.UseReflection Then
            MenuStr(0) = "Mirror effects: on"
        Else
            MenuStr(0) = "Mirror effects: off"
        End If
        If gOptions.UseShadow = 0 Then
            MenuStr(1) = "Shadows: off"
        ElseIf gOptions.UseShadow = 1 Then
            If RdCaps.ShadowNoTransparencyAviable Then
                MenuStr(1) = "Shadows: low"
            Else
                MenuStr(1) = "Shadows: low - N/A"
            End If
        Else
            If RdCaps.ShadowAviable Then
                MenuStr(1) = "Shadows: high"
            Else
                MenuStr(1) = "Shadows: high - N/A in " & Eng3d.GetFormatBPP(gOptions.Res.dxFormat) & " bit color depth"
            End If
        End If
        If gOptions.UseTransWalls Then
            MenuStr(2) = "Transparent walls: on"
        Else
            MenuStr(2) = "Transparent walls: off"
        End If
        If gOptions.UseSpecular Then
            If RdCaps.SpecularAviable Then
                MenuStr(3) = "Light reflections: on"
            Else
                MenuStr(3) = "Light reflections: on - N/A"
            End If
        Else
            MenuStr(3) = "Light reflections: off"
        End If
        If gOptions.UseAnisotropic Then
            If RdCaps.AnisotropicAviable Then
                MenuStr(4) = "Anisotropic texture filtering: on"
            Else
                MenuStr(4) = "Anisotropic texture filtering: on - N/A"
            End If
        Else
            MenuStr(4) = "Anisotropic texture filtering: off"
        End If
        MenuStr(5) = "Back"
        CreateMenuText = 6
    ElseIf MenuID = MENU_OPTIONS_GAMEPLAY Then
        If gOptions.EnableAction Then
            MenuStr(0) = "Action mode: on"
        Else
            MenuStr(0) = "Action mode: off"
        End If
        If gOptions.Cam = 0 Then
            MenuStr(1) = "Camera: behind, near"
        ElseIf gOptions.Cam = 1 Then
            MenuStr(1) = "Camera: behind, medium distance"
        ElseIf gOptions.Cam = 2 Then
            MenuStr(1) = "Camera: behind, far distance"
        ElseIf gOptions.Cam = 3 Then
            MenuStr(1) = "Camera: classic top-down"
        ElseIf gOptions.Cam = 4 Then
            MenuStr(1) = "Camera: modern top-down, low"
        ElseIf gOptions.Cam = 5 Then
            MenuStr(1) = "Camera: modern top-down, high"
        ElseIf gOptions.Cam = 6 Then
            MenuStr(1) = "Camera: onboard"
        ElseIf gOptions.Cam = 7 Then
            MenuStr(1) = "Camera: in front of"
        End If
        MenuStr(2) = "Arena size: " & gOptions.LandSize & " x " & gOptions.LandSize & " meters"
        MenuStr(3) = "Speed: " & Int((gOptions.MopedSpeed / 14) * 50) & " km/h"
        MenuStr(4) = "Bot count: " & gOptions.OpponentCnt
        If gOptions.OpponentSkill = 0 Then
            MenuStr(5) = "Bot skill: bad"
        ElseIf gOptions.OpponentSkill = 1 Then
            MenuStr(5) = "Bot skill: medium"
        ElseIf gOptions.OpponentSkill = 2 Then
            MenuStr(5) = "Bot skill: good"
        End If
        MenuStr(6) = "Point limit: " & gOptions.FragLimit
        If EditName Then
            MenuStr(7) = "Player's name: " & Trim$(gOptions.PlayerName) & "_"
        Else
            MenuStr(7) = "Player's name: " & Trim$(gOptions.PlayerName)
        End If
        MenuStr(8) = "Back"
        CreateMenuText = 9
    End If
End Function

Private Sub MasterGameController(ByVal RelTimeDiff As Single)
    Dim i As Long, CrashCnt As Long

    'Handle explosions
    gExplosions.DoEventsX RelTimeDiff
    
    '**************Handle shots***************
    gShots.DoEventsX RelTimeDiff

    '**************Handle shields************
    For i = 0 To gOptions.OpponentCnt
        gPlayer(i).HandleShield
    Next i

    '*******************Handle bikes**************
    If GameRuns Then
        'Steering of the player bike is done in the RunGame routine.
        For i = 0 To gOptions.OpponentCnt
            If gPlayer(i).Alive And Not i = PIdent Then
                'Call the AI routine to steer this bike.
                gPlayer(i).AI RelTimeDiff
            End If
        Next i
        'Call DoEventsX now for the bikes, since the steering process is done now.
        For i = 0 To gOptions.OpponentCnt
            gPlayer(i).DoEventsX RelTimeDiff, gOptions.MopedSpeed
        Next i
    Else
        If NotRunTime < 3 Then
            For i = 0 To gOptions.OpponentCnt
                gPlayer(i).DoEventsX RelTimeDiff, 3 - NotRunTime
            Next i
        Else
            For i = 0 To gOptions.OpponentCnt
                gPlayer(i).DoEventsX RelTimeDiff, 0
            Next i
        End If
    End If

    '******************Check the bikes for colissions**************
    For i = 0 To gOptions.OpponentCnt
        If gPlayer(i).Alive Then
            If Not gPlayer(i).CanPlace Then gPlayer(i).WillCrash = True
        End If
    Next i
    'Determine the number of bikes which are going to crash in this frame.
    CrashCnt = 0
    For i = 0 To gOptions.OpponentCnt
        If gPlayer(i).Alive And gPlayer(i).WillCrash Then CrashCnt = CrashCnt + 1
    Next i
    'Now, give points for every bike which is still alive.
    For i = 0 To gOptions.OpponentCnt
        If gPlayer(i).Alive Then
            If gPlayer(i).WillCrash Then
                gPlayer(i).Frags = gPlayer(i).Frags + CrashCnt - 1
                gPlayer(i).Crash
                'Create explosion
                gExplosions.CreateExplosion gPlayer(i).PosProp, False
                MopedAlive = MopedAlive - 1
            Else
                gPlayer(i).Frags = gPlayer(i).Frags + CrashCnt
            End If
        End If
    Next i


    '**************************Determine the leading player*************************
    LeaderScore = DoRanking
    If LeaderScore >= gOptions.FragLimit Then
        For i = 0 To gOptions.OpponentCnt
            gPlayer(i).Crash
        Next i
        GamePassed = True
        TimeFac = 1
    End If

    '*******************General*********************
    NotRunTime = NotRunTime + RelTimeDiff
    If Not GameRuns And NotRunTime >= 5 Then GameRuns = True
    If NotAliveTime >= 3 Then TimeFac = 3
    If MopedAlive < 2 Then
        GamePassed = True
        TimeFac = 1
    End If
End Sub

Private Sub RunGame()
    On Local Error GoTo Failed
    
    Dim i As Long, j As Long, x As Long, y As Long
    Dim fpsRendered As Long, OppShow As Long, RankLine As Long, RDataSize As Long
    Dim RelTimeDiff As Single, fpsT As Single, CamBlendTime As Single, TransErg As Single, SendT As Single, WaitTime As Single
    Dim fpsTxt As String, pntTxt As String
    Dim RdKeybPtr As Long, RdKeyb(255) As Byte
    Dim pMem As Long, SecondScore As Long
    Dim RDown As Boolean, LDown As Boolean, CDown As Boolean, ESCDown As Boolean, SUse As Boolean, ReturnDown As Boolean
    Dim StartOK As Boolean, FirstInit As Boolean, GameQuitQuest As Boolean

    Dim Cam As New TronCamera

    Dim GameLand As New cls3dPolygons
    Dim GameV(19) As Vertex
    Dim TexFloor As New cls2dTexture, TexWall As New cls2dTexture, TexMopedWall As New cls2dTexture, TexFire As New cls2dTexture
    Dim PicItem As New cls2dPicture
    Dim TexShotGreen As New cls2dTexture, TexShotRed As New cls2dTexture, TexSpeedup As New cls2dTexture, TexShotBlack As New cls2dTexture, TexShield As New cls2dTexture
    Dim fps As New cls2dText, TextRank As New cls2dText, bMatch As Boolean
    Dim FloorPlane As Revo3dPlane

    Dim TexRect As RECT
    Dim PointLight As D3DLIGHT8
    Dim FloorMat As D3DMATERIAL8
    Dim ActPos As D3DVECTOR2, CamDir As D3DVECTOR2, PrevDir As D3DVECTOR2, CamDiff As D3DVECTOR2, CamVec As D3DVECTOR2
    Dim NPos As D3DVECTOR, NDest As D3DVECTOR, CamDir3d As D3DVECTOR


    'Well, you won't believe it. We've got the variable declarations finished. *gg*


    '********************ONE TIME LOADING********************
    'Load textures
    If Not TexFloor.LoadFromFile("textures\floor.bmp", 0, 5, False) Then Exit Sub
    If Not TexWall.LoadFromFile("textures\wall.bmp", 0, 0, False) Then Exit Sub
    If Not TexMopedWall.LoadFromFile("textures\mopedwall.bmp", 0, 1, False) Then Exit Sub
    If Not TexFire.LoadFromFile("textures\fire.bmp", 0, 1, True) Then Exit Sub
    If TexFire.EditStart Then
        For x = 0 To 63
            For y = 0 To 63
                TransErg = 255 - 8 * Sqr((32 - x) * (32 - x) + (32 - y) * (32 - y))
                If TransErg > 0 Then
                    TexFire.Edit Vector2dMake(x, y), TransErg, -1, -1, -1
                Else
                    TexFire.Edit Vector2dMake(x, y), 0, -1, -1, -1
                End If
            Next y
        Next x
        TexFire.EditEnd
    End If
    If Not TexShotGreen.LoadFromFile("textures\shotgreen.bmp", 0, 1, False) Then Exit Sub
    If Not TexShotRed.LoadFromFile("textures\shotred.bmp", 0, 1, False) Then Exit Sub
    If Not TexSpeedup.LoadFromFile("textures\speedup.bmp", ColorMake(0, 0, 0), 1, False) Then Exit Sub
    If Not TexShotBlack.LoadFromFile("textures\shotblack.bmp", 0, 1, False) Then Exit Sub
    If Not TexShield.LoadFromFile("textures\shield.bmp", ColorMake(0, 0, 0), 1, False) Then Exit Sub
    PicItem.SetPosition gOptions.Res.ResX - 64, 0, gOptions.Res.ResX, 64

    'Load landscape
    GeometryInit gOptions.LandSize, gOptions.LandSize, 15, GameV
    GameLand.Initialize False
    GameLand.SetVertexData VarPtr(GameV(0)), 0, 20
    GameLand.SetMaterial Eng3d.GetDefaultMaterial
    GameLand.SetPolyFormat TRIANGLESTRIP
    GameLand.SetVisibility FRONTSIDE

    'Create shield object
    If Not Shield.LoadFromFile("objects\shield.x") Then Exit Sub
    Shield.SetTransparency 100

    'Create text objects
    fps.SetFont "Arial", 20, False, False, False
    fps.SetColor ColorMake(255, 255, 255)
    fps.SetPosition 0, 0, 400, 20
    TextRank.SetFont "Arial", 20, True, False, False
    TextRank.SetColor ColorMake(255, 255, 255)

    'Prepare walls
    gWallDesc.Initialize 100 * (gOptions.OpponentCnt + 1)

    'Prepare explosions
    gExplosions.Initialize 30, TexFire

    'Create the player objects.
    ReDim gPlayer(gOptions.OpponentCnt)
    For i = 0 To gOptions.OpponentCnt
        Set gPlayer(i) = New TronPlayer
    Next i

    'Prepare shots
    gShots.Initialize 10

    'Create light sources
    PointLight = PointLightMake(40, 40, 40, 10, Vector3dMake(gOptions.LandSize / 2, 30, gOptions.LandSize / 2), 1000)
    Eng3d.SetLight 0, PointLight
    PointLight = PointLightMake(0, 0, 0, 3, Vector3dMake(0, 0, 0), 1000)

    'Create planes
    FloorPlane = PlaneMake(Vector3dMake(0, 0.01, 0), Vector3dMake(0, 0.01, 1), Vector3dMake(1, 0.01, 0))
    FloorMat = MaterialMake(255, 255, 255, 0, 0, 0, 0, 75)

    'Anisotropic filter
    If gOptions.UseAnisotropic Then Eng3d.SetAnisotropicState True

    'General
    FirstInit = True
    LeaderScore = 0

    'Create the players.
    For i = 0 To gOptions.OpponentCnt
        If Not gPlayer(i).Initialize(Int(GetRandomVal(0, 4)) Mod 4, i, TexMopedWall) Then Exit Sub
    Next i

    Do While LeaderScore < gOptions.FragLimit
        '*********************LOADING BEFORE EVERY GAME*****************
        If Not FirstInit Then
            gShots.Delete
            gExplosions.DoEventsX 60
            Cam.DoEventsX 60
            For i = 0 To gOptions.OpponentCnt
                gPlayer(i).Crash
                gPlayer(i).Delete
            Next i
            gWallDesc.Delete
        End If
        'Generate startup positions for the bikes.
        For i = 0 To gOptions.OpponentCnt
            gPlayer(i).Reset
        Next i
        'Check if the startup positions aren't too close.
        Do
            StartOK = True
            For i = 0 To gOptions.OpponentCnt
                For j = 0 To gOptions.OpponentCnt
                    If Not i = j Then
                        If Abs(gPlayer(j).PosProp.x - gPlayer(i).PosProp.x) < 20 And Abs(gPlayer(j).PosProp.y - gPlayer(i).PosProp.y) < 20 Then
                            StartOK = False
                            Exit For
                        End If
                    End If
                Next j
            Next i
            If Not StartOK Then
                'Re-generate the startup positions.
                For i = 0 To gOptions.OpponentCnt
                    gPlayer(i).GeneratePos
                Next i
            End If
        Loop While Not StartOK

        'Init variables
        FirstInit = False
        CamDir = gPlayer(PIdent).DirProp
        CamDir3d = Vector3dMake(0, 0, 0)
        fpsRendered = 0
        OppShow = -1
        MopedAlive = gOptions.OpponentCnt + 1
        fpsT = 0
        CamBlendTime = 0
        NotRunTime = 0
        NotAliveTime = 0
        BlendTime = 0
        RDown = False
        LDown = False
        CDown = False
        ESCDown = False
        ReturnDown = False
        GameRuns = False
        GamePassed = False
        CamMode = 0
        TimeFac = 1
        'SendT = 0          multiplayer variables. Not kicked out in VB version, but also not implemented.
        GameQuitQuest = False
        'MPGameBreak = False

        Eng3d.SetLookDistance 1000

        '**************************RUNNING LOOP FOR ONE GAME*******************
        Eng3d.StartTimer
        Do While True
            RelTimeDiff = Eng3d.GetRelTimeDiff
            'Text outputs
            fpsT = fpsT + RelTimeDiff
            If fpsT >= 0.5 Then
                fpsTxt = Format$(fpsRendered / fpsT, "0.00") & " fps"
                fpsT = 0
                fpsRendered = 0
            End If
            fpsRendered = fpsRendered + 1
            RelTimeDiff = RelTimeDiff * TimeFac

            '********************Read keyboard and mouse inputs***********************
            Inp.ReadMouse
            RdKeybPtr = Inp.ReadKeyboard("")
            If RdKeybPtr = 0 Then Exit Sub
            CopyMemory RdKeyb(0), ByVal RdKeybPtr, 256
            'Left key
            If Not RdKeyb(DIK_LEFT) = 0 And Not LDown And gPlayer(PIdent).Alive And GameRuns Then
                gPlayer(PIdent).ChangeDir 0
                D3DXVec2Subtract CamDiff, gPlayer(PIdent).DirProp, CamDir
                PrevDir = CamDir
                CamBlendTime = 0.2
                LDown = True
            ElseIf RdKeyb(DIK_LEFT) = 0 Then
                LDown = False
            End If
            'Right key
            If Not RdKeyb(DIK_RIGHT) = 0 And Not RDown And gPlayer(PIdent).Alive And GameRuns Then
                gPlayer(PIdent).ChangeDir 1
                D3DXVec2Subtract CamDiff, gPlayer(PIdent).DirProp, CamDir
                PrevDir = CamDir
                CamBlendTime = 0.2
                RDown = True
            ElseIf RdKeyb(DIK_RIGHT) = 0 Then
                RDown = False
            End If
            'Space key (use an item)
            If Not RdKeyb(DIK_SPACE) = 0 And Not gPlayer(PIdent).ItemAviable = -1 And gPlayer(PIdent).Alive And GameRuns Then
                gPlayer(PIdent).FireItem
            End If
            'c key (change the camera)
            If Not RdKeyb(DIK_C) = 0 And Not CDown And gPlayer(PIdent).Alive And GameRuns Then
                gOptions.Cam = gOptions.Cam + 1
                gOptions.Cam = gOptions.Cam Mod 8
                If gOptions.Cam <= 3 Or gOptions.Cam >= 6 Then
                    Eng3d.SetLookDistance 1000, PI / 2.5
                Else
                    Eng3d.SetLookDistance 1000
                End If
                CDown = True
                Cam.BlendPosition Vector3dMake(0, 0, 0), 0
            ElseIf RdKeyb(DIK_C) = 0 Then
                CDown = False
            End If
            'ESC key
            If Not RdKeyb(DIK_ESCAPE) = 0 Then
                ESCDown = True
            ElseIf RdKeyb(DIK_ESCAPE) = 0 And ESCDown Then
                GameQuitQuest = Not GameQuitQuest
                ESCDown = False
            End If
            'Return key (quit a game, if it's already finished)
            If Not RdKeyb(DIK_RETURN) = 0 Or Not RdKeyb(DIK_NUMPADENTER) = 0 Then
                ReturnDown = True
            ElseIf RdKeyb(DIK_RETURN) = 0 And RdKeyb(DIK_NUMPADENTER) = 0 And ReturnDown Then
                If GameQuitQuest Then Exit Sub
                If GamePassed Then Exit Do
                ReturnDown = False
            End If


            '////////////////////G A M E   C O N T R O L L E R////////////////////
            If Not gPlayer(PIdent).Alive Then NotAliveTime = NotAliveTime + RelTimeDiff

            MasterGameController RelTimeDiff


            '////////////////////R E N D E R E R////////////////////
            '******************Camera setup****************
            ActPos = gPlayer(PIdent).PosProp

            If CamMode = 0 And NotRunTime >= 3 Then
                CamMode = 1
            ElseIf CamMode = 2 And NotRunTime >= 5 Then
                CamMode = 3
            ElseIf CamMode = 3 And Not gPlayer(PIdent).Alive Then
                CamMode = 4
            ElseIf CamMode = 5 And NotAliveTime >= 3 Then
                CamMode = 6
            ElseIf CamMode = 7 And BlendTime >= 1.5 Then
                CamMode = 6
            End If
            If CamMode = 0 Then
                'Game doesn't run yet, near bike camera.
                Cam.Dest = Vector3dMake(ActPos.x, 0, ActPos.y)
                CamVec = NormalVector2dMake(Vector2dMake(0, 0), gPlayer(PIdent).DirProp)
                Cam.Pos = Vector3dMake(ActPos.x - CamVec.x * 7, 6 - NotRunTime, ActPos.y - CamVec.y * 7)
            ElseIf CamMode = 1 Then
                'Game doesn't run yet, blend camera from near camera to the game camera.
                GetCamDesc ActPos, gPlayer(PIdent).DirProp, NPos, NDest
                Cam.BlendPosition NPos, 2
                Cam.BlendDestination NDest, 2
                If gOptions.Cam <= 3 Or gOptions.Cam >= 6 Then Cam.BlendAngle PI / 4, PI / 2.5, 2
                CamMode = 2
            'CamMode = 2 is just used for blending.
            ElseIf CamMode = 3 Then
                'Game runs. Set the camera accordingly to the camera option.
                'Change the camera direction if the bike changes its driving direction.
                If CamBlendTime > 0 Then
                    'Change the camera direction.
                    D3DXVec2Scale CamDir, CamDiff, 1 - CamBlendTime / 0.2
                    D3DXVec2Add CamDir, PrevDir, CamDir
                    CamBlendTime = CamBlendTime - RelTimeDiff
                Else
                    'Use the camera direction from the bike direction.
                    CamDir = gPlayer(PIdent).DirProp
                End If
                GetCamDesc ActPos, CamDir, NPos, NDest
                If (gOptions.Cam = 4 Or gOptions.Cam = 5) And Not CDown Then
                    'Blend camera to the new position.
                    Cam.BlendPosition NPos, 0.25
                    Cam.Dest = NDest
                Else
                    'Set camera immediately to the new position.
                    Cam.Pos = NPos
                    Cam.Dest = NDest
                End If
            ElseIf CamMode = 4 Then
                'The player just crashed, blend the camera.
                Cam.BlendPosition Vector3dMake(gOptions.LandSize / 2, 40, gOptions.LandSize / 2), 3
                If gOptions.Cam <= 3 Or gOptions.Cam >= 6 Then Cam.BlendAngle PI / 2.5, PI / 4, 3
                CamMode = 5
            'CamMode = 5 is just used for blending.
            ElseIf CamMode = 6 Then
                'If the player isn't alive any more, show another bike which is still alive.
                bMatch = False
                If OppShow = -1 Then
                    bMatch = True
                ElseIf Not gPlayer(OppShow).Alive Then
                    bMatch = True
                End If
                If bMatch Then
                    'Find a bike to show.
                    For i = 0 To gOptions.OpponentCnt
                        If Not i = OppShow And gPlayer(i).Alive Then
                            If Not OppShow = -1 Then
                                BlendTime = 0
                                CamMode = 7
                            End If
                            OppShow = i
                            Exit For
                        End If
                    Next i
                ElseIf Not OppShow = -1 Then
                    If TimeFac = 1 Then
                        'Use a special swing camera when the game is over.
                        D3DXVec3Subtract NDest, Vector3dMake(gPlayer(OppShow).PosProp.x - gPlayer(OppShow).DirProp.x * 5, GetRandomVal(10, 30), gPlayer(OppShow).PosProp.y - gPlayer(OppShow).DirProp.y * 5), Cam.Pos
                        D3DXVec3Normalize NDest, NDest
                        D3DXVec3Lerp CamDir3d, CamDir3d, NDest, RelTimeDiff * 3             'delay of 1/3 second
                        D3DXVec3Scale CamDir3d, CamDir3d, gOptions.MopedSpeed / 1.25 * RelTimeDiff
                        D3DXVec3Add NPos, Cam.Pos, CamDir3d
                        D3DXVec3Normalize CamDir3d, CamDir3d
                        If NPos.y < 10 Then NPos.y = 10
                        Cam.Pos = NPos
                    End If
                    Cam.Dest = Vector3dMake(gPlayer(OppShow).PosProp.x, 0, gPlayer(OppShow).PosProp.y)
                End If
            ElseIf CamMode = 7 Then
                'Static camera, if a bike explodes.
                BlendTime = BlendTime + RelTimeDiff
            End If
            Cam.DoEventsX RelTimeDiff

            '****************Set lights*************
            PointLight.Position = Cam.Pos
            Eng3d.SetLight 1, PointLight

            '****************Render*************
            Eng3d.RenderStart Vector2dMake(0, 0), Vector2dMake(0, 0)
            If gOptions.UseShadow = 2 Then
                Eng3d.RenderClear True, ColorMake(0, 0, 0), True, True
            Else
                Eng3d.RenderClear True, ColorMake(0, 0, 0), True, False
            End If
            If (gOptions.Cam <= 3 Or gOptions.Cam >= 6) And Cam.IsAngleBlended Then
                Eng3d.SetLookDistance 1000, Cam.GetAngle
            End If
            Cam.SetCamera
            'Render mirror effects, if enabled.
            If gOptions.UseReflection Then
                Eng3d.SetReflectionState VarPtr(FloorPlane), False
                'Render walls.
                Eng3d.SetGlobalLight ColorMake(200, 200, 200)
                Eng3d.SetLightState 0, False
                Eng3d.SetLightState 1, False
                Eng3d.SetZState False, False
                GameLand.SetTexture TexWall
                GameLand.Render 4, 2
                GameLand.Render 8, 2
                GameLand.Render 12, 2
                GameLand.Render 16, 2
                Eng3d.SetZState True, True
                'Render bikes.
                Eng3d.SetGlobalLight ColorMake(0, 0, 0)
                Eng3d.SetLightState 0, True
                Eng3d.SetLightState 1, True
                Eng3d.SetSpecularState gOptions.UseSpecular
                For i = 0 To gOptions.OpponentCnt
                    gPlayer(i).Render False, Shield
                Next i
                Eng3d.SetLightState 1, False
                Eng3d.SetSpecularState False
                'Render walls of the bikes.
                For i = 0 To gOptions.OpponentCnt
                    gPlayer(i).RenderWalls True
                Next i
                Eng3d.SetLightState 0, False
                'Render shots
                Eng3d.SetGlobalLight ColorMake(255, 255, 255)
                gShots.Render False
                Eng3d.SetReflectionState 0, False
            End If
            Eng3d.SetLightState 0, False
            Eng3d.SetLightState 1, False
            'Render floor.
            If gOptions.UseShadow = 0 Then Eng3d.SetZState False, False
            GameLand.SetTexture TexFloor
            If gOptions.UseReflection Then
                'Render transparent floor.
                Eng3d.SetGlobalLight ColorMake(255, 255, 255)
                GameLand.SetMaterial FloorMat
                GameLand.Render 0, 2
                Eng3d.SetGlobalLight ColorMake(200, 200, 200)
                GameLand.SetMaterial Eng3d.GetDefaultMaterial
            Else
                Eng3d.SetGlobalLight ColorMake(200, 200, 200)
                GameLand.Render 0, 2
            End If
            'Render arena walls.
            GameLand.SetTexture TexWall
            GameLand.Render 4, 2
            GameLand.Render 8, 2
            GameLand.Render 12, 2
            GameLand.Render 16, 2
            'Render shadows, if enabled.
            If Not gOptions.UseShadow = 0 Then
                Eng3d.SetZState True, False
                'Shadow on the floor
                If gOptions.UseShadow = 1 Then
                    SUse = Eng3d.SetShadowState(VarPtr(FloorPlane), Vector3dMake(CamVec.x * 10, 9, CamVec.y * 10), 255, DIRECTIONALLIGHT, False)
                ElseIf gOptions.UseShadow = 2 Then
                    SUse = Eng3d.SetShadowState(VarPtr(FloorPlane), Vector3dMake(CamVec.x * 10, 9, CamVec.y * 10), 100, DIRECTIONALLIGHT, False)
                End If
                If SUse Then
                    For i = 0 To gOptions.OpponentCnt
                        gPlayer(i).Render True, Nothing
                        gPlayer(i).RenderWalls False
                    Next i
                End If
                Eng3d.SetShadowState 0, Vector3dMake(0, 0, 0), 0, 0, False
            End If
            Eng3d.SetZState True, True
            'Render shots
            gShots.Render True
            'Render bikes
            Eng3d.SetSpecularState gOptions.UseSpecular
            Eng3d.SetGlobalLight ColorMake(0, 0, 0)
            Eng3d.SetLightState 0, True
            Eng3d.SetLightState 1, Not gOptions.Cam = 6
            gPlayer(PIdent).Render False, Shield
            Eng3d.SetLightState 1, True
            For i = 0 To gOptions.OpponentCnt
                If Not i = PIdent Then gPlayer(i).Render False, Shield
            Next i
            'Render walls of the bikes.
            Eng3d.SetLightState 1, False
            Eng3d.SetSpecularState False
            For i = 0 To gOptions.OpponentCnt
                gPlayer(i).RenderWalls False
            Next i
            Eng3d.SetLightState 0, False
            'Render explosions.
            Eng3d.SetGlobalLight ColorMake(255, 255, 255)
            gExplosions.Render Cam.Pos
            If gOptions.EnableAction And gPlayer(PIdent).Alive And Not gPlayer(PIdent).ItemAviable = -1 Then
                'Render item picture
                If gPlayer(PIdent).ItemAviable = 0 Then
                    PicItem.Render TexSpeedup
                ElseIf gPlayer(PIdent).ItemAviable = 1 Then
                    PicItem.Render TexShotGreen
                ElseIf gPlayer(PIdent).ItemAviable = 2 Then
                    PicItem.Render TexShotRed
                ElseIf gPlayer(PIdent).ItemAviable = 3 Then
                    PicItem.Render TexShotBlack
                ElseIf gPlayer(PIdent).ItemAviable = 4 Then
                    PicItem.Render TexShield
                End If
            End If
            'Render text outputs
            fps.Render fpsTxt
            If GameQuitQuest Then
                TextRank.SetFormat 1, 1
                TextRank.SetPosition 0, gOptions.Res.ResY - 20, gOptions.Res.ResX, gOptions.Res.ResY
                TextRank.SetColor ColorMake(255, 255, 255)
                TextRank.Render "Really quit? Return = yes, ESC = no"
            End If
            If GamePassed Then
                'Render the ranking table.
                TextRank.SetFormat 1, 1
                RankLine = 0
                If LeaderScore >= gOptions.FragLimit Then
                    TextRank.SetPosition 0, 20, gOptions.Res.ResX, 50
                    TextRank.SetColor ColorMake(255, 0, 0)
                    TextRank.Render "FINAL RESULTS"
                End If
                For i = 0 To gOptions.OpponentCnt
                    For j = 0 To gOptions.OpponentCnt
                        If gPlayer(j).Rank = i + 1 Then
                            TextRank.SetPosition gOptions.Res.ResX / 5, 100 + RankLine * 30, gOptions.Res.ResX * 4 / 5, 100 + (RankLine + 1) * 30
                            If j = PIdent Then
                                TextRank.SetColor ColorMake(255, 255, 0)
                            Else
                                TextRank.SetColor ColorMake(255, 255, 255)
                            End If
                            TextRank.Render gPlayer(j).Rank & ".: " & gPlayer(j).Name & " with " & gPlayer(j).Frags & " points"
                            RankLine = RankLine + 1
                        End If
                    Next j
                Next i
                TextRank.SetColor ColorMake(255, 255, 255)
                If Not GameQuitQuest Then
                    TextRank.SetPosition 0, gOptions.Res.ResY - 20, gOptions.Res.ResX, gOptions.Res.ResY
                    TextRank.Render "Press return to continue"
                End If
            Else
                If gPlayer(PIdent).Rank = 1 Then
                    SecondScore = -1

                    For i = 0 To gOptions.OpponentCnt
                        If Not i = PIdent And gPlayer(i).Frags > SecondScore Then SecondScore = gPlayer(i).Frags
                    Next i
                    If SecondScore = gPlayer(PIdent).Frags Then
                        pntTxt = "Points: " & gPlayer(PIdent).Frags & " (+0)"
                    Else
                        pntTxt = "Points: " & gPlayer(PIdent).Frags & " (" & SecondScore - LeaderScore & ")"
                    End If
                Else
                    pntTxt = "Points: " & gPlayer(PIdent).Frags & " (+" & LeaderScore - gPlayer(PIdent).Frags & ")"
                End If
                TextRank.SetPosition 0, gOptions.Res.ResY - 20, gOptions.Res.ResX, gOptions.Res.ResY
                TextRank.SetFormat 0, 1
                TextRank.Render pntTxt
            End If
            Eng3d.RenderEnd
            Eng3d.RenderShow
        Loop
    Loop
    For i = 0 To gOptions.OpponentCnt
        Set gPlayer(i) = Nothing
    Next i
Failed:
End Sub

'Sorts the field accordingly to the current ranking.
Private Function DoRanking() As Long
    Dim i As Long, j As Long, min As Long, MaxScore As Long
    Dim ActRank As Long, AddRank As Long
    Dim pField() As TronSwap, pSwap As TronSwap

    ReDim pField(gOptions.OpponentCnt)
    For i = 0 To gOptions.OpponentCnt
        pField(i).Frags = gPlayer(i).Frags
        pField(i).Ident = i
    Next i
    'Sort descending
    For i = 0 To gOptions.OpponentCnt - 1
        min = i
        For j = i + 1 To gOptions.OpponentCnt
            If pField(min).Frags < pField(j).Frags Then min = j
        Next j
        If Not min = i Then
            'Swap
            pSwap = pField(min)
            pField(min) = pField(i)
            pField(i) = pSwap
        End If
    Next i
    MaxScore = pField(0).Frags
    'Now, write the rankings to the player objects.
    For i = 0 To gOptions.OpponentCnt
        AddRank = AddRank + 1
        If i = 0 Then
            'Has less frags then the previous player, so increase the ranking.
            ActRank = ActRank + AddRank
            AddRank = 0
        ElseIf pField(i).Frags < pField(i - 1).Frags Then
            'Has less frags then the previous player, so increase the ranking.
            ActRank = ActRank + AddRank
            AddRank = 0
        End If
        gPlayer(pField(i).Ident).Rank = ActRank
    Next i
    DoRanking = MaxScore
End Function

Private Sub GetCamDesc(ActPos As D3DVECTOR2, CamDir As D3DVECTOR2, CamPos As D3DVECTOR, CamDest As D3DVECTOR)
    Dim CamVec As D3DVECTOR2

    If gOptions.Cam = 0 Then
        CamDest = Vector3dMake(ActPos.x + CamDir.x * 30, 0, ActPos.y + CamDir.y * 30)
        CamPos = Vector3dMake(ActPos.x - CamDir.x * 10, 6, ActPos.y - CamDir.y * 10)                'near distance back camera
    ElseIf gOptions.Cam = 1 Then
        CamDest = Vector3dMake(ActPos.x + CamDir.x * 30, 0, ActPos.y + CamDir.y * 30)
        CamPos = Vector3dMake(ActPos.x - CamDir.x * 20, 10, ActPos.y - CamDir.y * 20)               'medium distance back camera
    ElseIf gOptions.Cam = 2 Then
        CamDest = Vector3dMake(ActPos.x + CamDir.x * 30, 0, ActPos.y + CamDir.y * 30)
        CamPos = Vector3dMake(ActPos.x - CamDir.x * 30, 15, ActPos.y - CamDir.y * 30)               'far distance back camera
    ElseIf gOptions.Cam = 3 Then
        CamDest = Vector3dMake(ActPos.x + CamDir.x * 20, 0, ActPos.y + CamDir.y * 20)
        CamPos = Vector3dMake(ActPos.x + CamDir.x * 15, 75, ActPos.y + CamDir.y * 15)               'classical top-down camera
    ElseIf gOptions.Cam = 4 Then
        CamVec = Vector2dMake(gOptions.LandSize / 2 - ActPos.x, gOptions.LandSize / 2 - ActPos.y)
        D3DXVec2Normalize CamVec, CamVec
        D3DXVec2Scale CamVec, CamVec, 100
        CamDest = Vector3dMake(ActPos.x, 0, ActPos.y)
        CamPos = Vector3dMake(ActPos.x + CamVec.x, 50, ActPos.y + CamVec.y)                         'modern top-down camera, low
    ElseIf gOptions.Cam = 5 Then
        CamVec = Vector2dMake(gOptions.LandSize / 2 - ActPos.x, gOptions.LandSize / 2 - ActPos.y)
        D3DXVec2Normalize CamVec, CamVec
        D3DXVec2Scale CamVec, CamVec, 100
        CamDest = Vector3dMake(ActPos.x, 0, ActPos.y)
        CamPos = Vector3dMake(ActPos.x + CamVec.x, 100, ActPos.y + CamVec.y)                        'modern top-down camera, high
    ElseIf gOptions.Cam = 6 Then
        CamDest = Vector3dMake(ActPos.x, 0, ActPos.y)
        CamPos = Vector3dMake(ActPos.x - CamDir.x * 2, 2, ActPos.y - CamDir.y * 2)
        D3DXVec3Add CamDest, CamPos, Vector3dMake(CamDir.x * 15, 0, CamDir.y * 15)                  'onboard camera
    ElseIf gOptions.Cam = 7 Then
        CamDest = Vector3dMake(ActPos.x, 0, ActPos.y)
        CamPos = Vector3dMake(ActPos.x + CamDir.x * 3, 1, ActPos.y + CamDir.y * 3)
        D3DXVec3Add CamDest, CamPos, Vector3dMake(CamDir.x, 0, CamDir.y)                            'in front of-camera
    End If
End Sub







'Helper functions

Public Function GetRandomVal(ByVal pmin As Single, ByVal pmax As Single) As Single
    GetRandomVal = Rnd * (pmax - pmin) + pmin
End Function

Public Function InArea(ByVal pmin As Single, ByVal pmax As Single, ByVal pval As Single) As Boolean
    If (pval > pmin And pval < pmax) Or (pval > pmax And pval < pmin) Then InArea = True
End Function

Public Function InAreaDelta(ByVal pmin As Single, ByVal pmax As Single, ByVal pval As Single, ByVal Delta As Single) As Boolean
    If InArea(pmin - Delta, pmax + Delta, pval) Or InArea(pmin + Delta, pmax - Delta, pval) Then InAreaDelta = True
End Function

Public Function InLine(ByVal v11 As Single, ByVal v12 As Single, ByVal v21 As Single, ByVal v22 As Single) As Boolean
    If InArea(v11, v12, v21) Or InArea(v11, v12, v22) Then
        InLine = True
        Exit Function
    End If
    If InArea(v21, v22, v11) Or InArea(v21, v22, v12) Then InLine = True
End Function
