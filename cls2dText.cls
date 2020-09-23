VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2dText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FontHandle As Long
Private pFont As D3DXFont
Private REC As RECT
Private myFor As Long
Private Col As Long


Private Sub Class_Initialize()
    myFor = DT_LEFT Or DT_TOP
    Col = ColorMake(0, 0, 0)
    SetFont "", 16, False, False, False
    gEngine.RegisterObject Me, OBJ_2DTEXT
End Sub

Private Sub Class_Terminate()
    Cleanup
    DeleteObject FontHandle
    FontHandle = 0
    gEngine.UnregisterObject Me
End Sub

'//-----------------------------------------------------------------------------
'// Function: Cleanup
'// Desc: Terminates the text.
'//-----------------------------------------------------------------------------
Public Sub Cleanup()
    Set pFont = Nothing
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetFont
'// Desc: Sets a font for the text.
'// Param: FontName (name of the font, for example "Arial"), Size, Bold, Italic, Underline
'//-----------------------------------------------------------------------------
Public Sub SetFont(ByVal FontName As String, ByVal Size As Long, ByVal Bold As Boolean, ByVal Italic As Boolean, ByVal Underline As Boolean)
    On Local Error Resume Next
    Cleanup
    If gD3DX Is Nothing Or gD3DDevice Is Nothing Then Exit Sub
    FontHandle = CreateFont(Size, 0, 0, 0, IIf(Bold, 700, 400), Italic, Underline, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, FontName)
    Set pFont = gD3DX.CreateFont(gD3DDevice, FontHandle)
End Sub

Friend Sub ReSetFont()
    On Local Error Resume Next
    Cleanup
    Set pFont = gD3DX.CreateFont(gD3DDevice, FontHandle)
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetPosition
'// Desc: Sets the position for the text.
'// Param: x1, y1 (upper left corner), x2, y2 (lower right corner)
'//-----------------------------------------------------------------------------
Public Sub SetPosition(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    REC.Left = x1
    REC.Top = y1
    REC.Right = x2
    REC.bottom = y2
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetFormat
'// Desc: Sets the format of the text.
'// Param: FormatH (left = 0, center = 1, right = 2),
'// FormatV (top = 0, center = 1, bottom = 2)
'//-----------------------------------------------------------------------------
Public Sub SetFormat(ByVal FormatH As Long, ByVal FormatV As Long)
    myFor = 0
    If FormatH = 0 Then
        myFor = DT_LEFT
    ElseIf FormatH = 1 Then
        myFor = DT_CENTER
    Else
        myFor = DT_RIGHT
    End If
    If FormatV = 0 Then
        myFor = myFor Or DT_TOP
    ElseIf FormatV = 1 Then
        myFor = myFor Or DT_VCENTER
    Else
        myFor = myFor Or DT_BOTTOM
    End If
End Sub

'//-----------------------------------------------------------------------------
'// Function: SetColor
'// Desc: Sets the color of the text.
'// Param: Col (the color)
'//-----------------------------------------------------------------------------
Public Sub SetColor(ByVal ColParam As Long)
    Col = ColParam
End Sub

'//-----------------------------------------------------------------------------
'// Function: Render
'// Desc: Renders the 2d text.
'// Param: Str (the string)
'//-----------------------------------------------------------------------------
Public Sub Render(ByVal Str As String)
    On Local Error Resume Next
    pFont.DrawTextW Str, -1, REC, myFor, Col
End Sub
