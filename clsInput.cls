VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDI As DirectInput8
Private pKeyb As DirectInputDevice8
Private pMouse As DirectInputDevice8
Private MouseData As DIMOUSESTATE
Private KeybData(255) As Byte, KeybDown(255) As Byte


Private Sub Class_Terminate()
    Cleanup
End Sub

'//-----------------------------------------------------------------------------
'// Function: Initialize
'// Desc: Initializes RevoInput.
'// Param: hWnd (the hWnd of the 3d engine window), NoWinKeys (if windows keys
'// are supressed), NoWinMouse (if the windows cursor should be hidden)
'//-----------------------------------------------------------------------------
Public Function Initialize(ByVal hWnd As Long, ByVal NoWinKeys As Boolean, ByVal NoWinMouse As Boolean) As Boolean
    Dim DIFormat As DIDATAFORMAT
    
    On Local Error GoTo Failed
    
    Cleanup

    Set pDI = gDX.DirectInputCreate
    Set pKeyb = pDI.CreateDevice("GUID_SysKeyboard")
    pKeyb.SetCommonDataFormat DIFORMAT_KEYBOARD
    pKeyb.SetCooperativeLevel hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
    pKeyb.Acquire

    Set pMouse = pDI.CreateDevice("GUID_SysMouse")
    pMouse.SetCommonDataFormat DIFORMAT_MOUSE
    pMouse.SetCooperativeLevel hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
    pMouse.Acquire

    Initialize = True
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: Cleanup
'// Desc: Terminates RevoInput.
'//-----------------------------------------------------------------------------
Public Function Cleanup()
    If Not pMouse Is Nothing Then
        pMouse.Unacquire
        Set pMouse = Nothing
    End If
    If Not pKeyb Is Nothing Then
        pKeyb.Unacquire
        Set pKeyb = Nothing
    End If
    Set pDI = Nothing
End Function

'//-----------------------------------------------------------------------------
'// Function: ReadKeyboard
'// Desc: Reads the currently pressed keys into a buffer.
'// Param: AppendTo (A string where text inputs are appended. Can be used to read text
'// inputs from the keyboard.)
'// Returns: Pointer to the BYTE buffer (f(0) to f(255)), or NULL if an error occured.
'//-----------------------------------------------------------------------------
Public Function ReadKeyboard(ByRef AppendTo As String) As Long
    Dim i As Long
    Dim RdSign As String, UpStr As String
    Dim ShiftDown As Boolean

    On Local Error Resume Next
    Err.Number = 0

    CopyMemory KeybDown(0), KeybData(0), 256
    pKeyb.GetDeviceState 256, KeybData(0)
    If Err.Number = DIERR_INPUTLOST Then
        Err.Number = 0
        pKeyb.Acquire
        pKeyb.GetDeviceState 256, KeybData(0)
        If Not Err.Number = 0 Then Exit Function
    End If
    On Local Error GoTo Failed
    'Write inputs to the string.
    If KeybData(DIK_LSHIFT) Or KeybData(DIK_RSHIFT) Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
    For i = 0 To 255
        If KeybData(i) And Not KeybDown(i) Then
            'This key is now new pressed. The input possibly must be appended to the string.
            RdSign = GetChar(i, ShiftDown)
            If ShiftDown Then RdSign = UCase$(RdSign)
            If KeybData(DIK_BACK) Then
                If Not Len(AppendTo) = 0 Then AppendTo = Left$(AppendTo, Len(AppendTo) - 1)
            Else
                AppendTo = AppendTo & RdSign
            End If
        End If
    Next i
    ReadKeyboard = VarPtr(KeybData(0))
Failed:
End Function

'//-----------------------------------------------------------------------------
'// Function: ReadMouse
'// Desc: Reads the pressed mouse buttons and the mouse coordinates.
'// Returns: The mouse desc.
'//-----------------------------------------------------------------------------
Public Function ReadMouse() As DIMOUSESTATE
    On Local Error Resume Next
    pMouse.GetDeviceState Len(MouseData), MouseData
    If Err.Number = DIERR_INPUTLOST Then
        Err.Number = 0
        pMouse.Acquire
        pMouse.GetDeviceState Len(MouseData), MouseData
        If Not Err.Number = 0 Then Exit Function
    End If
    ReadMouse = MouseData
End Function

'//*****Private functions*****
Private Function GetChar(ByVal KeyCode As Long, ByVal ShiftDown As Boolean) As String
    If ShiftDown Then
        If KeyCode = DIK_0 Then
            GetChar = "="
        ElseIf KeyCode = DIK_1 Then
            GetChar = "!"
        ElseIf KeyCode = DIK_2 Then
            GetChar = """"
        ElseIf KeyCode = DIK_3 Then
            GetChar = "§"
        ElseIf KeyCode = DIK_4 Then
            GetChar = "$"
        ElseIf KeyCode = DIK_5 Then
            GetChar = "%"
        'DIK_6 is the &, but it is reserved for shortcart.
        ElseIf KeyCode = DIK_7 Then
            GetChar = "/"
        ElseIf KeyCode = DIK_8 Then
            GetChar = "("
        ElseIf KeyCode = DIK_9 Then
            GetChar = ")"
        ElseIf KeyCode = DIK_MINUS Then
            GetChar = "?"
        ElseIf KeyCode = DIK_SLASH Then
            GetChar = "_"
        ElseIf KeyCode = DIK_COMMA Then
            GetChar = ";"
        ElseIf KeyCode = DIK_PERIOD Then
            GetChar = ":"
        ElseIf KeyCode = DIK_OEM_102 Then
            GetChar = ">"
        End If
    End If
    If KeyCode = DIK_A Then
        GetChar = "a"
    ElseIf KeyCode = DIK_B Then
        GetChar = "b"
    ElseIf KeyCode = DIK_C Then
        GetChar = "c"
    ElseIf KeyCode = DIK_D Then
        GetChar = "d"
    ElseIf KeyCode = DIK_E Then
        GetChar = "e"
    ElseIf KeyCode = DIK_F Then
        GetChar = "f"
    ElseIf KeyCode = DIK_G Then
        GetChar = "g"
    ElseIf KeyCode = DIK_H Then
        GetChar = "h"
    ElseIf KeyCode = DIK_I Then
        GetChar = "i"
    ElseIf KeyCode = DIK_J Then
        GetChar = "j"
    ElseIf KeyCode = DIK_K Then
        GetChar = "k"
    ElseIf KeyCode = DIK_L Then
        GetChar = "l"
    ElseIf KeyCode = DIK_M Then
        GetChar = "m"
    ElseIf KeyCode = DIK_N Then
        GetChar = "n"
    ElseIf KeyCode = DIK_O Then
        GetChar = "o"
    ElseIf KeyCode = DIK_P Then
        GetChar = "p"
    ElseIf KeyCode = DIK_Q Then
        GetChar = "q"
    ElseIf KeyCode = DIK_R Then
        GetChar = "r"
    ElseIf KeyCode = DIK_S Then
        GetChar = "s"
    ElseIf KeyCode = DIK_T Then
        GetChar = "t"
    ElseIf KeyCode = DIK_U Then
        GetChar = "u"
    ElseIf KeyCode = DIK_V Then
        GetChar = "v"
    ElseIf KeyCode = DIK_W Then
        GetChar = "w"
    ElseIf KeyCode = DIK_X Then
        GetChar = "x"
    ElseIf KeyCode = DIK_Y Then
        GetChar = "z"
    ElseIf KeyCode = DIK_Z Then
        GetChar = "y"
    ElseIf KeyCode = DIK_NUMPAD0 Or KeyCode = DIK_0 Then
        GetChar = "0"
    ElseIf KeyCode = DIK_NUMPAD1 Or KeyCode = DIK_1 Then
        GetChar = "1"
    ElseIf KeyCode = DIK_NUMPAD2 Or KeyCode = DIK_2 Then
        GetChar = "2"
    ElseIf KeyCode = DIK_NUMPAD3 Or KeyCode = DIK_3 Then
        GetChar = "3"
    ElseIf KeyCode = DIK_NUMPAD4 Or KeyCode = DIK_4 Then
        GetChar = "4"
    ElseIf KeyCode = DIK_NUMPAD5 Or KeyCode = DIK_5 Then
        GetChar = "5"
    ElseIf KeyCode = DIK_NUMPAD6 Or KeyCode = DIK_6 Then
        GetChar = "6"
    ElseIf KeyCode = DIK_NUMPAD7 Or KeyCode = DIK_7 Then
        GetChar = "7"
    ElseIf KeyCode = DIK_NUMPAD8 Or KeyCode = DIK_8 Then
        GetChar = "8"
    ElseIf KeyCode = DIK_NUMPAD9 Or KeyCode = DIK_9 Then
        GetChar = "9"
    ElseIf KeyCode = DIK_SPACE Then
        GetChar = " "
    ElseIf KeyCode = DIK_COMMA Or KeyCode = DIK_DECIMAL Then
        GetChar = ","
    ElseIf KeyCode = DIK_PERIOD Then
        GetChar = "."
    ElseIf KeyCode = DIK_SUBTRACT Then
        GetChar = "-"
    ElseIf KeyCode = DIK_MINUS Then
        GetChar = "ß"
    ElseIf KeyCode = DIK_ADD Then
        GetChar = "+"
    ElseIf KeyCode = DIK_MULTIPLY Then
        GetChar = "*"
    ElseIf KeyCode = DIK_SLASH Then
        GetChar = "-"
    ElseIf KeyCode = DIK_OEM_102 Then
        GetChar = "<"
    End If
End Function
