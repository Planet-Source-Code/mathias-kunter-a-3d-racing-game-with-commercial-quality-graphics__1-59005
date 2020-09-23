VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'Kein
   Caption         =   "Revo Tron"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   2205
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'*******************************************************************************
'*******************************************************************************
'
'Here we go again...
'Well, turn your speakers on, hit F5 and don't believe your eyes :)
'However, it's highly recommended that you compile this game before running it.
'
'*******************************************************************************
'*******************************************************************************




'*******************************************************************************
'*******************************************************************************
'
'REVO TRON v 1.4 (Visual Basic)
'http://revotron.tripod.com
'
'This is Revo Tron version 1.4 for Visual Basic 6. It has been ported from C++
'to VB to show that commercial quality 3d is possible in VB. The 3d engine
'is included in this project, check out the cls3dEngine class. To simplify things,
'sound and network support wasn't ported. You can check out http://revotron.tripod.com
'to download the compiled C++ exe and the source code.
'
'Copyright by Mathias Kunter. You can mail me at mathiaskunter@yahoo.de
'
'*******************************************************************************
'*******************************************************************************



Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long



Private Sub Form_Load()
    On Local Error Resume Next
    
    MsgBox "Notice that you should compile this game to be able to play it well.", vbInformation
    
    Me.Show
    DoEvents
    
    'Play the start sound.
    OpenFile "revotron.mp3", 0
    PlayFile 0
    mdlTron.Run_The_Whole_Damn_Show
    CloseFile 0
    End
End Sub


Private Function OpenFile(ByVal sFile As String, ByVal iAlias As Integer) As Boolean
    Dim lPathLen As Long, sShortPath As String
    
    'Get the short path name of the mp3 file.
    lPathLen = GetShortPathName(sFile, vbNull, 0)
    If lPathLen = 0 Then Exit Function
    sShortPath = String$(lPathLen, Chr$(0))
    GetShortPathName sFile, sShortPath, lPathLen
    sShortPath = Left$(sShortPath, Len(sShortPath) - 1)
    
    'Close the previous opened file, if there is one.
    CloseFile iAlias
    'Open the mp3 file.
    If Not mciSendString("open " & sShortPath & " type MPEGVideo alias mp3" & iAlias, vbNull, 0, 0) = 0 Then Exit Function
    mciSendString "set mp3" & iAlias & " time format milliseconds", vbNull, 0, 0
    OpenFile = True
End Function

Private Sub PlayFile(ByVal iAlias As Integer)
    mciSendString "play mp3" & iAlias, vbNull, 0, 0
End Sub

Private Sub CloseFile(ByVal iAlias As Integer)
    mciSendString "close mp3" & iAlias, vbNull, 0, 0
End Sub
