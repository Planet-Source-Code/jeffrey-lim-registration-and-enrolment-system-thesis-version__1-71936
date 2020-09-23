Attribute VB_Name = "modAPI"

Option Explicit

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'For Transparent Form
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Public Function isTransparent(ByVal hwnd As Long) As Boolean
    On Error Resume Next
    Dim MSG As Long
    MSG = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (MSG And WS_EX_LAYERED) = WS_EX_LAYERED Then
      isTransparent = True
    Else
      isTransparent = False
    End If
    If err Then
      isTransparent = False
    End If
End Function

Public Function MakeTransparent(ByVal hwnd As Long, intOpacity As Integer) As Long
    Dim MSG As Long
    On Error Resume Next
    If intOpacity < 0 Or intOpacity > 255 Then
      MakeTransparent = 1
    Else
      MSG = GetWindowLong(hwnd, GWL_EXSTYLE)
      MSG = MSG Or WS_EX_LAYERED
      SetWindowLong hwnd, GWL_EXSTYLE, MSG
      SetLayeredWindowAttributes hwnd, 0, intOpacity, LWA_ALPHA
      MakeTransparent = 0
    End If
    If err Then
      MakeTransparent = 2
    End If
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
    Dim MSG As Long
    On Error Resume Next
    MSG = GetWindowLong(hwnd, GWL_EXSTYLE)
    MSG = MSG And Not WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, MSG
    SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
    MakeOpaque = 0
    If err Then
      MakeOpaque = 2
    End If
End Function
'-- eo: Form Transparency

