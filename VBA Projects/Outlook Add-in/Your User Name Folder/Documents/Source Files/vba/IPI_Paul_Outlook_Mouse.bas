Attribute VB_Name = "IPI_Paul_Outlook_Mouse"
Public Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As LongPtr
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As LongPtr
Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public crCoord As POINTAPI
