Attribute VB_Name = "IPI_Paul_Excel_Delcarations"
Option Explicit

Public Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As LongPtr
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParamaters As String, ByVal lpDirectory As String, ByVal nShowCmd As LongPtr) As LongPtr
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

