Attribute VB_Name = "IPI_Paul_Outlook_Declarations"
Public Declare PtrSafe Function PasteToControl Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, Optional ByVal wMsg As Long = &H302, _
    Optional ByVal wParam As Long = 0, Optional lParam As Any = 0&) As Long
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal plOperations As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As LongPtr) As LongPtr
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)

Public shrIPI As New IPI_Paul_Outlook_Shared, linkHTML As String
