Attribute VB_Name = "IPI_Paul_Outlook_Shell"
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParamaters As String, ByVal lpDirectory As String, ByVal nShowCmd As LongPtr) As LongPtr

Sub loadChromeURL(url As String)
    ShellExecute 0, "Open", "Chrome.exe", url, "", 1
End Sub

Sub loadExplorerLink(fPath As Variant)
    ShellExecute 0, "Open", "Explorer.exe", fPath, "", 1
End Sub

Sub loadApplicationLink(app As Variant, fPath As Variant)
    ShellExecute 0, "Open", app, fPath, "", 1
End Sub

Sub runShell(cmd As String, Optional strComputer As String = ".")
    Dim objWMIService As Object
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
    objWMIService.Create cmd
    DoEvents
    waitTill "00:00:02"
End Sub
