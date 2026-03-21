Attribute VB_Name = "IPI_Paul_Excel_Shell"
Option Explicit

Sub loadChromeURL(url As String)
    Dim hWnd As LongPtr
    hWnd = ShellExecute(0, "Open", "Chrome.exe", url, "", 1)
    hWnd = vbNull
End Sub

Sub loadExplorerLink(fPath As Variant)
    Dim hWnd As LongPtr
    hWnd = ShellExecute(0, "Open", "Explorer.exe", fPath, "", 1)
    hWnd = vbNull
End Sub

Sub loadApplicationLink(app As Variant, fPath As Variant)
    Dim hWnd As LongPtr
    hWnd = ShellExecute(0, "Open", app, fPath, "", 1)
    hWnd = vbNull
End Sub

Sub runShell(cmd As String, Optional strComputer As String = ".")
    Dim objWMIService As Object
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
    objWMIService.Create cmd
    DoEvents
    waitTill "00:00:02"
End Sub

