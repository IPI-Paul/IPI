Attribute VB_Name = "IPI_Paul_Outlook_LocalStorage"
Sub appendToFile(fPath As String, iStr As String)
    Dim fso As New FileSystemObject, ts As TextStream
    
    Set ts = fso.OpenTextFile(fPath, ForAppending, True)
    
    With ts
        .Write iStr
        .Close
    End With
    
    Set ts = Nothing
End Sub

Function readFile(fPath As String) As String
    Dim fso As New FileSystemObject, ts As TextStream, iStr As String
    
    Set ts = fso.OpenTextFile(fPath, ForReading, True)
    
    With ts
        iStr = .ReadAll
        .Close
    End With
    
    readFile = iStr
    Set ts = Nothing
End Function

Sub updateFile(fPath As String, iStr As String)
    Dim fso As New FileSystemObject
        
    Open fPath For Output As #1
        Print #1, iStr
    Close #1
End Sub

Sub viewInNotepad(iStr As String)
    Dim fso As New FileSystemObject, fPath As String
        
    fPath = Environ$("appdata") & "\IPI Paul\Outlook\Viewer\Temp.txt"
    If Not fso.FolderExists(Split(fPath, "Outlook", 2)(0)) Then MkDir Split(fPath, "Outlook", 2)(0)
    If Not fso.FolderExists(Split(fPath, "Viewer", 2)(0)) Then MkDir Split(fPath, "Viewer", 2)(0)
    If Not fso.FolderExists(Split(fPath, "Temp", 2)(0)) Then MkDir Split(fPath, "Temp", 2)(0)
        
    Open fPath For Output As #1
        Print #1, iStr
    Close #1
    
    loadApplicationLink "Notepad", fPath
End Sub
