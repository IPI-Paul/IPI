Attribute VB_Name = "IPI_Paul_Excel_LocalStorage"
Option Explicit

Sub appendToFile(fPath As String, iStr As String)
    Dim fso As New FileSystemObject, ts As TextStream, i As Long, j As Long

    Set ts = fso.OpenTextFile(fPath, ForAppending, True)

    With ts
        If Len(iStr) > 5000 Then
            For i = 1 To WorksheetFunction.RoundUp(Len(iStr) / 5000, 0)
                If i = 1 Then
                    j = 1
                Else
                    j = (i - 1) * 5000
                End If
'                For Each itm In Split(Mid(iStr, j, 5000), " ")
'                    'Debug.Print itm
'                    .write itm & " "
'                Next itm
                If i = WorksheetFunction.RoundUp(Len(iStr) / 5000, 0) And i * 5000 > Len(iStr) Then
                    .write Right(iStr, Len(iStr) - ((i - 1) * 5000))
                Else
                    .write Mid(iStr, j, 5000)
                End If
            Next i
        Else
            .write iStr
        End If
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

Sub streamToFile(fPath As String, iStr As String)
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 2
        .writetext CStr(iStr)
        .savetofile fPath, 2
        .Close
    End With
End Sub

Sub updateFile(fPath As String, iStr As String)
    Dim fso As New FileSystemObject
        
    Open fPath For Output As #1
        Print #1, iStr
    Close #1
End Sub

Sub viewInNotepad(iStr As String)
    Dim fso As New FileSystemObject, fPath As String
        
    fPath = Environ$("appdata") & "\IPI Paul\Excel\Viewer\Temp.txt"
    If Not fso.FolderExists(Split(fPath, "Excel", 2)(0)) Then MkDir Split(fPath, "Excel", 2)(0)
    If Not fso.FolderExists(Split(fPath, "Viewer", 2)(0)) Then MkDir Split(fPath, "Viewer", 2)(0)
    If Not fso.FolderExists(Split(fPath, "Temp", 2)(0)) Then MkDir Split(fPath, "Temp", 2)(0)
        
    Open fPath For Output As #1
        Print #1, iStr
    Close #1
    
    loadApplicationLink "Notepad", fPath
End Sub

