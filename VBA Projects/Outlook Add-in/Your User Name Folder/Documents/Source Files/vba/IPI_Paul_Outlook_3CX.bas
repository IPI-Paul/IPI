Attribute VB_Name = "IPI_Paul_Outlook_3CX"
Sub callNumber(num As String)
    Dim cmd As String
    
    If Len(num) > 1 Then
        For Each itm In Array("(", ")", " ")
            txt = Replace(num, itm, "")
        Next
        app = "C:\ProgramData\3CXPhone for Windows\PhoneApp\calltriggercmd.exe"
        param = " -cmd makecall:"
        cmd = """" & app & """" & param & num
        runShell cmd
    End If
End Sub

Sub callSelectedNumber()
    Dim num As String, olItem As Object
    
    Set olItem = getOlItem
    num = olItem.GetInspector.WordEditor.Application.Selection.Text
    callNumber num
    Set olItem = Nothing
End Sub
