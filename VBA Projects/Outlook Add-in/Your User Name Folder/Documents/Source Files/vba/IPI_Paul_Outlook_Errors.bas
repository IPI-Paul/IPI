Attribute VB_Name = "IPI_Paul_Outlook_Errors"

Public Sub ShowErrMsg()
    Dim txt As String
    
    txt = "Desc: " & Err.Description & vbNewLine
    txt = txt & "Err #: " & Err.Number
    MsgBox txt, vbExclamation, "Run-Time Error"
    End
End Sub

