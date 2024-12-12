VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IPI_Paul_Outlook_Notes 
   Caption         =   "Oulook Item Notes"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   OleObjectBlob   =   "IPI_Paul_Outlook_Notes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IPI_Paul_Outlook_Notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ListStyle As String = "'line-height: 1; padding: 0px; margin-top: 0px; margin-bottom: 0px;'"
Private urlFollow As Boolean

Private Sub cmdBold_Click()
    InkEdit1.SelBold = Not InkEdit1.SelBold
End Sub

Private Sub cmdItalic_Click()
    InkEdit1.SelItalic = Not InkEdit1.SelItalic
End Sub

Private Sub cmdLink_Click()
    IPI_Paul_Outlook_Link.Show 1
    If linkHTML > "" Then getHTML linkHTML, InkEdit1
End Sub

Private Sub cmdOrderedList_Click()
    getHTML "<ol style=" & ListStyle & "><li style=" & ListStyle & "></li></ol>", InkEdit1
End Sub

Private Sub cmdSave_Click()
    setProperties InkEdit1.Text = "", InkEdit1.TextRTF
End Sub

Private Sub cmdSecureSave_Click()
    Dim olItem As Object, id As Variant, sbj As String, rcvd As Date, sent As Date, db As String, rtf As String, sql As String
    
    db = getDbLocation
    Set olItem = getOlItem
    
    With olItem
        id = .PropertyAccessor.BinaryToString(.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x300B0102"))
        rcvd = .ReceivedTime
        sent = .SentOn
        sbj = .Subject
        If Year(sent) = 4501 Then sent = rcvd
    End With
    
    rtf = InkEdit1.TextRTF
    
    getRecords id
    DoEvents
    
    If Not shrIPI.aRst.State = 0 Then
        If shrIPI.aRst.Fields(0).Value = 0 And InkEdit1.Text > "" Then
            If cboSecureType = "ACCDB" Then
                sql = "INSERT INTO [Notes] ([EmailId], [DateReceived], [DateSent], [Subject], [EmailNotes]) " & _
                    "SELECT '" & id & "', '" & rcvd & "', '" & sent & "', '" & sbj & "', ?;"
            Else
                sql = "INSERT INTO [Notes] ([EmailId], [DateReceived], [DateSent], [Subject], [EmailNotes]) " & _
                    "SELECT '" & id & "', CAST('" & Format(rcvd, "yyyy-mm-dd hh:mm:ss") & "' AS datetime), CAST('" & Format(sent, "yyyy-mm-dd hh:mm:ss") & "' AS datetime), '" & sbj & "', ?;"
            End If
        Else
            If InkEdit1.Text > "" Then
                sql = "UPDATE [Notes] SET [EmailNotes] = ? WHERE [EmailId] = '" & id & "';"
            Else
                sql = "DELETE FROM [Notes] WHERE [EmailId] = '" & id & "';"
            End If
        End If
        
        If cboSecureType = "ACCDB" Then
            adoMsAccess sql, db, "", "", "Parameterised", rtf
        Else
            adoSQL sql, "OutlookNotes", "Parameterised", rtf
        End If
        DoEvents
        
        getRecords id
        DoEvents
        
        If Not shrIPI.aRst.Fields(0).Value = 0 And InkEdit1.Text > "" Then
            setProperties False, "[SECURE " & cboSecureType & "] " & id
        ElseIf InkEdit1.Text > "" Then
            MsgBox "Failed to securely save, please try again!", vbCritical, "Secure Save Failure"
        Else
            setProperties True, ""
        End If
    End If
End Sub

Private Sub cmdUnderline_Click()
    InkEdit1.SelUnderline = Not InkEdit1.SelUnderline
End Sub

Private Sub cmdUnorderedList_Click()
    getHTML "<ul style=" & ListStyle & "><li style=" & ListStyle & "></li></ul>", InkEdit1
End Sub

Private Sub cmdViewNotes_Click()
    On Error GoTo errHere
    
    IPI_Paul_Outlook_Secure_Notes.Show 0
exitHere:
    Exit Sub
errHere:
    MsgBox "No Records Found!"
End Sub

Private Sub getRecords(id As Variant)
    Dim sql As String, db As String
    
    db = getDbLocation
    
    sql = "SELECT COUNT(*) AS [Items] FROM [Notes] WHERE [EmailId] = '" & id & "'"
    If cboSecureType.Value = "ACCDB" Then
        adoMsAccess sql, db, "OLEDB", "", "OLEDB"
    Else
        adoSQL sql, "OutlookNotes"
    End If
    DoEvents
    
    If Not shrIPI.aRst.State = 0 Then
        shrIPI.aRst.MoveFirst
        DoEvents
    End If
End Sub

Private Sub InkEdit1_DblClick()
    urlFollow = True
End Sub

Private Sub InkEdit1_KeyUp(pKey As Long, ByVal ShiftKey As Integer)
    If ShiftKey = 2 Then
        Select Case pKey
            Case 66
                InkEdit1.SelBold = Not InkEdit1.SelBold
            Case 84
                InkEdit1.SelItalic = Not InkEdit1.SelItalic
            Case 85
                InkEdit1.SelUnderline = Not InkEdit1.SelUnderline
        End Select
    End If
End Sub

Private Sub InkEdit1_SelChange()
    On Error GoTo exitHere
    
    If urlFollow Then
        ShellExecute 0&, "open", Split(InkEdit1.SelText, """")(1), 0, 0, 1
        urlFollow = False
    End If
exitHere:
    Exit Sub
End Sub

Private Sub setProperties(Optional clear As Boolean = False, Optional rtf As String)
    Dim olItem As Object
    
    Set olItem = getOlItem
    
    With olItem
        If .UserProperties.Item("olIPINotes") Is Nothing Then
            .UserProperties.Add "olIPINotes", olText
            .UserProperties.Add "Has Notes", olText
        End If
        If clear Then
            .UserProperties.Item("olIPINotes").Value = ""
            .UserProperties.Item("Has Notes").Value = ""
        Else
            .UserProperties.Item("olIPINotes").Value = rtf
            .UserProperties.Item("Has Notes").Value = "Yes"
        End If
        .Save
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim olItem As Object, sql As String, db As String, myItem As Outlook.MailItem, txt As String, dl As Boolean
    
    cboSecureType.AddItem "ACCDB"
    cboSecureType.AddItem "SQL"
    cboSecureType.ListIndex = 1
    
    Set olItem = getOlItem
    db = getDbLocation
    
    With olItem
        id = .PropertyAccessor.BinaryToString(.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x300B0102"))
        If Not .UserProperties.Item("olIPINotes") Is Nothing Then
            If .UserProperties.Item("olIPINotes").Value > "" Then
                If InStr(1, .UserProperties.Item("olIPINotes").Value, "[SECURE") = 0 Then
                    txt = .UserProperties.Item("olIPINotes").Value
                    InkEdit1.TextRTF = txt
                    GoTo exitHere
                ElseIf InStr(1, .UserProperties.Item("olIPINotes").Value, "[SECURE") > 0 Then
                    If InStr(1, .UserProperties.Item("olIPINotes").Value, "SECURE ACCDB") > 0 Then
                        sql = "SELECT [EmailNotes] FROM [Notes] WHERE [EmailId] = '" & id & "'"
                        adoMsAccess sql, db, "OLEDB", "", "OLEDB"
                    Else
                        sql = "SELECT CAST([EmailNotes] AS text) AS 'EmailNotes' FROM [Notes] WHERE [EmailId] = '" & id & "'"
                        adoSQL sql, "OutlookNotes"
                    End If
                    DoEvents
                    
                    If shrIPI.aRst.State = 0 Then
                        dl = True
                    ElseIf shrIPI.aRst.BOF And shrIPI.aRst.EOF Then
                        dl = True
                    End If
                    
                    If dl Then Exit Sub
                    
                    shrIPI.aRst.MoveFirst
                    DoEvents
                    
                    If InStr(1, shrIPI.aRst.Fields(0).Value, "{\rtf") = 0 Then
                        txt = "<html><head><style>body{line-height:1;font-family:Arial;font-size:11pt;}</script></head><body>" & shrIPI.aRst.Fields(0).Value & "</body></html>"
                        getHTML txt, InkEdit1
                    Else
                        InkEdit1.TextRTF = shrIPI.aRst.Fields(0).Value
                    End If
                End If
            End If
        End If
    End With
exitHere:
    InkEdit1.SetFocus
End Sub
