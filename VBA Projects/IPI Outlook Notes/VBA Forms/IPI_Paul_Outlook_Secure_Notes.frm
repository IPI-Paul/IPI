VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IPI_Paul_Outlook_Secure_Notes 
   Caption         =   "Outlook Secure Notes"
   ClientHeight    =   10545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   OleObjectBlob   =   "IPI_Paul_Outlook_Secure_Notes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IPI_Paul_Outlook_Secure_Notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private emailId As String

Private Sub cmdViewEmail_Click()
    If emailId > "" Then findAndOpen "olIPINotes like '%" & emailId & "%'"
End Sub

Private Sub UserForm_Initialize()
    Dim db As String, sql As String
    
    db = getDbLocation & ""
    
    If IPI_Paul_Outlook_Notes.Controls("cboSecureType").Value = "ACCDB" Then
        sql = "SELECT EmailId, DateReceived AS [Date Received], DateSent AS [Date Sent], Subject FROM [Notes];"
        adoMsAccess sql, db, "OLEDB", "", "OLEDB"
    Else
        sql = "SELECT EmailId, DateReceived AS 'Date Received', DateSent AS 'Date Sent', Subject FROM Notes;"
        adoSQL sql, "OutlookNotes"
    End If
    DoEvents
    
    If Not shrIPI.aRst.State = 0 Then shrIPI.aRst.MoveFirst
    DoEvents
    
    With wbNotes
        .Navigate2 "about.htm"
        Do While .Busy Or .ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
        If Not shrIPI.aRst.State = 0 Then
            .Document.Write Join( _
                buildHTML( _
                    shrIPI.aRst, _
                    sCol:=Array(Array("EmailId", "display:none;")), _
                    hCol:=Array(Array("EmailId", "display:none;")), _
                    getTable:=True _
                ), _
                vbCrLf _
            )
        Else
            .Document.Write "No Records Found"
        End If
        .Document.Close
        DoEvents
    End With
End Sub

Private Sub wbNotes_TitleChange(ByVal Text As String)
    Dim txt As String, i As Integer, vl As String, itm As Variant, id As Integer, sql As String
    
    If wbNotes.ReadyState = READYSTATE_COMPLETE Then
        If wbNotes.Document.Title > "" Then
            i = 0
            vl = ""
            emailId = ""
            For Each itm In Split(wbNotes.Document.Title, ",")
                If CStr(itm) > "" Then
                    If i = 0 Then
                        id = Int(itm)
                    ElseIf i = 1 Then
                        If vl > "" Then vl = vl & ","
                        vl = vl & CStr(itm)
                    ElseIf id = 6 And i = 2 Then
                        emailId = Trim(CStr(itm))
                    End If
                    i = i + 1
                End If
            Next
            InkEdit1.TextRTF = ""
            If id = 6 Then
                db = getDbLocation
                If IPI_Paul_Outlook_Notes.Controls("cboSecureType").Value = "ACCDB" Then
                    sql = "SELECT [EmailNotes] FROM [Notes] WHERE [EmailId] = '" & emailId & "';"
                    adoMsAccess sql, db, "OLEDB", "", "OLEDB"
                Else
                    sql = "SELECT CAST([EmailNotes] AS text) AS 'EmailNotes' FROM [Notes] WHERE [EmailId] = '" & emailId & "';"
                    adoSQL sql, "OutlookNotes"
                End If
                DoEvents
                
                shrIPI.aRst.MoveFirst
                DoEvents
                
                If InStr(1, shrIPI.aRst.Fields(0).Value, "{\rtf") = 0 Then
                    txt = "<html><head><style>body{line-height:1;font-family:Arial;font-size:11pt;}</script></head><body>" & shrIPI.aRst.Fields(0).Value & "</body></html>"
                    getHTML txt, InkEdit1
                Else
                    InkEdit1.TextRTF = shrIPI.aRst.Fields(0).Value
                End If
                InkEdit1.SetFocus
            End If
        End If
    End If
End Sub
