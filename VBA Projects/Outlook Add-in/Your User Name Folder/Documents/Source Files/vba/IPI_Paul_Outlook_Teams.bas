Attribute VB_Name = "IPI_Paul_Outlook_Teams"
Sub callInTeams()
    Dim sip As String
    
    sip = getTeamsContact
    If sip > "" Then
        runShell getTeams & " sip:" & sip
        DoEvents
        SendKeys "(^+c)", True
        SendKeys "{NUMLOCK}", True
    End If
End Sub

Sub chatInTeams()
    Dim sip As String
    
    sip = getTeamsContact
    If sip > "" Then
        runShell getTeams & " sip:" & sip
        DoEvents
    End If
End Sub

Sub createTeamsMeeting()
    Dim sip As String
    
    sip = getTeamsContact
    If sip > "" Then
        runShell getTeams
        DoEvents
        SendKeys "(^4)", True
        waitTill
        SendKeys "(%+n)", True
        waitTill
        SendKeys getOutlookSubject, True
        waitTill
        SendKeys "{Tab}", True
        waitTill
        SendKeys sip, True
        waitTill "00:00:03"
        SendKeys "{Enter}", True
        SendKeys "{NUMLOCK}", True
    End If
End Sub

Private Function getTeamsContact() As String
    Dim olItem As Object, sip As String
    
    On Error Resume Next
    getTeamsContact = ""
    
    Set olItem = getOlItem
    sip = olItem.PropertyAccessor.GetProperty(PR_SENDER_SMTP_ADDRESS)
    If sip = "noreply@email.teams.microfost.com" Then
        sip = Split(olItem.SENDER, "(")(0)
        sip = StrConv(Trim(Split(sip, ",")(1)) & " " & Trim(Split(sip, ",")(0)), vbProperCase)
        'lookup a database
    End If
    If sip = "" Then
        For Each itm In getPropTagS
            If sip = "" Then
                sip = olItem.PropertyAccessor.GetProperty(itm)
                Exit For
            End If
        Next
    End If
    If sip = "" Then sip = olItem.SenderEmailAddress
    If sip = "" Then sip = olItem.SENDER
    If sip > "" Then getTeamsContact = sip
    Set olItem = Nothing
End Function

Private Function getTeams()
    getTeams = Environ$("UserProfile") & "\appdata\Local\Microsoft\Teams\Current\Teams.exe"
End Function
