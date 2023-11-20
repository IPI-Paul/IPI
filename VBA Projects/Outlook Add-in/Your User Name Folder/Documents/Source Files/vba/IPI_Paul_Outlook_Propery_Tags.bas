Attribute VB_Name = "IPI_Paul_Outlook_Propery_Tags"
Public Const PR_RECEIVED_BY_NAME As String = "http://schemas.microsoft.com/mapi/proptag/0x0040001E"
Public Const PR_SENT_REPRESENTING_NAME As String = "http://schemas.microsoft.com/mapi/proptag/0x0042001E"
Public Const PR_RECEIVED_BY_EMAIL_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x0076001E"
Public Const PR_SENT_REPRESENTING_EMAIL_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x0065001E"
Public Const PR_SENDER_EMAIL_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E"
Public Const PR_REPLY_RECIPIENT_NAMES As String = "http://schemas.microsoft.com/mapi/proptag/0x0050001E"
Public Const PR_SENDER_NAME As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E"
Public Const PR_DISPLAY_BCC As String = "http://schemas.microsoft.com/mapi/proptag/0x0E02001E"
Public Const PR_DISPLAY_CC As String = "http://schemas.microsoft.com/mapi/proptag/0x0E03001E"
Public Const PR_DISPLAY_TO As String = "http://schemas.microsoft.com/mapi/proptag/0x0E04001E"
Public Const PR_PRIMARY_SEND_ACCT As String = "http://schemas.microsoft.com/mapi/proptag/0x0E28001E"
Public Const PR_NEXT_SEND_ACC As String = "http://schemas.microsoft.com/mapi/proptag/0x0E29001E"
Public Const PR_BODY As String = "http://schemas.microsoft.com/mapi/proptag/0x1000001E"
Public Const PR_HTML As String = "http://schemas.microsoft.com/mapi/proptag/0x10130102"
Public Const PR_SENDER_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
Public Const RECIP As String = PR_RECEIVED_BY_EMAIL_ADDRESS & "," & PR_DISPLAY_TO & "," & PR_RECEIVED_BY_NAME & "," & PR_DISPLAY_CC & "," & PR_DISPLAY_BCC & "," & PR_REPLY_RECIPIENT_NAMES
Const SENDR As String = PR_SENDER_EMAIL_ADDRESS & "," & PR_SENT_REPRESENTING_EMAIL_ADDRESS & "," & PR_PRIMARY_SEND_ACCT & "," & PR_NEXT_SEND_ACC & "," & PR_SENDER_NAME & "," & _
    PR_SENT_REPRESENTING_NAME
Const PROPTAGS As String = RECIP & "," & SENDR & "," & PR_BODY & "," & PR_HTML
Public Enum PrTag
    RECEIVED_BY_EMAIL_ADDRESS
    DISPLAY_TO
    RECEIVED_BY_NAME
    DISPLAY_CC
    DISPLAY_BCC
    REPLY_RECIPIENT_NAMES
    SENDER_EMAIL_ADDRESS
    SENT_REPRESENTING_EMAIL_ADDRESS
    PRIMARY_SEND_ACCT
    NEXT_SEND_ACC
    SENDER_NAME
    SENT_REPRESENTING_NAME
    PR__BODY
    PR__HTML
End Enum
Public Enum rPrTag
    rRECEIVED_BY_EMAIL_ADDRESS
    rDISPLAY_TO
    rRECEIVED_BY_NAME
    rDISPLAY_CC
    rDISPLAY_BCC
    rREPLY_RECIPIENT_NAMES
End Enum
Public Enum sPrTag
    sSENDER_EMAIL_ADDRESS
    sSENT_REPRESENTING_EMAIL_ADDRESS
    sPRIMARY_SEND_ACCT
    sNEXT_SEND_ACC
    sSENDER_NAME
    sSENT_REPRESENTING_NAME
End Enum
Public Const wdOrientLandscape As Integer = 1, wdOrientPortrait As Integer = 0, wdActiveEndPageNumber As Integer = 3
Public Const wdBorderLeft As Integer = -2, wdBorderRight As Integer = -4, wdBorderVertical As Integer = -6, wdLineStyleNone As Integer = 0, wdLineStyleSingle As Integer = 1
Public Const wdAlignParagraphRight As Integer = 2, wdAlignParagraphCenter As Integer = 1, wdAlignParagraphLeft As Integer = 0

Public Function getPropTag(Optional idx As Variant = "") As Variant
    On Error GoTo errHere
    
    If idx = "" Then
        getPropTag = Split(PROPTAGS, ",")
    Else
        getPropTag = Split(PROPTAGS, ",")(idx)
    End If
    
errHere:
    If Err.Number Then ShowErrMsg
End Function

Public Function getPropTagR(Optional idx As Variant = "") As Variant
    On Error GoTo errHere
    
    If idx = "" Then
        getPropTagR = Split(RECIP, ",")
    Else
        getPropTagR = Split(RECIP, ",")(idx)
    End If
    
errHere:
    If Err.Number Then ShowErrMsg
End Function

Public Function getPropTagS(Optional idx As Variant = "") As Variant
    On Error GoTo errHere
    
    If idx = "" Then
        getPropTagS = Split(SENDR, ",")
    Else
        getPropTagS = Split(SENDR, ",")(idx)
    End If
    
errHere:
    If Err.Number Then ShowErrMsg
End Function

