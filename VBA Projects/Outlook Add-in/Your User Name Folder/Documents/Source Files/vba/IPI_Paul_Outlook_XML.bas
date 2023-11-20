Attribute VB_Name = "IPI_Paul_Outlook_XML"
Function getHyperlinks(olItem As Object) As ADODB.Recordset
    Dim http As MSXML2.XMLHTTP, htmlDoc As New MSHTML.HTMLDocument, dom As Object, elements As Object, coll As Object, rs As New ADODB.Recordset, flds As Variant
    
    Set coll = CreateObject("System.Collections.ArrayList")
    flds = Array("HRef", "Text")
    
    With rs
        Set .ActiveConnection = Nothing
        .CursorLocation = adUseClient '3
        .LockType = adLockBatchOptimistic '4
        For Each fld In flds
            With .Fields
                .Append fld, adVarChar, 255 '200,255
            End With
        Next fld
        .Open
    End With
    
    htmlDoc.Body.innerHTML = olItem.HTMLBody
    
    Set elements = htmlDoc.getElementsByTagName("a")
    For Each e In elements
        If Not coll.contains(e) Then
            coll.add e
            rs.AddNew flds, Array(e, e.outerText)
        End If
    Next e
    
    Set elements = htmlDoc.getElementsByTagName("img")
    For Each e In elements
        If Not coll.contains(e.src) Then
            coll.add e.src
            rs.AddNew flds, Array(e.src, e.outerText)
        End If
    Next e
    
    Set elements = htmlDoc.getElementsByTagName("imgagedata")
    For Each e In elements
        If Not coll.contains(e.src) Then
            coll.add e.src
            rs.AddNew flds, Array(e.src, e.outerText)
        End If
    Next e
    Set getHyperlinks = rs
    
exitHere:
    Set coll = Nothing
    Set elements = Nothing
End Function

Function getItemLink(olItem As Object, iStr As String) As String
    Dim rs As New ADODB.Recordset
    
    Set rs = getHyperlinks(olItem)
    With rs
        .Filter = "[Text] = '" & iStr & "'"
        If Not .BOF Then .MoveFirst
        If !hRef > "" Then getItemLink = !hRef
    End With

exitHere:
    Set rs = Nothing
End Function

Function getSimpleHTML() As String
    Dim olItem As Object, strHTML As String, http As New MSXML2.XMLHTTP, htmlDoc As New MSHTML.HTMLDocument, regExp As Variant
    
    regExp = Array(Array("o:p", "p"), Array("<B>", "<b>"), Array("</B>", "</b>"), Array("<BR>", "<br />"), _
        Array("<SPAN style=""mso-bookmark: _MailOriginal""><SPAN style='FONT-SIZE: 11PT; FONT-FAMILY: ""Calibri"",sans-serif; COLOR: #lf497d; mso-fareast-language: EN-US'>" & _
        "<?xml:namespace prefix = ""o"" ns = ""urn:schemas-microsoft-com:office:office"" /><p>&nbsp;</p></SPAN></SPAN>", ""), _
        Array("<?xml:namespace prefix = ""o"" ns = ""urn:schemas-microsoft-com:office:office"" />", ""), _
        Array("style=""mso-bookmark: _MailOriginal""", ""), Array("<SPAN ><SPAN", "<span"), Array("</SPAN></SPAN>", "</span>"), Array("<SPAN >", "<span>"), Array("</SPAN>", "</span>"))
    
    htmlDoc.Body.innerHTML = olItem.HTMLBody
    
    strHTML = ""
    For Each span In htmlDoc.getElementsByTagName("span")
        If InStr(span.outerHTML, "mso-bookmark: _MailOriginal") > 0 And Not InStr(1, Replace(span.outerHTML, "o:p", "p"), "<p>&nbsp;</p>") > 0 Then
            For Each xp In regExp
                strHTML = strHTML & Replace(span.outerHTML, xp(0), xp(1))
            Next xp
        End If
    Next
    getSimpleHTML = strHTML
End Function

Function getTextContent(olItem As Object) As String
    Dim http As New MSXML2.XMLHTTP, htmlDoc As New MSHTML.HTMLDocument
    
    htmlDoc.Body.innerHTML = olItem.HTMLBody
    getTextContent = htmlDoc.Body.outerText
End Function

Function removeLinks(txt As String, rs As ADODB.Recordset, Optional incl As Boolean = False) As String
    Dim regExp As Object, hLink As String, hText As String, hDoc As String, oMatch As match
    
    Set regExp = CreateObject("VBScript.RegExp")
    If Not rs.BOF Then rs.MoveFirst
    Do While Not rs.EOF
        hLink = rs!hRef
        hText = rs!Text
        If InStr(1, hLink, "/") > 0 Then
            hDoc = Split(hLink, "/")(UBound(Split(hLink, "/")))
        ElseIf InStr(1, hLink, "\") > 0 Then
            hDoc = Split(hLink, "\")(UBound(Split(hLink, "\")))
        Else
            hDoc = hLink
        End If
        rs.MoveNext
    Loop
    
    For Each regEx In Array("<", ">", "%20", "|")
        txt = Replace(txt, regEx, "")
    Next
    
    For Each pat In Array("\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b", "\b(https?|ftps|file)://[-A-Z0-9+&@#/%?=~_|$!:,.;]*[A-Z0-9+&@#/%=~_|$]", _
            "\b(http?|ftp?|file)://[-A-Z0-9+&@#/%?=~_|$!:,.;]*[A-Z0-9+&@#/%=~_|$]", "\bwww.[-A-Z0-9+&@#/%?=~_|$!:,.;]*[A-Z0-9+&@#/%=~_|$]", _
            "(?:^|(?:[^-a-zA-Z0-9_]))@([A-Za-z]+[A-Z-a-z0-9_]+)")
        With regExp
            .Pattern = pat
            .IgnoreCase = True
            .MultiLine = True
            .Global = True
            Set olMatches = .Execute(txt)
            
            For Each oMatch In olMatches
                txt = Replace(txt, oMatch, "")
            Next oMatch
        End With
    Next pat
    
    removeLinks = txt
    Set regExp = Nothing
End Function
