Attribute VB_Name = "IPI_Paul_Outlook_HTML"
Public frmLOut As FORMLAYOUT, ie(0 To 20) As IPI_Paul_Outlook_WebDocument, Document As New HTMLDocument
Public Const hdrBg As String = "RGB(0, 94, 184)", hdrFg As String = "RGB(255, 255, 255)"
Public Const tb1 As String = "  ", tb2 As String = tb1 & "  ", tb3 As String = tb2 & "  ", tb4 As String = tb3 & "  ", tb5 As String = tb4 & "  ", tb6 As String = tb5 & "  ", tb7 As String = tb6 & "  "
Public Const tblBg As String = "<table id=""tblDtl"" border=""1"">" & vbCrLf, tblEn As String = "</table></br>" & vbCrLf
Public Const thBg As String = tb2 & "<th>" & vbCrLf, thEnd As String = tb2 & "</th>" & vbCrLf
Public Const trBg As String = tb1 & "<tr onDblClick='highLt(rowIndex)' >" & vbCrLf, trEn As String = tb1 & "</tr>" & vbCrLf
Public Const tdBg As String = tb2 & "<td>" & vbCrLf, tdEn As String = tb2 & "</td>" & vbCrLf
Public Const stl As String = "<style id = ""tblStl"">" & vbCrLf & _
    tb1 & "table, body, input, select {" & vbCrLf & tb2 & "font-size:8pt;" & vbCrLf & tb1 & "}" & vbCrLf & _
    tb1 & "table {" & vbCrLf & tb2 & "border-collapse: collapse;" & vbCrLf & tb1 & "}" & vbCrLf & _
    tb1 & "th {" & vbCrLf & tb2 & "background-color: " & hdrBg & ";" & vbCrLf & tb2 & "color: " & hdrFg & ";" & vbCrLf & tb1 & "}" & vbCrLf & _
    tb1 & "td, th {" & vbCrLf & tb2 & "padding: 2px 7px 2px 7px;" & vbCrLf & tb1 & "}" & vbCrLf & _
    tb1 & "input, select {" & vbCrLf & tb2 & "outline: hidden;" & vbCrLf & tb2 & "border-width: 0 0 1px;" & vbCrLf & tb1 & "}" & vbCrLf & _
    "</style>" & vbCrLf

Private Function cboAct(Optional cols As Integer = 5) As String
    Dim sHTML As String
    
    sHTML = trBg & trEn
    sHTML = sHTML & trBg
    sHTML = sHTML & tb2 & "<td colspan=""" & cols & """ align=""right"">" & vbCrLf
    sHTML = sHTML & tb3 & "<select id=""selRun"" onChange=""runScript(id, value)"" class=""selClass"">" & vbCrLf
    sHTML = sHTML & tb4 & "<option value="""">Please select a function</option>" & vbCrLf
    sHTML = sHTML & tb4 & "<option value=""1"">Send Selection To Email</option>" & vbCrLf
    sHTML = sHTML & tb4 & "<option value=""2"">Send Selection To Word</option>" & vbCrLf
    sHTML = sHTML & tb4 & "<option value=""3"">User Defined Properties Clear</option>" & vbCrLf
    sHTML = sHTML & tb4 & "<option value=""4"">User Defined Properties Delete</option>" & vbCrLf
    sHTML = sHTML & tb4 & "<option value=""5"">User Defined Properties Save Map</option>" & vbCrLf
    sHTML = sHTML & tb4 & "<option value=""6"">User Defined Properties Update</option>" & vbCrLf
    sHTML = sHTML & tb3 & "</select>" & vbCrLf
    sHTML = sHTML & tdEn
    sHTML = sHTML & trEn
    cboAct = sHTML
End Function

Function buildPage(pgSet As WebPgSet)
    On Error GoTo errHere
    Dim val As Variant, rw As Integer, itms As Integer, fld As ADODB.Field, tr As String, iTxt As String, idx As Long
    
    tr = trBg
    rw = 1
    itms = 0
    idx = 0
    frmLOut.width = 0
    frmLOut.height = 0
    With pgSet
        For Each fld In .rs.Fields
            If IsArray(.nRow) Then
                If allIsInArray(.nRow, fld.name) Then tr = tr & trEn & trBg
            End If
            tr = tr & tb2 & "<th"
            If IsArray(.nCol) Then
                For Each col In .nCol
                    If col(0) = fld.name Then
                        tr = tr & " colspan='" & col(1) & "'"
                        Exit For
                    End If
                Next
            End If
            tr = tr & ">" & fld.name & thEnd
            itms = itms + Len(fld.name)
        Next fld
        If frmLOut.width < itms Then frmLOut.width = itms
        tr = tr & trEn
        If Not .rs.BOF Then .rs.MoveFirst
        While Not .rs.EOF
            tr = tr & trBg
            itms = 0
            For Each fld In .rs.Fields
                If IsArray(.nRow) Then
                    If allIsInArray(.nRow, fld.name) Then tr = tr & trEn & trBg
                End If
                tr = tr & tb2 & "<"
                If IsArray(.prntTag) Then
                    For Each tg In .prntTag
                        If tg(0) = fld.name Then
                            tr = tr & tg(1)
                            Exit For
                        End If
                        tr = tr & "td"
                    Next tg
                Else
                    tr = tr & "td"
                End If
                tr = tr & " id='" & Replace(fld.name, " ", "") & idx & "' onDblClick='runScript(" & idx & ", id)'"
                If IsArray(.prntName) Then
                    For Each nm In .prntName
                        If nm(0) = fld.name Then
                            tr = tr & " name='" & nm(1) & "'"
                            Exit For
                        End If
                    Next nm
                End If
                If IsArray(.nCol) Then
                    For Each col In .nCol
                        If col(0) = fld.name Then
                            tr = tr & " colspan='" & col(1) & "'"
                            Exit For
                        End If
                    Next col
                End If
                If IsArray(.prntStyle) Then
                    For Each iStl In .prntStyle
                        If iStl(0) = fld.name Then
                            tr = tr & " style='" & iStl(1) & "'"
                            Exit For
                        End If
                    Next iStl
                End If
                If IsArray(.prntClass) Then
                    For Each cl In .prntClass
                        If cl(0) = fld.name Then
                            tr = tr & " class='" & cl(1) & "'"
                            Exit For
                        End If
                    Next cl
                End If
                If IsArray(.prntTitle) Then
                    For Each tl In .prntTitle
                        If tl(0) = fld.name Then
                            tr = tr & " title='" & tl(1) & "'"
                            Exit For
                        End If
                    Next tl
                End If
                tr = tr & ">" & vbCrLf
                If IsArray(.elTag) Then
                    For Each tg In .elTag
                        If tg(0) = fld.name Then
                            tr = tr & tb3 & "<" & tg(1)
                            Exit For
                        End If
                    Next tg
                End If
                If IsArray(.elName) Then
                    For Each nm In .elName
                        If nm(0) = fld.name Then
                            tr = tr & " name='" & nm(1) & "'"
                            Exit For
                        End If
                    Next nm
                End If
                If IsArray(.elStyle) Then
                    For Each iStl In .elStyle
                        If iStl(0) = fld.name Then
                            tr = tr & " style='" & iStl(1) & "'"
                            Exit For
                        End If
                    Next iStl
                End If
                If IsArray(.elClass) Then
                    For Each cl In .elClass
                        If cl(0) = fld.name Then
                            tr = tr & " class='" & cl(1) & "'"
                            Exit For
                        End If
                    Next cl
                End If
                If IsArray(.elActn) Then
                    For Each act In .elActn
                        If act(0) = fld.name Then
                            tr = tr & act(1)
                            Exit For
                        End If
                    Next act
                End If
                If IsArray(.elTitle) Then
                    For Each tl In .elTitle
                        If tl(0) = fld.name Then
                            tr = tr & tl(1)
                            Exit For
                        End If
                    Next tl
                End If
                If IsArray(.elAlt) Then
                    For Each alt In .elAlt
                        If alt(0) = fld.name Then
                            tr = tr & alt(1)
                            Exit For
                        End If
                    Next alt
                End If
                If IsArray(.elSize) Then
                    For Each sz In .elSize
                        If sz(0) = fld.name Then
                            tr = tr & " size='" & sz(1) & "'"
                            itms = itms + sz(1) * 0.52
                            Exit For
                        End If
                    Next sz
                End If
                If IsArray(.elType) Then
                    For Each tp In .elType
                        If tp(0) = fld.name Then
                            tr = tr & " type='" & tp(1) & "'"
                            Exit For
                        End If
                    Next tp
                End If
                If IsArray(.elValue) Then
                    For Each vl In .elValue
                        If vl = fld.name Then
                            tr = tr & " value='" & Replace(fld.value, "'", "''") & "'"
                            Exit For
                        End If
                    Next vl
                End If
                If IsArray(.elTag) Then
                    For Each tg In .elTag
                        If tg(0) = fld.name Then
                            tr = tr & ">" & vbCrLf & tb1
                            Exit For
                        End If
                    Next tg
                End If
                tr = tr & tb3
                If IsArray(.elValue) Then
                    For Each vl In .elValue
                        If vl = fld.name Then Exit For
                        tr = tr & fld.value
                    Next vl
                Else
                    tr = tr & fld.value
                End If
                t = tr & vbCrLf
                If IsArray(.elTag) Then
                    For Each tg In .elTag
                        If tg(0) = fld.name Then
                            tr = tr & vbCrLf & tb3 & "</" & tg(1) & ">"
                            Exit For
                        End If
                    Next tg
                End If
                tr = tr & vbCrLf & tb2 & "</"
                If IsArray(.prntTag) Then
                    For Each tg In .prntTag
                        If tg(0) = fld.name Then
                            tr = tr & tg(1)
                            Exit For
                        End If
                        tr = tr & "td"
                    Next tg
                Else
                    tr = tr & "td"
                End If
                tr = tr & ">" & vbCrLf
                iTxt = ""
                If InStr(1, fld.value, "</a>") > 0 Then
                    iTxt = Split(fld.value, ">")(1)
                    iTxt = Split(aTxt, "<")(0)
                Else
                    iTxt = fld.value
                End If
                itms = itms + IIf(Len(iTxt) > Len(fld.name), Len(iTxt), Len(fld.name))
            Next fld
            If frmLOut.width < itms Then frmLOut.width = itms
            tr = tr & trEn
            rw = rw + 1
            idx = idx + 1
            .rs.MoveNext
        Wend
        frmLOut.height = rw
        tbl = ""
        If pgSet.winOpt > "" Then tbl = tbl & pgSet.winOpt & vbCrLf
        If pgSet.actOpt > "" Then tbl = tbl & pgSet.actOpt & vbCrLf
        If pgSet.selOpt > "" Then tbl = tbl & pgSet.selOpt & vbCrLf
        If .isIE Then
            tbl = tbl & tblBg & tr & tblEn
        Else
            tbl = tbl & jScript(1) & stl & tblBg & tr & cboAct(.rs.Fields.Count) & tblEn
        End If
    End With
    buildPage = tbl
    
exitHere:
    Exit Function
errHere:
    buildPage = Err.Description
    GoTo exitHere
End Function

Function showIEPage(tbl As String, Optional pStl As String = "")
    Dim pgDim As WebPgSet, num As Integer
    
    On Error GoTo errHere
    For i = 0 To 20
        If ie(i) Is Nothing Then
            If i > 0 Then
                If Not ie(i - 1) Is Nothing Then
                    pgDim.left = ie(i - 1).pos.left + 40
                    pgDim.top = ie(i - 1).pos.top + 40
                Else
                    pgDim.left = 0
                    pgDim.top = 0
                End If
            End If
            num = i
            pgDim.height = frmLOut.height
            pgDim.width = frmLOut.width
            Set ie(i) = New IPI_Paul_Outlook_WebDocument
            ie(i).initialize
            DoEvents
            ie(i).num = i
            ie(i).dims pgDim
            If pStl > "" Then ie(i).doc.getElementById("tblStl").outerHTML = pStl
            ie(i).doc.getElementById("tblDtl").outerHTML = tbl
            ie(i).doc.Close
            Exit For
        End If
    Next i
    
exitHere:
    Exit Function
errHere:
    Set ie(num) = Nothing
    MsgBox Err.Description
    GoTo exitHere
End Function

Function showWebPage(tbl As Variant, height As Variant, width As Variant, Optional iStl As Variant = "", Optional pScript As Variant = "", Optional Target As Variant = Nothing, _
        Optional left As Variant = 0, Optional top As Variant = 0, Optional lPos As Variant = 0, Optional tPos As Variant = 0, Optional modal As Integer = 0)
    If (height * 13.25) + 85 > 785 - top Then
        height = 785 - top
    Else
        height = (height * 13.25) + 85
    End If
    If width * 5.65 > 1024 - left Then
        width = 1024 - left
    ElseIf width * 5.65 < 100 Then
        width = 100
    Else
        width = width * 5.65
    End If
    If lPos > left Then left = lPos
    If tPos > top Then top = tPos
    IPI_Paul_Outlook_Interactive.height = height
    IPI_Paul_Outlook_Interactive.webTables.height = height - 50
    IPI_Paul_Outlook_Interactive.top = top
    IPI_Paul_Outlook_Interactive.left = left
    IPI_Paul_Outlook_Interactive.width = width + 14
    IPI_Paul_Outlook_Interactive.webTables.width = width
    IPI_Paul_Outlook_Interactive.webTables.Navigate2 "about.htm"
    Do While IPI_Paul_Outlook_Interactive.webTables.busy Or IPI_Paul_Outlook_Interactive.webTables.readyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    IPI_Paul_Outlook_Interactive.webTables.Document.Write pScript & iStl & tbl
    IPI_Paul_Outlook_Interactive.webTables.Document.Close
'    If TypeName(Target) = "Range" Then Set IPI_Paul_Outlook_Interactive.Target = Target
    IPI_Paul_Outlook_Interactive.Show 0
End Function

