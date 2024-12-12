Attribute VB_Name = "IPI_Paul_Outlook_Functions"
Private Const sp As String = "  ", tb1 As String = sp, tb2 As String = sp & sp, tb3 As String = tb2 & sp, tb4 As String = tb3 & sp, tb5 As String = tb4 & sp
Private Const tb6 As String = tb5 & sp, tb7 As String = tb6 & sp, trBgn As String = tb1 & "<tr>" & vbCrLf, trEnd As String = tb1 & "</tr>" & vbCrLf
Private Const thBgn As String = tb2 & "<th align='left'>" & vbCrLf, thEnd As String = tb2 & "</th>" & vbCrLf
Private Const tdBgn As String = tb2 & "<td>" & vbCrLf, tdEnd As String = tb2 & "</td>" & vbCrLf
Private Const bgColor As String = "RGB(0, 94, 184)", fgColor As String = "RGB(255, 255, 255)", selColor As String = "RGB(253, 233, 217)", unSelColor As String = "RGB(255, 255, 255)"

Function buildHTML(aRs As Object, Optional nRow As Variant = "", Optional nCol As Variant = "", Optional sCol As Variant = "", Optional hCol As Variant = "", _
    Optional getTable As Boolean = False _
)
    Dim val As Variant, rw As Integer, mwdth As Integer, mrws As Integer, itms As Integer, fld As Variant, tr As String, col As Variant, stl As Variant, i As Integer
    Dim tbl As String
    
    On Error GoTo exitHere
    
    Do While Not aRs.EOF
        rw = rw + 1
        aRs.MoveNext
        DoEvents
    Loop
    
    If Not aRs.BOF Then aRs.MoveFirst
    mwdth = 35
    mrws = 3
    itms = 0
    For Each fld In aRs.Fields
        If IsArray(nRow) Then
            If isInArray(nRow, fld.Name) Then tr = tr & trEnd & trBgn
        End If
        tr = tr & tb2 & "<th"
        If IsArray(nCol) Then
            For Each col In nCol
                If col(0) = fld.Name Then
                    tr = tr & " colspan='" & col(1) & "'"
                    Exit For
                End If
            Next
        End If
        If IsArray(hCol) Then
            For Each stl In sCol
                If stl(0) = fld.Name Then
                    tr = tr & " style='" & stl(1) & "'"
                    Exit For
                End If
            Next
        End If
        tr = tr & ">" & vbCrLf & tb3 & fld.Name & thEnd
        itms = itms + Len(fld.Name)
    Next fld
    If mwdth < itms Then mwdth = itms
    i = 0
    While Not aRs.EOF
        tr = tr & "<tr ondblclick=""runScript(" & i & ")"">" & vbCrLf
        itms = 0
        For Each fld In aRs.Fields
            If IsArray(nRow) Then
                If isInArray(nRow, fld.Name) Then tr = tr & trEnd & trBgn
            End If
            tr = tr & tb2 & "<td id=""" & fld.Name & i & """"
            If IsArray(nCol) Then
                For Each col In nCol
                    If col(0) = fld.Name Then
                        tr = tr & " colspan='" & col(1) & "'"
                        Exit For
                    End If
                Next col
            End If
            If IsArray(sCol) Then
                For Each stl In sCol
                    If stl(0) = fld.Name Then
                        tr = tr & " style='" & stl(1) & "'"
                        Exit For
                    End If
                Next
            End If
            tr = tr & ">" & vbCrLf
            If IsNumeric(fld.Value) Then
                If fld.Value < 0 Then tr = tr & "<font color=""red"">"
            End If
            tr = tr & tb3 & fld.Value
            If IsNumeric(fld.Value) Then
                If fld.Value < 0 Then tr = tr & "</fon>"
            End If
            tr = tr & tdEnd
            If InStr(1, fld.Value, "</a>") > 0 Then
                val = Split(fld.Value, ">")(1)
                val = Split(val, "<")(0)
            Else
                val = fld.Value
            End If
            If Not isInString(Array(fld.Value), "<a") Then
                itms = itms + IIf(Len(val) > Len(fld.Name), Len(val), Len(fld.Name))
            Else
                itms = itms + Len(fld.Name)
            End If
        Next fld
        mrws = mrws + Round(itms / 200)
        If mwdth < itms Then mwdth = itms
        tr = tr & trEnd
        DoEvents
        i = i + 1
        mrws = mrws + 1
        aRs.MoveNext
    Wend
    stl = getStyle
    tbl = "<table id=""MyTable"" border=""1"">" & tr & "</table>" & vbCrLf
    buildHTML = Array(tbl, stl, jScript("secNotes"))
    If Not aRs.State = 0 Then aRs.Close
exitHere:
    Set aRs = Nothing
    Exit Function
End Function

Function getDbLocation() As String
    getDbLocation = getDocsFolder & "\IPI Outlook Notes.accdb"
End Function

Function getDocsFolder()
    getDocsFolder = Replace(Environ$("temp"), "AppData\Local\Temp", "Documents")
End Function

Function getOlItem()
    Dim myItem As Object
    
    On Error Resume Next
    
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        Set myItem = ActiveInspector.CurrentItem
    Else
        Set myItem = ActiveExplorer.Selection(1)
    End If
    Set getOlItem = myItem
End Function

Function getStyle() As String
    Dim stl As String
    
    stl = "<style id=""MyStl"">" & vbCrLf
    stl = stl & "table, body, input {font-size: 8pt;}" & vbCrLf
    stl = stl & "table, tr, th, td {border-collapse: collapse}" & vbCrLf
    stl = stl & "th {background-color: " & bgColor & "; color: " & fgColor & ";}" & vbCrLf
    stl = stl & "td, th {padding: 2px 7px 2px 7px;}" & vbCrLf
    stl = stl & "input, select {outline: 0;, border-widht: 0 0 1px;}" & vbCrLf
    stl = stl & ".hidden {display: none;}" & vbCrLf
    stl = stl & ".noBorders {border: 0;}" & vbCrLf
    stl = stl & "</style>"
    getStyle = stl
End Function

Function isInArray(arr As Variant, val As Variant) As Boolean
    Dim obj As Variant
    
    For Each obj In arr
        If val = obj Then
            isInArray = True
            Exit For
        End If
    Next obj
End Function

Function isInString(arr As Variant, val As Variant) As Boolean
    Dim obj As Variant
    
    For Each obj In arr
        If IsNull(obj) Then
        ElseIf Not Len(obj) = Len(Replace(obj, val, "")) Then
            isInString = True
            Exit For
        End If
    Next obj
End Function

Function jScript(id) As String
    Dim sc As String, jq As String
    
    jq = getDocsFolder & "\Source Files\js\jquery-3.5.1.js"
    
    sc = vbCrLf & "<script type='text/javascript' language='javascript' src='file://" & jq & "'></script>" & vbCrLf
    sc = sc & "<script type='text/javascript' language='javascript' defer>" & vbCrLf
    sc = sc & tb1 & "function runScript(val) {" & vbCrLf
    If Not id = "secNotes" Then
        sc = sc & tb2 & "document.title = '" & id & ",' + val;" & vbCrLf
    Else
        sc = sc & tb2 & "tr = document.getElementsByTagName('tr')[val + 1];" & vbCrLf
        sc = sc & tb2 & "document.title = '6,' + tr.getElementsByTagName('td')[3].innerHTML + ',' + tr.getElementsByTagName('td')[0].innerHTML;" & vbCrLf
    End If
    sc = sc & tb1 & "}" & vbCrLf
    sc = sc & tb1 & "function getSelected() {" & vbCrLf
    sc = sc & tb2 & "var arr = '';" & vbCrLf
    sc = sc & tb2 & "var tds = document.getElementsByTagName('td');" & vbCrLf
    sc = sc & tb2 & "for (i = 0; i < tds.length; i++) {" & vbCrLf
    sc = sc & tb3 & "if (tds[i].style.backgroundColor == '" & selColor & "'.toLowerCase()) {" & vbCrLf
    sc = sc & tb4 & "if (arr != """") { arr = arr + ' ';}" & vbCrLf
    sc = sc & tb4 & "arr = arr + tds[i].innerText" & vbCrLf
    sc = sc & tb3 & "}" & vbCrLf
    sc = sc & tb2 & "}" & vbCrLf
    sc = sc & tb2 & "document.title = '9,' + arr;" & vbCrLf
    sc = sc & tb1 & "}"
    sc = sc & tb1 & "$(document).ready(function() {" & vbCrLf
    sc = sc & tb2 & "$('td').click(function() {" & vbCrLf
    sc = sc & tb3 & "if ($(this)[0].className != 'noNo') {" & vbCrLf
    sc = sc & tb4 & "if ((this).style.backgroundColor == '" & selColor & "'.toLowerCase()) {" & vbCrLf
    sc = sc & tb5 & "$(this).css('background-color', '" & unSelColor & "');" & vbCrLf
    sc = sc & tb4 & "} else {" & vbCrLf
    sc = sc & tb5 & "$(this).css('background-color', '" & selColor & "');" & vbCrLf
    sc = sc & tb4 & "}" & vbCrLf & vbCrLf
    sc = sc & tb4 & "getSelected();" & vbCrLf
    sc = sc & tb3 & "}" & vbCrLf
    sc = sc & tb2 & "});" & vbCrLf
    sc = sc & tb1 & "});"
    sc = sc & "</script>"
    jScript = sc
End Function
