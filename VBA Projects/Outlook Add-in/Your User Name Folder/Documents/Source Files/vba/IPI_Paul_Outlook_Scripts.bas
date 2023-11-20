Attribute VB_Name = "IPI_Paul_Outlook_Scripts"
Private cTd As String

Private Function getJavaFunc(iStr, func) As String
    Dim arr As Variant
    
    getJavaFunc = ""
    arr = Split(iStr, "function ")
    
    For Each itm In arr
        If left(itm, Len(func)) = func Then
            getJavaFunc = "function" & itm
            Exit For
        End If
    Next itm
    MsgBox getJavaFunc
End Function

Function getScript(fPath)
    Dim txt As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fl = fso.GetFile(fPath)
    Set ts = fl.OpenAsTextStream(ForReading, TristateUseDefault)
    txt = ""
    Do Until ts.AtEndOfStream
        If txt > "" Then txt = txt & vbNewLine
        txt = txt & ts.ReadLine
    Loop
    ts.Close
    getScript = txt

exitHere:
    Set ts = Nothing
    Set fl = Nothing
    Set fso = Nothing
End Function

Public Function jScript(id) As String
    Dim sHTML As String
    
'    sHTML = sHTML & "<script type='text/javascript' src='https://code.jquery.com/jquery-3.5.1.min.js'></script>" & vbCrLf
    sHTML = vbCrLf & "<script language=""javascript"" type=""text/javascript"">" & vbCrLf
    sHTML = sHTML & tb1 & "function runScript(hdr, iStr) {" & vbCrLf
    sHTML = sHTML & tb2 & "document.title='" & id & ", ' + hdr + ',' + iStr;" & vbCrLf
    sHTML = sHTML & tb1 & "}" & vbCrLf
    sHTML = sHTML & tb1 & "function highLt(idx) {" & vbCrLf
    sHTML = sHTML & tb2 & "if (document.getElementById(""tblDtl"").rows(idx).style.backgroundColor != ""rgb(253,233,217)"") {" & vbCrLf
    sHTML = sHTML & tb3 & "document.getElementById(""tblDtl"").rows(idx).style.backgroundColor = ""RGB(253, 233, 217)"";" & vbCrLf
    sHTML = sHTML & tb2 & "} else {" & vbCrLf
    sHTML = sHTML & tb3 & "document.getElementById(""tblDtl"").rows(idx).style.backgroundColor = ""RGB(255, 255, 255)"";" & vbCrLf
    sHTML = sHTML & tb1 & "}}" & vbCrLf
    sHTML = sHTML & "</script>" & vbCrLf
    jScript = sHTML
End Function

Private Sub scrpt(id, iStr)
    Dim rw(1 To 3) As HTMLTableRow, td(1 To 3) As HTMLTableCell, th(1 To 3) As Object, sHTML As String, tbl(1 To 3) As HTMLTable, oDoc As New HTMLDocument, inp As HTMLInputButtonElement
    Dim olItem As Object, bgCol As String, fgCol As String, css As New MSXML2.XMLHTTP, code As String, pStyle As String, fso As FileSystemObject, chk As ADODB.Recordset, uDef As UserProperty
    Dim wdDoc As Word.Document, xlWb As Excel.Workbook, col As Long, idx As Long, arr As Variant, dict As Dictionary, tArr(1 To 3) As Variant, acDb As Access.Application, cpt As String
    
    If isInArray(Array("0,0", "0,1", "0,2", "0,3", "0,4", "0,5", "0,6"), left(Trim(iStr), 3)) Then
        If Split(Trim(iStr), ",")(1) = 0 Then
            cTd = Split(Trim(iStr), ",", 3)(2)
            tmpClear
            tmpAddTxt Replace(cTd, ",", vbTab)
        ElseIf Split(Trim(iStr), ",")(1) = 1 Then
            tmpClear
        ElseIf Split(Trim(iStr), ",")(1) = 2 Then
            tmpAddClip
        ElseIf Split(Trim(iStr), ",")(1) = 3 Then
            tmpAddTbl
        ElseIf Split(Trim(iStr), ",")(1) = 4 Then
            tmpAddProp
        ElseIf Split(Trim(iStr), ",")(1) = 5 Then
            tmpPrefixTxt
        ElseIf Split(Trim(iStr), ",")(1) = 6 Then
            tmpSuffixTxt
        End If
    ElseIf Split(Trim(iStr), ",")(0) = 1 And Split(Trim(iStr), ",")(1) <= 3 Then
        If Document.getElementById("tblStl").innerHTML > "" Then
            pStyle = Document.getElementById("tblStl").outerHTML
        Else
            With css
                .Open "GET", Replace(Replace(Document.url, "JQuery%20Menu.html", ""), "#popupNested", "") & "styles/main.css", False
                .Send
                pStyle = "<style id=""tblStl"">" & .responseText & "</style>"
            End With
        End If
        arr = getColors(pStyle)
        bgCol = arr(0)
        fgCol = arr(1)
        If Split(Trim(iStr), ",")(1) <= 1 Then
            Set xlWb = CreateObject("Excel.Application").Workbooks.add()
            xlWb.Application.Visible = True
            xlWb.Worksheets.add
            xlWb.ActiveSheet.Rows(2).cells(1, 1).Select
            xlWb.Application.ActiveWindow.FreezePanes = True
            col = 1
            idx = 1
        Else
            Set xlWb = GetObject(, "Excel.Application").ActiveWorkbook
            col = xlWb.Application.ActiveCell.Column
            idx = xlWb.Application.ActiveCell.Row
            AppActivate xlWb.name
        End If
        Set tbl(1) = Document.getElementById("tblDtl")
        For Each rw(1) In tbl(1).Rows
            If rw(1).RowIndex = 0 Or isInArray(Array("0", "2"), Split(Trim(iStr), ",")(1)) Or (isInArray(Array("1", "3"), Split(Trim(iStr), ",")(1)) And rw(1).cells(0).Style.backgroundColor > "") Then
                For Each td(1) In rw(1).cells
                    If rw(1).RowIndex = 0 Then
                        xlWb.ActiveSheet.Rows(idx).cells(1, td(1).cellIndex + col).Interior.Color = CLng(bgCol)
                        xlWb.ActiveSheet.Rows(idx).cells(1, td(1).cellIndex + col).Font.Color = CLng(fgCol)
                        xlWb.ActiveSheet.Rows(idx).cells(1, td(1).cellIndex + col).Font.Bold = True
                    End If
                    xlWb.ActiveSheet.Rows(idx).cells(1, td(1).cellIndex + col).value = Trim(td(1).innerText)
                Next td(1)
                idx = idx + 1
            End If
        Next rw(1)
        xlWb.ActiveSheet.UsedRange.cells.EntireColumn.AutoFit
        If Split(Trim(iStr), ",")(1) <= 1 Then xlWb.ActiveSheet.UsedRange.cells.AutoFilter
        Set tbl(1) = Nothing
        Set xlWb = Nothing
    ElseIf Split(Trim(iStr), ",")(0) = 1 And Split(Trim(iStr), ",")(1) <= 5 Then
        Set tbl(1) = Document.getElementById("tblDtl")
        If Split(Trim(iStr), ",")(1) = 5 Then
            Set fso = New FileSystemObject
            fPath = Environ$("appdata") & "\IPI Paul\Outlook\User Defined Properties\uDefMap.tab"
            If fso.FileExists(fPath) Then
                Set chk = runSQL(dbPath:=CStr(Split(fPath, "uDef", 2)(0)), sql:="select * from [uDefMap.tab]", tp:="Text Test")
            End If
        End If
        Set dict = New Dictionary
        For Each tArr(1) In Split(Split(Trim(iStr), ",", 3)(2), ",")
            tArr(2) = Split(tArr(1), "|")
            If IsEmpty(dict(tArr(2)(1))) Then
                dict(tArr(2)(1)) = Array(CLng(tArr(2)(1)), Array(Trim(tArr(2)(2))))
            Else
                arr = dict(tArr(2)(1))(1)
                ReDim Preserve arr(UBound(arr) + 1)
                arr(UBound(arr)) = Trim(tArr(2)(2))
                dict(tArr(2)(1)) = Array(tArr(2)(1), arr)
            End If
        Next tArr(1)
        Set xlWb = GetObject(, "Excel.Application").ActiveWorkbook
        AppActivate xlWb.name
        If xlWb.ActiveSheet.FilterMode = True Then xlWb.ActiveSheet.ShowAllData
        For Each itm In dict.Items
            idx = 0
            If Split(Trim(iStr), ",")(1) = 4 Then
                idx = xlWb.Application.WorksheetFunction.match(Trim(tbl(1).Rows(0).cells(itm(0)).innerText), xlWb.ActiveSheet.AutoFilter.Range.Rows(1), 0)
            Else
                If fso.FileExists(fPath) Then
                    If Not chk.BOF Then chk.MoveFirst
                    chk.Filter = "[From]='" & Trim(tbl(1).Rows(0).cells(itm(0)).innerText) & "'"
                    If Not chk.BOF Then chk.MoveFirst
                    DoEvents
                    If Not chk.EOF And Not IsNull(chk!To.value) Then
                        idx = xlWb.Application.WorksheetFunction.match(Trim(chk!To.value), xlWb.ActiveSheet.AutoFilter.Range.Rows(1), 0)
                    End If
                End If
            End If
            If idx > 0 Then
                xlWb.ActiveSheet.Range(xlWb.ActiveSheet.AutoFilter.Range.Address).AutoFilter Field:=idx, Criteria1:=itm(1), Operator:=xlFilterValues
            End If
        Next
        Set tbl(1) = Nothing
        Set dict = Nothing
        Set xlWb = Nothing
    ElseIf Split(Trim(iStr), ",")(0) = 2 And Split(Trim(iStr), ",")(1) <= 2 Then
        fPath = "C:\Users\Paul\Documents\Source Files\accdb\Actor.accdb"
        cpt = "Actor : Database- " & fPath & " (Access 2007 - 2013 file format)"
        Set tbl(1) = Document.getElementById("tblDtl")
        Set rw(1) = tbl(1).Rows(0)
        Set acDb = GetObject(fPath)
        If Split(Trim(iStr), ",")(1) = 0 Then
            acDb.DoCmd.Close acQuery, "qryActor", False
        ElseIf Split(Trim(iStr), ",")(1) = 1 Then
            acDb.DoCmd.Close acForm, "frmActor", False
        Else
            acDb.DoCmd.Close acReport, "rptActor", False
        End If
        Set dict = New Dictionary
        For Each tArr(1) In Split(Split(Trim(iStr), ",", 3)(2), ",")
            tArr(2) = Split(tArr(1), "|")
            If IsEmpty(dict(tArr(2)(1))) Then
                dict(tArr(2)(1)) = Array(CLng(tArr(2)(1)), Array(Trim(tArr(2)(2))))
            Else
                arr = dict(tArr(2)(1))(1)
                ReDim Preserve arr(UBound(arr) + 1)
                arr(UBound(arr)) = Trim(tArr(2)(2))
                dict(tArr(2)(1)) = Array(Trim(rw(1).cells(CLng(tArr(2)(1))).innerText), arr)
            End If
        Next tArr(1)
        arr = Array()
        For Each itm In dict.Items
            ReDim Preserve arr(UBound(arr) + 1)
            arr(UBound(arr)) = Array("[" & itm(0) & "]", itm(1))
        Next
        acDb.Run "filterActor", arr
        DoEvents
        If Split(Trim(iStr), ",")(1) = 0 Then
            acDb.RunCommand acCmdAppRestore
            acDb.DoCmd.OpenQuery "qryActor", acViewNormal
            AppActivate cpt
        ElseIf Split(Trim(iStr), ",")(1) = 1 Then
            acDb.RunCommand acCmdAppMinimize
            acDb.DoCmd.OpenForm "frmActor", acFormDS
            AppActivate "Actor"
        Else
            acDb.RunCommand acCmdAppMaximize
            acDb.DoCmd.OpenReport "rptActor", acViewPreview
            AppActivate cpt
        End If
        Set rw(1) = Nothing
        Set tbl(1) = Nothing
        Set dict = Nothing
        Set acDb = Nothing
    ElseIf Split(Trim(iStr), ",")(0) = 3 And Split(Trim(iStr), ",")(1) <= 1 Then
        AppActivate ActiveWindow.Caption
        If Split(Trim(iStr), ",")(2) > "" Then
            searchOutlook fltBuild(Split(Trim(iStr), ",", 3)(2))
            DoEvents
        Else
            MsgBox "Nothing " & IIf(left(Trim(iStr), 3) = "3,0", "Highlighted", "Selected")
        End If
    ElseIf Split(Trim(iStr), ",")(0) = 3 And Split(Trim(iStr), ",")(1) <= 11 Then
        AppActivate ActiveWindow.Caption
        If Split(Trim(iStr), ",")(2) > "" Then
            If Split(Trim(iStr), ",")(1) <= 3 Then
                searchOutlook fltBuildUDef(Split(Trim(iStr), ",", 3)(2), uDefs)
            ElseIf isInArray(Array("4", "8"), Split(Trim(iStr), ",")(1)) Then
                searchOutlook fltBuildUDef(Split(Trim(iStr), ",", 3)(2), Array(uDefs(3)))
            ElseIf isInArray(Array("5", "9"), Split(Trim(iStr), ",")(1)) Then
                searchOutlook fltBuildUDef(Split(Trim(iStr), ",", 3)(2), Array(uDefs(2)))
            ElseIf isInArray(Array("6", "10"), Split(Trim(iStr), ",")(1)) Then
                searchOutlook fltBuildUDef(Split(Trim(iStr), ",", 3)(2), Array(uDefs(0)))
            ElseIf isInArray(Array("7", "11"), Split(Trim(iStr), ",")(1)) Then
                searchOutlook fltBuildUDef(Split(Trim(iStr), ",", 3)(2), Array(uDefs(1)))
            End If
            DoEvents
        Else
            MsgBox "Nothing " & IIf(isInArray(Array("3,2,", "3,8,", "3,9,", "3,10", "3,11"), left(Trim(iStr), 4)), "Highlighted", "Selected")
        End If
    ElseIf (Split(Trim(iStr), ",")(0) = 3 And Split(Trim(iStr), ",")(1) <= 15) Or (Split(Trim(iStr), ",")(0) = 4 And Split(Trim(iStr), ",")(1) <= 3) Then
        sHTML = sHTML & tblBg
        Set dict = New Dictionary
        For Each tArr(1) In Split(Split(Trim(iStr), ",", 3)(2), ",")
            tArr(2) = Split(tArr(1), "|")
            Set td(1) = Document.getElementsByTagName("tr")(CLng(tArr(2)(0))).cells(CLng(tArr(2)(1)))
            td(1).Style.backgroundColor = ""
            dict(tArr(2)(0)) = tArr(2)(1)
        Next tArr(1)
        If (Split(Trim(iStr), ",")(0) = 3 And Split(Trim(iStr), ",")(1) >= 14) Or (Split(Trim(iStr), ",")(0) = 4 And Split(Trim(iStr), ",")(1) >= 2) Then
            For Each rw(1) In Document.getElementsByTagName("tr")
                If rw(1).RowIndex = 0 Or Not IsEmpty(dict(CStr(rw(1).RowIndex))) Then
                    sHTML = sHTML & rw(1).outerHTML
                End If
            Next
        Else
            sHTML = sHTML & Document.getElementById("tblDtl").innerHTML
        End If
        For Each tArr(1) In Split(Split(Trim(iStr), ",", 3)(2), ",")
            tArr(2) = Split(tArr(1), "|")
            Set td(1) = Document.getElementsByTagName("tr")(CLng(tArr(2)(0))).cells(CLng(tArr(2)(1)))
            td(1).Style.backgroundColor = "rgb(253, 233, 217)"
        Next tArr(1)
        Set dict = Nothing
        Set td(1) = Nothing
        sHTML = sHTML & tblEn
        If Document.getElementById("tblStl").innerHTML > "" Then
            pStyle = Document.getElementById("tblStl").outerHTML
        Else
            With css
                .Open "GET", Replace(Replace(Document.url, "JQuery%20Menu.html", ""), "#popupNested", "") & "styles/main.css", False
                .Send
                pStyle = "<style id=""tblStl"">" & .responseText & "</style>"
            End With
        End If
        If isInArray(Array("3,12", "3,14"), Trim(left(iStr, 4))) Then
            Set olItem = Outlook.CreateItem(olMailItem)
            olItem.Display
            olItem.HTMLBody = olItem.HTMLBody & "<br /><br />" & pStyle & sHTML
        ElseIf isInArray(Array("3,13", "3,15"), Trim(left(iStr, 4))) Then
            Set olItem = getOlItem
            olItem.GetInspector.WordEditor.Application.Selection = "placeHere"
            olItem.HTMLBody = Replace(olItem.HTMLBody, "placeHere", pStyle & sHTML)
        End If
        If isInArray(Array("3,12", "3,13", "3,14", "3,15"), Trim(left(iStr, 4))) Then
            Set olItem = Nothing
        ElseIf isInArray(Array("4,0", "4,1", "4,2", "4,3"), Trim(left(iStr, 3))) Then
            arr = getColors(pStyle)
            bgCol = CLng(arr(0))
            fgCol = CLng(arr(1))
            If isInArray(Array("4,0", "4,2"), Trim(left(iStr, 3))) Then
                Set wdDoc = GetObject(, "Word.Application").ActiveDocument
            ElseIf isInArray(Array("4,1", "4,3"), Trim(left(iStr, 3))) Then
                Set wdDoc = CreateObject("Word.Application").Documents.add()
                wdDoc.Application.Visible = True
                With wdDoc.PageSetup
                    .Orientation = wdOrientLandscape
                    .TopMargin = CentimetersToPoints(1.75)
                    .LeftMargin = CentimetersToPoints(1.75)
                    .BottomMargin = CentimetersToPoints(1.75)
                    .RightMargin = CentimetersToPoints(1.75)
                End With
                DoEvents
            End If
            oDoc.Body.innerHTML = sHTML
            Set tbl(1) = oDoc.getElementsByTagName("table")(0)
            buildWdDoc wdDoc, tbl(1), CLng(bgCol), CLng(fgCol)
            Set tbl(1) = Nothing
            Set wdDoc = Nothing
        End If
    ElseIf isInArray(Array("3,16"), Trim(iStr)) Then
        Set olItem = getOlItem
        With olItem
            For i = 1 To .UserProperties.Count
                For Each uDef In .UserProperties
                    If Not uDef.name = "VSTO" Then
                        uDef.Delete
                    End If
                Next uDef
            Next i
            .Save
        End With
        Set olItem = noting
    ElseIf isInArray(Array("3,17"), Trim(iStr)) Then
        uDefPropMap Document
        DoEvents
        Set fso = New FileSystemObject
        fPath = Environ$("appdata") & "\IPI Paul\Outlook\User Defined Properties\uDefMap.tab"
        If fso.FileExists(fPath) Then
            Set chk = runSQL(dbPath:=CStr(Split(fPath, "uDef", 2)(0)), sql:="select * from [uDefMap.tab]", tp:="Text Test")
            If Not chk.BOF Then chk.MoveFirst
            Set tbl(1) = Document.getElementById("tblDtl")
            Set rw(1) = tbl(1).Rows(0)
            Set tbl(2) = IPI_Paul_Outlook_Interactive.webTables.Document.getElementById("tblDtl")
            Set th(2) = tbl(2).getElementsByTagName("input")
            i = 0
            For Each td(1) In rw(1).cells
                If LCase(td(1).tagName) = "th" Then
                    chk.Filter = "[From]='" & Trim(td(1).innerText) & "'"
                    If Not chk.BOF Then chk.MoveFirst
                    DoEvents
                    If Not chk.EOF And Not IsNull(chk!To.value) Then th(2)(i).value = Trim(chk!To.value)
                    DoEvents
                    i = i + 1
                End If
            Next td(1)
            Set chk = Nothing
            Set th(2) = Nothing
            Set tbl(2) = Nothing
            Set td(1) = Nothing
            Set rw(1) = Nothing
            Set tbl(1) = Nothing
        End If
        Set fso = Nothing
    ElseIf isInArray(Array("3,18"), Trim(iStr)) Then
        uDefPropShow
    ElseIf isInArray(Array("3,19", "3,20", "3,21", "3,22"), left(Trim(iStr), 4)) Then
        If Not isInArray(Array("3,19", "3,20", "3,21", "3,22"), Trim(iStr)) Then
            cTd = Join(Split(Split(Trim(iStr), ",", 3)(2), ","), " ")
            Set olItem = getOlItem
            With olItem
                For Each itm In uDefs
                    If .UserProperties.Item(itm) Is Nothing Then
                        .UserProperties.add itm, olText, True
                        .UserProperties(itm).ValidationText = match(uDefs, CStr(itm))
                    End If
                Next itm
                If isInArray(Array("3,19", "3,20"), left(Trim(iStr), 4)) Then
                    .UserProperties.Item("Invoices").value = Trim(cTd)
                Else
                    .UserProperties.Item("Orders").value = Trim(cTd)
                End If
                .Save
            End With
            Set olItem = Nothing
        Else
            ie(Document.cookie).navigate "javascript: getInner('" & id & "," & Trim(iStr) & "'); console.log('Done');"
        End If
    ElseIf isInArray(Array("3,23", "3,24"), left(Trim(iStr), 4)) Then
        If Not isInArray(Array("3,23", "3,24"), Trim(iStr)) Then
            uDefPropAdd
            DoEvents
            Set fso = New FileSystemObject
            fPath = Environ$("appdata") & "\IPI Paul\Outlook\User Defined Properties\uDefMap.tab"
            If Not fso.FileExists(fPath) Then
                MsgBox "There is no mapping file! Please create one using the Map Page Headers option and try again!", vbInformation, "Mapping File Missing"
            Else
                Set chk = runSQL(dbPath:=CStr(Split(fPath, "uDef", 2)(0)), sql:="select * from [uDefMap.tab]", tp:="Text Test")
                If Not chk.BOF Then chk.MoveFirst
                Set tbl(1) = Document.getElementById("tblDtl")
                Set rw(1) = tbl(1).Rows(CInt(Split(Trim(iStr), ",", 3)(2)))
                Set th(1) = tbl(1).getElementsByTagName("th")
                Set tbl(2) = IPI_Paul_Outlook_Interactive.webTables.Document.getElementById("tblDtl")
                Set th(2) = tbl(2).getElementsByTagName("input")
                i = 0
                For Each td(1) In rw(1).cells
                    If LCase(td(1).tagName) = "td" Then
                        chk.Filter = "[From]='" & Trim(th(1)(i).innerText) & "'"
                        If Not chk.BOF Then chk.MoveFirst
                        DoEvents
                        If Not chk.EOF And Not IsNull(chk!To.value) Then
                            If left(Trim(iStr), 4) = "3,23" Or (left(Trim(iStr), 4) = "3,24" And Not isInArray(uDefDocs, Trim(chk!To.value))) Then
                                For Each inp In th(2)
                                    Set td(2) = inp.parentElement.parentElement.Children(0)
                                    If Trim(td(2).innerText) = Trim(chk!To.value) Then
                                        inp.value = Trim(td(1).innerText)
                                        Exit For
                                    End If
                                Next inp
                            End If
                        End If
                        DoEvents
                        i = i + 1
                    End If
                Next td(1)
                For Each td(1) In tbl(1).getElementsByTagName("td")
                    If LCase(td(1).tagName) = "td" Then
                        If LCase(td(1).Style.backgroundColor) = "rgb(253, 233, 217)" Then
                            Set rw(1) = td(1).parentElement
                            chk.Filter = "[From]='" & Trim(th(1)(indexOf(rw(1), td(1))).innerText) & "'"
                            If Not chk.BOF Then chk.MoveFirst
                            DoEvents
                            If Not chk.EOF And Not IsNull(chk!To.value) Then
                                If isInArray(uDefDocs, Trim(chk!To.value)) Then
                                    For Each inp In th(2)
                                        Set td(2) = inp.parentElement.parentElement.Children(0)
                                        If Trim(td(2).innerText) = Trim(chk!To.value) Then
                                            If inp.value > "" Then inp.value = inp.value & " "
                                            inp.value = inp.value & Trim(td(1).innerText)
                                            Exit For
                                        End If
                                    Next inp
                                End If
                            End If
                            DoEvents
                        End If
                    End If
                Next td(1)
                Set chk = Nothing
                Set th(2) = Nothing
                Set td(2) = Nothing
                Set tbl(2) = Nothing
                Set th(1) = Nothing
                Set rw(1) = Nothing
                Set tbl(1) = Nothing
            End If
            Set fso = Nothing
        Else
            ie(Document.cookie).navigate "javascript: getRow('" & id & "," & Trim(iStr) & "'); console.log('Done');"
        End If
    ElseIf isInArray(Array("3,25", "3,26"), left(Trim(iStr), 4)) Then
        Set fso = New FileSystemObject
        fPath = Environ$("appdata") & "\IPI Paul\Outlook\User Defined Properties\uDefMap.tab"
        If Not fso.FileExists(fPath) Then
            MsgBox "There is no mapping file! Please create one using the Map Page Headers option and try again!", vbInformation, "Mapping File Missing"
        Else
            If Not isInArray(Array("3,25", "3,26"), Trim(iStr)) Then
                Set chk = runSQL(dbPath:=CStr(Split(fPath, "uDef", 2)(0)), sql:="select * from [uDefMap.tab]", tp:="Text Test")
                If Not chk.BOF Then chk.MoveFirst
                Set tbl(1) = Document.getElementById("tblDtl")
                Set th(1) = tbl(1).getElementsByTagName("th")
                Set rw(1) = tbl(1).Rows(CInt(Split(Trim(iStr), ",", 3)(2)))
                Set olItem = getOlItem
                With olItem
                    For Each itm In uDefs
                        If .UserProperties.Item(itm) Is Nothing Then
                            .UserProperties.add itm, olText, True
                            .UserProperties(itm).ValidationText = match(uDefs, CStr(itm))
                        End If
                    Next itm
                    i = 0
                    For Each td(1) In rw(1).cells
                        If LCase(td(1).tagName) = "td" Then
                            chk.Filter = "[From]='" & Trim(th(1)(i).innerText) & "'"
                            If Not chk.BOF Then chk.MoveFirst
                            If Not chk.EOF And Not IsNull(chk!To.value) Then
                                If left(Trim(iStr), 4) = "3,25" Or (left(Trim(iStr), 4) = "3,26" And Not isInArray(uDefDocs, Trim(chk!To.value))) Then
                                    .UserProperties.Item(Trim(chk!To.value)).value = Trim(td(1).innerText)
                                Else
                                    .UserProperties.Item(Trim(chk!To.value)).value = ""
                                End If
                            End If
                            i = i + 1
                        End If
                    Next td(1)
                    For Each td(1) In tbl(1).getElementsByTagName("td")
                        If LCase(td(1).tagName) = "td" Then
                            If LCase(td(1).Style.backgroundColor) = "rgb(253, 233, 217)" Then
                                Set rw(1) = td(1).parentElement
                                chk.Filter = "[From]='" & Trim(th(1)(indexOf(rw(1), td(1))).innerText) & "'"
                                If Not chk.BOF Then chk.MoveFirst
                                DoEvents
                                If Not chk.EOF And Not IsNull(chk!To.value) Then
                                    If isInArray(uDefDocs, Trim(chk!To.value)) Then
                                        If .UserProperties.Item(Trim(chk!To.value)).value > "" Then
                                            .UserProperties.Item(Trim(chk!To.value)).value = .UserProperties.Item(Trim(chk!To.value)).value & " "
                                        End If
                                        .UserProperties.Item(Trim(chk!To.value)).value = .UserProperties.Item(Trim(chk!To.value)).value & Trim(td(1).innerText)
                                    End If
                                End If
                                DoEvents
                            End If
                        End If
                    Next td(1)
                    .Save
                End With
                Set chk = Nothing
                Set th(1) = Nothing
                Set rw(1) = Nothing
                Set tbl(1) = Nothing
                 Set olItem = Nothing
           Else
                ie(Document.cookie).navigate "javascript: getRow('" & id & "," & Trim(iStr) & "'); console.log('Done');"
            End If
        End If
        Set fso = Nothing
    End If
    
exitHere:
    Set css = Nothing
    Set Document = Nothing
End Sub

Public Sub TitleChange(ByVal Text As String)
    Dim idx As Integer, iStr As String, itm As Variant, id As Integer, hdr As String
    
    On Error GoTo exitHere
    id = Split(Text, ",", 2)(0)
    iStr = ""
    If UBound(Split(Text, ",")) > 0 Then iStr = Split(Text, ",", 2)(1)
    scrpt id, iStr
    DoEvents

exitHere:
    Exit Sub
End Sub


Sub tt()
    MsgBox match(uDefs, "Invoices")
End Sub

