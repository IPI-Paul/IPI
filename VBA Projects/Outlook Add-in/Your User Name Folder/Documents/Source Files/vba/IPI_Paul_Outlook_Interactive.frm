VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IPI_Paul_Outlook_Interactive 
   Caption         =   "IPI Paul - Results Browser"
   ClientHeight    =   1260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2970
   OleObjectBlob   =   "IPI_Paul_Outlook_Interactive.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IPI_Paul_Outlook_Interactive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Target As excel.Range
Private lyt As FORMLAYOUT, scr As SCREEN

Private Sub cboRun_Change()
    If cboRun.ListIndex > 0 Then
        If cboRun.value = "Compact Window" Then
            setDefaults
            Me.top = 0
            Me.left = lyt.left
            Me.height = IIf(lyt.height < 380, lyt.height, 380)
            Me.width = 210
            webTables.width = Me.width - 14
            webTables.height = Me.height - 50
        ElseIf cboRun.value = "Minimize Window" Then
            setDefaults
            Me.top = (scr.height / 1.366) - 52
            Me.left = (scr.width / 1.333) - (780 + 202)
            Me.height = 47
            Me.width = 200
            webTables.width = Me.width - 14
            webTables.height = IIf(Me.height - 50 < 0, Me.height, Me.height - 50)
        ElseIf cboRun.value = "Restore Window" Then
            setDefaults
            Me.top = lyt.top
            Me.left = lyt.left
            Me.height = lyt.height
            Me.width = lyt.width
            webTables.width = Me.width - 14
            webTables.height = Me.height - 50
        End If
        cboRun.ListIndex = 0
    End If
End Sub

Private Sub setDefaults()
    If lyt.height = 0 Then
        lyt.height = Me.height
        lyt.width = Me.width
        lyt.top = Me.top
        lyt.left = Me.left
    End If
End Sub

Private Sub UserForm_Initialize()
    scr.height = GetSystemMetrics32(1)
    scr.width = GetSystemMetrics32(0)
    cboRun.AddItem ""
    cboRun.AddItem "Compact Window"
    cboRun.AddItem "Minimize Window"
    cboRun.AddItem "Restore Window"
End Sub

Private Sub webTables_TitleChange(ByVal Text As String)
    Dim idx As Integer, iStr As String, itm As Variant, id As Integer, hdr As String
    
    On Error GoTo exitHere
    If webTables.readyState = READYSTATE_COMPLETE Then
        If webTables.Document.Title > "" Then
            idx = 0
            iStr = ""
            For Each itm In Split(webTables.Document.Title, ",")
                If idx = 0 Then
                    id = Int(itm)
                ElseIf idx = 1 Then
                    hdr = itm
                Else
                    If iStr > "" Then iStr = iStr & ","
                    iStr = iStr & itm
                End If
                idx = idx + 1
            Next itm
            scrpt id, hdr, iStr
        End If
    End If
    
exitHere:
    Exit Sub
End Sub

Private Function rgb(red, green, blue)
    If red = "" Then red = 0
    If green = "" Then green = 0
    If blue = "" Then blue = 0
    rgb = ((red) + ((green) * 256) + ((blue) * 65536))
End Function

Private Sub scrpt(id, hdr, iStr)
    Dim Document As New HTMLDocument, rw As HTMLTableRow, sHTML As String, wdDoc As Word.Document, tbl As HTMLTable, oDoc As New HTMLDocument, pCell As HTMLTableCell, pInput As HTMLInputElement
    Dim bgCol As String, fgCol As String, olItem As Object, uName As String, fPath As String, sPath As String, fso As New FileSystemObject, chk As ADODB.Recordset, oDta As String
    
    Set Document = webTables.Document
    'MsgBox id & vbNewLine & hdr & vbNewLine & iStr
    If Trim(hdr) = "selRun" Then
        Document.getElementById(Trim(hdr)).selectedIndex = 0
        'sHTML = Document.getElementById("tblStl").outerHTML
        If Not isInArray(Array("5"), Trim(iStr)) Then
            sHTML = sHTML & tblBg
            For Each rw In Document.getElementsByTagName("tr")
                If rw.RowIndex = 0 Or LCase(rw.Style.backgroundColor) = "rgb(253,233,217)" Then
                    If Not rw.RowIndex = 0 Then
                        bgCol = rw.Style.backgroundColor
                        rw.Style.backgroundColor = ""
                    End If
                    sHTML = sHTML & rw.outerHTML
                    If Not rw.RowIndex = 0 Then rw.Style.backgroundColor = bgCol
                End If
            Next
            sHTML = sHTML & tblEn
        End If
        If Trim(iStr) = 1 Then
            Set olItem = getOlItem
            olItem.GetInspector.WordEditor.Application.Selection = "placeHere"
            olItem.HTMLBody = Replace(olItem.HTMLBody, "placeHere", Document.getElementById("tblStl").outerHTML & sHTML)
            Set olItem = Nothing
        ElseIf Trim(iStr) = 2 Then
            bgCol = Split(Document.getElementById("tblStl").outerHTML, "th {")(1)
            bgCol = Split(bgCol, "}")(0)
            For Each itm In Split(bgCol, ";")
                If InStr(Split(itm, ":")(0), "background-color") > 0 Then
                    bgCol = Trim(Split(itm, ":")(1))
                End If
                If InStr(Split(itm, ":")(0), "color") > 0 Then
                    fgCol = Trim(Split(itm, ":")(1))
                End If
            Next
            For Each itm In Array("RGB(", ")")
                bgCol = Replace(bgCol, itm, "")
                fgCol = Replace(fgCol, itm, "")
            Next itm
            Set wdDoc = GetObject(, "Word.Application").ActiveDocument
            oDoc.Body.innerHTML = sHTML
            Set tbl = oDoc.getElementsByTagName("table")(0)
            buildWdDoc wdDoc, tbl, rgb(Trim(Split(bgCol, ",")(0)), Trim(Split(bgCol, ",")(1)), Trim(Split(bgCol, ",")(2))), _
                rgb(Trim(Split(fgCol, ",")(0)), Trim(Split(fgCol, ",")(1)), Trim(Split(fgCol, ",")(2)))
        ElseIf isInArray(Array("3", "4", "6"), Trim(iStr)) Then
            If Not Document.getElementsByName("uDefInput") Is Nothing Then
                Set pCells = Document.getElementsByTagName("th")
                Set olItem = getOlItem
                With olItem
                    i = 0
                    j = 0
                    For Each pCell In pCells
                        sHTML = Trim(pCell.innerText)
                        If pCell.Title = "uDefName" Then
                            If Trim(iStr) = 3 Then
                                .UserProperties(sHTML).value = ""
                            ElseIf Trim(iStr) = 4 Then
                                .UserProperties(sHTML).Delete
                            Else
                                If .UserProperties.Item(sHTML) Is Nothing Then .UserProperties.add sHTML, olText, True
                                If Trim(Document.getElementsByName("uDefInput")(i).value) > "" Then
                                    .UserProperties(sHTML).value = Document.getElementsByName("uDefInput")(i).value
                                    .UserProperties(sHTML).ValidationText = i
                                End If
                                i = i + 1
                            End If
                        ElseIf InStr(Split(Document.getElementsByTagName("td")(j).outerHTML, ">")(0), "name=") > 0 Then
                            If Trim(iStr) = 3 Then
                                .UserProperties(sHTML).value = ""
                            ElseIf Trim(iStr) = 4 Then
                                .UserProperties(sHTML).Delete
                            Else
                                If .UserProperties.Item(sHTML) Is Nothing Then .UserProperties.add sHTML, olText, True
                                If Trim(Document.getElementsByTagName("td")(j).innerText) > "" Then
                                    .UserProperties(sHTML).value = Trim(Document.getElementsByTagName("td")(j).innerText)
                                    .UserProperties(sHTML).ValidationText = j
                                End If
                            End If
                        End If
                        j = j + 1
                    Next
                    .Save
                End With
                Set olItem = Nothing
            End If
        ElseIf Trim(iStr) = 5 Then
            fPath = Environ$("appdata") & "\IPI Paul\Outlook\User Defined Properties\uDefMap.tab"
            If Not fso.FolderExists(Split(fPath, "Outlook", 2)(0)) Then MkDir Split(fPath, "Outlook", 2)(0)
            If Not fso.FolderExists(Split(fPath, "User Defined Properties", 2)(0)) Then MkDir Split(fPath, "User Defined Properties", 2)(0)
            If Not fso.FolderExists(Split(fPath, "uDef", 2)(0)) Then MkDir Split(fPath, "uDef", 2)(0)
            
            sPath = Environ$("appdata") & "\IPI Paul\Outlook\User Defined Properties\schema.ini"
            If Not fso.FileExists(sPath) Then
                sHTML = "[uDefMap.tab]" & vbCrLf & "Format=TabDelimited" & vbCrLf & "ColNameHeader=True"
                appendToFile sPath, sHTML
            End If
            
            If Not fso.FileExists(fPath) Then
                sHTML = "From" & vbTab & "To"
            Else
                sHTML = ""
                Set chk = runSQL(dbPath:=CStr(Split(fPath, "uDef", 2)(0)), sql:="select * from [uDefMap.tab]", tp:="Text Test")
                If Not chk.BOF Then chk.MoveFirst
            End If
            Set tbl = Document.getElementById("tblDtl")
            j = 1
            For Each rw In tbl.Rows
                If rw.cells.Length > 1 And rw.RowIndex > 0 Then
                    i = 0
                    For Each pCell In rw.cells
                        If i = 0 Then
                            If fso.FileExists(fPath) Then
                                chk.Filter = "[From]='" & Trim(pCell.innerText) & "'"
                                If Not chk.BOF Then chk.MoveFirst
                                If Not chk.EOF Then
                                    If nz(Trim(chk!To.value)) = Trim(rw.cells(1).ChildNodes(0).value) Then
                                        GoTo chkSkip
                                    Else
                                        tmp = nz(chk!To.value)
                                        If MsgBox(Trim(pCell.innerText) & " is already mapped to " & chk!To.value & ", do you want to replace it", vbYesNo, "Already Mapped") = vbYes Then
                                            chk.Close
                                            oDta = Replace(readFile(fPath), Trim(pCell.innerText) & vbTab & tmp & vbCrLf, "")
                                            updateFile fPath, oDta
                                            DoEvents
                                            Set chk = runSQL(dbPath:=CStr(Split(fPath, "uDef", 2)(0)), sql:="select * from [uDefMap.tab] where [From] > '' order by [From]", tp:="Text Test")
                                            DoEvents
                                            If Not chk.BOF Then chk.MoveFirst
                                            oDta = "From" & vbTab & "To"
                                            Do While Not chk.EOF
                                                oDta = oDta & vbCrLf & chk!From.value & vbTab & chk!To.value
                                                chk.MoveNext
                                            Loop
                                            chk.Close
                                            updateFile fPath, oDta
                                            DoEvents
                                            Set chk = runSQL(dbPath:=CStr(Split(fPath, "uDef", 2)(0)), sql:="select * from [uDefMap.tab] where [From] > '' order by [From]", tp:="Text Test")
                                            DoEvents
                                            If Not chk.BOF Then chk.MoveFirst
                                            DoEvents
                                        Else
                                            GoTo chkSkip
                                        End If
                                    End If
                                End If
                            End If
                            If sHTML > "" Then sHTML = sHTML & vbCrLf
                            sHTML = sHTML & Trim(pCell.innerText) & vbTab
                        Else
                            sHTML = sHTML & Trim(pCell.ChildNodes(0).value)
                        End If
                        i = i + 1
                    Next pCell
chkSkip:
                End If
            Next rw
            If fso.FileExists(fPath) Then chk.Close
            appendToFile fPath, sHTML
            Set chk = runSQL(dbPath:=CStr(Split(fPath, "uDef", 2)(0)), sql:="select * from [uDefMap.tab] where [From] > '' order by [From]", tp:="Text Test")
            DoEvents
            If Not chk.BOF Then chk.MoveFirst
            oDta = "From" & vbTab & "To"
            Do While Not chk.EOF
                oDta = oDta & vbCrLf & chk!From.value & vbTab & chk!To.value
                chk.MoveNext
            Loop
            chk.Close
            updateFile fPath, oDta
            DoEvents
            loadApplicationLink "Notepad", fPath
        End If
    End If
    
exitHere:
    Set Document = Nothing
End Sub
