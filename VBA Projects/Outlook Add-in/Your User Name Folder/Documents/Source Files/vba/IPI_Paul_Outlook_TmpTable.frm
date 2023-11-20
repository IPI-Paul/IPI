VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IPI_Paul_Outlook_TmpTable 
   Caption         =   "IPI Paul- Temporary Table"
   ClientHeight    =   345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2160
   OleObjectBlob   =   "IPI_Paul_Outlook_TmpTable.frx":0000
End
Attribute VB_Name = "IPI_Paul_Outlook_TmpTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function buildRows()
    Dim pTable As HTMLTable, pRow As HTMLTableRow, pCell As HTMLTableCell, pDoc As HTMLDocument, pTbl As String, pStl, tr As String, js As String
    Dim wlim As Long, wArr As Variant, itm As Variant
    
    js = Replace(jScript(1), "'1, ' + ", "") & vbCrLf
    pStl = Replace(stl, "8pt", "9pt")
    pTbl = "<table id=""ProvTbl"" style=""border: 1;"">" & vbCrLf
    tr = trBg & thBg & tb3 & "Inserts" & vbCrLf & thEnd & trEn
    tr = tr & trBg & tb2 & "<td class=""func"">" & vbCrLf
    tr = tr & tb3 & "<br clear=""left"" /><select id=""actSel"" onChange=""runScript(0, value)"" />" & vbCrLf
    tr = tr & tb4 & "<option value="""">Select an option</option>" & vbCrLf
    tr = tr & tb4 & "<option value=""0,1"">Clear Table</option>" & vbCrLf
    tr = tr & tb4 & "<option value=""0,2"">Add Clipboard</option>" & vbCrLf
    tr = tr & tb4 & "<option value=""0,3"">Add Table Selection</option>" & vbCrLf
    tr = tr & tb4 & "<option value=""0,4"">Add Property</option>" & vbCrLf
    tr = tr & tb4 & "<option value=""0,5"">Append Prefix</option>" & vbCrLf
    tr = tr & tb4 & "<option value=""0,6"">Append Suffix</option>" & vbCrLf
    tr = tr & tb3 & "</select>" & vbCrLf
    tr = tr & tdEn & trEn
    pTbl = pTbl & tr & "</table>"
    webTable.Navigate2 "about.htm"
    Do While webTable.busy Or webTable.readyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    webTable.Document.Write js & pStl & pTbl
    webTable.Document.Close
    DoEvents
    Set pDoc = webTable.Document
    Set pTable = pDoc.getElementsByTagName("table")(0)
    Set pRow = pTable.insertRow(0)
    wArr = Array(Array(0, 40, 6.5), Array(40, 50, 5.25), Array(50, 60, 6.5), Array(60, 150, 5.5), Array(150, 1000, 7))
    wlim = 180
    tTbl.check
    If Not tTbl.pRst.state = 0 Then
        If Not tTbl.pRst.BOF Then tTbl.pRst.MoveFirst
        While Not tTbl.pRst.EOF
            Set pRow = pTable.insertRow(pTable.Rows.Length - 1)
            Set pCell = pRow.insertCell
            pCell.innerText = tTbl.pRst!Temp.value
            For Each itm In wArr
                If Len(tTbl.pRst!Temp.value) > itm(0) And Len(tTbl.pRst!Temp.value) < itm(1) Then
                    If wlim < itm(2) * Len(tTbl.pRst!Temp.value) And tTbl.pRst!Temp.value > "" Then wlim = itm(2) * Len(tTbl.pRst!Temp.value)
                    Exit For
                End If
            Next itm
            tTbl.pRst.MoveNext
        Wend
    End If
    
exitHere:
    buildRows = Array(pTable.Rows.Length, wlim)
    Set pcvell = Nothing
    Set pRow = Nothing
    Set pTable = Nothing
    Set pDoc = Nothing
    Exit Function
End Function

Private Sub buildTmpTable()
    Dim top As Long, left As Long, height As Long, width As Long, pDim As Variant
    
    top = 0: left = 0: height = 64: width = 180
    pDim = buildRows
    DoEvents
    height = height + ((pDim(0) - 1) * 12.05)
    width = pDim(1)
    IPI_Paul_Outlook_TmpTable.height = height
    webTable.height = height - 30
    IPI_Paul_Outlook_TmpTable.top = top
    IPI_Paul_Outlook_TmpTable.left = left
    IPI_Paul_Outlook_TmpTable.width = width
    webTable.width = width - 14
    
    apiMoveWindow FindWindow(vbNullString, Me.Caption), crCoord.X, 0, width, height + 20, 1
    IPI_Paul_Outlook_TmpTable.height = height
    webTable.height = height - 30
    IPI_Paul_Outlook_TmpTable.width = width
    webTable.width = width - 14
End Sub

Private Sub UserForm_Initialize()
    buildTmpTable
End Sub

Private Sub webTable_TitleChange(ByVal Text As String)
    On Error GoTo exitHere
    If webTable.Document.getElementById("actSel").selectedIndex > 0 Then
        TitleChange Text
        webTable.Document.getElementById("actSel").selectedIndex = 0
        buildTmpTable
    End If
    
exitHere:
    Exit Sub
End Sub
