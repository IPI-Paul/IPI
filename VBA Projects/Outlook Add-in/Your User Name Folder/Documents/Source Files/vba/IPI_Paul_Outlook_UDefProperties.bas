Attribute VB_Name = "IPI_Paul_Outlook_UDefProperties"
Function uDefDocs() As Variant
    uDefDocs = Array("Orders", "Invoices")
End Function

Sub uDefPropAdd()
    Dim pgSet As WebPgSet, tbl As String
    
    pgSet = uDefPropSet
    Set pgSet.rs = arrayToRecordset(Array(Array("Fields", uDefs), Array("Values", Array("", "", "", ""))))
    tbl = buildPage(pgSet)
    showWebPage tbl, frmLOut.height + 2, frmLOut.width + 23
    Set pgSet.rs = Nothing
End Sub

Sub uDefPropMap(pDoc As MSHTML.HTMLDocument)
    Dim pTable As HTMLTable, pRow As HTMLTableRow, pCell As HTMLTableCell, arr As Variant, tbl As String, pgSet As WebPgSet, vals As Variant

    Set pTable = pDoc.getElementById("tblDtl")
    Set pRow = pTable.Rows(0)

    arr = Array()
    vals = arr
    i = 0
    For Each pCell In pRow.cells
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = Trim(pCell.innerText)
        ReDim Preserve vals(UBound(vals) + 1)
        vals(UBound(vals)) = ""
        i = i + 1
    Next pCell

    pgSet = uDefPropSet
    Set pgSet.rs = arrayToRecordset(Array(Array("Fields", arr), Array("Values", vals)))
    tbl = buildPage(pgSet)
    showWebPage tbl, frmLOut.height + (i / 2), frmLOut.width + 23, pScript:=localStorage
End Sub

Function uDefPropSet() As WebPgSet
    Dim pgSet As WebPgSet
    
    pgSet.prntTag = Array(Array("Fields", "th"))
    pgSet.elClass = Array(Array("Values", "func"))
    pgSet.elName = Array(Array("Values", "uDefInput"))
    pgSet.elTag = Array(Array("Values", "input"))
    pgSet.elSize = Array(Array("Values", 100))
    pgSet.elType = Array(Array("Values", "text"))
    pgSet.elValue = Array("Values")
    pgSet.prntClass = Array(Array("Fields", "func"))
    pgSet.prntName = Array(Array("Fields", "uDefName"))
    pgSet.prntStyle = Array(Array("Fields", "text-align: left;"))
    pgSet.prntTitle = Array(Array("Fields", "uDefName"))
    uDefPropSet = pgSet
End Function

Sub uDefPropShow()
    Dim pgSet As WebPgSet, tbl As String, olItem As Object, uDef As UserProperty, props As Variant, names As Variant
    
    Set olItem = getOlItem
    
    i = 1
    props = Array()
    
    For Each uDef In olItem.UserProperties
        If Not uDef.name = "VSTO" Then
            ReDim Preserve props(UBound(props) + 1)
        End If
    Next
    names = props
    
    For Each uDef In olItem.UserProperties
        If Not uDef.name = "VSTO" Then
            props(CInt(uDef.ValidationText)) = Array(uDef.name, uDef.value)
            names(CInt(uDef.ValidationText)) = Array(uDef.name, "uDefName")
        End If
    Next
    
    If olItem.UserProperties.Count > 1 Then
        Set pgSet.rs = arrayToRecordset(props)
        pgSet.prntName = names
        tbl = buildPage(pgSet)
        showWebPage tbl, frmLOut.height + 1, frmLOut.width + ((UBound(props) + 1) * 2)
        Set pgSet.rs = Nothing
    Else
        MsgBox "There are no User Defined Properties!"
    End If
End Sub

Public Function uDefs() As Variant
    uDefs = Array("Vendor Id", "Vendor Name", "Orders", "Invoices")
End Function


