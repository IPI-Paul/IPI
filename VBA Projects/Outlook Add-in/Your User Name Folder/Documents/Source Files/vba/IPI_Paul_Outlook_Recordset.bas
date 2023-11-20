Attribute VB_Name = "IPI_Paul_Outlook_Recordset"
Public tTbl As New IPI_Paul_Outlook_TempTable

Function arrayToRecordset(arr As Variant) As ADODB.Recordset
    Dim rws As Long, rw As Long, rs As New ADODB.Recordset, hdr As Variant
    
    On Error GoTo errHere
    
    rw = 1
    rws = 1
    If IsArray(arr(0)) Then
        If IsArray(arr(0)(1)) Then rws = UBound(arr(0)(1))
    End If
    
    With rs
        Set .ActiveConnection = Nothing
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        With .Fields
            For Each hdr In arr
                If IsArray(hdr) Then
                    .Append hdr(0), adVarChar, 400
                Else
                    .Append hdr, 200, 400
                End If
            Next hdr
        End With
        .Open
    End With
    
    With rs
        If IsArray(arr(0)) Then
            If IsArray(arr(0)(1)) Then
                For i = rw - 1 To rws
                    .AddNew rs.Fields(0).name, arr(0)(1)(i)
                    For j = 1 To UBound(arr)
                        If Not j > UBound(arr) Then
                            If Not Len(arr(j)(1)(i)) = 2 Then .update .Fields(j).name, arr(j)(1)(i)
                        End If
                    Next j
                Next i
            Else
                For i = rw To rws
                    .AddNew rs.Fields(0).name, arr(0)(i)
                    For j = 1 To UBound(arr)
                        .update .Fields(j).name, arr(j)(i)
                    Next j
                Next i
            End If
        Else
            For i = rw To rws
                .AddNew rs.Fields(0).name, ""
                For j = 1 To UBound(arr)
                    .update .Fields(j).name, ""
                Next j
            Next i
        End If
    End With
    
    Set arrayToRecordset = rs
    
exitHere:
    Exit Function
errHere:
    MsgBox Err.Description
    GoTo exitHere
End Function

Sub tmpAddClip()
    Dim itms As Variant, itm As Variant
    
    tTbl.check
    
    For Each itms In Split(getClipboard, vbCrLf)
        If itms > "" Then
            For Each itm In Split(itms, " ")
                If itm > "" Then tTbl.add itm
            Next itm
        End If
    Next itms
End Sub

Sub tmpAddProp()
    Dim olItem As Object, rng As Variant, props As Variant, msg As String, isSet As Boolean, itm As Variant, idx As Variant
    
    Set olItem = getOlItem
    With olItem
        props = Array( _
            Array(0, "Order Id", .BillingInformation, Not IsObject(.BillingInformation)), _
            Array(1, "Vendor Id", .UserPorperties.Item("Vendor Id"), Not .UserPorperties.Item("Vendor Id") Is Nothing), _
            Array(2, "Vendor Name", .UserPorperties.Item("Vendor Name"), Not .UserPorperties.Item("Vendor Name") Is Nothing), _
            Array(3, "Invoices", .UserPorperties.Item("Invoices"), Not .UserPorperties.Item("Invoices") Is Nothing) _
            )
    End With
    hlp = ""
    
    For Each itm In props
        msg = msg & vbCrLf & itm(0) & ": " & itm(1)
    Next itm
    
    idx = InputBox("Which Property?", msg, "Property Value to Retrieve", 3)
    
    If props(idx)(3) And props(idx)(2) > "" Then
        tTbl.check
        tTbl.add props(idx)(2)
    End If
    
exitHere:
    Set olItem = Nothing
End Sub

Sub tmpAddTbl()
    Dim olItem As Object, rng As Variant, itm As Variant
    
    tTbl.check
    
    Set olItem = getOlItem
    Set rng = olItem.GetInspector.WordEditor.Application.Selection
    rng.Copy
    
    For Each itm In Split(getClipboard, vbCrLf)
        If itm > "" Then tTbl.add itm
    Next itm
End Sub

Sub tmpAddTxt(Optional txt As Variant = "")
    Dim itms As Variant, itm As Variant
    
    tTbl.check
    
    For Each itms In Split(txt, vbCrLf)
        If itms > "" Then
            For Each itm In Split(itms, vbTab)
                If itm > "" Then tTbl.add itm
            Next itm
        End If
    Next itms
End Sub

Sub tmpPrefixTxt(Optional txt As Variant = "")
    If txt = "" Then
        txt = InputBox("Please enter prefix to append", "Append Prefix", "0")
    End If
    
    If txt > "" Then
        tTbl.check
        If Not tTbl.pRst.BOF Then tTbl.pRst.MoveFirst
        While Not tTbl.pRst.EOF
            tTbl.pRst.update "Temp", txt & tTbl.pRst!Temp
            tTbl.pRst.MoveNext
        Wend
    End If
End Sub

Sub tmpSuffixTxt(Optional txt As Variant = "")
    If txt = "" Then
        txt = InputBox("Please enter prefix to append", "Append Prefix", "0")
    End If
    
    If txt > "" Then
        tTbl.check
        If Not tTbl.pRst.BOF Then tTbl.pRst.MoveFirst
        While Not tTbl.pRst.EOF
            tTbl.pRst.update "Temp", tTbl.pRst!Temp & txt
            tTbl.pRst.MoveNext
        Wend
    End If
End Sub

Sub tmpClear()
    If tTbl.pRst.state <> 0 Then tTbl.pRst.Close
End Sub
