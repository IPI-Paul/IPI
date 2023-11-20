Attribute VB_Name = "IPI_Paul_Outlook_Functions"
Function allIsInArray(arr, iStr) As Boolean
    For Each aStr In arr
        If iStr = aStr Then
            allIsInArray = True
            Exit Function
        End If
    Next
End Function

Function allIsInFields(flds, arr) As Boolean
    For Each fld In flds
        For Each aStr In arr
            If aStr = fld.name Then
                allIsInFields = True
                Exit Function
            End If
        Next
    Next
End Function

Function arrayIsInString(iStr, arr)
    For Each aStr In arr
        If InStr(iStr, aStr) Then
            arrayIsInString = True
            Exit Function
        End If
    Next
End Function

Function CentimetersToPoints(centimeter)
    CentimetersToPoints = centimeter * 28.3464567
End Function

Function getClipboard()
    Dim objData As New MSForms.DataObject
    
    getClipboard = ""
    On Error Resume Next
    objData.GetFromClipboard
    getClipboard = objData.GetText()
End Function

Function getColors(pStyle As Variant)
    bgCol = Split(pStyle, "th {")(1)
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
    getColors = Array(rgb(Trim(Split(bgCol, ",")(0)), Trim(Split(bgCol, ",")(1)), Trim(Split(bgCol, ",")(2))), _
                rgb(Trim(Split(fgCol, ",")(0)), Trim(Split(fgCol, ",")(1)), Trim(Split(fgCol, ",")(2))))
End Function

Function getFilePath() As String
    Dim fDialog As Office.FileDialog, varFile As Variant, wrd As New Word.Application
    
    On Error GoTo errHere
    
    Set fDialog = wrd.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Get File List"
        .InitialFileName = Environ$("userprofile") & "\Documents\"
        If .Show = True Then
            For Each varFile In .SelectedItems
                If Len(varFile) = 0 Then GoTo nothin
                getFilePath = varFile
            Next varFile
        Else
nothin:
            MsgBox "Error Opening File"
        End If
    End With
    
exitHere:
    wrd.Quit
    Set wrd = Nothing
    Set fDialog = Nothing
    Exit Function
errHere:
    If Err.Number Then ShowErrMsg
    GoTo exitHere
End Function

Function getOlItem() As Object
    On Error Resume Next
    
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        Set getOlItem = ActiveInspector.CurrentItem
    Else
        Set getOlItem = ActiveExplorer.Selection(1)
    End If
End Function

Function getOutlookSubject() As String
    Dim olItem As Object
    
    On Error Resume Next
    Set olItem = getOlItem
    getOutlookSubject = olItem.Subject
    Set olItem = Nothing
End Function

Function indexOf(rw As HTMLTableRow, td As HTMLTableCell) As Integer
    Dim cell As HTMLTableCell
    
    i = 0
    For Each cell In rw.cells
        If LCase(cell.tagName) = LCase(td.tagName) Then
            If cell.outerHTML = td.outerHTML Then Exit For
            i = i + 1
        End If
    Next
    indexOf = i
End Function

Function isInArray(arr, iStr) As Boolean
    For Each aStr In arr
        If iStr = aStr Then
            isInArray = True
            Exit Function
        End If
    Next
    For Each itm In Split(iStr, " ")
        For Each aStr In arr
            If itm = aStr Then
                isInArray = True
                Exit Function
            End If
        Next
    Next
End Function
Function match(arr As Variant, iStr As String) As Integer
    i = 0
    For Each itm In arr
        If itm = iStr Then Exit For
        i = i + 1
    Next
    match = i
End Function

Function max(ParamArray arr() As Variant)
    Dim mn As Long
    
    For Each itm In arr
        If min < itm Then min = itm
    Next
End Function

Function min(ParamArray arr() As Variant)
    Dim mn As Long
    
    For Each itm In arr
        If min = 0 Or min > itm Then min = itm
    Next
End Function

Function nz(obj As Variant, Optional tp As Variant = "")
    On Error Resume Next
    
    nz = tp
    If Not IsNull(obj) Then
        nz = obj
    End If
End Function

Function waitTill(Optional dur As String = "00:00:01") As Boolean
    Dim tNow As Date
    
    tNow = Now() + TimeValue(dur)
    While Now() < tNow
        DoEvents
    Wend
End Function
