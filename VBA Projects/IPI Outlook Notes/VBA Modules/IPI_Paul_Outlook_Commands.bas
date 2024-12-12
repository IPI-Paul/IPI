Attribute VB_Name = "IPI_Paul_Outlook_Commands"
Sub findAndOpen(strDASLFilter As String)
    Dim olNS As Outlook.NameSpace, olFld As Outlook.Folder, strScope As String, objSearch As Outlook.Search, acc As Variant, arr As Variant
    
    Set olNS = Application.GetNamespace("MAPI")
    Set olFld = olNS.GetDefaultFolder(olFolderInbox)

    On Error Resume Next

    arr = Array()
    For Each acc In olNS.Accounts
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = "'\\" & CStr(acc) & "\Inbox', '\\" & CStr(acc) & "\Sent Items', '\\" & CStr(acc) & "\Tasks', '\\" & CStr(acc) & "\Drafts', '\\" & CStr(acc) & "\Deleted Items'"
    Next
    
    strScope = "'Inbox', 'Sent Items', 'Tasks', 'Drafts', 'Deleted Items'"
    Set objSearch = Application.AdvancedSearch(Scope:=strScope, Filter:=strDASLFilter, SearchSubFolders:=True, Tag:="SearchFolder")
    DoEvents
    Sleep 1000
    If Not objSearch.Results.GetFirst Is Nothing Then
        objSearch.Results.GetFirst.Display
        DoEvents
        GoTo exitHere
    End If
    
'    strScope = CStr(arr(0))
'    Set objSearch = Application.AdvancedSearch(Scope:=strScope, Filter:=strDASLFilter, SearchSubFolders:=True, Tag:="SearchFolder")
'    Sleep 1000
'    If Not objSearch.Results.GetFirst Is Nothing Then
'        objSearch.Results.GetFirst.Display
'        GoTo exitHere
'    End If

    For i = 0 To 0
        strScope = CStr(arr(i))
        Set objSearch = Application.AdvancedSearch(Scope:=strScope, Filter:=strDASLFilter, SearchSubFolders:=True, Tag:="SearchFolder")
        DoEvents
        Sleep 1000
        If Not objSearch.Results.GetFirst Is Nothing Then
            objSearch.Results.GetFirst.Display
            DoEvents
            GoTo exitHere
        End If
    Next i
    
exitHere:
    Set olNS = Nothing
    Set olFld = Nothing
    Set objSearch = Nothing
End Sub

Sub getHTML(html As String, obj As InkEdit)
    Dim myItem As Outlook.MailItem
    
    Set myItem = Application.CreateItem(olMailItem)
    With myItem
        .HTMLBody = html
        With .GetInspector.WordEditor.windows(1).Selection
            .Find.Execute ""
            .collapse wdCollapseStart
            .MoveEnd WdUnits.wdStory, 1
            .Copy
        End With
    End With
    PasteToControl obj.hWnd
    myItem.Close olDiscard
End Sub

Sub addEmailNotes()
    IPI_Paul_Outlook_Notes.Show 0
End Sub
