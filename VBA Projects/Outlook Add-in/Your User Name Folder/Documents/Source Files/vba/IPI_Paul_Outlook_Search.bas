Attribute VB_Name = "IPI_Paul_Outlook_Search"
Public Const olBILL As String = "urn:schemas:contacts:billininformation", olSUBJ As String = "urn:schemas:httpmail:subject", olBOD As String = "urn:schemas:httpmail:textdescription"
Public Const olTO As String = "urn:schemas:httpmail:to", olFROM As String = "urn:schemas:httpmail:from", olDISP As String = "urn:schemas:httpmail:displayto"
Public Const advSrchUDef As String = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/@Prop"

Function fltBuild(iStr As Variant)
    Dim flt As String, arr As Variant, srch As String
    
    arr = Split(iStr, ",")
    For Each itm In Array(olBILL, olSUBJ, olBOD, olTO, olFROM, olDISP)
        srch = ""
        For i = 0 To UBound(arr)
            If srch > "" Then srch = srch & " or "
            srch = srch & "(""" & itm & """ like '%" & Trim(arr(i)) & "%')"
        Next
        If flt > "" Then flt = flt & " or "
        flt = flt & srch
    Next itm
    
    fltBuild = flt
End Function

Function fltBuildUDef(iStr As Variant, uArr As Variant)
    Dim flt As String, arr As Variant, srch As String
    
    arr = Split(iStr, ",")
    For Each itm In uArr
        srch = ""
        For i = 0 To UBound(arr)
            If srch > "" Then srch = srch & " or "
            srch = srch & "(""" & Replace(advSrchUDef, "@Prop", Replace(itm, " ", "%20")) & """ like '%" & Trim(arr(i)) & "%')"
        Next
        If flt > "" Then flt = flt & " or "
        flt = flt & srch
    Next itm
    
    fltBuildUDef = flt
End Function

Private Sub searchDelete(sName As String)
    Dim objStores As Outlook.Stores, objStore As Outlook.Store, objSearchFolders As Outlook.Folders, objSearchFolder As Outlook.Folder
    
    On Error Resume Next
    Set objStores = Session.Stores
    
    For Each objStore In objStores
        Set objSearchFolders = objStore.GetSearchFolders
        For Each objSearchFolder In objSearchFolders
            If objSearchFolder.name = sName Then
                objSearchFolder.Delete
            End If
        Next
    Next

exitHere:
    Set objSearchFolders = Nothing
    Set objStores = Nothing
End Sub

Sub searchOutlook(strDASLFilter As String)
    Dim strScope As String, objSearch As Search
    Const schFolder As String = "IPI Paul Search"

    On Error GoTo errHere
    searchDelete schFolder
    DoEvents

    strScope = "'Inbox', 'Sent Items', 'Tasks', 'Drafts'"
    
    Set objSearch = AdvancedSearch(Scope:=strScope, Filter:=strDASLFilter, SearchSubFolders:=True, Tag:="SearchFolder")
    objSearch.Results.GetFirst
    objSearch.Save schFolder
    DoEvents


    searchSelect "\\paul@ipi-international.co.uk\search folders\" & schFolder
    DoEvents
    AppActivate schFolder & " - paul@ipi-international.co.uk - Outlook"
    
exitHere:
    Set objSearch = Nothing
    Exit Sub
errHere:
    MsgBox Err.Description
    GoTo exitHere
End Sub

Private Sub searchSelect(sPath As String)
    Dim objStores As Outlook.Stores, objStore As Outlook.Store, objSearchFolders As Outlook.Folders, objSearchFolder As Outlook.Folder
    
    On Error Resume Next
    Set objStores = Session.Stores
    
    For Each objStore In objStores
        Set objSearchFolders = objStore.GetSearchFolders
        For Each objSearchFolder In objSearchFolders
            If objSearchFolder.FolderPath = sPath Then
                ActiveExplorer.SelectFolder objSearchFolder
                Exit For
                Exit For
            End If
        Next
    Next

exitHere:
    Set objSearchFolders = Nothing
    Set objStores = Nothing
End Sub
