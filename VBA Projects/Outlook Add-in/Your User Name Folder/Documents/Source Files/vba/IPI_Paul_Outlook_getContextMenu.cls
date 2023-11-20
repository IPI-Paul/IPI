VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPI_Paul_Outlook_getContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents olExpl As Outlook.Explorer
Attribute olExpl.VB_VarHelpID = -1
Public WithEvents olMsg As Outlook.MailItem
Attribute olMsg.VB_VarHelpID = -1
Public WithEvents olPst As Outlook.PostItem
Attribute olPst.VB_VarHelpID = -1
Public WithEvents olInsp As Outlook.Inspectors
Attribute olInsp.VB_VarHelpID = -1

Sub initialize()
    Set olExpl = Application.ActiveExplorer
    Set olInsp = Application.Inspectors
End Sub

Private Sub olExpl_Deactivate()
    On Error GoTo exitHere
    If TypeName(ActiveInspector.CurrentItem) = "MailItem" Then
        Set olMsg = ActiveInspector.CurrentItem
    ElseIf TypeName(ActiveInspector.CurrentItem) = "PostItem" Then
        Set olPst = ActiveInspector.CurrentItem
    End If
exitHere:
    Exit Sub
End Sub

Private Sub olExpl_SelectionChange()
    GetCursorPos crCoord
    On Error GoTo exitHere
    If TypeName(olExpl.Selection(1)) = "MailItem" Then
        Set olMsg = olExpl.Selection(1)
    ElseIf TypeName(olExpl.Selection(1)) = "PostItem" Then
        Set olPst = olExpl.Selection(1)
    End If
exitHere:
    Exit Sub
End Sub

Private Sub olInsp_NewInspector(ByVal Inspector As Inspector)
    On Error GoTo exitHere
    If TypeName(Inspector.CurrentItem) = "MailItem" Then
        Set olMsg = Inspector.CurrentItem
    ElseIf TypeName(Inspector.CurrentItem) = "PostItem" Then
        Set olPst = Inspector.CurrentItem
    End If
exitHere:
    Exit Sub
End Sub

Private Sub olMsg_CustomPropertyChange(ByVal name As String)
    On Error Resume Next
    If name = "VSTO" Then
        If olMsg.UserProperties(name).value > "" Then
            GetCursorPos crCoord
            runChange Split(olMsg.UserProperties(name).value, "|")
            olMsg.UserProperties(name).Delete
        End If
    End If
End Sub

Private Sub olPst_CustomPropertyChange(ByVal name As String)
    On Error Resume Next
    If name = "VSTO" Then
        If olPst.UserProperties(name).value > "" Then
            GetCursorPos crCoord
            runChange Split(olPst.UserProperties(name).value, "|")
            olPst.UserProperties(name).Delete
        End If
    End If
End Sub

Private Sub runChange(arr As Variant)
    Dim olItem As Object, tbl(1 To 3) As HTMLTable, rw(1 To 3) As HTMLTableRow, td(1 To 3) As HTMLTableCell, th(1 To 3) As Object, inp As HTMLInputElement
    
    On Error GoTo exitHere
    If arr(0) = "GetMenuContent" Then
        GetMenuContent
    ElseIf UBound(arr) > 1 Then
        If arr(1) = "doMenu" Then
            doMenu arr(2)
        ElseIf arr(1) = "HelpMacro" Then
            HelpMacro
        ElseIf arr(1) = "ListForm" Then
            ListForm
        ElseIf arr(1) = "ListIE" Then
            ListIE
        ElseIf arr(1) = "tmpTable" Then
            IPI_Paul_Outlook_TmpTable.Show 0
        ElseIf arr(1) = "uDefPropAdd" Then
            uDefPropAdd
            Set olItem = getOlItem
            Set tbl(1) = IPI_Paul_Outlook_Interactive.webTables.Document.getElementById("tblDtl")
            Set th(1) = tbl(1).getElementsByTagName("input")
            i = 0
            With olItem
                For Each inp In th(1)
                    Set td(1) = inp.parentElement.parentElement.Children(0)
                    If isInArray(uDefs, Trim(td(1).innerText)) Then
                        inp.value = .UserProperties(Trim(td(1).innerText)).value
                    End If
                Next inp
            End With
            Set olItem = Nothing
            Set td(1) = Nothing
            Set th(1) = Nothing
            Set rw(1) = Nothing
            Set tbl(1) = Nothing
        ElseIf arr(1) = "uDefPropShow" Then
            uDefPropShow
        End If
    End If
exitHere:
    Exit Sub
End Sub
