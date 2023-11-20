Attribute VB_Name = "IPI_Paul_Outlook_CommandBars"

Sub addCommBars()
    Dim commBar As CommandBar, nCommBar As CommandBarPopup

    Outlook.ActiveExplorer.Activate
    deleteCommBars
    For Each commBar In Outlook.ActiveExplorer.Selection(1).GetInspector.CommandBars
        If commBar.name > "" Then
            On Error GoTo skipHere
            Set nCommBar = commBar.Controls.add(Type:=msoControlPopup, Before:=1)
            With nCommBar
                .Tag = "Delete_Me_After"
                .Caption = "NewCommBarIn" & commBar.name
            End With
skipHere:
            If Err.Number <> 0 Then Err.Clear
            Resume Next
        End If
    Next
End Sub

Sub deleteCommBars()
    Dim commBar As CommandBar, nCommBar As CommandBarPopup

    Outlook.ActiveExplorer.Activate
    For Each commBar In Outlook.ActiveExplorer.Selection(1).GetInspector.CommandBars
        If commBar.name > "" Then
            For Each ctl In commBar.Controls
                With ctl
                    If .Tag = "Delete_Me_After" Then
                        .Delete
                    End If
                End With
            Next
        End If
    Next
End Sub


