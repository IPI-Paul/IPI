Attribute VB_Name = "IPI_Paul_Outlook_Menu"
Public Sub doMenu(Optional tName As Variant = "")
    Dim commBar As CommandBar, MENU_NAME As String, rs As ADODB.Recordset
    
    MENU_NAME = "commBar" & Replace(tName, " ", "_")
    
    On Error GoTo errHere
    For Each bar In ActiveExplorer.CommandBars
        If bar.name = "commBar" & Replace(tName, " ", "_") Then
            bar.Delete
        End If
    Next
    
    If tName = "Excel Test" Then
        Set rs = runSQL("Test.xlsx", "C:\Users\Paul\Documents\Source Files\xlsx\", "select * from [Compact$]", tName)
    ElseIf tName = "Ms Access Test" Then
        Set rs = runSQL("Test.accdb", "C:\Users\Paul\Documents\Source Files\accdb\", "select * from [Compact]", tName)
    ElseIf tName = "SQL Compact Edition Test" Then
        Set rs = runSQL("Test.sdf", "C:\Users\Paul\Documents\Source Files\sdf\", "select * from [Compact]", tName)
    ElseIf tName = "SQL Local Db Test" Then
        Set rs = runSQL("Test.mdf", "C:\Users\Paul\Documents\Source Files\mdf\", "select * from [Compact]", tName)
    ElseIf tName = "SQL Server Test" Then
        Set rs = runSQL("Studies", "", "select [actor_id], [first_name] + ' ' + [last_name] as [Name] from [actor]", tName)
    ElseIf tName = "Text Test" Then
        Set rs = runSQL("", "C:\Users\Paul\Documents\Source Files\tab\", "select * from [Test.tab]", tName)
    End If
    
    Set commBar = ActiveExplorer.CommandBars.add(MENU_NAME, Position:=msoBarPopup)
    Set menu = commBar.Controls.add(Type:=msoControlPopup)
    With menu
        .Caption = tName
        If Not rs Is Nothing Then
            If Not rs.BOF Then rs.MoveFirst
            While Not rs.EOF
                With .Controls.add
                    .Caption = rs.Fields(1).value
                    .OnAction = "menuSelected"
                    .Tag = Replace(.Caption, " ", "_")
                    .Parameter = tName & "|" & rs.Fields(0).value
                End With
                rs.MoveNext
                DoEvents
            Wend
            rs.Close
            Set rs = Nothing
            DoEvents
        End If
    End With
    On Error GoTo errHere
    ActiveExplorer.Activate
    commBar.ShowPopup crCoord.X, crCoord.Y

exitHere:
    Set rs = Nothing
    Set commBar = Nothing
    Exit Sub
errHere:
    MsgBox Err.Description
    Resume Next
End Sub

Sub menuSelected()
    Dim cItem As Object
    Const msg As String = "You can set User Defined Properties of emails or call other functions that utilise results passed in to the parameters" & vbNewLine & vbNewLine & "Here the id is:"
    
    With ActiveExplorer.CommandBars.ActionControl
        arr = Split(.Parameter, "|")
        If arr(0) = "Excel Test" Then
            MsgBox msg & arr(1)
        ElseIf arr(0) = "Ms Access Test" Then
            MsgBox msg & arr(1)
        ElseIf arr(0) = "SQL Compact Edition Test" Then
            MsgBox msg & arr(1)
        ElseIf arr(0) = "SQL Local Db Test" Then
            MsgBox msg & arr(1)
        ElseIf arr(0) = "SQL Server Test" Then
            MsgBox msg & arr(1)
        ElseIf arr(0) = "Text Test" Then
            MsgBox msg & arr(1)
        End If
    End With
End Sub
