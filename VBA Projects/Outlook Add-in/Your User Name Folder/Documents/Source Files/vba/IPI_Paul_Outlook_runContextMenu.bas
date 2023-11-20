Attribute VB_Name = "IPI_Paul_Outlook_runContextMenu"
Private nOlEvt As New IPI_Paul_Outlook_getContextMenu

Sub endEvent()
    Set nOlEvt = Nothing
End Sub

Sub startEvent()
    With nOlEvt
        .initialize
    End With
End Sub

Public Sub GetMenuContent()
    Dim xml As String, btn As Integer, frnt As String, bck As String, img As String, lbl As String, act As String, aTo As String
    Dim addIn As COMAddIn, automationObject As Object
    Set addIn = Application.COMAddIns("IPI Paul - Outlook AddIn")
    Set automationObject = addIn.Object

    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    
    frnt = "<button id="""
    img = """ imageMso="""
    lbl = """ label="""
    act = """ onAction=""runMacro"" tag="""
    bck = """ />" & vbCrLf
    
    btn = 1
    xml = xml & frnt & "btn" & btn & img & "Help" & lbl & "Help" & act & "VSTO|HelpMacro|" & bck
    xml = xml & "<menu id=""testMenu"" label=""SQL Tests"">" & vbCrLf
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "Consolidate" & lbl & "Excel Test" & act & "VSTO|doMenu|Excel Test" & bck
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "MicrosoftOnTheWeb01" & lbl & "List SQL Results with Internet Explorer" & act & "VSTO|ListIE|" & bck
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "FindDialog" & lbl & "List SQL Results in User Form" & act & "VSTO|ListForm|" & bck
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "InviteAttendees" & lbl & "Ms Access Test" & act & "VSTO|doMenu|Ms Access Test" & bck
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "DiagramRadialInsertClassic" & lbl & "SQL Compact Edition Test" & act & "VSTO|doMenu|SQL Compact Edition Test" & bck
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "DistributionListSelectMembers" & lbl & "SQL Local Db Test" & act & "VSTO|doMenu|SQL Local Db Test" & bck
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "DataGraphicIconSet" & lbl & "SQL Server Test" & act & "VSTO|doMenu|SQL Server Test" & bck
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "PositionAbsoluteMarks" & lbl & "Text Test" & act & "VSTO|doMenu|Text Test" & bck
    xml = xml & "</menu>" & vbCrLf
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "AccessListTasks" & lbl & "Temp Table" & act & "VSTO|tmpTable|" & bck
    xml = xml & "<menu id=""uDefMenu"" label=""User Defined Properties"">" & vbCrLf
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "MessageProperties" & lbl & "Show User Defined Properties" & act & "VSTO|uDefPropShow|" & bck
    btn = btn + 1
    xml = xml & frnt & "btn" & btn & img & "PropertyInsert" & lbl & "Update User Defined Properties" & act & "VSTO|uDefPropAdd|" & bck
    xml = xml & "</menu>" & vbCrLf
    xml = xml & "</menu>"
    
    automationObject.ImportData xml
End Sub

Public Sub HelpMacro()
    MsgBox "Help!"
End Sub

Public Sub ListForm()
    Dim rs As ADODB.Recordset, sHTML As String, pgSet As WebPgSet
    
    Set pgSet.rs = runSQL("Studies", "", "select * from [actor]", "SQL Server Test")
    sHTML = buildPage(pgSet)
    showWebPage sHTML, frmLOut.height, frmLOut.width, stl, jScript(0)
End Sub

Public Sub ListIE()
    Dim rs As ADODB.Recordset, sHTML As String, pgSet As WebPgSet
    
    Set pgSet.rs = runSQL("Studies", "", "select * from [actor]", "SQL Server Test")
    pgSet.isIE = True
    sHTML = buildPage(pgSet)
    showIEPage sHTML
    Set pgSet.rs = Nothing
End Sub

