Attribute VB_Name = "IPI_Paul_Outlook_Voice"
Sub readClipboard()
    Dim objData As New MSForms.DataObject
    
    objData.GetFromClipboard
    IPI_Paul_Outlook_Speech.Show 0
    IPI_Paul_Outlook_Speech.msg = objData.GetText()
    objData.Clear
End Sub

Sub readEntireEmail()
    Dim olItem As Object
    
    On Error Resume Next
    Set olItem = getOlItem
    
    IPI_Paul_Outlook_Speech.Show 0
    IPI_Paul_Outlook_Speech.msg = removeLinks(getTextContent(olItem), getHyperlinks(olItem))
    DoEvents
    Set olItem = Nothing
End Sub
    
Sub readOutFile()
    Dim fPath As String
    
    fPath = getFilePath
    If fPath > "" Then 'readOutStream fPath
        IPI_Paul_Outlook_Speech.Show 0
        IPI_Paul_Outlook_Speech.mFile = fPath
    End If
End Sub

Sub readOutLoud(txt)
    Dim oVoice As New SpVoice, oVoiceFile As SpFileStream, sFile As String
    
    oVoice.Speak txt
    DoEvents
    
    Set oVoice = Nothing
End Sub

Sub readOutStream(sFile As String)
    Dim oVoice As New SpVoice, oVoiceFile As New SpFileStream
    
    oVoiceFile.Open sFile
    oVoice.SpeakStream oVoiceFile
End Sub

Sub readSelectedText()
    Dim olItem As Object
    
    On Error Resume Next
    Set olItem = getOlItem
    
    IPI_Paul_Outlook_Speech.Show 0
    IPI_Paul_Outlook_Speech.msg = olItem.GetInspector.WordEditor.Application.Selection.Text
    Set olItem = Nothing
End Sub
