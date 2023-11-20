VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IPI_Paul_Outlook_Speech 
   Caption         =   "Read Out"
   ClientHeight    =   480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "IPI_Paul_Outlook_Speech.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IPI_Paul_Outlook_Speech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Voice As SpeechLib.SpVoice
Attribute Voice.VB_VarHelpID = -1
Private VoiceFile As SpFileStream
Public msg As String, mFile As String

Private Sub cmdRun_Click()
    On Error GoTo errHere
    
    Select Case cmdRun.Caption
    
        Case "Start"
            Call SpeakAgain
            cmdRun.Caption = "Pause"
        
        Case "Pause"
            Voice.Pause
            cmdRun.Caption = "Resume"
        
        Case "Resume"
            Voice.Resume
            cmdRun.Caption = "Pause"
            
    End Select
    
errHere:
    If Err.Number Then ShowErrMsg
End Sub

Private Sub Form_Load()
    On Error GoTo errHere
    
    Set Voice = New SpVoice
    Set VoiceFile = New SpFileStream
    Voice.AlertBoundary = SVEPhoneme
    cmdRun.Caption = "Start"
    
errHere:
    If Err.Number Then ShowErrMsg
End Sub

Private Sub UserForm_Initialize()
    Form_Load
End Sub

Private Sub Voice_EndStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
    ' Call SpeakAgain
    End
End Sub

Private Sub SpeakAgain()
    On Error GoTo errHere

    If msg > "" Then
        Voice.Speak msg, SVSFlagsAsync
    ElseIf mFile > "" Then
        VoiceFile.Open mFile
        Voice.SpeakStream VoiceFile, SVSFlagsAsync
    End If
    
    
errHere:
    If Err.Number Then ShowErrMsg
End Sub
