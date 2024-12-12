VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IPI_Paul_Outlook_Link 
   Caption         =   "Hyper Link Form"
   ClientHeight    =   720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15330
   OleObjectBlob   =   "IPI_Paul_Outlook_Link.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IPI_Paul_Outlook_Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Terminate()
    Dim hdr As String
    
    If txtHeader.Value = "" Then
        hdr = " " & Replace(txtLink.Value, """", "") & " "
    Else
        hdr = txtHeader.Value
    End If
    linkHTML = "<a href=""" & Replace(txtLink.Value, """", "") & """>" & hdr & "</a>"
End Sub
