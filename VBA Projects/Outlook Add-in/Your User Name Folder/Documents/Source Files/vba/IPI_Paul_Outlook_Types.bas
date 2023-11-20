Attribute VB_Name = "IPI_Paul_Outlook_Types"
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type FORMLAYOUT
    top As Long
    left As Long
    height As Long
    width As Long
End Type

Public Type SCREEN
    height As Long
    width As Long
End Type

Public Type WebPgSet
    actOpt As String
    elActn As Variant
    elClass As Variant
    elSize As Variant
    elStyle As Variant
    elTag As Variant
    elTitle As Variant
    elAlt As Variant
    elName As Variant
    elType As Variant
    elValue As Variant
    isIE As Boolean
    jScript As String
    nCol As Variant
    nRow As Variant
    prntClass As Variant
    prntName As Variant
    prntStyle As Variant
    prntTag As Variant
    prntTitle As Variant
    rs As ADODB.Recordset
    selOpt As String
    winOpt As String
    height As Long
    left As Long
    top As Long
    width As Long
End Type
