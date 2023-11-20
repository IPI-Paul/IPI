VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPI_Paul_Outlook_TempTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rs As ADODB.Recordset

Public Sub init()
    Set rs = New ADODB.Recordset
    
    With rs
        Set .ActiveConnection = Nothing
        .CursorLocation = adUseClient 'adUseClient=3 adUseServer=2
        .LockType = adLockBatchOptimistic 'adLockBatchOptimistic=4 adLockOptimistic=3 adLockPessimistic=2 adLockReadOnly=1
        With .Fields
            .Append "Temp", adVarChar, 400 'adVarChar=200 adArray=8192 adBigInt=20 adBinary=128 adBoolean=11 adBSTR=8 adChapter=136 adChar=129 adCurrency=6 adDate=7 adDBDate=133 adDBTime=134
                                           'adDBTimeStamp=135 adDecimal=14 adDouble=5 adEmpty=0 adError=10 adFileTime=64 adGUID=72 adIDispatch=9 adInteger=3 adIUnknown=13 adLongVarBinary=205
                                           'adLongVarChar=201 adLongVarWChar=203 adNumeric=131 adPropVariant=138 adSingle=4 adSmallInt=2 adTinyInt=16 adUnsignedBigInt=21 adUnsignedInt=19
                                           'adUnsignedSmallInt=18 adUnsignedTinyInt=17 adUserDefined=132 adVarBinary=204 adVariant=12 adVarNumeric=139 adVarWChar=202 adWChar=130
        End With
    End With
End Sub

Public Property Get pRst() As ADODB.Recordset
    Set pRst = rs
End Property

Public Sub add(rec As Variant)
    rs.AddNew "Temp", rec
End Sub

Public Sub check()
    If rs Is Nothing Then
        init
        rOpen
    ElseIf pRst.state = 0 Then
        init
        rOpen
    End If
End Sub

Public Sub rClose()
    rs.Close
End Sub

Public Sub rOpen()
    rs.Open
End Sub
