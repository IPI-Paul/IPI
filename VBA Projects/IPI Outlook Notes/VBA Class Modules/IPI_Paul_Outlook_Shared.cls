VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPI_Paul_Outlook_Shared"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private adoRst As ADODB.Recordset

Private Sub Class_Initialize()
    Set adoRst = New ADODB.Recordset
End Sub

Public Property Let aRst(rst)
    Set adoRst = rst
End Property

Public Property Get aRst() As ADODB.Recordset
    Set aRst = adoRst
End Property
