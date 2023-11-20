VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPI_Paul_Outlook_WebDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents iExpl As SHDocVw.InternetExplorer
Attribute iExpl.VB_VarHelpID = -1
Private idx As Integer
Public doc As HTMLDocument
Attribute doc.VB_VarHelpID = -1
Const url As String = "http://localhost:63140/Accordian%20Menu/JQuery%20Menu.html"

Sub initialize()
    Set iExpl = New SHDocVw.InternetExplorerMedium
    'Set ie = New SHDocVw.InternetExplorer
    iExpl.Visible = True
    iExpl.Navigate2 url
    
    While iExpl.busy Or iExpl.readyState <> READYSTATE_COMPLETE
        DoEvents
    Wend
    Set doc = iExpl.Document
End Sub

Public Property Let num(iNum As Integer)
    idx = iNum
End Property

Public Function pos() As WebPgSet
    pos.left = iExpl.left
    pos.top = iExpl.top
End Function

Public Sub dims(pgDim As WebPgSet)
    Dim pHeight As Long, pWdith As Long
    
    pHeight = min(1045 - pgDim.top, pgDim.height * 13.5)
    pwidth = min(1410 - pgDim.left, pgDim.width * 7.5)
    
    iExpl.height = pHeight
    iExpl.left = pgDim.left
    iExpl.top = pgDim.top
    iExpl.width = pwidth
End Sub

Private Sub iExpl_TitleChange(ByVal Text As String)
    Set Document = doc
    Document.cookie = idx
    TitleChange Text
End Sub
Private Sub iExpl_NameChange(ByVal Text As String)
    MsgBox iExpl.name
End Sub
Private Sub iExpl_OnQuit()
    Set ie(idx) = Nothing
End Sub

Public Sub navigate(url As String)
    iExpl.Navigate2 url
End Sub

Public Sub refresh()
    Dim tbl As String, pStl As String
    
    tbl = doc.getElementById("tblDtl").outerHTML
    pStl = nz(doc.getElementById("tblStl").outerHTML)
    navigate "javascript: docRefresh(); console.log('Done');"
    DoEvents
    iExpl.Document.getElementById("tblDtl").outerHTML = tbl
    iExpl.Document.getElementById("tblStl").outerHTML = pStl
End Sub
