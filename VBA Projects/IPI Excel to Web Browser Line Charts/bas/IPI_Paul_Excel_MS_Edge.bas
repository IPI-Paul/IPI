Attribute VB_Name = "IPI_Paul_Excel_MS_Edge"
Sub LoadHtml(frm As FormView, Optional html As String = "<html><body><h2>Hello from Excel</h2></body></html>", Optional large As Boolean = False)
    Dim browser As Object
    
    Set browser = CreateObject("ExcelWebView2Wrapper.WebViewHost")
    
    If frm.Title > "" Or frm.Top > 0 Or frm.Left > 0 Or frm.Width > 0 Or frm.Height > 0 Then
        browser.Show frm.Title, frm.Top, frm.Left, frm.Width, frm.Height
    Else
        browser.Show
    End If
    
    If large Then
        browser.LoadHtmlInMemory html
    Else
        browser.LoadHtml html
    End If
End Sub

Sub InjectScript()
    Dim browser As Object
    
    Set browser = CreateObject("ExcelWebView2Wrapper.WebViewHost")
    
    browser.Show
    
    browser.LoadHtml "<html><body><h2>Hello from Excel</h2></body></html>"
    browser.ExecuteScript "document.body.style.background='lightblue';"
End Sub

Sub NavigateTo(url As String)
    Dim browser As Object
    
    Set browser = CreateObject("ExcelWebView2Wrapper.WebViewHost")
    
    browser.Show
    browser.Navigate url
End Sub

