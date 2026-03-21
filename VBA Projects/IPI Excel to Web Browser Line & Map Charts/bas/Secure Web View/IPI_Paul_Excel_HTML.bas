Attribute VB_Name = "IPI_Paul_Excel_HTML"
Option Explicit

Sub showWebPage(Optional sh As String = "Temperature All Locations")
    Dim fso As New FileSystemObject, fPath As String, dta As String, hWnd As LongPtr, api As LongPtr, head As String
    Dim Top As Long, Left As Long, Height As Long, Width As Long, pDim As Variant, AppScript As String, rw As Range, itm As Variant, html As String
    Dim BubbleMap As String, DateHistogram As String, frm As FormView
    
    fPath = Environ$("appdata") & "\..\Local\Temp\IPI Excel to Web Browser Line Charts.html"
    html = ""
    head = Replace(Sheet1.HTMLHead.Value, "IPI Excel to Web Browser Line Charts", "IPI Excel to Web Browser " & Range("ChartType").Value & " Charts")
    dta = ""
    For Each rw In Sheets(sh).UsedRange.Rows
        dta = dta & Chr(34) & WorksheetFunction.TextJoin(Chr(34) & "," & Chr(34), False, rw) & Chr(34) & vbLf
    Next rw
    If Range("ChartType").Value = "Line" Then
        AppScript = Sheet1.AppScript.Value
        For Each itm In Array(Array("'timestamp'", "'" & Range("xAxis").Value & "'"), Array("'temperature'", "'" & Range("yAxis").Value & "'"), _
                Array("d.city", "d['" & Range("Legend").Value & "']"), Array("A Week of Temperature Around the World", Range("ChartTitle").Value), _
                Array("// xTickFormat", IIf(Range("xTickFormat").Value > "", ".tickFormat(timeFormat(" & Range("xTickFormat").Value & "))", "// xTickFormat")), _
                Array("// yTickFormat", IIf(Range("yTickFormat").Value > "", ".tickFormat(timeFormat(" & Range("yTickFormat").Value & "))", "// yTickFormat")), _
                Array("// LineMarkers", IIf(Range("LineMarkers").Value = "Yes", Sheet1.LineMarks.Value, "// LineMarkers")) _
        )
            AppScript = Replace(AppScript, CStr(itm(0)), CStr(itm(1)))
        Next itm
        For Each itm In Array("<!DOCTYPE html>", "<html>", "<head>", head, Sheet1.LineStyle.Value, "<script id='ipi'>", Sheet1.IPIScript.Value, "</script>", _
                "<script id='d3'>", Sheet1.D3Script.Value, "</script>", "<script id='colorApp'>", Sheet1.ColourLegend.Value, "</script>", _
                "<script id='tickMarksApp'>", Sheet1.MarkScript.Value, "</script>", "<script id='dropdownMenu'>", Sheet1.MenuScript.Value, "</script>", _
                "<script id='data'>", Sheet1.DataOpen.Value, dta, Sheet1.DataClose.Value, "</script>", "</head>", "<body>", Sheet1.HTMLBody.Value, _
                "<script id='App'>", AppScript, "</script>", "</body>", "<html>")
            html = html & CStr(itm)
        Next itm
    Else
        AppScript = Sheet1.MapAppScript.Value
        BubbleMap = Sheet1.BubbleMap.Value
        DateHistogram = Sheet1.DateHistogram.Value
        For Each itm In Array(Array("'Location Coordinates'", "'" & Range("Coordinates") & "'"), Array("'Total Dead and Missing'", "'" & Range("yAxis").Value & "'"), _
                Array("'Reported Date'", "'" & Range("xAxis").Value & "'"))
            AppScript = Replace(AppScript, CStr(itm(0)), CStr(itm(1)))
            BubbleMap = Replace(BubbleMap, CStr(itm(0)), CStr(itm(1)))
            DateHistogram = Replace(DateHistogram, CStr(itm(0)), CStr(itm(1)))
        Next itm
        For Each itm In Array("<!DOCTYPE html>", "<html>", "<head>", head, Sheet1.MapStyle.Value, "<script id='ipi'>", Sheet1.IPIScript.Value, "</script>", _
                "<script id='d3'>", Sheet1.D3Script.Value, "</script>", "<script id='topojson'>", Sheet1.TopoJSON.Value, "</script>", _
                "<script id='countries50m'>", Sheet1.Countries50m.Value, "</script>", "<script id='mapData'>", Sheet1.DataOpen.Value, dta, Sheet1.DataClose.Value, "</script>", _
                "<script id='bubbleMarks'>", Sheet1.BubbleMarks.Value, "</script>", "<script id='bubbleMap'>", BubbleMap, "</script>", _
                "<script id='axisBottom'>", Sheet1.AxisBottom.Value, "</script>", "<script id='axisLeft'>", Sheet1.AxisLeft.Value, "</script>", _
                "<script id='histogramMarks'>", Sheet1.HistogramMarks.Value, "</script>", "<script id='dateHistogram'>", DateHistogram, "</script>", _
                "</head>", "<body>", Sheet1.HTMLBody.Value, "<script id='App'>", AppScript, "</script>", "</body>", "<html>")
            html = html & CStr(itm)
        Next itm
    End If
    
    frm.Title = "IPI Excel to Web Browser " & Range("ChartType").Value & " Charts"
    frm.Top = Range("ChartTop").Value
    frm.Left = Range("ChartLeft").Value
    frm.Width = Range("ChartWidth").Value
    frm.Height = Range("ChartHeight").Value
    
    LoadHtml frm, html, Len(html) > 10000
End Sub

Sub viewChart()
    showWebPage Range("SheetName").Value
End Sub
