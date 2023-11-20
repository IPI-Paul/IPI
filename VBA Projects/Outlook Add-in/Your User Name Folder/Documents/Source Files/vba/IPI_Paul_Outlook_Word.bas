Attribute VB_Name = "IPI_Paul_Outlook_Word"

Public Sub buildWdDoc(doc, htmlTbl As HTMLTable, bgCol As Long, fgCol As Long)
    Set tbl = doc.Application.Selection.Tables.add(doc.Application.Selection.Range(), htmlTbl.Rows.Length, htmlTbl.Rows(0).cells.Length)
    With tbl.Borders
        .InsideLineStyle = wdLineStyleSingle
        .OutsideLineStyle = wdLineStyleSingle
        .InsideColor = bgCol
        .OutsideColor = bgCol
    End With
    pg = tbl.Rows(1).Range.Information(wdActiveEndPageNumber)
    rws = 1
    For i = 0 To htmlTbl.Rows.Length - 1
        If pg < tbl.Rows(rws).Range.Information(wdActiveEndPageNumber) Then
            pg = tbl.Rows(rws).Range.Information(wdActiveEndPageNumber)
            tbl.Rows.add (tbl.Rows(rws))
            For j = 0 To htmlTbl.Rows(0).cells.Length - 1
                With tbl.Rows(rws).cells(j + 1).Range
                    .Text = htmlTbl.Rows(0).cells(j).innerText
                    .Shading.BackgroundPatternColor = bgCol
                    .Font.Color = fgCol
                    .Font.Bold = True
                End With
            Next j
            rws = rws + 1
        End If
        If i = 0 Then
            For j = 0 To htmlTbl.Rows(0).cells.Length - 1
                With tbl.Rows(rws).cells(j + 1).Range
                    .Text = htmlTbl.Rows(0).cells(j).innerText
                    .Shading.BackgroundPatternColor = bgCol
                    .Font.Color = fgCol
                    .Font.Bold = True
                End With
            Next j
        Else
            For j = 0 To htmlTbl.Rows(0).cells.Length - 1
                tbl.Rows(rws).cells(j + 1).Range.Text = htmlTbl.Rows(i).cells(j).innerText
                tbl.Rows(rws).cells(j + 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            Next j
        End If
        rws = rws + 1
    Next i
'    With tbl.Rows(rws).cells
'        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
'        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
'    End With
'    rws = rws + 1
'    For j = 0 To lboTotals.ColumnCount - 1
'        tbl.Rows(rws).cells(j + 1).Range.Text = restyle(2, lboTotals.List(0, j))
'        tbl.Rows(rws).cells(j + 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
'    Next j
    Set tbl = Nothing
End Sub

