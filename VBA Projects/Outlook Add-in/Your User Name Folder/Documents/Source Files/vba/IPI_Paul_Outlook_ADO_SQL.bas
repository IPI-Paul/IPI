Attribute VB_Name = "IPI_Paul_Outlook_ADO_SQL"

Function runSQL(Optional db As String = "", Optional dbPath As String = "", Optional sql As String = "", Optional tp As Variant = "", Optional dLim As String = "") As ADODB.Recordset
    Dim connString As String, conn As New ADODB.Connection, rs As New ADODB.Recordset
    
    On Error GoTo errHere
    If tp = "Excel Test" Then
        With conn
            .CursorLocation = adUseClient
            .ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" & _
                "DBQ=" & dbPath & db & ";DefaultDir=C:\;Uid=Admin;Pwd=;"
            .Open
        End With
    ElseIf tp = "Ms Access Test" Then
        With conn
            .Mode = adModeShareDenyNone
            .Provider = "Microsoft.ACE.OLEDB.15.0"
            .ConnectionString = "Data Source=" & dbPath & db & ";"
            .Open
        End With
    ElseIf allIsInArray(Array("SQL Compact Edition Test", "SQL Local Db Test", "SQL Server Test"), tp) Then
        If tp = "SQL Compact Edition Test" Then
            connString = "Provider=Microsoft.SQLSERVER.CE.OLEDB.4.0;Persist Security Info=False;Data Source=" & dbPath & db & ";SSCE:Max Buffer Size=4096;"
        ElseIf tp = "SQL Local Db Test" Then
            connString = "Driver={SQL Server native Client 11.0};Server=(LocalDB)\MSSQLLocalDB;AttachDBFileName=" & dbPath & db & ";Database=" & db & ";Trusted_Connection=Yes;IntegratedSecurity=SSPI;"
        ElseIf tp = "SQL Server Test" Then
            connString = "Driver={SQL Server native Client 11.0};Server=Your_Domain\Your_Server;Database=" & db & ";Trusted_Connection=Yes;IntegratedSecurity=SSPI;"
        End If
        
        With conn
            .ConnectionTimeout = 120
            .Open connString
        End With
    ElseIf tp = "Text Test" Then
        With conn
            .Mode = adModeShareDenyNone
            .Provider = "Microsoft.ACE.OLEDB.15.0"
            .ConnectionString = "Data Source=" & dbPath & ";Extended Properties='text" & dLim & "'"
            .Open
        End With
    End If
    
    Set rs = conn.Execute(sql)
    If Not rs.BOF Then rs.MoveFirst
    Set runSQL = rs

exitHere:
    Exit Function
errHere:
    MsgBox Err.Description
    GoTo exitHere
End Function

