Attribute VB_Name = "IPI_Paul_Outlook_ADO_Connection"
Sub adoMsAccess(cSQL As Variant, fPath As Variant, Optional cType As Variant = "Excel", Optional cString As Variant, Optional cMethod As Variant = "Excel", _
    Optional param As Variant = "" _
)
    Dim Conn As New ADODB.Connection, Cmd As New ADODB.Command, Errs As Errors, i As Integer, dbConnect As String, errLoop As Error, strTmp As String, rst As New ADODB.Recordset
    Dim Pm As ADODB.Parameter, strm As New ADODB.Stream
    
    If Not shrIPI.aRst.State = 0 Then shrIPI.aRst.Close
    
    If Not cType = "Excel" Then
        dbConnect = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};" & _
            "Dbq=" & fPath & ";" & _
            "DefaultDir=C:\;" & _
            "Uid=Admin;" & _
            "Pwd=;"
    End If
    
'    Connection Object Methods

    On Error GoTo AdoError

    With Conn
        If cType = "Excel" Then
            .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fPath & ";Extended Properties=""Excel 12.0 Macro;MODE=READ;READONLY=TRUE;HDR=YES"""
        ElseIf cType = "OLEDB" Then
            .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fPath & ";"
        End If
        If isInArray(Array("Excel", "OLEDB"), cType) Then
            .Mode = adModeShareDenyNone
            .CursorLocation = 3
            .Open .ConnectionString
        End If
        If cType = "ConnectionString" Then
            .ConnectionString = dbConnect
            .Open
        ElseIf isInArray(Array("DSN", "ODBC"), cType) Then
            .Open dbConnect
        End If
        If isInArray(Array("DSN", "ODBC"), cType) Then .Close
    End With
    
'    Recordset Object Methods
'    Don't assume we have an object
    On Error GoTo AdoErrorLite
    
    With Conn
        If isInArray(Array("Command Execute", "Command Open", "Parameterised"), cMethod) Then
            .ConnectionString = dbConnect
            .Open
            DoEvents
        End If
        If cMethod = "OLEDB" Then
            rst.Open cSQL, .ConnectionString
            DoEvents
            If Not rst.BOF Then rst.MoveFirst
            DoEvents
        ElseIf cMethod = "Connection Execute" Then
            .Open dbConnect
            Set rst = .Execute(cSQL)
        ElseIf cMethod = "Command Execute" Then
            With Cmd
                .ActiveConnection = Conn
                DoEvents
                .CommandText = cSQL
                Set rst = .Execute
                DoEvents
            End With
        ElseIf cMethod = "Command Open" Then
            With Cmd
                .ActiveConnection = Conn
                .CommandText = cSQL
            End With
            rst.Open Cmd
        ElseIf cMethod = "Parameterised" Then
            With Cmd
                .ActiveConnection = Conn
                .CommandText = cSQL
                Set Pm = .CreateParameter("long_text", 203, 1, 200000) '203 = adLongVarWChar, 1 = adParamInput
                Pm.Value = param
                .Parameters.Append Pm
                Set rst = .Execute
            End With
        ElseIf cMethod = "Recordse Open" Then
            rst.Open cSQL, dbConnect, adOpenForwardOnly
        End If
        If isInArray(Array("Connection Execute", "OLEDB"), cMethod) Then
            rst.Save strm
            shrIPI.aRst.Open strm
            strm.Close
        End If
        If Not rst.State = 0 Then rst.Close
        Conn.Close
    End With
Done:
    Set Cmd = Nothing
    Set Conn = Nothing
    Set strm = Nothing
    Exit Sub
AdoError:
    i = 1
    On Error Resume Next
    
'    Enumerate errors collection and display properties of
'    each error object (if errors collection is filed out)
    Set Errs = Conn.Errors
    For Each errLoop In Errs
        With errLoop
            strTmp = strTmp & vbCrLf & "ADO Error # " & i & ":" & _
                vbCrLf & "  ADO Error # " & .Number & _
                vbCrLf & "  Description " & .Description & _
                vbCrLf & "  Source      " & .Source
            i = i + 1
        End With
    Next
AdoErrorLite:
'    Get VB Error objects information
    With Err
        strTmp = strTmp & vbCrLf & "VB Error # " & Str(Err.Number) & _
            vbCrLf & "  Generated by " & .Source & _
            vbCrLf & "  Description  " & .Description
    End With
    
    MsgBox strTmp
    
'    Clean up gracefully without risking infinite loop in error handler
    On Error GoTo 0
    GoTo Done
End Sub

Sub adoSQL(cSQL As String, Optional db As String = "LocalDb", Optional cMethod As String = "", Optional param As Variant = "")
    Dim Conn As New ADODB.Connection, Cmd As New ADODB.Command, connStr As String, Pm As ADODB.Parameter, rst As New ADODB.Recordset, strm As New ADODB.Stream
    
    If Not shrIPI.aRst.State = 0 Then shrIPI.aRst.Close
    
    On Error GoTo errHere
    
    connStr = "Driver={SQL Server native Client 11.0};Server=(localdb)\MSSQLLocalDB;Database=" & db & ";Trusted_Connection=Yes"
    
    If cSQL > "" Then
        With Conn
            .ConnectionString = connStr
            .Open
            If cMethod = "Parameterised" Then
                With Cmd
                    .ActiveConnection = Conn
                    .CommandText = cSQL
                    .CommandType = 1
                    Set Pm = .CreateParameter("long_text", 203, 1, 200000) '203 = adLongVarWChar, 1 = adParamInput
                    Pm.Value = param
                    .Parameters.Append Pm
                    Set rst = .Execute
                End With
            Else
                Set rst = .Execute(cSQL)
            End If
        End With
        If Not rst.State = 0 Then
            rst.Save strm
            shrIPI.aRst.Open strm
            strm.Close
            rst.Close
        End If
        Conn.Close
    End If
exitHere:
    Set Conn = Nothing
    Set Cmd = Nothing
    Set strm = Nothing
    Exit Sub
errHere:
    If cMethod = "Parameterised" Then
        MsgBox "No Records Found!"
        If Not shrIPI.aRst.State = 0 Then shrIPI.aRst.Close
    Else
        MsgBox Err.Description
    End If
    GoTo exitHere
End Sub