#[System.Windows.MessageBox]::Show($html);
function buildFilter() {
        $flt = $null;
        foreach ($row in $gridFilter.Rows) {
            if ($row.Cells[1].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " = ";
                $flt = $flt + (checkType -row $row) + $row.Cells[1].Value + (checkType -row $row);
            }
            if ($row.Cells[2].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + "[" + $row.Cells[0].Value + "] like '%" + $row.Cells[2].Value + "%'";
            }
            if ($row.Cells[3].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + "not [" + $row.Cells[0].Value + "] like '%" + $row.Cells[3].Value + "%'";
            }
            if ($row.Cells[4].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " in (";
                $flt = $flt + (checkType -row $row) + ($row.Cells[4].Value.split(",") -join (checkType -row $row) + "," + (checkType -row $row)) + (checkType -row $row) + ")";
            }
            if ($row.Cells[5].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + "not " + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " in (";
                $flt = $flt + (checkType -row $row) + ($row.Cells[5].Value.split(",") -join (checkType -row $row) + "," + (checkType -row $row)) + (checkType -row $row) + ")";
            }
            if ($row.Cells[6].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " >= ";
                $flt = $flt + (checkType -row $row) + $row.Cells[6].Value + (checkType -row $row);
            }
            if ($row.Cells[7].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " <= ";
                $flt = $flt + (checkType -row $row) + $row.Cells[7].Value + (checkType -row $row);
            }
        }
        return $flt;
}
function buildQuery($idx, $rIdx) {
    if ($idx -ne 0) {
        $sql = $null;
        foreach ($row in $gridFilter.Rows) {
            if ($row.Cells[14].Value -eq $true) {
                if ($sql -gt $null) {
                    $sql = $sql + [Environment]::NewLine + ",";
                    $whr = $whr + [Environment]::NewLine;
                } else {
                    $sql = $sql + [Environment]::NewLine;
                    $whr = "where " + [Environment]::NewLine + "not " + [Environment]::NewLine;
                }
                $sql = $sql + (formatColumns -row $row -col (grouping -row $row -col ((convertType -row $row -col ("[" + $row.Cells[0].Value + "]")))));
                if ($row.Cells[14].Value -eq $true -and ($row.Cells[10].Value -in ("cdbl", "cdate") -or $row.Cells[11].Value -gt "" -or $row.Cells[12].Value -in ("avg", "first", "last", "min", "max", "sum"))) {
                    $sql = $sql + " as [" + $row.Cells[0].Value + "] ";
                }
                $whr = $whr + "[" + $row.Cells[0].Value + "] & ";
            }
        }
        $whr = $whr.Substring(0, ($whr.Length -2)) + " = """"";
        $sql = "Select $sql " + [Environment]::NewLine + [Environment]::NewLine + "from " + [Environment]::NewLine + "[" + $cboNames.Text + "]";
        $sql = $sql + [Environment]::NewLine + [Environment]::NewLine + $whr + (buildFilter) + (groupBy) + (having);
        if ($rIdx -eq $oIdx[2]) {
            $sql = $sql + (orderBY);
        } 
        $txtSQL.Text = $sql
    }
}
function checkType($row) {
    $flt = $null;
    if ($row.Cells[9].Value -notin ("Double", "DateTime", "Decimal", "Int16", "Int32", "Int64", "Single", "TimeSpan", "UInt16", "UInt32", "UInt64") -and 
            $row.Cells[10].Value -notin ("cdbl", "cdate")) {
        $flt = "'";
    }
    if ($row.Cells[9].Value -in ("DateTime", "TimeSpan") -or $row.Cells[10].Value -eq "cdate") {
        $flt = "#";
    }
    return $flt;
}
function cleanUpExcel($filepath) {
    try {
        $xls = Get-Process | where {$_.ProcessName -like "Excel"};
        foreach ($xl in $xls) {
            $me = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application");
            $me.Workbooks($filepath.Split("\")[$filepath.Split("\").length-1]).Close($false);
            if ($xl.MainWindowTitle -eq "") {
                spps -Id $xl.Id;
            } 
        }
    } catch {}
}
function convertType($row, $col) {
    if ($row.Cells[10].Value -in ("cdbl", "cdate")) {
        return ($row.Cells[10].Value + "(iif($col = """", 0, iif(isnull($col) = true, 0, $col)))"); 
    } else {
        return $col; 
    }
}
function formatColumns($row, $col) {
    if ($row.Cells[11].Value -gt "") {
        return ("format($col,""" + $row.Cells[11].Value + """)");
    } else  {
        return $col;
    }
}
function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")| Out-Null;

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog;
    $OpenFileDialog.InitialDirectory= $initialDirectory;
    $OpenFileDialog.Filter = "All Files (*.*)| *.*";
	$OpenFileDialog.ShowDialog()| Out-Null;
    $OpenFileDialog.FileName;
}
function getNames() {
        $filepath = $txtFilePath.Text;
        cleanUpExcel -filepath $filepath;
        $app = New-Object -ComObject Excel.Application;
        $wb = $app.Workbooks.Open($filepath);
        
        foreach ($nm in $wb.Names) {
            if ($nm.Visible -eq $true -and $nm.Name -notcontains '!') {
                $cboNames.Items.Add($nm.Name);
            }
        }
        foreach ($sh in $wb.Sheets) {
            $cboNames.Items.Add($sh.Name + "$");
        }
                    
        $wb.Close($false);
        $app.Quit();
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app);
        cleanUpExcel -filepath $filepath;    
}
function groupBy() {
    $grp = $null;
    $grpd = $null;
    foreach ($row in $gridFilter.Rows) {
        if ($row.Cells[12].Value -gt "") {
            $grpd = $true;
            break;
        }
    }
    foreach ($row in $gridFilter.Rows) {
        if ($grpd -eq $true) {
            if ($row.Cells[12].Value -eq "group" -or ($row.Cells[12].Value -eq $null -and $row.Cells[14].Value -eq $true)) {
                if ($grp -gt $null) {
                    $grp = $grp + [Environment]::NewLine + ",";
                } else {
                    $grp = $grp + [Environment]::NewLine;
                }
                $grp = $grp + "[" + $row.Cells[0].Value + "]";
            } 
        }
    }
    if ($grpd -eq $true) {
        $grp = [Environment]::NewLine + [Environment]::NewLine + "Group By " + $grp;
    }
    return $grp;
}
function grouping($row, $col) {
    if ($row.Cells[14].Value -eq $true -and $row.Cells[12].Value -in ("avg", "first", "last", "min", "max", "sum")) {
        return ($row.Cells[12].Value + "($col)"); 
    } else {
        return $col;
    }
}
function having() {
    $hav = $null;
    foreach ($row in $gridFilter.Rows) {
        if ($row.Cells[8].Value -gt "") {
            if ($grp -gt $null) {
                $hav = $hav + [Environment]::NewLine + ",";
            } else {
                $hav = $hav + [Environment]::NewLine;
            }
            if ($row.Cells[8].Value -split "" -notcontains "=" -and $row.Cells[8].Value -split "" -notcontains "<" -and $row.Cells[8].Value -split "" -notcontains ">") {
                $hav = $hav + (grouping -row $row -col ((convertType -row $row -col ("[" + $row.Cells[0].Value + "]")))) + " = " + $row.Cells[8].Value;
            } else {
                $hav = $hav + (grouping -row $row -col ((convertType -row $row -col ("[" + $row.Cells[0].Value + "]")))) + $row.Cells[8].Value;
            }
        } 
    }
    if ($hav -gt $null) {
        $hav = [Environment]::NewLine + [Environment]::NewLine + "Having " + $hav;
    }
    return $hav;
}
function orderBy() {
    if ($oIdx[0] -gt $null -and $oIdx[0] -gt "" -and $oIdx -lt $gridFilter.Rows[$oIdx[1]].Cells[13].Value) {
        foreach ($row in $gridFilter.Rows) {
            if ($row.Index -ne $oIdx[1] -and $row.Cells[13].Value -ne "" -and $row.Cells[13].Value -ne $null -and $row.Cells[13].Value -gt $oIdx[0]) {
                $oIdx[2] = $row.Index;
                $nVal = ($row.Cells[13].Value -as "Double") - 1;
                $row.Cells[13].Value = "$nVal";
            }
        }
    }
    foreach ($row in $gridFilter.Rows) {
        if ($row.Index -ne $oIdx[1] -and $row.Cells[13].Value -ne "" -and $row.Cells[13].Value -ne $null -and $row.Cells[13].Value -eq $gridFilter.Rows[$oIdx[1]].Cells[13].Value) {
            foreach ($row in $gridFilter.Rows) {
                if ($row.Index -ne $oIdx[1] -and $row.Cells[13].Value -ne "" -and $row.Cells[13].Value -ne $null -and $row.Cells[13].Value -ge $gridFilter.Rows[$oIdx[1]].Cells[13].Value) {
                    $oIdx[2] = $row.Index;
                    $nVal = ($row.Cells[13].Value -as "Double") + 1;
                    $row.Cells[13].Value = "$nVal";
                }
            }
        }
    }
    $ord = @{};
    foreach ($row in $gridFilter.Rows) {
        if ($row.Cells[13].Value -ne "" -and $row.Cells[13].Value -ne $null) {
            $ord[(($row.Cells[13].Value -as "int") - 1)] = $row.Cells[0].Value;
        }
    }
    if ($ord[0] -gt "") {
        $ordBy = $null;
        for ($i = 0; $i -lt $ord.Count; $i++) {
            if ($ordBy -gt $null) {
                $ordBy = $ordBy + [Environment]::NewLine + ",";
            } else {
                $ordBy = $ordBy + [Environment]::NewLine;
            }
            $ordBy = $ordBy + "[" + $ord[$i] + "]";
        }
        $ordBy = [Environment]::NewLine + [Environment]::NewLine + "Order By "+ $ordBy;
        return $ordBy;
    } else {
        return $null;
    }
}
function resizeItems() {
    $gridResult.Size = New-Object System.Drawing.Size(($objForm.Width - 20), ($objForm.Height - 70)); 
    $txtSQL.Size = New-Object System.Drawing.Size((($objForm.Width - 20) / 4), ($objForm.Height - 70));
    $gridFilter.Location = New-Object System.Drawing.Size(((($objForm.Width - 20) / 4) + 2),25);
    $gridFilter.Size = New-Object System.Drawing.Size(((($objForm.Width - 28) / 4) * 3), ($objForm.Height - 70));
}
function runAction() {
    if ($cboAction.SelectedIndex -gt 0) {
        if ($cboAction.Text -eq "Edit Query") {
            $gridFilter.Visible = $true;
            $txtSQL.Visible = $true;
            $gridResult.Visible = $false;
        } elseif ($cboAction.Text -eq "Run Query") {
            $gridFilter.Visible = $false;
            $txtSQL.Visible = $false;
            $gridResult.Visible = $true;
            runUpdSQL;
        } elseif ($cboAction.Text -eq "Send To Open Email") {
            sendToOutlook;
        } elseif ($cboAction.Text -eq "View Previous Results") {
            $gridFilter.Visible = $false;
            $txtSQL.Visible = $false;
            $gridResult.Visible = $true;
        }

        $cboAction.SelectedIndex = 0;
    }
}
function runUpdSQL() {
        $filepath = $txtFilePath.Text;
        if ($cboDriver.Text -eq "JET") {
            cleanUpExcel -filepath $filepath;
            $app = New-Object -ComObject Excel.Application;
            $wb = $app.Workbooks.Open($filepath);
            $connString = "Provider = Microsoft.JET.OLEDB.4.0; Data Source=" + $filepath + ";Extended Properties=""Excel 8.0;IMEX=1;HDR=YES"";Persist Security Info=False;";
        } else {
            $connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + $filepath + ";Extended Properties=""Excel 12.0;IMEX=1;HDR=YES"";Persist Security Info=False;";
        }
        $conn = New-Object "System.Data.OleDb.OleDbConnection" $connString;
        $comm = New-Object "System.Data.OleDb.OleDbCommand";
        $commType = [System.Data.CommandType]"Text";
        $comm.CommandText = ($txtSQL.Text);
        $comm.Connection = $conn;

        $conn.Open();
        $adapter = New-Object "System.Data.OleDb.OleDbDataAdapter" $comm;
        $dt = New-Object System.Data.DataSet;
        $adapter.Fill($dt);
        $html = ($dt.Tables[0] | select * -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray | ConvertTo-Html -As Table -Fragment -Property *);
        Set-Variable -Name "html" -Value ($html) -Scope Global;

        if ($dt -ne $null){
            $gridResult.DataSource = $dt.Tables[0];
            $gridResult.Update();
        }

        $comm.Dispose();
        $conn.Close();
        $conn.Dispose();            
        if ($cboDriver.Text -eq "JET") {
            $wb.Close($false);
            $app.Quit();
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app);
            cleanUpExcel -filepath $filepath;
        }
}
function sendToOutlook() {
    $ol = New-Object -ComObject Outlook.Application
    $ol.ActiveInspector().WordEditor.Application.Selection = "placeHere";
    $ol.ActiveInspector().CurrentItem.htmlBody = $ol.ActiveInspector().CurrentItem.htmlBody.replace("placeHere", "$html");
}
function updateFilter() {
        $gridFilter.Visible = $true;
        $txtSQL.Visible = $true;
        $gridResult.Visible = $false;
        $filepath = $txtFilePath.Text;
        if ($cboDriver.Text -eq "JET") {
            cleanUpExcel -filepath $filepath;
            $app = New-Object -ComObject Excel.Application;
            $wb = $app.Workbooks.Open($filepath);
            $connString = "Provider = Microsoft.JET.OLEDB.4.0; Data Source=" + $filepath + ";Extended Properties=""Excel 8.0;IMEX=1;HDR=YES"";Persist Security Info=False;";
        } else {
            $connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + $filepath + ";Extended Properties=""Excel 12.0;IMEX=1;HDR=YES"";Persist Security Info=False;";
        }
        $conn = New-Object "System.Data.OleDb.OleDbConnection" $connString;
        $comm = New-Object "System.Data.OleDb.OleDbCommand";
        $commType = [System.Data.CommandType]"Text";
        $comm.CommandText = "select top 1 * from [" + $cboNames.Text + "]";
        $comm.Connection = $conn;

        $conn.Open();
        $adapter = New-Object "System.Data.OleDb.OleDbDataAdapter" $comm;
        $dt = New-Object System.Data.DataSet;
        $adapter.Fill($dt);

        $gridFilter.Rows.Clear();

        if ($dt -ne $null){
            $i = 1
            foreach ($col in $dt.Tables[0].Columns) {
                $cboIndex.Items.Add("$i");
                $i++;
            }
            foreach ($col in $dt.Tables[0].Columns) {
                $gridFilter.Rows.Add($col.ColumnName, $null, $null, $null, $null, $null, $null, $null, $null, $col.DataType.Name.ToString());
            }
        }

        $comm.Dispose();
        $conn.Close();
        $conn.Dispose();            
        if ($cboDriver.Text -eq "JET") {
            $wb.Close($false);
            $app.Quit();
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app);
            cleanUpExcel -filepath $filepath;
        }
}
function viewForm() {
    $objForm = New-Object System.Windows.Forms.Form;
    $objForm.text = "IPI International Excel Connections";
    $objForm.Size = New-Object System.Drawing.Size(1280,720);
    $objForm.StartPosition = "CenterScreen";

    $cboDriver = New-Object System.Windows.Forms.ComboBox;
    $cboDriver.Location = New-Object System.Drawing.Size(2, 2);
    $cboDriver.Size = New-Object System.Drawing.Size(49,20);
    $cboDriver.Items.Add('ACE');
    $cboDriver.Items.Add('JET');
    $cboDriver.SelectedIndex = 0;
    $objForm.Controls.Add($cboDriver);

    $txtFilePath = New-Object System.Windows.Forms.TextBox;
    $txtFilePath.Location = New-Object System.Drawing.Size(52,2);
    $txtFilePath.Size = New-Object System.Drawing.Size(660,25);
    $objForm.Controls.Add($txtFilePath);

    $btnUpdate = New-Object System.Windows.Forms.Button;
    $btnUpdate.Location = New-Object System.Drawing.Size(712,2);
    $btnUpdate.Size = New-Object System.Drawing.Size(30,21);
    $btnUpdate.Text = "...";
    $btnUpdate.Add_Click({$txtFilePath.Text = Get-FileName  -initialDirectory "$home\Documents"; getNames;});
    $objForm.Controls.Add($btnUpdate);

    $cboNames = New-Object System.Windows.Forms.ComboBox;
    $cboNames.Location = New-Object System.Drawing.Size(744,2);
    $cboNames.Size = New-Object System.Drawing.Size(300,21);
    $cboNames.Items.Add('');
    $cboNames.Add_SelectedIndexChanged({updateFilter;});
    $objForm.Controls.Add($cboNames);

    $cboAction = New-Object System.Windows.Forms.ComboBox;
    $cboAction.Location = New-Object System.Drawing.Size(1047,2);
    $cboAction.Size = New-Object System.Drawing.Size(210,21);
    $cboAction.Items.Add('');
    $cboAction.Items.Add('Edit Query');
    $cboAction.Items.Add('Run Query');
    $cboAction.Items.Add('Send To Open Email');
    $cboAction.Items.Add('View Previous Results');
    $cboAction.Add_SelectedIndexChanged({runAction;});
    $objForm.Controls.Add($cboAction);

    $txtSQL = New-Object System.Windows.Forms.TextBox;
    $txtSQL.Multiline = $true;
    $txtSQL.Location = New-Object System.Drawing.Size(2,25);
    $txtSQL.Size = New-Object System.Drawing.Size((($objForm.Width - 20) / 4), ($objForm.Height - 70));
    $objForm.Controls.Add($txtSQL);

    $gridFilter = New-Object System.Windows.Forms.DataGridView;
    $gridFilter.Location = New-Object System.Drawing.Size(((($objForm.Width - 20) / 4) + 2),25);
    $gridFilter.Size = New-Object System.Drawing.Size(((($objForm.Width - 28) / 4) * 3), ($objForm.Height - 70));
    $gridFilter.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True;
    $gridFilter.AutoSize = $false;
    $gridFilter.AutoSizeRowsMode = "AllCells";
    $gridFilter.AutoSizeColumnsMode = "AllCells";
    $gridFilter.ColumnCount = 10;
    $gridFilter.Columns[0].Name = "Column Name";
    $gridFilter.Columns[1].Name = "Equals";
    $gridFilter.Columns[2].Name = "Like";
    $gridFilter.Columns[3].Name = "Not Like";
    $gridFilter.Columns[4].Name = "Is In";
    $gridFilter.Columns[5].Name = "Is Not In";
    $gridFilter.Columns[6].Name = "From";
    $gridFilter.Columns[7].Name = "To";
    $gridFilter.Columns[8].Name = "Having";
    $gridFilter.Columns[9].Name = "Type";
    $gridFilter.Add_Click({Set-Variable -Name "oIdx" -Value ($gridFilter.CurrentRow.Cells[13].Value, $gridFilter.CurrentRow.Index, $gridFilter.CurrentRow.Index) -Scope Global;});
    $gridFilter.add_CellValueChanged({buildQuery -idx $_.ColumnIndex -rIdx $gridFilter.CurrentRow.Index;});
    $cboConvert = New-Object System.Windows.Forms.DataGridViewComboBoxColumn;
    $cboConvert.Name = "Convert To";
    $cboConvert.Width = 50;
    $cboConvert.Items.Add("");
    $cboConvert.Items.Add("cdbl");
    $cboConvert.Items.Add("cdate");
    $gridFilter.Columns.Add($cboConvert);
    $cboFormat = New-Object System.Windows.Forms.DataGridViewComboBoxColumn;
    $cboFormat.Name = "Format";
    $cboFormat.Width = 50;
    $cboFormat.Items.Add("");
    $cboFormat.Items.Add("#,###,###,##0.00");
    $cboFormat.Items.Add("#,###,###,##0");
    $cboFormat.Items.Add("#,###,###,##0%");
    $cboFormat.Items.Add("#,###,###,##0.00%");
    $cboFormat.Items.Add("dd/mm/yyyy hh:mm");
    $cboFormat.Items.Add("mmm-yy");
    $cboFormat.Items.Add("hh:mm");
    $gridFilter.Columns.Add($cboFormat);
    $cboGroup = New-Object System.Windows.Forms.DataGridViewComboBoxColumn;
    $cboGroup.Name = "Group By";
    $cboGroup.Width = 50;
    $cboGroup.Items.Add("");
    $cboGroup.Items.Add("group");
    $cboGroup.Items.Add("first");
    $cboGroup.Items.Add("last");
    $cboGroup.Items.Add("min");
    $cboGroup.Items.Add("max");
    $cboGroup.Items.Add("avg");
    $cboGroup.Items.Add("sum");
    $gridFilter.Columns.Add($cboGroup);
    $cboIndex = New-Object System.Windows.Forms.DataGridViewComboBoxColumn;
    $cboIndex.Name = "Index";
    $cboIndex.Width = 50;
    $cboIndex.Items.Add("");
    $gridFilter.Columns.Add($cboIndex);
    $chkShow = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn;
    $chkShow.Name = "Show";
    $chkShow.Width = 30;
    $gridFilter.Columns.Add($chkShow);
    $objForm.Controls.Add($gridFilter);

    $gridResult = New-Object System.Windows.Forms.DataGridView;
    $gridResult.Location = New-Object System.Drawing.Size(2,25);
    $gridResult.Size = New-Object System.Drawing.Size(($objForm.Width - 20), ($objForm.Height - 70));
    $gridResult.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True;
    $gridResult.AutoSize = $false;
    $gridResult.Visible = $false;
    $gridResult.AutoSizeRowsMode = "AllCells";
    $gridResult.AutoSizeColumnsMode = "AllCells";
    $objForm.Controls.Add($gridResult);

    $objForm.TopMost = $False;
    $objForm.Add_Shown({$objForm.Activate()});
    $objForm.Add_Resize({resizeItems;});
    [void]$objForm.ShowDialog();        
}
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing");
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");

$html = $null;
$oIdx = $null;
viewForm;
