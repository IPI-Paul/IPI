var ay, cols, db, fs, objectStore, request, flPth, nw, rn = window.name, rslt = '', PlayLst, pltb, prmG, prmO, prmW, pos, rwId, tmpFl1, tmpFl2, tmpFl3, tmpFl, upldList, ze, zo;

function callFunction(Indx) {
    document.getElementById('cboFunction').selectedIndex = 0;
    setTimeout(function () {
        if (Indx === 1) { deleteDb(); }
        if (Indx === 2) { createTable(); }
        if (Indx === 3) { deleteStore(); }
        if (Indx === 4) { returnResults(); }
        if (Indx === 5) { viewFrame(); }
        if (Indx === 6) { viewEditFrame(); }
        if (Indx === 7) { insertValues(); }
        if (Indx === 8) { createUpload(0); }
        if (Indx === 9) { createUpload(1); }
        if (Indx === 10) { createInsert(1); }
        if (Indx === 11) { createInsert(0); }
        if (Indx === 12) { deleteRow(); }
        if (Indx === 13) { deleteAll(); }
        if (Indx === 14) { resetPlayed(); }
        if (Indx === 15) { replaceErrors(); }
        if (Indx === 16) { iframeToDocument(); }
        if (Indx === 17) { tableToDocument(); }
        if (Indx === 18) { document.getElementById('Val2').value = document.location; }
        if (Indx === 19) { copyFrame(); }
        if (Indx === 20) { document.getElementById('Val2').value = document.getElementById('frmView').contentDocument.location; }
        if (Indx === 21) { document.getElementById('Val2').value = document.getElementById('frmView').contentDocument.getElementsByTagName('video')[0].src; document.getElementById('Val5').value = document.getElementById('Val2').value; }
        if (Indx === 22) { clearEdit(); }
        if (Indx === 23) { runScript(); }
        if (Indx === 24) { hideEditor(); SrcUpd(); }
        if (Indx === 25) { formatiFiles(); }
        if (Indx === 26) { formatMobilelite(); }
        if (Indx === 27) { hideiFrame(); hideEditor(); uploadTable(); }
        if (Indx === 28) { selfUpload(); }
    }, 250);
}
function clearEdit() {
    document.getElementById('txtView').value = '';
}
function copyFrame() {
    var prm = prompt('Text to Copy!', document.getElementById('frmView').contentDocument.all[0].childNodes[1].textContent);
}
function createInsert(tp) {
    var prm = '<table id="tblResults">', vl1 = '', vl2 = '', vl3 = '', vl4 = '', vl5 = '';
    if (tp === 1) {
        prm = prm + '<tr><td>&lt;table id="tblResults"&gt;<td>';
    }
    pltb = document.getElementById('tblResults');
    for (i = 1; i < pltb.rows.length; i++) {
        prm = prm + '<tr><td>';
        if (tp === 1) {
            prm = prm + '&lt;tr&gt;&lt;td&gt;';
        }
        prm = prm + "insert into [" + document.getElementById("strNm").value + "] ([" + pltb.rows[0].childNodes[0].textContent + "],[" + pltb.rows[0].childNodes[1].textContent + "],[" + pltb.rows[0].childNodes[2].textContent + "],[" + pltb.rows[0].childNodes[3].textContent + "],[" + pltb.rows[0].childNodes[4].textContent + "]) ";
        vl1 = pltb.rows[i].childNodes[0].textContent;
        vl2 = pltb.rows[i].childNodes[1].textContent;
        vl3 = pltb.rows[i].childNodes[2].textContent;
        vl4 = pltb.rows[i].childNodes[3].textContent;
        vl5 = pltb.rows[i].childNodes[4].textContent;
        vl1 = vl1.replace(/'/g, "''");
        vl2 = vl2.replace(/'/g, "''");
        vl4 = vl4.replace(/'/g, "''");
        vl5 = vl5.replace(/'/g, "''");
        prm = prm + "Values ('" + vl1 + "','" + vl2 + "','" + vl3 + "','" + vl4 + "','" + vl5 + "');";
        if (tp === 1) {
            prm = prm + '&lt;/td&gt;&lt;/tr&gt;';
        }
        prm = prm + '</td></tr>';
    }
    if (tp === 1) {
        prm = prm + '<tr><td>&lt;/table&gt;</td></tr>';
    }
    prm = prm + '<tr><td>--</td></tr></table>';
    document.getElementById('frmView').contentDocument.write(prm);
    document.getElementById('frmView').contentDocument.close();
    document.getElementById('frmView').height = 500;
    document.getElementById('frmView').style.visibility = 'visible';
}
function createTable() {
    tblName = document.getElementById('strNm').value;
    var idx1 = document.getElementById('Col1').value;
    var idx2 = document.getElementById('Col2').value;
    var idx3 = document.getElementById('Col4').value;
    var ver = request.result.version + 1;
    var database = request.result;
    database.close();
    request = indexedDB.open('IPIMediaDb', ver);
    request.onupgradeneeded = function (event) {
        db = event.target.result;
        objectStore = db.createObjectStore([tblName], { keyPath: 'Id', autoIncrement: true });
        objectStore.createIndex(idx1, idx1, { unique: false });
        objectStore.createIndex(idx2, idx2, { unique: false });
        objectStore.createIndex(idx3, idx3, { unique: false });
        db = request.result;
        db.close();
        request = "";
        listStores();
    };
}
function createUpload(tp) {
    var prm = '<table id="Ply">';
    pltb = document.getElementById('tblResults');
    prm = prm + "<th>" + pltb.rows[0].childNodes[0].textContent + '</th>';
    prm = prm + "<th>" + pltb.rows[0].childNodes[1].textContent + '</th>';
    prm = prm + "<th>" + pltb.rows[0].childNodes[2].textContent + '</th>';
    prm = prm + "<th>" + pltb.rows[0].childNodes[3].textContent + '</th>';
    prm = prm + "<th>" + pltb.rows[0].childNodes[4].textContent + '</th>';
    for (i = 1; i < pltb.rows.length; i++) {
        if (tp === 0 || (tp === 1 && pltb.rows[i].childNodes[5].childNodes[0].checked === true)) {
            prm = prm + '<tr>';
            prm = prm + "<td>" + pltb.rows[i].childNodes[0].textContent + '</td>';
            prm = prm + "<td>" + pltb.rows[i].childNodes[1].textContent + '</td>';
            prm = prm + "<td>" + pltb.rows[i].childNodes[2].textContent + '</td>';
            if (document.getElementById('Val4').value > '') {
                prm = prm + "<td>" + document.getElementById('Val4').value + '</td>';
            }
            else {
                prm = prm + "<td>" + pltb.rows[i].childNodes[3].textContent + '</td>';
            }
            prm = prm + "<td>" + pltb.rows[i].childNodes[4].textContent + '</td>';
            prm = prm + '</tr>';
        }
    }
    prm = prm + '</table>';
    document.getElementById('frmView').contentDocument.write(prm);
    document.getElementById('frmView').contentDocument.close();
    document.getElementById('frmView').height = 500;
    document.getElementById('frmView').style.visibility = 'visible';
}
function dbOpen() {
    request = indexedDB.open('IPIMediaDb');
    request.onerror = function (event) {
        request = '';
        alert('Database error: ' + event.target.error);
    };
}
function deleteDb() {
    if (confirm('Are you sure you want to delete the database?')) {
        var database = request.result;
        var ver = request.result.version + 1;
        database.close();
        var req = indexedDB.deleteDatabase('IPIMediaDb');
        req.onsuccess = function (e) {
            alert('Done');
        };
        req.onblocked = function (e) {
            alert('Blocked');
        };
    }
}
function deleteAll() {
    pltb = document.getElementById('tblResults');
    if (pltb.rows.length > 1) {
        if (confirm('Are you sure you want to delete all records from the ' + document.getElementById('strNm').value + ' store that are currently displayed?')) {
            if (!request) {
                dbOpen();
            }
            db = request.result;
            for (i = 1; i < pltb.rows.length; i++) {
                var req = db.transaction([tblName], 'readwrite').objectStore([tblName]).delete(parseInt(pltb.rows[i].childNodes[0].id));
            }
            returnResults();
        }
    }
    else {
        alert('Please ensure that there are some records displayed!');
    }
}
function deleteRow() {
    pltb = document.getElementById('tblResults');
    if (pltb.rows.length > 1) {
        if (confirm('Are you sure you want to delete records from the ' + document.getElementById('strNm').value + ' store that have been selected?')) {
            if (!request) {
                dbOpen();
            }
            db = request.result;
            for (i = 1; i < pltb.rows.length; i++) {
                if (pltb.rows[i].childNodes[5].childNodes[0].checked === true) {
                    var req = db.transaction([tblName], 'readwrite').objectStore([tblName]).delete(parseInt(pltb.rows[i].childNodes[0].id));
                }
            }
            returnResults();
        }
    }
    else {
        alert('Please ensure that there are some records displayed!');
    }
}
function deleteStore() {
    if (confirm('Are you sure you want to delete the ' + document.getElementById('strNm').value + ' store?')) {
        tblName = document.getElementById('strNm').value;
        var database = request.result;
        var ver = request.result.version + 1;
        database.close();
        request = indexedDB.open('IPIMediaDb', ver);
        request.onupgradeneeded = function (e) {
            db = request.result;
            db.deleteObjectStore([tblName]);
        };
    }
}
function formatiFiles() {
    var prm;
    prm = "document.getElementsByTagName('link')[0].setAttribute('href','./styles/iFiles.css'); ";
    prm = prm + "var a = '', b, c = ''; ";
    prm = prm + "a = decodeURIComponent(document.URL); ";
    prm = prm + "if (a.substring(a.length-1,a.length)=='/'){a = a.length;} else {a = a.length + 1;} ";
    prm = prm + "for (i=1;i<document.getElementsByTagName('a').length-2;i++) ";
    prm = prm + "{ ";
    prm = prm + "b = decodeURIComponent(document.getElementsByTagName('a')[i].href); ";
    prm = prm + "b = b.substring(a,b.length); ";
    prm = prm + "if (b.search('javascript') < 0 && b.search('#') < 0 && b.search('.config') < 0) {c = c + '<tr><td>' + b + '</td></tr>';} ";
    prm = prm + "} ";
    prm = prm + "document.write('<table id=" + '"Ply"' + ">' + c + '</table> <br />'); ";
    scrpt = 'JavaScript:' + prm;
    document.getElementById('frmView').contentWindow.document.location = scrpt;
}
function formatMobilelite() {
    var prm;
    prm = "var a, b = ''; ";
    prm = prm + "for (i=0;i<document.getElementsByTagName('a').length;i++) ";
    prm = prm + "{ ";
    prm = prm + "a = decodeURIComponent(document.getElementsByTagName('a')[i].href); ";
    prm = prm + "a = a.split('/')[a.split('/').length - 1]; ";
    prm = prm + "if (a.search('javascript') < 0 && a.search('#') < 0 && a.search('.config') < 0) {b = b + '<tr><td>' + a + '</td></tr>';} ";
    prm = prm + "} ";
    prm = prm + "document.write('<table id=" + '"Ply"' + ">' + b + '</table> <br />'); ";
    scrpt = 'JavaScript:' + prm;
    document.getElementById('frmView').contentWindow.document.location = scrpt;
}
function getFile() {
    document.getElementById('frmView').height = 580;
    document.getElementById('frmView').style.visibility = 'visible';
    document.getElementById('frmView').src = window.URL.createObjectURL(document.getElementById('getfiles').files[0]);
}
function hideEditor() {
    document.getElementById('txtView').style.visibility = 'hidden';
    document.getElementById('txtView').rows = 1;
}
function hideiFrame() {
    document.getElementById('frmView').style.visibility = 'hidden';
    document.getElementById('frmView').height = 0;
}
function iframeToDocument() {
    document.all[0].innerHTML = document.getElementById('frmView').contentDocument.getElementsByTagName('table')[0].outerHTML;
}
function Incr(a) {
    var i;
    if (document.getElementById('Whr3').value == 0) {
        i = 0;
    } else {
        i = 1;
    }
    n = parseInt(a) + i;
    b = n;
    rn = b;
    if (n === document.getElementById('rws').value) {
        window.clearInterval(PlayLst);
        document.location.reload();
    }
}
function indexSelect() {
    if (document.getElementById('idbIdx').selectedIndex !== 0) {
        document.getElementById('idxNm').value = document.getElementById('idbIdx').value;
        document.getElementById('idbIdx').selectedIndex = 0;
    }
}
function listIndexes(str) {
    var opt = '<option></option>';
    if (!request) {
        dbOpen();
    }
    db = request.result;
    idx = db.transaction(tblName).objectStore(tblName).indexNames;
    for (i = 0; i < idx.length; i++) {
        opt = opt + '<option value="' + idx[i] + '">' + idx[i] + '</option>';
    }
    document.getElementById('idbIdx').innerHTML = opt;
}
function listStores() {
    var opt = '<option></option>';
    dbOpen();
    request.onsuccess = function (event) {
        db = request.result;
        for (i = 0; i < db.objectStoreNames.length; i++) {
            opt = opt + '<option value="' + db.objectStoreNames[i] + '">' + db.objectStoreNames[i] + '</option>';
        }
        document.getElementById('idbStrs').innerHTML = opt;
        tblName = document.getElementById('strNm').value;
        listIndexes(tblName);
    };
}
function listSelect() {
    if (document.getElementById('idbStrs').selectedIndex !== 0) {
        document.getElementById('strNm').value = document.getElementById('idbStrs').value;
        document.getElementById('idbStrs').selectedIndex = 0;
    }
}
function playListCh(vl) {
    document.getElementById('frmView').src = vl;
    document.getElementById('frmView').height = 580;
    document.getElementById('frmView').style.visibility = 'visible';
}
function posStart() {
    var pos = prompt("Please specify what point to play from", 0);
    if (!pos == '') {
        if (pos.split(":").length > 2) {
            pos = parseFloat(pos.split(":")[0] * 60) + parseFloat(pos.split(":")[1]) + parseFloat(pos.split(":")[2] / 60);
        } else if (pos.search(":") >= 0) {
            pos = parseFloat(pos.split(":")[0]) + parseFloat(pos.split(":")[1] / 60);
        }
        document.getElementById("frmView").contentWindow.document.getElementById("Quicky").currentTime = pos * 60;
    }
}
function Rdm(a) {
    n = Math.round(Math.random(a) * parseInt(document.getElementById('rws').value));
    b = n;
    rn = b;
}
function replaceErrors() {
    pltb = document.getElementById('tblResults');
    if (pltb.rows.length > 1) {
        var err1, rplc1, col, prm = '';
        err1 = prompt('Enter the 1st error text here:');
        rplc1 = prompt('Enter the 1st replacement text here:');
        col = prompt('Enter the column number here:');
        col = col - 1;
        for (i = 1; i < pltb.rows.length; i++) {
            prm = pltb.rows[i].childNodes[col].textContent;
            var req = db.transaction([tblName], 'readwrite').objectStore([tblName]).get(parseInt(pltb.rows[i].childNodes[0].id));
            if (prm === err1) {
                req.onsuccess = function (e) {
                    var ply = e.target.result;
                    ply[pltb.rows[0].childNodes[col].textContent] = rplc1;
                    var upd = db.transaction([tblName], 'readwrite').objectStore([tblName]).put(ply);
                };
            }
            else if (prm.search(err1) >= 0) {
                req.onsuccess = function (e) {
                    var ply = e.target.result;
                    prm = ply[pltb.rows[0].childNodes[col].textContent];
                    prm = prm.replace(err1, rplc1);
                    prm = prm.replace(/&amp;/, '&');
                    ply[pltb.rows[0].childNodes[col].textContent] = prm;
                    var upd = db.transaction([tblName], 'readwrite').objectStore([tblName]).put(ply);
                };
            }
        }
        setTimeout(function () {
            returnResults();
        }, 1000);
    }
    else {
        alert('Please ensure that there are some records displayed!');
    }
}
function resetPlayed() {
    pltb = document.getElementById('tblResults');
    if (pltb.rows.length > 1) {
        if (confirm('Are you sure you want to reset all records currently displayed to Unplayed')) {
            if (!request) {
                dbOpen();
            }
            db = request.result;
            for (i = 1; i < pltb.rows.length; i++) {
                var req = db.transaction([tblName], 'readwrite').objectStore([tblName]).get(parseInt(pltb.rows[i].childNodes[0].id));
                req.onsuccess = function (e) {
                    var ply = e.target.result;
                    ply.Played = 0;
                    var upd = db.transaction([tblName], 'readwrite').objectStore([tblName]).put(ply);
                };
            }
            setTimeout(function () {
                returnResults();
            }, 1000);
        }
    }
    else {
        alert('Please ensure that there are some records displayed!');
    }
}
function returnResults() {
    tblName = document.getElementById('strNm').value;
    var arrTrack = new Array, vlTrack = document.getElementById('Whr1').value, arrFPath = new Array, vlFPath = document.getElementById('Whr2').value, arrPlayed = new Array, vlPLayed = document.getElementById('Whr3').value;
    var arrFav = new Array, vlFav = document.getElementById('Whr4').value, arrFUrl = new Array, vlFurl = document.getElementById('Whr5').value, pos, vl, trchk, fpchk, pldchk, favchk, frlchk, trckvl, fpvl, pldvl, favvl, frlvl;
    if (vlTrack > '') {
        vl = vlTrack.toLowerCase();
        pos = 0;
        vl = vl.replace(/, /g, ',');
        if (vl > '' && vl.substring(vl.length - 2, vl.length - 1) !== ',') {
            vl = vl + ',';
        }
        for (i = 0; i < vl.length - vl.replace(/,/g, '').length; i++) {
            arrTrack[i] = vl.substring(pos, vl.substring(pos, vl.length).search(',') + pos);
            pos = arrTrack[i].length + 1 + pos;
        }
    }
    if (vlFPath > '') {
        vl = vlFPath.toLowerCase();
        pos = 0;
        vl = vl.replace(/, /g, ',');
        if (vl > '' && vl.substring(vl.length - 2, vl.length - 1) !== ',') {
            vl = vl + ',';
        }
        for (i = 0; i < vl.length - vl.replace(/,/g, '').length; i++) {
            arrFPath[i] = vl.substring(pos, vl.substring(pos, vl.length).search(',') + pos);
            pos = arrFPath[i].length + 1 + pos;
        }
    }
    if (vlFav > '') {
        vl = vlFav.toLowerCase();
        pos = 0;
        vl = vl.replace(/, /g, ',');
        if (vl > '' && vl.substring(vl.length - 2, vl.length - 1) !== ',') {
            vl = vl + ',';
        }
        for (i = 0; i < vl.length - vl.replace(/,/g, '').length; i++) {
            arrFav[i] = vl.substring(pos, vl.substring(pos, vl.length).search(',') + pos);
            pos = arrFav[i].length + 1 + pos;
        }
    }
    if (vlFurl > '') {
        vl = vlFurl.toLowerCase();
        pos = 0;
        vl = vl.replace(/, /g, ',');
        if (vl > '' && vl.substring(vl.length - 2, vl.length - 1) !== ',') {
            vl = vl + ',';
        }
        for (i = 0; i < vl.length - vl.replace(/,/g, '').length; i++) {
            arrFUrl[i] = vl.substring(pos, vl.substring(pos, vl.length).search(',') + pos);
            pos = arrFUrl[i].length + 1 + pos;
        }
    }
    if (!request) {
        dbOpen();
    }
    var prmS = '', prmP, cols = '', rslt = '', rmn;
    document.getElementById('tblResults').innerHTML = '';
    if (document.getElementById('Col1').value > '') { cols = '<th>' + document.getElementById('Col1').value + '</th>'; }
    if (document.getElementById('Col2').value > '') { cols = cols + '<th>' + document.getElementById('Col2').value + '</th>'; }
    if (document.getElementById('Col3').value > '') { cols = cols + '<th>' + document.getElementById('Col3').value + '</th>'; }
    if (document.getElementById('Col4').value > '') { cols = cols + '<th>' + document.getElementById('Col4').value + '</th>'; }
    if (document.getElementById('Col5').value > '') { cols = cols + '<th>' + document.getElementById('Col5').value + '</th>'; }
    cols = cols + '<th width="10">Select</th>';
    db = request.result;
    if (document.getElementById('idxNm').value > '') {
        var arr = new Array;
        vl = document.getElementById('idxFlt').value;
        pos = 0;
        vl = vl.replace(/, /g, ',');
        if (vl > '' && vl.substring(vl.length -2,vl.length-1) !== ',') {
            vl = vl + ',';
        }
        for (i = 0; i < vl.length - vl.replace(/,/g, '').length; i++) {
            arr[i] = vl.substring(pos, vl.substring(pos, vl.length).search(',') + pos);
            pos = arr[i].length + 1 + pos;
            db.transaction(tblName).objectStore(tblName).index(document.getElementById('idxNm').value).openCursor(arr[i]).onsuccess = function (event) {
                var cursor = event.target.result;
                if (cursor) {
                    if (vlTrack > '') {
                        trchk = 0;
                        trckvl = cursor.value[document.getElementById('Col1').value].toLowerCase();
                        for (j = 0; j < arrTrack.length; j++) {
                            if (trckvl.search(arrTrack[j]) >= 0) {
                                trchk = 1;
                                break;
                            }
                        }
                    }
                    else {
                        trchk = 1;
                    }
                    document.getElementById('rws').value =trchk;
                    if (vlFPath > '') {
                        fpchk = 0;
                        fpvl = cursor.value[document.getElementById('Col2').value].toLowerCase();
                        for (j = 0; j < arrFPath.length; j++) {
                            if (fpvl.search(arrFPath[j]) >= 0) {
                                fpchk = 1;
                                break;
                            }
                        }
                    }
                    else {
                        fpchk = 1;
                    }
                    document.getElementById('rws').value = document.getElementById('rws').value + fpchk;
                    if (vlPLayed.length === 1) {
                        pldchk = 0;
                        if (parseInt(cursor.value[document.getElementById('Col3').value]) === parseInt(vlPLayed)) {
                            pldchk = 1;
                        }
                    }
                    else {
                        pldchk = 1;
                    }
                    document.getElementById('rws').value = document.getElementById('rws').value + pldchk;
                    if (vlFav > '') {
                        favchk = 0;
                        favvl = cursor.value[document.getElementById('Col4').value].toLowerCase();
                        for (j = 0; j < arrFav.length; j++) {
                            if (favvl.search(arrFav[j]) >= 0) {
                                favchk = 1;
                                break;
                            }
                        }
                    }
                    else {
                        favchk = 1;
                    }
                    document.getElementById('rws').value = document.getElementById('rws').value + favchk;
                    if (vlFurl > '') {
                        frlchk = 0;
                        frlvl = cursor.value[document.getElementById('Col5').value].toLowerCase();
                        for (j = 0; j < arrFUrl.length; j++) {
                            if (frlvl.search(arrFUrl[j]) >= 0) {
                                frlchk = 1;
                                break;
                            }
                        }
                    }
                    else {
                        frlchk = 1;
                    }
                    document.getElementById('rws').value = document.getElementById('rws').value + frlchk;
                    if (document.getElementById('rws').value === "11111") {
                        rslt = rslt + '<tr>';
                        prmP = '';
                        if (cursor.value[document.getElementById('Col5').value] > '') {
                            prmP = cursor.value[document.getElementById('Col2').value];
                            if (prmP > '') {
                                prmP = prmP.toString().replace(/\\/g, "\\\\");
                            }
                            prmS = cursor.value[document.getElementById('Col5').value];
                            prmS = prmS.toString().replace(/\\/g, "\\\\");
                            rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="' + [prmS] + '" fPath="' + [prmP] + '" plyd="' + cursor.value['Played'] + '" style="text-wrap:none; cursor:pointer;" onClick="TrackNmClick(id, textContent, &quot;' + [prmS] + '&quot;, &quot;' + [prmP] + '&quot;, &quot;' + cursor.value[document.getElementById('Col4').value] + '&quot;, &quot;' + [prmS] + '&quot;);" id="' + cursor.value.Id + '">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                        }
                        else if (cursor.value['FileURL'] > '' && cursor.value['FilePath'] > '') {
                            prmP = cursor.value['FilePath'];
                            if (prmP > '') {
                                prmP = prmP.toString().replace(/\\/g, "\\\\");
                            }
                            prmS = cursor.value['FileURL'];
                            prmS = prmS.toString().replace(/\\/g, "\\\\");
                            rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="' + [prmS] + '" fPath="' + [prmP] + '" plyd="' + cursor.value['Played'] + '" style="text-wrap:none; cursor:pointer;" onClick="TrackNmClick(id, textContent, &quot;' + [prmS] + '&quot;, &quot;' + [prmP] + '&quot;, &quot;' + cursor.value[document.getElementById('Col4').value] + '&quot;, &quot;' + [prmS] + '&quot;);" id="' + cursor.value.Id + '">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                        }
                        else if (cursor.value['FileURL'] > '') {
                            prmS = cursor.value['FileURL'];
                            prmS = prmS.toString().replace(/\\/g, "\\\\");
                            for (j = 1; j < prmS.length; j++){
                                if (prmS.substring(prmS.length - j, prmS.length + 1 - j) === "\\" || prmS.substring(prmS.length - j, prmS.length + 1 - j) === "/"){
                                    prmP = prmS.substring(0, prmS.length + 1 - j);
                                    break;
                                }
                            }
                            if (prmP > '') {
                                prmP = prmP.toString().replace(/\\/g, "\\\\");
                            }
                            rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="' + [prmS] + '" fPath="' + [prmP] + '" plyd="' + cursor.value['Played'] + '" style="text-wrap:none; cursor:pointer;" onClick="TrackNmClick(id, textContent, &quot;' + [prmS] + '&quot;, &quot;' + [prmP] + '&quot;, &quot;' + cursor.value[document.getElementById('Col4').value] + '&quot;, &quot;' + [prmS] + '&quot;);" id="' + cursor.value.Id + '">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                        }
                        else {
                            rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="" fPath="" plyd="" style="text-wrap:none; cursor:pointer;" onClick="document.getElementById(' + "'Val1'" + ').value=textContent;">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                        }
                        if (Col2.value > '') {
                            rslt = rslt + '<td style="word-wrap:break-word;" onClick="document.getElementById(' + "'Val2'" + ').value=textContent;document.getElementById(' + "'frmSrc'" + ').value=textContent;document.getElementById(' + "'frmView'" + ').src=textContent;">' + cursor.value[document.getElementById('Col2').value] + '</td>';
                        }
                        if (Col3.value > '') {
                            rslt = rslt + '<td onClick="document.getElementById(' + "'Val3'" + ').value=document.URL">' + cursor.value[document.getElementById('Col3').value] + '</td>';
                        }
                        if (Col4.value > '') {
                            rslt = rslt + '<td onClick="document.getElementById(' + "'Val4'" + ').value=textContent;">' + cursor.value[document.getElementById('Col4').value] + '</td>';
                        }
                        if (Col5.value > '') {
                            rslt = rslt + '<td style="word-wrap:break-word;" onClick="document.getElementById(' + "'Val5'" + ').value=textContent;">' + cursor.value[document.getElementById('Col5').value] + '</td>';
                        }
                        rslt = rslt + '<td width="10"><input type="checkbox" /></td></tr>';
                    }
                    cursor.continue();
                }
                else {
                    document.getElementById('tblResults').innerHTML = cols + rslt;
                    document.getElementById('rws').value = document.getElementById('tblResults').rows.length - 1;
                }
            };
        }
    }
    else {
        db.transaction(tblName).objectStore(tblName).openCursor().onsuccess = function (event) {
            var cursor = event.target.result;
            if (cursor) {
                if (vlTrack > '') {
                    trchk = 0;
                    trckvl = cursor.value[document.getElementById('Col1').value].toLowerCase();
                    for (i = 0; i < arrTrack.length; i++) {
                        if (trckvl.search(arrTrack[i]) >= 0) {
                            trchk = 1;
                            break;
                        }
                    }
                }
                else {
                    trchk = 1;
                }
                document.getElementById('rws').value = trchk;
                if (vlFPath > '') {
                    fpchk = 0;
                    fpvl = cursor.value[document.getElementById('Col2').value].toLowerCase();
                    for (i = 0; i < arrFPath.length; i++) {
                        if (fpvl.search(arrFPath[i]) >= 0) {
                            fpchk = 1;
                            break;
                        }
                    }
                }
                else {
                    fpchk = 1;
                }
                document.getElementById('rws').value = document.getElementById('rws').value + fpchk;
                if (vlPLayed.length === 1) {
                    pldchk = 0;
                    if (parseInt(cursor.value[document.getElementById('Col3').value]) === parseInt(vlPLayed)) {
                        pldchk = 1;
                    }
                }
                else {
                    pldchk = 1;
                }
                document.getElementById('rws').value = document.getElementById('rws').value + pldchk;
                if (vlFav > '') {
                    favchk = 0;
                    favvl = cursor.value[document.getElementById('Col4').value].toLowerCase();
                    for (i = 0; i < arrFav.length; i++) {
                        if (favvl.search(arrFav[i]) >= 0) {
                            favchk = 1;
                            break;
                        }
                    }
                }
                else {
                    favchk = 1;
                }
                document.getElementById('rws').value = document.getElementById('rws').value + favchk;
                if (vlFurl > '') {
                    frlchk = 0;
                    frlvl = cursor.value[document.getElementById('Col5').value].toLowerCase();
                    for (i = 0; i < arrFUrl.length; i++) {
                        if (frlvl.search(arrFUrl[i]) >= 0) {
                            frlchk = 1;
                            break;
                        }
                    }
                }
                else {
                    frlchk = 1;
                }
                document.getElementById('rws').value = document.getElementById('rws').value + frlchk;
                if (document.getElementById('rws').value === "11111") {
                    rslt = rslt + '<tr>';
                    prmP = '';
                    if (cursor.value[document.getElementById('Col5').value] > '') {
                        prmP = cursor.value[document.getElementById('Col2').value];
                        if (prmP > '') {
                            prmP = prmP.toString().replace(/\\/g, "\\\\");
                        }
                        prmS = cursor.value[document.getElementById('Col5').value];
                        prmS = prmS.toString().replace(/\\/g, "\\\\");
                        rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="' + [prmS] + '" fPath="' + [prmP] + '" plyd="' + cursor.value['Played'] + '" style="text-wrap:none; cursor:pointer;" onClick="TrackNmClick(id, textContent, &quot;' + [prmS] + '&quot;, &quot;' + [prmP] + '&quot;, &quot;' + cursor.value[document.getElementById('Col4').value] + '&quot;, &quot;' + [prmS] + '&quot;);" id="' + cursor.value.Id + '">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                    }
                    else if (cursor.value['FileURL'] > '' && cursor.value['FilePath'] > '') {
                        prmP = cursor.value['FilePath'];
                        if (prmP > '') {
                            prmP = prmP.toString().replace(/\\/g, "\\\\");
                        }
                        prmS = cursor.value['FileURL'];
                        prmS = prmS.toString().replace(/\\/g, "\\\\");
                        rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="' + [prmS] + '" fPath="' + [prmP] + '" plyd="' + cursor.value['Played'] + '" style="text-wrap:none; cursor:pointer;" onClick="TrackNmClick(id, textContent, &quot;' + [prmS] + '&quot;, &quot;' + [prmP] + '&quot;, &quot;' + cursor.value[document.getElementById('Col4').value] + '&quot;, &quot;' + [prmS] + '&quot;);" id="' + cursor.value.Id + '">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                    }
                    else if (cursor.value['FileURL'] > '') {
                        prmS = cursor.value['FileURL'];
                        prmS = prmS.toString().replace(/\\/g, "\\\\");
                        for (j = 1; j < prmS.length; j++) {
                            if (prmS.substring(prmS.length - j, prmS.length + 1 - j) === "\\" || prmS.substring(prmS.length - j, prmS.length + 1 - j) === "/") {
                                prmP = prmS.substring(0, prmS.length + 1 - j);
                                break;
                            }
                        }
                        if (prmP > '') {
                            prmP = prmP.toString().replace(/\\/g, "\\\\");
                        }
                        rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="' + [prmS] + '" fPath="' + [prmP] + '" plyd="' + cursor.value['Played'] + '" style="text-wrap:none; cursor:pointer;" onClick="TrackNmClick(id, textContent, &quot;' + [prmS] + '&quot;, &quot;' + [prmP] + '&quot;, &quot;' + cursor.value[document.getElementById('Col4').value] + '&quot;, &quot;' + [prmS] + '&quot;);" id="' + cursor.value.Id + '">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                    }
                    else {
                        rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="" fPath="" plyd="" style="text-wrap:none; cursor:pointer;" onClick="document.getElementById(' + "'Val1'" + ').value=textContent;">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                    }
                    if (Col2.value > '') {
                        rslt = rslt + '<td style="word-wrap:break-word;" onClick="document.getElementById(' + "'Val2'" + ').value=textContent;document.getElementById(' + "'frmSrc'" + ').value=textContent;document.getElementById(' + "'frmView'" + ').src=textContent;">' + cursor.value[document.getElementById('Col2').value] + '</td>';
                    }
                    if (Col3.value > '') {
                        rslt = rslt + '<td onClick="document.getElementById(' + "'Val3'" + ').value=document.URL">' + cursor.value[document.getElementById('Col3').value] + '</td>';
                    }
                    if (Col4.value > '') {
                        rslt = rslt + '<td onClick="document.getElementById(' + "'Val4'" + ').value=textContent;">' + cursor.value[document.getElementById('Col4').value] + '</td>';
                    }
                    if (Col5.value > '') {
                        rslt = rslt + '<td style="word-wrap:break-word;" onClick="document.getElementById(' + "'Val5'" + ').value=textContent;">' + cursor.value[document.getElementById('Col5').value] + '</td>';
                    }
                    rslt = rslt + '<td width="10"><input type="checkbox" /></td></tr>';
                }
                cursor.continue();
            }
            else {
                document.getElementById('tblResults').innerHTML = cols + rslt;
                document.getElementById('rws').value = document.getElementById('tblResults').rows.length - 1;
            }
        };
    }
    setTimeout(function () {
        if (document.getElementById('tblResults').rows.length === 1 && document.getElementById('idxNm').value > '') {
            vl = document.getElementById('idxFlt').value, pos = 0;
            vl = vl.replace(/, /g, ',');
            if (vl > '' && vl.substring(vl.length - 2, vl.length - 1) !== ',') {
                vl = vl + ',';
            }
            for (i = 0; i < vl.length - vl.replace(/,/g, '').length; i++) {
                arr[i] = vl.substring(pos, vl.substring(pos, vl.length).search(',') + pos);
                pos = arr[i].length + 1 + pos;
                rmn = arr[i].toString();
                db.transaction(tblName).objectStore(tblName).index(document.getElementById('idxNm').value).openCursor(IDBKeyRange.bound(rmn, rmn + '\uffff')).onsuccess = function (event) {
                    var cursor = event.target.result;
                    if (cursor) {
                        if (vlTrack > '') {
                            trchk = 0;
                            trckvl = cursor.value[document.getElementById('Col1').value].toLowerCase();
                            for (j = 0; j < arrTrack.length; j++) {
                                if (trckvl.search(arrTrack[j]) >= 0) {
                                    trchk = 1;
                                    break;
                                }
                            }
                        }
                        else {
                            trchk = 1;
                        }
                        document.getElementById('rws').value = trchk;
                        if (vlFPath > '') {
                            fpchk = 0;
                            fpvl = cursor.value[document.getElementById('Col2').value].toLowerCase();
                            for (j = 0; j < arrFPath.length; j++) {
                                if (fpvl.search(arrFPath[j]) >= 0) {
                                    fpchk = 1;
                                    break;
                                }
                            }
                        }
                        else {
                            fpchk = 1;
                        }
                        document.getElementById('rws').value = document.getElementById('rws').value + fpchk;
                        if (vlPLayed.length === 1) {
                            pldchk = 0;
                            if (parseInt(cursor.value[document.getElementById('Col3').value]) === parseInt(vlPLayed)) {
                                pldchk = 1;
                            }
                        }
                        else {
                            pldchk = 1;
                        }
                        document.getElementById('rws').value = document.getElementById('rws').value + pldchk;
                        if (vlFav > '') {
                            favchk = 0;
                            favvl = cursor.value[document.getElementById('Col4').value].toLowerCase();
                            for (j = 0; j < arrFav.length; j++) {
                                if (favvl.search(arrFav[j]) >= 0) {
                                    favchk = 1;
                                    break;
                                }
                            }
                        }
                        else {
                            favchk = 1;
                        }
                        document.getElementById('rws').value = document.getElementById('rws').value + favchk;
                        if (vlFurl > '') {
                            frlchk = 0;
                            frlvl = cursor.value[document.getElementById('Col5').value];
                            for (j = 0; j < arrFUrl.length; j++) {
                                if (frlvl.search(arrFUrl[j]) >= 0) {
                                    frlchk = 1;
                                    break;
                                }
                            }
                        }
                        else {
                            frlchk = 1;
                        }
                        document.getElementById('rws').value = document.getElementById('rws').value + frlchk;
                        if (document.getElementById('rws').value === "11111") {
                            rslt = rslt + '<tr>';
                            prmP = '';
                            if (cursor.value[document.getElementById('Col5').value] > '') {
                                prmP = cursor.value[document.getElementById('Col2').value];
                                if (prmP > '') {
                                    prmP = prmP.toString().replace(/\\/g, "\\\\");
                                }
                                prmS = cursor.value[document.getElementById('Col5').value];
                                prmS = prmS.toString().replace(/\\/g, "\\\\");
                                rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="' + [prmS] + '" fPath="' + [prmP] + '" plyd="' + cursor.value['Played'] + '" style="text-wrap:none; cursor:pointer;" onClick="TrackNmClick(id, textContent, &quot;' + [prmS] + '&quot;, &quot;' + [prmP] + '&quot;, &quot;' + cursor.value[document.getElementById('Col4').value] + '&quot;, &quot;' + [prmS] + '&quot;);" id="' + cursor.value.Id + '">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                            }
                            else if (cursor.value['FileURL'] > '' && cursor.value['FilePath'] > '') {
                                prmP = cursor.value['FilePath'];
                                if (prmP > '') {
                                    prmP = prmP.toString().replace(/\\/g, "\\\\");
                                }
                                prmS = cursor.value['FileURL'];
                                prmS = prmS.toString().replace(/\\/g, "\\\\");
                                rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="' + [prmS] + '" fPath="' + [prmP] + '" plyd="' + cursor.value['Played'] + '" style="text-wrap:none; cursor:pointer;" onClick="TrackNmClick(id, textContent, &quot;' + [prmS] + '&quot;, &quot;' + [prmP] + '&quot;, &quot;' + cursor.value[document.getElementById('Col4').value] + '&quot;, &quot;' + [prmS] + '&quot;);" id="' + cursor.value.Id + '">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                            }
                            else if (cursor.value['FileURL'] > '') {
                                prmS = cursor.value['FileURL'];
                                prmS = prmS.toString().replace(/\\/g, "\\\\");
                                for (j = 1; j < prmS.length; j++) {
                                    if (prmS.substring(prmS.length - j, prmS.length + 1 - j) === "\\" || prmS.substring(prmS.length - j, prmS.length + 1 - j) === "/") {
                                        prmP = prmS.substring(0, prmS.length + 1 - j);
                                        break;
                                    }
                                }
                                if (prmP > '') {
                                    prmP = prmP.toString().replace(/\\/g, "\\\\");
                                }
                                rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="' + [prmS] + '" fPath="' + [prmP] + '" plyd="' + cursor.value['Played'] + '" style="text-wrap:none; cursor:pointer;" onClick="TrackNmClick(id, textContent, &quot;' + [prmS] + '&quot;, &quot;' + [prmP] + '&quot;, &quot;' + cursor.value[document.getElementById('Col4').value] + '&quot;, &quot;' + [prmS] + '&quot;);" id="' + cursor.value.Id + '">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                            }
                            else {
                                rslt = rslt + '<td title="' + cursor.value[document.getElementById('Col1').value] + '" track="" fPath="" plyd="" style="text-wrap:none; cursor:pointer;" onClick="document.getElementById(' + "'Val1'" + ').value=textContent;">' + cursor.value[document.getElementById('Col1').value] + '</td>';
                            }
                            if (Col2.value > '') {
                                rslt = rslt + '<td style="word-wrap:break-word;" onClick="document.getElementById(' + "'Val2'" + ').value=textContent;document.getElementById(' + "'frmSrc'" + ').value=textContent;document.getElementById(' + "'frmView'" + ').src=textContent;">' + cursor.value[document.getElementById('Col2').value] + '</td>';
                            }
                            if (Col3.value > '') {
                                rslt = rslt + '<td onClick="document.getElementById(' + "'Val3'" + ').value=document.URL">' + cursor.value[document.getElementById('Col3').value] + '</td>';
                            }
                            if (Col4.value > '') {
                                rslt = rslt + '<td onClick="document.getElementById(' + "'Val4'" + ').value=textContent;">' + cursor.value[document.getElementById('Col4').value] + '</td>';
                            }
                            if (Col5.value > '') {
                                rslt = rslt + '<td style="word-wrap:break-word;" onClick="document.getElementById(' + "'Val5'" + ').value=textContent;">' + cursor.value[document.getElementById('Col5').value] + '</td>';
                            }
                            rslt = rslt + '<td width="10"><input type="checkbox" /></td></tr>';
                        }
                        cursor.continue();
                    }
                    else {
                        document.getElementById('tblResults').innerHTML = cols + rslt;
                        document.getElementById('rws').value = document.getElementById('tblResults').rows.length - 1;
                    }
                };
            }
        }
    }, 50);
    setTimeout(function () {
        SrtRsltRows();
    }, parseInt(document.getElementById('ldTm').value) * 1000);
}
function Rev(a) {
    n = parseInt(a) - 1;
    b = n;
    rn = b;
    if (n===0) {
        window.clearInterval(PlayLst);
        document.location.reload();
    }
}
function runScript() {
    scrpt = document.getElementById('txtView').value;
    scrpt = scrpt.replace(/\n/g, '');
    scrpt = 'JavaScript:' + scrpt;
    document.getElementById('frmView').contentWindow.document.location = scrpt;
}
function selfUpload() {
    nw = 0;
    upldList = window.setInterval("uploadRecord(1)", 10);
}
function srchTbl(search) {
    var cells = document.querySelectorAll('#tblResults tr'), opt = '<option></option>';
    for (var i = 0; i < cells.length; i++) {
        var txt = ' ' + cells[i].textContent.toLowerCase();
        if (txt.search(search.toLowerCase()) > 0) {
            opt = opt + '<option value="' + i + '">' + cells[i].childNodes[0].textContent + '</option>';
        }
    }
    document.getElementById('tblSel').innerHTML = opt;
}
function SrtRsltRows() {
    var rows;
    while (!document.getElementById("tblResults").rows) { }
    if (document.getElementById("tblResults")) {
        rows = document.getElementById("tblResults").rows;
        var tblRw = rows[0].innerHTML, d, rlen = rows.length, arr = new Array(), cells, clen, asc = 1;
        for (i = 1; i < rlen; i++) {
            arr[i - 1] = '<tr>' + rows[i].innerHTML + '</tr>';
        }
        arr.sort(function (a, b) {
            return (a === b) ? 0 : ((a > b) ? asc : -1 * asc);
        });
        for (i = 0; i < rlen - 1; i++) {
            tblRw = tblRw + arr[i];
        }
        document.getElementById("tblResults").innerHTML = tblRw;
    }
}
function startPlay1() {
    tblName = document.getElementById('strNm').value;
    if (document.getElementById("FlSt").value !== "3") {
        if (!request) {
            dbOpen();
        }
        db = request.result;
        var req = db.transaction([tblName], 'readwrite').objectStore([tblName]).get(parseInt(rwId));
        req.onsuccess = function (e) {
            var ply = e.target.result;
            ply.Played = 1;
            var upd = db.transaction([tblName], 'readwrite').objectStore([tblName]).put(ply);
            upd.onsuccess = function (e) {
                returnResults();
            };
        };
    }
    PlayLst = window.setInterval("Play_Track()", 1000);
    n = rn;
    m = 0;
    //Play_Track();
}
function tableToDocument() {
    document.all[0].innerHTML = document.getElementById('tblResults').outerHTML;
}
function tblSelect(selRw) {
    if (document.getElementById('tblSel').selectedIndex !== 0) {
        var trRw = document.getElementById('tblResults').rows[selRw];
        trRw.scrollIntoView(true);
        document.getElementById('tblSel').selectedIndex = 0;
    }
}
function Track_Name(b) {
    m = 0;
    pos = 1;
    flPth = pltb.rows[b].childNodes[0].getAttribute('fPath');
    rwId = pltb.rows[b].childNodes[0].id;
    document.getElementById('Val1').value = pltb.rows[b].childNodes[0].textContent;
    Extn(pltb.rows[b].childNodes[0].getAttribute('track'));
}
function TrackNmClick(trId, trNm, prmS, flPth, fvr, furl) {
    prmS = prmS.toString().replace(/\\/g,"\\\\");
    document.getElementById('Val1').value = trNm;
    document.getElementById('Val4').value = fvr;
    document.getElementById('Val5').value = furl;
    document.getElementById('frmView').height = 580;
    document.getElementById('frmView').style.visibility = 'visible';
    document.getElementById('frmView').scrollIntoView(true);
    rwId = trId;
    if (document.getElementById('Val2').value==='') {
        ze = prmS;
        zo = prmS;
        if (fvr.substring(0,5) ==='Radio') {
            radioPlay();
        }
        else {
            startPlay();
        }
    }
    else {
        flPth = flPth.replace(/<span>/g, '').replace(/<\/span>/g, '');
        prmSNew = prmS.replace(flPth, document.getElementById('Val2').value);
        ze = prmSNew;
        zo = prmS;
        if (fvr.substring(0, 5) === 'Radio') {
            radioPlay();
        }
        else {
            startPlay();
        }
    }
}
function UpdTbl() {
    m = 0;
    pos = 1;
    pltb = document.getElementById('tblResults');
}
function uploadTable() {
    nw = 0;
    upldList = window.setInterval("uploadRecord(0)", 10);
}
function viewEditFrame() {
    if (document.getElementById('txtView').style.visibility === 'hidden') {
        document.getElementById('txtView').rows = 9;
        document.getElementById('txtView').style.visibility = 'visible';
    }
    else {
        document.getElementById('txtView').style.visibility = 'hidden';
        document.getElementById('txtView').rows = 1;
    }
}
function viewFrame() {
    if (document.getElementById('frmView').style.visibility === 'hidden') {
        document.getElementById('frmView').height = 580;
        document.getElementById('frmView').style.visibility = 'visible';
    }
    else {
        document.getElementById('frmView').height = 0;
        document.getElementById('frmView').style.visibility = 'hidden';
    }
}
