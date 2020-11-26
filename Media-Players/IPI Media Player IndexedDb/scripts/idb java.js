function Extn(a) {
    window.clearInterval(PlayLst);
    if (document.getElementById('Val2').value === '') {
        ze = a;
        zo = a;
        //startPlay();
    }
    else {
        flPth = flPth.replace(/<span>/g, '').replace(/<\/span>/g, '');
        prmS = a.replace(flPth, document.getElementById('Val2').value);
        ze = prmS;
        zo = a;
        //startPlay();
    }
    document.getElementById("frmView").contentDocument.getElementById('Quicky').src = ze;
    document.getElementById("frmView").contentDocument.getElementById('Quicky').play();
    startPlay1();
}
function frameSrcCh(vl) {
    document.getElementById('frmSrc').value = prompt('Frame Source', vl);
    document.getElementById('frmView').height = 650;
    document.getElementById('frmView').style.visibility = 'visible';
    document.getElementById('frmView').src = vl;
}
function insertValues() {
    tblName = document.getElementById('strNm').value;
    var mediaData = [{
        [document.getElementById('Col1').value]: document.getElementById('Val1').value,
        [document.getElementById('Col2').value]: document.getElementById('Val2').value,
        [document.getElementById('Col3').value]: document.getElementById('Val3').value,
        [document.getElementById('Col4').value]: document.getElementById('Val4').value,
        [document.getElementById('Col5').value]: document.getElementById('Val5').value
    }];
    if (!request) {
        dbOpen();
    }
    db = request.result;
    var mediaObjectStore = db.transaction([tblName], 'readwrite').objectStore([tblName]);
    var ins = mediaObjectStore.add(mediaData[0]);
    ins.onsuccess = function (e) {
        document.getElementById('Val1').value = '';
        document.getElementById('Val2').value = '';
        document.getElementById('Val3').value = '';
        document.getElementById('Val4').value = '';
        document.getElementById('Val5').value = '';
        returnResults();
    };
    ins.onerror = function (e) {
        alert('There was the following error ' + e.target.result);
    };
}
function Play_Track() {
    m = m + 1;
    if (m > document.getElementById("PTm").value) {
        if (parseInt(document.getElementById("frmView").contentWindow.document.getElementById("Quicky").currentTime) === parseInt(document.getElementById("frmView").contentWindow.document.getElementById("Quicky").duration)) { pos = 0; }
        if (document.getElementById("frmView").contentWindow.document.getElementById("Quicky").currentTime === 0) { pos = 0; }
    }
    if (m > document.getElementById("PTm").value && pos === 0) {
        UpdTbl();
        if (document.getElementById("FlSt").value === "1") {
            Incr(rn);
        }
        if (document.getElementById("FlSt").value === "2") {
            window.clearInterval(PlayLst);
            document.location.reload();
        }
        if (document.getElementById("FlSt").value === "4") {
            Rev(rn);
        }
        if (document.getElementById("FlSt").value === "5") {
            Rdm(rn);
        }
        if (document.getElementById("FlSt").value === "6") {
            window.clearInterval(PlayLst);
            document.location.reload();
        }
        if (parseInt(document.getElementById('tblResults').rows[rn].childNodes[0].getAttribute('plyd')) !== 1) {
            Track_Name(rn);
        }
    }
}
function radioPlay() {
    if (ze.substring(ze.length - 1, ze.length) === '/' || ze.substring(ze.length - 4, ze.length) === 'aspx') {
        document.getElementById("frmView").src = '';
        setTimeout(function () {
            document.getElementById("frmView").contentDocument.location = ze;
        }, 250);
    }
    else {
       // document.getElementById("frmView").src = './';
        document.getElementById("frmView").contentDocument.all[0].innerHTML = '<video controls="true" x-webkit-airplay="allow" autoplay="true" id="Quicky" width="100%" height="100%" src="' + ze + '" />';
        setTimeout(function () { document.getElementById("frmView").contentDocument.getElementById('Quicky').play(); }, 5000);
        document.getElementById("frmView").height = 650;
    }
}
function SrcUpd() {
    var ml = 0;
    document.getElementById("frmView").src = "";
    a = document.getElementById("txtView").value;
    if ((a.search('/') < 0 && a.search('\\\\') < 0) || a.search('<table') >= 0) {
        if (a.search('<table') < 0) {
            document.getElementById("frmView").contentDocument.write("<table id=Ply style='font - size: 0.7em; '><tr><td>" + a.replace(/\n/g, "</td></tr><tr><td>") + "</tr></table>");
        }
        else {
            document.getElementById("frmView").contentDocument.write(a);
        }
        document.getElementById("frmView").contentDocument.close();
    }
    else {
        document.getElementById("frmView").contentWindow.document.write(a);
        document.getElementById("frmView").contentDocument.close();
        a = document.getElementById("frmView").contentDocument.firstChild.childNodes[1].innerHTML;
        if (a.substring(0, 3) !== "ftp" && document.getElementById('frmSrc').value.substring(0, 4) !== 'file') { a = "http://" + a; }
        a = a.replace("///", "<tr><td>");
        a = a.replace(/ftp:/, "<tr><td>ftp:");
        a = a.replace(/http:/, "<tr><td>http:");
        a = a.replace(" - /", "/");
        a = a.replace(/File:/g, "<tr><td>");
        a = a.replace(/ AM /g, "<tr><td>");
        a = a.replace(/ PM /g, "<tr><td>");
        a = a.replace("[To Parent Directory]", "");
        a = a.replace("Up to higher level directory", "");
        a = a.replace("Name 	Size 	Last Modified", "");
        if (a.search("Kingston Technology") > 0) {
            ml = 1
            a = a.replace("Kingston Technology", "");
            a = a.replace("/ SD_Card1/ Album", "");
            a = a.replace("Settings", "");
            a = a.replace("Back to Devices", "");
        }
        for (i = 1; i < a.length; i++) {
            a = a.replace("\n", "</td></tr><tr><td>");
        }
        document.getElementById("frmView").contentDocument.write("<table id=Ply>" + a + "</table>");
        document.getElementById("frmView").contentDocument.close();
        document.getElementById("frmView").contentDocument.write("<table id=Ply>" + document.getElementById("frmView").contentDocument.childNodes[0].childNodes[1].childNodes[document.getElementById("frmView").contentDocument.childNodes[0].childNodes[1].childNodes.length - 1].innerHTML + "</table>");
        document.getElementById("frmView").contentDocument.close();
        pltb = document.getElementById("frmView").contentDocument.getElementById("Ply");
        a = "<table id=Ply style='font-size:0.7em;'>";
        if (ml == 0) {
            if (pltb.rows[0].textContent > "") {
                d = pltb.rows[0].textContent;
                if (d.substring(0, 5) != "http:" || d.substring(0, 4) != "ftp:") {
                    //
                }
                else {
                    for (i = 1; i < d.length; i++) {
                        d = d.replace("/", "\\");
                    }
                }
            }
            if (d.substring(d.length - 1, d.length) != "\\" && d.substring(d.length - 1, d.length) != "/") {
                if (d.substring(0, 5) != "http:" || d.substring(0, 4) != "ftp:") {
                    d = "file:///" + d + "/";
                }
                else {
                    d = d + "\\";
                }
            }
            for (i = 1; i < pltb.rows.length; i++) {
                if (pltb.rows[i].innerText > "") {
                    e = pltb.rows[i].innerText;
                    if (e.search(",") > 0 && document.getElementById('frmSrc').value.substring(0, 4) === 'file') {
                        g = 0;
                    }
                    else {
                        g = 1;
                    }
                    if ((e.search(":") < 0 || e.search(",") > 0 || document.getElementById('frmSrc').value.substring(0, 4) !== 'file') && e.search("AlbumArt") < 0 && e.search("desktop.ini") < 0 && e.search("Folder.jpg") < 0 && e.search("javascript") < 0 && e.search('.config') < 0) {
                        for (j = 0; j < e.length; j++) {
                            if (e.substring(j, j + 1) === "." && e.length > j + 4) {
                                e = e.substring(g, j + 4);
                            }
                        }
                        if (pltb.rows[0].innerText.substring(0, 5) === "http:") {
                            //e = e.substring(e.search(":") + 1, e.length);
                            //e = e.substring(e.search(" ") + 1, e.length);
                            //e = e.substring(e.search(" ") + 1, e.length);
                            e = e.split('/')[e.split('/').length - 1]
                        }
                        f = d;
                        a = a + '<tr>';
                        a = a + '<td>' + e + '</td></tr>';
                    }
                }
            }
        }
        else {
            for (i = 3; i < pltb.rows.length; i++) {
                e = pltb.rows[i].innerText;
                if (e.search(" MB") < 0 && e.search(" KB") < 0 && e.search("AlbumArt") < 0 && e.search("desktop.ini") < 0 && e.search("Folder.jpg") < 0 && e.search("javascript") < 0 && e.search('.config') < 0) {
                    a = a + '<tr><td>' + e + '</td></tr>';
                }
            }
        }
        a = a + "</table>";
        document.getElementById("frmView").contentWindow.document.write(a);
        document.getElementById("frmView").contentDocument.close();
    }
}
function startPlay() {
    pltb = document.getElementById("tblResults");
    for (i = 1; i < pltb.rows.length; i++) {
        if (rwId===pltb.rows[i].childNodes[0].id) {
            rn = i;
            i = pltb.rows.length - 1;
        }
    }
    if (document.getElementById('Val4').value === 'YouTube') {
        document.getElementById('frmView').src = ze;
        setTimeout(function () { document.getElementById("frmView").contentDocument.getElementsByTagName('video')[0].id = 'Quicky'; }, 3000);
        setTimeout(function () { document.getElementById("frmView").contentDocument.getElementsByTagName('video')[0].autoplay = false; }, 4000);
    }
    //else if (ze.toString().search("\\\\") > 0) {
    //    document.getElementById('frmView').src = ze;
    //    document.getElementById('frmView').contentDocument.location = ze;
    //    setTimeout(function () { document.getElementById("frmView").contentDocument.getElementsByTagName('video')[0].id = 'Quicky'; }, 3000);
    //    setTimeout(function () { document.getElementById("frmView").contentDocument.getElementsByTagName('video')[0].autoplay = false; }, 4000);
    //}
    else {
        document.getElementById("frmView").contentDocument.all[0].innerHTML = '<video controls="true" x-webkit-airplay="allow" autoplay="true" id="Quicky" width="100%" height="100%" src="' + ze + '" />';
    }
    setTimeout(function () { document.getElementById("frmView").contentDocument.getElementById('Quicky').play(); }, 5000);
    if (document.getElementById('Val4').value !== 'YouTube') {
        startPlay1();
    }
}
function uploadRecord(tp) {
    var vl1 = '', vl2 = '', vl3 = '', vl4 = '', vl5 = '';
    tblName = document.getElementById('strNm').value;
    if (!request) {
        dbOpen();
    }
    if (tp === 0) {
        var tsv, fsrc;
        fsrc = document.getElementById("frmView").src;
        if (fsrc.search("file://") >= 0 && fsrc.search("iDb.txt") >= 0) {
            tsv = document.getElementById("frmView").contentDocument.childNodes[1].innerHTML;
            tsv = tsv.split('pre-wrap">')[1];
            tsv = tsv.split("</pre>")[0];
            tsv = tsv.replace(/\t/g, "</td><td>");
            tsv = tsv.replace(/\n/g, "</td></tr><tr><td>");
            tsv = '<table id="Ply"><tr><td>' + tsv + '</td></tr></table>';
            document.getElementById("frmView").src = "";
            document.getElementById("frmView").contentDocument.write(tsv);
            document.getElementById("frmView").contentDocument.close();
        }
        pltb = document.getElementById("frmView").contentDocument.getElementById("Ply");
        if (pltb.rows[0].cells.length === 1) {
            vl2 = document.getElementById("Val2").value;
            if (vl2.length > 0) {
                if (vl2.substring(vl2.length - 1, vl2.length) !== '/' && vl2.substring(vl2.length - 1, vl2.length) !== '\\') {
                    vl2 = vl2 + '/';
                }
            }
            else {
                if (nw < pltb.rows.length) {
                    if (pltb.rows[nw].cells[0].title > '') {
                        vl2 = pltb.rows[nw].cells[0].title;
                    }
                }
            }
            vl3 = 0;
            vl4 = document.getElementById("Val4").value;
        }
    }
    else if (tp === 1) {
        pltb = document.getElementById('frmView').contentDocument.getElementById('tblResults');
    }
    if (nw < pltb.rows.length) {
        if (tp === 0 && pltb.rows[0].cells.length === 1) {
            prm = pltb.rows[nw].textContent;
        }
        else {
            prm = pltb.rows[nw].childNodes[0].textContent;
        }
        if (prm !== '--' && prm > '' && (tp === 0 && pltb.rows[0].cells.length === 1 || ((tp === 0 || tp === 1) && nw > 0))) {
            if (tp === 0 && pltb.rows[0].cells.length === 1) {
                vl1 = prm;
                vl1 = vl1.replace(/\n/g, "");
                vl5 = vl2 + vl1;
                prm = [{
                    [document.getElementById("Col1").value] : vl1,
                    [document.getElementById("Col2").value] : vl2,
                    [document.getElementById("Col3").value] : vl3,
                    [document.getElementById("Col4").value] : vl4,
                    [document.getElementById("Col5").value] : vl5
                    }];
            }
            if (tp === 1 || (tp === 0 && pltb.rows[0].cells.length === 5)) {
                vl1 = pltb.rows[nw].childNodes[0].textContent;
                vl2 = pltb.rows[nw].childNodes[1].textContent;
                vl3 = pltb.rows[nw].childNodes[2].textContent;
                vl4 = pltb.rows[nw].childNodes[3].textContent;
                vl5 = pltb.rows[nw].childNodes[4].textContent;
                vl1 = vl1.replace(/\n/g, "");
                prm = [{
                    [pltb.rows[0].childNodes[0].textContent] : vl1, 
                    [pltb.rows[0].childNodes[1].textContent] : vl2, 
                    [pltb.rows[0].childNodes[2].textContent] : vl3, 
                    [pltb.rows[0].childNodes[3].textContent] : vl4, 
                    [pltb.rows[0].childNodes[4].textContent] : vl5
                }];
            }
            cs = vl5.replace(/" "/g, "");
            cs = cs.substring(cs.length - 3, cs.length);
            cs = cs.toUpperCase();
            if (((((document.getElementById("SvTp").value === "1"  || (
                document.getElementById("SvTp").value === "2" && (cs === "MP3" || cs === "CDA" || cs === "WMA" ||
                cs === "WAV" || cs === "MPA")) || (document.getElementById("SvTp").value === "3" && (cs === "MPG" ||
                cs === "PNG" || cs === "JPG" || cs === "PCX" || cs === "GIF" || cs === "BMP")) || (
                document.getElementById("SvTp").value === "4" && (cs === "AVI" || cs === "PEG" || cs === "DVX" ||
                cs === "DVD" || cs === "FLV" || cs === "MPE" || cs === "MPV" || cs === "ASF" || cs === "MP4")) ||
                (document.getElementById("SvTp").value === cs)))))) {
                document.getElementById('Val1').value = nw + 1 + ' of ' + pltb.rows.length + ' loaded!';
                db = request.result;
                var mediaObjectStore = db.transaction([tblName], 'readwrite').objectStore([tblName]);
                mediaObjectStore.add(prm[0]);
            }
        }
    }
    else {
        window.clearInterval(upldList);
        returnResults();
    }
    nw = nw + 1;
}
