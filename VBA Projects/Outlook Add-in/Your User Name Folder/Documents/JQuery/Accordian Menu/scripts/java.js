var rowIndex, colIndex, $td, $title = document.title;

function appPrefix() {
    window.history.back();
    var txt = prompt("Please give prefix text: ");
    if (txt != "") {
        $td.innerText = txt + $td.innerText;
    }
}
function appPrefixCol() {
    window.history.back();
    var txt = prompt("Please give prefix text: ");
    if (txt != "") {
        $('#tblDtl tr').each(function () {
            $(this).find("td").each(function ($idx) {
                if ($idx == colIndex) {
                    this.innerText = txt + this.innerText;
                }
            })
        });
    }
}
function appSuffix() {
    window.history.back();
    var txt = prompt("Please give suffix text: ");
    if (txt != "") {
        $td.innerText = $td.innerText + txt;
    }
}
function appSuffixCol() {
    window.history.back();
    var txt = prompt("Please give suffix text: ");
    if (txt != "") {
        $('#tblDtl tr').each(function () {
            $(this).find("td").each(function ($idx) {
                if ($idx == colIndex) {
                    this.innerText = this.innerText + txt;
                }
            })
        });
    }
}
function clrHighLt() {
    window.history.back();
    var td = document.getElementsByTagName("td"), i;
    for (i = 0; i < td.length; i++) {
        if (td[i].style.backgroundColor != "") {
            td[i].style.backgroundColor = "";
        }
    }
}
function docRefresh() {
    document.location.reload();
}
function getHighLt() {
    var arr = "", td = document.getElementsByTagName("td"), i;
    for (i = 0; i < td.length; i++) {
        if (td[i].style.backgroundColor == "rgb(253, 233, 217)") {
            if (arr != "") {
                arr = arr + ",";
            }
            arr = arr + td[i].innerText.trim();
        }
    }
    document.title = "0,0,0," + arr;
    window.setTimeout(function () { document.title = $title; }, 1000);
    return arr;
}
function getHighLtArr() {
    var arr = "";
    $('#tblDtl tr').each(function ($rdx) {
        $(this).find("td").each(function ($idx) {
            if (this.style.backgroundColor == "rgb(253, 233, 217)") {
                if (arr != "") {
                    arr = arr + ",";
                }
                arr = arr + $rdx + "|" + $idx + "|" + this.innerText.toString().trim();
            }
        })
    });
    return arr;
}
function getInner(idx) {
    document.location = "#popupNested";
    if (idx.substring(4, idx.length) == 9 || idx.substring(4, idx.length) == 11) {
        document.title = idx + ',' + $td.innerText;
    } else {
        document.title = idx + ',' + getHighLt();
    }
}
function getRow(idx) {
    document.location = "#popupNested";
    document.title = idx + ',' + rowIndex;
}
function getSel() {
    window.history.back();
    if (window.getSelection || document.getSelection) {
        var oSel = (window.getSelection ? window : document).getSelection();
    }
    return oSel;
}
function highLt(idx) {
    if (document.getElementById("tblDtl").rows[idx].style.backgroundColor == "rgb(253, 233, 217)") {
        document.getElementById("tblDtl").rows[idx].style.backgroundColor = "RGB(253, 233, 217)";
    } else {
        document.getElementById("tblDtl").rows[idx].style.backgroundColor = "RGB(255, 255, 255)";
    }
}
function keepSel() {
    var oSel = getSel();
    if (oSel) {
        $td.innerText = $td.innerText.toString().replace($td.innerText, oSel);
    }
}
function keepSelCol() {
    var oSel = getSel();
    var $oSel = oSel.toString().trim();
    if (oSel) {
        $('#tblDtl tr').each(function () {
            $(this).find("td").each(function ($idx) {
                if ($idx == colIndex && this.innerText.indexOf($oSel) != -1) {
                    this.innerText = this.innerText.toString().replace(this.innerText, $oSel);
                }
            })
        });
    }
}
function remSel() {
    var oSel = getSel();
    if (oSel) {
        $td.innerText = $td.innerText.toString().replace(oSel, '');
    }
}
function remSelCol() {
    var oSel = getSel();
    var $oSel = oSel.toString().trim();
    if (oSel) {
        $('#tblDtl tr').each(function () {
            $(this).find("td").each(function ($idx) {
                if ($idx == colIndex) {
                    this.innerText = this.innerText.toString().replace($oSel, '');
                }
            })
        });
    }
}
function rplSel() {
    var oSel = getSel();
    var $oSel = oSel.toString().trim();
    if (oSel) {
        var rpl = prompt("What would you like to replace the selection with", $oSel);
        if (rpl != "") {
            oSel.getRangeAt(0).deleteContents();
            oSel.getRangeAt(0).insertNode(document.createTextNode(rpl));
        }
    }
}
function rplTxt() {
    var oSel = getSel();
    if (oSel) {
        var rpl = prompt("What would you like to replace the selection with", oSel);
        if (rpl != "") {
            $td.innerText = $td.innerText.toString().split(oSel).join(rpl);
        }
    }
}
function rplTxtCol() {
    var oSel = getSel();
    var $oSel = oSel.toString().trim();
    if (oSel) {
        var rpl = prompt("What would you like to replace the selection with", oSel);
        if (rpl != "") {
            $('#tblDtl tr').each(function () {
                $(this).find("td").each(function ($idx) {
                    if ($idx == colIndex && this.innerText.indexOf($oSel) != -1) {
                        this.innerText = this.innerText.toString().split($oSel).join(rpl);
                    }
                })
            });
        }
    }
}
function runScript(iStr) {
    window.history.back();
    document.title = iStr;
    window.setTimeout(function () {document.title = $title;}, 1000);
}
function setIndex(td, rIndex, cIndex) {
    $td = td;
    rowIndex = rIndex;
    colIndex = cIndex;
}
