Function DrvPthCh
	if DrvPth.value = ".." then 
		if instrRev(frmSrc.value,"\",len(frmSrc.value)-1) > 0 then
			frmSrc.value = left(frmSrc.value,instrRev(frmSrc.value,"\",len(frmSrc.value)-1))
		elseif instrRev(frmSrc.value,"/",len(frmSrc.value)-1) > 0 then
			frmSrc.value = left(frmSrc.value,instrRev(frmSrc.value,"/",len(frmSrc.value)-1))
		end if
	else		
		frmSrc.value = DrvPth.value
	end if
    FlSch()
End Function

Function FlSch
    on error resume next
    FsLd
	if Left(frmSrc.value,2)="./" then
		ft = "/"
		Set f = fs.GetFolder(root & right(frmSrc.value,len(frmSrc.value)-2))
	else
		ft = "\"
		Set f = fs.GetFolder(frmSrc.value)
	end if
	Set sf = f.SubFolders
	Set fc = f.Files
	s = frmSrc.value
    tmpFl = ""
    tmpFl1 = ""
    tmpFl2 = ""
    tmpFl3 = ""
    ay = "<OPTION value='..'>..</OPTION>"
    ay = ay & "<OPTION value=""" & frmSrc.value & """ selected>" & frmSrc.value & "</OPTION>"
    Val2.value = frmSrc.value
    For Each f1 In sf
		ay = ay & "<OPTION value=""" & s & f1.Name & ft & """>" & s & f1.Name & ft & "</OPTION>"
	next
    FlSch1 s,fc
'    For Each f1 In sf
'		SbFlSch(s & f1.Name & ft)
'	next
    DrvPth.outerHTML = "<SELECT id=DrvPth name=DrvPth onchange='DrvPthCh();'>" & ay & "</SELECT>"
'    document.GetElementById("Dta").outerHTML = "<table id='Dta'>" & tmpFl1 & tmpFl2 & tmpFl3 & tmpFl &"</table>"
    document.getElementById("frmView").style.visibility = "visible"
    document.getElementById("frmView").height = 500
    document.getElementById("frmView").contentDocument.write("<table id=Ply>" & tmpFl1 & tmpFl2 & tmpFl3 & tmpFl & "</table>")
    document.getElementById("frmView").contentDocument.close()
'    SrtRsltRows
End Function 

Function FlSch1(fa,fb)
    i = 0
    For Each fd In fb
		i = i + 1
		if i = 2000 then 
			TFl tmpFl,1
			tmpFl = ""
		elseif i = 4000 then
			TFl tmpFl,2
			tmpFl = ""
		elseif i = 6000 then
			TFl tmpFl,3
			tmpFl = ""
		end if
    	if left(fd.Name,len("AlbumArt")) <> "AlbumArt" and fd.Name <> "desktop.ini" and fd.Name <> "Folder.jpg" then
            tmpFl = tmpFl & "<tr><td>" & fd.Name & "</td></tr>"
        end if
    Next
End Function

Function FsLd
	Set fs = CreateObject("Scripting.FileSystemObject")
End Function

Function GoSbCl
	set WshShell = CreateObject("WScript.Shell")
	WshShell.Run "%windir%\explorer " & chr(34) & DrvPth.Value & chr(34)
End Function

Function SbFlSch(a)
	Set f4 = fs.GetFolder(a)
	Set sf1 = f4.SubFolders
	Set fc1 = f4.Files
	For Each f5 In sf1
		ay = ay & "<OPTION value='" & a & f5.Name & ft & "'>" & a & f5.Name & ft & "</OPTION>"
		SbFlSch1(a & f5.Name & ft)
	next
	FlSch1 a,fc1
End Function

Function SbFlSch1(aa)
	Set f6 = fs.GetFolder(aa)
	Set sf2 = f6.SubFolders
	Set fc2 = f6.Files
	For Each f7 In sf2
		ay = ay & "<OPTION value='" & aa & f7.Name & ft & "'>" & aa & f7.Name & ft & "</OPTION>"
		SbFlSch1(aa & f7.Name & ft)
	next
	FlSch1 aa,fc2
End Function

Function SbFlSch2(ab)
	Set f8 = fs.GetFolder(ab)
	Set sf3 = f8.SubFolders
	Set fc3 = f8.Files
	For Each f9 In sf3
		ay = ay & "<OPTION value='" & ab & f9.Name & ft & "'>" & ab & f9.Name & ft & "</OPTION>"
	next
	FlSch1 ab,fc3
End Function