Function RequisitionAction(obj)
	dim arrTemp
	if obj.selectedIndex <> 0 then
		arrTemp = split(obj.value,"?")
		document.formname.mrs.value = arrTemp(0)
		document.formname.sAct.value = arrTemp(1)
		document.formname.hAction.value = obj(obj.selectedIndex).text
		document.formname.action = "MRApprovalEntry.asp"
		document.formname.submit()
	end if

end Function

Function IssueAction(obj)
	dim arrTemp
	if obj.selectedIndex <> 0 then
		arrTemp = split(obj.value,"?")
		document.formname.mrs.value = arrTemp(0)
		document.formname.sAct.value = arrTemp(1)
		document.formname.hAction.value = obj(obj.selectedIndex).text
		'document.formname.action = "mrsHeaderDetails.asp"
		document.formname.action = "mrsIssueItemEntry.asp"
		document.formname.submit()
	end if
end Function

Function AmendAction(obj)
	dim arrTemp
	if not obj.selectedIndex = 0 then
		str = "Do you want the MRS to be " & obj(obj.selectedIndex).text
		if not confirm(str) then
			obj.selectedIndex = 0
			exit Function
		end if
		
		document.formname.hSelected.value = obj.name
		document.formname.hWhichMRS.value = "AM"
		arrTemp = split(obj.name,"Z")

		Set objhttp = CreateObject("Microsoft.XMLHTTP")

		objhttp.Open "POST","mrsMgmtInsert.asp?hSelected="&obj.name&"&hWhichMRS=AM&sAction="&obj.value, false
		objhttp.send
		
		if Left(objhttp.responseText,3) = "MRS" then
			set idAM = eval("idAmend"&arrTemp(1))
			idAM.innerHTML = objhttp.responseText
			if obj.value = "C" then
				set selObj = eval("document.formname.selAmendZ"&arrTemp(1))
				selObj.disabled = true
				set selObj = eval("idAMHref"&arrTemp(1))
				selObj.href = "#"
			elseif obj.value = "O" then
				set selObj = eval("document.formname.selAmendZ"&arrTemp(1))
				selObj.options.length = 1
				selObj.length = selObj.length+1
				selObj.options(selObj.length-1).text = "Un Hold"
				selObj.options(selObj.length-1).Value = "U"
				selObj.length = selObj.length+1
				selObj.options(selObj.length-1).text = "Cancel"
				selObj.options(selObj.length-1).Value = "C"
			elseif obj.value = "U" then
				set selObj = eval("document.formname.selAmendZ"&arrTemp(1))
				selObj.options.length = 1
				selObj.length = selObj.length+1
				selObj.options(selObj.length-1).text = "On Hold"
				selObj.options(selObj.length-1).Value = "O"
				selObj.length = selObj.length+1
				selObj.options(selObj.length-1).text = "Cancel"
				selObj.options(selObj.length-1).Value = "C"
			end if
		else
			MsgBox objhttp.responseText
		end if

	end if

end Function

Function ApproveAction(obj)
	dim arrTemp
	if not obj.selectedIndex = 0 then
		str = "Do you want the MRS to be " & obj(obj.selectedIndex).text
		if not confirm(str) then
			obj.selectedIndex = 0
			exit Function
		end if

		document.formname.hSelected.value = obj.name
		document.formname.hWhichMRS.value = "AP"
		arrTemp = split(obj.name,"Z")

		Set objhttp = CreateObject("Microsoft.XMLHTTP")

		objhttp.Open "POST","mrsMgmtInsert.asp?hSelected="&obj.name&"&hWhichMRS=AP&sAction="&obj.value, false
		objhttp.send
		
		if Left(objhttp.responseText,3) = "MRS" then
			set idAP = eval("idApprove"&arrTemp(1))
			idAP.innerHTML = objhttp.responseText
			if obj.value = "C" then
				set selObj = eval("document.formname.selApproveZ"&arrTemp(1))
				selObj.disabled = true
				set selObj = eval("idAPHref"&arrTemp(1))
				selObj.href = "#"
			elseif obj.value = "O" then
				set selObj = eval("document.formname.selApproveZ"&arrTemp(1))
				selObj.options.length = 1
				selObj.length = selObj.length+1
				selObj.options(selObj.length-1).text = "Un Hold"
				selObj.options(selObj.length-1).Value = "U"
				selObj.length = selObj.length+1
				selObj.options(selObj.length-1).text = "Cancel"
				selObj.options(selObj.length-1).Value = "C"
			elseif obj.value = "U" then
				set selObj = eval("document.formname.selApproveZ"&arrTemp(1))
				selObj.options.length = 1
				selObj.length = selObj.length+1
				selObj.options(selObj.length-1).text = "On Hold"
				selObj.options(selObj.length-1).Value = "O"
				selObj.length = selObj.length+1
				selObj.options(selObj.length-1).text = "Cancel"
				selObj.options(selObj.length-1).Value = "C"
			end if
		else
			MsgBox objhttp.responseText
		end if

	end if

end Function

