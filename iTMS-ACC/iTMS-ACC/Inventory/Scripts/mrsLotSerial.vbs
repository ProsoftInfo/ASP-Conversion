dim objTemp,Root,newElem
dim iLineNo,iPickNo,iQty,ii,sLot,sBin,sStore,sOrgID,dSchMinDate,dSchMaxDate

Function checkNumbers(val)
	dim valid,temp,i
	valid = "0123456789."
	for i=1 to len(val)
		temp = mid(val,i,1)
		if Instr(1,valid,temp) > 0 then
			checkNumbers = true
		else
			checkNumbers = false
			exit for
		end if
	next
end Function

Function ChangeDet(sRadValue)
	dim ictr,i
	ictr = document.formname.hiCtr.value
	if sRadValue = "M" then
		for i=1 to ictr
			set objD = eval("document.formname.txtDate"&i)
			objD.value = ""
			objD.readonly = false
		next	
	else
		for i=1 to ictr
			set objD = eval("document.formname.txtDate"&i)
			objD.value = ""
			objD.value = document.formname.ctlDDate.GetDate
			objD.readonly = true
		next	
	end if
end Function

Function fnInit(obj,dMaxDate)
	if document.formname.hiCtr.value = 0 then exit function

	set objTemp = window.dialogArguments

	Set Root = objTemp.documentElement
	
	arrTemp = split(obj,"`")
	
	iLineNo = arrTemp(1)
	iPickNo = arrTemp(2)
	sLot = arrTemp(3)
	sStore = arrTemp(4)
	sBin = arrTemp(5)
	
	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(0).nodeValue = iLineNo and HeaderNode.Attributes.Item(1).nodeValue = iPickNo then
			dSchMinDate = HeaderNode.Attributes.Item(7).nodeValue
			dSchMaxDate = HeaderNode.Attributes.Item(8).nodeValue
		end if
	next	
	'alert(dMaxDate)
	'alert(dSchMinDate)
	document.formname.ctlDDate.SetMaxDate = dMaxDate
	document.formname.ctlDDate.SetMinDate = dSchMinDate
	
	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(0).nodeValue = iLineNo and HeaderNode.Attributes.Item(1).nodeValue = iPickNo then
			if HeaderNode.HaschildNodes() then
				For Each SHNode In HeaderNode.childNodes
					if StrComp(Trim(SHNode.NodeName),"SERIALHEADER") = 0 then

						if SHNode.Attributes.Item(0).nodeValue = "M" then
							document.formname.radDate(0).checked = true
						else
							document.formname.radDate(1).checked = true
						end if

						document.formname.ctlDDate.SetDate = SHNode.Attributes.Item(1).nodeValue

						For Each HNode In SHNode.childNodes
							if StrComp(Trim(HNode.NodeName),"SERIALDETAILS") = 0 then
								ii = ii + 1
								set Q = eval("document.formname.txtQty"&ii)
								Q.value = HNode.Attributes.Item(1).nodeValue
								set D = eval("document.formname.txtDate"&ii)
								D.value = HNode.Attributes.Item(2).nodeValue
							end if
						next

					end if
				next
			end if
			exit function
		end if
	Next
end Function

Function CheckSubmit(todaysdate)
	dim ictr,objQ,iQtyTot,objSTQ,objSerial,i
	
	ictr = document.formname.hiCtr.value
	
	if ictr = "0" then exit function 
	
	if not (document.formname.radDate(0).checked or document.formname.radDate(1).checked) then
		alert("Select Date")
		document.formname.radDate(0).focus()
		exit function
	end if

	if document.formname.radDate(1).checked then
		sTempDate = document.formname.ctlDDate.GetDate

		if(datediff("d",dSchMinDate,sTempDate)) < 0 then
			alert("Date should be greater than or equal to Schedule Date")
			exit function
		end if
		if(datediff("d",todaysdate,sTempDate)) > 0 then
			alert("Date should be less than or equal to Today's Date")
			exit function
		end if
	end if
	
	for i=1 to ictr
		set objQ = eval("document.formname.txtQty"&i)
		set objSTQ = eval("document.formname.txtStQty"&i)
		set objD = eval("document.formname.txtDate"&i)

		if trim(objQ.value) = "" then
			msgbox "Enter Quantity",0,"Quantity"
			objQ.select()
			exit function
		elseif not checkNumbers(objQ.value) then
			msgbox "Enter Numerals Only",0,"Numerals"
			objQ.select()
			exit function
		else
			if document.formname.radDate(0).checked and trim(objQ.value = "") then
				alert("Enter Quantity")
				objQ.select()
				exit function
			end if
			if document.formname.radDate(0).checked and cdbl(objQ.value > 0) and trim(objD.value = "") then
				alert("Enter Date")
				objD.select()
				exit function
			end if
			if document.formname.radDate(0).checked and cdbl(objQ.value > 0) and trim(objD.value <> "") then
				if not vd(objD.value,todaysdate) then
					alert("Invalid date")
					objD.select()
					exit function
				end if

				sTempDate = objD.value

				if(datediff("d",dSchMinDate,sTempDate)) < 0 then
					alert("Date should be greater than or equal to Schedule Date")
					objD.select()
					exit function
				end if
				if(datediff("d",todaysdate,sTempDate)) > 0 then
					alert("Date should be less than or equal to Today's Date")
					objD.select()
					exit function
				end if

			end if
			if (cdbl(objQ.value) > cdbl(objSTQ.value)) then
				msgbox "Quantity Issue should be equal to or less than Stock Quantity",0,"Quantity"
				objQ.select()
				exit function
			end if

			iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
		end if

	next

	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(0).nodeValue = iLineNo and HeaderNode.Attributes.Item(1).nodeValue = iPickNo then
			if HeaderNode.HaschildNodes() then
				For Each HNode In HeaderNode.childNodes
					if StrComp(Trim(HNode.NodeName),"SERIALHEADER") = 0 then
						set a = HeaderNode.removeChild(HNode)
					end if
				next
			end if
		end if
	Next

	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(0).nodeValue = iLineNo and HeaderNode.Attributes.Item(1).nodeValue = iPickNo then
			Set newElem = objTemp.createElement("SERIALHEADER")
			if document.formname.radDate(0).checked then
				newElem.setAttribute "DATETYPE", "M"
			else
				newElem.setAttribute "DATETYPE", "S"
			end if
			newElem.setAttribute "DATEVALUE", document.formname.ctlDDate.GetDate

			for i=1 to ictr
				set objSerial = eval("document.formname.hSerial"&i)
				Set newElem1 = objTemp.createElement("SERIALDETAILS")
				newElem1.setAttribute "SERIALNO", trim(objSerial.value)
				newElem1.setAttribute "QTY", trim(eval("document.formname.txtQty"&i&".value"))
				newElem1.setAttribute "TDATE", trim(eval("document.formname.txtDate"&i&".value"))
				newElem.appendChild newElem1
			next
			HeaderNode.appendChild newElem
		end if
	Next

	if (cdbl(iQtyTot) > cdbl(idQty.innerText)) then
		msgbox "Total Quantity Issue should be equal to or less than Quantity Requested",0,"Quantity"
		exit function
	end if

	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(0).nodeValue = iLineNo and HeaderNode.Attributes.Item(1).nodeValue = iPickNo then
			HeaderNode.setAttribute "ISSQTY", cdbl(iQtyTot)
			window.close
			exit function
		end if
	Next
	
end Function

Function window_onunload() 
	
	set window.returnValue = objTemp.documentElement
	window.close()
end Function
