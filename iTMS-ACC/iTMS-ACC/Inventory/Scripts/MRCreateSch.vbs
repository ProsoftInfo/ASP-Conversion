dim objTemp,Root,newElem
dim iClass,iItem,iQty,iEntNo

Function checkNumbers(val,sStr)
	dim valid,temp,i
	if sStr = "Y" then
		valid = "0123456789."
	else
		valid = "0123456789"
	end if
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

Function fnInit(sItem,sClass,sQty,sEntNo)
	iClass = sClass
	iItem = sItem
	iQty = sQty
	iEntNo = sEntNo
	
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement
	
end Function

Function setMax(objp)
dim ii,objDa,objQa
	
	
	for ii = 1 to 12 
		set objDa = eval("document.formname.txtD"&ii)
		set objQa = eval("document.formname.txtQ"&ii)
		objDa.value = ""
		objQa.value = ""
	next		
	if objP.selectedIndex = 1 then
		for ii = 1 to 12 
			set objDa = eval("document.formname.txtD"&ii)
			objDa.maxlength = 10
		next		
	elseif objP.selectedIndex = 2 or objP.selectedIndex = 3 then
		for ii = 1 to 12 
			set objDa = eval("document.formname.txtD"&ii)
			objDa.maxlength = 6
		next		
	elseif objP.selectedIndex = 4 then
		for ii = 1 to 12 
			set objDa = eval("document.formname.txtD"&ii)
			objDa.maxlength = 8
		next		
	end if
end Function

Function CheckSubmit(todaysdate)
	dim i,iQtyTot,sTemp
	dim objQ,objD
	sTemp = "--"
	if document.formname.selschtype.selectedIndex = 0 then
		MsgBox "Select Schedule Type",0,"Schedule Type"
		document.formname.selschtype.focus
		exit function
	end if
	
	i = 1
	if document.formname.selschtype.selectedIndex = "1" then
		for i = 1 to 12 
			set objD = eval("document.formname.txtD"&i)
			set objQ = eval("document.formname.txtQ"&i)
		
			if trim(objD.value) <> "" then
				if (not vd(objD.value,todaysdate)) then
					MsgBox "Invalid Date",0,"Invalid Date"
					objD.select()
					Exit Function
				elseif DateDiff("d",todaysdate,objD.value) < 0 then
					MsgBox "Date should be greater or equal to Today's Date",0,"Invalid Date"
					objD.select()
					Exit Function
				elseif trim(objQ.value) = "" then
					msgbox "Enter Quantity",0,"Quantity"
					objQ.select()
					exit function
				elseif not checkNumbers(objQ.value,"Y") then
					msgbox "Enter Numerals Only",0,"Numerals"
					objQ.select()
					exit function
				else
					if InStr(1,sTemp,objD.value) > 0 then
						msgbox "No duplication date allowed",0,"Duplication"
						objD.select()
						exit function
					end if
					sTemp = sTemp & objD.value
					iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
				end if
			end if
		next
	end if

	if document.formname.selschtype.selectedIndex = "2" then
		
		for i = 1 to 12 
			
			set objD = eval("document.formname.txtD"&i)
			set objQ = eval("document.formname.txtQ"&i)
		
			if trim(objD.value) <> "" then
				if trim(len(objD.value)) <> 6 then
					MsgBox "Invalid Format (Month 2 and Year 4 characters)",0,"Invalid Format"
					objD.select()
					Exit Function
				elseif not checkNumbers(objD.value,"N") then
					msgbox "Invalid Format",0,"Invalid Format"
					objD.select()
					exit function
				elseif (cdbl(left(objD.value,2)) < 0) or (cdbl(left(objD.value,2)) > 12) then
					MsgBox "Invalid Month",0,"Invalid Month"
					objD.select()
					Exit Function
				elseif (cdbl(mid(trim(objD.value),3)) < cdbl(year(date()))) then
					MsgBox "Year should be greater than or equal to Current Year",0,"Invalid Year"
					objD.select()
					Exit Function
				elseif (cdbl(left(objD.value,2)) < cdbl(month(date())) and cdbl(mid(trim(objD.value),3)) = cdbl(year(date()))) then
					MsgBox "Month should be greater than or equal to Current month",0,"Invalid Month"
					objD.select()
					Exit Function
				elseif trim(objQ.value) = "" then
					msgbox "Enter Quantity",0,"Quantity"
					objQ.select()
					exit function
				elseif not checkNumbers(objQ.value,"Y") then
					msgbox "Enter Numerals Only",0,"Numerals"
					objQ.select()
					exit function
				else
					if InStr(1,sTemp,objD.value) > 0 then
						msgbox "No duplication date allowed",0,"Duplication"
						objD.select()
						exit function
					end if
					sTemp = sTemp & objD.value
					iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
				end if
			end if
		next
	end if

	if document.formname.selschtype.selectedIndex = "3" then
		
		for i = 1 to 12 
			
			set objD = eval("document.formname.txtD"&i)
			set objQ = eval("document.formname.txtQ"&i)
		
			if trim(objD.value) <> "" then
				if trim(len(objD.value)) <> 6 then
					MsgBox "Invalid Format (Week 2 and Year 4 characters)",0,"Invalid Format"
					objD.select()
					Exit Function
				elseif not checkNumbers(objD.value,"N") then
					msgbox "Invalid Format",0,"Invalid Format"
					objD.select()
					exit function
				elseif (cdbl(left(objD.value,2)) < 0) or (cdbl(left(objD.value,2)) > 52) then
					MsgBox "Invalid Week",0,"Invalid Week"
					objD.select()
					Exit Function
				elseif (cdbl(mid(trim(objD.value),3)) < cdbl(year(date()))) then
					MsgBox "Year should be greater than or equal to Current Year",0,"Invalid Year"
					objD.select()
					Exit Function
				elseif (cdbl(left(objD.value,2)) < cdbl(DatePart("ww",date())) and cdbl(mid(trim(objD.value),3)) = cdbl(year(date()))) then
					MsgBox "Week should be greater than or equal to Current Week",0,"Invalid Week"
					objD.select()
					Exit Function
				elseif trim(objQ.value) = "" then
					alert(DatePart("ww",date()))
					msgbox "Enter Quantity",0,"Quantity"
					objQ.select()
					exit function
				elseif not checkNumbers(objQ.value,"Y") then
					msgbox "Enter Numerals Only",0,"Numerals"
					objQ.select()
					exit function
				else
					if InStr(1,sTemp,objD.value) > 0 then
						msgbox "No duplication date allowed",0,"Duplication"
						objD.select()
						exit function
					end if
					sTemp = sTemp & objD.value
					iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
				end if
			end if
		next
	end if

	if document.formname.selschtype.selectedIndex = "4" then
		
		for i = 1 to 12 
			
			set objD = eval("document.formname.txtD"&i)
			set objQ = eval("document.formname.txtQ"&i)
		
			if trim(objD.value) <> "" then
				if trim(len(objD.value)) <> 8 then
					MsgBox "Invalid Format (Month 2, Week 2 and Year 4 characters)",0,"Invalid Format"
					objD.select()
					Exit Function
				elseif not checkNumbers(objD.value,"N") then
					msgbox "Invalid Format",0,"Invalid Format"
					objD.select()
					exit function
				elseif (cdbl(left(objD.value,2)) < 0) or (cdbl(left(objD.value,2)) > 12) then
					MsgBox "Invalid Month",0,"Invalid Month"
					objD.select()
					Exit Function
				elseif (cdbl(mid(objD.value,3,2)) < 0) or (cdbl(mid(objD.value,3,2)) > 4) then
					MsgBox "Invalid Week",0,"Invalid Week"
					objD.select()
					Exit Function
				elseif (cdbl(mid(trim(objD.value),5)) < cdbl(year(date()))) then
					MsgBox "Year should be greater than or equal to Current Year",0,"Invalid Year"
					objD.select()
					Exit Function
				elseif (cdbl(left(objD.value,2)) < cdbl(month(date())) and cdbl(mid(trim(objD.value),5)) = cdbl(year(date()))) then
					MsgBox "Month should be greater than or equal to Current month",0,"Invalid Month"
					objD.select()
					Exit Function
				elseif trim(objQ.value) = "" then
					msgbox "Enter Quantity",0,"Quantity"
					objQ.select()
					exit function
				elseif not checkNumbers(objQ.value,"Y") then
					msgbox "Enter Numerals Only",0,"Numerals"
					objQ.select()
					exit function
				else
					if InStr(1,sTemp,objD.value) > 0 then
						msgbox "No duplication date allowed",0,"Duplication"
						objD.select()
						exit function
					end if
					sTemp = sTemp & objD.value
					iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
				end if
			end if
		next
	end if

	if (cdbl(iQtyTot) < cdbl(iQty)) or (cdbl(iQtyTot) > cdbl(iQty)) then
		msgbox "Total Schedule Quantity should be equal to Quantity Required",0,"Schedule Quantity"
		exit function
	end if

	i = 1
	Set Root = objTemp.documentElement
	'alert(Root.xml)
	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
		'alert(iEntNo & "**"& iItem)
			if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(2).nodeValue = iClass and HeaderNode.Attributes.Item(1).nodeValue = iItem then
				For Each PageNode In HeaderNode.childNodes
					if StrComp(PageNode.nodeName,"Schedule") = 0 then
						for i = 1 to 12 
							Set newElem = objTemp.createElement("ScheduleDetails")
							newElem.setAttribute "SNO", i
							newElem.setAttribute "NEED", eval("document.formname.txtD"&i&".value")
							newElem.setAttribute "QTY", eval("document.formname.txtQ"&i&".value")
							newElem.setAttribute "TYPE", trim(document.formname.selschtype(document.formname.selschtype.selectedIndex).value)
							PageNode.appendChild newElem
						next
					end if
				next
				window.close
				exit function
			end if
		next
	end if
end Function

Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	'window.close()
end Function

