dim objTemp,Root,RootO,j,iTotQty
dim iClass,iItem,sUsage,sOrgID
j = 0
iTotQty = 0

Function fnInit(sOrg,sClass,sItm,sUsag)
	set objTemp = window.dialogArguments
	Set RootO = objTemp.documentElement

	iClass = sClass
	iItem = sItm
	sUsage = sUsag
	sOrgID = sOrg

	sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItem&"]/PackingDet/PCode"
	Set PackNode = RootO.Selectnodes(sExp)
	i = 0
	if PackNode.Length > 0 then
		for i = 0 to PackNode.Length - 1
			set oQty = eval("document.formname.txtQty"&i+1)
			set oMC = eval("document.formname.hPONO"&i+1)

			if oMC.value = PackNode.Item(i).Attributes.getNamedItem("PONO").Value then
				oQty.value = PackNode.Item(i).Attributes.getNamedItem("QTY").Value
			end if
		next
	end if
	i = 1
end Function

Function CheckSubmit()
	iCtr = cdbl(document.formname.hiCtr.value)
	
	if iCtr = 0 then exit Function
	
	iTotQty = 0
	
	for i = 1 to iCtr
		set oQty = eval("document.formname.txtQty"&i)
		if trim(oQty.value) = "" then
			alert("Enter Quantity")
			oQty.select()
			exit function
		elseif cdbl(oQty.value) = 0 then
			alert("Quantity cannot be ZERO")
			oQty.select()
			exit function
		else
			iTotQty = cdbl(iTotQty) + cdbl(oQty.value)	
		end if
	next
	
	sExp ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItem&"]/PackingDet/PCode"
	Set PackNode = RootO.Selectnodes(sExp)

	if (cdbl(iTotQty) <> cdbl(idQty.innerText)) then
		alert("Quantity breakup should be equal to Quantity Issue")
		exit function
	end if

	sExp1 ="//ITEM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItem&"]"

	Set HeaderNode = RootO.Selectnodes(sExp1)

	i = 0
	if PackNode.Length > 0 then
		for i = 0 to PackNode.Length - 1
			set oQty = eval("document.formname.txtQty"&i+1)
			set oMC = eval("document.formname.hPONO"&i+1)

			if oMC.value = PackNode.Item(i).Attributes.getNamedItem("PONO").Value then
				PackNode.Item(i).Attributes.getNamedItem("PONO").Value = trim(oMC.value)
				PackNode.Item(i).Attributes.getNamedItem("QTY").Value = trim(oQty.value)
			end if
		next

	else
		Set newElem1 = objTemp.createElement("PackingDet")
		i = 1
		for i = 1 to iCtr
			set oQty = eval("document.formname.txtQty"&i)
			set oMC = eval("document.formname.hPONO"&i)

			Set newElem2 = objTemp.createElement("PCode")
			newElem2.setAttribute "PONO", trim(oMC.value)
			newElem2.setAttribute "QTY", trim(oQty.value)

			newElem1.appendChild newElem2
		next
		HeaderNode.Item(0).appendChild newElem1
	end if

	window.close
end Function

Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	window.close()
end Function
