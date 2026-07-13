dim objTemp,Root,RootO
dim iClass,iItem

Set Root = PickData.documentElement

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

Function fnInit(sClass,sItm,iEntryNo)

	set objTemp = window.dialogArguments
	Set RootO = objTemp.documentElement

	iClass = sClass
	iItem = sItm

	'alert(RootO.xml)

	if document.formname.hiCtr.value = ""  then exit function
	
	sExp ="//ITM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItem&" and PICKNO = "& iEntryNo &" ]/PickDet/PICK"
	Set ItemNode = RootO.Selectnodes(sExp)
	For iCounter = 0 to ItemNode.Length - 1
		iSerialNumber = ItemNode.Item(iCounter).Attributes.getNamedItem("SNO").Value 
		iLineNumber = ItemNode.Item(iCounter).Attributes.getNamedItem("LINENO").Value 
		set Q = eval("document.formname.txtIss"&iSerialNumber)
		Q.value = ItemNode.Item(iCounter).Attributes.getNamedItem("ISSQTY").Value 
	next
	
end Function

Function CheckLot(obj,iCtr,iQty,iItem,iClass,sAttID)
	dim sItem,sClass,aobj,iSerialNo
	
	arrTemp = split(obj.name,"`")
	
	iEntryNo = arrTemp(1)
	iIssueEntryNo = arrTemp(2)
	sLot = arrTemp(3)
	sStore = arrTemp(4)
	sBin = arrTemp(5)
	
	
	Set RootO = objTemp.documentElement
	sTempValues = obj.name&"`"&iQty

	sExp ="//ITM [ @CLACODE = "&iClass&" and @ITMCODE = "&iItem&" and @PICKNO = "& iEntryNo &"]/PickDet"
	Set ItemNode = RootO.Selectnodes(sExp)
	
	if ItemNode.Length > 0 then
		Set PickData.documentElement = ItemNode.Item(0)
	end if			

'	set OutDataValue = showModalDialog("mrsPickLotPoP.asp?sTemp="&sTempValues,PickData,"dialogHeight:380px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
'	'alert(PickData.xml)
'	Set Root = PickData.documentElement
'	'alert(Root.xml)
'	For Each HeaderNode In Root.childNodes
'		if HeaderNode.Attributes.Item(10).nodeValue = iCtr then
'			if HeaderNode.HaschildNodes() then
'				set Q = eval("document.formname.txtIss"&iCtr)
'				Q.value = HeaderNode.Attributes.Item(6).nodeValue
'			end if
'			exit function
'		end if
'	Next
sOrgID = document.formname.hOrgID.value

'hValueZ:1:10/CLTG:4:0:550:144:7540:010101::1:::

	sTempValues = "hValueZ:"&iCtr&":"&sLot&":"&sStore&":"&sBin&"::"&iClass&":"&iItem&":"&sOrgID&"::"&iEntryNo&"::"&sAttID&":"
	'sTempValues = obj.value&":"&iClass&":"&iItem&":"&sOrg&":"&InvRecNo&":"&iEntNo&":"&sOptName&":"&sAttID&":"&sAttList
	'alert(sTempValues)

	set OutDataValue = window.showModalDialog("IssuePickLotPop.asp?sTemp="&sTempValues&"&IssNo="&iIssueEntryNo,PickData,"dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No")
	sExp = "//PickDet/PICK[@LOC = """&sStore&""" and @BIN = """&sBin&""" and @ITEMENTRYNO="""&iEntryNo&"""]/Pagination"
	Set Root = PickData.documentElement
	Set Tempnode = Root.Selectnodes(sExp)
	If Tempnode.Length > 0 Then
		sTempValues = split(Tempnode.Item(0).Attributes.Item(0).Nodevalue,"```")
		while UBound(sTempValues) = 0
			sTempValues = replace(Tempnode.Item(0).Attributes.Item(0).Nodevalue,"|","&")
			set OutDataValue = window.showModalDialog("IssuePickLotPop.asp?sTemp="&sTempValues&"&IssNo="&iIssueEntryNo,PickData,"dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No")
			sExp = "//PickDet/PICK[@LOC = """&sStore&""" and @BIN = """&sBin&""" and @ITEMENTRYNO="""&iEntryNo&"""]/Pagination"
			Set Tempnode = Root.Selectnodes(sExp)
			If Tempnode.Length > 0 Then
				sTempValues = split(Tempnode.Item(0).Attributes.Item(0).Nodevalue,"```")
			end if
		wend
		
		if UBound(sTempValues)>0 then
		    if left(trim(sTempValues(1)),2)="NO" then exit function
	    end if 'if UBound(sTempValues)>0 then
		
	end if
	
	
	'alert(Root.xml)
	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(10).nodeValue = iCtr then
			if HeaderNode.HaschildNodes() then
				set Q = eval("document.formname.txtIss"&iCtr)
				Q.value = HeaderNode.getAttribute("ISSQTY")
			end if
			exit function
		end if
	Next

end Function

Function CheckSubmit(todaysdate,iItem,iClass,iEntryNo)
	dim i,iSerialNo,iTotWOQty,iTotWQty,arrTemp

	iTotWOQty = 0
	iTotWQty = 0
'alert(iItem &"***"& iClass)
	for i=0 to document.formname.elements.length - 1
		if document.formname.elements(i).type = "text" then
			if (Instr(document.formname.elements(i).name,"txtQty") > 0) then
				if (trim(document.formname.elements(i+2).value) <> "") then
					if(not checkNumbers(document.formname.elements(i+2).value)) then
						msgbox "Enter Numerals Only",0,"Numerals"
						document.formname.elements(i+2).select()
						exit function
					else
						if not document.formname.elements(i).value = "-" then
							iTotWQty = cdbl(iTotWQty) + cdbl(document.formname.elements(i+2).value)
							if (cdbl(document.formname.elements(i+2).value)) > cdbl(document.formname.elements(i).value) then
								msgbox "Quantity Issue should be less than or equal to Quantity Marked",0,"Quantity"
								document.formname.elements(i+2).select()
								exit function
							end if
							if (cdbl(document.formname.elements(i+2).value)) > cdbl(document.formname.elements(i-1).value) then
								msgbox "Quantity Issue should be less than or equal to Stock Quantity",0,"Quantity"
								document.formname.elements(i+2).select()
								exit function
							end if
						else
							iTotWOQty = cdbl(iTotWOQty) + cdbl(document.formname.elements(i+2).value)
							if (cdbl(document.formname.elements(i+2).value)) > cdbl(document.formname.elements(i-1).value) then
								msgbox "Quantity Issue should be less than or equal to Stock Quantity",0,"Quantity"
								document.formname.elements(i+2).select()
								exit function
							end if
						end if
					
					'	if (document.formname.elements(i+1).disabled and trim(document.formname.elements(i+2).value) <> "") then
					'		if(trim(document.formname.elements(i+3).value) = "") then
					'			alert("Enter Date")
					'			document.formname.elements(i+3).select()
					'			exit function
					'		end if
					'		if not vd(trim(document.formname.elements(i+3).value),todaysdate) then
					'			alert("Invalid date")
					'			document.formname.elements(i+3).select()
					'			exit function
					'		end if
					'		
					'	end if

						if document.formname.elements(i+2).value = "" then
							sQty = "0"
						else
							sQty = cdbl(document.formname.elements(i+2).value)
						end if
					
						sTotQty = cdbl(sTotQty) + sQty

					end if
				else
					sQty = "0"
					sTotQty = cdbl(sTotQty) + sQty
				end if
			end if
		end if
	next

	if (cdbl(sTotQty) = 0) then
		alert("Enter Quantity Issue or select from Serial number(s)")
		exit function
	elseif (cdbl(iTotWQty) > cdbl(idWQty.innerText)) then
		alert("Quantity Issue from Lot Selection should be less than or equal to (" & idWQty.innerText & ")")
		exit function
	elseif (cdbl(iTotWOQty) > cdbl(idWOQty.innerText)) then
		alert("Quantity Issue from without Lot Selection should be less than or equal to (" & idWOQty.innerText & ")")
		exit function
	elseif (cdbl(sTotQty) > cdbl(idQty.innerText)) then
		alert("Quantity Issue should be less than or equal to total quantity available for Pick (" & idQty.innerText & ")")
		exit function
	end if

	Set Root = PickData.documentElement
	
	iSerialNo = 0

'alert(RootO.xml)
	sExp ="//ITM [ @CLACODE ="""&iClass&""" and @ITMCODE="""&iItem&""" and @ItemEntNo ="""& iEntryNo &"""]/PickDet"
	Set ItemNode = RootO.Selectnodes(sExp)
	'alert(sExp &" = "& ItemNode.length)
	
	if ItemNode.Length > 0 then
		ItemNode.Item(0).Attributes.getNamedItem("TOT").Value = sTotQty
	else
		Root.setAttribute "TOT", trim(sTotQty)
	end if


	sExp ="//ITM [ @CLACODE ="""&iClass&""" and @ITMCODE ="""&iItem&""" and @ItemEntNo ="""& iEntryNo &"""]/PickDet/PICK"
	Set ItemNode = RootO.Selectnodes(sExp)
	'alert(sExp &" = "& ItemNode.length)
	if ItemNode.Length > 0 then
		For iCounter = 0 to ItemNode.Length - 1
			iSerialNo = ItemNode.Item(iCounter).Attributes.getNamedItem("SNO").Value 
			iLineNumber = ItemNode.Item(iCounter).Attributes.getNamedItem("LINENO").Value 
			set Q = eval("document.formname.txtIss"&iSerialNo)
			ItemNode.Item(iCounter).Attributes.getNamedItem("ISSQTY").Value = Q.value
		next
	else
		For Each HeaderNode In Root.childNodes
			iSerialNo = HeaderNode.Attributes.Item(10).nodeValue
			iLineNo = HeaderNode.Attributes.Item(0).nodeValue
			set objQty = eval("document.formname.txtIss"&iSerialNo)
			if objQty.value = "" then
				HeaderNode.setAttribute "ISSQTY", ""
			else
				HeaderNode.setAttribute "ISSQTY", cdbl(objQty.value)
			end if
		Next

		For Each HeaderNode In RootO.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iClass and HeaderNode.Attributes.Item(1).nodeValue = iItem and HeaderNode.getAttribute("ItemEntNo")=iEntryNo then
				HeaderNode.appendChild Root
			end if
		Next
	end if
'	alert(RootO.xml)
	window_onunload()
	
end Function

Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	window.close()
end Function
