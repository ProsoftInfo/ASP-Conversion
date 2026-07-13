dim OutDataValue,Root

Set Root = OutData.documentElement
Function SetDate(dDate)
IF DateValue(dDate) < DateValue(document.formname.hMinDate.Value) or DateValue(dDate) > DateValue(document.formname.hMaxDate.Value)  then
	document.formname.ctlIssDate.setDate = document.formname.hMaxDate.value	
Else
	document.formname.ctlIssDate.setDate = dDate	
End IF
End Function
Function MinDate()
	If DateValue(document.formname.ctlIssDate.getDate) < Datevalue(document.formname.hMinDate.Value) or DateValue(document.formname.ctlIssDate.getDate) > Datevalue(document.formname.hMaxDate.Value)  then
		Alert("Issue Date Should Be With in the Financial Year " & document.formname.hMinDate.Value &" and " &document.formname.hMaxDate.Value)
		document.formname.ctlIssDate.setDate = document.formname.hMaxDate.Value
		exit function
	End IF	
End Function
Function Back()
	'document.formname.action = "mrsHeaderDetails.asp"
	document.formname.action = "MRSMGMTLIST.ASP?hCheck=I"
	document.formname.submit()
end Function
   
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

Function DisplayDet(sText)
	FontFace="Verdana,8,,bold" 
	arrTemp = split(sText,"|")
	TopicText = "To Purchase Requisition : " & arrTemp(0) & vbcrlf
	TopicText = TopicText & "For Stock Transfer         : " & arrTemp(1)

	'document.formname.penDet.TextPopup TopicText, FontFace, 10,10,0,0
end Function

Function DisplayStock(obj)
	dim sItem,sClass,sLoc,sBin
	arrTemp = split(obj.name,"Z")
 '	alert(obj.name)
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)
	sLoc = arrTemp(5)
	sBin = arrTemp(6)

	sTempValues = sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&iEntNo&":"&sOptName&":"&sLoc&":"&sBin

	showModalDialog "itmStockPoP.asp?sTemp="&sTempValues,OutData,"dialogHeight:450px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No"
end Function

Function LoadData()
	dim Q,S,i,ic
	
	exit function
	for i=0 to document.formname.elements.length - 1
		if (document.formname.elements(i).type = "text") then
			if left(document.formname.elements(i).name,7) = "txtQtyA" then
				arrTemp = split(document.formname.elements(i).name,"A")
				sItem = arrTemp(2)
				sClass = arrTemp(1)

				Set RootO = Data.documentElement
				For Each HeaderNodeO In RootO.childNodes
					if HeaderNodeO.NodeName = "ITEM" then
						if HeaderNodeO.Attributes.Item(0).nodeValue = sItem and HeaderNodeO.Attributes.Item(1).nodeValue = sClass then
							'sTemp = cstr(HeaderNodeO.Attributes.Item(0).nodeValue)&"A"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)&"A"&sQty&"A"&iValue

							'sTemp = cstr(HeaderNodeO.Attributes.Item(0).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)
							'set Q = eval("document.formname.txtQtyReqX"&sTemp)
							'Q.value = HeaderNodeO.Attributes.Item(4).nodeValue

							'sTemp = cstr(HeaderNodeO.Attributes.Item(0).nodeValue)&"A"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)
							'set Q = eval("document.formname.txtQtyA"&sTemp)
							'Q.value = HeaderNodeO.Attributes.Item(5).nodeValue

							'sTemp = cstr(HeaderNodeO.Attributes.Item(0).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)
							'set Q = eval("document.formname.txtQtyPrX"&sTemp)
							Q.value = "0"
							'Q.value = HeaderNodeO.Attributes.Item(5).nodeValue

							'sTemp = cstr(HeaderNodeO.Attributes.Item(0).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)
							'set Q = eval("document.formname.txtQtyTraX"&sTemp)
							'Q.value = "0"
							'Q.value = HeaderNodeO.Attributes.Item(5).nodeValue
						end if
					end if
				next
			end if
		end if
	next
end Function

Function GetAddDetails(sOrg,iclass,iItem,iMRSNo)
	dim  sIssuedFor,iQty
	iQty = eval("document.formname.txtQtyPPX"&iClass&"x"&iItem).value
	sTempValues = iclass&"|"&iItem&"|"&sOrg&"|"&iMRSNo&"|"&iQty
	'alert OutData.XML
	set OutDataValue = showModalDialog("IssueAddDetails.asp?sTemp="&sTempValues,OutData,"dialogHeight:400px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No")		
End Function

Function CheckSch(obj)
	dim sItem,sClass,a,aobj
	dim qty
	arrTemp = split(obj.name,":")
	
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)
	set qty = eval("document.formname.txtQtyA"+cstr(sClass)+"A"+cstr(sItem)+"A"+cstr(iEntNo))

	sTempValues = qty.value&":"&sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&document.formname.hOrgID.value&":"&iEntNo&":"&sOptName

	OutDataValue = showModalDialog("mrsIssueSchedulePoP.asp?sTemp="&sTempValues,Data,"dialogHeight:480px;dialogWidth:390px;center:Yes;help:No;resizable:No;status:No")
End Function

Function GetSch(obj)
	dim sItem,sClass
	dim sch,qty

	arrTemp = split(obj.name,"X")
	
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	set sch = eval("document.formname.hSchA"+cstr(sClass)+"A"+cstr(sItem)+"A"+cstr(iEntNo))
	
	if sch.value <> "S" then exit function	
	
	if trim(obj.value) = "" then exit function
	if cdbl(obj.value) = 0 then exit function

	set qty = eval("document.formname.txtQtyA"+cstr(sClass)+"A"+cstr(sItem)+"A"+cstr(iEntNo))

	sTempValues = obj.value&":"&sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&document.formname.hOrgID.value&":"&iEntNo

	set OutDataValue = showModalDialog("mrsIssueScheduleEntryPoP.asp?sTemp="&sTempValues,OutData,"dialogHeight:490px;dialogWidth:330px;center:Yes;help:No;resizable:No;status:No")

	For Each HeaderNodeO In Root.childNodes
		if HeaderNodeO.NodeName = "ITEMDETAILS" then
			bsFlag = false
			if HeaderNodeO.Attributes.Item(0).nodeValue = iEntNo and HeaderNodeO.Attributes.Item(1).nodeValue = sItem and HeaderNodeO.Attributes.Item(2).nodeValue = sClass then
				For Each childNod In HeaderNodeO.childNodes
					if childNod.NodeName = "SCHEDULE" then
						bsFlag = true
						exit for
					end if
				next
				if not bsFlag then
					sTemp = cstr(HeaderNodeO.Attributes.Item(2).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(0).nodeValue)
					set Q = eval("document.formname.txtQtyPPX"&sTemp)
					Q.value = "0"
					'Q.readOnly = false
				end if
			end if
		end if
	next

end Function

Function CheckQty(obj)
	dim sItem,sClass,a
	arrTemp = split(obj.name,":")
	'alert(obj.name)	
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)
	sTempValues = sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&iEntNo&":"&sOptName

	OutDataValue = showModalDialog("mrsIssueQtyParaPoP.asp?sTemp="&sTempValues,OutData,"dialogHeight:385px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No")
end Function

Function DisplayItem(obj)
	sTempValues = obj

	showModalDialog "itmDetailsPop.asp?sTemp="&sTempValues,"","dialogHeight:400px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No"
end Function

Function CheckLot(obj,iQty)
	dim sItem,sClass,sLoc,sBin,sType,sOrgID
	
	arrTemp = split(obj.name,"Z")
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)
	sLoc = arrTemp(3)
	sBin = arrTemp(4)
	sOrgID = arrTemp(5)
	sArrList = arrTemp(7)
	
	sTempValues = sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&iEntNo&":"&sOptName&":"&sLoc&":"&sBin&":"&iQty &":"&sType&":"&sUsage&":"& sOrgID&":"& sArrList
	'alert(OutData1.xml)
	
	set OutDataValue = showModalDialog("MRPickDetPoP.asp?sTemp="&sTempValues,OutData1,"dialogHeight:310px;dialogWidth:500px;center:Yes;help:No;resizable:No;status:No")
	'alert(OutDataValue.xml)

	Set Root = OutData1.documentElement
'alert("A="&OutData1.xml)
	For Each HeaderNodeO In Root.childNodes
		if HeaderNodeO.NodeName = "ITEM" then
			bFlag = true
			if HeaderNodeO.Attributes.Item(0).nodeValue = iEntNo and HeaderNodeO.Attributes.Item(1).nodeValue = sItem  and HeaderNodeO.Attributes.Item(2).nodeValue = sClass then
				For Each childNod In HeaderNodeO.childNodes
					if childNod.NodeName = "Pick" then
					 
						bFlag = false
						sTemp = cstr(HeaderNodeO.Attributes.Item(2).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(0).nodeValue)
						set Q = eval("document.formname.txtQtyPPX"&sTemp)						
						Q.value = childNod.Attributes.Item(0).nodeValue
						if not cdbl(childNod.Attributes.Item(0).nodeValue) = 0 then
							Q.readOnly = true
							Q.focus()
						else
							Q.readOnly = false
						end if
						exit function
					end if
				next
				if bFlag then
					sTemp = cstr(HeaderNodeO.Attributes.Item(2).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(0).nodeValue)
					set Q = eval("document.formname.txtQtyPPX"&sTemp)
					Q.value = "0"
					'Q.readOnly = false
				end if
			end if
		end if
	next

end Function


Function CheckST(obj)
	dim sItem,sClass,a,aobj
	arrTemp = split(obj.name,":")
	'alert(obj.name)
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)
	sOrgId = arrTemp(5)

	set aobj = eval("document.formname.hStoName"&cstr(sClass)&"A"&cstr(sItem)&"A"&cstr(iEntNo))

	sTempValues = sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&aobj.value&":"&iEntNo&":"&sOptName&":"&sOrgId

	Set OutDataValue = showModalDialog("mrsIssueSTPoP.asp?sTemp="&sTempValues,OutData,"dialogHeight:330px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
'alert(OutDataValue.xml)
	'alert(OutData.xml)
	
	for i=0 to document.formname.elements.length - 1
		if (document.formname.elements(i).type = "text") then
			'arrTemp = split(document.formname.elements(i).name,"A")
			'sItem = arrTemp(2)
			'sClass = arrTemp(1)
			Set RootO = OutData.documentElement
			For Each HeaderNodeO In RootO.childNodes
				if HeaderNodeO.NodeName = "ITEM" then
					if HeaderNodeO.Attributes.Item(0).nodeValue = sClass and HeaderNodeO.Attributes.Item(1).nodeValue = sItem then
						sTemp = cstr(HeaderNodeO.Attributes.Item(0).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)
						set Q = eval("document.formname.txtQtyTraX"&sTemp)
						Q.value = HeaderNodeO.Attributes.Item(6).nodeValue
						exit function
					end if
				end if
			next
		end if
	next
end Function

Function CheckSubmit(todaysdate)

	dim i,iCtrr, j
	dim arrTemp,AddDetNode,subNode,sInvoiceSelected
	set Root1 = OutData.documentElement
	IF Root1.haschildnodes then 
		For each node in Root1.childnodes 
			If trim(node.NodeName) = "ITEM" then
				Root1.RemoveChild node
			End If 
		Next
	End IF	
	
	 
	If (datediff("d",todaysdate,document.formname.ctlIssDate.GetDate)) > 0 then
		alert("Issue Date should be less than or equal to Today's Date")
		exit function
	End If
	dReqDate = document.formname.hReqDate.value
'	alert(dReqDate)
'	alert(document.formname.ctlIssDate.getDate())
'	alert(datediff("d",dReqDate,document.formname.ctlIssDate.GetDate))
'	IF (datediff("d",dReqDate,document.formname.ctlIssDate.GetDate)) > 0 then	
'		alert("Issue Date should be less than Requisition Date")
'		exit Function
'	End If 
	
	iTotQuantity = 0
	'alert(document.formname.elements.length)
	'alert(document.formname.hMRSNo.value)
IF document.formname.hMRSNo.value <> "" then 	

	for i=0 to document.formname.elements.length - 1
		'alert("Print="&document.formname.elements(i).type)
		if document.formname.elements(i).type = "text" then
			if (Instr(document.formname.elements(i).name,"txtQtyA") > 0) then
				iCtrr = iCtrr + 1
				iTotIssueQty = 0
				iIssueQty = 0
				iTotSTQty = 0
				iTotPRQty = 0
				'alert(document.formname.elements(i+5).value)			
				if document.formname.elements(i+3).value = "" then
					iIssueQty = 0
				else
					iIssueQty = cdbl(document.formname.elements(i+3).value)				
				end if

				if document.formname.elements(i+4).value = "" then
					iSTIssueQty = 0
				else
					iSTIssueQty = cdbl(document.formname.elements(i+4).value)					
				end if

				sRecBy = cStr(document.formname.txtRecBy.value)
				sRemarks = cStr(document.formname.Remarks.value)
				if document.formname.elements(i+5).value = "" then
					iPRIssueQty = 0
				else
					iPRIssueQty = cdbl(document.formname.elements(i+5).value)
				end if
				
				if (trim(document.formname.elements(i+3).value) <> "") then
					if(not checkNumbers(document.formname.elements(i+3).value)) then
						msgbox "Enter Numerals Only",0,"Numerals"
						document.formname.elements(i+3).select()
						exit function
					end if
					iTotIssueQty = cdbl(iTotIssueQty) + iIssueQty
				else
					iTotIssueQty = cdbl(iTotIssueQty) + 0
				end if

				if (trim(document.formname.elements(i+4).value) <> "") then
					if(not checkNumbers(document.formname.elements(i+4).value)) then
						msgbox "Enter Numerals Only",0,"Numerals"
						document.formname.elements(i+4).select()
						exit function
					end if
					iTotSTQty = cdbl(iTotSTQty) + iSTIssueQty
				else
					iTotSTQty = cdbl(iTotSTQty) + 0
				end if

				if (trim(document.formname.elements(i+5).value) <> "") then
					if(not checkNumbers(document.formname.elements(i+5).value)) then
						msgbox "Enter Numerals Only",0,"Numerals"
						document.formname.elements(i+5).select()
						exit function
					end if
					iTotPRQty = cdbl(iTotPRQty) + iPRIssueQty
				else
					iTotPRQty = cdbl(iTotPRQty) + 0
				end if

				
				arrTemp = split(document.formname.elements(i).name,"A")

				sItem = arrTemp(2)
				sClass = arrTemp(1)
				iEntNo = arrTemp(3)
	'		alert(iIssueQty &"**"&iSTIssueQty &"**"&iPRIssueQty)
				document.formname.elements(i+6).value = cdbl(iIssueQty) + cdbl(iSTIssueQty) + cdbl(iPRIssueQty)
				iTotQuantity = cdbl(iTotQuantity) + cdbl(document.formname.elements(i+6).value)
				iTotQuantity = cdbl(iTotQuantity) + cint(document.formname.elements(i+6).value)

				if (cint(document.formname.elements(i+6).value) > cint(document.formname.elements(i+1).value)) then
					msgbox "Total Quantity should be less than or equal to Quantity Pending (" & cint(document.formname.elements(i+1).value) & ")",0,"Quantity"
					document.formname.elements(i+3).select()
					exit function
				end if
				
				Set Root = OutData1.documentElement
				'alert("B4="&Root.xml)
				'alert(document.formname.elements(i-2).name)	
				For Each HeaderNode In Root.childNodes
					'if HeaderNode.Attributes.Item(0).nodeValue = sClass and HeaderNode.Attributes.Item(1).nodeValue = sItem then
					if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
						HeaderNode.setAttribute "SSTORE", trim(document.formname.elements(i-2).value)
						HeaderNode.setAttribute "REQQTY", trim(document.formname.elements(i).value)
						HeaderNode.setAttribute "REQBY", trim(sRecBy)
						HeaderNode.setAttribute "REMARKS", trim(sRemarks)
						HeaderNode.setAttribute "ITEMTYPE", trim(document.formname.hItmType.value)
						HeaderNode.setAttribute "ISSUEDATE", document.formname.ctlIssDate.GetDate
						
						if trim(document.formname.elements(i+3).value) = "" then
							HeaderNode.setAttribute "ISSQTY", "0"
						else
							HeaderNode.setAttribute "ISSQTY", trim(document.formname.elements(i+3).value)
						end if
						
						if trim(document.formname.elements(i+4).value) = "0" then
							HeaderNode.setAttribute "TRAQTY", trim(document.formname.elements(i+4).value)
						else
							HeaderNode.setAttribute "TRAQTY", trim(document.formname.elements(i+4).value)
							if (document.formname.elements(i+9).selectedIndex = "0") then
								For Each HNode In HeaderNode.childNodes
									if StrComp(Trim(HNode.NodeName),"STSchedule") = 0 then
										set a = HeaderNode.removeChild(HNode)
									end if
								next

								Set newElem = OutData1.createElement("STSchedule")
								newElem.setAttribute "STYPE", trim(document.formname.elements(i+9).value)
								newElem.setAttribute "SVALUE", date()
								HeaderNode.appendChild newElem
							end if
						end if
						
						if trim(document.formname.elements(i+5).value) = "" then
							HeaderNode.setAttribute "PRQTY", "0"
						elseif trim(document.formname.elements(i+5).value) = "0" then
							HeaderNode.setAttribute "PRQTY", "0"
						else
							HeaderNode.setAttribute "PRQTY", trim(document.formname.elements(i+5).value)
							if (document.formname.elements(i+10).selectedIndex = "0") then
								For Each HNode In HeaderNode.childNodes
									if StrComp(Trim(HNode.NodeName),"PRSchedule") = 0 then
										set a = HeaderNode.removeChild(HNode)
									end if
								next

								Set newElem = OutData1.createElement("PRSchedule")
								newElem.setAttribute "STYPE", trim(document.formname.elements(i+10).value)
								newElem.setAttribute "SVALUE", date()
								HeaderNode.appendChild newElem
							end if
						end if

						HeaderNode.setAttribute "ORGCODE", trim(document.formname.hOrgID.value)
						HeaderNode.setAttribute "IVALUE", iClosing

					end if
				Next

			end if
		end if
	next
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
ElseIF document.formname.hMRSNo.value = "" then 

	Set Root = OutData1.documentElement
'	 alert(Root.xml)
	 sPickFlag = False
	For Each HeaderNode In Root.childNodes
		IF trim(HeaderNode.NodeName) = trim("ITEMDETAILS") then
		'if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
			sTemp = cstr(HeaderNode.Attributes.Item(2).nodeValue)&"X"&cstr(HeaderNode.Attributes.Item(1).nodeValue)&"X"&cstr(HeaderNode.Attributes.Item(0).nodeValue)
			set IQ = eval("document.formname.txtQtyPPX"&sTemp)
			iIssueQty = IQ.value
			'Set SQ =  eval("document.formname.txtQtyTraX"&sTemp)
			'iSTIssueQty = SQ.value
			'Set SQ =  eval("document.formname.txtQtyPrX"&sTemp)
			'iPRIssueQty = SQ.value
			'alert(iIssueQty &"**"&iSTIssueQty &"**"&iPRIssueQty)
			'eval("document.formname.txtQtyTotX"&sTemp).value = cdbl(iIssueQty) + cdbl(iSTIssueQty) + cdbl(iPRIssueQty)
			'eval("document.formname.txtQtyTotX"&sTemp).value = cdbl(iIssueQty)
			'iTotQuantity = cdbl(iTotQuantity) + eval("document.formname.txtQtyTotX"&sTemp).value 
			iTotQuantity = iIssueQty
			'alert(HeaderNode.getAttribute("ENTRYNO") & "==" & HeaderNode.getAttribute("ITEMCODE"))		
			for each node in HeaderNode.childnodes 
				if trim(node.NodeName) = "Pick" then	
					sPickFlag = True
					set picknode = node	
				else
					sPickFlag = False
				end if
			next
			
			for each hnode in Root1.childnodes 
				for each node in hnode.childNodes
					if trim(node.NodeName) = "AddDet" then	
						sAddNodeFlag = True
						set AddDetNode = node	
					else
						sAddNodeFlag = False
					end if
				next
			next
			sRecBy = cStr(document.formname.txtRecBy.value)
			sRemarks = cStr(document.formname.Remarks.value)
			iSTIssueQty = 0
			iPRIssueQty	= 0					
			Set Elem1 = OutData.createElement("ITEM")
		 
			Elem1.setAttribute "ENTRYNO",HeaderNode.getAttribute("ENTRYNO")
			Elem1.setAttribute "ITMCODE", HeaderNode.getAttribute("ITEMCODE")
			Elem1.setAttribute "CLACODE",HeaderNode.getAttribute("CLASSCODE")			
			Elem1.setAttribute "ITMNAME", HeaderNode.getAttribute("ITEMNAME")
			Elem1.setAttribute "SSTORE", ""
			Elem1.setAttribute "REQQTY","0" 'HeaderNode.getAttribute("QTY")
			
			Elem1.setAttribute "REQBY", trim(sRecBy)
			Elem1.setAttribute "REMARKS", trim(sRemarks)
			Elem1.setAttribute "ITEMTYPE", trim(document.formname.hItmType.value)
			Elem1.setAttribute "ISSUEDATE", document.formname.ctlIssDate.GetDate 
			
			Elem1.setAttribute "ISSQTY",iIssueQty
			Elem1.setAttribute "TRAQTY", iSTIssueQty
			Elem1.setAttribute "PRQTY", iPRIssueQty
			Elem1.setAttribute "IVALUE", iTotQuantity
			Elem1.setAttribute "ORGCODE",HeaderNode.getAttribute("UNIT")
			sOrgCode = HeaderNode.getAttribute("UNIT")
			Elem1.setAttribute "MRSNO",""
			Elem1.setAttribute "MRSDATE",""
			Elem1.setAttribute "ATTRIBUTELIST",HeaderNode.getAttribute("ATTRIBUTELIST")
			Elem1.setAttribute "CREATEDBY",document.formname.hUserId.value
			Elem1.setAttribute "CREATEDON",document.formname.ctlIssDate.getDate
			if document.formname.chkReqType.checked=true then
				Elem1.setAttribute "RETURNABLE","0"
			Else
				Elem1.setAttribute "RETURNABLE","1"
			end if
			Root1.Appendchild Elem1
			
			IF sPickFlag = True then Elem1.appendchild picknode
			if sAddNodeFlag = true then Elem1.appendchild AddDetNode
		End IF 'IF trim(HeaderNode.NodeName) = trim("ITEMDETAILS") then
	Next
'alert("OutData="&OutData.xml)
'alert("OutData1="&OutData1.xml)

set Root1 = OutData.documentElement

OutData1.removechild Root
Set Root = OutData1.createElement("ISSTYPE")
OutData1.appendChild Root

if Root1.haschildnodes then 
	for each rnode in Root1.childnodes 
		if trim(rnode.nodename) = trim("ITEM") then 
			Set ItemNode = rnode
			Root.Appendchild ItemNode	
		end if
	next
end if




'document.formname.hIssForType.value = document.formname.selIssueFor.value
'alert("Total="& iTotQuantity)
End IF
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'exit function 
	Set Root = OutData.documentElement
	'alert Root.XML
	' This is the code which is used to check additional details
	if document.formname.hUsage.value = "MAT" then
		sExp1 ="//ITEM [ @CLASSCODE = "&sClass&" and @ITEMCODE = "&sItem&" and @ENTRYNO = "&iEntNo&"]/AddDet/WorkCenter/MachineCenter"
		sExp1 ="//ITEM/AddDet"
		Set MCNode = Root.Selectnodes(sExp1)
		sExp1 ="//ITEM"
		Set MCNode1 = Root.Selectnodes(sExp1)
		iCnt =  CInt(MCNode1.length)- CInt(MCNode.length)
		'if MCNode.length = 0 then
		if MCNode.length < MCNode1.length Then
			If j = 0 Then
				j = 1
				if Not confirm("Additional Detail Is Not Given For " & iCnt & " Items Do You Want To Proceed...!?") then
					Exit Function						
				end if										
			End If
		end if			
	End If
	Set Root = OutData1.documentElement

'	msgbox Root.xml
'	exit function

	'if (cdbl(iTotQuantity) = 0) then
	'	alert("Enter Quantity (Issue or Stock Transfer or PR) for any one Item of MR")
	'	exit function
	'end if

	'document.formname.B15.disabled = True
	
	
	'msgbox OutData1.xml
	'exit function
	

	Root.setAttribute "ISSTYPE",document.formname.selIssType(document.formname.selIssType.selectedIndex).value
	Root.setAttribute "ISSFORCODE", trim(document.formname.hUsage.value)
	Root.setAttribute "ISSFORTYPE", trim(document.formname.hIssForType.value)
	Root.setAttribute "PARTYCODE",trim(document.formname.hPartyCode.value)
	Root.setAttribute "POConfirm","N"
	Root.setAttribute "SInvConfirm","N"
	Root.setAttribute "Invoice","A"
	Root.setAttribute "GPConfirm","N"
	Root.setAttribute "ProConfirm","N"
	Root.setAttribute "MCallFrom","MRIssue"
	Root.setAttribute "RedirectTo","ISSUEMGMT.ASP"
	Root.setAttribute "MRNo",""
	
'	set ndRootUsage = RefData.documentElement
'	if ndRootUsage.hasChildNodes() then
'		sAppRefType = ndRootUsage.getAttribute("RefType")
'		for each ndChild in ndRootUsage.childNodes
'			if ndChild.nodeName="Reference" then
'				sAppRefNo = ndChild.getAttribute("ReferenceNo")
'				sAppRefDate = ndChild.getAttribute("ReferenceDate")
'			end if
'		next
'	end if
	
	Root.setAttribute "AppRefType",sAppRefType
	Root.setAttribute "AppRefNo",sAppRefNo
	Root.setAttribute "AppRefDate",sAppRefDate
	Root.setAttribute "ConsumptionAccHead",""
	Root.setAttribute "IssueToCode",""
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'objhttp.Open "POST","XMLSave.asp?Name=mrsIssueData", false	'
	'objhttp.send OutData1.XMLDocument							'
	'	if objhttp.responseText <> "" then						'
	'		Msgbox(objhttp.responseText)						'	
	'	else													'
	'		document.formname.ACTION="mrsIssueInsert.asp"		'
	'		document.formname.submit()							'	
	'	end if													'	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Set Root = OutData1.documentElement
'	alert(Root.xml)
	
'	exit function
	Dim sTempArr,sTempValue,sTempInvArr,sForInvNo
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	
	objhttp.Open "POST","XMLSave.asp?Name=mrsIssueData&SessionFlag=true", false
	objhttp.send OutData1.XMLDocument
	
	document.formname.action = "mrsIssueInsert.asp?CallFrom=MR"
	document.formname.submit

	
'	if trim(objhttp.responseText)<>"" then
'			alert(objhttp.responseText)
'			document.formname.B15.disabled = False	
'	else
'		if document.formname.selIssType(document.formname.selIssType.selectedIndex).value ="M" then
'			Msgbox ("MR has been Marked for Pick Issue")
'			window.location.href = "ISSUEMGMT.ASP"
'		else
'			window.location.href = "ISSUEMGMT.ASP"
'		end if
'	end if
	
'=====================================================================		
'	'window.location.href = "mrsMgmtEntry.asp?sCheck=I"

'=====================================================================


'If cstr(Trim(objhttp.responseText)) = "T" Then 
'	msgbox "Issue Has Done...!"'
'window.location.href = "mrsMgmtEntry.asp?sCheck=I"
'Else
'	'alert(objhttp.responseText)
'End If


end Function

Function setSch(sItem,sClass,objSch)
	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = sItem and HeaderNode.Attributes.Item(1).nodeValue = sClass then
				For Each PageNode In HeaderNode.childNodes
					sTemp = PageNode.Attributes.Item(0).nodeValue
				next
			end if
		next
	end if
	for ic = 0 to objSch.length
		if sTemp = objSch.options(ic).value then
			objSch.selectedIndex = ic
			exit function
		end if
	next
end Function


Function CheckSTPRSch(obj,todaysdate,sWho)
	dim sItem,sClass,a,qty
	arrTemp = split(obj.name,"Z")
	'alert(obj.name)
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)
 
	if sWho = "ST" then
		set qty = eval("document.formname.txtQtyTraX"+cstr(sClass)+"X"+cstr(sItem)+"X"+cstr(iEntNo))
	 
	elseif sWho = "PR" then
		set qty = eval("document.formname.txtQtyPrX"+cstr(sClass)+"X"+cstr(sItem)+"X"+cstr(iEntNo))
	end if


	if (trim(qty.value)="") then
		MsgBox "Enter Quantity",0,"Quantity"
		qty.focus()
		obj.selectedIndex=0
		exit function
	elseif(not checkNumbers(qty.value)) then
		msgbox "Enter Numerals Only",0,"Numerals"
		qty.focus()
		obj.selectedIndex=0
		exit function
	end if

	if cdbl(qty.value)=0 then 
		obj.selectedIndex=0
		exit function
	end if
	
	Set Root = OutData.documentElement
	For Each HeaderNode In Root.childNodes
		'if HeaderNode.Attributes.Item(0).nodeValue = sClass and HeaderNode.Attributes.Item(1).nodeValue = sItem then
		if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
			if HeaderNode.HaschildNodes() then
				For Each HNode In HeaderNode.childNodes
					if sWho = "ST" then
						if StrComp(Trim(HNode.NodeName),"STSchedule") = 0 or StrComp(Trim(HNode.NodeName),"ScheduleDetails") = 0 then
							set a = HeaderNode.removeChild(HNode)
						end if
					elseif sWho = "PR" then
						if StrComp(Trim(HNode.NodeName),"PRSchedule") = 0 or StrComp(Trim(HNode.NodeName),"ScheduleDetails") = 0 then
							set a = HeaderNode.removeChild(HNode)
						end if
					end if
				next
			end if
		end if
	Next

	if (obj.selectedIndex = "1") then
		value = prompt("Enter No of Days","0")
		if (isNull(value)) then
			obj.selectedIndex=0
			exit function
		elseif (trim(value)="") then
			obj.selectedIndex=0
			exit function
		else
			if(trim(value)="") then
				msgbox "Enter Number of Days",0,"Number of Days"
				obj.selectedIndex=0
				exit function
			else
				if(not checkNumbers(value)) then
					msgbox "Enter Numerals Only",0,"Numerals"
					obj.selectedIndex=0
					exit function
				else
					Set Root = OutData.documentElement
					For Each HeaderNode In Root.childNodes
						if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
							if sWho = "ST" then
								Set newElem = OutData.createElement("STSchedule")
							elseif sWho = "PR" then
								Set newElem = OutData.createElement("PRSchedule")
							end if
							newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
							newElem.setAttribute "SVALUE", trim(value)
							HeaderNode.appendChild newElem
						end if
					Next
				end if
			end if
		end if		
	end if
	if (obj.selectedIndex = "2") then
		value=prompt("Enter the Date","")
		if (isNull(value)) then
			obj.selectedIndex=0
			exit function
		elseif (trim(value)="") then
			objType.selectedIndex=0
			objValue.value=	""
			exit function
		else
			if (not vd(value,todaysdate)) then
				MsgBox "Invalid Date",0,"Invalid Date"
				obj.selectedIndex=0
				Exit Function
			end if
			if (DateDiff("d",todaysdate,value) < 0) then
				MsgBox "Date should be greater or equal to Today's Date",0,"Invalid Date"
				obj.selectedIndex=0
				Exit Function
			else
				Set Root = OutData.documentElement
				For Each HeaderNode In Root.childNodes
					'if HeaderNode.Attributes.Item(0).nodeValue = sClass and HeaderNode.Attributes.Item(1).nodeValue = sItem then
					if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
						if sWho = "ST" then
							Set newElem = OutData.createElement("STSchedule")
						elseif sWho = "PR" then
							Set newElem = OutData.createElement("PRSchedule")
						end if
						newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
						newElem.setAttribute "SVALUE", trim(value)
						HeaderNode.appendChild newElem
					end if
				Next
			end if
		end if
	end if
	if (obj.selectedIndex = "3") then
		Set Root = OutData.documentElement
		For Each HeaderNode In Root.childNodes
			'if HeaderNode.Attributes.Item(0).nodeValue = sClass and HeaderNode.Attributes.Item(1).nodeValue = sItem then
			if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
				if sWho = "ST" then
					Set newElem = OutData.createElement("STSchedule")
				elseif sWho = "PR" then
					Set newElem = OutData.createElement("PRSchedule")
				end if
				newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
				newElem.setAttribute "SVALUE", ""
				HeaderNode.appendChild newElem
			end if
		Next
		sTempValues = qty.value&":"&sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&iEntNo&":"&sOptName

		if sWho = "ST" then
			Set OutDataValue = showModalDialog("mrsSTSchedulePoP.asp?sTemp="&sTempValues,OutData,"dialogHeight:510px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")
		elseif sWho = "PR" then
			Set OutDataValue = showModalDialog("mrsPRSchedulePoP.asp?sTemp="&sTempValues,OutData,"dialogHeight:510px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")
		end if
		
	end if			
end Function

Function RemoveXML()
	dim Node, HeaderNode, Flag, i, obj
	
	Set Root = OutData.documentElement
	Flag = False
	
	for Each HeaderNode in Root.childNodes
		if strcomp(HeaderNode.nodename,"ITEMDETAILS") = 0 then
			for Each Node in HeaderNode.childNodes
				HeaderNode.removeChild Node
				Flag = True
			next
		end if
	next
	
	if Flag then
		for i = 0 to document.formname.elements.length - 1
			if document.formname.item(i).type = "text" then
				if instr(1,document.formname.item(i).Name,"txtQtyPPX") > 0 then
					document.formname.item(i).value = 0 
				end if
			end if
		next
	end if
	
end Function