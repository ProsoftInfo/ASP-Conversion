dim OutDataValue,Root

Set Root = OutData2.documentElement
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
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)
	sLoc = arrTemp(5)
	sBin = arrTemp(6)

	sTempValues = sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&iEntNo&":"&sOptName&":"&sLoc&":"&sBin

	showModalDialog "itmStockPoP.asp?sTemp="&sTempValues,OutData2,"dialogHeight:450px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No"
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
							Q.value = "0"
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
	set OutDataValue = showModalDialog("IssueAddDetails.asp?sTemp="&sTempValues,OutData2,"dialogHeight:400px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No")		
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

	set OutDataValue = showModalDialog("mrsIssueScheduleEntryPoP.asp?sTemp="&sTempValues,OutData2,"dialogHeight:490px;dialogWidth:330px;center:Yes;help:No;resizable:No;status:No")

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

	OutDataValue = showModalDialog("mrsIssueQtyParaPoP.asp?sTemp="&sTempValues,OutData2,"dialogHeight:385px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No")
end Function

Function DisplayItem(obj)
	sArrValues = Split(obj,"A")
	sTempValues = sArrValues(0)&"A"&sArrValues(2)&"A"&sArrValues(1)&"A"&sArrValues(3)
	showModalDialog "itmDetailsPop.asp?sTemp="&sTempValues,"","dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No"
end Function

Function CheckLot(obj,iQty,sAttributeList)
	dim sItem,sClass,sLoc,sBin,sType,sOrgID,sAttList,sArrList,sAttID,sTempAttribute
	'alert(sAttributeList)
	arrTemp = split(obj.name,"Z")
	'alert(obj.name)
	sItem = arrTemp(2)
	sClass = arrTemp(1)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)
	sLoc = arrTemp(3)
	sBin = arrTemp(4)
	sOrgID = arrTemp(5)
	'sAttList = arrTemp(6)
	sTempAttribute = sAttributeList 
	sAttributeList = Replace(sAttributeList,"#","$")
	
	sArrAttTemp = Split(sAttributeList,",")
	
	    For iCnt = 0 to UBound(sArrAttTemp)
	        sArrAttSub1 = Split(sArrAttTemp(iCnt),"@")
	        sArrAttribute = split(sArrAttSub1(0),"$")
	        if UBound(sArrAttSub1)>0 then
	            if Trim(sArrAttribute(1))<>"0" then
	                sAttList = sAttList &","& sArrAttribute(0)
	                sAttID = sAttID &","& sArrAttribute(1)
	            end if 'if Trim(sArrAttribute(1))<>"0" then    
	        else
	            sAttList =sAttList &","& sArrAttSub1(0)
    	    end if 'if UBound(sArrAttTemp)>0 then
	    Next
    	if trim(sAttList)<>"" then
		    sAttList = Mid(sAttList,2)
		    sAttID = mid(sAttID,2)
	    else
		    sAttList = ""
		    sAttID = ""
	    end if ' if trim(arrTemp(6))<>"" then
'	alert(sAttList)
'	alert(sAttID)
	sPackFlag = document.formname.hPickPackFlag.value 
	if document.formname.selIssType.checked = true then
		sType = "M"
	else
		sType = "F"
	end if
	
	sTempValues = sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&iEntNo&":"&sOptName&":"&sLoc&":"&sBin&":"&iQty &":"&sType&":"&sUsage&":"& sOrgID&":"& sAttList&":"&sAttID
	
	set OutDataValue = showModalDialog("mrsPickDetailPoP.asp?sTemp="&sTempValues&"&AttributeList="&sAttributeList&"&PickPackFlag="&sPackFlag,OutData1,"dialogHeight:350px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
	Set Root = OutData1.documentElement
	if Root.getAttribute("DONE")="NO" then exit function
	
	For Each HeaderNodeO In Root.childNodes
		if HeaderNodeO.NodeName = "ITEMDETAILS" then
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
							'Q.focus()
						else
							Q.readOnly = false
						end if
						HeaderNodeO.setAttribute "ONLYLOT",childNod.getAttribute("ONLYLOT")
						exit function
					end if
				next
				if bFlag then
					sTemp = cstr(HeaderNodeO.Attributes.Item(2).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(1).nodeValue)&"X"&cstr(HeaderNodeO.Attributes.Item(0).nodeValue)
					set Q = eval("document.formname.txtQtyPPX"&sTemp)
					Q.value = "0"
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

	Set OutDataValue = showModalDialog("mrsIssueSTPoP.asp?sTemp="&sTempValues,OutData2,"dialogHeight:330px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")

	for i=0 to document.formname.elements.length - 1
		if (document.formname.elements(i).type = "text") then
			'arrTemp = split(document.formname.elements(i).name,"A")
			'sItem = arrTemp(2)
			'sClass = arrTemp(1)
			Set RootO = OutData2.documentElement
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
	dim arrTemp,AddDetNode,subNode,sInvoiceSelected,sArrList,sAttributeID
	dim iReturableCnt,iPrimaryCnt
	
	set Root1 = OutData2.documentElement
	IF Root1.haschildnodes then 
		For each node in Root1.childnodes 
			If trim(node.NodeName) = "ITEM" then
				Root1.RemoveChild node
			End If 
		Next
	End IF	
	'alert(Root1.xml)
	if document.formname.hAutoConsumption.value="Y" then
	    if document.formname.selAccHead.selectedIndex = 0 or document.formname.selAccHead.selectedIndex = 1 then
	        alert("Select the Account Head")
	        document.formname.selAccHead.focus
	        exit function
	    end if
	end if
	 
	If (datediff("d",todaysdate,document.formname.ctlIssDate.GetDate)) > 0 then
		alert("Issue Date should be less than or equal to Today's Date")
		exit function
	End If
	dReqDate = document.formname.hReqDate.value
	iTotQuantity = 0
	'alert(document.formname.elements.length)
	
	Set Root = OutData1.documentElement
	
	if trim(document.formname.hType.value)="SUB" then
		sIssToCode = document.formname.hIssueToCode.value
		if trim(sIssToCode)=""  or IsNull(sIssToCode) then
			alert("Please Select Issued To")
			exit function
		end if
		if Root.hasChildNodes() then
		    bProcessFlag = false
		    for each ndNode in Root.childNodes
		        if trim(ndNode.nodeName)=trim("SubContract") then
		            bProcessFlag=true
		        end if
		    next
		end if
		if bProcessFlag=false then
		    alert("Please select the Process Details")
		    exit function
		end if 
	end if 
	
	
	
	iReturableCnt = 0
	iPrimaryCnt = 0
	For Each HeaderNode In Root.childNodes
		IF trim(HeaderNode.NodeName) = trim("ITEMDETAILS") then
		    sItem = HeaderNode.getAttribute("ITEMCODE")
		    sClass = HeaderNode.getAttribute("CLASSCODE")
		    iEntNo =  HeaderNode.getAttribute("ENTRYNO")
		    sItemName = HeaderNode.getAttribute("ITEMNAME")
		    sMatType = ""
		    
		    set objIssQty = eval("document.formname.txtQtyPPX"&trim(sClass)&"X"&trim(sItem)&"X"&trim(iEntNo))
		    
		    iIssQty = objIssQty.value
		    if trim(iIssQty)="" or IsNull(iIssQty) then iIssQty = "0"
		    
		    if cdbl(iIssQty)<=0 then 
		        alert("Issue Quantity Should be Greater then Zero for "& sItemName)
		        exit function
		    end if
		    set objRetType = eval("document.formname.selReturnZ"&trim(sItem)&"Z"&trim(sClass)&"Z"&trim(iEntNo))
		    if objRetType.selectedIndex>0 then
			    iReturableCnt = iReturableCnt + 1
			end if
		End IF
	Next
	
	if trim(document.formname.hType.value)="SUB" then
	    if iReturableCnt <=0 then
	        alert("SubContract Issue Should have minimum One Returnable Item")
	        exit function
	    end if
	end if 
	
	sIssMode =document.formname.hIssMode.value
	sIssEntNo = document.formname.hIssEntryNo.value
	
	
	
'	alert(Root.xml)
'	alert(OutData2.xml)
	 sPickFlag = False
	 
	For Each HeaderNode In Root.childNodes
		IF trim(HeaderNode.NodeName) = trim("ITEMDETAILS") then
		'if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
		
			sTemp = cstr(HeaderNode.Attributes.Item(2).nodeValue)&"X"&cstr(HeaderNode.Attributes.Item(1).nodeValue)&"X"&cstr(HeaderNode.Attributes.Item(0).nodeValue)
			set IQ = eval("document.formname.txtQtyPPX"&sTemp)
			iIssueQty = IQ.value
			sTemp = cstr(HeaderNode.getAttribute("ITEMCODE"))&"Z"&cstr(HeaderNode.getAttribute("CLASSCODE"))&"Z"&cstr(HeaderNode.Attributes.Item(0).nodeValue)
			iRequested = Eval("document.formname.txtRequestedZ"&sTemp).value
			iToIssue = Eval("document.formname.txtToIssueZ"&sTemp).value
			sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
			
			sIssMode = document.formname.hIssMode.value
			if Trim(sRefType)="11" then 
			    if Trim(sIssMode)<>"E" then
			        if cdbl(iIssueQty) > (CDbl(iRequested)-cdbl(iToIssue)) then
			            alert("Issue Quantity Must be Less than or Equal to Request Quantity")
			            exit function
			        end if 
			    end if 'if Trim(sIssMode)<>"E" then
			end if
			
			iItemCode = HeaderNode.getAttribute("ITEMCODE")
			iClassCode = HeaderNode.getAttribute("CLASSCODE")			
			iEntNo = HeaderNode.getAttribute("ENTRYNO")	
			
			iTotQuantity = iIssueQty
			for each node in HeaderNode.childnodes 
			'alert(node.xml)
				if trim(node.NodeName) = "Pick" then	
					sPickFlag = True
					set picknode = node	
				else
					sPickFlag = False
				end if
			next
			for each hnode in Root1.childnodes 
			    if trim(hnode.getAttribute("ENTRYNO"))=iEntNo and trim(hnode.getAttribute("ITEMCODE"))=iItemcode and trim(hnode.getAttribute("CLASSCODE"))=iClassCode then
				    for each node in hnode.childNodes
					    if trim(node.NodeName) = "AddDet" then	
						    sAddNodeFlag = True
						    set AddDetNode = node	
					    else
						    sAddNodeFlag = False
					    end if
				    next
			    end if 'if trim(hnode.getAttribute("ENTRYNO"))=iEntNo and trim(hnode.getAttribute("ITEMCODE"))=iItemcode and trim(hnode.getAttribute("CLASSCODE"))=iClassCode then
			next
			sArrList = split(trim(HeaderNode.getAttribute("ATTRIBUTELIST")),"#")
			if uBound(sArrList)>0 then
				if trim(sArrList(1))<>"0" and trim(sArrList(1))<>"" then
					sAttributeID = sArrList(1)
				elseif trim(sArrList(0))<>"0" then
					sAttributeID = sArrList(0)
				end if
			elseif ubound(sArrList)=0 then
				if trim(sArrList(0))<>"0" then
					sAttributeID = sArrList(0)
				end if
			end if
			
			sRecBy = cStr(document.formname.txtRecBy.value)
			sRemarks = cStr(document.formname.Remarks.value)
			iSTIssueQty = 0
			iPRIssueQty	= 0					
			
			iItemCode = HeaderNode.getAttribute("ITEMCODE")
			iClassCode = HeaderNode.getAttribute("CLASSCODE")			
			iEntNo = HeaderNode.getAttribute("ENTRYNO")			
			Set Elem1 = OutData2.createElement("ITEM")
		 
			Elem1.setAttribute "ENTRYNO",iEntNo
			Elem1.setAttribute "ITMCODE",iItemCode
			Elem1.setAttribute "CLACODE",iClassCode
			Elem1.setAttribute "ITMNAME", HeaderNode.getAttribute("ITEMNAME")
			Elem1.setAttribute "SSTORE", ""
			Elem1.setAttribute "REQQTY","0" 'HeaderNode.getAttribute("QTY")
			
			Elem1.setAttribute "REQBY", trim(sRecBy)
			Elem1.setAttribute "REMARKS", trim(sRemarks)
			Elem1.setAttribute "ITEMTYPE", trim(document.formname.hItemType.value)
			Elem1.setAttribute "ISSUEDATE", document.formname.ctlIssDate.GetDate 
			Elem1.setAttribute "ISSQTY",iIssueQty
			Elem1.setAttribute "TRAQTY", iSTIssueQty
			Elem1.setAttribute "PRQTY", iPRIssueQty
			Elem1.setAttribute "IVALUE", iTotQuantity
			Elem1.setAttribute "ORGCODE",HeaderNode.getAttribute("UNIT")
			sOrgCode = HeaderNode.getAttribute("UNIT")
			if trim(sRefType)="11" then
				Elem1.setAttribute "MRSNO",document.formname.hRefNo.value
				Elem1.setAttribute "MRSDATE",document.formname.hRefDate.value
			else
				Elem1.setAttribute "MRSNO",""
				Elem1.setAttribute "MRSDATE",""
			end if
			Elem1.setAttribute "ATTRIBUTELIST",sAttributeID
			Elem1.setAttribute "CREATEDBY",document.formname.hUserId.value
			Elem1.setAttribute "CREATEDON",document.formname.ctlIssDate.getDate
			Root1.Appendchild Elem1
			Elem1.setAttribute "RefNo",HeaderNode.getAttribute("RefNo")
			
			Elem1.setAttribute "ONLYLOT",HeaderNode.getAttribute("ONLYLOT")
			set objRetType = eval("document.formname.selReturnZ"&trim(iItemCode)&"Z"&trim(iClassCode)&"Z"&trim(iEntNo))
			
		    sRetType = objRetType(objRetType.selectedIndex).value
		    if trim(sRetType)="N" then
			    Elem1.setAttribute "RETURNABLE","N"
				Elem1.setAttribute "RETURNITEM","S"
			elseif trim(sRetType)="Y" then
			    Elem1.setAttribute "RETURNABLE","Y"
				Elem1.setAttribute "RETURNITEM","S"
			else
				Elem1.setAttribute "RETURNABLE","Y"
				Elem1.setAttribute "RETURNITEM","D"
			end if
			
			sLotOrPackFlag = HeaderNode.getAttribute("ONLYLOT")
			
			sItemRefNo = HeaderNode.getAttribute("RefNo")
			sItemCode = HeaderNode.getAttribute("ITEMCODE")
			sClassCode = HeaderNode.getAttribute("CLASSCODE")
			sArrItemRefNo = split(sItemRefNo,",")
			
			if UBound(sArrItemRefNo)>0 then
			 set OutValue = showmodaldialog("IssGetQtyForMixCodes.asp?ItemCode="&sItemCode&"&ClassCode="&sClassCode&"&MixCodes="&sItemRefNo&"&TotQty="&iissueQty,"","dialogWidth:300px;dialogHeight:300px;Status:No")
			    while OutValue.getAttribute("Action")<>"Done"
			       alert("Please Enter the Mix Quantity Details")
			       OutValue = showmodaldialog("IssGetQtyForMixCodes.asp?ItemCode="&sItemCode&"&ClassCode="&sClassCode&"&MixCodes="&sItemRefNo&"&TotQty="&iissueQty,"","dialogWidth:300px;dialogHeight:300px;Status:No")
			    wend
    			
			    if OutValue.hasChildNodes() then
			        Elem1.appendChild(OutValue)
			    end if
			end if
			
			IF sPickFlag = True then Elem1.appendchild picknode
			if sAddNodeFlag = true then Elem1.appendchild AddDetNode
			Elem1.setAttribute "MatType",""
'			if trim(document.formname.hType.value)="SUB" then
'			    set ObjMatType = eval("document.formname.radMatTypeZ"&iClassCode&"A"&iItemCode&"A"&iEntNo)
'			    if ObjMatType(0).checked then 
'			        sMatType = ObjMatType(0).value
'			    elseif ObjMatType(1).checked then 
'			        sMatType = ObjMatType(1).value
'			    elseif ObjMatType(2).checked then 
'			        sMatType = ObjMatType(2).value
'			    end if 
'				Elem1.setAttribute "MatType",sMatType
'				if trim(sMatType)="P" then
'				    iPrimaryCnt = iPrimaryCnt + 1
'				end if
'			end if 
		Elseif trim(HeaderNode.NodeName)=trim("SubContract") then
			Root1.appendChild HeaderNode
		End IF 'IF trim(HeaderNode.NodeName) = trim("ITEMDETAILS") then
	Next
	
	set Root1 = OutData2.documentElement

	OutData1.removechild Root
	Set Root = OutData1.createElement("ISSTYPE")
	OutData1.appendChild Root
	if Root1.haschildnodes then 
		for each rnode in Root1.childnodes 
			if trim(rnode.nodename) = trim("ITEM") then 
				Set ItemNode = rnode
				Root.Appendchild ItemNode	
			elseif trim(rnode.nodename)=trim("SubContract") then
			    set ProcessNode = rnode
			    Root.AppendChild ProcessNode
			end if
		next
	end if
	Set Root = OutData2.documentElement
	' This is the code which is used to check additional details
	if document.formname.hIssueToCode.value = "MAT" then
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
	if Root.hasChildNodes() then
		for each ndItem in Root.childNodes
			if ndItem.nodeName=trim("ITEM") then
				sItem = ndItem.getAttribute("ITMCODE")
				sClass = ndItem.getAttribute("CLACODE")
				sIssQty = ndItem.getAttribute("ISSQTY")
				iEntNo = ndItem.getAttribute("ENTRYNO")
				
				set objhttp = CreateObject("Microsoft.XMLHTTP")
				objhttp.open "GET","../../Common/GetItemRcptNumbering.asp?ItemCode="&sItem,false
				objhttp.send 
				if trim(objhttp.responsetext)<>"" then
					sReceiptNumbering = objhttp.responsetext
				end if
				if not ndItem.hasChildNodes() then
					if trim(sReceiptNumbering)="N" then
					    if ndItem.hasChildNodes() then
					        if ndItem.getAttribute("ITMCODE")=sItem and ndItem.getAttribute("CLACODE")=sClass and ndItem.getAttribute("ENTRYNO")=iEntNo then
					            for each ndPick in ndItem.childNodes
					                if trim(ndPick.nodeName)="Pick" then
					                    ndItem.removeChild ndPick
					                    exit for
					                end if
					            next
					        end if 'if ndItem.getAttribute("ITMCODE")=sItem and ndItem.getAttribute("CLACODE")=sClass and ndItem.getAttribute("ENTRYNO)=iEntNo then
					    end if
				        set ndPick = OutData1.createElement("Pick")
				            ndPick.setAttribute "TOT",sIssQty
					        ndPick.setAttribute "NoofPack","0"
					    ndItem.appendChild ndPick
					    
						set objhttp = CreateObject("Microsoft.XMLHTTP")
						objhttp.open "GET","../../Common/GetStoreDetailsForItem.asp?ItemCode="&sItem&"&ClassCode="&sClass,false
						objhttp.send 
						if trim(objhttp.responseXML.xml)<>"" then
							StoreDetails.loadXML(objhttp.responseXML.xml)
						end if
						set ndRootStore = StoreDetails.documentElement
						if ndRootStore.hasChildNodes() then
							if trim(ndRootStore.childNodes.length)="1" then
								for each ndStore in ndRootStore.childNodes
									set ndPickDet = OutData1.createElement("STORE")
										ndPickDet.setAttribute "LOC",ndStore.getAttribute("LocNo")
										ndPickDet.setAttribute "BIN",ndStore.getAttribute("BinNo")
										ndPickDet.setAttribute "LOTNO","N/A"
										ndPickDet.setAttribute "INVRECNO",""
										ndPickDet.setAttribute "QTYISS",sIssQty
										ndPickDet.setAttribute "NoofPack","0"
									ndPick.appendChild ndPickDet
								next
							else
								alert("Multiple Stores available please select the Issue Quantity for specify store")
								document.formname.submit
							end if 'if trim(ndRootStore.childNodes.length)=1" then
						end if 'if ndRootStore.hasChildNodes() then
					else
						set ndPick = OutData1.createElement("Pick")
						ndPick.setAttribute "TOT",sIssQty
						ndPick.setAttribute "NoofPack","0"
						ndItem.appendChild ndPick
						
						set objhttp = CreateObject("Microsoft.XMLHTTP")
						objhttp.open "GET","../../Common/GetStoreDetailsForItem.asp?ItemCode="&sItem&"&ClassCode="&sClass,false
						objhttp.send 
						if trim(objhttp.responseXML.xml)<>"" then
							StoreDetails.loadXML(objhttp.responseXML.xml)
						end if
						set ndRootStore = StoreDetails.documentElement
						if ndRootStore.hasChildNodes() then
							if trim(ndRootStore.childNodes.length)="1" then
								for each ndStore in ndRootStore.childNodes
									set ndPickDet = OutData1.createElement("PICK")
										ndPickDet.setAttribute "LOC",ndStore.getAttribute("LocNo")
										ndPickDet.setAttribute "BIN",ndStore.getAttribute("BinNo")
										ndPickDet.setAttribute "LOTNO","N/A"
										ndPickDet.setAttribute "INVRECNO",""
										ndPickDet.setAttribute "QTYISS",sIssQty
										ndPickDet.setAttribute "NoofPack","0"
									ndPick.appendChild ndPickDet
								next
							else
								alert("Multiple Stores available please select the Issue Quantity for specify store")
								document.formname.submit
							end if 'if trim(ndRootStore.childNodes.length)=1" then
						end if 'if ndRootStore.hasChildNodes() then
						
					end if 'if trim(sReceiptNumbering)="N" then
				else 'if not ndItem.hasChildNodes() then
				    if trim(sIssMode)="E" then
				        if trim(sReceiptNumbering)="N" then
				            if ndItem.hasChildNodes() then
				                if ndItem.getAttribute("ITMCODE")=sItem and ndItem.getAttribute("CLACODE")=sClass and ndItem.getAttribute("ENTRYNO")=iEntNo then
					                for each ndPick in ndItem.childNodes
					                    if trim(ndPick.nodeName)="Pick" then
					                        ndItem.removeChild ndPick
					                        exit for
					                    end if
					                next
					            end if 'if ndItem.getAttribute("ITMCODE")=sItem and ndItem.getAttribute("CLACODE")=sClass and ndItem.getAttribute("ENTRYNO)=iEntNo then
				            end if
				            
				            set ndPick = OutData1.createElement("Pick")
						    ndPick.setAttribute "TOT",sIssQty
						    ndPick.setAttribute "NoofPack","0"
						    ndItem.appendChild ndPick
				            
				            set objhttp = CreateObject("Microsoft.XMLHTTP")
						    objhttp.open "GET","../../Common/GetStoreDetailsForItem.asp?ItemCode="&sItem&"&ClassCode="&sClass,false
						    objhttp.send 
						    if trim(objhttp.responseXML.xml)<>"" then
							    StoreDetails.loadXML(objhttp.responseXML.xml)
						    end if
						    set ndRootStore = StoreDetails.documentElement
						    if ndRootStore.hasChildNodes() then
							    if trim(ndRootStore.childNodes.length)="1" then
								    for each ndStore in ndRootStore.childNodes
									    set ndPickDet = OutData1.createElement("STORE")
										    ndPickDet.setAttribute "LOC",ndStore.getAttribute("LocNo")
										    ndPickDet.setAttribute "BIN",ndStore.getAttribute("BinNo")
										    ndPickDet.setAttribute "LOTNO","N/A"
										    ndPickDet.setAttribute "INVRECNO",""
										    ndPickDet.setAttribute "QTYISS",sIssQty
										    ndPickDet.setAttribute "NoofPack","0"
									    ndPick.appendChild ndPickDet
								    next
							    else
								    alert("Multiple Stores available please select the Issue Quantity for specify store")
								    document.formname.submit
							    end if 'if trim(ndRootStore.childNodes.length)=1" then
						    end if 'if ndRootStore.hasChildNodes() then
				        end if 'if trim(sReceiptNumbering)="N" then
				    end if  'if trim(sIssMode)="E" then
				end if ' if not ndItem.hasChildNodes() then
			end if
		next
	end if
	
	
	    
	if document.formname.selIssType.checked = true then
		sIssType = "M"
	else
		sIssType = "F"
	end if
	
	Root.setAttribute "ISSTYPE",sIssType
	Root.setAttribute "ISSTOTYPE", trim(document.formname.hIssueToType.value)
	Root.setAttribute "ISSTOCODE", trim(document.formname.hIssueToCode.value)
	Root.setAttribute "ISSTOSUBCODE", trim(document.formname.hIssueToSubCode.value)
	Root.setAttribute "POConfirm","N"
	Root.setAttribute "SInvConfirm","N"
	Root.setAttribute "Invoice","A"
	Root.setAttribute "GPConfirm","N"
	Root.setAttribute "ProConfirm","N"
	Root.setAttribute "MCallFrom","MRIssue"
	Root.setAttribute "RedirectTo","ISSUEMGMT.ASP"
	sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
	if trim(sRefType)<>"N" then
	    Root.setAttribute "AppRefType",sRefType
	else
	    Root.setAttribute "AppRefType",""
	end if
	Root.setAttribute "AppRefNo",document.formname.hRefNo.value
	Root.setAttribute "AppRefDate",document.formname.hRefDate.value
	Root.setAttribute "ConsumptionAccHead",document.formname.selAccHead(document.formname.selAccHead.selectedIndex).value
	Root.setAttribute "IssueToCode",document.formname.hIssueToCode.value
	Root.setAttribute "PickPackFlag",document.formname.hPickPackFlag.value
	Root.setAttribute "IssFrom",document.formname.hIssFrom.value
	Root.setAttribute "Returnable","N"
	Root.setAttribute "ReturnItem","S"
	Root.setAttribute "TYPE",document.formname.hType.value
	Set Root = OutData1.documentElement
		if trim(document.formname.hType.value)="SUB" then
		    set outValue = showModalDialog("IssSubConPurDetPop.asp?sHead=Subcontract Order Creation&OrgCode="&sOrgCode,POrder,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			POrder.LoadXML(outValue.xml)
			'alert(POrder.xml)
			set RootSub = POrder.documentElement
			if RootSub.hasChildNodes() then
				for each subNode in RootSub.childNodes
					if strcomp(subNode.nodeName,"PURACC")=0 then
						Root.appendChild subNode
					end if
				next
			end if 'if RootSub.hasChildNodes() then
			
		'	set outValue = showModalDialog("IssSubConSalnvPop.asp?sHead=Proforma Invoice Creation&OrgCode="&sOrgCode,SalesInvoice,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		'	SalesInvoice.LoadXML(outValue.xml)
		'	
		'	set RootSub = SalesInvoice.documentElement
		'	if RootSub.hasChildNodes() then
		'		for each subNode in RootSub.childNodes
		'			if strcomp(subNode.nodeName,"SALINV")=0 then
		'				Root.appendChild subNode
		'			end if
		'		next
		'	end if
			
			
			set OutValue = ShowModalDialog("SubContOrdConfirmPop.asp?ItemType="&trim(document.formname.hItemType.value),"","dialogHeight:400px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No")
			sPOConfirm = OutValue.getAttribute("Done")
				
		'    if trim(sPOConfirm)="Y" then
		'	    set outValue = showModalDialog("CommonConfirmPop.asp?sHead=Proforma Invoice Creation&CallFrom=SUB","","dialogHeight:180px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No")
		'	    ConfData.LoadXML(outValue.xml)
		'	    ProformaConfirm = outValue.getAttribute("ProInv")
		'	else
			    ProformaConfirm ="N"
		'	end if
			
			set RootSub = ConfData.documentElement
			Root.setAttribute "Invoice","P"
			Root.setAttribute "POConfirm",sPOConfirm
			Root.setAttribute "ProConfirm",ProformaConfirm
			
		elseif trim(document.formname.hIssueToCode.value)="DIS" then
			
			if trim(sReceiptNumbering)="S" then
				sLotOrPackFlag = "P"
			end if
			
			sTemp = sIssType &":"& document.formname.hPickPackFlag.value  &":"& sLotOrPackFlag
			
			set outValue = showModalDialog("CommonConfirmPop.asp?sHead=Sales Invoice Creation&CallFrom=DIS&Issue="&sTemp,"","dialogHeight:200px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No")
			'alert(outValue.xml)
			SalesInvoice.LoadXML(outValue.xml)
			
			set RootSub = SalesInvoice.documentElement
			Root.setAttribute "SInvConfirm",RootSub.getAttribute("Confirm")
			Root.setAttribute "Invoice",RootSub.getAttribute("Invoice")
			sInvoiceSelected = RootSub.getAttribute("Invoice")
			sPOConfirm = RootSub.getAttribute("Confirm")
			
			
			set outValue = showModalDialog("IssSubConSalnvPop.asp?sHead=Sales Invoice Creation&OrgCode="&sOrgCode,SalesInvoice,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			SalesInvoice.LoadXML(outValue.xml)
			set RootSub = SalesInvoice.documentElement
			do while trim(RootSub.getAttribute("Confirm"))="N"
			    alert("Invoice Type and Type of Sale Must be Select")
			    set outValue = showModalDialog("IssSubConSalnvPop.asp?sHead=Sales Invoice Creation&OrgCode="&sOrgCode,SalesInvoice,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			    SalesInvoice.LoadXML(outValue.xml)
    		    set RootSub = SalesInvoice.documentElement
			loop
			'alert(RootSub.xml)
		    if RootSub.hasChildNodes() then
				for each subNode in RootSub.childNodes
					if strcomp(subNode.nodeName,"SALINV")=0 then
						Root.appendChild subNode
					end if
				next
			end if
		elseif trim(document.formname.hIssueToCode.value)="SER" or trim(document.formname.hIssueToCode.value)="JWK" then
			set outValue = showModalDialog("CommonConfirmPop.asp?sHead=Gate Pass Creation&CallFrom=SER","","dialogHeight:180px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No")
			'alert(outValue.xml)
			GatePass.LoadXML(outValue.xml)
			
			set RootSub = GatePass.documentElement
			Root.setAttribute "Confirm","Y"
			Root.setAttribute "POConfirm","N"
			Root.setAttribute "SInvConfirm","N"
			Root.setAttribute "GPConfirm",RootSub.getAttribute("Confirm")
			sPOConfirm = RootSub.getAttribute("Confirm")
			
			set ndServices = GatePass.createElement("SERVICES")
			ndServices.setAttribute "Transport",""
			ndServices.setAttribute "TakenBy",""
			ndServices.setAttribute "DelivertyBy",""
			ndServices.setAttribute "Remarks",""
			Root.appendChild ndServices
		end if
	

   
	
	'exit function
	Dim sTempArr,sTempValue,sTempInvArr,sForInvNo
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	if trim(sIssMode)="E" then
	    objhttp.Open "POST","XMLSave.asp?Name=mrsIssueDataEdit&SessionFlag=true", false
	else
	    objhttp.Open "POST","XMLSave.asp?Name=mrsIssueData&SessionFlag=true", false
	end if
	objhttp.send OutData1.XMLDocument
	sCallFrom = document.formname.hCallFrom.value
	if trim(sIssMode)="E" then
	    sFileName = "mrsIssueUpdate.asp?hCallFrom="&sCallFrom&"&IssEntNo="&sIssEntNo
    else
        sFileName = "mrsIssueInsert.asp?hCallFrom="&sCallFrom
        
    end if 
	document.formname.action = sFileName
	document.formname.submit

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
	
	Set Root = OutData2.documentElement
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
					Set Root = OutData2.documentElement
					For Each HeaderNode In Root.childNodes
						if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
							if sWho = "ST" then
								Set newElem = OutData2.createElement("STSchedule")
							elseif sWho = "PR" then
								Set newElem = OutData2.createElement("PRSchedule")
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
				Set Root = OutData2.documentElement
				For Each HeaderNode In Root.childNodes
					'if HeaderNode.Attributes.Item(0).nodeValue = sClass and HeaderNode.Attributes.Item(1).nodeValue = sItem then
					if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
						if sWho = "ST" then
							Set newElem = OutData2.createElement("STSchedule")
						elseif sWho = "PR" then
							Set newElem = OutData2.createElement("PRSchedule")
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
		Set Root = OutData2.documentElement
		For Each HeaderNode In Root.childNodes
			'if HeaderNode.Attributes.Item(0).nodeValue = sClass and HeaderNode.Attributes.Item(1).nodeValue = sItem then
			if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem and HeaderNode.Attributes.Item(2).nodeValue = sClass then
				if sWho = "ST" then
					Set newElem = OutData2.createElement("STSchedule")
				elseif sWho = "PR" then
					Set newElem = OutData2.createElement("PRSchedule")
				end if
				newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
				newElem.setAttribute "SVALUE", ""
				HeaderNode.appendChild newElem
			end if
		Next
		sTempValues = qty.value&":"&sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&iEntNo&":"&sOptName

		if sWho = "ST" then
			Set OutDataValue = showModalDialog("mrsSTSchedulePoP.asp?sTemp="&sTempValues,OutData2,"dialogHeight:510px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")
		elseif sWho = "PR" then
			Set OutDataValue = showModalDialog("mrsPRSchedulePoP.asp?sTemp="&sTempValues,OutData2,"dialogHeight:510px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")
		end if
		
	end if			
end Function

Function RemoveXML()
	dim Node, HeaderNode, Flag, i, obj
	
	Set Root = OutData2.documentElement
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


