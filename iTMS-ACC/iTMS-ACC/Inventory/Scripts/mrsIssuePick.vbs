dim objTemp,Root,newElem,ndTemp
dim iClass,iItem,iQty,i,Q,j,iarrAttributeList

Function ShowDet(sDesc)
	if sDesc = "-" then exit function
	FontFace="Verdana,8,,bold" 
	arrTemp = split(sDesc,"|")
	TopicText = "Passed For Count From : " & arrTemp(0) & vbcrlf
	TopicText = TopicText & "Passed For Count To      : " & arrTemp(1)

	document.formname.penDet.TextPopup TopicText, FontFace, 10,10,0,0

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

Function fnInit(sItem,sClass,sEntNo,sAttriList)

	iClass = sClass
	iItem = sItem
	iEntNo = sEntNo
	iarrAttributeList = sAttriList
	'alert OutData.XML
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement
	Root.setAttribute "DONE","NO"
	Set RootO = OutData.documentElement
	'alert("Root="&Root.xml)
	'alert("RootO="&RootO.xml)
	iCtr = document.formname.hCtr.value
	
	
	
	
	For Each HeaderNodeO In Root.childNodes   
		if HeaderNodeO.NodeName = "ITEMDETAILS" then
			if HeaderNodeO.Attributes.getNamedItem("CLASSCODE").Value = iClass and HeaderNodeO.Attributes.getNamedItem("ITEMCODE").Value  = iItem and HeaderNodeO.Attributes.getNamedItem("ENTRYNO").Value  = iEntNo and HeaderNodeO.Attributes.getNamedItem("ATTRIBUTELIST").Value = iarrAttributeList then
				'idItemName.innerHTML = HeaderNodeO.Attributes.Item(4).nodeValue & "&nbsp;"
				For Each childNod In HeaderNodeO.childNodes					
					if childNod.NodeName = "Pick" then
						RootO.Attributes.getNamedItem("TOT").value = childNod.Attributes.getNamedItem("TOT").value 
						For Each childNodO In childNod.childNodes
					        For Each Node in RootO.ChildNodes
					            i = Node.getAttribute("Count")
					    	    if childNodO.Attributes.getNamedItem("LOTNO").Value = Node.Attributes.getNamedItem("LOTNO").Value and childNodO.Attributes.getNamedItem("LOC").Value = Node.Attributes.getNamedItem("LOC").Value and childNodO.Attributes.getNamedItem("BIN").Value = Node.Attributes.getNamedItem("BIN").Value then
					    	        set Q = eval("document.formname.txtIssZ"&i)
							        Q.value = childNodO.Attributes.getNamedItem("QTYISS").Value
							        eval("document.formname.txtTotPackZ"&i).value=childNodO.getAttribute("NoofPack")
								end if
							next
							
					    	For Each Node in RootO.ChildNodes
					    	    if childNodO.Attributes.getNamedItem("LOTNO").Value = Node.Attributes.getNamedItem("LOTNO").Value and childNodO.Attributes.getNamedItem("LOC").Value = Node.Attributes.getNamedItem("LOC").Value and childNodO.Attributes.getNamedItem("BIN").Value = Node.Attributes.getNamedItem("BIN").Value then
								    RootO.removeChild Node
								    RootO.appendChild childNodO
							    end if
							next
						next
						HeaderNodeO.removeChild childNod
					end if
				next
			end if
		end if
	next
	
	if trim(document.formname.hType.value)="M" and trim(document.formname.hPickPackFlag.value)="N" then
	    sExp = "//Pick/PICK"
	    Set Tempnode = RootO.Selectnodes(sExp)
	    If Tempnode.Length > 0  and Tempnode.Length<=1 Then
		    if RootO.hasChildNodes() then
			    sArrTemp= split(document.formname.hTemp.value,":")
    			
			    sOptName = document.formname.hOptName.value
			    sOrgID = document.formname.hOrgID.value
			    sAttList = document.formname.hAttList.value
			    sAttID	= document.formname.hAttID.value
    			
			    CheckLot 1,iEntNo,iClass,iItem,sOrgID,"",sOptName,sAttID,sAttList
		    end if
	    End If
	end if 'if trim(document.formname.hType.value)="M" and trim(document.formname.hPickPackFlag.value)="N" then
	
	'alert(Root.xml)
	'alert "CHK="& RootO.XML
end Function

Function CheckSubmit()
	dim i,Root
	dim arrTemp
	Dim sTotNoofPack
	
	Set ndRoot = objTemp.documentElement
	ndRoot.setAttribute "DONE","YES"
	
	Set Root = OutData.documentElement
	
	'alert("PickSt="&Root.xml)
	'exit function
	
	for i=0 to document.formname.elements.length - 1

		if document.formname.elements(i).type = "text" then			
			if (Instr(document.formname.elements(i).name,"txtQtyZ") > 0) then
				if (trim(document.formname.elements(i+1).value) = "") then
					msgbox "Enter Quantity Issue" ,0,"Quantity Issue"
					document.formname.elements(i+1).select()
					exit function
				elseif(not checkNumbers(document.formname.elements(i+1).value)) then
					msgbox "Enter Numerals Only",0,"Numerals"
					document.formname.elements(i+1).select()
					exit function
				else
					sQty = document.formname.elements(i).value
					
					if (cdbl(document.formname.elements(i+1).value) > cdbl(sQty)) then
						msgbox "Quantity Issue should be equal to or less than Stock Quantity "&sQty,0,"Quantity"
						document.formname.elements(i+1).select()
						exit function
					end if
					sTotQty = cdbl(sTotQty) + cdbl(document.formname.elements(i+1).value)					
				end if
			end if
		end if 
	next
	
IF iQty <> "" then 
	if (cdbl(sTotQty) > cdbl(idQty.innerHTML)) then
		msgbox "Quantity Issue should be equal to or less than Quantity Pending "&idQty.innerHTML,0,"Quantity"
		exit function
	end if
End IF	
	iCtr = document.formname.hCtr.value
	
	if not sTotQty = 0 then
			For Each HeaderNode In Root.childNodes   		
			iCnt = HeaderNode.getattribute("Count")
				set Q = eval("document.formname.txtIssZ"&iCnt)
				HeaderNode.setAttribute "QTYISS", trim(Q.value)
				HeaderNode.setAttribute "NoofPack",eval("document.formname.txtTotPackZ"&iCnt).value
				sTotNoofPack = cint(sTotNoofPack)+cint(eval("document.formname.txtTotPackZ"&iCnt).value)
			next
		
		
		Root.Attributes.getNamedItem("TOT").Value = sTotQty
		
		sRcptNumbering = document.formname.hRcptNumbering.value

		if trim(document.formname.hType.value)="M" and trim(document.formname.hPickPackFlag.value)="N" and trim(sRcptNumbering)="LS" then
			if document.formname.radLotPack(0).checked = true then
				sFlagValue = document.formname.radLotPack(0).value
			else
				sFlagValue = document.formname.radLotPack(1).value
			end if
		End if 'if trim(document.formname.hType.value)="M" and trim(document.formname.hPickPackFlag.value)="N" and trim(sRcptNumbering)="LS" then
		
        Root.setAttribute "ONLYLOT",sFlagValue

		Set RootO = objTemp.documentElement
		'alert("AAA="&RootO.xml)
		For Each HeaderNodeO In RootO.childNodes
			 'alert iClass & ";" & iItem & ";" & iEntNo
			if HeaderNodeO.NodeName = "ITEMDETAILS" then 
				if HeaderNodeO.Attributes.getNamedItem("CLASSCODE").Value = iClass and HeaderNodeO.Attributes.getNamedItem("ITEMCODE").Value  = iItem and replace(replace(replace(HeaderNodeO.Attributes.getNamedItem("ATTRIBUTELIST").Value,Chr(39),"~~"),"#","$"),":","@") = iarrAttributeList then
					For Each childNod In HeaderNodeO.childNodes
						if StrComp(Trim(childNod.NodeName),"Pick") = 0 then
							set oRem = HeaderNodeO.removeChild(childNod)
						end if						
					next
					'	alert(trim(Q.value))
						HeaderNodeO.setAttribute "QTY",""' trim(Q.value)
					'alert sTotQty
					'alert(Root.getAttribute("TOT"))
					if not Root.Attributes.getNamedItem("TOT").Value = "" then				
						HeaderNodeO.appendChild(OutData.documentElement)
					end if
					exit for
				end if
			end if
		next
	else
		Root.Attributes.getNamedItem("TOT").Value = "0"
		Set RootO = objTemp.documentElement
		For Each HeaderNodeO In RootO.childNodes
			if HeaderNodeO.NodeName = "ITEM" then
				if HeaderNodeO.Attributes.getNamedItem("ENTRYNO").Value = iEntNo And HeaderNodeO.Attributes.getNamedItem("CLACODE").Value = iClass and HeaderNodeO.Attributes.getNamedItem("ITMCODE").Value  = iItem and replace(HeaderNodeO.Attributes.getNamedItem("ATTRIBUTELIST").Value,Chr(39),"~~") = iarrAttributeList  then
					For Each childNod In HeaderNodeO.childNodes
						if StrComp(Trim(childNod.NodeName),"Pick") = 0 then
							set oRem = HeaderNodeO.removeChild(childNod)
						end if
					next
					exit for
				end if
			end if
		next

	end if	
	Root.setAttribute "NoofPack",sTotNoofPack
	window.close()
end Function

Function window_onunload() 
	'alert objTemp.documentElement
	set window.returnValue = objTemp.documentElement
	window.close()
end Function

Function CheckLot(iCtr,iEntNo,iClass,iItem,sOrg,InvRecNo,sOptName,sAttID,sAttList)
	dim sItem,sClass,aobj,iSerialNo,sTotNoofPack,Obj
	set obj = eval("document.formname.hValueZ"&iCtr)
	 
	arrTemp = split(obj.value,":")
	iLineNo = arrTemp(1)
	sLot = arrTemp(2)
	sStore = arrTemp(3)
	sBin = arrTemp(4)
	iStQty = arrTemp(5)
	
	Set Root = OutData.documentElement 
	sTempValues = obj.value&":"&iClass&":"&iItem&":"&sOrg&":"&InvRecNo&":"&iEntNo&":"&sOptName&":"&sAttID&":"&sAttList
	'alert(Root.xml)
	'alert(sTempValues)
	
	set OutDataValue = window.showModalDialog("mrsPickDetailLotPop.asp?sTemp="&sTempValues,OutData,"dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No")
	sExp = "//Pick/PICK[@LOC = """&sStore&""" and @BIN = """&sBin&""" and @LOTNO = """&sLot&"""]/Pagination"
	Set Tempnode = Root.Selectnodes(sExp)
	'alert(Root.xml)
	If Tempnode.Length > 0 Then
	'alert(Tempnode.Item(0).Attributes.Item(0).Nodevalue)
		sTempValues = split(Tempnode.Item(0).Attributes.Item(0).Nodevalue,"```")
		while UBound(sTempValues) = 0
			sTempValues = replace(Tempnode.Item(0).Attributes.Item(0).Nodevalue,"|","&")
			set OutDataValue = window.showModalDialog("mrsPickDetailLotPop.asp?sTemp="&sTempValues,OutData,"dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No")
			sExp = "//Pick/PICK[@LOC = """&sStore&""" and @BIN = """&sBin&""" and @LOTNO = """&sLot&"""]/Pagination"
			Set Tempnode = Root.Selectnodes(sExp)
			If Tempnode.Length > 0 Then
	'		    alert(Tempnode.Item(0).Attributes.Item(0).Nodevalue)
				sTempValues = split(Tempnode.Item(0).Attributes.Item(0).Nodevalue,"```")
			end if
		wend
		
		if UBound(sTempValues)>0 then
		    if left(trim(sTempValues(1)),2)="NO" then exit function
	    end if 'if UBound(sTempValues)>0 then
		
	end if
	
	sType = document.formname.hType.value
    sPickPack = document.formname.hPickPackFlag.value
    sRcptNum = document.formname.hRcptNumbering.value
    if trim(sType) = "M" and trim(sPickPack)="N" and trim(sRcptNum)="LS" then
        set objradLot = eval("document.formname.radLotPack")
	    objradLot(1).checked = true
	end if 


	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
			set Q = eval("document.formname.txtIssZ"&iLineNo)
			Q.value = HeaderNode.Attributes.getNamedItem("QTYISS").Value
			if trim(Q.value)="" or isnull(Q.value) then Q.value="0"
			sNoofPack = Root.getAttribute("NoofPack")
			if trim(sNoofPack)<>"" and trim(sNoofPack)<>"0" then
				if trim(Q.value)<>"0" then
					eval("document.formname.txtTotPackZ"&iLineNo).value = 	Root.getAttribute("NoofPack")
				end if
			else
				eval("document.formname.txtTotPackZ"&iLineNo).value = 	"0"
			end if
			exit function
		end if
	Next
	
end Function