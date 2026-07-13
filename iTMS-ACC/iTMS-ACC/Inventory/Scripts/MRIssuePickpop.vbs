dim objTemp,Root,newElem
dim iClass,iItem,iQty,i,Q,j,iEntNo

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

Function fnInit(sItem,sClass,sEntNo)

	iClass = sClass
	iItem = sItem
	iEntNo = sEntNo
	'alert OutData.XML
	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement

	Set RootO = OutData.documentElement
	'alert("Root="&Root.xml)
	'alert("RootO="&RootO.xml)
	iCtr = document.formname.hCtr.value
	
	For Each HeaderNodeO In Root.childNodes   
		if HeaderNodeO.NodeName = "ITEM" then
			if HeaderNodeO.getAttribute("CLACODE") = iClass and HeaderNodeO.getAttribute("ITMCODE") = iItem and HeaderNodeO.getAttribute("ENTRYNO") = iEntNo  then
				For Each childNod In HeaderNodeO.childNodes					
					if childNod.NodeName = "Pick" then
						RootO.Attributes.getNamedItem("TOT").value = childNod.Attributes.getNamedItem("TOT").value 
						For Each childNodO In childNod.childNodes
							'alert(childNodO.Attributes.getNamedItem("QTYISS").Value)
							if childNodO.Attributes.getNamedItem("QTYISS").Value <> "" then
								i = i + 1
								set Q = eval("document.formname.txtIss"&i)
								Q.value = childNodO.Attributes.getNamedItem("QTYISS").Value
								eval("document.formname.txtTotPackZ"&i).value=childNodO.getAttribute("NoofPack")
							end if
							For Each Node in RootO.ChildNodes
								if childNodO.Attributes.getNamedItem("LOTNO").Value = Node.Attributes.getNamedItem("LOTNO").Value and childNodO.Attributes.getNamedItem("LOC").Value = Node.Attributes.getNamedItem("LOC").Value and childNodO.Attributes.getNamedItem("BIN").Value = Node.Attributes.getNamedItem("BIN").Value then
									RootO.removeChild Node
									RootO.appendChild ChildNodO
								end if
							next
						next
						HeaderNodeO.removeChild childNod
					end if
				next
			end if
		end if
	next
	'alert "CHK="& RootO.XML
end Function

Function CheckSubmit()
	dim i,Root
	dim arrTemp
	Dim sTotNoofPack

	Set Root = OutData.documentElement
	'alert("PickSt="&Root.xml)
	'exit function
	
	for i=0 to document.formname.elements.length - 1

		if document.formname.elements(i).type = "text" then			
			if (Instr(document.formname.elements(i).name,"txtQty") > 0) then
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
	'alert(sTotQty)
	if not sTotQty = 0 then
	'alert(Root.xml)
		 'For i = 1 to iCtr
		 i=0
			For Each HeaderNode In Root.childNodes   		
				i= i + 1	
				set Q = eval("document.formname.txtIss"&i)
				'alert(trim(Q.value))
				HeaderNode.setAttribute "QTYISS", trim(Q.value)
				'HeaderNode.setAttribute "TOT", trim(Q.value)
				HeaderNode.setAttribute "NoofPack",eval("document.formname.txtTotPackZ"&i).value
				sTotNoofPack = cint(sTotNoofPack)+cint(eval("document.formname.txtTotPackZ"&i).value)
			next
		'Next
	
		'alert Root.XML
		
		Root.setAttribute "TOT", sTotQty
		
		Set RootO = objTemp.documentElement
		'alert(RootO.xml)
		For Each HeaderNodeO In RootO.childNodes
			if strcomp(HeaderNodeO.NodeName,"ITEM") = 0 then 
				if strcomp(HeaderNodeO.getAttribute("CLACODE"),iClass)=0 and strcomp(HeaderNodeO.getAttribute("ITMCODE"),iItem)=0 and strcomp(HeaderNodeO.getAttribute("ENTRYNO"),iEntNo)=0 then
					For Each childNod In HeaderNodeO.childNodes
						if StrComp(Trim(childNod.NodeName),"Pick") = 0 then
							set oRem = HeaderNodeO.removeChild(childNod)
						end if						
					next
					HeaderNodeO.setAttribute "ISSQTY",""' trim(Q.value)
						if not Root.getAttribute("TOT")= "" then				
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
				if HeaderNodeO.getAttribute("ENTRYNO") = iEntNo And HeaderNodeO.getAttribute("CLACODE") = iClass and HeaderNodeO.getAttribute("ITMCODE") = iItem then
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
	'alert("PickEd="&Root.xml)
	window.close()
end Function

Function window_onunload() 
	'alert objTemp.documentElement
	set window.returnValue = objTemp.documentElement
	window.close()
end Function

Function CheckLot(obj,iEntNo,iClass,iItem,sOrg,sOptName,InvRecNo)
	dim sItem,sClass,aobj,iSerialNo,sTotNoofPack
	' alert(Obj.name)
	arrTemp = split(obj.name,":")
	iLineNo = arrTemp(1)
	sLot = arrTemp(2)
	sStore = arrTemp(3)
	sBin = arrTemp(4)
	iStQty = arrTemp(5)
	
	Set Root = OutData.documentElement 
	sTempValues = obj.name&":"&iClass&":"&iItem&":"&sOrg&":"&InvRecNo&":"&iEntNo&":"&sOptName

	set OutDataValue = showModalDialog("mrsPickDetailLotPop.asp?sTemp="&sTempValues,OutData,"dialogHeight:490px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No")
	sExp = "//Pick/PICK[@LOC = """&sStore&""" and @BIN = """&sBin&""" and @LOTNO = """&sLot&"""]/Pagination"
	Set Tempnode = Root.Selectnodes(sExp)
	If Tempnode.Length > 0 Then
	'alert(Tempnode.Item(0).Attributes.Item(0).Nodevalue)
		sTempValues = split(Tempnode.Item(0).Attributes.Item(0).Nodevalue,"```")
		while UBound(sTempValues) = 0
			sTempValues = replace(Tempnode.Item(0).Attributes.Item(0).Nodevalue,"|","&")
			set OutDataValue = window.showModalDialog("mrsPickDetailLotPop.asp?sTemp="&sTempValues,OutData,"dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No")
			sExp = "//Pick/PICK[@LOC = """&sStore&""" and @BIN = """&sBin&""" and @LOTNO = """&sLot&"""]/Pagination"
			Set Tempnode = Root.Selectnodes(sExp)
			If Tempnode.Length > 0 Then
				sTempValues = split(Tempnode.Item(0).Attributes.Item(0).Nodevalue,"```")
			end if
		wend
	end if
	
	if UBound(sTempValues)>0 then
		if left(trim(sTempValues(1)),2)="NO" then exit function
	end if 'if UBound(sTempValues)>0 then
	

	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.getNamedItem("LOTNO").Value = sLot and HeaderNode.Attributes.getNamedItem("LOC").Value = sStore and HeaderNode.Attributes.getNamedItem("BIN").Value = sBin  then
			set Q = eval("document.formname.txtIss"&iLineNo)
			Q.value = HeaderNode.Attributes.getNamedItem("QTYISS").Value
			eval("document.formname.txtTotPackZ"&iLineNo).value = Root.getAttribute("NoofPack")	
			exit function
		end if
	Next

end Function
