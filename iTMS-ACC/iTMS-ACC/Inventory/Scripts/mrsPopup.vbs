Function DisplayDet(sText)
	FontFace="Verdana,8,,bold" 
	arrTemp = split(sText,"|")
	TopicText = "To Purchase Requisition : " & arrTemp(0) & vbcrlf
	TopicText = TopicText & "For Stock Transfer         : " & arrTemp(1)

	'document.formname.penDet.TextPopup TopicText, FontFace, 10,10,0,0
end Function

Function CheckQty(obj)
	dim sItem,sClass,a
	arrTemp = split(obj.name,":")

	sClass = arrTemp(1)
	sItem = arrTemp(2)
	iEntNo = arrTemp(3)
	sOptName = arrTemp(4)

	sTempValues = sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&iEntNo&":"&sOptName
	
	OutDataValue = showModalDialog("mrsIssueQtyParaPoP.asp?sTemp="&sTempValues,OutData,"dialogHeight:385px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No")
end Function

Function DisplayItem(obj)
	sTempValues = obj

	showModalDialog "itmDetailsPop.asp?sTemp="&sTempValues,idOrgName.innerText,"dialogHeight:360px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No"
end Function

Function CheckSch(obj,i)
	dim sItem,sClass,a,aobj
	dim qty
	arrTemp = split(obj.name,":")

	sClass = arrTemp(1)
	sItem = arrTemp(2)
	iCounter = arrTemp(3)
	sOptName = arrTemp(4)

	set qty = eval("document.formname.txtQtyIssue"+i)

	sTempValues = qty.value&":"&sItem&":"&sClass&":"&document.formname.hMRSNo.value&":"&document.formname.hOrgID.value&":"&iCounter&":"&sOptName

	OutDataValue = showModalDialog("mrsIssueSchedulePoP.asp?sTemp="&sTempValues,Data,"dialogHeight:480px;dialogWidth:390px;center:Yes;help:No;resizable:No;status:No")
end Function
