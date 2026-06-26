
Function getRatePerQtyUoM(ORGID,CLASSCODE,ITEMCODE,QTYUOM,RATEUOM,RATE)
	
	'' To fetch the Option to base rate and option to base operator for the selected
	'' item,classification, organisation
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","XMLgetUoMConvFactor.asp?ORGID=" & ORGID & "&CLASSCODE="& CLASSCODE & "&ITEMCODE="& ITEMCODE & "&QTYUOM="& QTYUOM & "&RATEUOM="& RATEUOM & "&FLAG=P" ,false
	objhttp.send 
	'msgbox objhttp.responsetext 
		
	if objhttp.responseXML.xml <> "" then
		tempData.loadXML objhttp.responseXML.xml
		Set Root = tempData.documentElement
		
		For each nd in Root.childNodes
			iOptToBaseRate = nd.getAttribute("OptionToBaseRate")
			iOptToBaseOperator = nd.getAttribute("OptionToBaseOperator")
		Next
	end if

	''test data ----------------------------------------
	'Qty - 1000 KG	Rate = 10000 / Bls 
	'Option to base rate = 175 : Option to base Operator = 1

	'10000 / 175 = 57.15 / KG

	'Rate/ QTy UoM = 57.15 / KG
	'' -----------------------------------------------

	'' if conversion operator is 0 : Multiply
	if cint(iOptToBaseOperator) = 0 then	' Multiply
		dRatePerQtyUoM = cdbl(RATE) * cdbl(iOptToBaseRate)
	elseif cint(iOptToBaseOperator) = 1 then ' Divide
		'' if conversion operator is 1 : divide
		dRatePerQtyUoM = cdbl(RATE) / cdbl(iOptToBaseRate)
	end if

	getRatePerQtyUoM = dRatePerQtyUoM	'RETURN RATE/QTY UOM
	
End Function