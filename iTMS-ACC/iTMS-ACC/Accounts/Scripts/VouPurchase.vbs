Dim TaxRoot

FUNCTION setTaxPercentage(sCatCode,sTaxCode,objText)

	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
	IF Cdbl(Trim(objText.value)) < 0 Then 
		Msgbox ("Tax Value Should be Greater than Zero ")
		objText.focus()
		Exit Function
	End IF
	if trim(objText.value)<>"" then
		if IsNumeric(objText.value) then
			For Each oNodTemp in TaxRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=sCatCode and oNodTemp.Attributes.Item(1).nodeValue=sTaxCode then
					oNodTemp.Attributes.Item(4).nodeValue=objText.value
					popTax
					exit function
				end if
			next
		
		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if
	end if		
END FUNCTION

FUNCTION setTaxAmount(sCatCode,sTaxCode,objText)

	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
	
	
	dInvAmount=document.formname.hAmount.value
	dBasicTotal=document.formname.hBasicValue.value
	dTotal	=document.formname.hAmount.value
	
	if trim(objText.value)<>"" then
		if IsNumeric(objText.value) then
			For Each oNodTemp in TaxRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=sCatCode and oNodTemp.Attributes.Item(1).nodeValue=sTaxCode then
					sTaxMode=oNodTemp.Attributes.Item(2).nodeValue 
					sFormula=oNodTemp.Attributes.Item(3).nodeValue 
					if sTaxMode="P" then
						oNodTemp.Attributes.Item(4).nodeValue = calPercentage(sFormula,dBasicTotal,dTotal,objText.value)
					else
						oNodTemp.Attributes.Item(4).nodeValue = objText.value
						oNodTemp.Attributes.Item(5).nodeValue = objText.value
						
					end if	
					popTax
					exit function
				end if
			next
		
		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if
	end if		
END FUNCTION

FUNCTION popTax()

	dim dInvAmount,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax
	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next

	dInvAmount = document.formname.hAmount.value
	dBasicTotal= document.formname.hBasicValue.value
	dTotal	= document.formname.hAmount.value
	
	For Each oNodEntry in TaxRoot.childNodes
	
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue 
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue 
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue 
		sFormula=oNodEntry.Attributes.Item(3).nodeValue 
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue
		
		
		IF Cstr(sCatCode) <> "0" and Cstr(sTaxCode) <> "0" Then
			iRndOff = oNodEntry.Attributes.Item(7).nodeValue
		Else
			iRndOff = 0
		End IF
		
		dTaxValue=FormatNumber(dTaxValue,2,,,0)
		
		if sTaxMode="P" then
			IF Cstr(sCatCode) <> "0" and Cstr(sTaxCode) <> "0" Then
				dTax=CalculateTax(sFormula,dBasicTotal,dTotal,dTaxValue)
				eval("document.formname.txtTaxPer"&sCatCode&sTaxCode).value=dTaxValue
			End IF
		else
			dTax=dTaxValue
		end if
				
		If iRndOff = 1 Then
			dTax = Round(dTax, 0)
		End If
		dInvAmount=dInvAmount + CDbl(dTax)
		
		dTax=FormatNumber(dTax,2,,,0)
		oNodEntry.Attributes.Item(5).nodeValue=dTax
		
		IF Cstr(sCatCode) <> "0" and Cstr(sTaxCode) <> "0" Then
			eval("document.formname.txtTaxValue"&sCatCode&sTaxCode).value=dTax
		End IF
	next	
	
	IF Cstr(document.Formname.hRndChk.Value) = "1" Then
		dRoundedInvvalue = round(dInvamount,0)
		dRoundedoff = Round(cdbl(dRoundedInvvalue) - cdbl(dInvamount),2)
	Else
		dRoundedInvvalue = dInvamount
		dRoundedoff = 0
	End IF
	
	dRoundedInvvalue=FormatNumber(dRoundedInvvalue,2,,,0)
	
	'Msgbox "&&& " & dRoundedoff
	document.formname.txtroundoff.value = dRoundedoff
	document.formname.txtInvValue.value=dRoundedInvvalue
		
		'Taxroot.Attributes.Item(0).Nodevalue = dBasicTotal
		'TaxRoot.Attributes.Item(1).nodeValue=dTotal
		'TaxRoot.Attributes.Item(2).nodeValue=dRoundedInvvalue	
		'TaxRoot.Attributes.Item(3).nodeValue=dRoundedoff	
		
		
	Taxroot.Attributes.Item(0).Nodevalue = dRoundedInvvalue
	TaxRoot.Attributes.Item(1).nodeValue = dTotal
	TaxRoot.Attributes.Item(2).nodeValue = dRoundedInvvalue
	TaxRoot.Attributes.Item(3).nodeValue = dRoundedoff
	
	'Msgbox TaxRoot.Attributes.Item(3).nodeValue
	
	'document.formname.txtInvValue.value = formatNumber(dInvAmount,2,,,0)
	
END FUNCTION

FUNCTION CalculateTax(sFormula,dBValue,dDValue,dPercentage)


dim saTemp,dTaxAmount,iCounter,iTemp
dim oNodTemp
dim saTemp1

saTemp=Split(sFormula,",")
if trim(saTemp(0))="BV" then
	dTaxAmount=dBValue
	iTemp=1
elseif trim(saTemp(0))="BD" then
	dTaxAmount=dDValue
	iTemp=1
else
	dTaxAmount=0
	iTemp=0
end if

for iCounter=iTemp to UBound(saTemp)
	saTemp1=Split(trim(saTemp(iCounter)),"#")
	For Each oNodTemp in TaxRoot.childNodes
		if oNodTemp.Attributes.Item(0).nodeValue=trim(saTemp1(0)) and oNodTemp.Attributes.Item(1).nodeValue=trim(saTemp1(1)) then
			dTaxAmount=CDbl(dTaxAmount)+CDbl(oNodTemp.Attributes.Item(5).nodeValue)
		end if
	next
next
CalculateTax=FormatNumber(dTaxAmount*(cdbl(dPercentage)/100),2,,,0)
END FUNCTION

FUNCTION calPercentage(sFormula,dBValue,dDValue,dAmount)

dim saTemp,dTaxAmount,iCounter,iTemp
dim oNodTemp
dim saTemp1

saTemp=Split(sFormula,",")
if trim(saTemp(0))="BV" then
	dTaxAmount=dBValue
	iTemp=1
elseif trim(saTemp(0))="BD" then
	dTaxAmount=dDValue
	iTemp=1
else
	dTaxAmount=0
	iTemp=0
end if

for iCounter=iTemp to UBound(saTemp)
	saTemp1=Split(trim(saTemp(iCounter)),"#")
	For Each oNodTemp in TaxRoot.childNodes
		if oNodTemp.Attributes.Item(0).nodeValue=trim(saTemp1(0)) and oNodTemp.Attributes.Item(1).nodeValue=trim(saTemp1(1)) then
			dTaxAmount=CDbl(dTaxAmount)+CDbl(oNodTemp.Attributes.Item(5).nodeValue)
		end if
	next
next
if cdbl(dTaxAmount)>0 then
	calPercentage=FormatNumber(CDbl(dAmount)/cdbl(dTaxAmount)*100,2,,,0)
else
	calPercentage=0
end if	

END FUNCTION
