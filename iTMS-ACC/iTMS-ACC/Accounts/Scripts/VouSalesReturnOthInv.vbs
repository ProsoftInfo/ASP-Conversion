Dim TaxRoot
FUNCTION setQty(objTemp,iSno,sCallTy)
dim dQty,dAmount,dTotal,dRate
Dim dOldQty,dNewQty,dRatePer,sAccTy,dNewRateVal,dOldRate,dNewRate
Dim dDisAmt,dDisPer,dOldDis,dNewDis
dTotal=0
	'alert(objTemp.value&","&iSno&","&sCallTy)
	sAccTy = document.formname.hAccType.Value
	dNewQty = Trim(objTemp.value)
	dOldQty = Eval("document.all.tOldQty"&iSno).innerHtml
	dNewRate = Eval("document.formname.txtRate"&iSno).value
	dOldRate = Eval("document.all.tOldRate"&iSno).innerHtml
	dNewDis = Eval("document.formname.txtDis"&iSno).Value
	dOldDis = Eval("document.all.tOldDis"&iSno).innerHtml
	
	if Trim(dNewRate)="" or IsNull(dNewRate) then 
	    dNewRate = 0
	    Eval("document.formname.txtRate"&iSno).value = "0"
	end if
	
	if Trim(dOldRate)="" or IsNull(dOldRate) then 
	    dOldRate = 0
	    Eval("document.all.tOldRate"&iSno).innerHtml = "0"
	end if
	
	dNewQty = CDbl(dNewQty)
	dOldQty = CDbl(dOldQty)

	dNewRate = CDbl(dNewRate)
	dOldRate = CDbl(dOldRate)
	
	dNewDis = Cdbl(dNewDis)
	dOldDis = Cdbl(dOldDis)
	
	'IF Cstr(document.formname.hNoteType.Value) = "C" Then
	
		'IF Cstr(sCallTy) = "Q" Then
		'	IF dNewQty > dOldQty Then
		'		Msgbox "Quantity Should be Less than the Invoiced Quantity "
		'		objTemp.focus()
		'		Exit Function
		'	End IF
		'End IF
		
		
		'IF dNewRate > dOldRate Then
		'	Msgbox "Rate Should be Less than the Invoiced Rate "
		'	Eval("document.formname.txtRate"&iSNo).focus()
		'	Exit Function
		'End IF
		
		'IF dNewDis > dOldDis Then
		'	Msgbox "Discount Percentage Should be Less than the Invoiced Discount Percentage "
		'	Eval("document.formname.txtDis"&isNo).focus()
		'	Exit Function
		'End IF
		
	'Else
	
		'IF Cstr(sCallTy) = "Q" Then
		'	IF dNewQty < dOldQty Then
		'		Msgbox "Quantity Should be Greater than the Invoiced Quantity "
		'		objTemp.focus()
		'		Exit Function
		'	End IF
		'End IF
				
		'IF dNewRate < dOldRate Then
		'	Msgbox "Rate Should be Greater than the Invoiced Rate "
		'	Eval("document.formname.txtRate"&iSNo).focus()
		'	Exit Function
		'End IF
		
		'IF dNewDis < dOldDis Then
		'	Msgbox "Discount Percentage Should be Greater than the Invoiced Discount Percentage "
		'	Eval("document.formname.txtDis"&iSNo).focus()
		'	Exit Function
		'End IF
		
	'End IF
	
	
		

	if trim(objTemp.value)="" then
		dQty=0
	elseif IsNumeric(objTemp.value)=false then
		Msgbox("Enter Numeric values for Quantity")
		objTemp.focus
		exit Function
	elseif CDbl(objTemp.value)<0 or CDbl(objTemp.value)>9999999.999 then
		Msgbox("Quantity should be >=0 and < 9999999.999")
		objTemp.focus
		exit Function
	else
		dQty=objTemp.value	
	end if	

	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="Details" then
			set DeatilRoot=oNodTemp
		end if
	next
	'alert(DeatilRoot.xml)
	
	For Each oNodTemp in DeatilRoot.childNodes	
		IF oNodTemp.NodeName = "Entry" Then
			if oNodTemp.Attributes.GetNamedItem("No").value=iSno then
				oNodTemp.Attributes.GetNamedItem("Qty").value=dQty			
				'dRate=FormatNumber(Round(oNodTemp.Attributes.GetNamedItem("Rate").value,0),2,,,0)
				'dRate = FormatNumber(Round(dNewRate,0),2,,,0)
				dRate = FormatNumber(dNewRate,2,,,0)
				oNodTemp.Attributes.GetNamedItem("Rate").value =dNewRate 
				
				'dRatePer = oNodTemp.Attributes.GetNamedItem("RatePer").value
				IF Cdbl(dRatePer) = 0 Then
					dRatePer = 1
				End IF
				'Msgbox dRate &" " & dQty & " " & dRatePer
				dAmount=(CDbl(dRate)/CDbl(dRatePer)) * CDbl(dQty)
				dDisPer = Eval("document.formname.hDisPer"&iSno).Value
			    IF cdbl(dDisPer) <> 0 Then
				    dDisAmount=(CDbl(trim(dDisPer))/100)*dAmount
				    dAmount = FormatNumber(dAmount,2,0,0,0)
				    dDisAmount = FormatNumber(dDisAmount,2,0,0,0)
				    dDisNewAmt = eval("document.formname.txtDis"&iSno).value
				    if CDbl(dRate)> 0 then
				        if dDisAmount<>dDisNewAmt then
				            dDisPer = (cdbl(dDisNewAmt)/cdbl(dRate))*100
				            dDisPer = FormatNumber(dDisPer,2,0,0,0)
				            oNodTemp.Attributes.GetNamedItem("DisPer").value=dDisPer
				            dDisAmount = dDisNewAmt
				        end if
				    else
				        sDisAmount = 0
				    end if 'if CDbl(dRate)> 0 then
				    if CDbl(dQty) =0 then dDisAmount  =0 
				    dAmount=dAmount-dDisAmount
				Else
				    dDisAmount = eval("document.formname.txtDis"&iSNo).value
				    if CDbl(dRate)>0  then
				        dDisPer = (cdbl(dDisAmount)/cdbl(dRate))*100
				        dDisPer = FormatNumber(dDisPer,2,0,0,0)
				    else
				        dDisAmount =0
				    end if 'if CDbl(dRate)>0 then
				    if CDbl(dQty) =0 then dDisAmount  =0 
				    
				    dAmount = dAmount - dDisAmount
				    oNodTemp.Attributes.GetNamedItem("DisPer").value=dDisPer
			    End IF
				dTotal=dTotal+CDbl(dAmount)
				oNodTemp.Attributes.GetNamedItem("DisAmount").value=dDisAmount
				oNodTemp.Attributes.GetNamedItem("Amount").value=dAmount

			else
				dTotal=dTotal+CDbl(oNodTemp.Attributes.GetNamedItem("Amount").value)
			end if
		End IF
	next
	document.formname.txtTotal.value = FormatNumber(dTotal,2,,,0)
	eval("document.formname.txtDis"&iSno).value=FormatNumber(dDisAmount,2,,,0)
	eval("document.formname.txtAmount"&iSno).value=FormatNumber(dAmount,2,,,0)
	DeatilRoot.Attributes.GetNamedItem("BasicValue").value = dTotal
	DeatilRoot.Attributes.GetNamedItem("Discount").value = dDisPer
	DeatilRoot.Attributes.GetNamedItem("ActualValue").value = dTotal
	popTax	
	
	IF document.formname.SelCrAgain.Value = "A" Then
		document.formname.txtCrNoteValue.value = document.formname.txtInvValue.Value
	End IF
	SetRetVal document.formname.SelCrAgain,"0"
END FUNCTION

FUNCTION setTotal(objTemp,iSno)

dim dAmount,dTotal
dTotal=0
	if trim(objTemp.value)="" then
		dAmount=0
	elseif IsNumeric(objTemp.value)=false then
		Msgbox("Enter Numeric values for Amount")
		objTemp.focus
		exit Function
	elseif CDbl(objTemp.value)<0 or CDbl(objTemp.value)>9999999999.99 then
		Msgbox("Amount should be >=0 and < 9999999999.99")
		objTemp.focus
		exit Function
	else
		dAmount=objTemp.value	
	end if	

	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="Details" then
			set DeatilRoot=oNodTemp
		end if
	next
	For Each oNodTemp in DeatilRoot.childNodes
		if  oNodTemp.Attributes.GetNamedItem("No").value=iSno then
			oNodTemp.Attributes.GetNamedItem("Amount").value=dAmount
			dTotal=dTotal+CDbl(dAmount)
		else
			dTotal=dTotal+CDbl(oNodTemp.Attributes.GetNamedItem("Amount").value)
		end if
	next

	document.formname.txtTotal.value=FormatNumber(dTotal,2,,,0)
	objTemp.value=FormatNumber(dAmount,2,,,0)
	DeatilRoot.Attributes.GetNamedItem("BasicValue").value=dTotal
	DeatilRoot.Attributes.GetNamedItem("Discount").value="0"
	DeatilRoot.Attributes.GetNamedItem("ActualValue").value=dTotal
	
	IF document.formname.SelCrAgain.Value <> "A" Then
		popTax
	Else
	    popTax
		dTotal = FormatNumber(dTotal,2,,,0)
		'document.formname.txtInvValue.Value = dTotal
		
	End IF
	
	'IF document.formname.SelCrAgain.Value = "A" Then
		document.formname.txtCrNoteValue.value = document.formname.txtInvValue.Value
	'End IF
	
END FUNCTION

FUNCTION setTaxPercentage(sCatCode,sTaxCode,objText)
	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
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
	IF document.formname.SelCrAgain.Value = "A" Then
		document.formname.txtCrNoteValue.value = document.formname.txtInvValue.Value
	End IF
END FUNCTION

FUNCTION setTaxAmount(sCatCode,sTaxCode,objText)
	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
	
	dInvAmount=document.formname.txtTotal.value
	dBasicTotal=dInvAmount
	dTotal	=dInvAmount
	
	if trim(objText.value)<>"" then
		if IsNumeric(objText.value) then
			For Each oNodTemp in TaxRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=sCatCode and oNodTemp.Attributes.Item(1).nodeValue=sTaxCode then
					sTaxMode=oNodTemp.Attributes.Item(2).nodeValue
					sFormula=oNodTemp.Attributes.Item(3).nodeValue
					
					if sTaxMode="P" then
						oNodTemp.Attributes.Item(4).nodeValue=calPercentage(sFormula,dBasicTotal,dTotal,objText.value)
					else
						oNodTemp.Attributes.Item(4).nodeValue=objText.value
						oNodTemp.Attributes.Item(5).nodeValue=objText.value
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
	IF document.formname.SelCrAgain.Value = "A" Then
		document.formname.txtCrNoteValue.value = document.formname.txtInvValue.Value
	End IF
END FUNCTION

FUNCTION popTax()
dim dInvAmount,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax,sTotpackvalue
Dim sStr,sTempNode,sCheckExp,sCheckNode
	set RootNode=TaxData.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
	
	dInvAmount = document.formname.txtTotal.value
	dBasicTotal = dInvAmount
	dTotal	= dInvAmount
	
	
	For Each oNodEntry in TaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue 
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue 
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue 
		sFormula=oNodEntry.Attributes.Item(3).nodeValue 
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue
		'sTotpackvalue = oNodEntry.Attributes.Item(8).nodeValue 

		If sCatCode = "0" and sTaxCode = "0" and sTaxMode = "0" Then
			iRndOff = 0
'			dTaxValue = 0
'			dTax = 0
		Else
			iRndOff = oNodEntry.Attributes.Item(7).nodeValue
		End If
		
		if IsObject(eval("document.formname.txtTaxValue"&sCatCode&sTaxCode)) then
			if sTaxMode="P" then
				dTax=CalculateTax(sFormula,dBasicTotal,dTotal,dTaxValue)
				eval("document.formname.txtTaxPer"&sCatCode&sTaxCode).value=dTaxValue
			elseif sTaxmode = "Q" then
				'msgbox cdbl(Qtysum)
				'Msgbox dTaxvalue
				dTax = cdbl(dTaxvalue) * cdbl(Qtysum)	
			elseif sTaxMode = "K" then
				'msgbox cdbl(Qtysum)
				'Msgbox sTotpackvalue
				dTax=sTotpackvalue * cdbl(Qtysum)
				'Msgbox dTax							
			else
				dTax=dTaxValue
			end if

			If sCatCode = "0" and sTaxCode = "0" and sTaxMode = "0" Then
			Else
				If iRndOff = 1 Then
					dTax = Round(dTax,0)
				End If
			End If
			IF document.formname.SelCrAgain.Value = "A" Then
			    If sCatCode <> "0" and sTaxCode <> "0" and sTaxMode <> "0" Then
			        dTax=0
			    End if
			End if
			
			dInvAmount=dInvAmount+CDbl(dTax)
			
			dTax=FormatNumber(dTax,2,,,0)
			
		    oNodEntry.Attributes.Item(5).nodeValue=dTax
		    eval("document.formname.txtTaxValue"&sCatCode&sTaxCode).value=dTax
		end if	
	next
	
	dInvAmount = FormatNumber(dInvAmount,2,,,0)
	document.formname.txtInvValue.value = FormatNumber(round(dInvamount),2,,,0)
	iRoundedInvvalue = document.formname.txtInvValue.value
	iRoundedoff = Round(cdbl(iRoundedInvvalue) - cdbl(dInvamount),2)
	
	sCheckExp = "//TaxDetails[@RoundOffValue]"
	Set sCheckNode = TaxRoot.selectNodes(sCheckExp)
	
		
	TaxRoot.Attributes.Item(0).nodeValue=document.formname.txtInvValue.value
	TaxRoot.Attributes.Item(1).nodeValue=document.formname.txtTotal.value
	TaxRoot.Attributes.Item(2).nodeValue=document.formname.txtTotal.value
	IF sCheckNode.Length <> 0 Then
		TaxRoot.Attributes.Item(3).nodeValue=iRoundedoff
	End IF
	
	sStr = "//Tax[@CatCode=0 and @TaxCode=0 and @TaxMode=0]"
	Set sTempNode = RootNode.selectNodes(sStr)
	
	IF sTempNode.length <> 0 Then
		sTempNode.Item(0).Attributes.Item(5).Value = iRoundedoff
		document.formname.txtTaxValue00.value = iRoundedoff
	End IF
	
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

If Trim(sFormula) = "" or Trim(sFormula) = "0" Then Exit Function

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

calPercentage=FormatNumber(CDbl(dAmount)/cdbl(dTaxAmount)*100,2,,,0)

END FUNCTION

Function Qtysum()
	Set Rootnode = TaxData.documentElement
	sExp = "//Details/Entry"
	Set sTempnode = Rootnode.selectnodes(sExp)
	if sTempnode.Length > 0 then
	Qty = 0
	For Ictr = 0 to sTempnode.length - 1
		Qty = Cdbl(Qty) + cdbl(sTempnode.Item(ICtr).Attributes(3).nodevalue)
	Next 
	Else
		Qty  = 0
	End if
	QtySum = Qty
End function
