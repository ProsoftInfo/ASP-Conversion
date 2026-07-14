<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	popTaxDetails.asp
	'Module Name				:	Purchase (Transactions-Invoice)
	'Author Name				:
	'Created On					:
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	invPurInvoiceHeaderEntry.asp
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/PurchaseTermsConditions.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS-Invoice Form Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/calcAlternateUoM.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/popTaxDetails.js"></SCRIPT>

<script type="application/xml" id="TaxFormData" data-itms-xml-island="1"><Root/></script>
<script type="text/plain" data-itms-legacy-client-script="1">
Dim ItemTaxData,PurchType
Dim TaxRoot,RootNode,oNodTemp
Set ItemTaxData = window.dialogArguments
'------------------------------------------------------------------------------------------
Function init()
	dim root,ItemNode,i,TaxName,sCatCode,sTaxCode,TaxMode,TaxPer,TaxVal
	dim oRow,headerCell
	'____________________________Getting ItemNames,PurTypeBasicVal and NettBasicVal_____________
	dim dNetTotal,dDisTotal,dAmount,dRate,dQty,dDisAmount,ItemNm,RootNode,oNodTemp,ItemRoot,PurType
	set RootNode=ItemTaxData.documentElement
	PurchType=trim(document.formname.hPurType.value)
	PurchTypeName=trim(document.formname.hPurTypeName.value)
	set ItemRoot=RootNode.Selectnodes("//ItemDetails/Item[@PurchaseType="& PurchType &"]")
	dNetTotal = 0
	dDisTotal = 0
	dTBasicVal = 0
	ItemNm	=""
	For i=0 to  ItemRoot.Length-1
			set oNodTemp = ItemRoot.Item(i)
			PurType=oNodTemp.Attributes.getNamedItem("PurchaseType").value
			ItemNm =ItemNm + oNodTemp.Attributes.getNamedItem("ItmDescription").value + ","
			dDisTotal=CDbl(dDisTotal)+CDbl(oNodTemp.Attributes.Item(7).nodeValue)
			dQty=oNodTemp.Attributes.Item(4).nodeValue

			'' For Alt. UoM
			'dRate=oNodTemp.Attributes.Item(5).nodeValue
			dRate = oNodTemp.Attributes.getNamedItem("RatePerQtyUoM").value

			dTBasicVal = dTBasicVal + (CDbl(dQty)*CDbl(dRate))
			dAmount=(CDbl(dQty)*CDbl(dRate))-CDbl(oNodTemp.Attributes.Item(7).nodeValue)
			dNetTotal=CDbl(dNetTotal)+dAmount
	next
	if ItemNm<>"" then
		ItemNm=mid(ItemNm,1,len(ItemNm)-1)
	end if
	document.all.spnItemName.innerhtml=ItemNm
	document.all.spnPurType.innerhtml=PurchTypeName
	document.all.spnAmount.innerhtml=FormatNumber(Round(dNetTotal,2),2,,,0)
	document.all.spnBasicValue.innerhtml=FormatNumber(Round(dTBasicVal,2),2,,,0)
	'____________________________________________________________________________________
	set root=ItemTaxData.DocumentElement
	Set ItemNode=Root.Selectnodes("//TaxDetails[@PurchaseType="& PurchType &"]/Tax")
	for i = 0 to ItemNode.Length - 1
		TaxName=ItemNode.Item(i).text
		sCatCode=trim(ItemNode.Item(i).Attributes.getNamedItem("CatCode").value )
		sTaxCode=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxCode").value )
		TaxMode=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxMode").value )

		if trim(sCatCode) <> "0" and trim(sCatCode) <> "0" then
			if TaxMode="F" then
				TaxPer=""
				TaxVal=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxValue").value )
			else
				TaxVal=""
				TaxPer=ItemNode.Item(i).Attributes.getNamedItem("TaxValue").value
			end if

			set oRow = document.all.tblTaxDet.insertRow(document.all.tblTaxDet.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.className="ExcelSerial"
			headerCell.innerhtml=i+1
			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			headerCell.className="ExcelSerial"
			headerCell.innerhtml=TaxName
			headerCell.align="left"

			set headerCell=oRow.insertCell()
			if TaxMode<>"F" then
				set oText = document.createElement("<input type=""text"" name=""txtTaxPer" & sCatCode & sTaxCode & """ value="""& TaxPer &""" size=""4"" onBlur=""setTaxPercentage('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" class=""FormElem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.innerhtml=headerCell.innerhtml + " %"
				headerCell.className="ExcelFieldCell"
			else
				headerCell.className="ExcelSerial"
			end if

			headerCell.align="center"

			set headerCell=oRow.insertCell()
			if TaxMode="F" then
				set oText = document.createElement("<input type=""text"" name=""txtTaxValue" & sCatCode & sTaxCode & """ value="""& TaxVal &""" size=""11"" onBlur=""setTaxAmount('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" class=""FormElem"">")
			else
				'set oText = document.createElement("<input type=""text"" name=""txtTaxValue" & sCatCode & sTaxCode & """ value="""& TaxVal &""" size=""11"" onBlur=""setTaxAmount('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" ReadOnly class=""FormElemRead"">")
				set oText = document.createElement("<input type=""text"" name=""txtTaxValue" & sCatCode & sTaxCode & """ value="""& TaxVal &""" size=""11"" style=""text-align: Right"" ReadOnly class=""FormElemRead"">")
			end if
			headerCell.className="ExcelFieldCell"
			headerCell.appendChild(oText)
			headerCell.width = "10"

			headerCell.align="center"
		end if 'if trim(sCatCode) <> "0" and trim(sCatCode) <> "0" then
	Next

	set oRow = document.all.tblTaxDet.insertRow(document.all.tblTaxDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelSerial"
	headerCell.colspan="3"
	headerCell.innerhtml="Total"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	'set oText = document.createElement("<input type=""text"" name=""txtTotTax"" value="""& TaxVal &""" size=""11"" onBlur=""setTaxAmount('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" ReadOnly class=""FormElemRead"">")
	set oText = document.createElement("<input type=""text"" name=""txtTotTax"" value="""& TaxVal &""" size=""11""  style=""text-align: Right"" ReadOnly class=""FormElemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelFieldCell"
	headerCell.align="center"
	popTax
end function

'-------------------------------------------------------------------------------------------

FUNCTION setTaxPercentage(sCatCode,sTaxCode,objText)
dim TaxNodes
	set RootNode=ItemTaxData.documentElement

	set TaxNodes=RootNode.SelectNodes("//TaxDetails[@PurchaseType="& PurchType &"]/Tax")
	if trim(objText.value)<>"" then
		if IsNumeric(objText.value) then
			For i=0 to  TaxNodes.Length - 1
				set oNodTemp=TaxNodes.item(i)
				if oNodTemp.Attributes.Item(0).nodeValue=sCatCode and oNodTemp.Attributes.Item(1).nodeValue=sTaxCode then
					oNodTemp.Attributes.Item(4).nodeValue=objText.value
					popTax
					exit function
				end if
			Next
		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if
	end if
END FUNCTION
'-------------------------------------------------------------------------------------------
FUNCTION popTax()
dim dInvAmount,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax,TaxNodes
dim nDisplayTotal

set RootNode=ItemTaxData.documentElement
set TaxNodes=RootNode.SelectNodes("//TaxDetails[@PurchaseType="& PurchType &"]")
if TaxNodes.Length>0 then
	set TaxRoot=TaxNodes.Item(0)
	dInvAmount=document.all.spnAmount.innerhtml
	dBasicTotal=document.all.spnBasicValue.innerhtml
	dTotal	=document.all.spnAmount.innerhtml

	if trim(dInvAmount) = "" then dInvAmount = 0
	if trim(dBasicTotal) = "" then dBasicTotal = 0
	if trim(dTotal) = "" then dTotal = 0

	nDisplayTotal =0
'''	nDisplayTotal = CDbl(dTotal)

	For Each oNodEntry in TaxRoot.childNodes

		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue
		sFormula=oNodEntry.Attributes.Item(3).nodeValue
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue

		If sCatCode <> "0" and sTaxCode <> "0" then		' Except round off node - 09 Aug 04


			' Added on 16-Jun-04 - should test
			if sTaxMode="P" then
				dTax=CalculateTax(sFormula,dBasicTotal,dTotal,dTaxValue,PurchType)
				eval("document.formname.txtTaxPer"&sCatCode&sTaxCode).value=dTaxValue
			elseif sTaxMode = "Q" then
				dTax = cdbl(dTaxvalue) * cdbl(Qtysum)
			elseif sTaxMode = "K" then
				sTotpackvalue = 0
				sExp = "//Tax[@CatCode = "&sCatCode&" and @TaxCode ="&sTaxCode&"]/Taxpack"
				set sNode = TaxRoot.Selectnodes(sExp)
				''alert taxroot.xml
				If sNode.Length > 0 then
					For IMCtr = 0 to sNode.Length - 1
						'Msgbox sNode.Item(iMCtr).Attributes.Item(3).nodevalue
						sTotpackvalue = cdbl(sTotpackvalue) + cdbl(sNode.Item(iMCtr).Attributes.Item(3).nodevalue)
					Next
				end if
				dTax=sTotpackvalue
			else
				dTax=dTaxvalue
			end if

			dTRndoff = oNodEntry.getAttribute("Rndoff")

			if trim(dTRndoff) = "1" then
				'alert(" dTRndoff = " + dTRndoff)

				'dtax = Round(dTax,0)
				'alert( dtax)
				dtax = RndOff(dTax)
			End if


			dTax=FormatNumber(dTax,2,,,0)
			oNodEntry.Attributes.Item(5).nodeValue=dTax

			dInvAmount=dInvAmount+CDbl(dTax)

			eval("document.formname.txtTaxValue"&sCatCode&sTaxCode).value=dTax

			nDisplayTotal = nDisplayTotal + dTax
'''			eval("document.formname.txtSubTaxValue"&sCatCode&sTaxCode).value=FormatNumber(nDisplayTotal,2,,,0)

		end if
	Next

	'msgbox " dInvamount = " & dInvamount
'''	dRoundedInvvalue = RndOff(dInvamount)
'''	dRoundedoff = Round(cdbl(dRoundedInvvalue) - cdbl(dInvamount),2)

'''	dRoundedoff = FormatNumber(dRoundedoff,2,,,0)
'''	dRoundedInvvalue=FormatNumber(dRoundedInvvalue,2,,,0)

'''	document.formname.txtRoundOff.value = dRoundedoff



	' set value for round off node - 09 Aug 04
'''	For Each oNodEntry in TaxRoot.childNodes
'''		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
'''		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue

'''		If sCatCode = "0" and sTaxCode = "0" then
'''			oNodEntry.setAttribute "TaxValue",document.formname.txtRoundOff.value
'''			oNodEntry.setAttribute "TaxAmount",document.formname.txtRoundOff.value
			'msgbox document.formname.txtRoundOff.value
'''		End if
'''	Next

'''	document.formname.txtInvValue.value=dRoundedInvvalue
	TaxRoot.Attributes.Item(0).nodeValue=dBasicTotal
	TaxRoot.Attributes.Item(1).nodeValue=dTotal
'''	TaxRoot.Attributes.Item(2).nodeValue=dRoundedInvvalue
'''	TaxRoot.Attributes.Item(3).nodeValue=dRoundedoff
'''	TaxRoot.setAttribute "InvAmtWithoutRoundOff", dInvamount
'_____________________________________________________________

	'alert(nDisplayTotal)
	TaxRoot.setAttribute "TotalTax", FormatNumber(nDisplayTotal,2,,,0)
	TaxRoot.setAttribute "SubTotal", FormatNumber((dTotal+nDisplayTotal),2,,,0)
	document.formname.txtTotTax.value=nDisplayTotal
end if
END FUNCTION
'----------------------------------------
FUNCTION CalculateTax(sFormula,dBValue,dDValue,dPercentage,sPurType)
dim saTemp,dTaxAmount,iCounter
dim oNodTemp,iTemp
dim saTemp1

set RootNode=ItemTaxData.documentElement
if trim(sPurType) ="0" then
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
else
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			if trim(oNodTemp.getAttribute("PurchaseType")) = trim(sPurType) then
				set TaxRoot=oNodTemp
			end if
		end if
	next
end if

saTemp=Split(sFormula,",")
if trim(saTemp(0)) = "BV" then
	dTaxAmount = dBValue
	iTemp = 1
elseif trim(saTemp(0))="BD" then
	dTaxAmount = dDValue
	iTemp = 1
else
	dTaxAmount = 0
	iTemp = 0
end if
for iCounter = iTemp to UBound(saTemp)
	saTemp1=Split(trim(saTemp(iCounter)),"#")
	For Each oNodTemp in TaxRoot.childNodes
		if oNodTemp.Attributes.Item(0).nodeValue=trim(saTemp1(0)) and oNodTemp.Attributes.Item(1).nodeValue=trim(saTemp1(1)) then
			dTaxAmount=CDbl(dTaxAmount)+CDbl(oNodTemp.Attributes.Item(5).nodeValue)
		end if
	next
next

if trim(dPercentage)<>"" then
	CalculateTax=dTaxAmount*(cdbl(dPercentage)/100)
else
	CalculateTax=dTaxAmount
end if

END FUNCTION
'-------------------------------------------------------------------------------------------
FUNCTION setTaxAmount(sCatCode,sTaxCode,objText)

	Dim sTaxCatType,TaxNodes

	Set RootNode=ItemTaxData.documentElement
	set TaxNodes=RootNode.SelectNodes("//TaxDetails[@PurchaseType="& PurchType &"]/Tax")

	dInvAmount=document.all.spnAmount.innerhtml
	dBasicTotal=document.all.spnBasicValue.innerhtml
	dTotal	=document.all.spnAmount.innerhtml

	if trim(objText.value)<>"" then
		if IsNumeric(objText.value) then
			For i=0 to  TaxNodes.Length - 1
				set oNodTemp=TaxNodes.item(i)
				if oNodTemp.Attributes.Item(0).nodeValue=sCatCode and oNodTemp.Attributes.Item(1).nodeValue=sTaxCode then

					sTaxMode = oNodTemp.Attributes.Item(2).nodeValue
					sFormula = oNodTemp.Attributes.Item(3).nodeValue
					sTaxCatType = oNodTemp.getAttribute("TaxCategoryType")

					if sTaxMode="F" then
						'if sTaxCetType <> "C" and trim(objtext.value) < 0 then
						'	MsgBox ("Enter Numeric Value")
						'	objText.select
						'	exit function
						'Else
							oNodTemp.Attributes.Item(4).nodeValue=objText.value
							oNodTemp.Attributes.Item(5).nodeValue=objText.value
						'End if
					elseif sTaxMode="P" then
						oNodTemp.Attributes.Item(4).nodeValue = CalculateTax(sFormula,dBasicTotal,dTotal,objText.value,PurchType)
					else ' K & Q case
						oNodTemp.Attributes.Item(4).nodeValue=objText.value
						oNodTemp.Attributes.Item(5).nodeValue=objText.value
					end if
					popTax
					exit function
				end if
			Next
		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if
	end if
END FUNCTION
'-------------------------------------------------------------------------------------------
Function window_onunload()
	Set window.returnvalue= ItemTaxData.documentElement
End Function
'-------------------------------------------------------------------------------------------

</script>
</HEAD>
<%
Dim sPurTypeName

Dim nPtype

Dim rsTemp

nPtype =  Request.QueryString("PurType")

set rsTemp = Server.CreateObject("ADODB.RecordSet")

with rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT PURTYPESHORTNAME,PURCHASETYPENAME FROM APP_M_PURCHASETYPES  where upper(isNull(Active,'Y')) = 'Y'  ORDER BY PURCHASETYPE"
	.ActiveConnection = con
	.Open
end with
set rsTemp.ActiveConnection = nothing

sPurTypeName = ""
if not rsTemp.EOF then
	sPurTypeName = rsTemp(1)
end if
rsTemp.Close
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload=init()>

<form method="POST" name="formname" action="">
<input type=hidden name="hPurType" value="<%=nPtype%>">
<input type=hidden name="hPurTypeName" value="<%=sPurTypeName%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
          Purchase Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
<TABLE id="Table16" cellSpacing=0 cellPadding=0 border=0 width="100%"  >
<TR>
<TD class="TabBodyWithTopLine">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
					<td align="center" colspan="3" class="MiddlePack" height="7" width="600">
					</td>
                </tr>
                <tr>
					<td align="center" width="5" class="ClearPixel" height="2">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
					<td align="center" class="ClearPixel" width="6" height="2">
                        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
                </tr>
                <tr>
					<td align="center" colspan="3" class="MiddlePack" height="7" width="600">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
                </tr>
                <tr>
					<td align="center" width="5" class="ClearPixel" height="2">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
					<td valign="top">
                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                <!---------------------------------------------->
                <tr>
					<td  class="FieldCell">Item Name</td>
					<td  class="FieldCellSub">
					<span id="spnItemName" class="Dataonly"></span></td>
				 </tr>
				<tr>
					<td  class="FieldCell">Purchase Type</td>
					<td  class="FieldCellSub">
					<span id="spnPurType" class="Dataonly"></span></td>
				 </tr>
				 <tr>
					<td  class="FieldCell">Basic Value</td>
					<td  class="FieldCellSub">
					<span id="spnBasicValue" class="Dataonly"></span></td>
				 </tr>
				 <tr>
					<td  class="FieldCell">Nett Basic Value</td>
					<td  class="FieldCellSub">
					<span id="spnAmount" class="Dataonly"></span></td>
				 </tr>
				<!------------------------------------------------>
                  <tr>
                    <td valign="top" colspan="2">
                      <div class="frmbody" id="frm1" style="width:450;">
                        <table border="0" cellspacing="1" class="ExcelTable" id="tblTaxDet" width="100%">
                          <tr>
                            <td class="ExcelHeaderCell" align="center" width=5 >
                              <p align="center">S.No</td>
                            <td class="ExcelHeaderCell"  align="center" >Tax<br>Name</td>
                            <td class="ExcelHeaderCell" align="center" >Tax<br>Percentage
                            </td>
                            <td class="ExcelHeaderCell"  align="center" >Tax<br>Value
                            </td>
						</tr>

				       </table>
                      </div>
                    </td>
                  </tr>
                </table>
								</td>
								<td align="center" class="ClearPixel" width="6" height="2">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                                <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<input type="button" value="Done" name="B8" onClick="window.close()"  class="ActionButton" tabindex="3" >
															<input type="button" value="Cancel" name="B10"  onClick="window.close()" class="ActionButton" tabindex="3" >
															<input type="reset" value="Reset" name="B9" class="ActionButton" tabindex="3" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="6">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="BottomPack" colspan="3">
								</td>
                                </tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

