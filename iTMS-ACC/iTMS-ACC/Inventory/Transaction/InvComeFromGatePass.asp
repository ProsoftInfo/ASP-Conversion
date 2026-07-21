<%	option explicit	%>
<%
	'Program Name				:	InvComeFromGatepass.asp
	'Module Name				:	Gate Pass - Service
	'Author Name				:	Ragavendran R
	'Created On					:	August 26,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:
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
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>

<!-- #include file="../../include/DatabaseConnection.asp"-->
<!-- #include File="../../include/PurchaseTermsConditions.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/purpopulate.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/PurChkItemSpecPack.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<!---XML Data Island---->
<script type="application/xml" data-itms-xml-island="1" id="TempData"> <Root/> </script>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="ITEMData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="TaxFormData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="PurTypeData">
</script>
<script type="application/xml" data-itms-xml-island="1" id="ItemTaxData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="InvoiceDet" data-src="<%="../temp/transaction/InvItemValue_PUR_"&Session.SessionID&".xml"%>"></script>

<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/calcAlternateUoM.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></SCRIPT>


<SCRIPT type="text/plain" data-itms-legacy-client-script="1">

Dim sInvAgainst,sInvRefNum,bFlag
Dim TaxRoot,RootNode,oNodTemp,ItemRoot, dBasicVal,dTBasicVal
'------------------------------------------------------------------------------------------
function GetAddiDesc(sPassClassCode,sPassItemCode,sPassEntryNo)
Dim sAddDesc

	set ndSample = showModalDialog("invPurInvEntry_ItemAddDesc.asp?ClassCode=" + trim(sPassClassCode) + "&ItemCode=" + trim(sPassItemCode) ,InvoiceDet,"dialogHeight:200px;dialogWidth:450px;status:no")

	set Root = ndSample
	sAddDesc = Root.getAttribute("ItemAddiDesc")

	if trim(sAddDesc) <> "" then
		Set Root=InvoiceDet.documentElement

		if Root.hasChildNodes then
			For each Node1  in Root.ChildNodes
				If Node1.NodeName = "ItemDetails" then
					For Each NodeDet in Node1.ChildNodes
						if trim(NodeDet.getAttribute("ItemCode")) = trim(sPassItemCode) and trim(NodeDet.getAttribute("ClassificationCode")) = trim(sPassClassCode) and trim(NodeDet.getAttribute("SourceEntryNo"))=trim(sPassEntryNo)  then
							NodeDet.setAttribute "ItemAddiDesc",sAddDesc
						end if
					Next
				End If 'If Node1.NodeName = "ItemDetails" then
			Next 'For each Node1  in Root.ChildNodes

		end if 'if Root.hasChildNodes then
	end if 'if trim(sAddDesc) <> "" then

end function
'------------------------------------------------------------------------------------------
Function Amount_onChange()
	Dim Root
	set Root = InvoiceDet.documentElement
CalculateItemValue sPurType
if Root.hasChildNodes then
	Set ItemNode=Root.Selectnodes("//ItemDetails/Item")
	for i = 0 to ItemNode.Length - 1
			ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
			ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
			iEntryNo = ItemNode.Item(i).Attributes.getNamedItem("SourceEntryNo").value
			if document.formname.chkConsolidated.checked=true then
				ItemNode.Item(i).Attributes.getNamedItem("Rate").value =document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("RatePerQtyUoM").value=document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("Amount").value=document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("ItemValue").value=document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("ItemRate").value=document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value=document.formname.txtAmount.value
			else
				ItemNode.Item(i).Attributes.getNamedItem("Rate").value ="0"
				ItemNode.Item(i).Attributes.getNamedItem("RatePerQtyUoM").value="0"
				ItemNode.Item(i).Attributes.getNamedItem("Amount").value="0"
				ItemNode.Item(i).Attributes.getNamedItem("ItemValue").value="0"
				ItemNode.Item(i).Attributes.getNamedItem("ItemRate").value="0"
				ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value="0"
			end if
	next
popTax
end if  'if Root.hasChildNodes then
End Function
'------------------------------------------------------------
function AddItem()
Dim oRow,headerCell,Root,iCtr,ItemNode,sExp,i,Node,InvQty,Rate
Dim DisPer,DisAmount,NettBasic,ItemCode,ClassCode,TaxMode,TaxVal,TaxPer,TaxName
Dim sCatCode,sTaxCode,UomCode,sPurchaseType,iEntryNo

ClearTable
Set Root=InvoiceDet.documentElement
iCtr=1

sPurchaseType = document.Formname.cmbPurType.value

if Root.hasChildNodes then
	Set ItemNode=Root.Selectnodes("//ItemDetails/Item")
	for i = 0 to ItemNode.Length - 1
			ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
			ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
			InvQty = ItemNode.Item(i).Attributes.getNamedItem("Qty").value
			Rate=FormatNumber(ItemNode.Item(i).Attributes.getNamedItem("Rate").value,5,,,0)
			DisPer=ItemNode.Item(i).Attributes.getNamedItem("DisPer").value
			DisAmount=ItemNode.Item(i).Attributes.getNamedItem("DisAmount").value
			NettBasic=ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value
			UomCode=ItemNode.Item(i).Attributes.getNamedItem("Uom").value
			iEntryNo = ItemNode.Item(i).Attributes.getNamedItem("SourceEntryNo").value

			set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)
			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtSerA"& ClassCode &"A"& ItemCode &"A"& iEntryNo&""" value="""&iCtr&""" size=""1"" class=""FormelemRead"" style=""text-align=center"" READONLY>" )
			headerCell.appendChild(oText)
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""checkbox""  name=""chkDeleteA"& ClassCode &"A"& ItemCode &"A"& iEntryNo &""" value=""Y""  class=""FormElem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"

			sItemDesc = ItemNode.Item(i).Attributes.getNamedItem("ItmDescription").value
			set headerCell=oRow.insertCell()
			headerCell.innerHTML="<a class='ExcelDisplayLink' href=# onClick=GetAddiDesc(" & ClassCode & "," & ItemCode & ","& iEntryNo&") >" & sItemDesc & "</a>"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Left"

			set headerCell=oRow.insertCell()
			headercell.innerHtml = UomCode + "<input type=""hidden"" name=""hRateUOM"& ClassCode &"Z"& ItemCode&"Z"& iEntryNo &"""  >"
			headercell.innerHtml = headercell.innerHtml + "<input type=""hidden"" name=""hRatePerQtyUoM"& ClassCode &"Z"& ItemCode&"Z"& iEntryNo &"""  >"
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtInvQty"& ClassCode &"Z"& ItemCode &"Z"& iEntryNo &""" value="& InvQty &" onBlur=""DisplayAmount('Q','"& ItemCode &"','"& ClassCode &"','"& iEntryNo &"',this)"" style=""text-align: Right"" size=""10"" class=""Formelem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelInputCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtInvRate"& ClassCode &"Z"& ItemCode &"Z"& iEntryNo &""" value="& Rate &" onBlur=""DisplayAmount('R','"& ItemCode &"','"& ClassCode &"','"& iEntryNo &"',this)""  style=""text-align: Right"" size=""11"" class=""Formelem"">")
			headerCell.appendChild(oText)
			headercell.innerHtml = headercell.innerHtml + "&nbsp;&nbsp;<a><Img border=""0"" src=""../../assets/images/iTMS%20Icons/EntryIcon.gif"" onclick=""SetRateUOM('"& trim(ClassCode) & "','" & trim(ItemCode) &"','" & UomCode & "','SP','"& iEntryNo&"')"" alt=""Select Rate UOM"" style=""cursor:hand""></a>"
			headerCell.width = "100"
			headerCell.className="ExcelInputCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtValue"& ClassCode &"Z"& ItemCode&"Z"& iEntryNo &""" value="""" size=""11"" style=""text-align: Right"" readonly class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="FormelemRead"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtDisPer"& ClassCode &"Z"& ItemCode&"Z"& iEntryNo & """ value="& DisPer &" onBlur=""DisplayAmount('D','"& ItemCode &"','"& ClassCode &"','"& iEntryNo &"',this)"" style=""text-align: Right"" size=""6"" class=""Formelem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelInputCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtDisAmount"& ClassCode &"Z"& ItemCode &"Z"& iEntryNo &""" value="& DisAmount &" Readonly maxlength=13  style=""text-align: Right"" class=""FormelemRead"" size=""11"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"


			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtNetValue"& ClassCode &"Z"& ItemCode &"Z"& iEntryNo&""" value="& NettBasic &" size=""11"" style=""text-align: Right"" readonly class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"


			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" readonly name=""txtItemRateDisplay" & ClassCode & "Z" & ItemCode &"Z"& iEntryNo& """ size=""11""  style=""text-align: Right"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""checkbox"" name=""chkVat" & ClassCode & "Z" & ItemCode &"Z"& iEntryNo & """ class=""Formelem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			iCtr=iCtr+1

	next
	'To display Total
	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "6"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml = "[&nbsp;<input type=Checkbox name=chkConsolidated onClick=chkConsolidated_onClick()><Font color=Yellow>Enter Conslidated Value</Font>&nbsp;&nbsp;]"
	headerCell.innerhtml= headerCell.innerhtml & "<B>Total</B>"
	headerCell.align="Right"

'	set headerCell=oRow.insertCell()
'	headerCell.colspan = "6"
'	headerCell.className="ExcelSerial"
'	headerCell.innerhtml= "<B>Total</B>"
'	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  readonly  name=""txtBasicValue"" size=""11"" style=""text-align: Right"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Center"

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"


	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text"" readonly  name=""txtDisValue"" size=""11"" style=""text-align: Right"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Center"


	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  name=""txtAmount"" size=""11"" style=""text-align: Right""  onChange=""Amount_onChange()""  class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Center"

	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text"" readonly name=""txtTotalItemRate" & sPurchaseType & """ size=""11""  style=""text-align: Right"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Center"

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"

	'To display tax details
	Set ItemNode=Root.Selectnodes("//TaxDetails/Tax")
	for i = 0 to ItemNode.Length - 1

		TaxName=ItemNode.Item(i).text
		sCatCode=trim(ItemNode.Item(i).Attributes.getNamedItem("CatCode").value )
		sTaxCode=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxCode").value )
		TaxMode=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxMode").value )
		if TaxMode="F" then
			TaxPer=""
			TaxVal=trim(ItemNode.Item(i).Attributes.getNamedItem("TaxValue").value )
		else
			TaxVal=""
			TaxPer=ItemNode.Item(i).Attributes.getNamedItem("TaxValue").value
		end if
		set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

		set headerCell=oRow.insertCell()
		headerCell.colspan = "8"
		headerCell.className="ExcelSerial"
		headerCell.innerhtml=TaxName
		headerCell.align="Right"

		set headerCell=oRow.insertCell()
		if TaxMode<>"F" then
			set oText = document.createElement("<input type=""text"" name=""txtTaxPer" & sCatCode & sTaxCode & """ value="""& TaxPer &""" size=""6"" onBlur=""setTaxPercentage('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" class=""Formelem"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.innerhtml=headerCell.innerhtml + " %"
		end if
		headerCell.className="ExcelFieldCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		if TaxMode="F" then
			set oText = document.createElement("<input type=""text"" name=""txtTaxValue" & sCatCode & sTaxCode & """ value="""& TaxVal &""" size=""11"" onBlur=""setTaxAmount('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" class=""Formelem"">")
		else
			set oText = document.createElement("<input type=""text"" name=""txtTaxValue" & sCatCode & sTaxCode & """ value="""& TaxVal &""" size=""11"" onBlur=""setTaxAmount('"& sCatCode &"','"& sTaxCode &"',this)"" style=""text-align: Right"" ReadOnly class=""FormelemRead"">")
		end if
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelFieldCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		set oText = document.createElement("<input type=""text""  readonly  name=""txtSubTaxValue" & sCatCode & sTaxCode & """ size=""11"" style=""text-align: Right"" class=""FormelemRead"">")
		headerCell.appendChild(oText)
		headerCell.width = "10"
		headerCell.className="ExcelAverageCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		headerCell.width = "10"
		headerCell.className="ExcelAverageCell"

	next

end if 'if Root.hasChildNodes then

	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "9"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml="<B>Rounded Off</B>("
	set oText = document.createElement("<input type=""radio"" value=""Y"" name=""rdRndOff"" onclick=""RoundOffInv()"" class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.innerhtml=headerCell.innerhtml + "Yes"
	set oText = document.createElement("<input type=""radio"" value=""N"" name=""rdRndOff"" onclick=""RoundOffInv()"" class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.innerhtml = headerCell.innerhtml + "No)"
	headerCell.align="Right"


	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text"" NAME=""txtRoundOff"" style=""text-align: Right"" size=""11"" ReadOnly onBlur=""AssignRoundOffValue()"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "9"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml="<B>Invoice Value</B>"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  readonly  NAME=""txtInvValue"" style=""text-align: Right"" size=""11"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	document.all.ImgDeleteIcon.disabled= false

	if document.formname.cmbPurType.selectedIndex = 0 then
		alert("Select Purchase Type")
		document.formname.cmbPurType.focus
		exit function
	end if

	if Root.hasChildNodes then
		Set ItemNode=Root.Selectnodes("//ItemDetails/Item")
		for i = 0 to ItemNode.Length - 1
			ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
			ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
			iEntryNo = ItemNode.Item(i).Attributes.getNamedItem("SourceEntryNo").value

			set obj = eval("document.formname.txtInvRate"&ClassCode&"Z"&ItemCode&"Z"&iEntryNo)
			DisplayAmount "R",ItemCode ,ClassCode,iEntryNo,obj
		next
	end if 'end if 'if Root.hasChildNodes then

end function
'*****************************************************************

'-------------------------------------------------------------------------------------------------
Function chkConsolidated_onClick()
Dim Root,ItemNode
Dim ItemCode,ClassCode,iEntryNo,i
Dim sPurType
sPurType = document.formname.cmbPurType(document.formname.cmbPurType.selectedIndex).value
Set Root=InvoiceDet.documentElement
	if Root.hasChildNodes then
		Set ItemNode=Root.Selectnodes("//ItemDetails/Item")
		for i = 0 to ItemNode.Length - 1
				ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
				ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
				iEntryNo = ItemNode.Item(i).Attributes.getNamedItem("SourceEntryNo").value

				if document.formname.chkConsolidated.checked=true then

		'			ItemNode.Item(i).Attributes.getNamedItem("Rate").value ="0"
		'			ItemNode.Item(i).Attributes.getNamedItem("RatePerQtyUoM").value="0"
		'			ItemNode.Item(i).Attributes.getNamedItem("Amount").value="0"
		'			ItemNode.Item(i).Attributes.getNamedItem("ItemValue").value="0"
		'			ItemNode.Item(i).Attributes.getNamedItem("ItemRate").value="0"
		'			ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value="0"

					eval("document.formname.txtInvRate"&cstr(ClassCode)&"Z"&cstr(ItemCode)&"Z"&cstr(iEntryNo)).disabled = true
					eval("document.formname.txtDisPer"&cstr(ClassCode)&"Z"&cstr(ItemCode)&"Z"&cstr(iEntryNo)).disabled = true
					eval("document.formname.txtInvRate"&cstr(ClassCode)&"Z"&cstr(ItemCode)&"Z"&cstr(iEntryNo)).value = "0"
					eval("document.formname.txtDisPer"&cstr(ClassCode)&"Z"&cstr(ItemCode)&"Z"&cstr(iEntryNo)).value = "0"
					eval("document.formname.txtValue"&cstr(ClassCode)&"Z"&cstr(ItemCode)&"Z"&cstr(iEntryNo)).value = "0"
					eval("document.formname.txtNetValue"&cstr(ClassCode)&"Z"&cstr(ItemCode)&"Z"&cstr(iEntryNo)).value = "0"
					eval("document.formname.txtItemRateDisplay"&cstr(ClassCode)&"Z"&cstr(ItemCode)&"Z"&cstr(iEntryNo)).value = "0"

					eval("document.formname.txtAmount").className="FormElem"
				elseif document.formname.chkConsolidated.checked=False then
					eval("document.formname.txtInvRate"&cstr(ClassCode)&"Z"&cstr(ItemCode)&"Z"&cstr(iEntryNo)).disabled = false
					eval("document.formname.txtDisPer"&cstr(ClassCode)&"Z"&cstr(ItemCode)&"Z"&cstr(iEntryNo)).disabled = false
					eval("document.formname.txtAmount").className="FormElemRead"
					eval("document.formname.txtAmount").value="0"
				end if
		next

	end if  'if Root.hasChildNodes then

CalculateItemValue sPurType

if Root.hasChildNodes then
	Set ItemNode=Root.Selectnodes("//ItemDetails/Item")
	for i = 0 to ItemNode.Length - 1
			ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
			ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
			iEntryNo = ItemNode.Item(i).Attributes.getNamedItem("SourceEntryNo").value
			if document.formname.chkConsolidated.checked=true then
				ItemNode.Item(i).Attributes.getNamedItem("Rate").value =document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("RatePerQtyUoM").value=document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("Amount").value=document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("ItemValue").value=document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("ItemRate").value=document.formname.txtAmount.value
				ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value=document.formname.txtAmount.value
			else
				ItemNode.Item(i).Attributes.getNamedItem("Rate").value ="0"
				ItemNode.Item(i).Attributes.getNamedItem("RatePerQtyUoM").value="0"
				ItemNode.Item(i).Attributes.getNamedItem("Amount").value="0"
				ItemNode.Item(i).Attributes.getNamedItem("ItemValue").value="0"
				ItemNode.Item(i).Attributes.getNamedItem("ItemRate").value="0"
				ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value="0"
			end if
	next
end if  'if Root.hasChildNodes then

End Function
'******************************************
function PurchaseTypeWise()
dim oRow,headerCell,Root,iCtr,ItemNode,sExp,i,Node,InvQty,Rate,DisPer,DisAmount,NettBasic,ItemCode,ClassCode,TaxMode,TaxVal,TaxPer,TaxName
dim TaxNode,PurType,j,nTempLastItemCtr
ClearTable
Set Root=InvoiceDet.documentElement
iCtr=1
nTempLastItemCtr = 0


if Root.hasChildNodes then
	Set TaxNode=Root.Selectnodes("//TaxDetails")
	for j = 0 to TaxNode.Length - 1
		PurType=TaxNode.Item(j).Attributes.getNamedItem("PurchaseType").value
		Set ItemNode=Root.Selectnodes("//ItemDetails/Item[@PurchaseType="& PurType &"]")
		for i = 0 to ItemNode.Length - 1
			nTempLastItemCtr = nTempLastItemCtr + 1
		Next
	Next
end if

if Root.hasChildNodes then
	Set TaxNode=Root.Selectnodes("//TaxDetails")
	for j = 0 to TaxNode.Length - 1
		PurType=TaxNode.Item(j).Attributes.getNamedItem("PurchaseType").value

		Set ItemNode=Root.Selectnodes("//ItemDetails/Item[@PurchaseType="& PurType &"]")
		for i = 0 to ItemNode.Length - 1
				ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
				ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
				InvQty = ItemNode.Item(i).Attributes.getNamedItem("Qty").value
				Rate=FormatNumber(ItemNode.Item(i).Attributes.getNamedItem("Rate").value,5,,,0)
				DisPer=ItemNode.Item(i).Attributes.getNamedItem("DisPer").value
				DisAmount=ItemNode.Item(i).Attributes.getNamedItem("DisAmount").value
				NettBasic=ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value
				iEntryNo = ItemNode.Item(i).Attributes.getNamedItem("SourceEntryNo").value

				UomCode=ItemNode.Item(i).Attributes.getNamedItem("Uom").value

				set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)
				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtSerA"& ClassCode &"A"& ItemCode &"A"& iEntryNo &""" value="""&iCtr&""" size=""1"" class=""FormelemRead"" style=""text-align=center"" READONLY>" )
				headerCell.appendChild(oText)
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""checkbox""  name=""chkDeleteA"& ClassCode &"A"& ItemCode &"A"& iEntryNo &""" value=""Y""  class=""FormElem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"

				sItemDesc = ItemNode.Item(i).Attributes.getNamedItem("ItmDescription").value
				set headerCell=oRow.insertCell()
				headerCell.innerHTML="<a class='ExcelDisplayLink' href=# onClick=GetAddiDesc(" & ClassCode & "," & ItemCode & ","& iEntryNo&") >" & sItemDesc & "</a>"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="Left"

				set headerCell=oRow.insertCell()
				headercell.innerHtml = UomCode + "<input type=""hidden"" name=""hRateUOM"& ClassCode &"Z"& ItemCode &"""  >"
				headercell.innerHtml = headercell.innerHtml + "<input type=""hidden"" name=""hRatePerQtyUoM"& ClassCode &"Z"& ItemCode&"Z"& iEntryNo &"""  >"
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtInvQty"& ClassCode &"Z"& ItemCode&"Z"& iEntryNo &""" value="& InvQty &" onBlur=""DisplayPurAmount('Q','"& ItemCode &"','"& ClassCode &"','& iEntryNo &',this)"" style=""text-align: Right"" size=""10"" class=""Formelem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelInputCell"
				headerCell.align="center"


				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtInvRate"& ClassCode &"Z"& ItemCode&"Z"& iEntryNo &""" value="& Rate &" onBlur=""DisplayPurAmount('R','"& ItemCode &"','"& ClassCode &"','"& iEntryNo&"',this)""  style=""text-align: Right"" size=""11"" class=""Formelem"">")
				headerCell.appendChild(oText)
				headercell.innerHtml = headercell.innerHtml + "&nbsp;&nbsp;<a><Img border=""0"" src=""../../assets/images/iTMS%20Icons/EntryIcon.gif"" onclick=""SetRateUOM('"& trim(ClassCode) & "','" & trim(ItemCode) &"','" & UomCode & "','MP','"& iEntryNo &"')"" alt=""Select Rate UOM"" style=""cursor:hand""></a>"
				headerCell.width = "100"
				headerCell.className="ExcelInputCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtValue"& ClassCode &"Z"& ItemCode&"Z"& iEntryNo &""" value="""" size=""11"" style=""text-align: Right"" readonly class=""FormelemRead"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="FormelemRead"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtDisPer"& ClassCode &"Z"& ItemCode &"Z"& iEntryNo &""" value="& DisPer &" onBlur=""DisplayPurAmount('D','"& ItemCode &"','"& ClassCode &"','"& iEntryNo &"',this)"" style=""text-align: Right"" size=""6"" class=""Formelem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelInputCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtDisAmount"& ClassCode &"Z"& ItemCode &"Z"& iEntryNo &""" value="& DisAmount &" ReadOnly maxlength=13  style=""text-align: Right"" class=""FormelemRead"" size=""11"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"


				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtNetValue"& ClassCode &"Z"& ItemCode &"Z"& iEntryNo &""" value="& NettBasic &" size=""11"" style=""text-align: Right"" readonly class=""FormelemRead"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"


				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" readonly name=""txtItemRateDisplay" & ClassCode & "Z" & ItemCode &"Z"& iEntryNo & """ size=""11""  style=""text-align: Right"" class=""FormelemRead"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""checkbox"" name=""chkVat" & ClassCode & "Z" & ItemCode &"Z"& iEntryNo & """ class=""Formelem"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				iCtr=iCtr+1
		next

		if ItemNode.Length > 0 then

			'To display Total
			set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.colspan = "6"
			headerCell.className="ExcelSerial"
			headerCell.innerhtml="<B>Total</B>"
			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" ReadOnly name=""txtBasicValue"& PurType &""" size=""11"" style=""text-align: Right"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Center"

			set headerCell=oRow.insertCell()
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"


			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text""  ReadOnly  name=""txtDisValue"& PurType &""" size=""11"" style=""text-align: Right"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Center"


			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text""  ReadOnly  name=""txtAmount"& PurType &""" size=""11"" style=""text-align: Right"" onBlur=""Amount_onChange()"" class=""FormelemRead"" >")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" readonly name=""txtTotalItemRate" & PurType & """ size=""11"" style=""text-align: Right""  class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Center"

			'Adding Total And subTotal
			set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.colspan = "9"
			headerCell.className="ExcelSerial"
			headerCell.innerhtml="Total Tax Value"
			set oText = document.createElement("<img name=""imgTaxDet"& PurType &""" border=""0"" src=""../../assets/images/iTMS%20Icons/EntryIcon.gif"" alt=""View Tax Details"" onclick=""ShowTaxDet('"& PurType  &"')"" width=""15"" height=""15"" style=""cursor:hand"">")
			headerCell.appendChild(oText)

			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" NAME=""txtTot"& PurType &""" style=""text-align: Right"" size=""11"" readonly class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" NAME=""txtTaxDisplay"& PurType &""" style=""text-align: Right"" size=""11"" readonly class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Right"

			set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.colspan = "9"
			headerCell.className="ExcelSerial"
			headerCell.innerhtml="Sub Total"
			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text""  ReadOnly  NAME=""txtSubTot"& PurType &""" style=""text-align: Right"" size=""11"" class=""FormelemRead"">")
			headerCell.appendChild(oText)
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Right"

			set headerCell=oRow.insertCell()
			headerCell.width = "10"
			headerCell.className="ExcelDisplayCell"
			headerCell.align="Right"

		end if ' if ItemNode.Length > 0 then
	next
end if 'if Root.hasChildNodes then

	'Adding InvoiceValue and Roundoff
	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "9"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml="<B>Rounded Off</B>("
	set oText = document.createElement("<input type=""radio"" value=""Y"" onclick=""RoundOffInv()"" name=""rdRndOff"" class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.innerhtml=headerCell.innerhtml + "Yes"
	set oText = document.createElement("<input type=""radio"" value=""N"" onclick=""RoundOffInv()"" name=""rdRndOff"" class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.innerhtml = headerCell.innerhtml + "No)"
	headerCell.align="Right"


	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text"" NAME=""txtRoundOff"" ReadOnly onBlur=""AssignRoundOffValue()"" style=""text-align: Right"" size=""11""  class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set oRow = document.all.tblItemDet.insertRow(document.all.tblItemDet.rows.length)

	set headerCell=oRow.insertCell()
	headerCell.colspan = "9"
	headerCell.className="ExcelSerial"
	headerCell.innerhtml="<B>Invoice Value</B>"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  ReadOnly  Name=""txtInvValue"" style=""text-align: Right"" size=""11"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"

	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text""  ReadOnly  Name=""mTxtTotal"" size=""11""  style=""text-align: Right"" class=""FormelemRead"">")
	headerCell.appendChild(oText)
	headerCell.width = "10"
	headerCell.className="ExcelDisplayCell"
	headerCell.align="Right"


	document.all.ImgDeleteIcon.disabled= false

	if Root.hasChildNodes then
		Set ItemNode=Root.Selectnodes("//ItemDetails/Item")
		for i = 0 to ItemNode.Length - 1
			ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
			ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value

			set obj = eval("document.formname.txtInvRate"&ClassCode&"Z"&ItemCode)
			DisplayPurAmount "R",ItemCode ,ClassCode,iEntryNo,obj
		next
	end if 'if Root.hasChildNodes then

end function
'-------------------------------------------------------------------------------------------
function SetRateUOM(iClassCode,iItemCode,sUOM,sPurchaseTypeNature,iEntryNo)
Dim sOrgID,sRateUOM

Dim nRatePerQtyUOM

sOrgID = document.formname.hOrgID.value

set obj = eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&iEntryNo )
nRatePerQtyUOM = obj.value

set obj = eval("document.formname.hRateUOM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&iEntryNo )
sRateUOM = obj.value

if trim(sRateUOM) = "" then sRateUOM = sUOM

Set ndSample= showModalDialog("invPurInvEntry_RateUomPop.asp?hOrgID=" & trim(sOrgID) & "&hClassCode="+trim(iClassCode)+"&hItemCode="+trim(iItemCode)+"&hUOM=" + trim(sUOM)+"&hRateUOM=" + trim(sRateUOM) +"&hRatePerQtyUOM="+trim(nRatePerQtyUOM),"","dialogHeight:250px;dialogWidth:300px;status:no")
'alert(ndSample.xml)

set Root = ndSample
sRateUOM = ""
nRatePerQtyUoM = 0.00

if trim(Root.getAttribute("RateUOM")) <> "" then
	sRateUOM = Root.getAttribute("RateUOM")
	nRatePerQtyUoM = Root.getAttribute("RatePerQtyUoM")
else
	sRateUOM = sUOM
	nRatePerQtyUoM = eval("document.formname.txtInvRate"&iClassCode&"Z"&iItemCode&"Z"&iEntryNo).value
end if

set obj = eval("document.formname.hRateUOM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&iEntryNo )
obj.value = sRateUOM
set obj = eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&iEntryNo )
obj.value = nRatePerQtyUoM

'alert(sRateUOM)
'alert(nRatePerQtyUoM)

getQtyUoMRate sOrgID,iClassCode,iItemCode,eval("document.formname.hRateUoM"&iClassCode&"Z"&iItemCode),sUOM

set obj = eval("document.formname.txtInvRate"&iClassCode&"Z"&iItemCode)
if UCase(Trim(sPurchaseTypeNature)) = "SP" then 'single purchase type
	DisplayAmount "R",iItemCode,iClassCode,iEntryNo,obj
else	' item wise purchase type
	DisplayPurAmount "R",iItemCode,iClassCode,iEntryNo,obj
end if
end function
'-------------------------------------------------------------------------------------------

Function getQtyUoMRate(ORGID,CLASSCODE,ITEMCODE,OBJ,sUOM)

Dim iOptToBaseRate,iOptToBaseOperator,QTYUOM,RATE,iCtr,dQTY,dRatePerQtyUoM,dRateTotalQtyUoM

'QTYUOM = eval("document.formname.cmbUoM"&CLASSCODE&"Z"&ITEMCODE).value
QTYUOM = sUOM
arrUoM = split(QTYUOM,":")
QTYUOM = arrUoM(0)

RATE = eval("document.formname.txtInvRate"&CLASSCODE&"Z"&ITEMCODE).value
dQTY = eval("document.formname.txtInvQty"&CLASSCODE&"Z"&ITEMCODE).value

RATEUOM = OBJ.value
if trim(RATEUOM) <> "" then
	arrUoM = split(RATEUOM,":")
	RATEUOM = arrUoM(0)
end if 'if trim(RATEUOM) <> "" then

if trim(RATE) = "" then RATE = 0
if (CLASSCODE = "0" or CLASSCODE="TEMP") then
	dRatePerQtyUoM = RATE
else
	If trim(RATEUOM) <> trim(QTYUOM) Then
		dRatePerQtyUoM = getRatePerQtyUoM(ORGID,CLASSCODE,ITEMCODE,QTYUOM,RATEUOM,RATE)
		'alert("test ....")
	Else
		dRatePerQtyUoM = RATE
	End If
end if

eval("document.formname.txtInvRate"&CLASSCODE&"Z"&ITEMCODE).value  = formatnumber(dRatePerQtyUoM,5,,,0)

End function
'-------------------------------------------------------------------------------------------
Function DisplayAmount(sPara,iItemCode,iClassCode,iEntryNo,objText)

dim dBasicTotal,dDisTotal,dNetTotal,dRate,dQty,dDisPer


	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="ItemDetails" then
			set ItemRoot=oNodTemp
		end if
	next

	For Each oNodTemp in ItemRoot.childNodes
		sCurrentItemPurType = trim(oNodTemp.Attributes.getNamedItem("PurchaseType").value )
	Next

	'alert(sCurrentItemPurType)
	if trim(sCurrentItemPurType) = "" or trim(sCurrentItemPurType) = "0" then
		Alert("Select Purchase Type")
		exit function
	end if

	dBasicTotal	= 0
	dDisTotal	= 0
	dNetTotal	= 0

	dQty = 0
	dRate = 0.00
	dDisPer = 0.00

	if sPara = "Q" then dQty = objText.value
	if sPara = "R" then dRate = objText.value
	if sPara = "D" then dDisPer = objText.value

	if trim(objText.value)<>"" then

		if IsNumeric(objText.value) then
			For Each oNodTemp in ItemRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=iItemCode and oNodTemp.Attributes.Item(1).nodeValue=iClassCode and oNodTemp.Attributes.Item(10).nodeValue=iEntryNo then

					if sPara = "Q" then oNodTemp.SetAttribute "Qty",dQty
					if sPara = "R" then

						oNodTemp.SetAttribute "Rate",FormatNumber(dRate,5,,,0)
						if trim(eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&trim(iEntryNo)).value) <> "" then
							oNodTemp.SetAttribute "RatePerQtyUoM",	eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&trim(iEntryNo)).value
							sRateUOM = eval("document.formname.hRateUOM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&trim(iEntryNo)).value
							if trim(sRateUOM) <> "" then
								Arr1 = Split(sRateUOM,":")
								sRateUOM = Arr1(0)
							end if
							oNodTemp.SetAttribute "RateUOM", sRateUOM
						else
							oNodTemp.SetAttribute "RatePerQtyUoM",FormatNumber(dRate,5,,,0)
							oNodTemp.SetAttribute "RateUOM", oNodTemp.getAttribute("Uom")
						end if
					end if 'if sPara = "R" then
					if sPara = "D" then oNodTemp.SetAttribute "DisPer",dDisPer


					dQty		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Qty").value ))
					dRate		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Rate").value ))
					dDisPer		= CDbl(trim(oNodTemp.Attributes.getNamedItem("DisPer").value ))

					if dQty > 0 and dRate > 0  then
						dBasicAmount=CDbl(dQty)*CDbl(dRate)

						dDisAmount =round((dDisPer/100)* dBasicAmount,2)
						dNetAmount =round(dBasicAmount - dDisAmount,2)

						oNodTemp.SetAttribute "DisAmount",dDisAmount
						oNodTemp.SetAttribute "NettBasic",dNetAmount

						dBasicAmount=FormatNumber(Round(dBasicAmount,2),2,,,0)
						dDisAmount=FormatNumber(dDisAmount,2,,,0)
						dNetAmount=FormatNumber(dNetAmount,2,,,0)

						eval("document.formname.txtValue"&iClassCode&"Z"&iItemCode&"Z"&iEntryNo).value=dBasicAmount
						eval("document.formname.txtDisAmount"&iClassCode&"Z"&iItemCode&"Z"&iEntryNo).value=dDisAmount
						eval("document.formname.txtNetValue"&iClassCode&"Z"&iItemCode&"Z"&iEntryNo).value=dNetAmount
					end if 'if dQty > 0 and dRate > 0  then


				end if
			next

			For Each oNodTemp in ItemRoot.childNodes

				dQty		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Qty").value ))
				dRate		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Rate").value ))
				dDisPer		= CDbl(trim(oNodTemp.Attributes.getNamedItem("DisPer").value ))

				if dQty > 0 and dRate > 0  then
					dBasicAmount=CDbl(dQty)*CDbl(dRate)

					dDisAmount=round((dDisPer/100)* dBasicAmount,2)

					dNetAmount =  round(dBasicAmount - dDisAmount,2)

					dBasicTotal	= dBasicTotal + dBasicAmount
					dDisTotal	= dDisTotal + dDisAmount
					dNetTotal	= dNetTotal + dNetAmount
				end if 'if dQty > 0 and dRate > 0  then

			next

			document.formname.txtBasicValue.value	= FormatNumber(Round(dBasicTotal,2),2,,,0)
			document.formname.txtDisValue.value		= FormatNumber(Round(dDisTotal,2),2,,,0)
			document.formname.txtAmount.value		= FormatNumber(Round(dNetTotal,2),2,,,0)

			popTax


			For Each oNodTemp in ItemRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=iItemCode and oNodTemp.Attributes.Item(1).nodeValue=iClassCode and oNodTemp.Attributes.Item(10).nodeValue=iEntryNo then
					CalcItemValueForEachItem iItemCode,iClassCode
					if sPara = "R" then
						objText.value= FormatNumber(trim(objText.value),5,,,0)
					elseif sPara = "D" then
						objText.value= FormatNumber(trim(objText.value),2,,,0)
					end if
				end if
			next

			'CalcAmountForLastItem
		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if 'if IsNumeric(objText.value) then

	end if 'if trim(objText.value)<>"" then
End Function

'-------------------------------------------------------------

Function DisplayPurAmount(sPara,iItemCode,iClassCode,iEntryNo,objText)
dim sCurrentItemPurType
dim dBasicTotal,dDisTotal,dNetTotal,dRate,dQty,dDisPer


	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="ItemDetails" then
			set ItemRoot=oNodTemp
		end if
	next

	dBasicTotal	= 0
	dDisTotal	= 0
	dNetTotal	= 0

	dQty = 0
	dRate = 0.00
	dDisPer = 0.00

	if sPara = "Q" then dQty = objText.value
	if sPara = "R" then dRate = objText.value
	if sPara = "D" then dDisPer = objText.value

	if trim(objText.value)<>"" then

		if IsNumeric(objText.value) then
			For Each oNodTemp in ItemRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=iItemCode and oNodTemp.Attributes.Item(1).nodeValue=iClassCode then

					sCurrentItemPurType = trim(oNodTemp.Attributes.getNamedItem("PurchaseType").value )

					if sPara = "Q" then oNodTemp.SetAttribute "Qty",dQty
					if sPara = "R" then
						oNodTemp.SetAttribute "Rate",FormatNumber(dRate,5,,,0)
						if trim(eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&iEntryNo).value) <> "" then
							oNodTemp.SetAttribute "RatePerQtyUoM",	eval("document.formname.hRatePerQtyUoM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&iEntryNo).value
							sRateUOM = eval("document.formname.hRateUOM" & trim(iClassCode) & "Z" & trim(iItemCode)&"Z"&iEntryNo).value
							if trim(sRateUOM) <> "" then
								Arr1 = Split(sRateUOM,":")
								sRateUOM = Arr1(0)
							end if
							oNodTemp.SetAttribute "RateUOM", sRateUOM
						else
							oNodTemp.SetAttribute "RatePerQtyUoM",FormatNumber(dRate,5,,,0)
							oNodTemp.SetAttribute "RateUOM", oNodTemp.getAttribute("Uom")
						end if
					end if 'if sPara = "R" then

					if sPara = "D" then oNodTemp.SetAttribute "DisPer",dDisPer


					dQty		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Qty").value ))
					dRate		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Rate").value ))
					dDisPer		= CDbl(trim(oNodTemp.Attributes.getNamedItem("DisPer").value ))

					if dQty > 0 and dRate > 0  then
						dBasicAmount=CDbl(dQty)*CDbl(dRate)

						dDisAmount = round((dDisPer/100)* dBasicAmount,2)
						dNetAmount = round(dBasicAmount - dDisAmount,2)

						oNodTemp.SetAttribute "DisAmount",dDisAmount
						oNodTemp.SetAttribute "NettBasic",dNetAmount

						dBasicAmount=FormatNumber(Round(dBasicAmount,2),2,,,0)
						dDisAmount=FormatNumber(dDisAmount,2,,,0)
						dNetAmount=FormatNumber(dNetAmount,2,,,0)


						eval("document.formname.txtValue"&iClassCode&"Z"&iItemCode&"Z"&iEntryNo).value=dBasicAmount
						eval("document.formname.txtDisAmount"&iClassCode&"Z"&iItemCode&"Z"&iEntryNo).value=dDisAmount
						eval("document.formname.txtNetValue"&iClassCode&"Z"&iItemCode&"Z"&iEntryNo).value=dNetAmount



					end if 'if dQty > 0 and dRate > 0  then


				end if
			next

			For Each oNodTemp in ItemRoot.childNodes

				If sCurrentItemPurType = trim(oNodTemp.Attributes.getNamedItem("PurchaseType").value ) then

					dQty		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Qty").value ))
					dRate		= CDbl(trim(oNodTemp.Attributes.getNamedItem("Rate").value ))
					dDisPer		= CDbl(trim(oNodTemp.Attributes.getNamedItem("DisPer").value ))

					if dQty > 0 and dRate > 0  then
						dBasicAmount=CDbl(dQty)*CDbl(dRate)

						dDisAmount = round((dDisPer/100)* dBasicAmount,2)

						dNetAmount = round(dBasicAmount - dDisAmount,2)

						dBasicTotal	= dBasicTotal + dBasicAmount
						dDisTotal	= dDisTotal + dDisAmount
						dNetTotal	= dNetTotal + dNetAmount
					end if 'if dQty > 0 and dRate > 0  then

				end if ' if sCurrentItemPurType = trim(oNodTemp.Attributes.getNamedItem("PurchaseType").value )

			next

			eval("document.formname.txtBasicValue" &sCurrentItemPurType).value = FormatNumber(Round(dBasicTotal,2),2,,,0)
			eval("document.formname.txtDisValue" &sCurrentItemPurType).value = FormatNumber(Round(dDisTotal,2),2,,,0)
			eval("document.formname.txtAmount" &sCurrentItemPurType).value = FormatNumber(Round(dNetTotal,2),2,,,0)

			popTax1


			For Each oNodTemp in ItemRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=iItemCode and oNodTemp.Attributes.Item(1).nodeValue=iClassCode then
					CalcItemValueForEachItem iItemCode,iClassCode
					if sPara = "R" then
						objText.value= FormatNumber(trim(objText.value),5,,,0)
					elseif sPara = "D" then
						objText.value= FormatNumber(trim(objText.value),2,,,0)
					end if
				end if
			next

			'CalcAmountForLastItem

		else
			MsgBox ("Enter Numeric Value")
			objText.select
		end if 'if IsNumeric(objText.value) then

	end if 'if trim(objText.value)<>"" then
End Function
'-------------------------------------------------------------
FUNCTION setTaxPercentage(sCatCode,sTaxCode,objText)
	set RootNode=InvoiceDet.documentElement

	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			FOR EACH Node2 in oNodTemp.ChildNodes
				if Node2.NodeName = "Header" then
					sPurchaseType	= Node2.getAttribute("PurchaseType")
				end if 'if Node2.NodeName = "Header" then
			next
		end if

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
					CalculateItemValue sPurchaseType
					'CalcAmountForLastItem
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
FUNCTION setTaxAmount(sCatCode,sTaxCode,objText)

	Dim sTaxCatType

	Set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes

		if oNodTemp.nodeName="InvoiceHeader" then
			FOR EACH Node2 in oNodTemp.ChildNodes
				if Node2.NodeName = "Header" then
					sPurchaseType	= Node2.getAttribute("PurchaseType")
				end if 'if Node2.NodeName = "Header" then
			next
		end if
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	Next

	dInvAmount=document.formname.txtAmount.value
	dBasicTotal=document.formname.txtBasicValue.value
	dTotal	=document.formname.txtAmount.value

	if trim(objText.value)<>"" then
		if IsNumeric(objText.value) then
			For Each oNodTemp in TaxRoot.childNodes
				if oNodTemp.Attributes.Item(0).nodeValue=sCatCode and oNodTemp.Attributes.Item(1).nodeValue=sTaxCode then

					sTaxMode = oNodTemp.Attributes.Item(2).nodeValue
					sFormula = oNodTemp.Attributes.Item(3).nodeValue
					sTaxCatType = oNodTemp.getAttribute("TaxCategoryType")

					if sTaxMode="F" then
							oNodTemp.Attributes.Item(4).nodeValue=objText.value
							oNodTemp.Attributes.Item(5).nodeValue=objText.value
					elseif sTaxMode="P" then

					else ' K & Q case
						oNodTemp.Attributes.Item(4).nodeValue=objText.value
						oNodTemp.Attributes.Item(5).nodeValue=objText.value
					end if
					popTax
					CalculateItemValue sPurchaseType
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

dim dInvAmount,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax
dim nDisplayTotal

	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
	dInvAmount=document.formname.txtAmount.value
	dBasicTotal=document.formname.txtBasicValue.value
	dTotal	=document.formname.txtAmount.value

	if trim(dInvAmount) = "" then dInvAmount = 0
	if trim(dBasicTotal) = "" then dBasicTotal = 0
	if trim(dTotal) = "" then dTotal = 0

	nDisplayTotal = CDbl(dTotal)

	For Each oNodEntry in TaxRoot.childNodes

		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue
		sFormula=oNodEntry.Attributes.Item(3).nodeValue
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue

		If sCatCode <> "0" and sTaxCode <> "0" then		' Except round off node - 09 Aug 04

			if sTaxMode="P" then
				dTax=CalculateTax(sFormula,dBasicTotal,dTotal,dTaxValue,0)
				eval("document.formname.txtTaxPer"&sCatCode&sTaxCode).value=dTaxValue
			elseif sTaxMode = "Q" then
				dTax = cdbl(dTaxvalue) * cdbl(Qtysum)
			elseif sTaxMode = "K" then
				sTotpackvalue = 0
				sExp = "//Tax[@CatCode = "&sCatCode&" and @TaxCode ="&sTaxCode&"]/Taxpack"
				set sNode = TaxRoot.Selectnodes(sExp)
				If sNode.Length > 0 then
					For IMCtr = 0 to sNode.Length - 1
						sTotpackvalue = cdbl(sTotpackvalue) + cdbl(sNode.Item(iMCtr).Attributes.Item(3).nodevalue)
					Next
				end if
				dTax=sTotpackvalue
			else
				dTax=dTaxvalue
			end if

			dTRndoff = oNodEntry.getAttribute("Rndoff")

			if trim(dTRndoff) = "1" then
				dtax = RndOff(dTax)
			End if

			dTax=FormatNumber(dTax,2,,,0)
			oNodEntry.Attributes.Item(5).nodeValue=dTax

			dInvAmount=dInvAmount+CDbl(dTax)

			eval("document.formname.txtTaxValue"&sCatCode&sTaxCode).value=dTax

			nDisplayTotal = nDisplayTotal + dTax
			eval("document.formname.txtSubTaxValue"&sCatCode&sTaxCode).value=FormatNumber(nDisplayTotal,2,,,0)

		end if
	Next


	dRoundedInvvalue = RndOff(dInvamount)

	dRoundedoff = Round(cdbl(dRoundedInvvalue) - cdbl(dInvamount),2)

	dRoundedoff = FormatNumber(dRoundedoff,2,,,0)
	dRoundedInvvalue=FormatNumber(dRoundedInvvalue,2,,,0)

	'alert(dRoundedoff)
	if cdbl(dRoundedoff) <> 0.00 then
		document.formname.rdRndOff(0).Checked=true
	else
		document.formname.rdRndOff(1).Checked=true
	end if

	if document.formname.rdRndOff(0).Checked=true then
		document.formname.txtRoundOff.value = dRoundedoff
		document.formname.txtInvValue.value = dRoundedInvvalue
	else
		document.formname.txtRoundOff.value="0.00"
		document.formname.txtInvValue.value = formatnumber(dInvamount,2,,,0)
	end if


	' set value for round off node - 09 Aug 04
	For Each oNodEntry in TaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue

		If sCatCode = "0" and sTaxCode = "0" then
			oNodEntry.setAttribute "TaxValue",document.formname.txtRoundOff.value
			oNodEntry.setAttribute "TaxAmount",document.formname.txtRoundOff.value
		End if
	Next

	TaxRoot.Attributes.Item(0).nodeValue=dBasicTotal
	TaxRoot.Attributes.Item(1).nodeValue=dTotal
	TaxRoot.setAttribute "TotalTax", dInvamount-dTotal
	TaxRoot.setAttribute "SubTotal", dInvamount
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			set TempNode=oNodTemp
			exit for
		end if
	next

	for Each oNodTemp in TempNode.childNodes
		oNodTemp.setAttribute "InvValue",document.formname.txtInvValue.value
		oNodTemp.setAttribute "RoundOff",document.formname.txtRoundOff.value
	next

END FUNCTION
'-------------------------------------------------------------------------------------------
FUNCTION CalculateTax(sFormula,dBValue,dDValue,dPercentage,sPurType)
dim saTemp,dTaxAmount,iCounter,TaxRoot
dim oNodTemp,iTemp
dim saTemp1

set RootNode=InvoiceDet.documentElement
if trim(sPurType) ="0" then
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
			exit for
		end if
	next
else
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			if trim(oNodTemp.getAttribute("PurchaseType")) = trim(sPurType) then
				set TaxRoot=oNodTemp
				exit for
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

Function ClearTable()
	dim j
	for j=2 to document.all.tblItemDet.rows.length - 1
		document.all.tblItemDet.deleteRow(2)
	next
End Function
'-------------------------------------------------------------------
function getTaxDet()
	Dim ndEntry,ndTax,sTaxName,sFormName,iTaxCode,iCatCode,iFormNo,dPerAmt,sMode
	Dim hPurchaseType,objhttp,sForUnit,objhttp1
	dim RtTax,root,DeleteNode,ItemNode
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set Root=InvoiceDet.documentElement
	'To add Tax details node
	Set DeleteNode = Root.Selectnodes("//TaxDetails")
	for i=0 to (DeleteNode.Length-1)
		root.RemoveChild(DeleteNode.Item(i))
	next
	sForUnit = document.formname.hOrgID.value
	hPurchaseType = document.Formname.cmbPurType.value

	if hPurchaseType=0 then
		document.all.imgPurchaseDet.disabled=false
	else
		document.all.imgPurchaseDet.disabled=true
	end if

	objhttp.Open "GET","XMLGetTaxDetails.asp?PurType="&hPurchaseType&"&ForUnit="&sForUnit,false
	objhttp.send

	If objhttp.responsexml.xml <> "" then
		'msgbox "objhttp.responseXML.xml"+objhttp.responseXML.xml
		TaxFormData.loadXML objhttp.responseXML.xml
		Set RtTax = TaxFormData.documentElement
		If RtTax.hasChildNodes Then
			Set Root=InvoiceDet.documentElement
			For Each ndEntry in RtTax.childNodes
				Root.appendchild ndEntry
			next
		end if
	end if

	'To set Puchasetype for each Item
	if Root.hasChildNodes then

		Set ItemNode = Root.Selectnodes("//InvoiceHeader/Header")
		for i = 0 to ItemNode.Length - 1
			ItemNode.Item(i).Attributes.getNamedItem("PurchaseType").value= hPurchaseType
		next

		Set ItemNode = Root.Selectnodes("//ItemDetails/Item")
		for i = 0 to ItemNode.Length - 1
			ItemNode.Item(i).Attributes.getNamedItem("PurchaseType").value= hPurchaseType
		next

		Set ItemNode = Root.Selectnodes("//TaxDetails/Tax")
		for i = 0 to ItemNode.Length - 1
			ItemNode.Item(i).setAttribute "ItemValue", "0"
		next
	end if

	set objhttp1 = CreateObject("Microsoft.XMLHTTP")
	objhttp1.Open "POST","XMLSavePur.asp?Mod=PUR&Name=InvItemValue", false
	objhttp1.send InvoiceDet.XMLDocument

	additem()
end function
'------------------------------------------------------------------

Function showTaxFormPopUp(sOrgId,sInvDt)
Dim ndSample,iPurTypeVal,sPurType,sPartyName,rootTax,ndEntry

Set rootTax = TaxFormData.documentElement

If document.formname.cmbPurType.value  = "0" Then
	msgbox "Select Purchase Type"
	document.formname.cmbPurType.focus()
	Exit Function
Else
	iPurTypeVal = document.formname.cmbPurType.selectedIndex
	sPurType	= document.formname.cmbPurType(document.formname.cmbPurType.selectedIndex).text
	sPartyName	= document.formname.txtPartyName.value
	Set ndSample= showModalDialog("invPurTaxFormPop.asp?OrgID="+sOrgId+"&PurType="+cstr(iPurTypeVal)+"&PurTypeName="+cstr(sPurType)+"&Party="+sPartyName+"&InvDt='"+cstr(sInvDt)+"'","","dialogHeight:500px;dialogWidth:900px;status:no")
	If ndSample.hasChildNodes Then
		For each ndEntry in ndSample.childNodes
			rootTax.appendchild ndEntry
		Next
	End If
End If
End Function
'----------------------------------------------------------------------------------------
Function ShowPurchaseDet(sOrgID,sItemType)
dim ndTemp,objhttp1,NoOfItems
set RootNode=InvoiceDet.DocumentElement
set objhttp1 = CreateObject("Microsoft.XMLHTTP")
objhttp1.Open "POST","XMLSavePur.asp?Mod=PUR&Name=InvItemValue", false
objhttp1.send InvoiceDet.XMLDocument

Set ndTemp =showModalDialog("popItemPurType.asp",InvoiceDet,"dialogHeight:400px;dialogWidth:650px;status:no;help:no")
set InvoiceDet.DocumentElement=ndTemp
PurchaseTypeWise()
popTax1
end function

'-----------------------------------------------------------------------------------------
Function ShowTaxDet(PurType)
dim ndTemp,ItemNode
Set InvoiceDet.DocumentElement = showModalDialog("popTaxDetails.asp?PurType="+ PurType,InvoiceDet,"dialogLeft:0px,dialogTop:0Px,dialogHeight:600px;dialogWidth:500px;status:no")
popTax1
CalculateItemValue(PurType)
end function

'-----------------------------------------------------------------------------------------

Function getReferenceDetail()
bFlag = False
sInvAgainst = ""
sInvRefNum = ""
if SpnRcptCode.innerText <> "" then
	bFlag = True
	sInvAgainst = "Receipt"
	sInvRefNum = document.formname.hRcptno.value
End if
End Function
'------------------------------------------------------------------------------------------
Function displayItemDetail(sOrgMD,sFlag)
Dim objHttp,Root,ndHeader,ndItem
Dim sPartyName, dQty
Dim iClassCode,iItemCode,iQtyRecd,sClassDesc,sItemDesc

getReferenceDetail()

If bFlag = False Then
	msgbox "Select Reference Number",0,"Invoice"
	exit function
End if
Set objhttp = CreateObject("MSXML2.XMLHTTP")
objhttp.Open "GET","XMLGetInvItem.asp?Flag="&sFlag&"&sOrgID="&sOrgID&"&InvoiceAgainst="& sInvAgainst& "&RefNum=" & sInvRefNum ,false
objhttp.send


If objhttp.responseXML.xml <> "" then
	OutData.loadXML objhttp.responseXML.xml
	Set Root = OutData.documentElement
	if Root.HaschildNodes() then

	'' To fill in Party type select box based on party code selection
	For Each ndHeader In Root.childNodes
		if ndHeader.nodename = "PartyType" then
			document.formname.cmbPartyType.length = 0
			For each ndTemp in ndHeader.childnodes
				sPartyType = ndTemp.getAttribute("Type")
				spartySubType = ndTemp.getAttribute("SubTypeCode")
				sPartySubtypeName = ndTemp.getAttribute("SubTypeName")

				document.formname.cmbPartyType.length = document.formname.cmbPartyType.length+1

				document.formname.cmbPartyType.options(document.formname.cmbPartyType.length-1).text = sPartySubtypeName
				document.formname.cmbPartyType.options(document.formname.cmbPartyType.length-1).Value = cstr(spartySubType) + "|" + cstr(sPartyType)

			Next

		End if
	Next

	For Each ndHeader In Root.childNodes

		if ndHeader.nodename = "Header" then

			document.formname.hPartyCode.value = ndHeader.Attributes.Item(3).nodeValue
			sPartyName = ndHeader.Attributes.Item(4).nodeValue
			document.formname.cmbPartyType.value = ndHeader.Attributes.Item(5).nodeValue&"|"&ndHeader.Attributes.Item(6).nodeValue
			document.formname.txtPartyName.value = sPartyName

			if trim(ndHeader.Attributes.Item(7).nodeValue) <> "" and trim(ndHeader.Attributes.Item(7).nodeValue)<>"0" then
				document.formname.txtSuppInvNo.value  =  ndHeader.Attributes.Item(7).nodeValue
				document.formname.txtSuppInvDT.value  =  ndHeader.Attributes.Item(8).nodeValue

			else
				document.formname.txtSuppInvNo.readOnly = false
				document.formname.txtSuppInvNo.className = "formelem"
				document.formname.ctlDate.enable = true
			end if

		End if
	Next
	end if
End if
End Function
'-------------------------------------------------------------------------------------------
'=====================================================================================================
'-------------------------------------------------------------------------------------------
FUNCTION popTax1()

dim dInvAmount,sCatCode,sTaxCode,sTaxMode,sFormula,dTaxValue,dTax,TaxNode,i,j,PurType,ItemNd,TotInvAmount
dim nDisplayTotal,dRndTotInvAmount
set RootNode=Invoicedet.documentElement
TotInvAmount=0
Set TaxNode=RootNode.SelectNodes("//TaxDetails")
for i=0 to TaxNode.length-1
	PurType=TaxNode.item(i).Attributes.getnamedItem("PurchaseType").value
	set ItemNd=RootNode.SelectNodes("//ItemDetails/Item[@PurchaseType="& PurType &"]")

	dInvAmount=eval("document.formname.txtAmount"&PurType).value
	dBasicTotal=eval("document.formname.txtBasicValue"&PurType).value
	dTotal	=eval("document.formname.txtAmount"&PurType).value

	if trim(dInvAmount) = "" then dInvAmount = 0
	if trim(dBasicTotal) = "" then dBasicTotal = 0
	if trim(dTotal) = "" then dTotal = 0

	nDisplayTotal = 0
	nDispTax = 0
'''	nDisplayTotal = CDbl(dTotal)
	set TaxRoot=TaxNode.item(i)
	For Each oNodEntry in TaxRoot.childNodes

		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue
		sTaxMode=oNodEntry.Attributes.Item(2).nodeValue
		sFormula=oNodEntry.Attributes.Item(3).nodeValue
		dTaxValue=oNodEntry.Attributes.Item(4).nodeValue
		'alert(dTaxValue)

		If sCatCode <> "0" and sTaxCode <> "0" then		' Except round off node - 09 Aug 04

			if sTaxMode="P" then
				dTax=CalculateTax(sFormula,dBasicTotal,dTotal,dTaxValue,PurType)
			elseif sTaxMode = "Q" then
				dTax = cdbl(dTaxvalue) * cdbl(Qtysum)
			elseif sTaxMode = "K" then
				sTotpackvalue = 0
				sExp = "//Tax[@CatCode = "&sCatCode&" and @TaxCode ="&sTaxCode&"]/Taxpack"
				set sNode = TaxRoot.Selectnodes(sExp)
				If sNode.Length > 0 then
					For IMCtr = 0 to sNode.Length - 1
						sTotpackvalue = cdbl(sTotpackvalue) + cdbl(sNode.Item(iMCtr).Attributes.Item(3).nodevalue)
					Next
				end if
				dTax=sTotpackvalue
			else
				dTax=dTaxvalue
			end if

			dTRndoff = oNodEntry.getAttribute("Rndoff")

			if trim(dTRndoff) = "1" then
				dtax = RndOff(dTax)
			End if

			dTax=FormatNumber(dTax,2,,,0)
			oNodEntry.Attributes.Item(5).nodeValue=dTax

			dInvAmount=dInvAmount+CDbl(dTax)

			nDisplayTotal = nDisplayTotal + dTax

			if trim(oNodEntry.getAttribute("AccHead")) <> "0" then
				nDispTax = nDispTax + dTax
			end if

		end if 'If sCatCode <> "0" and sTaxCode <> "0" then
	Next

	dRoundedInvvalue = RndOff(dInvamount)
	dRoundedoff = Round(cdbl(dRoundedInvvalue) - cdbl(dInvamount),2)

	dRoundedoff = FormatNumber(dRoundedoff,2,,,0)
	dRoundedInvvalue=FormatNumber(dRoundedInvvalue,2,,,0)

'''	document.formname.txtRoundOff.value = dRoundedoff



	' set value for round off node - 09 Aug 04
	For Each oNodEntry in TaxRoot.childNodes
		sCatCode=oNodEntry.Attributes.Item(0).nodeValue
		sTaxCode=oNodEntry.Attributes.Item(1).nodeValue

		If sCatCode = "0" and sTaxCode = "0" then
			oNodEntry.setAttribute "TaxValue",document.formname.txtRoundOff.value
			oNodEntry.setAttribute "TaxAmount",document.formname.txtRoundOff.value
		End if
	Next

	TaxRoot.Attributes.Item(0).nodeValue=dBasicTotal
	TaxRoot.Attributes.Item(1).nodeValue=dTotal
	TaxRoot.setAttribute "TotalTax", FormatNumber(nDisplayTotal,2,,,0)
	TaxRoot.setAttribute "SubTotal", dInvamount

	eval("document.formname.txtTot" & PurType).value=FormatNumber(nDisplayTotal,2,,,0)
	eval("document.formname.txtTaxDisplay" & PurType).value=FormatNumber(nDispTax,2,,,0)

	eval("document.formname.txtSubTot" & PurType).value=FormatNumber(dInvAmount,2,,,0)
	TotInvAmount=TotInvAmount+dInvAmount
next

dRndTotInvAmount = RndOff(TotInvAmount)
dRoundedoff = Round(cdbl(dRndTotInvAmount) - cdbl(TotInvAmount),2)

dRoundedoff = FormatNumber(dRoundedoff,2,,,0)
dRndTotInvAmount=FormatNumber(dRndTotInvAmount,2,,,0)

'alert(dRoundedoff)
if cdbl(dRoundedoff) <> 0.00 then
	document.formname.rdRndOff(0).Checked=true
else
	document.formname.rdRndOff(1).Checked=true
end if


if document.formname.rdRndOff(0).Checked=true  then
	document.formname.txtRoundOff.value = dRoundedoff
	document.formname.txtInvValue.value=dRndTotInvAmount
else
	document.formname.txtRoundOff.value = "0.00"
	document.formname.txtInvValue.value= Formatnumber(TotInvAmount,2,,,0)
end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			set TempNode=oNodTemp
			exit for
		end if
	next

	for Each oNodTemp in TempNode.childNodes
		oNodTemp.setAttribute "InvValue",document.formname.txtInvValue.value
		oNodTemp.setAttribute "RoundOff",document.formname.txtRoundOff.value
	next

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
END FUNCTION
'---------------------------------------------------------------------------------
function Next_Click(sPara)
	dim objhttp1
	dim ItemNd,Qty,Rate,i,DisPer,Node,Nd,InvVal,nRetValue
	Dim ndNarration,sConsolidate

	nRetValue = 0


	If trim(document.formname.txtSuppInvNo.value)  = "" Then
		msgbox "Enter Supplier Invoice No.",0,"Invoice"
		document.formname.txtSuppInvNo.focus()
		Exit Function
'	elseif document.formname.cmbBillType.selectedIndex = 0 Then
'		MsgBox "Select Bill Type "
'		document.formname.cmbBillType.focus()
'		exit function
	elseIf datediff("d",document.formname.ctldate.getdate,document.formname.hCurrDate.value) < 0 then
		MsgBox "Invoice date cannot be greater than the current date"
		exit function
	elseif document.formname.cmbPurType.selectedIndex = 0 Then
		MsgBox "Select Purchase Type "
		document.formname.cmbPurType.focus()
		exit function
	elseif document.formname.cmbInvCat.selectedIndex = 0 Then
		MsgBox "Select Category "
		document.formname.cmbInvCat.focus()
		exit function
	elseif trim(document.formname.txtNarration.value)="" then
		MsgBox "Enter Narration"
		document.formname.txtNarration.focus
		exit function
	End If
	if document.formname.chkConsolidated.checked = true then
		sConsolidate="Y"
	else
		sConsolidate = "N"
	end if

	set RootNode=Invoicedet.documentElement

	set ItemNd=RootNode.SelectNodes("//ItemDetails/Item")

	''Blocked by Ragav
'	for i=0 to ItemNd.Length-1
'		Qty = cdbl(ItemNd.Item(i).Attributes.getNamedItem("Qty").value)
'		Rate = cdbl(ItemNd.Item(i).Attributes.getNamedItem("Rate").value)
'		DisPer = cdbl(ItemNd.Item(i).Attributes.getNamedItem("DisPer").value)
'		if Qty <= 0 then
'			msgbox "Invoice Quantity should be greater than 0"
'			nRetValue = 1
'		elseif Rate<=0 then
'			msgbox "Invoice Rate should be greater than 0"
'			nRetValue = 1
'		elseif DisPer < 0 then
'			msgbox "Discount Percentage should not be less than 0"
'			nRetValue = 1
'		end if
'		if nRetValue = 1 then exit for
'	next

''end

	if document.formname.txtAmount.value<=0 then
		alert("Total Value Greater than zero")
		exit function
	end if

	if nRetValue = 0 then
		For Each oNodTemp in RootNode.childNodes
			if oNodTemp.nodeName="InvoiceHeader" then
				set Node=oNodTemp
				exit for
			end if
		next
		for each Nd in Node.ChildNodes
			InvVal = Nd.Attributes.getNamedItem("InvValue").Value
		next
		if InvVal<=0 then
			msgbox "Invoice Value should be greater than 0"
			nRetValue = 1
		end if
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		set TempNode=RootNode.SelectNodes("//InvoiceHeader/Header")
		if TempNode.Length > 0 then
			TempNode.item(0).SetAttribute "InvValue",document.formname.txtInvValue.value
			TempNode.item(0).SetAttribute "RoundOff",document.formname.txtRoundOff.value
			TempNode.item(0).SetAttribute "Remarks",document.formname.mTextAreaRemarks.value

			TempNode.item(0).SetAttribute "SuppInvNo",document.formname.txtSuppInvNo.value
			TempNode.item(0).SetAttribute "SuppInvDt",document.formname.ctldate.getdate

			TempNode.item(0).SetAttribute "ItemType",document.formname.hItemType.value
			TempNode.item(0).SetAttribute "InvCategory",document.formname.cmbInvCat.value
			TempNode.item(0).SetAttribute "BillType","P" 'document.formname.cmbBillType.value
			TempNode.item(0).setAttribute "GPNo",document.formname.hGPNo.value
		end if
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	end if 'if nRetValue = 0 then

	set ndNarration = Invoicedet.createElement("Narration")
	ndNarration.text = document.formname.txtNarration.value
	RootNode.appendChild ndNarration

'alert(RootNode.xml)
	if sPara = "V" then
		Next_Click = true
		 exit function
	end if

		CalcAmountForLastItem

		AddRoundOffNode document.formname.txtRoundOff.value,document.formname.txtInvValue.value
		set objhttp1 = CreateObject("Microsoft.XMLHTTP")
		objhttp1.Open "POST","XMLSavePur.asp?Mod=PUR&Name=NewInvItemValue", false
		objhttp1.send InvoiceDet.XMLDocument


		if nRetValue = 0 then
			Next_Click = true
		else
			Next_Click = false
			exit function
		end if
		'exit function

		document.formname.action = "InvPurInvEntryService_AccDetails.asp?Consolidate="&sConsolidate
		document.formname.submit()
end function

'---------------------------------------------------------------------------------
Function CalcItemValueForEachItem(iItemCode,iClassCode)
	Dim ObjRoot,Node1,nSumOfItemRate,sItemPurType,nTotal
	Dim sPurchaseType


	Set ObjRoot = InvoiceDet.documentElement

	For Each oNodTemp in ObjRoot.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			set Node=oNodTemp
			exit for
		end if
	next
	for each Nd in Node.ChildNodes
		dInvAmount		= Nd.Attributes.getNamedItem("InvValue").Value
		sPurchaseType	= Nd.Attributes.getNamedItem("PurchaseType").Value
	next

	nSumOfItemRate = 0.00
	nTotal = 0.00
	dTemp=0
	sItemPurType = ""
	for each Node1 in ObjRoot.ChildNodes


		if Node1.nodeName = "ItemDetails" then

			dTemp=0

			if sItemPurType = "" then
				For each Node2 in Node1.ChildNodes
					If Node2.getAttribute("ItemCode") = iItemCode and Node2.getAttribute("ClassificationCode") = iClassCode then
						sItemPurType=Node2.getAttribute("PurchaseType")
					End if 'If Node2.getAttribute("ItemCode") = iItemCode and Node2.getAttribute("ClassificationCode") = iClassCode then
				Next 'For each Node2 in Node1.ChildNodes
			end if 'if sItemPurType = "" then
			iItemCount = 0
			nPos = 0
			if trim(sItemPurType) <> "" then
				For each Node2 in Node1.ChildNodes
					If trim(Node2.getAttribute("PurchaseType")) = trim(sItemPurType) then
						Node2.setAttribute "Amount", Node2.getAttribute("NettBasic")
						iItemCount = iItemCount + 1
						nPos = Node2.getAttribute("EntryNo")
					End if 'If trim(Node2.getAttribute("PurchaseType")) = trim(sItemPurType) then
				Next 'For each Node2 in Node1.ChildNodes
			end if 'if sItemPurType <> "" then


			For each Node2 in Node1.ChildNodes



				If sItemPurType = Node2.getAttribute("PurchaseType") then


					dItemValue=Node2.getAttribute("Amount")

					if cdbl(dItemValue) > 0 then
						'msgbox " first = "& trim(dItemValue)

						set TaxNodes=ObjRoot.SelectNodes("//TaxDetails[@PurchaseType="& sItemPurType &"]")

						if TaxNodes.Length > 0 then

							set oNodTax=TaxNodes.Item(0)

							dTotBasicVal	= oNodTax.getAttribute("Basicvalue")
							dTotNetVal		= oNodTax.getAttribute("NettValue")
							dInvAmount		= oNodTax.getAttribute("SubTotal")


								For Each oNodEntry in oNodTax.childNodes


									sCatCode	=	oNodEntry.getAttribute("CatCode")
									sTaxCode	=	oNodEntry.getAttribute("TaxCode")
									sTaxMode	=	oNodEntry.getAttribute("TaxMode")
									sFormula	=	oNodEntry.getAttribute("TaxFormula")
									dTaxValue	=	oNodEntry.getAttribute("TaxValue")
									if sTaxMode="F" then
										dTax = (cdbl(dTaxValue)/cdbl(dTotNetVal)) * cdbl(dItemValue)
									elseif sTaxMode="P" then
										'alert(sFormula & " -------------- " & trim(dItemValue)  & " ********** " & trim(dTaxValue) )
										dTax=ItemValueCalculateTax(oNodTax,sFormula,dItemValue,dItemValue,dTaxValue,sItemPurType)
									else
										dTax = (cdbl(dTaxValue)/cdbl(dTotNetVal)) * cdbl(dItemValue)
									end if

									saTemp= Split(CStr(dTax),".")
									if UBound(saTemp)>0 then
										oNodEntry.setAttribute "ItemValue", saTemp(0)&"."&left(saTemp(1),2)
									else
										oNodEntry.setAttribute "ItemValue", saTemp(0)
									end if

								Next


								For Each oNodEntry in oNodTax.childNodes
									if cint(oNodEntry.getAttribute("AccHead"))=0 then
										if trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then
											dItemValue=CDbl(dItemValue)+CDbl(oNodEntry.getAttribute("ItemValue"))
										end if
									else
										if trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then
											dTemp = CDbl(dTemp) + CDbl(oNodEntry.getAttribute("TaxAmount"))
										end if
									end if
								next

							dItemTotal=CDbl(dItemValue)+CDbl(dItemTotal)
							Node2.setAttribute "Amount", dItemValue
						end if 'if TaxNodes.Length > 0 then
					end if 'if cdbl(dItemValue) > 0 then
				end if 'If sItemPurType = Node2.getAttribute("PurchaseType") then

			Next'FOR EACH Node2 in Node1.ChildNodes

			'alert(nPos)
			if nPos > 0 then

				nPos = nPos  - 1
				if  iItemCount > 0 then

					dItemTotal = CDbl(dTemp/iItemCount) + CDbl(dItemTotal)

					dTemp = Node1.childNodes(nPos).getAttribute("Amount")
					dTemp = CDbl(dTemp) + CDbl(dInvAmount)-CDbl(dItemTotal)

					Node1.childNodes(nPos).SetAttribute "Amount",dTemp

				end if 'if  iItemCount > 0 then
			end if 'if nPos > 0 then

			For each Node2 in Node1.ChildNodes

				dProdRate		= Node2.getAttribute("Rate")
				iQty			= Node2.getAttribute("Qty")
				nTempRate		= Node2.getAttribute("Rate")
				dItemNetBasic	= Node2.getAttribute("NettBasic")
				dItemValue		= Node2.getAttribute("Amount")

				If sItemPurType = Node2.getAttribute("PurchaseType") then
					if cdbl(dItemValue) > 0 then
						''Item Rate
						dItemRate = cdbl(dItemValue) / cdbl(iQty)
						dItemRate = round(dItemRate,4)
						dItemValue = FormatNumber(dItemValue,2,,,0)

						Node2.setAttribute "ItemValue", dItemValue
						Node2.setAttribute "ItemRate", dItemRate


						If iQty > 0 and nTempRate > 0 then
							If dItemRate > 0 then
								eval("document.formname.txtItemRateDisplay"&Node2.getAttribute("ClassificationCode")&"Z"&Node2.getAttribute("ItemCode")&"Z"&Node2.getAttribute("SourceEntryNo")).value=dItemValue
								nSumOfItemRate = cdbl(nSumOfItemRate) + cdbl(dItemValue)
							End If 'If dItemRate > 0 then
						End If 'If iQty > 0 and nTempRate > 0 then


					end if 'if cdbl(dItemValue) > 0 then
				end if 'If sItemPurType = Node2.getAttribute("PurchaseType") then


				if dItemValue > 0 and iQty > 0 and dProdRate > 0 then
					nTotal = cdbl(nTotal) + cdbl(Node2.getAttribute("Amount"))
				end if

			Next 'For each Node2 in Node1.ChildNodes

		end if 'if Node1.nodeName = "ItemDetails" then

	next 'for each Node1 in ObjRoot.ChildNodes

	if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
		eval("document.formname.txtTotalItemRate" + sPurchaseType).value	= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
	else

		for each NodeTax in ObjRoot.ChildNodes
			if NodeTax.nodeName = "TaxDetails" then
				sPType= NodeTax.getAttribute("PurchaseType")
				nTotal = cdbl(nTotal) + cdbl( eval("document.formname.txtTaxDisplay" & sPType).value )
			end if
		next

		eval("document.formname.txtTotalItemRate" + sItemPurType).value		= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
		document.formname.mTxtTotal.value = FormatNumber(RndOff(nTotal),2,,,0)
	end if

End Function


'---------------------------------------------------------------------------------
Function CalculateItemValue(sPassPurType)
	Dim ObjRoot,Node1
	Dim nSumOfItemRate,nTotal
	Dim sPurchaseType,sItemPurType

	Set ObjRoot = InvoiceDet.documentElement
	'alert(ObjRoot.xml)

	nTotal = 0.00
	for each Node1 in ObjRoot.ChildNodes

		if Node1.nodeName = "InvoiceHeader" then

			FOR EACH Node2 in Node1.ChildNodes
				if Node2.NodeName = "Header" then
					sPurchaseType	= Node2.getAttribute("PurchaseType")
					dTotInvVal		= Node2.getAttribute("InvValue")
				end if 'if Node2.NodeName = "Header" then
			next
		end if 'if Node1.nodeName = "InvoiceHeader" then

		if document.formname.chkConsolidated.checked=true then
			if ObjRoot.hasChildNodes then
				Set ItemNode=ObjRoot.Selectnodes("//ItemDetails/Item")
				for i = 0 to ItemNode.Length - 1
						ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
						ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
						iEntryNo = ItemNode.Item(i).Attributes.getNamedItem("SourceEntryNo").value
						if document.formname.chkConsolidated.checked=true then
							ItemNode.Item(i).Attributes.getNamedItem("Rate").value =document.formname.txtAmount.value
							ItemNode.Item(i).Attributes.getNamedItem("RatePerQtyUoM").value=document.formname.txtAmount.value
							ItemNode.Item(i).Attributes.getNamedItem("Amount").value=document.formname.txtAmount.value
							ItemNode.Item(i).Attributes.getNamedItem("ItemValue").value=document.formname.txtAmount.value
							ItemNode.Item(i).Attributes.getNamedItem("ItemRate").value=document.formname.txtAmount.value
							ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value=document.formname.txtAmount.value
						else
							ItemNode.Item(i).Attributes.getNamedItem("Rate").value ="0"
							ItemNode.Item(i).Attributes.getNamedItem("RatePerQtyUoM").value="0"
							ItemNode.Item(i).Attributes.getNamedItem("Amount").value="0"
							ItemNode.Item(i).Attributes.getNamedItem("ItemValue").value="0"
							ItemNode.Item(i).Attributes.getNamedItem("ItemRate").value="0"
							ItemNode.Item(i).Attributes.getNamedItem("NettBasic").value="0"
						end if
				next
			popTax
			end if  'if ObjRoot.hasChildNodes then
		end if


		if Node1.nodeName = "ItemDetails" then


			dTemp=0
			iItemCount = 0
			nPos = 0

			if trim(sPassPurType) <> "" then
				For each Node2 in Node1.ChildNodes

					If trim(Node2.getAttribute("PurchaseType")) = trim(sPassPurType) then
						if document.formname.chkConsolidated.checked=true then
							Node2.setAttribute "Amount", document.formname.txtAmount.value
						else
							Node2.setAttribute "Amount", Node2.getAttribute("NettBasic")
						end if 'if document.formname.chkConsolidated.checked=true then
						iItemCount = iItemCount + 1
						nPos = Node2.getAttribute("EntryNo")
					End if 'If trim(Node2.getAttribute("PurchaseType")) = trim(sPassPurType) then

				Next 'For each Node2 in Node1.ChildNodes

			end if 'if trim(sPassPurType) <> "" then


			For each Node2 in Node1.ChildNodes

				sItemPurType=Node2.getAttribute("PurchaseType")

				if trim(sPassPurType) = trim(sItemPurType) then

					dItemValue=Node2.getAttribute("Amount")




					set TaxNodes=ObjRoot.SelectNodes("//TaxDetails[@PurchaseType="& sPassPurType &"]")

					if TaxNodes.Length > 0 then

						set oNodTax=TaxNodes.Item(0)

						dTotBasicVal	= oNodTax.getAttribute("Basicvalue")
						dTotNetVal		= oNodTax.getAttribute("NettValue")
						dInvAmount		= oNodTax.getAttribute("SubTotal")

						if cdbl(dTotBasicVal) > 0 and cdbl(dTotNetVal) > 0 then

							For Each oNodEntry in oNodTax.childNodes

								sCatCode	=	oNodEntry.getAttribute("CatCode")
								sTaxCode	=	oNodEntry.getAttribute("TaxCode")
								sTaxMode	=	oNodEntry.getAttribute("TaxMode")
								sFormula	=	oNodEntry.getAttribute("TaxFormula")
								dTaxValue	=	oNodEntry.getAttribute("TaxValue")

								if sTaxMode="F" then
									dTax = (cdbl(dTaxValue)/cdbl(dTotNetVal)) * dItemValue
								elseif sTaxMode="P" then
									dTax=ItemValueCalculateTax(oNodTax,sFormula,dItemValue,dItemValue,dTaxValue,sItemPurType)
								else
									dTax = (cdbl(dTaxValue)/cdbl(dTotNetVal)) * cdbl(dItemValue)
								end if

								saTemp= Split(CStr(dTax),".")

								if UBound(saTemp)>0 then
									oNodEntry.setAttribute "ItemValue", saTemp(0)&"."&left(saTemp(1),2)
								else
									oNodEntry.setAttribute "ItemValue", saTemp(0)
								end if
							Next

							For Each oNodEntry in oNodTax.childNodes
								if cint(oNodEntry.getAttribute("AccHead"))=0 then	'' considering Acc head = 0 only
									if trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then
										dItemValue=CDbl(dItemValue)+CDbl(oNodEntry.getAttribute("ItemValue"))
									end if
								else
									if trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then
										dTemp = CDbl(dTemp) + CDbl(oNodEntry.getAttribute("TaxAmount"))
									end if
								end if
							next
						end if 'if cdbl(dTotBasicVal) > 0 and cdbl(dTotNetVal) > 0 then
						dItemTotal=CDbl(dItemValue)+CDbl(dItemTotal)
						Node2.setAttribute "Amount", dItemValue

					end if 'if TaxNodes.Length > 0 then
				end if 'if trim(sPassPurType) = trim(sItemPurType) then

			Next'FOR EACH Node2 in Node1.ChildNodes

			if nPos > 0 then
				nPos = nPos - 1

				if  iItemCount > 0 then
					dItemTotal = CDbl(dTemp/iItemCount) + CDbl(dItemTotal)

					dTemp = Node1.childNodes(nPos).getAttribute("Amount")
					dTemp = CDbl(dTemp) + CDbl(dInvAmount)-CDbl(dItemTotal)
					Node1.childNodes(nPos).SetAttribute "Amount",dTemp
				end if 'if  iItemCount > 0 then
			end if 'if nPos > 0 then

			For each Node2 in Node1.ChildNodes

				dProdRate		= Node2.getAttribute("Rate")
				iQty			= Node2.getAttribute("Qty")
				dItemNetBasic	= Node2.getAttribute("NettBasic")
				dItemValue		= Node2.getAttribute("Amount")
				sItemPurType	= Node2.getAttribute("PurchaseType")


				if trim(sPassPurType) = trim(sItemPurType) then


					''Item Rate
					dItemRate = cdbl(dItemValue) / cdbl(iQty)
					dItemRate = round(dItemRate,4)
					dItemValue = FormatNumber(dItemValue,2,,,0)

					Node2.setAttribute "ItemValue", dItemValue
					Node2.setAttribute "ItemRate", dItemRate

					iItemCode = Node2.getAttribute("ItemCode")
					iClassCode = Node2.getAttribute("ClassificationCode")
					iEntryNo = Node2.getAttribute("SourceEntryNo")

					If iQty > 0 and dProdRate > 0 then
						If dItemRate > 0 then

							eval("document.formname.txtItemRateDisplay"&iClassCode&"Z"&iItemCode&"Z"&iEntryNo).value=dItemValue
							nSumOfItemRate = cdbl(nSumOfItemRate) + cdbl(dItemValue)

						End If 'If dItemRate > 0 then
					End If 'If iQty > 0 and dProdRate > 0 then

					if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
						eval("document.formname.txtTotalItemRate" + sPurchaseType).value	= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
					else
						eval("document.formname.txtTotalItemRate" + sItemPurType).value		= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
					end if
				end if 'if trim(sPassPurType) = trim(sItemPurType) then

				if dItemValue > 0 and iQty > 0 and dProdRate > 0 then
					nTotal = cdbl(nTotal) + cdbl(dItemValue)
				end if

			Next

		end if 'if Node1.nodeName = "ItemDetails" then

	next 'for each Node1 in ObjRoot.ChildNodes

	if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then

	else
		for each NodeTax in ObjRoot.ChildNodes
			if NodeTax.nodeName = "TaxDetails" then
				sPType= NodeTax.getAttribute("PurchaseType")
				nTotal = cdbl(nTotal) + cdbl( eval("document.formname.txtTaxDisplay" & sPType).value )
				'alert("for = " + trim(nTotal) )
			end if
		next

		nTotal = cdbl(nTotal) + cdbl(document.formname.txtRoundOff.value)
		document.formname.mTxtTotal.value = FormatNumber(nTotal,2,,,0)
	end if


End Function
'--------------------------------------------------------------------------------------------
Function CalcAmountForLastItem()

	Dim ObjRoot,Node1,sRoundOffAccHead
	Dim nSumOfItemRate,nTotal,nRoundOffValue
	Dim sPurchaseType,sItemPurType,sPrevItemPurchaseType
	Dim nLastSourceEntryNo


	sRoundOffAccHead = trim(document.formname.hRoundOffAccHead.value)

	Set ObjRoot = InvoiceDet.documentElement
	'alert(ObjRoot.xml)

	nTotal = 0.00
	for each Node1 in ObjRoot.ChildNodes

		if Node1.nodeName = "InvoiceHeader" then

			FOR EACH Node2 in Node1.ChildNodes
				if Node2.NodeName = "Header" then
					dTotInvVal		= Node2.getAttribute("InvValue")
					sPurchaseType	= Node2.getAttribute("PurchaseType")
					nRoundOffValue	= Node2.getAttribute("RoundOff")
				end if 'if Node2.NodeName = "Header" then
			next
		end if 'if Node1.nodeName = "InvoiceHeader" then


		if Node1.nodeName = "ItemDetails" then

			iItemCount = 0
			'alert(Node1.xml)
			if Node1.hasChildNodes then
				iItemCount = Node1.ChildNodes.length
				nLastItemCode  = Node1.childNodes(iItemCount - 1).getAttribute("ItemCode")
				nLastClassCode = Node1.childNodes(iItemCount - 1).getAttribute("ClassificationCode")
				nLastSourceEntryNo = Node1.childNodes(iItemCount - 1).getAttribute("SourceEntryNo")
			end if

			if sRoundOffAccHead	 = "0" then
				For each Node2 in Node1.ChildNodes


					if trim(Node2.getAttribute("ItemCode")) = trim(nLastItemCode) and trim(Node2.getAttribute("ClassificationCode")) = trim(nLastClassCode) and trim(Node2.getAttribute("SourceEntryNo"))=trim(nLastSourceEntryNo) then

						dItemValue=Node2.getAttribute("Amount")

						if document.formname.RdRndOff(0).Checked=true then
							dItemValue = CDbl(dItemValue) + CDbl(nRoundOffValue)
						else
							dItemValue = CDbl(dItemValue) - CDbl(nRoundOffValue)
						end if 'if document.formname.RdRndOff(0).Checked=true then

						Node2.setAttribute "Amount", dItemValue

					end if 'if trim(Node2.getAttribute("ItemCode")) = trim(nLastItemCode) and trim(Node2.getAttribute("ClassificationCode")) = trim(nLastClassCode) then

				Next'FOR EACH Node2 in Node1.ChildNodes
			end if 'if sRoundOffAccHead	 = "0" then

			sPrevItemPurchaseType = ""
			For each Node2 in Node1.ChildNodes


				dItemValue		= Node2.getAttribute("Amount")
				if eval("document.formname.chkVat"&Node2.getAttribute("ClassificationCode")&"Z"&Node2.getAttribute("ItemCode")&"Z"&Node2.getAttribute("SourceEntryNo")).checked then
					Node2.setAttribute "VAT", "Y"
				else
					Node2.setAttribute "VAT", "N"
				end if

				if trim(Node2.getAttribute("ItemCode")) = trim(nLastItemCode) and trim(Node2.getAttribute("ClassificationCode")) = trim(nLastClassCode) then
					iQty			= Node2.getAttribute("Qty")
					dItemNetBasic	= Node2.getAttribute("NettBasic")

					sItemPurType	= Node2.getAttribute("PurchaseType")

					if trim(sPrevItemPurchaseType) = "" or trim(sPrevItemPurchaseType) <> trim(sItemPurType) then
						nSumOfItemRate = 0.00
						sPrevItemPurchaseType = sItemPurType
					end if

					dItemRate = cdbl(dItemValue) / cdbl(iQty)
					dItemRate = round(dItemRate,4)
					dItemValue = FormatNumber(dItemValue,2,,,0)

					Node2.setAttribute "ItemValue", dItemValue
					Node2.setAttribute "ItemRate", dItemRate

					iItemCode = Node2.getAttribute("ItemCode")
					iClassCode = Node2.getAttribute("ClassificationCode")
					iEntryNo = Node2.getAttribute("SourceEntryNo")

					If iQty > 0 and dItemRate > 0 then
						If dItemRate > 0 then

							eval("document.formname.txtItemRateDisplay"&iClassCode&"Z"&iItemCode&"Z"&iEntryNo).value=dItemValue
							nSumOfItemRate = cdbl(nSumOfItemRate) + cdbl(dItemValue)

						End If 'If dItemRate > 0 then
					End If 'If iQty > 0 and dItemRate > 0 then

					if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
						eval("document.formname.txtTotalItemRate" + sPurchaseType).value	= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
					else
						eval("document.formname.txtTotalItemRate" + sItemPurType).value		= FormatNumber(Round(nSumOfItemRate,5),2,,,0)
					end if
				end if 'if trim(Node2.getAttribute("ItemCode")) = trim(nLastItemCode) and trim(Node2.getAttribute("ClassificationCode")) = trim(nLastClassCode) then
				nTotal = cdbl(nTotal) + cdbl(dItemValue)

			Next

		end if 'if Node1.nodeName = "ItemDetails" then

	next 'for each Node1 in ObjRoot.ChildNodes

	if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
	else
		for each NodeTax in ObjRoot.ChildNodes
			if NodeTax.nodeName = "TaxDetails" then
				sPType= NodeTax.getAttribute("PurchaseType")
				nTotal = cdbl(nTotal) + cdbl( eval("document.formname.txtTaxDisplay" & sPType).value )
			end if
		next

		document.formname.mTxtTotal.value	= FormatNumber(nTotal,2,,,0)
	end if

End Function
'--------------------------------------------------------------------------------------------
Function ItemValueCalculateTax(oNodTaxRoot,sFormula,dBValue,dDValue,dPercentage,sItemPurType)
dim saTemp,dTaxAmount,iCounter,iTemp,TaxRoot
dim oNodTemp
dim saTemp1

if trim(sItemPurType) ="0" then
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set TaxRoot=oNodTemp
		end if
	next
else
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			if trim(oNodTemp.getAttribute("PurchaseType")) = trim(sItemPurType) then
				set TaxRoot=oNodTemp
			end if
		end if
	next
end if 	'if trim(sItemPurType) ="0" then


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
	For Each oNodTemp in oNodTaxRoot.childNodes
		if oNodTemp.Attributes.Item(0).nodeValue=trim(saTemp1(0)) and oNodTemp.Attributes.Item(1).nodeValue=trim(saTemp1(1)) then
			dTaxAmount=CDbl(dTaxAmount)+CDbl(oNodTemp.getAttribute("ItemValue") )
		end if
	next
next
if trim(dPercentage)<>"" then
	ItemValueCalculateTax=dTaxAmount*(cdbl(dPercentage)/100)
else
	ItemValueCalculateTax=dTaxAmount
end if

End function
'-------------------------------------------------------------------------------------------
Function AssignRoundOffValue()

if document.formname.RdRndOff(0).Checked=true then
	if CDbl(document.formname.txtRoundOff.value) <> 0 then
		set RootNode=InvoiceDet.documentElement
		set TempNode=RootNode.SelectNodes("//InvoiceHeader/Header")
		if TempNode.Length > 0 then
			if CDbl(document.formname.txtRoundOff.value) <> cdbl(TempNode.item(0).Attributes.GetNamedItem("RoundOff").value) then

				InvVal=cdbl(TempNode.item(0).Attributes.GetNamedItem("InvValue").value)
				InvVal = CDbl(InvVal) - cdbl(TempNode.item(0).Attributes.GetNamedItem("RoundOff").value) + CDbl(document.formname.txtRoundOff.value)
				TempNode.item(0).SetAttribute "InvValue",InvVal
				TempNode.item(0).SetAttribute "RoundOff",document.formname.txtRoundOff.value
				RoundOffInv()
			end if
		end if 'if TempNode.Length > 0 then
	end if 	'if CDbl(document.formname.txtRoundOff.value) <> 0 then
end if 'if document.formname.RdRndOff(0).Checked=true then
end Function
'-------------------------------------------------------------------------------------------
Function RoundOffInv()
dim TempNode,InvVal,RndOffVal,blnRoundOffNodeExist,RndOffAmt

set RootNode=InvoiceDet.documentElement
set TempNode=RootNode.SelectNodes("//InvoiceHeader/Header")
if TempNode.Length > 0 then
	sPurchaseType	= TempNode.item(0).Attributes.GetNamedItem("PurchaseType").value
	InvVal=cdbl(TempNode.item(0).Attributes.GetNamedItem("InvValue").value)
	RndOffVal=cdbl(TempNode.item(0).Attributes.GetNamedItem("RoundOff").value)
	RndOffAmt = RndOffVal
end if

'alert(InvVal)

if document.formname.RdRndOff(0).Checked=true then
	'InvVal = CDbl(InvVal) +  CDbl(RndOffVal)
else
	InvVal = CDbl(InvVal) -  CDbl(RndOffVal)
	RndOffVal = 0.00
end if
document.formname.txtRoundOff.value = FormatNumber(RndOffVal,2,,,0)
document.formname.txtInvValue.value = FormatNumber(InvVal,2,,,0)

if Trim(sPurchaseType) <> "" and  Trim(sPurchaseType) <> "0" then
else
	document.formname.mTxtTotal.value = FormatNumber(InvVal,2,,,0)
end if

end function
'---------------------------------------------------------------------------------
Sub AddRoundOffNode(nPassRndOffAmt,nInvVal)

Dim Node1,TempNode,NewElem

Dim blnRoundOffNodeExist


	set RootNode=InvoiceDet.documentElement

	For each Node1 in RootNode.ChildNodes

		If trim(Node1.NodeName) = "InvoiceHeader" then
			for each TempNode in Node1.ChildNodes
				if trim(TempNode.NodeName) = "Header" then
					TempNode.setAttribute "InvValue",nInvVal
				end if
			next
		end if

		If trim(Node1.NodeName) = "TaxDetails" then
			'find round of tax node is exist or not , if not exist add it
			blnRoundOffNodeExist = false
			for each TempNode in Node1.ChildNodes
				set NewElem = TempNode
				if TempNode.getAttribute("CatCode") = "0" and TempNode.getAttribute("TaxCode") = "0" then
					blnRoundOffNodeExist = true
				end if
			Next

			if not blnRoundOffNodeExist then
				set NewElem = InvoiceDet.CreateElement("Tax")
				Node1.AppendChild NewElem
			end if 	 'end if 'if not blnRoundOffNodeExist then

			NewElem.setAttribute "CatCode","0"
			NewElem.setAttribute "TaxCode","0"
			NewElem.setAttribute "TaxMode","0"
			NewElem.setAttribute "TaxFormula","0"
			NewElem.setAttribute "TaxValue",nPassRndOffAmt
			NewElem.setAttribute "TaxAmount",nPassRndOffAmt
			NewElem.setAttribute "AccHead",document.formname.hRoundOffAccHead.value
			NewElem.setAttribute "Formnumber","0"
			NewElem.setAttribute "TransAmt","0"
			NewElem.setAttribute "ItemValue","0"
			NewElem.text = "ROUND OFF"

			'alert("nPassRndOffAmt = " + trim(nPassRndOffAmt) )
		End If 'If trim(Node1.NodeName) = "TaxDetails" then

	Next 'For each Node1 in RootNode.ChildNodes

end sub
'---------------------------------------------------------------------------------
Function ShowPrefDet()
	dim TempNode
	dim Curr,Mod1,Mop,IssueBank,PayTerm,Bop,Transporter
	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="InvoiceHeader" then
			set TempNode=oNodTemp
			exit for
		end if
	next
	for Each oNodTemp in TempNode.childNodes
		Curr=oNodTemp.Attributes.getnameditem("Currency").value
		Mod1=oNodTemp.Attributes.getnameditem("DespatchMode").value
		Mop=oNodTemp.Attributes.getnameditem("PaymentMode").value
		IssueBank=oNodTemp.Attributes.getnameditem("IssueBank").value
		PayTerm=oNodTemp.Attributes.getnameditem("PayTerms").value
		Bop=oNodTemp.Attributes.getnameditem("PricingBasis").value
		Transporter=oNodTemp.Attributes.getnameditem("Transporter").value
	next
	set TempNode = showModalDialog("InvPurInvoiceEntryPref.asp?Mod1="+ Mod1 +"&Mop="+ Mop +"&IssueBank="+IssueBank +"&PayTerm="+PayTerm +"&Bop="+Bop +"&Transporter="+Transporter,TempNode,"dialogLeft:0px;dialogTop:0Px;dialogHeight:250px;dialogWidth:600px;status:no")
end function

'---------------------------------------------------------------------------------

Function DeleteItems()
	dim root,sExp,ItemNode,itr,iItemDel,iClassDel,objSel,objSer,iEntryDel
	set root = InvoiceDet.DocumentElement
	sExp ="//ItemDetails/Item"
	Set ItemNode = Root.Selectnodes(sExp)
	if ItemNode.Length > 0 then
		for itr = 0 to ItemNode.Length - 1
			iItemDel = ItemNode.item(itr).Attributes.getNamedItem("ItemCode").value
			iClassDel = ItemNode.item(itr).Attributes.getNamedItem("ClassificationCode").value
			iEntryDel = ItemNode.item(itr).Attributes.getNamedItem("SourceEntryNo").value
			set objSel = eval("document.formname.chkDeleteA"&CStr(iClassDel)&"A"&CStr(iItemDel)&"A"&cstr(iEntryDel))
			set objSer = eval("document.formname.txtSerA"&CStr(iClassDel)&"A"&CStr(iItemDel)&"A"&cstr(iEntryDel))
			if objSel.checked then DeleteItem objSel,objSer.value + 2
		next
	end if
End Function

'-------------------------------------------------------------------------------

Function DeleteItem(sobj,iRow)
	dim itr,iItem,iClass,arrTemp,root,sExp,DeleteNode,oNode,iser,iItemDel,iClassDel,objSer,PurType,i,TaxNodes,Node,ItemRoot
	Dim iEntry,iEntryDel
	itr = 0
	arrTemp = split(sobj.name,"A")
	iItem = arrTemp(2)
	iClass = arrTemp(1)
	iEntry = arrTemp(3)
	set root = InvoiceDet.DocumentElement
	if root.hasChildNodes then
		For Each oNodTemp in Root.childNodes
		if oNodTemp.nodeName="ItemDetails" then
			set ItemRoot=oNodTemp
		end if
		next

		sExp = "//Item[@ItemCode = "&iItem&" and @ClassificationCode = "&iClass&" and @SourceEntryNo = "& iEntry &"]"
		Set DeleteNode = ItemRoot.Selectnodes(sExp)
		if DeleteNode.Length > 0 then

			PurType=DeleteNode.Item(0).Attributes.getNamedItem("PurchaseType").value
			Set oNode = ItemRoot.RemoveChild(DeleteNode.Item(0))
			if document.formname.cmbPurType.value=0 and PurType<>"0" then
				sExp = "//TaxDetails"
				Set TaxNodes = Root.Selectnodes(sExp)
				for i=0 to TaxNodes.Length-1
					if TaxNodes.Item(i).Attributes.getNamedItem("PurchaseType").value=Purtype then
						exit for
						msgbox "i="&i
					end if
				next
				iRow=iRow + 3*i
				document.all.tblItemDet.deleteRow(iRow - 1)
			else
				document.all.tblItemDet.deleteRow(iRow - 1)
			end if
		end if

		sExp = "//ItemDetails/Item[@PurchaseType = "& PurType &"]"
		set ItemNode= Root.Selectnodes(sExp)
		if ItemNode.Length=0 then
			sExp = "//TaxDetails[@PurchaseType = "& PurType &"]"
			Set DeleteNode = Root.Selectnodes(sExp)
			if DeleteNode.Length > 0 then
				if document.formname.cmbPurType.value=0 and PurType<>"0" then
					document.all.tblItemDet.deleteRow(iRow - 1)
					document.all.tblItemDet.deleteRow(iRow - 1)
					document.all.tblItemDet.deleteRow(iRow - 1)
				else
					for i=2 to (document.all.tblItemDet.rows.length - 1)
						document.all.tblItemDet.deleteRow(2)
					next
				end if
				Set oNode = root.RemoveChild(DeleteNode.Item(0))
			end if

		end if

		iser = 0
		set TaxNodes=Root.SelectNodes("//TaxDetails")
		for i= 0 to TaxNodes.Length-1
			PurType=TaxNodes.item(i).Attributes.GetNamedItem("PurchaseType").value
			sExp ="//ItemDetails/Item[@PurchaseType = "& PurType &"]"
			Set ItemNode = Root.Selectnodes(sExp)

			if ItemNode.Length > 0 then
				for itr = 0 to ItemNode.Length - 1
					iser = iser + 1
					iItemDel = ItemNode.item(itr).Attributes.getNamedItem("ItemCode").value
					iClassDel = ItemNode.item(itr).Attributes.getNamedItem("ClassificationCode").value
					iEntryDel = ItemNode.item(itr).Attributes.getNamedItem("SourceEntryNo").value
					set objSer = eval("document.formname.txtSerA"&CStr(iClassDel)&"A"&CStr(iItemDel)&"A"&cstr(iEntryDel))
					objSer.value = iser
				next
			end if
		next
		if TaxNodes.Length=0 then
			for i=2 to (document.all.tblItemDet.rows.length - 1)
				document.all.tblItemDet.deleteRow(2)
			next
		else
			if document.formname.cmbPurType.value=0 and PurType<>"0" then
				setValues1
			else
				SetValues
			end if
		end if
	end if
End Function

'=========================================================================
FUNCTION setValues1()
	dim dNetTotal,dDisTotal,dAmount,dRate,dQty,dDisAmount,i,PurType,j,ItemNd
	set RootNode=Invoicedet.documentElement
	Set TaxNode=RootNode.SelectNodes("//TaxDetails")
	for i=0 to TaxNode.length-1
		PurType=TaxNode.item(i).Attributes.getnamedItem("PurchaseType").value
		set ItemNd=RootNode.SelectNodes("//ItemDetails/Item[@PurchaseType="& PurType &"]")

		dNetTotal = 0
		dDisTotal = 0
		dTBasicVal = 0

		for j=0 to ItemNd.length-1
			set oNodTemp=ItemNd.item(j)
			dDisTotal=CDbl(dDisTotal)+CDbl(oNodTemp.Attributes.Item(7).nodeValue)
			dQty=oNodTemp.Attributes.Item(4).nodeValue

			dRate= oNodTemp.Attributes.getNamedItem("Rate").value

			dTBasicVal = dTBasicVal + (CDbl(dQty)*CDbl(dRate))
			dAmount=(CDbl(dQty)*CDbl(dRate))-CDbl(oNodTemp.Attributes.Item(7).nodeValue)
			dNetTotal=CDbl(dNetTotal)+dAmount
		next
		eval("document.formname.txtDisValue" &PurType).value=FormatNumber(Round(dDisTotal,2),2,,,0)
		eval("document.formname.txtAmount" &PurType).value=FormatNumber(Round(dNetTotal,2),2,,,0)
		eval("document.formname.txtBasicValue" &PurType).value=FormatNumber(Round(dTBasicVal,2),2,,,0)
	next
	popTax1
END FUNCTION
'-------------------------------------------------------------------------------------------
FUNCTION setValues()
dim dNetTotal,dDisTotal,dAmount,dRate,dQty,dDisAmount
	set RootNode=InvoiceDet.documentElement
	For Each oNodTemp in RootNode.childNodes
		if oNodTemp.nodeName="ItemDetails" then
			set ItemRoot=oNodTemp
		end if
	next

	dNetTotal = 0
	dDisTotal = 0
	dTBasicVal = 0

			For Each oNodTemp in ItemRoot.childNodes
					dDisTotal=CDbl(dDisTotal)+CDbl(oNodTemp.Attributes.Item(7).nodeValue)
					dQty=oNodTemp.Attributes.Item(4).nodeValue

					dRate= oNodTemp.Attributes.getNamedItem("Rate").value

					dTBasicVal = dTBasicVal + (CDbl(dQty)*CDbl(dRate))
					dAmount=(CDbl(dQty)*CDbl(dRate))-CDbl(oNodTemp.Attributes.Item(7).nodeValue)
					dNetTotal=CDbl(dNetTotal)+dAmount
			next
			document.formname.txtDisValue.value=FormatNumber(Round(dDisTotal,2),2,,,0)
			document.formname.txtAmount.value=FormatNumber(Round(dNetTotal,2),2,,,0)
			document.formname.txtBasicValue.value=FormatNumber(Round(dTBasicVal,2),2,,,0)
			popTax
END FUNCTION

'=========================================================================
function ViewItemValues()
	if Next_Click("V") then
		Set Root=InvoiceDet.documentElement
		showModalDialog "InvPurItemValueViewPop.asp",InvoiceDet,"dialogHeight:435px;dialogWidth:700px;Status:no;help:no"
	end if
end function
'--------------------------------------------------------------------------------------------
Function SelectAll()
	dim i, objCheck, Flag
	if document.formname.cmbPurType.selectedIndex = 0 Then
		MsgBox "Select Purchase Type "
		document.formname.cmbPurType.focus()
		exit function
	end if
	if document.formname.ChkAll.checked then
		Flag = True
	else
		Flag = False
	end if

	set root = InvoiceDet.DocumentElement
	sExp ="//ItemDetails/Item"
	Set ItemNode = Root.Selectnodes(sExp)
	if ItemNode.Length > 0 then
		for itr = 0 to ItemNode.Length - 1
			iItemDel = ItemNode.item(itr).Attributes.getNamedItem("ItemCode").value
			iClassDel = ItemNode.item(itr).Attributes.getNamedItem("ClassificationCode").value
			set objCheck = eval("document.formname.chkVat"&CStr(iClassDel)&"Z"&CStr(iItemDel))
			objCheck.Checked = Flag
		next
	end if
End Function

</script>

<%

'declaring variables
Dim iParTypeID,sParTypeName,sParType,sTraUnit,sRecptAg,sParSubType
Dim rsItem,dcrs,oDom,objDom2,Root,objFs,sAccRoot,sTempNode,SubNode,SubNode1,rsTemp
Dim newElem1
Dim iGPNo,iActivityNo,iPartyCode,iClassCode,iItemCode,iQtyRecd,iEntryNo
Dim sOrgID,sRoundOffHead,sSql,sPartyName,sClassDesc,sItemDesc,UoMCode,sItemType
Dim dOrderRate,dItemValue,sExp,sRateType

set rsTemp = Server.CreateObject("ADODB.RecordSet")
Set rsItem = Server.CreateObject("ADODB.Recordset")
Set dcrs = Server.CreateObject("ADODB.RecordSet")
set objDOM2 = server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iGPNo = trim(Request.QueryString("iGPNo"))
sOrgID = trim(Request.QueryString("ForUnit"))
sRateType = trim(Request.QueryString("RateType"))

iActivityNo = "10"		'Activity Number for "RECEIPT INVOICE"

dOrderRate=0
dItemValue=0

Set Root = oDOM.createElement("Root")
oDOM.appendChild Root


If objfs.FileExists(Server.MapPath("../xmldata/General.xml")) then
	objDOM2.Load server.MapPath("../xmldata/General.xml")
	Set sAccroot = objdom2.documentElement
	sExp = "//ROUNDOFF"
	Set sTempnode = sAccroot.Selectnodes(sExp)
	If sTempnode.Length > 0 then
		sRoundoffHead = sTempnode.Item(0).Attributes.Item(0).nodevalue
	else
		sRoundoffHead = "0"
	end if
else
	sRoundoffHead = "0"
End if



sSql = "Select GatePassNo,OrganisationCode,InvoiceType,ForInvoiceNo,PartyCode,TypeOfItems,MarkedOn,DCCode,"&_
	   "NoofPacks,PackingType,Transport,TakenBy,DeliveryBy,RefType from FORGATEPASSHEADER where GatePassNo = "& iGPNo  &""
dcrs.open sSql,con
if not dcrs.eof then
	iPartyCode = dcrs(4)
	sItemType = dcrs(5)
end if
dcrs.close



sSql = "Select PartyName from App_M_PartyMaster where PartyCode=" & trim(iPartyCode) & ""
'Response.Write sSql
With rsItem
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
End With
Set rsItem.ActiveConnection = nothing
if not rsItem.EOF then
	sPartyName = rsItem(0)
End if
rsItem.Close

sSql = "Select PartyType,PartySubType from VWORGParty where PartyCode = "& trim(iPartyCode)
'Response.Write sSql
rsItem.Open sSql,con
if not rsItem.EOF then
	sParType = trim(rsItem(0))
	sParSubType = trim(rsItem(1))
end if
rsItem.Close

	Set SubNode= oDOM.createElement("InvoiceHeader")
	Root.appendChild SubNode
	Set SubNode1=oDOM.createElement("ItemDetails")
	Root.appendChild SubNode1

	Set NewElem1 = oDOM.createElement("Header")
	newElem1.setAttribute "OrgID", sOrgID
	newElem1.setAttribute "Party", sPartyName
	newElem1.setAttribute "PurchaseType", ""
	newElem1.setAttribute "Currency",""
	newElem1.setAttribute "InvAgainst", "Receipt"
	newElem1.setAttribute "RefNum", ""
	newElem1.setAttribute "PartyCode", iPartyCode
	newElem1.setAttribute "PartyType",sParType
	newElem1.setAttribute "PartySubType",sParSubType
	newElem1.setAttribute "CurrencyNo", ""
	newElem1.setAttribute "DespatchMode",""
	newElem1.setAttribute "PaymentMode",""
	newElem1.setAttribute "PayTerms",""
	newElem1.setAttribute "IssueBank",""
	newElem1.setAttribute "BenificiaryBank",""
	newElem1.setAttribute "PricingBasis",""
	newElem1.setAttribute "Transporter",""
	newElem1.setAttribute "LoadingPort",""
	newElem1.setAttribute "DestPort",""
	newElem1.setAttribute "Remarks",""
	newElem1.setAttribute "SuppInvNo",""
	newElem1.setAttribute "SuppInvDt",""
	newElem1.setAttribute "TransporterFlag",""
	newElem1.setAttribute "PoNo",""
	newElem1.setAttribute "ConfNum",""
	newelem1.setAttribute "InvoiceFlag",""
	newelem1.setAttribute "InvValue",0
	newelem1.setAttribute "RoundOff",0
	newelem1.setAttribute "SuppCode",""
	newelem1.setAttribute "ItemType",""
	SubNode.appendChild NewElem1

	'==============================================================================
	sSql ="Select isNull(ClassificationCode,0),isNull(ItemCode,0),isNull(Quantity,0),"&_
	" isNull(Description,''),isNull(InvoicedUoM,'') from FORGATEPASSDETAILS where GatePassNo = "& iGPNo

	'Response.Write ssql
		With rsItem
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		End With

		Set rsItem.ActiveConnection = nothing

		Set iClassCode = rsItem(0)
		Set iItemCode= rsItem(1)
		Set iQtyRecd = rsItem(2)
		Set sItemDesc = rsItem(3)
		Set UomCode=rsItem(4)
	'	set sStockType = rsItem(7)
	'	set nActItemRate = rsItem(8)

		iEntryNo = 0
		If not rsItem.EOF then
		Do While Not rsItem.EOF

			'Response.Write "<p> sRefNum = "& sRefNum
			'' to add Qty Validation for Receipt

				sSql = "Select ItemDescription from vwItem where ItemCode ="& iItemCode &"  and ClassificationCode = "& iClassCode &""
				rsTemp.Open sSql,con
				if not rsTemp.EOF then
					sItemDesc = trim(rsTemp(0))
				end if
				rsTemp.Close


					iEntryNo = iEntryNo + 1
					Set newElem1 = oDOM.createElement("Item")

					newElem1.setAttribute "ItemCode", iItemCode
					newElem1.setAttribute "ClassificationCode", iClassCode
					newElem1.setAttribute "ItmDescription", sItemDesc
					newElem1.setAttribute "Uom", UomCode
					newElem1.setAttribute "Qty", iQtyRecd
					newElem1.setAttribute "Rate","0"
					newElem1.setAttribute "DisPer", "0"
					newElem1.setAttribute "DisAmount", "0"
					newElem1.setAttribute "NettBasic", "0"
					newElem1.setAttribute "UomDesc", UomCode
					newElem1.setAttribute "EntryNo", iEntryNo
					newElem1.setAttribute "RatePerQtyUoM", "0"
					newElem1.setAttribute "SourceEntryNo", iEntryNo
					newElem1.setAttribute "PurchaseType", ""
					newElem1.setAttribute "Amount", "0"
					newElem1.setAttribute "ItemValue", "0"
					newElem1.setAttribute "ItemRate", "0"
					newElem1.setAttribute "RateUOM", ""
					newElem1.setAttribute "StockType",""
					newElem1.setAttribute "VAT", ""
					SubNode1.appendChild NewElem1
			rsItem.MoveNext
		Loop
		End if
		rsItem.Close
		'Root.appendChild SubNode1
		'*******************************************************************************************************
		oDOM.save server.MapPath("..\temp\transaction\InvItemValue_PUR_"&Session.SessionID&".xml")

%>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="addItem()">

<form method="POST" name="formname" action>
	<input type=hidden name="hItemtype" value="<%=sItemType%>">
	<input type="hidden" name="hOrgID" value="<%=sOrgID%>">
	<input type="hidden" name="hPartyCode" value="<%=iPartyCode%>">
	<input type="hidden" name="hCurrDate" value="<%=FormatDate(Date)%>">
	<input type="hidden" name="hPartyType" value="<%%>">
	<input type="hidden" name="hPSubType" value="<%%>">
	<input type="hidden" name="txtSuppInvDt" value="">
	<Input type="hidden" name="hConfNum" value="<%%>" >
	<input type="hidden" name="hFlag" value="<%%>">
	<input type="hidden" name="hRcptno" value="<%%>">
	<input type="hidden" name="hRcptCode" value="<%%>">
	<input type="hidden" name="txtPartyName" value="">
	<input type="hidden" name="hRoundOffAccHead" value="<%=sRoundoffHead%>">

	<input type="hidden" name="hSuppInvDate" value="<%%>">
	<input type="hidden" name="hGPNo" value="<%=iGPNo%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Purchase Invoice Entry - Service Return
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td>
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0">
											<tr>
												<td class="FieldCell" valign="top">Supplier Name
												</td>
												<td class="FieldCellSub" colspan="4">
													<span class="Dataonly" id="spnSuppName"><%=sPartyName%>&nbsp;</span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Supplier Type
												</td>
												<td class="FieldCellSub" colspan="4">
													<select size="1" class="Formelem" name="cmbPartyType">
															<option value="0" selected>Select</option>
															<%populatePartyTypeForParty sParType,sParSubType %>
													</select>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Supplier Inv. No. / Dt.
												</td>
												<td class="FieldCellSub" colspan="4">
													<table>
													<tr><td>
													<input type="text" name="txtSuppInvNo" size="30" class="FormElem" align=top value="<%%>">
													</td><td>
 													<% ' Function Call to Insert Date Picker
													Response.Write InsertDatePicker("ctlDate")
													%>
													</td>
													<td>
													<!--	<select size="1" class="Formelem" name="cmbBillType">
															<option value="0" selected>Select Bill type</option>
															<option value="P">Credit bill</option>
															<option value="C">Cash bill</option>
														</select>-->
													</td>
													</tr>
													</table>
												</td>
											</tr>

											<tr>
												<td class="FieldCell" valign="top">Purchase Type
												</td>
												<td class="FieldCellSub" colspan="4">
												<select size="1" class="FormElem" name="cmbPurType" onchange="getTaxDet()">
														<option  value="" selected>Select</option>
														<option  value="0">--------------ITEMWISE-------------</option>
														<%

															popSelPurTypeFull("0")
														%>
													</select>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Currency
												</td>
												<td class="FieldCellSub"><select size="1" name="D29" class="Formelem">
														<%
														populateCurrency
								  						'popSelCurrency(Curr1)
								  						%>
													</select>
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">Preferences&nbsp&nbsp
												<a>
													<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" width="11" height="11" onclick="ShowPrefDet()" alt="Preferences" style="cursor:hand">
													</a>
												</td>
												<td class="FieldCellSub" align="left">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Category
												</td>
												<td class="FieldCellSub" colspan="3">
												<select size="1" name="cmbInvCat" class="FormElem" >
													<option value="0">Select</option>
													<%

													sSql = "Select CategoryCode,CategoryName From APP_M_InvoiceCategory Where ApplicableFor = 'P' Order By CategoryName "
													With dcrs
															.CursorLocation = 3
															.CursorType = 3
															.Source = sSql
															.ActiveConnection = Con
															.Open
														End With
														Set dcrs.ActiveConnection = Nothing
														Do While Not dcrs.EOF
														%>
														<Option value="<%=dcrs(0)%>"><%=dcrs(1)%></Option>
														<%
														dcrs.MoveNext
														loop
														dcrs.Close
													%>


													</select>
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Narration
												</td>
												<td class="FieldCellSub" colspan=4>
													<textarea id="txtNarration" class="FormElem" rows=2 cols=50></textarea>
												</td>
											</tr>
										</table>
									</td>
									<td>
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td>
									</td>
									<td valign="top" width="100%">
										<div class="frmBody" id="frm2" style="width: 585; height:242;">
											<table border="0" cellspacing="1" class="ExcelTable" width="1000" id="tblItemDet">
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">
														<a href="#"><img name="ImgDeleteIcon" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onClick="DeleteItems()" alt="Delete Selected Item"></a>
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description
														<a><img name=imgPurchaseDet border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" disabled=true alt="Enter Purchase Details" onclick="ShowPurchaseDet('<%=sOrgID%>','<%=sItemType%>')" width="15" height="15" style="cursor:hand"></a>
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Invoice<br>
														UoM
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Invoice<br>
														Quantity
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Invoice<br>
														Rate
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Basic<br>
														Value
													</td>
													<td class="ExcelHeaderCell" align="center" colspan="2">Discount
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Nett<br>
														Basic
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">Item Value
													</td>
													<td class="ExcelHeaderCell" align="center" rowspan="2">VAT
													<input type="Checkbox" name="ChkAll" class="Formelem" onClick="SelectAll()">
													</td>
													<!--td class="ExcelHeaderCell" align="center" rowspan="2">Purchase Type <a href="InvoiceTaxDetails.html" target="_flank">
														<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" width="11" height="11">
														</a>
													</td-->
												</tr>

												<tr>
													<td class="ExcelHeaderCell" align="center">%
													</td>
													<td class="ExcelHeaderCell" align="center">Value
													</td>
												</tr>

											</table>
										</div>
									</td>
									<td>
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td>
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td width="65" class="FieldCell" valign="top">Remarks
												</td>
												<td><textarea rows="2" name="mTextAreaRemarks" cols="93" class="FormElem"></textarea>
												</td>
											</tr>

										</table>
									</td>
									<td>
									</td>
								</tr>

								<tr>
									<td colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td>
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
											        <!--<input type="Button" value="View Item Values"	name="BtnViewItem" onClick="ViewItemValues()" class="ActionButton" tabindex="3" style="width:120">-->
													<input type="Button" value="Next"				name="BtnNext" onclick="javascript: return Next_Click('S')" class="ActionButton" tabindex="3">
 													<input type="reset"  value="Cancel"				name="BtnCancel" class="ActionButton" tabindex="4">
 													<input type="reset"  value="Reset"				name="BtnReset" class="ActionButton" tabindex="4">
												</td>
											</tr>

										</table>
									</td>
									<td>
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td colspan="3" class="BottomPack">
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
</body>
</html>
