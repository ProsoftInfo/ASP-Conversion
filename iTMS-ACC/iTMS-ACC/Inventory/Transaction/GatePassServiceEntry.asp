<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassServiceEntry.asp
	'Module Name				:	Gate Pass - Service
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	APRIL 05, 2010
	'Modified On				:	Jan 06,2011
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/CheckPrevFinYear.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
	Dim sOrgID,sOrgName,rsObj
	sOrgID = Session("organizationcode")
	set rsObj = Server.CreateObject("ADODB.Recordset")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Gate Pass</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<script type="application/xml" id="UoMData" data-itms-xml-island="1" data-src="../../inventory/xmldata/Uom.xml"></script>
<script type="application/xml" id="OutData" data-itms-xml-island="1"><Root/></script>
<script type="application/xml" id="Data" data-itms-xml-island="1"><ROOT></ROOT></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/gatePassServiceEntry.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,iCounter1,ssupcode,iser,iTempCode,Partyname,iRcptNo,sPartyCode
j = 0
iser = 0
iCounter1 = 0

Function popSuppAgent()

	sForUnit = document.formname.hUnitID.value+":O"
	nFlag = 1

	'note : value for hSelectMode is M( Multiple ) / S(Single)
	set OutValue=showModalDialog("SupplierSelect.asp?Unit="+sForUnit+"&hSelectMode=S&Flag="+cstr(nFlag),OutData,"status:no")
	'msgbox OutValue.xml

	sAct = UCase(trim(OutValue.getAttribute("Action")))
	sQuery = trim(OutValue.getAttribute("PassQuery"))
	if ucase(trim(sAct)) <> "CLOSE" then
		do while sAct <> "DONE"
			set OutValue=showModalDialog("SupplierSelect.asp?" & sQuery,OutData,"status:no")
			sAct = UCase(trim(OutValue.getAttribute("Action")))
			sQuery = trim(OutValue.getAttribute("PassQuery"))

			if ucase(Trim(sAct)) = "CLOSE" then exit do
		loop
	end if 'if ucase(trim(sAct)) <> "CLOSE" then
	'alert(OutValue.xml)

	If not OutValue.hasChildNodes Then 	exit function


	For each Node2 in OutValue.childNodes
		if ucase(Node2.nodename) = ucase("Supplier") then
			sPartyCode = trim(Node2.getAttribute("SuppCode"))
			Partyname = trim(Node2.getAttribute("SuppShortCode"))
			document.formname.txtRefName.value = Partyname
		end if
	Next
End Function

Function ClearTable()
	dim i
	for i=1 to document.all.tblLot.rows.length - 1
		document.all.tblLot.deleteRow(1)
	next
	j = 0
end Function

Function AddDetails()
		Dim sDesc

		'alert(document.formname.selUOM.disabled)

		set Root = OutData.documentElement
		if trim(document.formname.txtDesc.value) = "" and trim(document.formname.txtItemDesc.value) = "" then
			alert("Enter Description or Select Item...!")
			document.formname.txtDesc.select
			exit function
		elseif trim(document.formname.txtQty.value) = "" then
			alert("Enter Quantity")
			document.formname.txtQty.select
			exit function
		elseif not document.formname.selUOM.disabled and ( document.formname.selUOM.value = "select" or trim(document.formname.selUOM.value) = "" ) then
			alert("Select UOM")
			document.formname.selUOM.focus
			exit function
		else
			If trim(document.formname.txtItemDesc.value) <> "" Then
				sDesc = ucase(Trim(document.formname.txtItemDesc.value))
			End If

			If trim(document.formname.txtDesc.value) <> "" Then
				If Trim(sDesc) = "" Then
					sDesc = ucase(trim(document.formname.txtDesc.value))
				Else
					sDesc = sDesc & " - " & ucase(trim(document.formname.txtDesc.value))
				End If
			End If


			sApplFormJJ = "N"
			if document.formname.ChkFormJJ.checked then
				sApplFormJJ = "Y"
			end if

			sExp = "//ROOT/DETAILS [@DESC = '" & replace(sDesc,"'"," ") & "']"

			set Node1 = Root.selectNodes(sExp)
			'alert Node1.length
			if Node1.length > 0 then
				alert("Description Already entered")
				document.formname.txtDesc.select
				exit function
			end if


			Set Node = OutData.createElement("DETAILS")
			Node.setAttribute "DESC",sDesc
			Node.setAttribute "QTY",document.formname.txtQty.value
			Node.setAttribute "ITEMCODE", document.formname.hItem.value
			Node.setAttribute "CLASSCODE", document.formname.hClass.value
			Node.setAttribute "ITEMDESC", ucase(Trim(document.formname.txtItemDesc.value))
			Node.setAttribute "OTHERDESC", ucase(trim(document.formname.txtDesc.value))
			Node.setAttribute "UOM", ucase(trim(document.formname.SelUOM.value))
			Node.setAttribute "VALUE",document.formname.txtValue.value
			Node.setAttribute "FORMJJ",sApplFormJJ
			Node.setAttribute "REASON",TRIM(document.formname.txtReason.value)
			Root.appendChild Node

			set oRow = document.all.tblDetails.insertRow(j+1)

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=j+1
			headerCell.className="ExcelSerial"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=sDesc
			headerCell.className="ExcelDisplayCell"
			headerCell.align="left"

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=document.formname.txtQty.value
			headerCell.className="ExcelDisplayCell"
			headerCell.align="right"


			set headerCell=oRow.insertCell()
			headerCell.innerHTML=document.formname.SelUOM.value
			headerCell.className="ExcelDisplayCell"
			headerCell.align="right"

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=document.formname.txtValue.value
			headerCell.className="ExcelDisplayCell"
			headerCell.align="right"

			set headerCell=oRow.insertCell()
			if sApplFormJJ = "Y" then
				set oText = document.createElement("<input type=""checkbox"" value=""Y"" name=""chkbox"& j+1 &""" class=""FormElem"" checked >" )
			else
				set oText = document.createElement("<input type=""checkbox"" value=""Y"" name=""chkbox"& j+1 &""" class=""FormElem"" >" )
			end if
			headerCell.appendChild(oText)

			headerCell.className="ExcelDisplayCell"
			headerCell.align="right"

			set headercell = orow.insertcell()
			headercell.className="ExcelDisplayCell"
			headercell.align="left"
			headercell.innerText = TRIM(document.formname.txtReason.value)

			j = j + 1

			document.formname.txtDesc.value = ""
			document.formname.txtQty.value = ""

			document.formname.txtItemDesc.value = ""
			document.formname.hItem.Value =  ""
			document.formname.hClass.value =  ""
			document.formname.selUOM.value  = "select"

			document.formname.txtValue.value = ""
			document.formname.ChkFormJJ.checked = false

		end if

End Function

Function CheckSubmit(todaysdate)
'	if document.formname.selUnit.selectedIndex = "0" then
'		alert("Select For Unit")
'		document.formname.selUnit.focus
'		exit function
'	else
   ' if document.formname.selItmType.selectedIndex = "0" then
'		alert("Select Item Type")
'		document.formname.selItmType.focus
'		exit function
'	else
    if document.formname.txtRefName.value = "" then
		alert("Select Supplier")
		exit function
	elseif len(trim(document.formname.txtRemarks.value)) > 200 then
		alert("Remarks should be less than 200 characters")
		document.formname.txtRemarks.select
		exit function
	else
		itr = 0
		set root = OutData.DocumentElement
		sExp ="//DETAILS"
		Set ItemNode = Root.Selectnodes(sExp)
		if ItemNode.Length = 0 then
			alert("Enter Details")
			exit function
		else

			for nTempCtr = 0 to ItemNode.Length -1
				set ObjChk1 = eval("document.formname.chkbox" + trim(nTempCtr+1) )
				if ObjChk1.checked then
					ItemNode.item(nTempCtr).Attributes(8).value ="Y"
				else
					ItemNode.item(nTempCtr).Attributes(8).value ="N"
				end if
			next
		end if

		Set newElem = OutData.createElement("HEADER")
		newElem.setAttribute "FORUNIT", document.formname.hUnitID.value
		newElem.setAttribute "ITEMTYPE",""' document.formname.selItmType.value
		newElem.setAttribute "SUPPAGENT", sPartyCode
		newElem.setAttribute "REMARKS", trim(document.formname.txtRemarks.value)

		newElem.setAttribute "Transport", trim(document.formname.txtTransport.value)
		newElem.setAttribute "TakenBy", trim(document.formname.txtTakenBy.value)
		newElem.setAttribute "DeliveryBy", trim(document.formname.txtDeliveryBy.value)
		Root.appendChild newElem
	end if

	'alert(OutData.xml)
	'exit function

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","GatePassServiceInsert.asp", false
	objhttp.send OutData.XMLDocument
	'alert(objhttp.responseText)
	'exit function

	if objhttp.responseText = "" then
		if confirm("Gate Pass for Service has been created. Do you want to create another one?") then
			window.location.href "GatePassServiceEntry.asp"
		else
			window.location.href "GATEPASSSELECTION.ASP"
		end if
	else
		alert(objhttp.responseText)
	end if

end Function

Function SelectItem()
Dim	sIType, iItem,iClass,objhttp

sUnit = document.formname.hUnitID.value

        sTempValWindowSize = GetWindowSizeForPopup("1")
		sArrTempValWindowSize = split(sTempValWindowSize,":")
		sProgramName = sArrTempValWindowSize(0)
		sPopupHeight = sArrTempValWindowSize(1)
		sPopupWidth = sArrTempValWindowSize(2)

		set OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=S&Flag="+cstr(nFlag)&"&hDispButt="&bAddButton&"&hDispItem="&sDispItem&"&CallFrom="&sCallFrom,Data,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))

	    Set Root = Data.documentElement
	    'alert(Root.xml)
	    If not Root.hasChildNodes Then 	exit function
	    if Root.hasChildNodes() then
		    For Each HeaderNode In Root.childNodes
			    'alert(HeaderNode.xml)
			    document.formname.txtItemDesc.value = Trim(HeaderNode.getAttribute("ItemName"))
			    document.formname.hItem.Value =  HeaderNode.getAttribute("ItemCode")
			    document.formname.hClass.value =  HeaderNode.getAttribute("ClassCode")
			    document.formname.SelUOM.value = Trim(HeaderNode.getAttribute("StoresUoM"))
			    document.formname.selUOM.disabled = true
		    next
	    end if


end Function


</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
	<form method="POST" name="formname">
	<input type=hidden name="hItem" value="">
	<input type=hidden name="hClass" value="">
	<input type=hidden name="hUnitID" value="<%=sOrgID%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Gate Pass - Service
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>
								<tr>
									<td align="center">
									</td>
									<td width="100%" colspan="2">
										<div align="left">
											<table border="0" cellspacing="0" cellpadding="0" width="572">
											<!--	<tr>
													<td class="FieldCellSub">For Unit</td>
													<td class="FieldCellSub">
														<select size="1" name="selUnit" class="FormElem">
															<option value="select">Select</option>
															<%	'Calling the Function which populates Organization Unit list
																populateUnit
															%>
														</select>
													</td>
												</tr>-->

											<!--	<tr>
													<td class="FieldCellSub">Item Type</td>
													<td class="FieldCellSub" valign="top">
                                                        <select size="1" name="selItmType" class="FormElem">
															<option value="select">Select</option>
															<%	'Calling the Function which populates the Item Type list
														'		populateItemType
															%>
														</select>
													</td>
												</tr>-->
												<tr>
													<td class="FieldCellSub">Party</td>
													<td class="FieldCellSub">
														<input type="text" name="txtRefName" value size="60" class="FormElemRead" readonly>&nbsp;
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Party" width="11" height="11" onClick="popSuppAgent()"></a>
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Item Description</td>
													<td class="FieldCellSub">
														<input type="text" name="txtItemDesc" value size="60" class="FormElemRead" readonly>&nbsp;
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item" width="11" height="11" onClick="SelectItem()"></a>
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Other Description</td>
													<td class="FieldCellSub">
														<input type="text" name="txtDesc" maxlength=50 size="60" class="FormElem">
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Reason</td>
													<td class="FieldCellSub">
														<input type="text" name="txtReason" maxlength=35 size="60" class="FormElem" value="SENT FOR REPAIRS - TO BE RETURNED">
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Quantity & Value</td>
													<td class="FieldCellSub">
														<input type="text" name="txtQty" value size="12" class="FormElem" onkeypress="DoKeyPress('Y',7,3)">&nbsp;
														<select size="1" name="selUOM" class="FormElem">
															<option value="select">Select</option>
															<%
															populateUoM()
															%>
														</select>
														&nbsp;
														<input type="text" name="txtValue" size="12" class="FormElem" onkeypress="DoKeyPress('Y',7,2)" >&nbsp;
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Form JJ Applicable</td>
													<td class="FieldCellSub">
														<input type="checkbox" name="ChkFormJJ" value="Y" class="FormElem" >
														<input type="button" value="Add" name="B3" class="AddButtonX" onclick = "AddDetails()">
													</td>
												</tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack"></td>
								</tr>

								<tr>
									<td align="center"></td>
									<td width="100%" colspan="2">
										<div class="frmBody" id="frm1" style="width: 585; height:230;">
											<table border="0" cellspacing="1" class="ExcelTable" width="585" id=tblDetails>
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10">
														<p align="center">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center" width=500>
														Item Description
													</td>
													<td class="ExcelHeaderCell" align="center">Quantity</td>
													<td class="ExcelHeaderCell" align="center">UoM</td>
													<td class="ExcelHeaderCell" align="center">Value</td>
													<td class="ExcelHeaderCell" align="center">Form JJ Applicable</td>
													<td class="ExcelHeaderCell" align="center">Reason</td>
												</tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td colspan="4">
										<table border="0" cellspacing="0" cellpadding="0" width="572">
											<tr>
												<td class="FieldCell" Rowspan="3">Remarks</td>
												<td class="FieldCellSub" Rowspan="3" >
													<textarea rows="3" name="txtRemarks" cols="50" class="FormElem"></textarea>
												</td>
												<td class="FieldCell">Transport</td>
												<td class="FieldCellSub">
													<input type="text" name="txtTransport" maxlength=50 size="50" class="FormElem">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Taken By</td>
												<td class="FieldCellSub">
													<input type="text" name="txtTakenBy" maxlength=50 size="50" class="FormElem">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Delivery By</td>
												<td class="FieldCellSub">
													<input type="text" name="txtDeliveryBy" maxlength=50 size="50" class="FormElem">
												</td>
											</tr>
										</table>
									</td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" colspan="2">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<input type="button" value="Save" name="BtnSubmit" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date)%>')">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
 													&nbsp;
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="BottomPack">
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

