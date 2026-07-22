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
	'Created On					:	APRIL 05,2010
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!-- #include File="../../include/populate.asp" -->
<%
	Dim oDOM,Root,nodeDetail,objRs,objRs1,hRoot,suppNode
	Dim sQuery,sRemarks,sTransport,sDelivery,sTakenBy,sForUnit,sItemType,sSupAgent,sDesc
	Dim nGatePassNo
	Dim sItemCode,sClassCode,sItemDesc,sOtherDesc,sReason
	nGatePassNo= Request.QueryString("GatePassNo")
	sForUnit = Session("organizationcode")

	set oDOM = CreateObject("Microsoft.XMLDOM")
	set objRs = server.CreateObject("ADODB.Recordset")
	set objRs1 = server.CreateObject("ADODB.Recordset")

	sQuery = "Select GATEPASSNO,ORGANISATIONCODE,INVOICETYPE,PARTYCODE,isNull(TYPEOFITEMS,''),APPLICATIONCODE,"&_
			 "MARKEDON,isNull(REMARKS,''),STATUS,isNull(Transport,''),isNull(TakenBy,''),isNull(DeliveryBy,''),DCCODE from  FORGATEPASSHEADER "&_
			 "WHERE GatePassNo ="& nGatePassNo

	'Response.Write sQuery
	set Root = oDOM.createElement("Root")
			oDOM.appendChild Root
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sQuery
		.Open
	end with
	if not objrs.EOF then
		set hRoot = oDOM.createElement("HEADER")
			hRoot.setAttribute "ITEMTYPE",objRs(4)
			hRoot.setAttribute "FORUNIT",objrs(1)
			hRoot.setAttribute "REMARKS",objrs(7)
			hRoot.setAttribute "SUPPAGENT",objrs(3)
			hRoot.setAttribute "Transport",objRs(9)
			hRoot.setAttribute "TakenBy",objrs(10)
			hRoot.setAttribute "DeliveryBy",objrs(11)
			Root.appendChild hRoot
			sSupAgent = objrs(3)
	end if
	objrs.Close
	'sQuery = "Select OrgnPartyCode,SupplierCode,PartyName,PartyCode,PartyType,PartySubType from vwSupplierAddress where PartyCode = "&sSupAgent
	sQuery = "Select OrgnPartyCode,B.PartyCode,PartyName,A.PartyCode,PartyType,PartySubType from APP_M_PartyMaster A,APP_R_OrgParty B where A.PartyCode = B.PartyCode and A.PartyCode = "& sSupAgent
	'Response.Write sQuery
	with objrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source =sQuery
		.Open
	end with
	if not objrs.EOF then
		set suppNode = oDOM.createElement("Supplier")
			suppNode.setAttribute "SuppShortCode",objrs(0)
			suppNode.setAttribute "SuppCode",objrs(1)
			suppNode.setAttribute "SuppName",objrs(2)
			suppNode.setAttribute "AgentCode","N"
			suppNode.setAttribute "AgentName",""
			suppNode.setAttribute "PartyCode",objrs(3)
			suppNode.setAttribute "PartyType",objrs(4)
			suppNode.setAttribute "PartySubType",objrs(5)
		Root.appendChild suppNode
	else
			with objRs1
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source ="Select PartyType,PartySubType,PartyCode,PartyName from vwOrgParty where PartyCode = "& sSupAgent
				.Open
			end with
			if not objRs1.EOF then
				set suppNode = oDOM.createElement("Supplier")
					suppNode.setAttribute "SuppShortCode",""
					suppNode.setAttribute "SuppCode",""
					suppNode.setAttribute "SuppName",objrs1(3)
					suppNode.setAttribute "AgentCode","N"
					suppNode.setAttribute "AgentName",""
					suppNode.setAttribute "PartyCode",objrs1(2)
					suppNode.setAttribute "PartyType",objrs1(0)
					suppNode.setAttribute "PartySubType",objrs1(1)
				Root.appendChild suppNode
			end if
			objRs1.Close
	end if
	objrs.Close
	sQuery = "Select GATEPASSNO,ENTRYNO,isNull(ITEMCODE,0),isNull(CLASSIFICATIONCODE,0),QUANTITY,isNull(DESCRIPTION,''),"&_
			 "INVOICEDUOM,isNull(ItemValue,0),isNull(FormJJ,'N'),isNull(Reason,'') from FORGATEPASSDETAILS Where GatePassNo= "& nGatePassNo

	'Response.Write sQuery
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sQuery
		.Open
	end with
	if not objrs.EOF then
		do while not objrs.EOF
			SET nodeDetail = oDOM.createElement("DETAILS")
				sItemCode = objrs(2)
				sClassCode = objrs(3)
				sOtherDesc = objrs(5)
				sReason = objRs(9)
				if sItemCode <>"" and sClassCode<>"" then
					with objRs1
						.CursorLocation = 3
						.CursorType = 3
						.ActiveConnection = con
						.Source = "Select ItemDescription from vwItem where ItemCode ="& sItemCode  &"  and ClassificationCode = "&sClassCode
						.Open
					end with
					if not objRs1.EOF then
						sItemDesc = objrs1(0)
						sItemDesc = Replace(sItemDesc,"'","")
					end if
					objRs1.Close
				end if 'if sItemCode <>"" and sClassCode<>"" then
				if sItemDesc <>"" and sOtherDesc<>"" then
					sDesc= sItemDesc  &"-"& sOtherDesc
				elseif sItemDesc <>"" and trim(sOtherDesc)="" then
					sDesc= sItemDesc
				else
					sDesc = sOtherDesc
				end if
			nodeDetail.setAttribute "OTHERDESC",objrs(5)
			nodeDetail.setAttribute "QTY",objrs(4)
			nodeDetail.setAttribute "ITEMCODE",sItemCode
			nodeDetail.setAttribute "CLASSCODE",sClassCode
			nodeDetail.setAttribute "UOM",objrs(6)
			nodeDetail.setAttribute "VALUE",objrs(7)
			nodeDetail.setAttribute "FORMJJ",objrs(8)
			nodeDetail.setAttribute "DESC",sDesc
			nodeDetail.setAttribute "ITEMDESC",sItemDesc
			nodeDetail.setAttribute "REASON",sReason
			Root.appendChild nodeDetail
			objrs.MoveNext
		loop
	end if
	objrs.Close

oDOM.save server.MapPath("../temp/transaction/GatePassServiceAmd.xml")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Gate Pass</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<script type="application/xml" data-itms-xml-island="1" id="UoMData" data-src="../../inventory/xmldata/Uom.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="Data"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="newData" data-src="../temp/transaction/GatePassServiceAmd.xml"></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/gatePassServiceEntryAmd.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="Init()">
	<form method="POST" name="formname">
	<input type=hidden name="hItem" value="">
	<input type=hidden name="hClass" value="">
	<input type=hidden name="hForUnit" value="<%=sForUnit%>">
	<input type=hidden name="hItemType" value="<%=sItemType%>">
	<input type=hidden name="hRemarks" value="<%=sRemarks%>">
	<input type=hidden name="hTransport" value="<%=sTransport%>">
	<input type=hidden name="hDeliveryBy" value="<%=sDelivery%>">
	<input type=hidden name="hTakenBy" value="<%=sTakenBy%>">
	<input type=hidden name="hPartyCode" value="<%=sSupAgent%>">
	<input type=hidden name="hRows" value="">
	<input type=hidden name="hGatePassNo" value="<%=nGatePassNo%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Gate Pass - Service
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
												<!--<tr>
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

												<!--<tr>
													<td class="FieldCellSub">Item Type</td>
													<td class="FieldCellSub" valign="top">
                                                        <select size="1" name="selItmType" class="FormElem">
															<option value="select">Select</option>
															<%	'Calling the Function which populates the Item Type list
																'populateItemType
															%>
														</select>
													</td>
												</tr>-->

												<tr>
													<td class="FieldCellSub">Party</td>
													<td class="FieldCellSub">
														<input type="text" name="txtRefName" value size="60" class="formelemread" readonly>&nbsp;
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Party" width="11" height="11" onClick="popSuppAgent()"></a>
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Item Description</td>
													<td class="FieldCellSub">
														<input type="text" name="txtItemDesc" value size="60" class="formelemread" readonly>&nbsp;
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item" width="11" height="11" onClick="SelectItem()"></a>
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Other Description</td>
													<td class="FieldCellSub">
														<input type="text" name="txtDesc" maxlength=50 size="60" class="formelem">
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Reason</td>
													<td class="FieldCellSub">
														<input type="text" name="txtReason" maxlength=35 size="60" class="formelem" value="SENT FOR REPAIRS - TO BE RETURNED">
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Quantity & Value</td>
													<td class="FieldCellSub">
														<input type="text" name="txtQty" value size="12" class="formelem" onkeypress="DoKeyPress('Y',7,3)">&nbsp;
														<select size="1" name="selUOM" class="FormElem">
															<option value="select">Select</option>
															<%
															populateUoM()
															%>
														</select>
														&nbsp;
														<input type="text" name="txtValue" size="12" class="formelem" onkeypress="DoKeyPress('Y',7,2)" >&nbsp;
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Form JJ Applicable</td>
													<td class="FieldCellSub">
														<input type="checkbox" name="ChkFormJJ" value="Y" class="formelem" >
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
													<td class="ExcelHeaderCell" align="center" width=10>
														<img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onclick="DelItem()">
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
													<textarea rows="3" name="txtRemarks" cols="50" class="Formelem"></textarea>
												</td>
												<td class="FieldCell">Transport</td>
												<td class="FieldCellSub">
													<input type="text" name="txtTransport" maxlength=50 size="50" class="formelem">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Taken By</td>
												<td class="FieldCellSub">
													<input type="text" name="txtTakenBy" maxlength=50 size="50" class="formelem">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Delivery By</td>
												<td class="FieldCellSub">
													<input type="text" name="txtDeliveryBy" maxlength=50 size="50" class="formelem">
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
													<p align="center">
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

