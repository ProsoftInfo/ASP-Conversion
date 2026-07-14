<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	stockLotSerPopNew.asp
	'Module Name				:	Inventory (Opening Stock Lot and Serial Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 28, 2003
	'Modified By				:	RAGAVENDRAN R
	'Modified On				:	May 24,2010 'Update the Auto Generation Entry Selection as a default
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
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/UoMDecimal.asp"-->
<!--#include virtual="/include/ItemDisplay.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->
<%
	Dim oDom,Root,HeaderNode,newElem

	'Declaration of Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	' Declaration of variables
	Dim dcrs,rsTemp
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set rsTemp = Server.CreateObject("ADODB.RecordSet")

	dim iItem,iClass,sOrgID,sType,arrTemp,sStoName,iQty,sItmType,sFormCode
	dim bPackNoSerFlag,sComCode,sStoresUom,sTemp
	dim arrUoM,sUoMDesc,sUoMCode,sCheck,sPackNoFlag,bPackFlag
	Dim ct, StDate,sItemName,sAttributeID,sQuery
	ct = 0
	bPackNoSerFlag = false
	bPackFlag = false

	sTemp = trim(Request.QueryString("sTemp"))
	 'Response.write "sTemp="&sTemp
sTemp = replace(sTemp,"'","~")
sTemp = Replace(sTemp,Chr(34),"~~")
	arrTemp = split(sTemp ,"``")
	sType	= arrTemp(0)
	iItem	= arrTemp(1)
	iClass	= arrTemp(2)
	sOrgID	= arrTemp(3)
	iQty = arrTemp(6)
	sStoName = arrTemp(7)
	sStoresUom = arrTemp(8)
	sItmType =  arrTemp(11)
	sAttributeID=arrTemp(12)
'	Response.Write "sAttributeID = "& sAttributeID
	'Response.Write iItem
	if iItem = "0" then
		sItemName = replace(arrTemp(10),"~~","'")
	else
		sItemName = replace(ItemDisplay(iItem,"NULL"),"~~","'")
	end if
	If UBound(arrTemp) = 10 Then
		StDate = arrTemp(10)
	Else
		StDate  = arrTemp(11)
	End If
	IF  sItmType = "" then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMTYPEID FROM INV_M_CLASSIFICATION WHERE GROUPCODE = " & iClass & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then
			sItmType = trim(dcrs(0))
		end if
		dcrs.Close
	End IF
	
	''blocked by ragav on Feb 27, 2012
	''begin
'	with dcrs
'		.CursorLocation = 3
'		.CursorType = 3
'		'.Source = "SELECT COMPANYITEMCODE FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ""
'		.Source = "SELECT ITEMTYPEID FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ""
'		.ActiveConnection = con
'		.Open
'	end with
'	set dcrs.ActiveConnection = nothing
'	if not dcrs.EOF then
'		'sComCode = trim(dcrs(0))
'		sItmType = dcrs(0)
'		'sFormCode = mid(sComCode,8,1)
''	end if
'	dcrs.Close
''end

	'IF sFormCode <> "" then
	'	if sItmType <> "YRN" then
	'		sFormCode = "NULL"
	'	else
	'		sFormCode = Pack(sFormCode)
	'	end if
	'End IF
	if sItmType <> "YRN" then
		sFormCode = "NULL"
	else
		sFormCode = Pack(sFormCode)
	end if

	if sItmType = "YRN" then
		' Packing Number - Number Series
		' OrganisationCode,SeriesNumber,SeriesCode,Date
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'PN' AND ITEMTYPE = " & Pack(sItmType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
			.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'PN' AND ITEMTYPE = 'YRN' AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if dcrs.EOF then
			bPackNoSerFlag = true
		end if
		dcrs.close
	end if

	Set Root = oDOM.createElement("Packing")
	oDOM.appendChild Root
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT PACKINGCODE,PACKINGSHORTNAME,PACKINGNAME FROM APP_M_PACKINGTYPE WHERE PACKINGCODE IN (SELECT PACKINGCODE FROM SAL_R_ITEMTYPEPACK WHERE ITEMTYPEID = " & Pack(sItmType) & ")"
		.Source = "SELECT PACKINGCODE,PACKINGSHORTNAME,PACKINGNAME FROM APP_M_PACKINGTYPE" 'WHERE PACKINGCODE IN (SELECT PACKINGCODE FROM SAL_R_ITEMTYPEPACK WHERE ITEMTYPEID = " & Pack(sItmType) & ")"
		.ActiveConnection = con
		.Open
	end with
	'Response.Write dcrs.source
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		Do While Not dcrs.EOF
			Set newElem = oDOM.createElement("PACK")
			newElem.setAttribute "PACKCODE",trim(dcrs(0))
			newElem.setAttribute "PACKSHNAME",trim(dcrs(1))
			newElem.setAttribute "PACKNAME",trim(dcrs(2))
			Root.appendChild newElem
		dcrs.MoveNext
		Loop
	else
		bPackFlag = true
	end if
	dcrs.Close

	oDOM.Save server.MapPath("../temp/master/PACKING"&Session.SessionID&".xml")

	set oDOM = nothing

	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	Set Root = oDOM.createElement("Selling")
	oDOM.appendChild Root


	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS WHERE CODE = " & sFormCode & " AND ITEMTYPEID = " & Pack(sItmType) & ""
		.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		Do While Not dcrs.EOF
			Set newElem = oDOM.createElement("SELL")
			newElem.setAttribute "UNITID",trim(dcrs(0))
			newElem.setAttribute "UNITNAME",trim(dcrs(1))
			Root.appendChild newElem
		dcrs.MoveNext
		Loop
	else
		bPackFlag = true
	end if
	dcrs.Close

	oDOM.Save server.MapPath("../temp/master/SELLING"&Session.SessionID&".xml")

	set oDOM = nothing

	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	Set Root = oDOM.createElement("SellingForm")
	oDOM.appendChild Root

	GetSellingForm

	oDOM.Save server.MapPath("../temp/master/SELLINGFORM"&Session.SessionID&".xml")

	arrUoM = split(DisplayUoM(iItem),":")
	sUoMCode = arrUoM(0)
	sUoMDesc = arrUoM(1)
	sCheck = UoMDecimal(sUoMCode)

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT NUMBERINGTYPE FROM PRD_M_PACKINGNUMBERSERIESCHECK WHERE ORGANISATIONCODE = " & Pack(sOrgID)
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sPackNoFlag  = trim(dcrs(0))
	else
		sPackNoFlag = ""
	end if
	dcrs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock - Lot / Serial Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" id="Data" data-itms-xml-island>
<root/>
</script>
<script type="application/xml" id="SerialData" data-itms-xml-island>
<ROOT/>
</script>
<script type="application/xml" id="FabricData" data-itms-xml-island>
<ROOT/>
</script>

<script type="application/xml" id="PackingData" data-itms-xml-island data-src="<%="../temp/master/PACKING"&Session.SessionID&".xml"%>"></script>
<script type="application/xml" id="SellingData" data-itms-xml-island data-src="<%="../temp/master/SELLING"&Session.SessionID&".xml"%>"></script>
<script type="application/xml" id="SellingFormData" data-itms-xml-island data-src="<%="../temp/master/SELLINGFORM"&Session.SessionID&".xml"%>"></script>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/stockLotSerialModern.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="LoadDetails('<%=trim(sTemp)%>');init();">
<%	if sPackNoFlag = "" then %>
<SCRIPT LANGUAGE=javascript>
if (window.ITMSStockLotSerial) {
	window.ITMSStockLotSerial.closeForMissingPackingNumber();
}
</SCRIPT>
<%	end if
	if bPackFlag then
%>
<!-- Legacy packing type/selling form warning was disabled here. -->
<%	end if %>

<form method="POST" name="formname" action="">
<input type="hidden" name="hRec" value="<%=sType%>">
<input type="hidden" name="sTemp" value="<%=sTemp%>">
<input type="hidden" name="hItemType" value="<%=sItmType%>">
<input type="hidden" name="hPackNoFlag" value="<%=sPackNoFlag%>">
<input type="hidden" name="hDate" value="<%=StDate%>">
<input type="hidden" name="hAttributeList" value="<%=sAttributeID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot / Serial Details
		</td>
    </tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
											<td class="FieldCell"></td>
											<td class="FieldCellSub"></td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCell" colspan="2"></td>
                                        </tr>

<!--Table 1-->
										<tr>
											<td align="left" class="FieldCell" width="60%">

												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class='GroupTitleLeft' width="10">&nbsp;</td>
																	<td class='GroupTitle' width="50"><p align="center">Details</td>
																	<td class='GroupTitleRight'><p align="left">&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td class=GroupTable>
															<table cellpadding="0" cellspacing="0">
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>
																<tr>
																    <td class="FieldCell"></td>
																    <td class="FieldCell">Item</td>
																    <td class="FieldCellSub">
																        <span class="DataOnly" id="idItemName"><%=sItemName%>&nbsp;</span>
																    </td>
																</tr>
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Store -- Bin</td>
																    <td class="FieldCellSub">
																        <span class="DataOnly" id="idStoreName"><%=sStoName%>&nbsp;</span>
																    </td>
																</tr>
																<!--<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Nett Quantity&nbsp;</td>
																    <td class="FieldCellSub">
																        <span class="DataOnly" id="idQty"><%=iQty%></span> -
																        <span class="DataOnly" ><%=sStoresUom%>&nbsp;</span>
																    &nbsp;Quantity Entered&nbsp;
																		<span class="DataOnly" id="idQtyEntered">0</span> -
																        <span class="DataOnly"><%=sStoresUom%>&nbsp;</span>
																    </td>
																</tr>-->
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Quantity Entered&nbsp;</td>
																    <td class="FieldCellSub">
																    	<span class="DataOnly" id="idQtyEntered">0</span> -
																        <span class="DataOnly"><%=sStoresUom%>&nbsp;</span>
																    </td>
																</tr>
																
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>

															</table>
											            </td>
													</tr>
												</table>
											</td>

											<td align="left" class="FieldCell" width="40%">
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class='GroupTitleLeft' width="10">&nbsp;</td>
																	<td class='GroupTitle' width="90"><p align="center">Alternate UoM</td>
																	<td class='GroupTitleRight'><p align="left">&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td class=GroupTable>
															<table cellpadding="0" cellspacing="0">
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>
																<tr>
																    <td class="FieldCell"></td>
																    <td class="FieldCell">Quantity in Gross</td>
																    <td class="FieldCellSub">
																        <input type="text" name="txtAltGross" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem" >
																		<select size="1" name="selAltUom" class="FormElem" onChange="var t=this.options[this.selectedIndex].text;document.getElementById('idAltUom').textContent=(t.toLowerCase()==='select'?'':t)">
																			<option value="select">Select</option>
																			<%
																				populateUoMList sStoresUom
																			%>
																		</select>
																    </td>
																</tr>
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Quantity in Nett</td>
																    <td class="FieldCellSub">
																        <input type="text" name="txtAltNett" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem" >
																        <span class="DataOnly" id="idAltUom"></span>
																    </td>
																</tr>
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>
															</table>
											            </td>
													</tr>
												</table>
											</td>
										</tr>
<!--Table 1-->

<!--Table 2-->
										<tr>
											<td align="center" class="FieldCell" width="100%" colspan="2">
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0" width="100%">
																<tr>
																	<td class='GroupTitleLeft' width="10">&nbsp;</td>
																	<td class='GroupTitle' width="100"><p align="center">Quantity Details</td>
																	<td class='GroupTitleRight'><p align="left">&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td class=GroupTable>
															<table cellpadding="0" cellspacing="0">
																<tr>
																	<td align="center" colspan="5" class="MiddlePack">
																	</td>
																</tr>
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Receipt Numbering&nbsp;</td>
																    <td class="FieldCell">
																    <%	if sType = "S" then %>
																        <input type="text" name="txtRcptNumbering" READONLY size="15" value="SERIAL" class="FormElemRead">
																    <%	Elseif sType = "LS" then %>
																        <input type="text" name="txtRcptNumbering" READONLY size="15" value="LOT\SERIAL" class="FormElemRead">
																    <%	Elseif sType = "L" then %>
																        <input type="text" name="txtRcptNumbering" READONLY size="15" value="LOT" class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtRcptNumbering" size="15" class="FormElem">
																    <%	end if %>
																    </td>
																    <td class="FieldCellSub">Lot Number</td>
																    <td class="FieldCellSub">

																		<% If sType = "S" Then %>
																			<select size="1" name="sellot" class="FormElem" OnChange="" disabled >
																			<option value="select">select</option>
																			<option value="N">New Lot No.</option>
																		</select>
																		<input type="text" name="txtLotNumber" size="10" value="" class="FormElem" disabled>
																		<% Else %>
																    		<select size="1" name="sellot" class="FormElem" OnChange="DispLot()">
																				<option value="select">select</option>
																				<!--<option value="N">New Lot No.</option>-->
																				<option value="N">New Lot No.</option>
																			</select>
																			<input type="text" name="txtLotNumber" size="10" value="" class="FormElem">
																		<% End If %>

																	<!--<input type="checkbox" name="chkNew" value="1" >&nbsp;Auto Generate&nbsp;-->
																		<input type="checkbox" name="chkNew" value="1" checked>&nbsp;Auto Generate&nbsp;
																		<%	if sType = "S" then %>
																			Serial
																		<%	Elseif sType = "LS" then %>
																		    Lot & Serial
																		<%	Elseif sType = "L" then %>
																		    Lot
																		<%	else %>

																		<%	end if %>

																    </td>
																</tr>


																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Quantity Gross&nbsp;&nbsp;&nbsp;</td>
																    <td class="FieldCell">
																        <!--input type="radio" value="G" name="radQtyIn" class="FormElem">  Gross&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																        <input type="radio" value="N" name="radQtyIn" class="FormElem">  Nett &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-->
																    <%	if ucase(sItmType) = "FAB" then %>
																        <input type="text" name="txtQtyIn" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem" >
																	<%	else %>
																		<input type="text" name="txtQtyIn" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem" >
																    <%	end if	%>

																    </td>
																    <td class="FieldCellSub">Tare weight&nbsp;</td>
																    <td class="FieldCellSub">
																		<input type="text" name="txtTare" size="11" maxlength=10 onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem"> <%=sStoresUom%>
																		&nbsp; Rate / <%=sStoresUom%> :&nbsp;&nbsp;
																		<input type="text" name="txtValue" size="10" value="" class="FormElem">
																    </td>
																</tr>

																<tr>
																	<td class="FieldCell"></td>
														            <td class="FieldCell">Selling Form</td>
														            <td class="FieldCell">
																		<!--input type="radio" value="I" name="radSellType" class="FormElem" onclick="DisableselSellType(this)">  Individual
																		<input type="radio" value="U" name="radSellType" class="FormElem" onclick="DisableselSellType(this)">  Uniform&nbsp;&nbsp;-->
																		<select size="1" name="selSellType" class="FormElem">
																			<option value="select">-N/A-</option>
																		</select>
																	</td>
																	<td class="FieldCellSub">W/t. Per Form</td>
																		<!--input type="radio" value="I" name="radWeight" class="FormElem" onclick="DisableTxtWt(this)">  Individual
																		<input type="radio" value="U" name="radWeight" class="FormElem" onclick="DisableTxtWt(this)">  Uniform&nbsp;&nbsp;-->
																	<td class="FieldCellSub">
																		<input type="text" name="txtWeight" size="11" class="FormElem" value="0">
																		&nbsp; No of Packs : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																		<input type="text" name="txtNoOfPacks" size="11" class="FormElem">
														            </td>
														        </tr>
																<tr>
																	<td class="FieldCell"></td>
														            <td class="FieldCell">Packing Type</td>
														            <td class="FieldCell">
																		<!--input type="radio" value="I" name="radPackType" class="FormElem" onclick="DisableselPack(this)">  Individual
																		<input type="radio" value="U" name="radPackType" class="FormElem" onclick="DisableselPack(this)">  Uniform&nbsp;&nbsp;-->
																		<select size="1" name="selPackType" class="FormElem">
																			<option value="select">Select</option>
																		</select>
																	</td>
														            <td class="FieldCellSub">W/t. Per Pack</td>
														            <td class="FieldCellSub">
																		<input type="text" name="txtWeightPerPack" size="11" class="FormElem">
																		&nbsp; Tare w/t per pack:&nbsp;
																		<input type="text" name="txtPackTare" size="11" class="FormElem">
																	</td>

																</tr>
																<tr>
																	<td class="FieldCell"></td>
																	<td class="FieldCell">Packing Form &nbsp	</td>
																	<td class="FieldCell">
																		<select size="1" name="selForm" class="FormElem">
																			<option value="select">-N/A-</option>
																		</select>
																	</td>
																	<td class="FieldCellSub">Packing Quality </td>
																	<td class="FieldCellSub">
																		<select size="1" name="selStage" class="FormElem">
																		<option value="select">-N/A-</option>
																		<%
																			PopulateStage sItmType
																		%>
																		</select>&nbsp;&nbsp;&nbsp;&nbsp;
																		<%if trim(sAttributeID)="" or IsNull(sAttributeID) or sAttributeID="NULL" then%>
																				<input type="button" value="Add Details" name="B3" class="AddButtonX" onClick="AddDetails('<%=sCheck%>')">
																		<%end if%>
																	</td>
																</tr>
																<%if trim(sAttributeID)<>"" and not IsNull(sAttributeID) and sAttributeID<>"NULL" then%>
																<tr>
																	<td class="FieldCell"></td>
																	<td class="FieldCell">Attributes &nbsp	</td>
																	<td class="FieldCell">
																		<select size="1" name="selAttribute" class="FormElem">
																			<%
																			'	sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,ItemTypeID from  INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where "&_
																			'		" O.ItemTypeAttributeID = A.ItemTypeAttributeID and ItemTypeID='"& sItmType &"' and A.ItemTypeAttributeID in ("& sAttributeID &") Group by A.ItemTypeAttributeID,A.ItemTypeAttributeName,ItemTypeID"
																				sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,'' from  INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where "&_
																				    " O.ItemTypeAttributeID = A.ItemTypeAttributeID and A.ItemTypeAttributeID in ("& sAttributeID &") Group by A.ItemTypeAttributeID,A.ItemTypeAttributeName"
																				rsTemp.Open sQuery,con
																				if not rsTemp.EOF then
																					Response.Write "<option value="& trim(rsTemp(0))&">"&Trim(rsTemp(1))&"</option>"
																				end if
																				rsTemp.Close

																			'	sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,O.OptionValue,O.OptionName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O "&_
																			'	" where O.ItemTypeAttributeID = A.ItemTypeAttributeID and ItemTypeID = '"& sItmType &"' and A.ItemTypeAttributeID = "& sAttributeID
																				sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,O.OptionValue,O.OptionName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O "&_
																				    " where O.ItemTypeAttributeID = A.ItemTypeAttributeID and A.ItemTypeAttributeID = "& sAttributeID
																				rsTemp.Open sQuery,con
																				if not rsTemp.EOF then
																					do while not rsTemp.EOF
																						Response.Write "<option value="& trim(rsTemp(2))&">"&Trim(rsTemp(3))&"</option>"
																						rsTemp.MoveNext
																					loop
																				end if
																				rsTemp.Close
																			%>
																		</select>
																	</td>
																	<td class="FieldCellSub"><input type="button" value="Add Details" name="B3" class="AddButtonX" onClick="AddDetails('<%=sCheck%>')"></td>
																	<td class="FieldCellSub">
																	</td>
																</tr>
																<%end if%>
																<!--tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Lot Number</td>
																    <td class="FieldCellSub">
																    <%	if sType = "S" then %>
																        <input type="text" name="txtLotNumber" READONLY size="5" value="N/A" class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtLotNumber" size="15" class="FormElem">
																    <%	end if %>
																    &nbsp;
																    Serial No. Prefix&nbsp;
																    <%	if sType = "L" then %>
																        <input type="text" name="txtPrefix" READONLY value=0 size="12" maxlength=12 class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtPrefix" size="12" maxlength=12 class="FormElem">
																    <%	end if %>
																    &nbsp;From&nbsp;
																    <%	if sType = "L" then %>
																        <input type="text" name="txtSerialFrom" READONLY value=0 size="5" maxlength=5 class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtSerialFrom" size="5" maxlength=5 class="FormElem" onkeypress="DoKeyPress('<%=sCheck%>',5,1)">
																    <%	end if %>
																    &nbsp;To&nbsp;
																    <%	if sType = "L" then %>
																        <input type="text" name="txtSerialTo" READONLY value=0 size="5" maxlength=5 class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtSerialTo" size="5" maxlength=5 class="FormElem" onkeypress="DoKeyPress('<%=sCheck%>',5,1)">
																    <%	end if %>
																    &nbsp;Suffix&nbsp;
																    <%	if sType = "L" then %>
																        <input type="text" name="txtSuffix" READONLY value=0 size="13" maxlength=13 class="FormElemRead">
																    <%	else %>
																        <input type="text" name="txtSuffix" size="13" maxlength=13 class="FormElem">
																    <%	end if %>
																    </td>
																</tr-->

																<!--tr>
																	<td class="FieldCell"></td>
																	<%	if sType = "S" then %>
																    <td class="FieldCell">Serial Value</td>
																    <td class="FieldCell">
																		<input type="radio" value="I" name="radValue" class="FormElem" onclick="DisableTxtValue(this)">  Individual
																		<input type="radio" value="U" name="radValue" class="FormElem" onclick="DisableTxtValue(this)">  Uniform&nbsp;&nbsp;
																		<input type="text" name="txtValue" size="11" onkeypress="DoKeyPress('Y',7,3)" class="FormElem">
																    </td>
																	<%	else	%>
																    <td class="FieldCell">Lot Value</td>
																    <td class="FieldCellSub">
																		<input type="text" name="txtValue" size="15" onkeypress="DoKeyPress('Y',7,3)" class="FormElem">
																    </td>
																	<%	end if	%>
																</tr-->
																<!--tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Quantity</td>
																    <td class="FieldCell">
																		<input type="radio" value="I" name="radQty" class="FormElem" onclick="DisableTxtQty(this)">  Individual
																		<input type="radio" value="U" name="radQty" class="FormElem" onclick="DisableTxtQty(this)">  Uniform&nbsp;&nbsp;
																		<input type="text" name="txtQuantity" size="11" onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem">
																    </td>
																</tr>
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr-->
															</table>
                                                        </td>
													</tr>
												</table>
                                            </td>
										</tr>
<!--Table 2-->

<!--Table 3-->

<!--Table 3-->

                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td>
									<div class="frmBody" id="frm2" style="width: 100%; height:170;">
										<table border="0" cellspacing="1" id="tblLot" class="ExcelTable" width="100%">
											<tr>
											<%	if ucase(sItmType) = "FAB" then %>
												<td class="ExcelHeaderCell" align="Left" colspan="10" >Lot No. / Value [Rate/<%=sStoresUom%>]</td>
											<%	else	%>
												<td class="ExcelHeaderCell" align="Left" colspan="9" >Lot No. / Value [Rate/<%=sStoresUom%>]</td>
											<%	end if	%>
											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="10">S.No.</td>
												<!--td class="ExcelHeaderCell" align="center" rowspan="2">Serial No.</td-->
												<td class="ExcelHeaderCell" align="center" colspan="2">Received</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Packing Type</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="100">Packing No</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="50">Value</td>
											<%	if ucase(sItmType) = "FAB" then %>
												<td class="ExcelHeaderCell" align="center" rowspan="2" width="50">Quantity Breakup</td>
											<%	end if
												'if ucase(sType) = "S" then %>

											<%	'end if%>
												<td class="ExcelHeaderCell" align="center" colspan="2">Selling</td>
												<!--td class="ExcelHeaderCell" align="center" rowspan="2">Selling Form</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">W/t. Per Form</td-->
												<td class="ExcelHeaderCell" align="center" rowspan="2">Packing Form</td>

											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center">Gross</td>
												<td class="ExcelHeaderCell" align="center">Tare</td>
												<td class="ExcelHeaderCell" align="center">Form</td>
												<td class="ExcelHeaderCell" align="center">W/t. Per</td>
											</tr>

										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>
							<!--Begining of Code added by Tajudeen-->
							<tr>
							    <td class="FieldCell" colspan="2"><p align="center">
									<!--input type="button" value="Add Serial" name="B4" class="AddButtonX" onClick="AddSerial()">
									<input type="button" value="  Add Lot " name="B5" class="AddButtonX"  onClick="AddLot()"-->
								</td>
							</tr>
							<!--End of Code added by Tajudeen-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()" >
                                                    <input type="reset" value="Reset" name="B2" class="ActionButton" onClick="clearAll()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
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

<%
	' Function to populate Stage
	Function PopulateStage(sItmType)
		' Declaration of variables
		Dim dcrs, StageCode,StageName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STAGEID, STAGENAME FROM INV_M_STAGE WHERE ITEMTYPEID = " & Pack(sItmType) & " AND (CATEGORYID IS NULL OR CATEGORYID = 'Q')"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set StageCode = dcrs(0)
		set StageName = dcrs(1)

		do while Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(StageCode)&""">"&trim(StageName)&"</OPTION>" &vbcrlf)
			dcrs.movenext
		loop
		dcrs.Close

	End Function
%>

<%
	' Function to Display UoM
	Function DisplayUoM(iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ")"
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to Get Selling Form
	Function GetSellingForm()
		' Declaration of variables
		Dim dcrs,dcrs1,iCodeLen,iCodeSize,sForm
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	'Response.Write "sItmType="&sItmType
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT CODELENGTH FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form' AND ITEMTYPEID = " & Pack(sItmType) & ""
			.Source = "SELECT CODELENGTH FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		If not dcrs.EOF Then
			iCodeLen = cint(dcrs(0))
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				'.Source = "SELECT ISNULL(SUM(CODELENGTH)+1,0) FROM APP_M_CODETYPES WHERE DISPLAYORDER < (SELECT DISPLAYORDER FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form') AND ITEMTYPEID = " & Pack(sItmType) & ""
				.Source = "SELECT ISNULL(SUM(CODELENGTH)+1,0) FROM APP_M_CODETYPES WHERE DISPLAYORDER < (SELECT DISPLAYORDER FROM APP_M_CODETYPES WHERE LOWER(CODETYPENAME) = 'form') "
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			If not dcrs1.EOF Then
				iCodeSize = cint(dcrs1(0))
			end if
			dcrs1.Close
		end if
		dcrs.Close

		if iCodeSize > 0 and iCodeLen > 0 then
			sForm = trim(mid(sComCode,iCodeSize,iCodeLen))

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				'.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS WHERE CODE = " & Pack(sForm) & " AND ITEMTYPEID = " & Pack(sItmType) & ""
				.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS "
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing
			If not dcrs.EOF Then
				do while not dcrs.EOF
					'Response.Write "<option value="""&trim(dcrs(0))&""">"&trim(trim(dcrs(1)))&"</option>" & vbCrLf

					Set newElem = oDOM.createElement("SELLFORM")
					newElem.setAttribute "UNITID",trim(dcrs(0))
					newElem.setAttribute "UNITNAME",trim(dcrs(1))

					Root.appendChild newElem

				dcrs.MoveNext
				loop
			end if
			dcrs.Close
		end if
	End Function
%>

<%
	' Function to populate UoM List
	Function populateUoMList(sUoM)
		' Declaration of variables
		Dim dcrs,sUomDesc,sUomShDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMDESCRIPTION,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE NOT IN (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ") AND UOMCODE <> " & Pack(sUoM) & " ORDER BY UOMCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUoMCode = dcrs(0)
		set sUomDesc = dcrs(1)
		set sUomShDesc = dcrs(2)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUoMCode)&""">"&trim(sUomShDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
