<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	CatalogItemNonePopEntry.asp
	'Module Name				:	Inventory (Opening Stock Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	January 31, 2007
	'Modified By				:	S.MAHESHWARI
	'Modified On				:	October 05,2007
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
<%
	Dim oDom,Root,HeaderNode,newElem

	'Declaration of Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	' Declaration of variables
	Dim dcrs
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	dim iItem,iClass,sOrgID,sType,arrTemp,sStoName,iQty,sItmType,sFormCode
	dim bPackNoSerFlag,sComCode,sStoresUom,sTemp,sItemName,sAttr
	dim arrUoM,sUoMDesc,sUoMCode,sCheck,sPackNoFlag,bPackFlag

	bPackNoSerFlag = false
	bPackFlag = false

	sTemp = trim(Request.QueryString("sTemp"))
	sTemp = replace(sTemp,chr(34),"~~")
'	Response.Write "sTemp="& sTemp
	arrTemp = split(sTemp ,"``")
	sType	= arrTemp(0)
	iClass	= arrTemp(1)
	sOrgID	= arrTemp(2)
	iQty = arrTemp(5)
	sStoName = arrTemp(6)
	sStoresUom = arrTemp(7)
	sItemName = arrTemp(9)
	sItmType = arrTemp(10)
	sAttr = arrTemp(11)
	if sAttr <> "NULL" then sAttr = Replace(sAttr,":",",")
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

	sCheck = UoMDecimal(sStoresUom)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock - Details</TITLE>
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
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/catalogItemNonePop.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="LoadDetails('<%=trim(replace(sTemp,"'","`"))%>')">
<form method="POST" name="formname" action="">
<input type="hidden" name="hSno" value="">
<input type="hidden" name="hRec" value="<%=sType%>">
<input type="hidden" name="sTemp" value="<%=sTemp%>">
<input type="hidden" name="hItemType" value="<%=sItmType%>">
<input type="hidden" name="hPackNoFlag" value="<%=sPackNoFlag%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Stock Details
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
											<td align="left" class="FieldCell" >

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
																        <span class="DataOnly" id="idItemName"><%=replace(sItemName,"~~",chr(34))%>&nbsp;</span>
																    </td>
																</tr>
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Store -- Bin</td>
																    <td class="FieldCellSub">
																        <span class="DataOnly" id="idStoreName"><%=sStoName%>&nbsp;</span>
																    </td>
																</tr>
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Nett Quantity&nbsp;</td>
																    <td class="FieldCellSub">
																        <span class="DataOnly" id="idQty"><%=iQty%></span> -
																        <span class="DataOnly" ><%=sStoresUom%>&nbsp;</span>
																    &nbsp;Quantity Entered&nbsp;
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
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>
																<tr>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Value</td>
																    <td class="FieldCellSub">
																		<input type="text" name="txtValue" size="15" onkeypress="DoKeyPress('Y',7,3)" class="FormElem">
																    </td>
																	<td class="FieldCell"></td>
																    <td class="FieldCell">Quantity</td>
																    <td class="FieldCellSub">
																		<input type="text" name="txtQuantity" size="11" onkeypress="DoKeyPress('<%=sCheck%>',7,3)" class="FormElem">
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
<!--Table 2-->

<!--Table 3-->
										<tr>
											<td align="center" class="FieldCell" width="100%" colspan="2">
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td>
															<table cellpadding="0" cellspacing="0">
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>
																<tr>
																	<td align="center" width="6">
																		<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																	</td>
																	<td valign="top" > <span class="ExcelHeaderCell">&nbsp; Attributes &nbsp;</span>
																          <table border="0" cellspacing="1" class="ExcelTable" width="450">
																			<tr>
																			  <td class="ExcelDisplayCell" align="left">
																		<%
																			with dcrs
																				.CursorLocation = 3
																				.CursorType = 3
																				.Source = "SELECT ITEMTYPEATTRIBUTEID,ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES WHERE ITEMTYPEATTRIBUTEID IN (" & sAttr & ") ORDER BY 1"
																				.ActiveConnection = con
																				.Open
																			end with
																			set dcrs.ActiveConnection = nothing
																			if not dcrs.EOF then
																				Do While Not dcrs.EOF

																		%>


																				<select size="1" name="selAttrZ<%=trim(dcrs(0))%>" class="FormElem">
																					<option value="select"><%=trim(dcrs(1))%></option>
																			<%
																				populateAttrValues trim(dcrs(0))
																			%>
																				</select>

																		<%
																				dcrs.MoveNext
																				loop
																			end if
																			dcrs.Close

																		%>

																			  </td>
																			</tr>
																		</table>
																	</td>
																</tr>
																<tr>
																	<td align="center" colspan="3" class="MiddlePack">
																	</td>
																</tr>
																<tr>
														            <td class="FieldCell" colspan=3 align=center>
																		<input type="button" value="Add Details" name="B3" class="AddButtonX" onClick="AddDetails('<%=sCheck%>')">
																	</td>
														        </tr>
															</table>
                                                        </td>
													</tr>
												</table>
                                            </td>
										</tr>
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
									<div class="frmBody" id="frm2" style="width: 100%; height:220;">
										<table border="0" cellspacing="1" id="tblLot" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="Left" colspan="5" >Attribute</td>
											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" >
													<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif"  height="15" onClick="DeleteItems()"></a>
												</td>
												<td class="ExcelHeaderCell" align="center" width="50">Quantity</td>
												<td class="ExcelHeaderCell" align="center" width="50">Value</td>
												<td class="ExcelHeaderCell" align="center" width="50">BarCode</td>
											</tr>
										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>
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
	' Function to Display UoM
	Function DisplayUoM(iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & ")"
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
	' Function to populate UoM List
	Function populateUoMList(sUoM)
		' Declaration of variables
		Dim dcrs,sUomDesc,sUomShDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMDESCRIPTION,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT"' WHERE UOMCODE NOT IN (SELECT STORESUOM FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ") AND UOMCODE <> " & Pack(sUoM) & " ORDER BY UOMCODE"
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

<%
	' Function to populate Attribute Values List
	Function populateAttrValues(iLevel)
		' Declaration of variables
		Dim dcrs,iOptionCode,sOptionName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OPTIONVALUE,OPTIONNAME FROM INV_M_ITEMTYPEOPTIONS WHERE ITEMTYPEATTRIBUTEID = " & iLevel & " ORDER BY 1"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set iOptionCode = dcrs(0)
		set sOptionName = dcrs(1)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(iOptionCode)&""">"&trim(sOptionName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>

