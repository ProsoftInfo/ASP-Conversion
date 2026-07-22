<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	itmDetailsPop.asp
	'Module Name				:	Inventory (Item)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	April 15, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrsIssueItemEntry.asp
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
<!-- #include File="../../include/populate.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Details </TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmDetailsPopModern.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit()">
<%
	dim dcrs,dcrs1,dcrs2,dcrs3

	dim sOrgName,sClassName,sClassCode,iItmCode,sItmDescr,sItmShDesc,sItmAddDesc,sItmType,sItmController
	dim sItmStore,sItmManu,sItmPur,sItmSales,sItmDesc,sItmStoreF,sItmStoreO,sItmManuF,sItmManuO
	dim sItmPurF,sItmPurO,sItmSalesF,sItmSalesO,sOrgCode,sComCode,iCtr,sItmTypeCode,arrTemp
	dim sVarCode,sAttrValue,sAttrOpValue
	Dim sTemp,sAttributeList,iOptVal,sOptName,i,rsAtt,sSql
	'Response.Write "<textarea>"&Request.QueryString&"</textarea>"
	'sOrgName = trim(Request.QueryString("sOrgName"))
	arrTemp = split(trim(Request.QueryString("sTemp")),"A")
	sClassCode = trim(arrTemp(1))
	iItmCode = trim(arrTemp(2))
	sOrgCode = trim(arrTemp(3))
	'sAttributeList = trim(arrTemp(4))
	sAttributeList =  Request.QueryString("AttbVal")
	'Response.Write "sAttributeList="& sAttributeList
	set dcrs = Server.CreateObject("ADODB.RecordSet")
	set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	set dcrs2 = Server.CreateObject("ADODB.RecordSet")
	set dcrs3 = Server.CreateObject("ADODB.RecordSet")
	set rsAtt = Server.CreateObject("ADODB.RecordSet")
	If trim(sAttributeList) <> "" then
		sOptName = ""
		sTemp = split(sAttributeList,",")
		For i = 0 to UBOUND(sTemp)
			iOptVal = sTemp(i)
			sSql = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "&iOptVal&" "
			rsAtt.Open sSql,con
			If not rsAtt.EOF then
				sOptName =sOptName &","& rsAtt(0)
			End If
			rsAtt.Close
		Next
	End If
	IF sOptName <> "" then
		sOptName = " [" & mid(sOptName,2) &"] "
	End IF


	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT SHORTDESCRIPTION,ITEMDESCRIPTION,ADDITIONALDESCRIPTION,STORESUOM,PURCHASEUOM,PURTOSTORERATE,PURTOSTOREOPERATOR,MANUFACTURINGUOM,MANTOSTORERATE,MANTOSTOREOPERATOR,SALESUOM,SALETOSTORERATE,SALETOSTOREOPERATOR,ITEMTYPENAME,EMPLOYEENAME,COMPANYITEMCODE,GROUPNAME FROM VWITEMDETAILS WHERE ITEMCODE = " & iItmCode & " AND GROUPCODE = " & sClassCode & ""
		'.Source = "SELECT SHORTDESCRIPTION,ITEMDESCRIPTION,ADDITIONALDESCRIPTION,STORESUOM,ISNULL(PURCHASEUOM,'1'),ISNULL(PURTOSTORERATE,0),PURTOSTOREOPERATOR,ISNULL(MANUFACTURINGUOM,'1'),ISNULL(MANTOSTORERATE,0),MANTOSTOREOPERATOR,ISNULL(SALESUOM,'1'),ISNULL(SALETOSTORERATE,0),SALETOSTOREOPERATOR,ITEMTYPENAME,EMPLOYEENAME,COMPANYITEMCODE,GROUPNAME FROM VWITEMDETAILS WHERE ITEMCODE = " & iItmCode & " AND GROUPCODE = " & sClassCode & ""
		'.Source = "SELECT SHORTDESCRIPTION,ITEMDESCRIPTION,ADDITIONALDESCRIPTION,STORESUOM,ISNULL(PURCHASEUOM,'1'),ISNULL(PURTOSTORERATE,0),PURTOSTOREOPERATOR,ISNULL(MANUFACTURINGUOM,'1'),ISNULL(MANTOSTORERATE,0),MANTOSTOREOPERATOR,ISNULL(SALESUOM,'1'),ISNULL(SALETOSTORERATE,0),SALETOSTOREOPERATOR,ITEMTYPENAME,COMPANYITEMCODE,GROUPNAME FROM VWITEM WHERE ITEMCODE = " & iItmCode & " "
		.Source = "SELECT SHORTDESCRIPTION,ITEMDESCRIPTION,ADDITIONALDESCRIPTION,STORESUOM,ISNULL(PURCHASEUOM,'1'),ISNULL(PURTOSTORERATE,0),PURTOSTOREOPERATOR,ISNULL(MANUFACTURINGUOM,'1'),ISNULL(MANTOSTORERATE,0),MANTOSTOREOPERATOR,ISNULL(SALESUOM,'1'),ISNULL(SALETOSTORERATE,0),SALETOSTOREOPERATOR,COMPANYITEMCODE,GROUPNAME FROM VWITEM WHERE ITEMCODE = " & iItmCode & " "
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		sItmShDesc = dcrs(0)
		sItmDescr = dcrs(1)
		sItmAddDesc = dcrs(2)
		sItmStore = dcrs(3)
		sItmPur = dcrs(4)
		sItmPurF = dcrs(5)
		sItmPurO = dcrs(6)
		sItmManu = dcrs(7)
		sItmManuF = dcrs(8)
		sItmManuO = dcrs(9)
		sItmSales = dcrs(10)
		sItmSalesF = dcrs(11)
		sItmSalesO = dcrs(12)
		'sItmType = dcrs(13)
		'sItmController = dcrs(14)
		sItmController = ""
		sComCode = dcrs(13)
		sClassName = dcrs(14)

		if sItmPurO = "1" then
			sItmPurO = "/"
		elseif sItmPurO = "0" then
			sItmPurO = "*"
		end if
		if sItmPur = "1" then
			sItmPurO = "N/A"
		end if

		if sItmManuO = "1" then
			sItmManuO = "/"
		elseif sItmManuO = "0" then
			sItmManuO = "*"
		end if
		if sItmManu = "1" then
			sItmManuO = "N/A"
		end if

		if sItmSalesO = "1" then
			sItmSalesO = "/"
		elseif sItmSalesO = "0" then
			sItmSalesO = "*"
		end if
		if sItmSales = "1" then
			sItmSalesO = "N/A"
		end if

	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(COMPANYITEMCODE,'-') FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sVarCode = trim(dcrs(0))
	end if
	dcrs.Close
	IF trim(sOptName) <> "" then sItmDescr = sItmDescr & sOptName
	'Response.Write "sItmDescr ="&sItmDescr
%>
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Item Details
		</td>
    </tr>
    <tr>
    	<td class="TopPack"></td>
    </tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
									<table cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td class="FieldCell" width="135"> Organization</td>
											<td class="FieldCellSub">
												<span class="DataOnly" id="idOrgName">&nbsp;</span>
                                            </td>
											<td rowspan="9" width="225" valign="top">
												<%
													with dcrs
														.CursorLocation = 3
														.CursorType = 3
														.Source = "SELECT ISNULL(ITEMTYPEID,'') FROM INV_M_CLASSIFICATION WHERE GROUPCODE = " & sClassCode & ""
														.ActiveConnection = con
														.Open
													end with
													set dcrs.ActiveConnection = nothing
													'Response.Write dcrs.source
													if not dcrs.EOF then
														sItmTypeCode = trim(dcrs(0))
														if sItmTypeCode <> "" then 'Condition added by Maheshwari on 23rd Oct 2007
															with dcrs1
																.CursorLocation = 3
																.CursorType = 3
																'.Source = "SELECT DISTINCT HEADERID,ITEMTYPEHEADERNAME FROM INV_M_ITEMTYPEHEADER WHERE HEADERID = 2 AND ITEMTYPEID = " & Pack(sItmTypeCode) & " ORDER BY HEADERID"
																.Source = "SELECT DISTINCT HEADERID,ITEMTYPEHEADERNAME FROM INV_M_ITEMTYPEHEADER WHERE HEADERID = 2  ORDER BY HEADERID"
																.ActiveConnection = con
																.Open
															end with
															set dcrs1.ActiveConnection = nothing
															if not dcrs1.EOF then
													%>

															<DIV class=frmBody id=frm2 style="width: 220;">
																<table border="0" cellspacing="1" class="ExcelTable" width="100%">
																	<tr>
																		<td class="ExcelHeaderCell" align="center" width="25">S.No</td>
																		<%	if sItmTypeCode = "YRN" then %>
																				<td class="ExcelHeaderCell" align="center" width="90">Additional Specification</td>
																		<%	else %>
																				<td class="ExcelHeaderCell" align="center" width="90"><%=trim(dcrs1(1))%></td>
																		<%	end if %>
																		<td class="ExcelHeaderCell" align="center" width="61">Value</td>
																	</tr>
																</table>
															</div>
															<DIV class=frmBody id=frm1 style="width: 220; height: 135">
																<table border="0" cellspacing="1" class="ExcelTable" width="220">
																<%
																	iCtr = 0
																	with dcrs2
																		.CursorLocation = 3
																		.CursorType = 3
																		'.Source = "SELECT ITEMTYPEATTRIBUTEID,HEADERID,ITEMTYPEATTRIBUTENAME,ITEMTYPEATTRIBUTETYPE FROM INV_M_ITEMTYPEATTRIBUTES WHERE HEADERID = " & trim(dcrs1(0)) & " AND ITEMTYPEID = " & Pack(sItmTypeCode) & " ORDER BY ITEMTYPEATTRIBUTEID"
																		.Source = "SELECT ITEMTYPEATTRIBUTEID,HEADERID,ITEMTYPEATTRIBUTENAME,ITEMTYPEATTRIBUTETYPE FROM INV_M_ITEMTYPEATTRIBUTES WHERE HEADERID = " & trim(dcrs1(0)) & " ORDER BY ITEMTYPEATTRIBUTEID"
																		.ActiveConnection = con
																		.Open
																	end with
																	set dcrs2.ActiveConnection = nothing
																	do while not dcrs2.EOF
																		iCtr = iCtr + 1

																		with dcrs3
																			.CursorLocation = 3
																			.CursorType = 3
																			.Source = "SELECT ATTRIBUTEDATA,OPTIONVALUE FROM INV_M_ITEMTYPEVALUE WHERE ITEMTYPEATTRIBUTEID = " & trim(dcrs2(0)) & " AND ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode& " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
																			.ActiveConnection = con
																			.Open
																		end with
																		set dcrs3.ActiveConnection = nothing

																		if not dcrs3.EOF then
																			sAttrValue = trim(dcrs3(0))
																			sAttrOpValue = trim(dcrs3(1))
																		else
																			sAttrOpValue = "NULL"
																		end if
																		dcrs3.Close
																%>
																	<tr>
																		<td class="ExcelSerial" align="center" width="25"><%=iCtr%></td>
																		<td class="ExcelDisplayCell" align="left" valign="top" width="90"><%=trim(dcrs2(2))%></td>
																		<td class="ExcelDisplayCell" width="10">
																			<%	if lcase(trim(dcrs2(3))) = "numeric" then %>
																			<input type="text" NAME="txt<%=cint(trim(dcrs2(0)))%>" VALUE="<%=sAttrValue%>" size="20" class="FormElemRead" READONLY>
																			<%	elseif lcase(trim(dcrs2(3))) = "string" then %>
																			<input type="text" NAME="txt<%=cint(trim(dcrs2(0)))%>" VALUE="<%=sAttrValue%>" size="20" class="FormElemRead" READONLY>
																			<%	elseif lcase(trim(dcrs2(3))) = "options" then %>
																			<input type="text" NAME="a<%=cint(trim(dcrs2(0)))%>" VALUE="<%=populateOptionList(trim(dcrs2(0)),sAttrOpValue)%>" size="20" class="FormElemRead" READONLY>
																			<%	end if %>
																		</td>
																	</tr>
																	<%	dcrs2.MoveNext
																		loop
																		dcrs2.Close
																	%>
																</table>
															</div>
														<%
																	end if
																	dcrs1.Close
															end if
														end if 'if sItmTypeCode <> ""
												dcrs.Close
											%>

                                            </td>
										</tr>
										<tr>
											<td class=FieldCell width="135"> Item Code</td>
											<td class=FieldCellSub>
												<span class="DataOnly"><%=sComCode%>&nbsp;</span>&nbsp;
												<%	if sItmTypeCode = "YRN" then %>

														<img border="0" name="bn:<%=sComCode%>:<%=sClassCode%>:<%=iItmCode%>:<%=sOrgCode%>" src="../../assets/images/iTMS%20Icons/Details.gif" width="15" height="15" style="cursor:hand;" alt="View Item Specification" onClick="DisplayItemCode(this,'<%=GetItemSpec(sComCode)%>')">

												<%	end if %>
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell width="135"> Item Variant Code</td>
											<td class=FieldCellSub>
												<SPAN class="DataOnly"><%=sVarCode%>&nbsp;</SPAN>
											</td>
										</tr>
										<tr>
											<td class=FieldCell width="135"> Short Description</td>
											<td class=FieldCellSub>
												<SPAN class="DataOnly"><%=sItmShDesc%>&nbsp;</SPAN>
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell width="135"> Description</td>
											<td class=FieldCellSub>
                                            <SPAN class="DataOnly"><%=sItmDescr%>&nbsp;</SPAN>
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell width="135"> Additional Description</td>
											<td class=FieldCellSub>
                                            <SPAN class="DataOnly"><%=sItmAddDesc%>&nbsp;</SPAN>
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell width="135"> Item Type</td>
											<td class=FieldCellSub>
                                            <SPAN class="DataOnly"><%=sItmType%>&nbsp;</SPAN>
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell width="135"> Controller</td>
											<td class=FieldCellSub>
                                            <SPAN class="DataOnly"><%=sItmController%>&nbsp;</SPAN>
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell width="135"> Item Classification</td>
											<td class=FieldCellSub>
                                            <SPAN class="DataOnly"><%=sClassName%>&nbsp;</SPAN>
                                            </td>
										</tr>
									</table>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td align="center" class="FieldCell" width="50%">
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="147"><p align="center">UoM &amp; Basic Conversion
                                                            </td>
												</center>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable>
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack colspan="2"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub width="87"> Stores</td>
															<td>
                                                            <SPAN class="DataOnly"><%=DisplayUoM(sItmStore)%>&nbsp;</SPAN>
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCellSub width="87"> Purchase</td>
															<td>
                                                            <SPAN class="DataOnly"><%=DisplayUoM(sItmPur)%>&nbsp;</SPAN>
                                                            </td>
														</tr>
												</center>
														<tr>
															<td class=FieldCellSub width="87"> Manufacturing</td>
															<td>
                                                            <SPAN class="DataOnly"><%=DisplayUoM(sItmManu)%>&nbsp;</SPAN>
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCellSub width="87"> Sales</td>
															<td>
                                                            <SPAN class="DataOnly"><%=DisplayUoM(sItmSales)%>&nbsp;</SPAN>
                                                            </td>
														</tr>
														<tr>
															<td class=MiddlePack colspan="2"> </td>
														</tr>
													</table>
                                                            </td>
														</tr>
													</table>
											</td>
											<td align="center" class="FieldCell" width="5">
                                                <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
											<td align="center" class="FieldCell" width="50%" valign="bottom">
                                                <table border="0" cellspacing="1" class="TableOutlineOnly" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" width="140"><p align="center">Conversion</p>
                                        </td>
                                        <td class="ExcelHeaderCell" width="50"><p align="center">Factor</td>
                                        <td class="ExcelHeaderCell"><p align="center">Operator</td>
                                            </tr>
                                            <tr>
                                        <td class="MiddlePack" colspan="3"></td>
                                            </tr>
                                            <tr>
                                        <td class="FieldCell" width="140">Stores to Purchase</td>
                                        <td width="50">
                                            <SPAN class="DataOnly"><%=sItmPurF%>&nbsp;</SPAN>
												</td>
												<td>
										        <SPAN class="DataOnly"><%=sItmPurO%>&nbsp;</SPAN>
                                        </td>
                                            </tr>
                                            <tr>
                                        <td class="FieldCell" width="140">Stores to Manufacturing</td>
                                        <td width="50">
                                        <SPAN class="DataOnly"><%=sItmManuF%>&nbsp;</SPAN>
                                        </td>
                                        <td >
                                         <SPAN class="DataOnly"><%=sItmManuO%>&nbsp;</SPAN>
                                        </td>
                                            </tr>
                                            <tr>
                                        <td class="FieldCell" width="140">Stores to Sales</td>
                                        <td width="50">
                                         <SPAN class="DataOnly"><%=sItmSalesF%>&nbsp;</SPAN>
                                        </td>
                                        <td>
                                         <SPAN class="DataOnly"><%=sItmSalesO%>&nbsp;</SPAN>
                                        </td>
                                            </tr>
                                            <tr>
                                        <td class="MiddlePack" colspan="3"></td>
                                            </tr>
                                                </table>
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
								</td>
							</tr>
							<%
								Dim iBomApplicability

								sSql ="Select isNull(BOMApplicability,0) from INV_M_ItemMaster where ItemCode = "& iItmCode
								dcrs.open sSql,con
								if not dcrs.eof then
									iBomApplicability = dcrs(0)
								end if
								dcrs.close

								if Trim(iBomApplicability)="1" then

							%>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td align="center" class="FieldCell" colspan="3" width="100%">
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="100"><p align="center">Bill Of Materials
                                                            </td>
												</center>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable>
																<center>

																	<table cellpadding="0" cellspacing="1" width="100%">
																		<tr>
																			<td class="ExcelHeaderCell" align="center">S.No</td>
																			<td class="ExcelHeaderCell" align="center">Item Description</td>
																			<td class="ExcelHeaderCell" align="center">Quantity</td>
																			<td class="ExcelHeaderCell" align="center">UOM</td>
																			<td class="ExcelHeaderCell" align="center">Type</td>
																			<td class="ExcelHeaderCell" align="center">Consumable</td>
																		</tr>
																	<%
																		Dim iSNo
																		iSNo = 0
																		sSql = "Select ItemDescription,Quantity,UOM,Type,Consumable from INV_M_ItemMasterBOM A,VWItem B where A.BOMItemCode=B.ItemCode and A.ItemCode = "& iItmCode
																		dcrs1.Open sSql,con
																		if not dcrs1.EOF then
																			do while not dcrs1.EOF
																				iSNo = iSNo + 1
																				%>
																					<tr>
																						<td class="ExcelSerial" align="center"><%=iSNo%></td>
																						<td class="ExcelDisplayCell"><%=trim(dcrs1(0))%></td>
																						<td class="ExcelDisplayCell" align="right"><%=FormatNumber(dcrs1(1),2,0,0,0)%></td>
																						<td class="ExcelDisplayCell" align="center"><%=trim(dcrs1(2))%></td>
																						<td class="ExcelDisplayCell" align="center">
																							<%
																								if Trim(dcrs1(3))="F" then
																									Response.Write "Final Component"
																								else
																									Response.Write "Assembly"
																								end if
																							%>
																						</td>
																						<td class="ExcelDisplayCell" align="center">
																							<%
																								if Trim(dcrs1(4))="Y" then
																									Response.Write "Yes"
																								else
																									Response.Write "No"
																								end if
																							%>
																						</td>
																					</tr>
																				<%
																				dcrs1.MoveNext
																			loop
																		end if 'if not dcrs1.EOF then
																		dcrs1.Close
																	%>
																	</table>
																</center>
                                                            </td>
														</tr>
													</table>
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
								</td>
							</tr>

							<%end if 'if Trim(iBomApplicability)="1" then%>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<input type="button" value="Close" name="B1" class="ActionButton" onClick="javascript:window.close()" >
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

<%
	' Function to populate UoM
	Function DisplayUoM(sUoM)
		' Declaration of variables
		Dim dcrs,sUoMDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = " & Pack(sUoM) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMDesc = dcrs(0)

		if Not dcrs.EOF then
			DisplayUoM = sUoMDesc
		else
			DisplayUoM = "N/A"
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to populate Option List
	Function populateOptionList(iAID,iOpVal)
		' Declaration of variables
		Dim dcrs,iOptVal,sOptName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OPTIONVALUE,OPTIONNAME FROM INV_M_ITEMTYPEOPTIONS WHERE ITEMTYPEATTRIBUTEID = " & iAID & " AND OPTIONVALUE = " & iOpVal & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set iOptVal = dcrs(0)
		set sOptName = dcrs(1)

		if Not dcrs.EOF then
			populateOptionList = trim(sOptName)
		else
			populateOptionList = ""
		end if
		dcrs.Close

	End Function
%>

<%
	' Function to populate Codes List
	Function GetItemSpec(iID)
		' Declaration of variables
		Dim dcrs,sCodeType,iCodeTypeLen,arrCodeTypeLen,arrCode,iCtr,arrCodeType
		dim sTemp,sCode,iLen,iInnerCtr,iRecCount,sValue,arrCodeTypeFromLen
		iLen = 0
		arrCodeTypeLen = array("1","3","2","1","1","2","1","1","1","1","1")
		arrCodeTypeFromLen = array("1","2","5","7","8","9","11","12","13","14","15")

		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CODETYPENAME FROM APP_M_CODETYPES ORDER BY CODETYPE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		arrCodeType = dcrs.getRows()
		dcrs.Close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT * FROM APP_M_CODEMASTER ORDER BY CODETYPE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		arrCode = dcrs.getRows()
		iRecCount = cdbl(dcrs.RecordCount)
		dcrs.Close

		for iCtr = 1 to 11
			iLen = iLen + arrCodeTypeLen(iCtr - 1)
			sTemp = mid(iID,arrCodeTypeFromLen(iCtr - 1),arrCodeTypeLen(iCtr - 1))
			for iInnerCtr = 0 to iRecCount - 1
				if lcase(arrCode(0,iInnerCtr)) = lcase(sTemp) and arrCode(2,iInnerCtr) = iCtr then
					sValue = sValue & "|" & arrCodeType(0,iCtr - 1) & ":" & arrCode(1,iInnerCtr)
				end if
			next
		next
		GetItemSpec = mid(sValue,2)
	End Function
%>
