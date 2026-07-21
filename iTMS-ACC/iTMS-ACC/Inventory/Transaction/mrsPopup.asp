<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	mrsPopup.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	TAJUDEEN S
	'Created On					:	June 08, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrs
	'Procedures/Functions Used	:	populateStore
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
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : MR Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data"><root/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsPopup.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
	function Print()
	{
		if(document.formname.hCtr.value==0)
		{
			if(confirm("No Records Found.\nDo you want to print.")==1)
			{
				//PrintWindow( "../reports/PRNMRDetails.asp?sTemp=" + document.formname.sTemp.value);
				PrintWindow( "../reports/PRNMRCreateDetails.asp?sTemp=" + document.formname.sTemp.value);
			}
		}
		else
			//PrintWindow( "../reports/PRNMRDetails.asp?sTemp=" + document.formname.sTemp.value);
			PrintWindow( "../reports/PRNMRCreateDetails.asp?sTemp=" + document.formname.sTemp.value);
	}
-->
</SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
	Function PrintNew()
		sTempValues = document.formname.hData.value
		PrintWindow( "../reports/PRNPickedDetails.asp?sTemp=" + sTempValues)
	End Function
</script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<%
	Dim dcrs,dcrs1,rstemp, iMRSNo, iQtyReq, iQtyIssued, iQtyAppr, iQtyPur, iQtyTransfer
	dim iItemCode, iClassCode, sItemName, sOrgId, iCtr, iQtyPending, sOrgName
	dim arrSchTemp, sSchTemp, sSchTempValue, sMRSDate, sType, sIssue, sMRStatus
	dim sIssueCode,sTemp,sAttribList,sOptName,iOptVal,i,iCounter,sPassData,sAttList,sArrList,sQuery
	Dim sIssToStr,sIssTypeCodestr,sArrRef,sReferenceName

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set rsTemp = Server.CreateObject("ADODB.RecordSet")

	iMRSNo =  Request.QueryString("MRSNO")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ORGUNITSHORTDESCRIPTION, CONVERT(VARCHAR,MRSDATE,103), MRSTYPE, 0, MRSFORUNIT,ISNULL(MRSCODE,MRSNUMBER),IssToType,IssToCode,IssToSubCode,IsNull(AppRefType,''),AppRefNo,IsNull(IssueTypeCode,'GEN') FROM VwMRSList WHERE MRSNUMBER = " & iMRSNo
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.eof then
		sOrgName = Trim(dcrs(0))
		sMRSDate = Trim(dcrs(1))
		sType = Trim(dcrs(2))
		sIssue = Trim(dcrs(3))
		sOrgId = Trim(dcrs(4))
		sIssueCode = Trim(dcrs(5))
		sIssToStr = IssuedToString(dcrs(6),dcrs(7),dcrs(8))

		sIssTypeCodestr = GetRcptIssName(dcrs(11))

		if trim(dcrs(9))<>"" and trim(dcrs(10))<>"" then
		    sArrRef = split(GetInfoRefType(dcrs(9),dcrs(10),sOrgID),":")
		    if trim(sArrRef(0))<>"" then
		        sReferenceName= sArrRef(0) & " ("& sArrRef(1) &" - "& sArrRef(2) &")"
		    else
		        sReferenceName = "None"
		    end if
		else
		    sReferenceName = "None"
		end if


	end if
	dcrs.close

	sTemp = sOrgId&":"&sMRSDate&":"&sMRSDate&":::"&sType&":"&iMRSNo&":"&sMRSDate
	sPassData = iMRSNo & ":" & sMRSDate & ":" & ""


	if sType = "0" then
		sType = "Returnable"
	else
		sType = "Non Returnable"
	end if


%>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname">
<input type=hidden name="hMRSNo" value="<%=iMRSNo%>">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hData" value="<%=sPassData%>">
<OBJECT id=penDet type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11" VIEWASTEXT>
<PARAM name="Command" value="HH Version">
</OBJECT>
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Material Requisition - Item Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%"  >
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">

										<tr>
										    <td class="FieldCell">Requested By</td>
										    <td class="FieldCellSub"><span class="DataOnly"><%=sIssToStr%>&nbsp;</span></td>
										    <td class="FieldCell">Requisition No - Date</td>
										    <td class="FieldCellSub"><span class="DataOnly"><%=sIssueCode%></span> - <span class="DataOnly"><%=sMRSDate%></span></td>
										</tr>
										<tr>
										    <td class="FieldCell">Requested For&nbsp;</td>
										    <td class="FieldCellSub"><span class="DataOnly" id="idOrgName"><%=sIssTypeCodestr%>&nbsp;</span></td>
										    <td class="FieldCell">Reference Type&nbsp;</td>
										    <td class="FieldCellSub"><span class="DataOnly" id="Span1"><%=sReferenceName%>&nbsp;</span></td>
										</tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel"></td>
								<td valign="top" class="FieldCell" width="100%"><center>
                                    <div align="left">
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td><center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="FieldCell">
																<DIV class="frmbody" id="frm2" style="width: 555; height:285;">
																	<table border="0" cellspacing="1" class="ExcelTable" width="100%">
																		<tr>
																			<td class="ExcelHeaderCell" align="center" width="10" rowspan="3">S.No.</td>
																			<td class="ExcelHeaderCell" align="center" colspan="4">Requisition Details</td>
																		</tr>
																	    <tr>
																			<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description</td>
																			<td class="ExcelHeaderCell" align="center">Approved</td>
																			<td class="ExcelHeaderCell" align="center">Pending</td>
																			<td class="ExcelHeaderCell" align="center">By Date's</td>
																	    </tr>
																		<tr>
																			<td class="ExcelHeaderCell" align="center">Issued</td>
																			<td class="ExcelHeaderCell" align="center">Tra. / PR</td>
																			<td class="ExcelHeaderCell" align="center">Quality</td>
																		</tr>
																	<%
																		with dcrs
																			.CursorLocation = 3
																			.CursorType = 3
																			.Source = "SELECT ISNULL(QUANTITYREQUESTED,0),ISNULL(QUANTITYAPPROVED,0),ISNULL(QUANTITYISSUED,0),(ISNULL(QUANTITYTOPURCHASE,0) - ISNULL(QUANTITYPURCHASED,0)),(ISNULL(QUANTITYFORTRANSFER,0) - ISNULL(QUANTITYTRANSFERRED,0)), ITEMCODE, CLASSIFICATIONCODE, ORGANISATIONCODE,MRSITEMSTATUS,ISNULL(ITEMATTRIBUTES,''),ISNULL(ICOUNTER,0)  FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNo & " ORDER BY ICOUNTER"
																			.ActiveConnection = con
																			.Open
																		end with
																		set dcrs.ActiveConnection = nothing

																		do while not dcrs.eof
																			sMRStatus = ""
																			iQtyReq = Trim(dcrs(0))
																			iQtyAppr = Trim(dcrs(1))
																			iQtyIssued = Trim(dcrs(2))
																			iQtyPur = Trim(dcrs(3))
																			iQtyTransfer = Trim(dcrs(4))
																			iItemCode = Trim(dcrs(5))
																			iClassCode = Trim(dcrs(6))
																			sOrgId = Trim(dcrs(7))
																			sMRStatus = Trim(dcrs(8))
																			sAttribList = Trim(dcrs(9))
																			iCounter = Trim(dcrs(10))

																			iQtyPending = cdbl(iQtyAppr) - (cdbl(iQtyIssued) + (cdbl(iQtyPur) + cdbl(iQtyTransfer)))
																			'sItemName = ItemDisplay(iItemCode,iClassCode)
																			'Response.Write "sAttribList="&sAttribList&"-----------"
																			sAttList = split(sAttribList,":")
																			if trim(sAttribList)<>"" then

																			    IF trim(sAttList(0))  = "" or IsNull(sAttList(0)) then
																				    sOptName = ""
																			    Else

																				    sArrList = split(sAttList(0),"#")
																				    if ubound(sArrList)>0 then
																					    sOptName = FunAttribName(sAttList(0))
																				    else
																					    sQuery = "Select OptionName from INV_M_ItemTypeOptions where OptionValue="& sAttlist(0)
																					    rstemp.open squery,con
																					    if not rstemp.eof then
																						    sOptName = " ["& trim(rstemp(0)) &"]"
																					    end if
																					    rstemp.close
																				    end if
																			    end IF
																			end if 'if trim(sAttribList)<>"" then
																			arrSchTemp = split(GetSchedule(sOrgID,iClassCode,iItemCode,iMRSNo,iCounter),":")
																			sSchTemp = ""


																			if UBound(arrSchTemp) > 0 then
																				sSchTemp = arrSchTemp(0)
																				sSchTempValue = arrSchTemp(1)
																			end if

																			if iQtyAppr = 0 then iQtyAppr = iQtyReq
																			if sMRStatus = "040103" then
																				sMRStatus = "*"
																			else
																				sMRStatus = ""
																			end if
																			iCtr = iCtr + 1

																			with dcrs1
																				.CursorLocation = 3
																				.CursorType = 3
																				.Source = "SELECT ITEMDESCRIPTION FROM VWITEM WHERE ITEMCODE = " & iItemCode & " "
																				.ActiveConnection = con
																				.Open
																			end with
																			if not dcrs1.Eof then
																				sItemName = dcrs1(0)
																			end if
																			dcrs1.close

																			IF sOptName <> "" then  sItemName = sItemName & sOptName

																	%>
																		<tr>
																			<td class="ExcelSerial" align="center" rowspan="2"><%=iCtr%></td>
																			<td class="ExcelDisplayCell" rowspan="2" align="left">
																				<a href="#" class="ExcelDisplayLink" name="lnkA<%=cstr(iClassCode)%>A<%=cstr(iItemCode)%>A<%=sOrgId%>A<%=iCounter%>A<%=sOptName%>" onClick="javascript:DisplayItem(this.name)"><%=sItemName%> &nbsp;<%=sMRStatus%></a>
																			</td>
																			<td class="ExcelDisplayCell" width="10" align="right">
																				<input type="text" name="txtQtyApproved<%=iCtr%>" size="12" value="<%=iQtyAppr%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
																			</td>
																			<td class="ExcelDisplayCell" width="10" align="right">
																				<input type="text" name="txtQtyPending<%=iCtr%>" size="12" value="<%=iQtyPending%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
																			</td>
																			<td class="ExcelDisplayCell" width="91" align="right">
																		<%
																			if sSchTemp = "I" then
																				Response.Write trim(sSchTempValue)
																			elseif sSchTemp = "W" then
																				Response.Write "Within " &sSchTempValue& " Days"
																			elseif sSchTemp = "D" then
																				Response.Write sSchTempValue
																			elseif sSchTemp = "S" then
																				Response.Write sSchTempValue
																		%>
																			<a href="#">
																				<img name="btn:<%=iClassCode%>:<%=iItemCode%>:<%=iCounter%>:<%=sOptName%>" border="0" src="../../assets/images/iTMS%20Icons/Details.gif" width="15" height="15" alt="Schedule Details" onClick="CheckSch(this,'<%=ictr%>')">
																			</a>
																		<%	end if %>
																			</td>
																		</tr>
																		<tr>
																			<td class="ExcelDisplayCell" width="10" align="right">
																				<input type="text" name="txtQtyIssue<%=iCtr%>" size="12" value="<%=iQtyIssued%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
																			</td>
																			<td class="ExcelDisplayCell" width="10" align="right">
																				<input type="text" name="txtQtyPurchase<%=ictr%>" size="10" value="<%=cdbl(iQtyPur)+cdbl(iQtyTransfer)%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right;cursor:hand;FONT-WEIGHT: bold" alt="BreakUp Details" onClick="DisplayDet('<%=iQtyPur%>|<%=iQtyTransfer%>')">
																			</td>
																			<td class="ExcelDisplayCell" align="right">
																				<a href="#">
																					<img name="btn:<%=iClassCode & ":" & iItemCode & ":" & iCounter & ":" & sOptName %>" border="0" src="../../assets/images/iTMS%20Icons/Details.gif" width="15" height="15" alt="Quality Parameters" onClick="CheckQty(this)">
																				</a>
																			</td>
																	<%
																			dcrs.MoveNext
																		loop
																		dcrs.Close
																	%>
																	<Input Type=Hidden name="hCtr" Value="<%=iCtr%>" >
																	<Input Type=Hidden name="sTemp" Value="<%=sTemp%>" >
																	</table>
																</div>
															</td>
														</tr>
														<tr>
															<td class=MiddlePack colspan="3"></td>
														</tr>
													</table>
                                                </td>
											</tr>
										</table>
                                    </div>
								</td>
								<td align="center" class="ClearPixel"></td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td align="left" class="FieldCell" colspan="2" width="100%">* Rejected Item
								</td>
							</tr>

                            <tr>
								<td align="center" class="ClearPixel" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell" align="center">
                                                <input type="button" value="IssuePrint" name="BtnPrint" class="ActionButton" onclick="Print()">
                                                <input type="button" value="Print" name="BtnPrint" class="ActionButton" onclick="PrintNew()">
                                                <input type="button" value="Close" name="BtnClose" class="ActionButton" onclick="window.close()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
	' Function to populate Schedule Type
	Function GetSchedule(sOrgID,iClass,iItem,iMRSNoP,iCounter)
		' Declaration of variables
		Dim dcrs,dcrs1,sReqBy,sReqValue
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			IF iCounter <> 0 then
				.Source = "SELECT REQUIREDBY,REQUIREDVALUE FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNoP & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND ICOUNTER = " & iCounter & " "
			Else
				.Source = "SELECT REQUIREDBY,REQUIREDVALUE FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNoP & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " "
			End IF

			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sReqBy = dcrs(0)
		set sReqValue = dcrs(1)
		if Not dcrs.EOF then
			if not sReqValue = "S" then
				GetSchedule = sReqBy&":"&sReqValue
			else
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT SCHEDULEDON FROM INV_T_MRSITEMSCHEDULES WHERE MRSNUMBER = " & iMRSNoP & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " ORDER BY SCHEDULENO"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing
				set sReqValue = dcrs1(0)
				if Not dcrs1.EOF then
					GetSchedule = sReqBy&":"&sReqValue
				end if
				dcrs1.Close
			end if
		end if
		dcrs.Close

	End Function
%>
