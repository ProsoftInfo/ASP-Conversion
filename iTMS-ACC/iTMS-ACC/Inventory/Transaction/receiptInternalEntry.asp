<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	receiptInternalEntry.asp
	'Module Name				:	Inventory (Receipt Creation)
	'Author Name				:	KUMAR K A
	'Created On					:
	'Modified By				:
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
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!--#include file="../../include/CommonFunctions.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
	Dim iCtr,arrTemp,sTemp,sSource,sOrgID,sDept,iItem, iClass , sReceiptType, sReceiptName
	Dim dcrs, sSalesEli, sPurEli, sManEli, sEligible, arrUoM, sUoMCode, sUoMDesc, sCheck
	Dim sTempMonYr, sMonYr, arrFin, sFinFrom, sFinTo,sType,sItmType,sQuery,Arr1,sPassData
	Dim dtCurrDate,sAttributeList,sFinPeriod,Arr,sMinDate,sMaxDate,sAutoInternalRcptAccount
	Dim iRcptNo,sMode
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	sOrgID = Session("organizationcode")

	iRcptNo = Request("RcptNo")
	if Trim(iRcptNo)<>"" then
	    sMode = "E"
	else
	    sMode = "N"
	end if

	sFinPeriod = session("FinPeriod")
	Arr = split(sFinPeriod,":")
	sMinDate = "01/04/"& Arr(0)
	sMaxDate = "31/03/"& Arr(1)


	iItem = trim(Request.Form("hItmCode"))

	sSource = "N"
	sTemp = trim(Request.Form("hSelectedValue"))


	if trim(sOrgID) = "" then
		If iSAApplicationPop <> "" then
			sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY OUDEFINITIONID"
		Else
			sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
		End If

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			sOrgID = dcrs(0)
		end if
		dcrs.Close
	end if

	sTempMonYr = mid(FormatDate(date),4,2)
	sMonYr = sTempMonYr&Year(date())

	arrFin = split(Session("FinPeriod"),":")
	sFinFrom = "01/04/"& arrFin(0)
	sFinTo = "31/03/"& arrFin(1)

	if DateDiff("D",FormatDate(sFinTo),FormatDate(date)) > 0 then
		dtCurrDate = sFinTo
	else
		dtCurrDate = FormatDate(date)
	end if

	sQuery = "Select IsNull(AutoInternalRcptAccounting,'N') from INV_M_ApplicationSetup"
	dcrs.open sQuery,con
	if not dcrs.eof then
	    sAutoInternalRcptAccount =dcrs(0)
	else
	    sAutoInternalRcptAccount = "N"
	end if
	dcrs.close


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Receipt Creation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!--<ITEM CLACODE="<%=trim(iClass)%>" ITMCODE="<%=trim(iItem)%>" QTY="" MRSNO="<%="N"%>" ISSNO="<%="N"%>" />-->
<script type="application/xml" data-itms-xml-island="1" id="OutData2">
	<ROOT DEPT="" SOURCE="<%=sSource%>" ORGCODE="<%=sOrgID%>" STYPE="" ITEMTYPE="" PACKNUM="" SRCREFTYPE="" SRCREFNO="" RCPTNUMBERINV="" APPREFTYPE="" APPREFNO="" APPREFDATE="">
	</ROOT>
</script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemData">
<root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="NewData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" id="StoreData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="RefXML"><Root/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/newReceipt.js"></script>
<script LANGUAGE=javascript SRC="../scripts/receiptInternalEntry.js"></script>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>

<BODY leftMargin=0 topMargin=0 onLoad="EditInit('<%=iRcptNo%>');setdate()">
<form method="POST" name="formname">
<input type=hidden name="hSelectedValue" value="">
<input type=hidden name="hItemNames" value="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hIType" value="<%=sItmType%>">
<input type=hidden name="hOrgCode" value="<%=sOrgID%>">
<input type=hidden name="hItmCode" value="<%=iItem%>">
<input type=hidden name="hClassCode" value="<%=iClass%>">
<INPUT TYPE=HIDDEN NAME="hStoresUom" VALUE="<%=sUoMDesc%>">
<INPUT TYPE=HIDDEN NAME="hReceiptType" VALUE="<%=sReceiptType%>">
<INPUT TYPE=HIDDEN NAME="hDept" VALUE="<%=sDept%>">
<INPUT TYPE=HIDDEN NAME="hType" VALUE="<%=sType%>">
<input type=hidden name="hCurrDate" value="<%=dtCurrDate%>">
<input type=hidden name="hAttributeList" value="<%=sAttributeList%>">
<input type="hidden" name="hItemType" value="">
<input type="hidden" name="hCtr" value="0">
<input type="hidden" name="hMinDate" value="<%=sMinDate%>">
<input type="hidden" name="hMaxDate" value="<%=sMaxDate%>">
<input type="hidden" name="hAutoAccount" value="<%=sAutoInternalRcptAccount%>" />
<input type="hidden" name="hMode" value="<%=sMode%>" />
<input type="hidden" name="hRcptNo" value="<%=iRcptNo%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Internal Receipts
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
								<td valign="top" class="FieldCell">
                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                       <!-- <tr>
                                            <td class="FieldCell">Select Unit</td>
                                            <td class="FieldCellSub" colspan=3>
												<select size="1" name="selUnit" class="FormElem" onChange="resetAll('U')">
													<!--option value="select">Select</option-->
													<%	'Calling the Function which populates Organization Unit list
													'	populateUnitSelected(sOrgID)
													%>
												<!--</select>
											</td>
                                        </tr>-->
                                        <!--<tr>
                                            <td class="FieldCell">Usage</td>
                                            <td class="FieldCellSub" colspan=3>
												<select size="1" name="selDepart" class="FormElem" onChange="resetAll('R');DoChanges(this)" <%if sPassData <> "" then Response.Write " Disabled " %> >
													<option value="select">Select</option>
													<%	'Calling the Function which populates Department list
														'populateDepartment(sDept)
													%>
												</select>
                                            </td>
                                        </tr>-->
                                       <tr>
                                            <td class="FieldCell">Received From</td>
                                            <td class="FieldCellSub">
												<select size="1" name="selDepart" class="FormElem" onChange="resetAll('R');DoChanges(this)" <%if sPassData <> "" then Response.Write " Disabled " %> >
													<option value="select">Select</option>
													<%	'Calling the Function which populates Department list
														ReceivedFrom(sOrgID)
													%>
												</select>&nbsp;
												<span id="spanWCName" class="dataonly"></span>
												<input type="hidden" name="hWCCode" value="">
                                            </td>
                                            <td class="FieldCellSub">
                                                Received On
                                            </td>
                                            <td class="FieldCellSub">
                                                <%
                                                    InsertDatePicker("ctlRcvdOn")
                                                %>
                                            </td>
										<!--<tr>
										    <td class="FieldCell">Select Type</td>
										    <td class="FieldCellSub" colspan=3>
												<select size="1" name="selAddType" class="FormElem" onChange="DoDisable(this)"  <%if sDept <> "PRD" then Response.Write " Disabled " %> >
													<option value="N" <% If sType = "N" Then Response.Write "Selected" %> >Select</option>
													<option value="W" <% If sType = "W" Then Response.Write "Selected" %>>Work Center</option>
													<option value="P" <% If sType = "P" Then Response.Write "Selected" %>>Packing</option>
													<option value="M" <% If sType = "M" Then Response.Write "Selected" %>>Mixing</option>
													<option value="T" <% If sType = "T" Then Response.Write "Selected" %>>Waste</option>
											    </select>
											</td>
										</tr>-->
										<!--<tr>
										    <td class="FieldCell">Item Type</td>
										    <td class="FieldCellSub">
											    <select size="1" name="selItmType" class="FormElem" onChange="resetAll('T')">
													<option value="select">Select</option>
													<%	'Calling the Function which populates the Item Type list
												'		populateItemTypeSelected sItmType
													%>
												</select>
											</td>
										</tr>-->

                                        <!--tr>
                                            <td class="FieldCell">Source Reference</td>
                                            <td class="FieldCellSub" colspan="3">
												<select size="1" name="selSrc" class="FormElem" onChange="CheckType(this)">
													<option value="select">Select</option>
													<option value="M">MR / Direct Issue</option>
													<option value="N">None</option>
												</select>
                                            </td>
                                        </tr-->

											<!--<tr>
											   <td class="FieldCell">Select Item</td>
											   <td class="FieldCellSub" colspan=4>
											   <% If iItem <> "" Then %>
											   <span class="DataOnly" ><%=ItemDisplay(iItem,iClass)%>&nbsp;</span>
											   <% End If %>
											   <a onClick="Search()" href="#">
													<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif"  align="top" width="11" height="11" alt="Select Item">
												</a>
												</td>
											 </tr>-->


											<!--<tr>
												<td class=FieldCell> Receipt Numbering</td>
												<td class='FieldCellSub'>
													<Span Class="DataOnly"><%=sReceiptName%>&nbsp;</Span>
                                                </td>
											</tr>-->

											<!--<tr>
												<td class=FieldCell valign="top"> Storage Location</td>
												<td class='FieldCellSub' colspan="4">
													<select size="5" name="selStorage" class="FormElem" multiple onBlur="GetStockDet(this,'<%=sCheck%>')">
													<%	'Calling the Function which populates the Store list
														'populateStores sOrgID, sEligible
													%>
													</select>
                                                </td>
											</tr>-->

											<!--<tr>
												<td class=FieldCell >Source Reference</td>
												<td class='FieldCellSub' colspan="4">
												<%	if sDept <> "PRD" then %>
													<input type="radio" value="S" name="radSource" class="FormElem" onclick="SelectReference('S')"> Sales Order&nbsp;
													<input type="radio" value="P" name="radSource" class="FormElem" onclick="SelectReference('P')"> Purchase Order&nbsp;
													<input type="radio" value="R" name="radSource" class="FormElem" onclick="SelectReference('R')"> Production Order&nbsp;
													<input type="radio" value="M" name="radSource" class="FormElem" onclick="SelectReference('M')"> Mixing&nbsp;
													<input type="radio" value="I" name="radSource" class="FormElem" onclick="SelectReference('I')"> Issue&nbsp;
													<input type="radio" value="N" name="radSource" class="FormElem" onclick="SelectReference('N')" CHECKED> None
                                                <%	elseif sDept = "PRD" then %>
													<input type="radio" value="S" name="radSource" class="FormElem" onclick="SelectReference('S')" DISABLED> Sales Order&nbsp;
													<input type="radio" value="P" name="radSource" class="FormElem" onclick="SelectReference('P')" DISABLED> Purchase Order&nbsp;
													<input type="radio" value="R" name="radSource" class="FormElem" onclick="SelectReference('R')" CHECKED> Production Order&nbsp;
													<input type="radio" value="M" name="radSource" class="FormElem" onclick="SelectReference('M')" > Mixing&nbsp;
													<input type="radio" value="I" name="radSource" class="FormElem" onclick="SelectReference('I')"> Issue&nbsp;
													<input type="radio" value="N" name="radSource" class="FormElem" onclick="SelectReference('N')" DISABLED> None
                                                <%	end if %>
                                                </td>
											</tr>-->
											<tr>
											    <td class="FieldCell">Reference Name</td>
											    <td class="FieldCellSub">
											       <select name="selRefName" class="FormElem" Onchange="GetDetails()">
													<%
													    RefTypePop 4,4
													%>
													</select>
													<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click Here to Edit Usage Information" width="11" height="11" onClick="GetDetails()"></a>
													&nbsp;<span id="RefNoDate" class="dataonly"></span>
													<input type="hidden" name="hRefNo" value="">
													<input type="hidden" name="hRefDate" value="">
											    </td>
											    <td class="FieldCellSub">
                                                    Created By
                                                </td>
                                                <td class="FieldCellSub">
                                                    <span id="spancreatedby" class="DataOnly"><%=session("username")%></span>
                                                </td>
											</tr>



											<!--<tr>
												<td class=FieldCell valign="top"></td>
												<td class='FieldCellSub' colspan="4">
													<input type="text" name="txtSource" size="20" maxlength=30 class="FormElem">
													&nbsp;<span id="spanSource" class="dataonly"></span>
                                                </td>
											</tr>-->


                                    </table>
								</td>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>

							<tr>

								<td valign="top" colspan="3">
									<table border="0" cellspacing="1" width="100%" >

										<tr>
											<td align="center"></td>
											<td>
												<DIV class="frmBody" id="frm1" style="width: 750; height:340;">
													<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width=100%>
														<tr>
															<td class="ExcelHeaderCell" align="center" rowspan="2" width="10">S.No.</td>
															<td class="ExcelHeaderCell" align="center" rowspan="2">Item Name</td>
															<td class="ExcelHeaderCell" align="center" rowspan="2">Storage Location</td>
															<td class="ExcelHeaderCell" align="center" colspan="2">Stock Details</td>
															<td class="ExcelHeaderCell" align="center" rowspan="2" width="75">Lot & Serial</td>
															<td class="ExcelHeaderCell" align="center" rowspan="2" width="75">By Product?</td>
														</tr>
														<tr>
															<td class="ExcelHeaderCell" align="center">Net. Quantity</td>
															<td class="ExcelHeaderCell" align="center">Unit Rate</td>
														</tr>
													</table>
													<input type="button" name="btnAddItem" class="AddButtonX" value="Add Item" onclick="AddItem()">
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
														<td valign="middle" class="ActionCell" align="center">
														<% 'Response.Write sFinFrom & "," & sFinTo %>
															<input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit('<%=sFinFrom%>','<%=sFinTo%>','<%=dtCurrDate%>')">
															<input type="reset" value="Reset" name="B1" class="ActionButton">
															<input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('MATERIALRECEIPTS.ASP?RCPT=A')">
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
	' Function to populate Department
	Function populateDepartment(Dept)
		' Declaration of variables
		Dim dcrs,sDepartCode,sDepartDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DEPTNO,DEPTNAME FROM INV_M_DEPARTMENT ORDER BY DEPTNO"
			.Source = "SELECT ISSUEDFORCODE,ISSUEDFORDESCRIPTION FROM INV_M_ISSUEDFOR ORDER BY ISSUEDFORDESCRIPTION"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sDepartCode = dcrs(0)
		set sDepartDesc = dcrs(1)

		Do While Not dcrs.EOF
			If Dept = trim(sDepartCode) Then
				Response.Write("<OPTION VALUE="""&trim(sDepartCode)&""" Selected>"&trim(sDepartDesc)&" </OPTION>" &vbcrlf)
			Else
				Response.Write("<OPTION VALUE="""&trim(sDepartCode)&""">"&trim(sDepartDesc)&"</OPTION>" &vbcrlf)
			End If
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
<%
	' Function to populate Store
	Function populateStores(sOrgID,sEligible)
		' Declaration of variables
		Dim dcrs,dcrs1,sLoc,sBin,sBinName,sLocName,sLocCode,imaxLoc,sSql
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT LOCATIONNUMBER,LOCATIONNAME,APPLICABLEFOR FROM Inv_M_Storage WHERE OUDEFINITIONID = " & Pack(sOrgID) & " AND APPLICABLEFOR IN ('IN') ORDER BY 1"
			'Response.Write dcrs.source
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then
			Do While Not dcrs.EOF
				sLoc = trim(dcrs(0))
				sLocName = trim(dcrs(1))
				sLocCode = trim(dcrs(2))

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT BINNUMBER,BINNAME,BINCODE FROM Inv_M_StoreBinDetails WHERE LOCATIONNUMBER = " & sLoc & " ORDER BY BINNUMBER"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					do while not dcrs1.EOF
						Response.Write("<OPTION VALUE="""&sLoc&"-"&trim(dcrs1(0))&"-"&sLocCode&""">"&sLocName&" -- "&trim(dcrs1(1))&"</OPTION>" &vbcrlf)
					dcrs1.MoveNext
					loop
				else
					Response.Write("<OPTION VALUE="""&sLoc&"-NULL-"&sLocCode&""">"&sLocName&"</OPTION>" &vbcrlf)
				end if
				dcrs1.Close

			dcrs.MoveNext
			Loop
		else
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'PUR','PURCHASE', " &_
				" 'PU','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'IOO','INSPECTION-OUTORDER', " &_
				" 'OI','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'POI','INSPECTION-PREORDER', " &_
				" 'POI','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'SAL','SALES', " &_
				" 'SA','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'INV','INVENTORY', " &_
				" 'IN','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'PI','INSPECTION PROCESS', " &_
				" 'PI','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_ORGSTORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'PSI','POST SALE', " &_
				" 'PSI','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				imaxLoc = trim(dcrs1(0))
			end if
			dcrs1.Close

			sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
				"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
				"(" & Pack(sOrgID) & "," & imaxLoc & ",'MAN','MANUFACTURING', " &_
				" 'MA','1','0','1000',0)"
			'Response.Write sSql & "<BR>"
			con.Execute sSql

			Server.Execute ("XMLStorageDefault.asp")

			populateStores sOrgID,sEligible

		end if
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
    ' Function to Display ReceivedFrom
	Function ReceivedFrom(sOrgCode)
		Dim sQuery,rsTemp,objrs
        set rsTemp = Server.CreateObject("ADODB.Recordset")
        set objrs = Server.CreateObject("ADODB.Recordset")

	    sQuery = "Select DeptShortName,DepartmentName from APP_M_Departments"
	    rsTemp.Open sQuery,con
	    if not rsTemp.Eof then
		    do while not rsTemp.EOF
			    Response.Write "<option value='"& trim(rsTemp(0)) &"'>"&rsTemp(1)&"</option>"
			    rsTemp.MoveNext
		    loop
		end if
	    rsTemp.Close
	    'Response.Write "<option value='Party'>Party</option>"
	    Response.Write "<option value='Unit'>Other Unit</option>"
		    sQuery = "Select OuDefinitionID,OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where Len(OuDefinitionID)>4 and OuDefinitionID not in('"& sOrgCode &"')"
		    objrs.Open sQuery,con
		    if not objrs.EOF then
		        do while not objrs.EOF
		            Response.Write "<option value='"& trim(objrs(0)) &"'>&nbsp;&nbsp;&nbsp;"&trim(objrs(1))&"</option>"
		            objrs.MoveNext
		        loop
		    end if
		    objrs.Close

	End Function
%>
