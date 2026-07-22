<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	IssueUsageSelPop.asp
	'Module Name				:	Inventory (Direct Issue)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	MARCH 15,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None

%>
<%
	Dim rsTemp,objrs
	Dim sOrgCode,sQuery
	sOrgCode = Request.QueryString("OrgID")
	set rsTemp = Server.CreateObject("ADODB.Recordset")
	set objrs = Server.CreateObject("ADODB.Recordset")
%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>MR Issue - Item Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="RefType"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="RefData"><Root Done="N"/></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/issueUsageSelPop.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init()">
<form method="POST" name="formname">
<input type=hidden name="hUnit" value="<%=sOrgCode%>">
<input type=Hidden name="hUsage" value="">
<input type=hidden name="hSupplier" value="">
<input type=hidden name="hSupplierName" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Material Issue
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>

		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td width="10px"></td>
					<TD class=TabBodywithtopline>

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
							</tr>
							<tr>
								<td width="10">
								</td>
								<td>
									<table border=0 cellspacing=1 cellpadding=0 width="100%" class="ExcelTable">
										<tr>
											<td class="ExcelHeaderCell" align="Center"> Usage</td>
										</tr>
										<tr>

											<td class="FieldCellSub" valign="top" align="Center" >
												<select size="12" name="selUsage" class="FormElem">
													<!--<option value="select">Select</option>-->
												<%	'Calling the Function which populates Usage List
													populateUsage
												%>
												</select>
											</td>
										</tr>
									</table>
								</td>
								<td width=5></td>
								<td valign="top">
									<table border=0 cellspacing=0 cellpadding=0 width="100%" class="ExcelTable" >
											<tr>
												<td class="ExcelHeaderCell" align="Center"><b> Issue To</b></td>
											</tr>
											<tr>
												<td class="FieldCellSub" valign="top">
													<select size="12" name="selIssueFor" class="FormElem" onChange="popParty()">
													<%
														populateIssueToSel(sOrgCode)
													%>
													</select>
												</td>
											</tr>
											<!--<tr>
												<td class="FieldCellSub">Ref. Name</td>
												<td class="FieldCellSub" valign="top">
													<select size="1" name="SelRefType" class="FormElem" >
														<option value="Select">Select</option>
														<%
														    RefTypePop 2,4
														%>
													</select>
												</td>
											</tr>
											<tr>
												<td class="FieldCellSub"><input type="radio" value="WI"  name="radDoc" class=FormElem onClick="SelDoc()">With Doc.</td>
												<td class="FieldCellSub"><input type="radio" value="WO" name="radDoc" class=FormElem onClick="SelDoc()" checked>Without Doc.</td>
											</tr>
											<tr>
												<td class="FieldCellSub">Ref. No</td>
												<td class="FieldCellSub"> <textarea name="txtRef" class="FormElemRead" READONLY cols=20 rows=4></textarea></td>
											</tr>-->
									</table>
								</td>
								<td width=5></td>
							</tr>
							<tr>
								<td Height="26" colspan=5>
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel" colspan=5>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="5" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="5" class="ActionCell" >
									<input type=button name="btnProceed" value="Proceed" class="ActionButtonX" onClick="FinalSubmit()">
									<input type=button name="btnReset" value="Reset" class="ActionButtonX">
								</td>
							</tr>


                        </table>
					</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

<%
	' Function to populate Usage
	Function populateUsage()
		' Declaration of variables
		Dim dcrs,sUsageCode,sUsageDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISSUEDFORCODE,ISSUEDFORDESCRIPTION FROM INV_M_ISSUEDFOR ORDER BY ISSUEDFORCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUsageCode = dcrs(0)
		set sUsageDesc = dcrs(1)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUsageCode)&""">"&trim(sUsageDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
<%
	' Function to populate the Cost Center list
	Function populateCostCenter()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT COSTCENTERHEAD,CCACCOUNTDESCRIPTION FROM VWORGCOSTCENTER WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND USEABLE = 1 ORDER BY COSTCENTERHEAD"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>

<%
	' Function to populate the Account Head list
	Function populateAccountHead()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT DISTINCT ACCOUNTHEAD,ACCOUNTDESCRIPTION,ACCOUNTHEADCODE FROM VWORGGLHEADS WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND ACCOUNTHEAD IN (SELECT ACCOUNTHEAD FROM ACC_R_GLACCAPPLICATIONS WHERE AVAILABLEINAPPLN IN (4,5,6) AND OUDEFINITIONID = " & Pack(sUnit) & ") ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(2)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function

	Function Issue()
		MsgBox "ok"
	End Function
%>
<%
Function populateIssueToSel(sOrgCode)
	sQuery = "Select DepartmentNo,DepartMentName from MS_DEPARTMENTS"
	rsTemp.Open sQuery,con
	if not rsTemp.Eof then
		do while not rsTemp.EOF
			Response.Write "<option value='"& trim(rsTemp(0)) &"'>"&rsTemp(1)&"</option>"
			'	if strcomp(trim(rsTemp(1)),"Work Center")=0 then
			'		Response.Write "<option value='M'>&nbsp;&nbsp;&nbsp;Mixing</option>"
			'		sQuery = "Select WorkCenterCode,WorkCenterName from PRD_M_WORKCENTER  where OrganisationCode = '"& sOrgCode &"'"
			'		objrs.Open sQuery,con
			'		if not objrs.EOF then
			'			do while not objrs.EOF
			'				Response.Write "<option value='"&trim(objrs(0))&"'>&nbsp;&nbsp;&nbsp;"&objrs(1)&"</option>"
			'				objrs.MoveNext
			'			loop
			'		end if
			'		objrs.Close
			'		Response.Write "<option value='P'>&nbsp;&nbsp;&nbsp;Packing</option>"
			'	end if
			rsTemp.MoveNext
		loop
		Response.Write "<option value='Party'>Party</option>"
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
	end if
	rsTemp.Close
End Function
%>
