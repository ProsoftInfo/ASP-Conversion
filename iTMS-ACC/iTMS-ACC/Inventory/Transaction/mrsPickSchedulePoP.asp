<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsPickSchedulePoP.asp
	'Module Name				:	Inventory (Direct Issue)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Feb 07,2013
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None

%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Pick Schedule</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="TempData"><Root></Root></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsPickSchedulePop.js"></SCRIPT>
</head>
<%
Dim rsTemp,objrs
	Dim sOrgCode,sQuery,sItemName,iItemCode,iClassCode,AttributeList,sArrTemp
	sOrgCode = Request.QueryString("OrgID")
	set rsTemp = Server.CreateObject("ADODB.Recordset")
	set objrs = Server.CreateObject("ADODB.Recordset")
	sArrTemp= split(Request.QueryString("sTemp"),":")
	iItemCode = sArrTemp(0)
	iClassCode=sArrTemp(1)
	AttributeList =  Request.QueryString("AttributeList")
	AttributeList = replace(replace(AttributeList,"$","#"),"@",":")
	
	sQuery = "Select ItemDescription from INV_M_ItemMaster where ItemCode = "& iItemCode &" and ClassificationCode = " & iClassCode
	rsTemp.open sQuery,con
	if not rsTemp.eof then
	    sItemName = trim(rsTemp(0))
	end if
	rsTemp.close
%>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init('<%=iItemCode%>','<%=iClassCode%>','<%=AttributeList%>')">
<form method="POST" name="formname">
<input type="hidden" name="hUnit" value="<%=sOrgCode%>">
<input type="hidden" name="hItemCode" value="<%=iItemCode%>">
<input type="hidden" name="hClassCode" value="<%=iClassCode%>">
<input type="hidden" name="hAttributeList" value="<%=AttributeList%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Pick Schedule
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
								<td align="center" class="ClearPixel" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td width="5"></td>
								<td>
								    <table width="100%">
								        <tr>
								            <td class="fieldcell">
								                Item Name
								            </td>
								            <td class="fieldcell">
								                <span id="SpanItemName" class="DataOnly"><%=sItemName%></span>
								            </td>
								        </tr>
								        <tr>
								            <td class="fieldcell">
								                Issued Quantity
								            </td>
								            <td class="fieldcell">
								                <span id="spanIssQty" class="DataOnly">0</span>
								            </td>
								        </tr>
								        <tr>
								            <td class="fieldcell">
								                Schedule Date
								            </td>
								            <td class="fieldcell">
								                <%
								                    InsertDatePicker("ctlScheduleDate")
								                %>
								            </td>
								        </tr>
								        <tr>
								            <td class="fieldcell">
								                Schedule Quantity
								            </td>
								            <td class="fieldcell">
								                <input type="text" name="txtSchQty" class="FormElem" value="0" style="width:80px;text-align:right;">&nbsp;&nbsp;<input type="button" name="btnAddSchedule" value="Add Schedule" class="AddButtonX" onclick="AddSchedule()">
								            </td>
								        </tr>
								    </table>
								</td>
								<td width="5"></td>
							</tr>
							<tr>
								<td width="5"></td>
								<td>
								    <table width="100%" id="tblSchedule" class="ExcelTable">
								        <tr>
								            <td class="ExcelHeaderCell" align="center">S.No.</td>
								            <td class="ExcelHeaderCell" align="center">
								                Schedule Date
								            </td>
								            <td class="ExcelHeaderCell" align="center">
								                Schedule Quantity
								            </td>
								        </tr>
								    </table>
								</td>
								<td width="5"></td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="ActionCell" >
									<input type=button name="btnDone" value="Done" class="ActionButtonX" onClick="FinalSubmit()">
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
