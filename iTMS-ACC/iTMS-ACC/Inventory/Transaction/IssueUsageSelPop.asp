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
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Dim objTemp
set objTemp = window.dialogArguments
'***********************************************************
Function Init()
Dim Root,iCnt,subNode,RefNode
	set Root = objTemp.documentElement
	if Root.hasChildNodes() then
		Root.setAttribute "Done","N"
		for iCnt = 0 to cint(document.formname.selUsage.length)-1
			if trim(document.formname.selUsage(iCnt).value) = trim(Root.getAttribute("Usage")) then
				document.formname.selUsage.selectedIndex = iCnt
			end if
		next
		For each subNode in Root.childNodes
			if strcomp(subNode.nodeName,"Ref")=0 then
				set RefNode = subNode
			end if
		Next
		document.formname.selIssueFor.value = RefNode.getAttribute("Issue")
		popParty()
	end if
End Function
'=====================================================================
Function FinalSubmit()
Dim Root,RefNode,PartyNode
Dim sTemp,ObjValue

    if document.formname.selUsage.selectedIndex = -1 then
        document.formname.selUsage.focus
        alert("Select Usage")
        exit function
    end if

    if document.formname.selIssueFor.selectedIndex = -1 then
        document.formname.selIssueFor.focus
        alert("Select Issue To")
        exit function
    end if

	ObjValue = document.formname.selUsage(document.formname.selUsage.selectedIndex).value

	set Root = objTemp.documentElement
		Root.setAttribute "Usage",ObjValue
		Root.setAttribute "UsageName",document.formname.selUsage(document.formname.selUsage.selectedIndex).text
		Root.setAttribute "Done","Y"
		Root.setAttribute "IssueTo",document.formname.selIssueFor(document.formname.selIssueFor.selectedIndex).value
		Root.setAttribute "IssueToName",document.formname.selIssueFor(document.formname.selIssueFor.selectedIndex).text

	For each node in Root.childNodes
		if strcomp(node.nodeName,"Party")=0 then
			Root.removeChild(node)
		elseif strcomp(node.nodeName,"Ref")=0 then
			Root.removeChild(node)
		end if
	next

	set PartyNode = objTemp.createElement("Party")
		sTemp = split(document.formname.hSupplierName.value&":"&document.formname.hSupplier.value,":")

	if (sTemp(0)="" and sTemp(1)="" ) or (isNull(sTemp(0)) and isNull(sTemp(1))) then
		PartyNode.setAttribute "Name",""
		PartyNode.setAttribute "Code",""
	else
		PartyNode.setAttribute "Name",sTemp(0)
		PartyNode.setAttribute "Code",sTemp(1)
	end if
	Root.appendChild PartyNode

	set RefNode = objTemp.createElement("Ref")
	RefNode.setAttribute "Issue",document.formname.selIssueFor(document.formname.selIssueFor.selectedIndex).value
	RefNode.setAttribute "IssName",document.formname.selIssueFor(document.formname.selIssueFor.selectedIndex).text
	Root.appendChild RefNode
	window.close
End Function
'=====================================================
Function window_onunload()
window.returnvalue = objTemp.documentElement
End Function
'=========================================================
Function popParty()
Dim OutValue,ObjValue,IssVal
	IssVal = document.formname.selIssueFor(document.formname.selIssueFor.selectedIndex).value
	if trim(document.formname.selUsage.selectedIndex)="-1" then
		alert("Select Usage")
		document.formname.selUsage.focus
		document.formname.selIssueFor.value = "A"
		exit function
	end if
	if lcase(trim(IssVal))=lcase("Party") then
		document.formname.hUsage.value = document.formname.selUsage(document.formname.selUsage.selectedIndex).value
		ObjValue = document.formname.hUsage.value

		sOrgID = document.formname.hUnit.value
	    set	OutValue = showModalDialog("../../Common/PartySelection.asp?orgID="&sOrgID,PartyData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	    sQuery = OutValue.getAttribute("PassQuery")
	    if OutValue.getAttribute("Action")="CLOSE" then exit function

		while OutValue.getAttribute("Action")<>"Done"
		set	OutValue = showModalDialog("../../Common/PartySelection.asp?"&sQuery,PartyData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		    sQuery = OutValue.getAttribute("PassQuery")
	        if OutValue.getAttribute("Action")="CLOSE" then exit function
		wend
		if OutValue.hasChildNodes() then
		    For each ndChild in OutValue.childNodes
		        document.formname.hSupplierName.value = ndChild.getAttribute("RetField0")
		        document.formname.hSupplier.value = ndChild.getAttribute("RetField1")
		    Next
		end if
	end if ' if lcase(trim(IssVal))=lcase("Party") then
End Function
'********************************************************

</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
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
													<select size="12" name="selIssueFor" class="FormElem" onchange("issue") onChange="popParty()">
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