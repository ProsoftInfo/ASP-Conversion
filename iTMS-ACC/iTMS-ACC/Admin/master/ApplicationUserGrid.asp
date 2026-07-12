
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ApplicationUserGrid.asp
	'Module Name				:	Admin(Master)
	'Author Name				:	UMAMAHESWARI S
	'Created On					:	December 06 ,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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

<!--#include file="../../include/DatabaseConnection.asp"-->
<!-- #include File="../../include/populate.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" ID = "OutData"><Root/></script>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<Script>
function appUserGridTrim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function appUserGridField(name) {
	var frm = document.formname;
	return frm && (frm.elements[name] || frm[name]) || null;
}

function Validate() {
	document.formname.hEmpName.value = appUserGridTrim(appUserGridField("TxtEmpName").value);
	document.formname.hLoginID.value = appUserGridTrim(appUserGridField("txtLoginID").value);
	document.formname.submit();
}

function ChkReset() {
	if (appUserGridField("TxtEmpDesig")) {
		appUserGridField("TxtEmpDesig").value = "";
	}
	document.formname.submit();
}

function GotoAction(sPara, UserAccessMode) {
	var count = parseInt(appUserGridField("hCnt").value, 10) || 0;
	var selectedCount = 0;
	var loginId = "";
	var item;
	for (var i = 1; i <= count; i += 1) {
		item = appUserGridField("Chkbox" + i);
		if (item && item.checked) {
			if (appUserGridTrim(item.value) !== "") {
				loginId = item.value;
			}
			selectedCount += 1;
		}
	}
	if (selectedCount > 1) {
		alert("Select any one Employee");
		return false;
	}
	if (selectedCount === 0 && appUserGridTrim(sPara) !== "CRN") {
		alert("Select Employee");
		return false;
	}
	if (appUserGridTrim(sPara) === "CRN") {
		document.formname.action = "UserCreationEntry.asp?UAM=" + UserAccessMode;
		document.formname.submit();
		return true;
	}
	if (appUserGridTrim(sPara) === "EDT") {
		document.formname.action = "AmendUserCreationDetailsEntry.asp?sTemp=" + loginId + ":" + appUserGridField("hUnitID").value;
		document.formname.submit();
		return true;
	}
	return false;
}

function Paginate(nPage) {
	document.formname.hPageSelection.value = nPage;
	document.formname.submit();
}

function ShowRoleDetails(sLoginID, nEmployeeNo) {
	document.formname.action = "AppActivityRole.asp?EmpNo=" + nEmployeeNo;
	document.formname.submit();
}
</Script>
<% 	Dim sSql,sSql1,iCtr,sLoginID,sUnitID,sEmpValue,sEmpName
	Dim iCurrentPage,iTotalPage,lnPage,nSlNo,sDefValue

	Const iPageSize = 15
	Dim objRs,objRs1,objRs2

	Set objRs = Server.CreateObject("ADODB.RecordSet")
	Set objRs1 = Server.CreateObject("ADODB.RecordSet")
	Set objRs2 = Server.CreateObject("ADODB.RecordSet")

	iCurrentPage=CInt(Request.Form("hPageSelection"))
	if iCurrentPage=0 then iCurrentPage=1

	sEmpName = Request("hEmpName")
	sLoginID = Request("hLoginID")
	'sUnitID   = Request("selUnitId")

	if trim(sUnitID) = "" then	sUnitID = Session("organizationcode")

%>
</Head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
<form method="POST" name="formname">
	<Input type=hidden name=hEmpName value="<%=sEmpName%>">
	<Input type=hidden name=hLoginID value="<%=sLoginID%>">
	<Input type=hidden name=hUnitID value="<%=sUnitID%>">
	<INPUT TYPE=HIDDEN NAME="hCallFrom" VALUE="FG">
	<table border="0" width="100%" cellspacing="0" cellpadding="0" >
		<tr>
		<td align="center" class="PageTitle" height="20">
			<p align="center">Application Users
		</td>
		</tr>

		<tr>
		<td valign="top">
			<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
				<tr>
				<td height="1" valign="bottom" class="TabCellEnd">&nbsp;
				</td>
                </tr>

				<tr>
				<TD class=TabBody>
				<!--<td class="TabBodyWithTopLine"><div style="height:130px;">-->
					<table border="0" cellpadding="0" cellspacing="0" >
						<tr>
						<td align="center" colspan="3" class="MiddlePack" height="7" >
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
						</td>
						</tr>

						<tr>
							<td>
							</td>
							<td valign="top" width="100%">
							</td>

							<td>
							</td>
						</tr>

						<tr>
						<td align="center" width="5" class="ClearPixel">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
						</td>
						<td valign="top" width="100%">
						<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
						<tr>
						<td>
						<div>
						<table class="CollapseBand" cellspacing="0" cellpadding="0" >
						<tr>
						<td valign="center">
						<a style="width: 1em; height: 1em;" title="" onclick="return Div_OnClick(idUnprocessed,event);" >
						<img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
						</a>
						</td>
						<td valign="right" class="SubTitle">
						</td>
						</tr>
						</table>

						<table border="0" cellpadding="0" cellspacing="0">
						<tr>
						<td width="100%">
						<div id="idUnprocessed" style="width: 575; display: none">
						<table cellpadding="0" cellspacing="0">

						<!--<tr>
							<td class="FieldCellSub">Unit Name</td>
							<td class="FieldCellSub" colspan="2">
							<select size="1" name="selUnitId" class="FormElem">
								<%'populateUnits%>
							</select>
							</td>
						</tr>-->
						<tr>
							<td class="FieldCellSub">First Name</td>
							<td class="FieldCellSub" colspan="2">
								<input type = text name=TxtEmpName  class="Formelem">
							</td>
						</tr>

						<tr>
							<td class="FieldCellSub">Login ID</td>
							<td class="FieldCellSub" colspan="2">
								<input type = text name=txtLoginID  class="Formelem">
							</td>
						</tr>

						<tr>
							<td class="FieldCellSub"></td>
							<td class="FieldCellSub" >
								<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
							</td>
							<td class="FieldCell" >
								<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()">
							</td>
						</tr>
						</table>
						</div>
						</td>
						</tr>

						<tr>
						<td align="center" class="MiddlePack">
						</td>
						</tr>

						</table>
						</div>
						</td>
						</tr>

						</table>
						</td>
						<td align="center" class="ClearPixel" width="5">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
						</td>
						</tr>

						<tr>
						<td align="center" class="MiddlePack" colspan="3">
						</td>
						</tr>

						<tr>
						<td>
						</td>
						<td valign="top" width="100%">
							<table border="0" cellspacing="1" class="ExcelTable" width="100%">

								<tr>
									<td class="ExcelHeaderCell" align="center" width="10" >S.No.</td>
									<td class="ExcelHeaderCell" align="center" width="10">
										<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15" height="15"></a>
									</td>
									<td class="ExcelHeaderCell" align="center" >Employee ID</td>
									<td class="ExcelHeaderCell" align="center" >Login ID</td>
									<td class="ExcelHeaderCell" align="center" >Employee Name</td>
									<td class="ExcelHeaderCell" align="center" >Phone</td>
									<td class="ExcelHeaderCell" align="center" >Designation</td>
									<td class="ExcelHeaderCell" align="center" >Role</td>
								</tr>

							<%
							sSql = "Select LoginID,InternalUserID,isNull(UserName,''),isNull(MiddleName,''),isNull(LastName,'') ,Designation From DCS_User where InternalUserID > 0"

							if Trim(sEmpName) <> "" then
								sSql = sSql & " and UserName like '%"& sEmpName &"%'"
							End if

							if Trim(sLoginID) <> "" then
								sSql = sSql & " and LoginID like '%"& sLoginID &"%'"
							End if

							With objRs
								.ActiveConnection = Con
								.CursorLocation = 3
								.CursorType = 3
								.Source = sSql
								.Open
							End With

							Set objRs.ActiveConnection = Nothing
							iCtr = 1
							nSlNo = 1

							If not objRs.EOF then
								objRs.PageSize = iPageSize
								objRs.AbsolutePage = iCurrentPage
								iTotalPage = objRs.PageCount
							End if 'If not objRs.EOF then

								sEmpValue = ""

								Do while not objRs.EOF and nSlNo <= objRs.PageSize
									%>
									<tr>
										<td class="ExcelSerial" align="center" ><%=nSlNo%></td>
										<td class="ExcelDisplayCell" align="center" width="10">
											<input type="checkbox" name="Chkbox<%=nSlNo%>" value="<%=objrs(0)%>">
										</td>
										<td class="ExcelDisplayCell" align="center" ><%=objRs(1)%></td>
										<td class="ExcelDisplayCell" align="Left"><%=objRs(0)%></td>
										<td class="ExcelDisplayCell" align="Left"><%Response.Write objRs(2) & " " & objRs(3) & " " & objRs(4)%></td>
										<td class="ExcelDisplayCell" align="left"></td>
										<td class="ExcelDisplayCell" align="Left" ><%=objRs(5)%></td>
										<td class="ExcelDisplayCell" align="center">
											<p align="center"></p><img border="0" style="cursor: pointer" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="top" width="11" height="11" alt="Role Details" onclick="ShowRoleDetails('<%=objrs(0)%>','<%=objrs(1)%>')" ></td>
										</td>
									</tr>

									<%
										nSlNo = nSlNo + 1
										iCtr = iCtr + 1
										objRs.MoveNext
								Loop
							objRs.Close
							'Response.Write "<p>iCtr="&iCtr
								%>
								<input type="hidden" name="hCnt" value="<%=iCtr-1%>">

							</table>
							</div>
							</td>
								<td>
								</td>
                            </tr>

							<tr>
								<td colspan="3" class="MiddlePack">
								</td>
							</tr>

							<tr>
							<td align="center" width="5" class="ClearPixel">
							</td>
							<td valign="top" align="right">
							<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
							<!--<input type=hidden name="hCnt" value=<%'=iCtr -1  %>>-->
							<input type=hidden name="hPageSelection" value="0">

							<%	If iTotalPage >= 2 Then
							if iCurrentPage = 1 then
							%>
							<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
							<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
							<%		else%>
							<input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
							<input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
							<%		end if	%>
							<SELECT class="FormElem" onChange="Paginate(this.options[this.selectedIndex].value)" id=select1 name=select1>
							<%
							For lnPage = 1 To iTotalPage
							If lnPage = iCurrentPage Then
							%>
							<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
							<%		else	%>
							<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
							<%		end if
							next
							%>
							</SELECT>
							<%
							if iCurrentPage = iTotalPage then
							%>
							<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
							<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

							<%		else	%>
							<input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
							<input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotalPage%>')" id=button8 name=button8>
							<%		end if
							End If
							%>
							</td>
							<td align="center" class="ClearPixel" width="5">
							<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							</tr>

							<tr>
							<td>
								<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							</td>
							<td valign="top">

							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td valign="middle" class="ActionCell">
									<p align="center">
										<!--<select size="1" name="Choice" class="FormElem">
											<option Value="SEL">Select </option>
											<option Value="CRN">Create</option>
											<option Value="EDT">Edit</option>
										</select>-->
										<Input type="button" value="Create Internal User" name="ButOpt" class="ActionButtonX" tabindex="3" onclick="GotoAction('CRN','I')" >
										<Input type="button" value="Create External User" name="ButOpt" class="ActionButtonX" tabindex="3" onclick="GotoAction('CRN','E')" >
										<Input type="button" value="Edit" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction('EDT','N')" >
										<!--<Input type="button" value="Proceed" name="ButOpt" class="ActionButton" tabindex="3" onclick="GotoAction()" >-->
									</td>
								</tr>
							</table>

							</td>
								<td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td colspan="3" class="BottomPack">
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
Function populateUnits()
	' Declaration of variables
	Dim dcrs,sUnitID,sUnitName
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM VWITEMLIST) ORDER BY ORGANIZATIONUNITID"
		.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	set sUnitID = dcrs(0)
	set sUnitName = dcrs(1)
	If not dcrs.EOF then
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
	end if
	dcrs.Close

End Function
%>
