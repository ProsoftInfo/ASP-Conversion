<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MasClassificationNameEntry.asp
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 16, 2002
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	MasClassificationNameInsert.asp
	'Procedures/Functions Used	:	populateItemType
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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Classification</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/masClassificationCreate.js"></SCRIPT>
<script type="application/xml" id="TempData" data-itms-xml-island="1"><Root></Root></script>
<script>
function trimValue(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function CatAddNew() {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog("MasCategoryAddEditPop.asp", "", "dialogWidth:500px;dialogHeight:300px;Status:No;", popCategorty);
	} else {
		window.open("MasCategoryAddEditPop.asp", "_blank", "width=500,height=300,resizable=no,status=no");
		window.setTimeout(popCategorty, 500);
	}
}

function popCategorty() {
	var objhttp = new XMLHttpRequest();
	var data = window.TempData || document.TempData;
	var root;
	var category;
	var option;
	objhttp.open("GET", "XMLGetCategoryDetails.asp", false);
	objhttp.send(null);
	if (trimValue(objhttp.responseText) !== "" && data && data.loadXML) {
		data.loadXML(objhttp.responseText);
	}
	root = data && data.documentElement;
	if (root && root.hasChildNodes() && document.formname.hspClass.value.length === 3) {
		document.formname.selCategory.length = 0;
		for (var i = 0; i < root.childNodes.length; i += 1) {
			category = root.childNodes[i];
			if (category.nodeType !== 1) {
				continue;
			}
			option = document.createElement("option");
			option.value = trimValue(category.attributes[0] && category.attributes[0].nodeValue);
			option.text = trimValue(category.attributes[1] && category.attributes[1].nodeValue);
			document.formname.selCategory.add(option);
		}
	}
}
</script>
</HEAD>
<%
	Dim oDom,fs,Root,PGNode
	dim scatID,scatName,scatShName,arrTemp,iClass,sDisabled

	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set fs = CreateObject("Scripting.FileSystemObject")

	dim spClass

	spClass = trim(Request.Form("pGroup"))
	if not (isNull(spClass) or isEmpty(spClass) or spClass = "") then
	%>
	<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload ="popCategorty()">
	<%
	if len(spClass) > 3 then
		arrTemp = split(spClass,":")
		iClass = arrTemp(1)
		sDisabled = " DISABLED "
	elseif len(spClass) = 3 then
		sDisabled = ""
		iClass = "00"
	end if
%>

<form method="POST" name="formname" action="" TARGET="bodyFrame">
<input type=hidden name=hItmType value="">
<input type='hidden' name="hspClass" value="<%=spClass%>">
	<table border="0" cellspacing="0" width="100%" cellpadding="0">
		<tr>
			<td class="ExcelHeaderCell" colspan="3"><p align="center">Define New Classification</td>
		</tr>
		<tr>
			<td width="10" colspan="3" class="MiddlePack"></td>
		</tr>
		<tr>
			<td width="5"></td>
			<td>
				<table cellpadding="0" cellspacing="0" width="100%" border="0">
					<tr>
						<td>
							<table cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td class='GroupTitleLeft' width="10">&nbsp;</td>
									<td class='GroupTitle' width="60"><p align="center">Details</td>
									<td class='GroupTitleRight'><p align="left">&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class=GroupTable>
							<table cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td class=MiddlePack colspan="2"> </td>
								</tr>
							<%	if len(spClass) = 3 then %>
								<tr>
									<td class=FieldCellSub width="62"> Category</td>
									<td class='FieldCellSub'>
										<select size="1" name="selCategory" class="FormElem">
											<option value="select">Select</option>
										</select>&nbsp;&nbsp;&nbsp;<input type=button name=BtnAddNew class="ActionButtonX" value="Add New" onclick="CatAddNew()">&nbsp;
									</td>
								</tr>
							<%	end if%>
								<tr>
									<td class=FieldCellSub width="62"> Name</td>
									<td class='FieldCellSub'>
										<input type="text" name="txtClassName" size="30" maxlength=40 class="FormElem">
									</td>
								</tr>
								<!--<tr>
									<td class=FieldCellSub width="62"> Item Type</td>
									<td class='FieldCellSub'>
										<select size="1" name="selItemType" class="FormElem" <%=sDisabled%>>
											<%	'Calling the Function which populates the Item Type list
											'	populateItemType iClass
											%>
										</select>
									</td>
								</tr>-->
								<tr>
									<td class=MiddlePack colspan="2"> </td>
								</tr>
								<tr>
									<td class=ActionCell colspan="2"> <p align="center">
								    <input type="button" value="Save" name="B2" class="ActionButton" onClick="javascript:CheckSubmit()">
								    <input type="reset" value="Reset" name="B3" class="ActionButton"></td>
								</tr>
							</table>
						</td>
						<td width="5"></td>
					</tr>
				</table>
				<INPUT type=hidden value="<%=spClass%>" name=hpGroup>
			</td>
		</tr>
	</table>
</form>
</BODY>
<%	end if %>

</HTML>

<%
	' Function to populate the Item Type list
	Function populateItemType(iClassP)
		' Declaration of variables
		Dim dcrs,stypID,stypName,sPTypeID

		dim iSAApplicationPop,iSAProcessPop,iSAActivityPop,iEmpNoPopulate

		iSAApplicationPop = Session("iApplication")
		iSAProcessPop = Session("iProcess")
		iSAActivityPop = Session("iActivity")
		iEmpNoPopulate = Session("employeenumber")

		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMTYPEID FROM INV_M_CLASSIFICATION WHERE GROUPCODE = " & iClassP & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		If not dcrs.EOF then
			sPTypeID = trim(dcrs(0))
		end if
		dcrs.Close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
'			if iSAApplicationPop <> "" then
'			.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE WHERE ITEMTYPEID IN (SELECT DISTINCT ITEMTYPEID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY ITEMTYPENO"
'			else
			.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE ORDER BY ITEMTYPENO"
'			end if
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
		    Do While Not dcrs.EOF
				if sPTypeID = stypID then
					Response.Write("<OPTION VALUE="""&trim(stypID)&""" SELECTED>"&trim(stypName)&"</OPTION>" &vbcrlf)
				else
					Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				end if
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>

