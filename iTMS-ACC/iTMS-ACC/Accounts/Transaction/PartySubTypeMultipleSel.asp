<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PartySubTypeMultipleSel.asp 
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MAHESHWARI S
	'Created On					:	Feb 26,2007
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>TDS</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT SRC="../../scripts/RoundOff.js"></SCRIPT>
<SCRIPT SRC="../../scripts/checkdate.js"></SCRIPT>
<SCRIPT SRC="../../scripts/TempItem.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<!-- XML Data Island -->
 



<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/VouTransactions.js"></script>

<SCRIPT>
var objTemp = null;

function parseXml(text) {
	return new DOMParser().parseFromString(text || "<AccHead/>", "text/xml");
}

function dialogDocument() {
	var args = window.dialogArguments;
	if (!args && window.opener && window.opener.__itmsDialogArgs) {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		if (match) {
			args = window.opener.__itmsDialogArgs[decodeURIComponent(match[1])];
		}
	}
	if (args && args.nodeType === 9) {
		return args;
	}
	if (args && args.nodeType === 1) {
		return args.ownerDocument;
	}
	if (args && args.documentElement) {
		return args;
	}
	if (args && args.XMLDocument) {
		return args.XMLDocument;
	}
	if (typeof args === "string") {
		return parseXml(args);
	}
	return parseXml("<AccHead/>");
}

function childElements(node, nodeName) {
	var result = [];
	var wanted = String(nodeName || "").toLowerCase();
	for (var i = 0; node && i < node.childNodes.length; i += 1) {
		if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
			result.push(node.childNodes[i]);
		}
	}
	return result;
}

function Init() {
	var root;
	var nodes;
	var iCnt;
	var sChkVal;
	var item;
	objTemp = dialogDocument();
	root = objTemp.documentElement;
	iCnt = parseInt(document.formname.hCnt.value, 10) || 0;
	nodes = childElements(root, "Details");
	for (var n = 0; n < nodes.length; n += 1) {
		sChkVal = [nodes[n].getAttribute("PartyType") || "", nodes[n].getAttribute("PartySubType") || "", nodes[n].getAttribute("PartySubTypeName") || ""].join("?");
		for (var i = 1; i <= iCnt; i += 1) {
			item = document.formname.elements["ChkType" + i];
			if (item && String(item.value).trim() === sChkVal.trim()) {
				item.checked = true;
			}
		}
	}
}

function CheckSubmit() {
	var root;
	var nodes;
	var iCnt;
	var item;
	var parts;
	var elem;
	objTemp = objTemp || dialogDocument();
	root = objTemp.documentElement;
	nodes = childElements(root, "Details");
	for (var n = 0; n < nodes.length; n += 1) {
		root.removeChild(nodes[n]);
	}
	iCnt = parseInt(document.formname.hCnt.value, 10) || 0;
	for (var i = 1; i <= iCnt; i += 1) {
		item = document.formname.elements["ChkType" + i];
		if (item && item.checked) {
			parts = String(item.value || "").split("?");
			elem = objTemp.createElement("Details");
			elem.setAttribute("PartyType", parts[0] || "");
			elem.setAttribute("PartySubType", parts[1] || "");
			elem.setAttribute("PartySubTypeName", parts[2] || "");
			root.appendChild(elem);
		}
	}
	Win_UnLoad();
}

function Win_UnLoad() {
	var value = objTemp && objTemp.documentElement ? objTemp.documentElement : objTemp;
	window.returnValue = value;
	window.returnvalue = value;
	if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
		window.ITMSModernCompat.returnModalValue(value);
	}
	window.close();
}
</script>
</HEAD>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="Init()">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr>
		<td align="center" class="TopPack">
		</td>
    </tr>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		Account Head </td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5"  >
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell" width="100%">
									 
								</td>
								<td align="center" class="ClearPixel">
								</td>
							</tr>
                        <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" >
								</td>
                        </tr>
                        <tr>
							<td align="center" class="ClearPixel">
							</td>
							<td valign="top" class="FieldCell" width="100%" align="center">
                            <DIV class=frmBody id=frm3 style="width: 360; height:380;">
			
								<table border="0" cellspacing="1" class="BodyTable" width="100%" id=TabAccHead>
								<% dim objRs,objRs1,sQuery,iCnt
								dim sOrgId,sCallTy,sParType,iParSubType,sParTypeName

								set objRs = Server.CreateObject("ADODB.Recordset")
								set objRs1 = Server.CreateObject("ADODB.Recordset")
								sOrgId = Request("Unit")
								 
								'Response.Write sUnit
								sCallTy = Request("sCallTy")
								

								IF CStr(sCallTy) <> "P" Then
									sQuery="select distinct PartyType,PartySubType,SubTypeName from vwOrgPartyType where OUDefinitionID='"&sOrgId&"'"
								Else
									sQuery = "select distinct PartyType,PartySubType,SubTypeName from vwOrgPartyType where OUDefinitionID='"&sOrgId&"' "&_
											 "and ((PartyType = 'CR' and PartySubType >= 2) or PartyType = 'DR') "
								End IF

									With objRs
										.CursorLocation = 3
										.CursorType = 3
										.Source = sQuery
										.ActiveConnection = con
										.Open
									End with

									Set objRs.Activeconnection = nothing

									if not objRs.EOF then
										 
										 iCnt = 0
										Do while not objRs.EOF
											
											sParType=objRs(0)
											iParSubType = objRs(1)
											sParTypeName = Replace(objRs(2)	,"&"," and ")
											sQuery="select count(1) from APP_R_OrgParty where PartyType='"&sParType&"' and PartySubType="&iParSubType&" and OUDefinitionID='"&sOrgId&"'"
											with objRs1
												.CursorLocation =3
												.CursorType =3
												.Source = sQuery
												.ActiveConnection = con
												.Open
											end with
											set objRs1.ActiveConnection=nothing
											if CDbl(objRs1(0)) then
											iCnt = iCnt + 1
													%>
													<tr>
													<td class="FieldCellSub" valign="top"></td>
													<td class="FieldCell" valign="top">
													<input type="checkbox" name="ChkType<%=iCnt%>" value="<%=sParType%>?<%=iParSubType%>?<%=sParTypeName%>" class="formelem" >  
													</td>
													<td class="FieldCellSub" valign="top"></td>
													<td class="FieldCell" valign="top"><%=sParTypeName%></td>
													</tr>		
																				
													<%
													
													end if			
													objRs1.Close
												  	objRs.MoveNext		  	
												Loop				
												 
											end if
											objRs.Close
											%>
								<input type=hidden name="hCnt" value="<%=iCnt%>">
								</table>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
									<td valign="middle" class="ActionCell">
                                        <p align="center">
                                        <input type="button" value="Done" name="B2" class="ActionButton" tabindex="3" onclick = "CheckSubmit()">
                                        
									</td>
									</tr>
								</table>
							</div>
							</td>
						</tr>
							
							
					</table>
					</td>
				</tr>
				

					
							
													
						
</table>
</form>
</BODY>
</HTML>
