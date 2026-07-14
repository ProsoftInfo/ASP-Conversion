<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItmAccHead.asp
	'Module Name				:	Inventory
	'Author Name				:	Ragavendran R
	'Created On					:	
	'Modified By				:	
	'Tables Used				:	
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	
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
<!--#include virtual="/include/GetSerialDetail.asp"-->

<%
Dim dcrs
Dim sOrgCode,arrAHead,sOAH,sCAH
Dim iCounter,iCtr
set dcrs = Server.CreateObject("ADODB.Recordset")
sOrgCode = session("organizationcode")
sOAH = Request("OAH")
sCAH = Request("CAH")
Response.write "<font color=red>"

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT IA.ACCOUNTHEAD,IA.ACCOUNTHEADCODE,IA.ACCOUNTDESCRIPTION FROM ACC_M_GLACCOUNTHEAD IA,Acc_R_GLAccApplications D where IA.AccountHead = D.AccountHead and AvailableInAppln = 4 ORDER BY IA.ACCOUNTHEAD"
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

iCounter = dcrs.RecordCount

if not dcrs.EOF then
	arrAHead = dcrs.GetRows()
end if
dcrs.Close


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Item Account Head</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" id="AccHeadData" data-itms-xml-island="1"><Root></Root></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/selection.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT>
function selectedAccount(select) {
	var option = select.options[select.selectedIndex];
	var value = option ? option.value : "0";
	var name = option ? option.text : "";
	if (String(value).replace(/^\s+|\s+$/g, "").toLowerCase() === "select") {
		return { value: "0", name: "" };
	}
	return { value: value, name: name };
}

function CheckSubmit() {
	var opening = selectedAccount(document.formname.selOpening);
	var closing = selectedAccount(document.formname.selClosing);
	var data = window.AccHeadData || document.AccHeadData;
	var root = data.documentElement;
	var child = data.createElement("AccHead");
	while (root.firstChild) {
		root.removeChild(root.firstChild);
	}
	child.setAttribute("OAHV", opening.value);
	child.setAttribute("OAHN", opening.name);
	child.setAttribute("CAHV", closing.value);
	child.setAttribute("CAHN", closing.name);
	root.appendChild(child);
	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.returnAndClose(root);
	}
}
</SCRIPT>
</HEAD>
<BODY>
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Account Head
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top" class=TabBodyWithTopLine>
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<TD >
						<table cellpadding="0" cellspacing="0" width="100%">
							<!--<tr>
								<td class=FieldCellSub> Purchase</td>
								<td class='FieldCellSub'>
									<select size="1" name="selPurchase" class="FormElem">
										<option value="select">Select</option>
								<%
								'	for iCtr = 0 to iCounter - 1
								'		if arrAHead(3,iCtr) = "P" or arrAHead(3,iCtr) = "A" then
								'			Response.Write "<option value="""&arrAHead(0,iCtr)&""">"&arrAHead(2,iCtr)&"</option>"
								'		end if
								'	next
								%>
									</select>
								</td>
								<td class='FieldCellSub'></td>
							</tr>
							<tr>
								<td class=FieldCellSub> Sub-Contracting</td>
								<td class='FieldCellSub'>
									<select size="1" name="selSubCon" class="FormElem">
										<option value="select">Select</option>
								<%'	iCtr = 0
								'	for iCtr = 0 to iCounter - 1
								'		if arrAHead(3,iCtr) = "S" or arrAHead(3,iCtr) = "A" then
								'			Response.Write "<option value="""&arrAHead(0,iCtr)&""">"&arrAHead(2,iCtr)&"</option>"
								'		end if
								'	next
								%>
									</select>
								</td>
								<td class='FieldCellSub'></td>
							</tr>
							<tr>
								<td class=FieldCellSub> Finished Goods / Sales</td>
								<td class='FieldCellSub'>
									<select size="1" name="selSales" class="FormElem">
										<option value="select">Select</option>
								<%'	iCtr = 0
								'	for iCtr = 0 to iCounter - 1
								'		if arrAHead(3,iCtr) = "F" or arrAHead(3,iCtr) = "A" then
								'			Response.Write "<option value="""&arrAHead(0,iCtr)&""">"&arrAHead(2,iCtr)&"</option>"
								'		end if
								'	next
								%>
									</select>
								</td>
								<td class='FieldCellSub'></td>
							</tr>-->
							<tr>
								<td class="FieldCellSub"> Opening Stock</td>
								<td class="FieldCellSub">
									<select size="1" name="selOpening" class="FormElem">
										<option value="select">Select</option>
								<%	iCtr = 0
									for iCtr = 0 to iCounter - 1
									'	if arrAHead(3,iCtr) = "O" or arrAHead(3,iCtr) = "A" then
									    if trim(sOAH)=trim(arrAHead(0,iCtr)) then
											Response.Write "<option value="""&arrAHead(0,iCtr)&""" selected>"&arrAHead(2,iCtr)&"</option>"
										else
										    Response.Write "<option value="""&arrAHead(0,iCtr)&""">"&arrAHead(2,iCtr)&"</option>"
										end if
									'	end if
									next
								%>
									</select>
								</td>
								<td class="FieldCellSub"></td>
							</tr>
							<tr>
								<td class=FieldCellSub> Closing Stock</td>
								<td class='FieldCellSub'>
									<select size="1" name="selClosing" class="FormElem">
										<option value="select">Select</option>
								<%	iCtr = 0
									for iCtr = 0 to iCounter - 1
									'	if arrAHead(3,iCtr) = "C" or arrAHead(3,iCtr) = "A" then
									    if trim(sCAH)=trim(arrAHead(0,iCtr)) then
											Response.Write "<option value="""&arrAHead(0,iCtr)&""" selected>"&arrAHead(2,iCtr)&"</option>"
										else
										    Response.Write "<option value="""&arrAHead(0,iCtr)&""">"&arrAHead(2,iCtr)&"</option>"
										end if 
									'	end if
									next
								%>
									</select>
								</td>
								<td class='FieldCellSub'></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
                    <td class="actioncell" align="center">
                        <table>
                            <tr>
                                <td align="center">
                                    <input type="button" name="btnDone" value="Done" class="ActionButton" onclick="CheckSubmit()" />
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
