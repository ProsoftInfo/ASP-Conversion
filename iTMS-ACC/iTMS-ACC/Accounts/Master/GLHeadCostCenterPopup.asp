<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GLHeadCostCenterPopup.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 25,2010
	'Modified On				:
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
<!--#include virtual="/include/sessionVerify.asp"-->
<%
dim sGlHeadName,iAccHead,sArrUnit,sUnits,iCnt,sQuery,sUnitID,sUnitName,iSNo
dim objRs
'XML DOM Variables
Dim oDOM,nodHeader,Root,nodUnit,sGroupName,sSelCostCode

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iAccHead = Request.QueryString("AccHead")
sGLHeadName = Request.QueryString("HeadName")
sGroupName	= Request.QueryString("GroupName")
sUnits = Request.QueryString("Units")
sSelCostCode = Request.QueryString("hSelCostCode")
'Response.Write "sSelCostCode= "& sSelCostCode
sQuery = "Select OUDefinitionID,OrgUnitDescription from DCS_OrganizationUnitDefinitions where Len(OUDefinitionID) > 4 and OUDefinitionID = '"& sUnits  &"'"
objRs.Open sQuery,con
if not objRs.EOF then
	sUnitName  = objRs(1)
end if
objRs.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT src="../../scripts/Selection.js"></SCRIPT>
<script>
function final()
{
		if (document.formname.selTobox.length != 0)
		{
			finaldone('selTobox','hSelectedValue');
		}
}

function finaldone(rightCombo,hiddenFieldName)
{
	var i,par
	par="";

	for (i=0;i<(document.forms[0].elements[rightCombo].options.length)-1;i++)
	{
		par= par+document.forms[0].elements[rightCombo].options[i].value+",";
	}
	if(document.forms[0].elements[rightCombo].options.length==0)
	{
		//par="m-1"
		alert("Select the item");
		return false;
	}
	else
	{
		par= par+document.forms[0].elements[rightCombo].options[i].value;
	}

	document.forms[0].elements[hiddenFieldName].value=par;
}

</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/ModalReturnCompat.js"></script>
<script>
window.__itmsPopupCompat = { type: "glHeadCostCenterPopup" };
window.ITMSModalReturnCompat.install(function () {
	return window.ITMSModalReturnCompat.dialogArgumentsRoot();
});
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="init()" >

<form method="POST" name="formname">
    <input type=hidden name="hSelectedValue" value="<%=sSelCostCode%>">
    <input type=hidden name="hUnitCode" value="<%=sUnits%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"><%=sUnitName%><br>Cost Center
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
                                    <table cellpadding="0" cellspacing="0" width="100%">
									    <tr>
										<td class="FieldCell" width="145" valign="top">GL Classification</td>
										<td><span class="DataOnly"><%=sGroupName%></span>
										</td>
									    </tr>
									<tr>
									<td class="FieldCell" width="145">GL Account Head Name</td>
									<td>
										<span class="DataOnly"><%=sGlHeadName%></span>
									</td>
									</tr>
                                    </table>
								</td>
								<td align="center" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td valign="top" colspan=3 class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td align="center"  class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td>
								<div class="frmbody" style="height:175px;width:300px">
                                    <table border="0" cellspacing="1" width="100%" class="ExcelTable">
									<tr>
										<td class="ExcelHeaderCell">S.No.
										</td>
										<td class="ExcelHeaderCell">
										</td>
										<td class="ExcelHeaderCell" align=center>
											Cost Center
										</td>
									</tr>
									<%
										iSNo = 1
											sQuery="select CCGroupCode,CostCenterHead,CCGroupName,CCHeadCode,CCAccountDescription from VwOrgCostCenter where OUDefinitionID='"& sUnits &"' "

										'	Response.Write "<option>"& sQuery & "</option>"
											with objRs
												.CursorLocation =3
												.CursorType =3
												.Source =sQuery
												.ActiveConnection = con
												.Open
											end with
											set objRs.ActiveConnection=nothing

											do while not objRs.EOF
											%>
												<tr>
													<td class="ExcelHeaderCell" width=10><%=iSNo%>
													</td>
													<td class="ExcelDisplayCell" width=10>
														<%
														Response.Write "<input type=Checkbox name=chkCostCenter"&iSNo&" value="& sUnits &":"& objrs(0) &":" & objrs(1) & ">"
														%>
													</td>
													<td class="ExcelDisplayCell">
													<%=objrs(2)&":"& objRs(4)%>
													</td>
												</tr>
											<%
												iSNo = iSNo + 1
												objRs.MoveNext
											loop
											objRs.Close
									%>
						        <input type=hidden name=hTotalCostCenter value="<%=iSNo-1%>">
                              </table>
                              </div>
								</td>
								<td align="center"  class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td valign="top" colspan=3 class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td valign="top" colspan=3 class="BottomPack">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Save" name="next" class="ActionButton" ONCLICK="CheckSubmit()"  tabindex="3" >
																<input type="button" value="Close" name="close" class="ActionButton" onClick="Popup_Close()">
														</td>
													</tr>
												</table>
								</td>
							</tr>
							<tr>
								<td valign="top" colspan=3 class="BottomPack">
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
