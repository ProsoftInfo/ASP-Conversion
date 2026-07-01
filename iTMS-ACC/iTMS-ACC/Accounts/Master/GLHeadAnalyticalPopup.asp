<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GLHeadAnalyticalPopup.asp
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<%
dim sGlHeadName,iAccHead,iSNo
dim objRs,objRs1,objRs2,objRs3,sExp3,UnitNode,sExp2
'XML DOM Variables
Dim oDOM,nodHeader,Root,nodUnit,sGroupName,sUnitName
Dim sExp,TempNode,sSelUnits,iCounter

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")
Set objRs3 = Server.CreateObject("ADODB.RecordSet")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
dim sQuery,sArrUnit ,iCnt,sUnits,sSelAnyCode

iAccHead = Request.QueryString("AccHead")
sGLHeadName = Request.QueryString("HeadName")
sGroupName	= Request.QueryString("GroupName")
sUnits = Request.QueryString("Units")
sSelAnyCode = Request.QueryString("hSelAnayCode")
'Response.Write "sSelAnyCode = "& sSelAnyCode
if iAccHead="" then iAccHead = 0

sQuery = "Select OUDefinitionID,OrgUnitDescription from DCS_OrganizationUnitDefinitions where Len(OUDefinitionID) > 4 and OUDefinitionID = '"& sUnits  &"'"
objRs1.Open sQuery,con
if not objRs1.EOF then
	sUnitName = objRs1(1)
end if
objRs1.Close

'Response.Write "iAccHead = "& iAccHead
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript src="../../scripts/Selection.js"></SCRIPT>
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
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/ModalReturnCompat.js"></script>
<script language="javascript">
window.__itmsPopupCompat = { type: "glHeadAnalyticalPopup" };
window.ITMSModalReturnCompat.install(function () {
	return window.ITMSModalReturnCompat.dialogArgumentsRoot();
});
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="init()" >

<form method="POST" name="formname" >
   <input type=hidden name="hSelectedValue" value="<%=sSelAnyCode%>">
   <input type="hidden" name="hUnitCode" value="<%=sUnits%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"><%=sUnitName%><br>Analytical Code
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
								<td align="center" colspan=3 class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center"  class="MiddlePack">
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
								<td align="center"  class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>

							</tr>
							<tr>
								<td align="center" colspan=3  class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center"  class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>

								<td align=center>
								<div class="frmbody" style="height:175px;width:300px">
			                      <table border="0" cellpadding="0" cellspacing="1" class="ExcelTable" width=100%>
									<tr>
									<td class="ExcelHeaderCell" width=10 >S.No.
									</td>
									<td class="ExcelHeaderCell" width=10>
									</td>
									<td class="ExcelHeaderCell" align=center >Analytical Head
									</td>
                                    </tr>
									<%
										iSNo = 1
											sQuery="select AnalyticalCode FROM Acc_M_AnalyticalHeads order by AnalyticalCode "
											with objRs1
												.CursorLocation =3
												.CursorType =3
												.Source = sQuery
												.ActiveConnection = con
												.Open
											end with
											set objRs1.ActiveConnection=nothing
											while not objRs1.EOF
												sQuery="select distinct AHGroupCode,AHGroupName,AnalyticalShortName from VwOrgAnalytical "&_
												" where AnalyticalCode="&objRs1(0)&" and OUDefinitionID = '"& sUnits &"' order by AHGroupName"
											'	Response.Write "<option>"& sQuery &"</option>"	& vbCrLf
													with objRs2
														.CursorLocation =3
														.CursorType =3
														.Source = sQuery
														.ActiveConnection = con
														.Open
													end with
													set objRs2.ActiveConnection=nothing
													if not objRs2.EOF then
													%>
													<tr>
													<td class="ExcelHeaderCell" width=10 align=center>
													<%=iSNo%>
													</td>
													<td class="ExcelDisplayCell" width=10>
													<%
														Response.Write "<input type=checkbox name=chkAnalyticalZ"&iSNo&" value="& sUnits &":"& objRs1(0) &":"& objRs2(0) & ">"
													%>
													</td>
													<td class="ExcelDisplayCell">
													<%=objRs2(1)&":"& objRs2(2)%>
													</td>
													</tr>
													<%
													iSNo = iSNo + 1
													end if
													objRs2.Close
												objRs1.MoveNext
											wend
											objRs1.Close

									%>
                                     </table>
                                     </div>
                                     <input type=hidden name="hTotalAnalCode" value="<%=iSNo-1%>">
								</td>
								<td align="center"  class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>

							<tr>
								<td align="center"  class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="BottomPack">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Save" name="Save" class="ActionButton" ONCLICK="CheckSubmit()"  >
																<input type="reset" value="Close" name="Close" class="ActionButton" onClick="Popup_Close()" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center"  class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td valign="top" colspan=3  class="BottomPack">
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
