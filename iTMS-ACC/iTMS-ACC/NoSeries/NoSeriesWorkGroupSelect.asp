<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	NoSeriesWorkGroupSelect.asp
	'Module Name				:	Maintenance (Master)
	'Author Name				:	Kalaiselvi R
	'Created On					:	01 january 2009
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
<%
Dim sOrgId,sGetVal,sQuery,sItemVal

Dim saTemp

Dim objRs

'---------- Getting Values From Party Head Selection Page -------------

	sGetVal=Request.QueryString("Value")
	saTemp=split(sGetVal,":")
	sOrgId=saTemp(0)
	sItemVal = saTemp(1)

	Set objRs = Server.CreateObject("ADODB.RecordSet")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Work Group</TITLE>

<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Selection.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/NoSeriesEntryCompat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" LANGUAGE=javascript onunload="return window_onunload()">
<form method="POST" name="formname" action="">

<div align="center">

						<table border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" height="2">
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%" align="center">
                                    <table cellpadding="0" cellspacing="0">
                                <tr>
									<td class="FieldCell">
									</td>
                                </tr>
                                <tr>
                            <td class="FieldCell">
<%

	sQuery = " Select WorkGroupCode,WorkGroupName from PRD_M_WORKGROUP order by WorkGroupName"
	
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	%>
	 <select size="10" name="SelWorkGroup"  class="FormElem" multiple>
	<%
		

	

	If not objRs.EOF then
		Do While Not objRs.EOF
			
			Response.Write("<OPTION VALUE="& trim(objRs(0)) &">"&  trim(objRs(1)) &"</OPTION>")
			
			objRs.MoveNext
			
		Loop
	end if
	objRs.Close

%>
 </select>
                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5" height="2">
                                    <img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Done" name="B7" onclick="checksubmit()" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="B8" onClick="finalcancel()" class="ActionButton">
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="BottomPack" colspan="3">
								</td>
                                </tr>
						</table>
 </div>
</form>
</BODY>
</HTML>
