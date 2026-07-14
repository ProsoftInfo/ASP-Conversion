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
	'Program Name				:	NoSeriesItemSel.asp
	'Module Name				:	Sales (Master)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	25 May 2004
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
dim sOrgId,saTemp,sGetVal,objRs,sQuery,sType,sTitleName,sItemVal

'---------- Getting Values From Party Head Selection Page -------------
sGetVal=Request.QueryString("Value")
saTemp=split(sGetVal,":")
sOrgId=saTemp(0)
sType =saTemp(1)
sItemVal = saTemp(2)

IF CStr(sType) = "I" Then
	sTitleName = "Item Type "
Elseif CStr(sType) = "S" Then
	sTitleName = "Sale Type "
Elseif CStr(sType) = "A" Then
	sTitleName = "Commission Agent "
Else
	sTitleName = "Invoice Type "
End IF

Set objRs = Server.CreateObject("ADODB.RecordSet")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE><%=sTitleName%></TITLE>

<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData"><account/></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Selection.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/NoSeriesEntryCompat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" LANGUAGE=javascript onunload="return window_onunload()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hCallTy" Value="<%=sType%>">
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
 <!--input type="text" name="txtSearch" size="35"  ONKEYUP="selectTheItem(this,'selPartyHead')"  class="FormElem"-->
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell">
<%
'------------- Recordset For the Items based on the Criteria ----------------
dim iTypeCode,sTypeName,sTypeCode2

IF CStr(sType) = "I" Then
	sQuery = "Select ItemTypeID,ItemTypeName from Inv_M_ItemType "
Elseif CStr(sType) = "A" Then
	sQuery = "Select Distinct P.PartyCode,P.PartyName,P.OrgnPartyCode From APP_M_PartyMaster P, "&_
			 "APP_R_OrgParty R Where P.PartyCode = R.PartyCode and  "&_
			 "R.PartyType = 'CR' and R.PartySubType = 1 Order By P.PartyName "

Elseif CStr(sType) = "S" Then
	sQuery = "SELECT InvoiceType,InvoiceTypeName FROM Sal_M_InvoiceTypes  ORDER BY InvoiceType"
End IF

IF CStr(sType) = "I" or CStr(sType) = "S" or CStr(sType) = "A" Then
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
	if objRs.RecordCount>10 then
	%>
	 <select size="10" name="selPartyHead"  class="FormElem" multiple>
	<%else%>
	 <select size="<%=objRs.RecordCount%>" name="selPartyHead"  class="FormElem" multiple>
	<%
	end if
	set iTypeCode = objRs(0)
	set sTypeName = objRs(1)


	'----------- Displaying Party Subtypes in Selection Control -----------

	If not objRs.EOF then
		Do While Not objRs.EOF
			IF CStr(sType) = "A" Then
				sTypeCode2 = objRs(0)&":"&objRs(2)
				Response.Write("<OPTION VALUE="&sTypeCode2&">"&trim(sTypeName)&"</OPTION>")
			Else
				Response.Write("<OPTION VALUE="&iTypeCode&">"&trim(sTypeName)&"</OPTION>")
			End IF
			objRs.MoveNext
			sTypeCode2 = ""
		Loop
	end if
	objRs.Close
Elseif CStr(sType) = "V" Then
%>
	<select size="8" name="selPartyHead"  class="FormElem" multiple>

		<!-- Blocked by ragav<option value="A">MILL SALES </option>
		<option value="T">TRANSFERS TO DEPOT</option>
		<option value="U">TRANSFER TO GROUP UNITS / COMPANIES</option>
		<option value="C">TRANSFER TO CONVERTORS FOR CONVERSION</option>
		<option value="S">INVOICE FOR SAMPLE ITEMS</option>
		<option value="I">STOCK ISSUED FOR SALES</option>
		<option value="D">UNPROCESSED ORDERS / DIRECT</option>
		<option value="R">TRANSFER INVOICE - RETURN OF SUBCONTRACT ITEMS</option> end -->
		<!--option value="L">LABOUR CHARGES / CONVERSION CHARGES</option-->
		<Option Value="CB">Cash Bill</Option>
		<Option Value="NEB">Non Excise Bill</Option>
		<Option Value="EB">Excise Bill</Option>
	</Select>

<%
End IF
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
                                                                 <input type="reset" value="Reset" name="B9" class="ActionButton" >
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
