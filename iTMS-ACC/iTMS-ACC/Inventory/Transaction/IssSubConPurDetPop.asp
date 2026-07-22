<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	IssSubConPurDetPop.asp
	'Module Name				:	Inventory (Issue - SubContract Case)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	MARCH 23,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None

%>
<%
	Dim sHeading,sOrgCode
	sHeading = Request.QueryString("sHead")
	sOrgCode = Request.QueryString("OrgCode")
	'Response.Write sOrgCode
%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/purpopulate.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="RefData"><Root Confirm="N"/></script>
<script type="application/xml" data-itms-xml-island="1" id="PurTypeData"><Root/></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/issSubConPurDetPop.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad ="init()">
<form method="POST" name="formname">
<INPUT TYPE=HIDDEN NAME="hPurType" value="">
<INPUT TYPE=HIDDEN NAME="hPurTypeName" value="">
<INPUT TYPE=HIDDEN NAME="selUnit" value="<%=sOrgCode%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"><%=sHeading%>
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
									<table border=0 cellspacing=0 cellpadding=0 width="100%">
										<tr>
												<td class="FieldCell" valign="top">Purchase Type
												</td>
												<td class="FieldCellSub" colspan="4">
												<select size="1" class="FormElem" name="cmbPurType" onChange="ChangePurType()">
														<option  value="" selected>Select</option>
														<option  value="0">--------------ITEMWISE-------------</option>
														<%
															popSelPurTypeFull("")
														%>
													</select>
												</td>
											</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="ActionCell">
									<input type=button name="btnProceed" value="Submit" class="ActionButtonX" onClick="FinalSubmit()">
									<input type=button name="btnReset" value="Cancel" class="ActionButtonX" onClick="window.close">
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
Function popAccountHead(sOrgID,iSelAccHead,sStockType)

Dim dcrs,iAccHeadNo,sAccHeadName,iCCExist,iAHExist,sSqlTemp

Set dcrs = Server.CreateObject("ADODB.RecordSet")

if trim(sStockType) = ""  then sStockType = "S"

sStockType = ucase(trim(sStockType))


	'sSqlTemp = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode  from VwAccHeadForInventApp A,Acc_M_GLAccountHead B where A.OUDefinitionID='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
	if trim(sStockType) = "C" then
		sSqlTemp = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode  from VwAccHeadForFAApp A,Acc_M_GLAccountHead B where A.OUDefinitionID='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
	else
		sSqlTemp = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode  from VwAccheadforPurchaseApp A,Acc_M_GLAccountHead B where A.OUDefinitionID='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
	end if
	'Response.Write "AccHead="&sSqlTemp
	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSqlTemp
		.ActiveConnection = con
		.Open
	End With

	Set dcrs.ActiveConnection = Nothing

	Set iAccHeadNo = dcrs(0)
	Set sAccHeadName = dcrs(1)
	Set iCCExist = dcrs(2)
	Set iAHExist = dcrs(3)

	Do while not dcrs.EOF
		if trim(iSelAccHead) = trim(iAccHeadNo) then
			Response.Write "<option value="&trim(iAccHeadNo)&" selected>" & trim(sAccHeadName) & "</option>" & vbCr
		Else
			Response.Write "<option value="&trim(iAccHeadNo)&">" & trim(sAccHeadName) & "</option>" & vbCr
		End if
		dcrs.MoveNext
	Loop
'End if
dcrs.Close

End Function
%>
