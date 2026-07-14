<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ConfigureProfitLossAcc.asp
	'Module Name				:	ACCOUNTS (Master Amendment)
	'Author Name				:	Manohar Prabhu .R
	'Created On					:	Nov 27 2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:	April 19,2011
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
<!--#include virtual="/include/Accpopulate.asp"-->
<%
	Dim sQuery,Objrs1,Objrs2,Objrs3,iCtr,sOrgID,sCatCode
	Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs2 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs3 = Server.CreateObject("ADODB.RecordSet")

	sOrgID = session("organizationcode")
	sCatCode = Request.Form("selCategory")

	IF CStr(sCatCode) = "" Then
		sCatCode = "0"
	End IF

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><title>iTMS</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="/Scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = {
	type: "glAliasConfiguration",
	page: "ConfigureProfitLossAcc.asp",
	schedulePage: "SchSetup_ForPL.asp",
	breakupPage: "SchBreakupSetup_forPL.asp"
};
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="setBalPlUpdate.asp">
	<Input type="hidden" name="hOrgID" value="<%=sOrgID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle"><p align="center">GL A/C
          Head Alias
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%"  >
				<TR>
					<td height="20px" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="105px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
										<tr>
											<td align="center">GL A/C Alias
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="110px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<a href="SchSetup_ForPL.asp"><td align="center">Schedule Setup</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="165px">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<a href="SchBreakupSetup.asp"><td align="center">Schedule Breakup Setup</td></a>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="75px">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
									  	<a href="PLSetup.asp"><td align="center">PL Setup</td></a>
									</tr>
								  </table>
								</td>
								<!--
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<a href="BSSetup.asp"><td align="center">BS Setup</td></a>
								  	</tr>
								  </table>
								</td>
								-->
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%" >
				<TR>
					<TD class="TabBody">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td class="FieldCellSub" width="50px">Category</td>
								<td class="FieldCell" colspan="2" width="110px">
									<select size="1" name="selCategory" class="FormElem" onChange="DisplayVal()">
										<Option Value="0" Selected>Select</Option>
											<%
												sQuery = "SELECT CategoryCode, CategoryName from Acc_M_AccountCategory order by CategoryCode "
												With Objrs1
													.CursorLocation = 3
													.CursorType = 3
													.Source = sQuery
													.ActiveConnection = Con
													.Open
												End WIth
												Set Objrs1.ActiveConnection = Nothing
												Do While Not Objrs1.EOF
													IF CStr(sCatCode) = CStr(Objrs1(0)) Then
											%>
														<Option Value="<%=Objrs1(0)%>" Selected><%=Objrs1(1)%></Option>
											<% Else %>
														<Option Value="<%=Objrs1(0)%>"><%=Objrs1(1)%></Option>
											<%
												End IF
												Objrs1.MoveNext
												Loop
												Objrs1.Close
											%>
										</select>
									</td>
								</tr>
							</table>

							<table border="0" cellpadding="0" cellspacing="0" width="100%">

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
   <%IF CStr(sOrgID) <> "" and sCatCode <> "0" Then %>
                            <tr>
								<td align="center"></td>
								<td valign="top" width="100%">
<DIV class=frmBody id=frm4 style="width: 585; height:270;">

<table BORDER="0" CELLSPACING="1"   CELLPADDING="0" class="ExcelTable">
<tr>
<td class="ExcelHeaderCell" colspan="3"><p align="left">Group
  Name</td>
</tr>

<tr>
<td class="ExcelHeaderCell" width="10"><p></td>
<td class="ExcelHeaderCell"><p>A/C Head Name</td>
<td class="ExcelHeaderCell">A/C Head Alias</td>
</tr>
<%
	sQuery = "SELECT CategoryCode, CategoryName from Acc_M_AccountCategory Where CategoryCode = '"&sCatCode&"' "
	With Objrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End WIth
	Set Objrs1.ActiveConnection = Nothing
	Do While Not Objrs1.EOF
%>
<tr>
<td align="left" class="ExcelDisplayCell" colspan="3"><b><%=Objrs1(1)%></b></td>
</tr>
<%
	sQuery = "SELECT AccountsGroupCode,AccountsGroupName,AccountsParentGroup,GroupCategory FROM  "&_
			 "Acc_M_AccountGroups Where GroupCategory = '"&Objrs1(0)&"' ORDER BY GroupHierarchy "
	With Objrs2
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End WIth
	Set Objrs2.ActiveConnection = Nothing
	Do While Not Objrs2.EOF
%>
<tr>
<td align="left" class="ExcelDisplayCell" colspan="3"><b>
<%
	For iCtr = 0 To Len(Objrs2(0)) * 1
		Response.Write("&nbsp;")
	Next
	Response.Write(Objrs2(1))
%>
</b></td>
</tr>
<%
	sQuery = "Select Distinct AccountHead,isNull(AccHeadAlias,AccountDescription) AccHeadAlias,AccountDescription  "&_
			 "from VwOrgGLHeads   "&_
			 "where AccountsGroupCode='"&Objrs2(0)&"' and OUDefinitionID = '"&sOrgID&"' "
	With Objrs3
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = Con
		.Open
	End WIth
	Set Objrs3.ActiveConnection = Nothing
	Do While Not Objrs3.EOF
%>
<tr>
<td align="left" class="ExcelDisplayCell" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <font color="#FF0000"><%=Objrs3(2)%></font></td>
<td align="center" class="ExcelInputCell">
  <p align="left">
  <%IF CStr(Objrs3(1)) = "" Then %>
		<input type="text" name="txtAcc<%=Objrs3(0)%>" size="60" class="FormElem" Value="<%=Objrs3(2)%>" onBlur="CheckVal(this)"></td>
	<%Else%>
		<input type="text" name="txtAcc<%=Objrs3(0)%>" size="60" class="FormElem" Value="<%=Objrs3(1)%>" onBlur="CheckVal(this)"></td>
	<%End IF %>
</tr>

<%
	Objrs3.MoveNext
	Loop
	Objrs3.Close

	Objrs2.MoveNext
	Loop
	Objrs2.Close

	Objrs1.MoveNext
	Loop
	Objrs1.Close
%>

</table>
</DIV>
								</td>
								<td align="center">
								</td>
                            </tr>

                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
     <% End IF %>

                 <tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td class="ActionCell">
                                                <input type="button" value="Save" name="B4" class="ActionButton" onClick="CheckSubmit()">
                                                <input type="button" value="Schedule Setup" name="B7" class="ActionButtonX" onClick="Sch()">
                                                <input type="button" value="Sch Breakup Setup" name="B6" class="ActionButtonX" onClick="SchBrk()">
                                                <input type="button" value="PL Setup" name="B8" class="ActionButton" onClick="PL()">
                                                <!--<input type="button" value="BS Setup" name="B10" class="ActionButton" onClick="BS()">-->

												<!--input type="button" value="Cancel" name="B5" class="ActionButton"-->

											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
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
