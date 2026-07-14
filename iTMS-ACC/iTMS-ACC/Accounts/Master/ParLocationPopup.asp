<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParLocationPopup.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 15,2010
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
<!--#include virtual="/include/sessionVerify.asp"-->
<%
'XML DOM Variables
Dim oDOM,nodHeader,Root,Objrs,sQuery,sRecCount,Objrs1
dim sCity,sState,sCountry,sParName,sParCode,sECNo,sPanNo,sLocalTax,sCentralTax,iPartyCode
Dim sDelType,sLastLocNo,sTinNo
' Create our DOM Document Objects
iPartyCode = Request.QueryString("PartyCode")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set Objrs = Server.CreateObject("ADODB.RecordSet")
Set Objrs1 = Server.CreateObject("ADODB.RecordSet")

sQuery = "Select PartyCode,PartyName,OrgnPartyCode,isNull(City,''),isNull(State,''),isNull(ExciseControlCode,''),isNull(IncomeTaxPANNo,''),"&_
		 "isNull(LocalSTNoandDT,''),isNull(CentralSTNoandDT,''),isNull(TINNumber,'') from APP_M_PartyMaster where PartyCode = "& iPartyCode

Objrs.Open sQuery,con
if not objrs.EOF then
	sParName = Objrs(1)
	sParCode = Objrs(2)
	sCity = Objrs(3)
	sState = Objrs(4)
	sECNo  = Objrs(5)
	sPanNo = Objrs(6)
	sLocalTax = Objrs(7)
	sCentralTax = Objrs(8)
	sTinNo = Objrs(9)
end if
Objrs.Close


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<base target="_self">
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<!-- XML Data Island --><script type="application/xml" data-itms-xml-island="1" id="OutData">
<Location>
<%
if trim(iPartyCode)<>"" then
	sQuery = "Select LocationCode,isNull(Location,''),isNull(LocationAddress1,''),isNull(LocationAddress2,''),isNull(City,''),isNull(State,''),isNull(Country,''), "&_
		     "isNull(LocalSTNoandDT,''),isNull(CentralSTNoandDT,''),isNull(ExciseControlNo,''),isNull(IncomeTaxPanno,'') From APP_M_PartyLocations "&_
		     "Where PartyCode = "&iPartyCode&" "

	With Objrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	sRecCount = Objrs.RecordCount
	Set Objrs.ActiveConnection = Nothing
	Do While Not Objrs.EOF
		sLastLocNo = Objrs(0)

%>
<Loc No="<%=Objrs(0)%>" Name="<%=Replace(Objrs(1),"&"," and ")%>" Address1="<%=Objrs(2)%>" Address2="<%=Objrs(3)%>" City="<%=Objrs(4)%>"
State="<%=Objrs(5)%>" Country="<%=Objrs(6)%>" ECCNo="<%=Objrs(7)%>" SalesLocal="<%=Objrs(8)%>"
SalesCentral="<%=Objrs(9)%>" PANNo="<%=Objrs(10)%>" Status="<%=CheckLoc(iPartyCode,Objrs(0))%>" />
<%
	Objrs.MoveNext
	loop
	Objrs.Close
end if ' if trim(iPartyCode)<>"" then
%>
</Location>
</script>
<%
	IF CStr(sLastLocNo) = "" Then
		sLastLocNo = 0
	End IF
%>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<script>
function validate()
{
	if (trim(document.formname.txtLocationName.value)=="")
	{
		alert("Enter Location Name");
		document.formname.txtLocationName.select();
		return false;
	}
	if (trim(document.formname.txtCity.value)=="")
	{
		alert("Enter City");
		document.formname.txtCity.select();
		return false;
	}

	return true;
}
</script>
<script>
window.__itmsPopupCompat = { type: "partyLocationPopup" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="PopulateLoc()">

<form method="POST" name="formname">
<Input type="hidden" name="hRecCount" value="<%=sRecCount+1%>">
<Input type="hidden" name="hEntNo" value="<%=sLastLocNo%>">
<input type="hidden" name="hPartyCode" value="<%=iPartyCode%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Party Location</p>
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
								<td align="center">
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=FieldCell>
													<table cellpadding="0" cellspacing="0">
													<tr>
															<td class="FieldCell" width="115" valign="top">Party Name</td>
															<td class="FieldsubCell" ><span class="DataOnly"><%=sParName%></td>
													</tr>
													<tr>
															<td class="FieldCell" width="115" valign="top">Party Code</td>
															<td class="FieldsubCell"><span class="DataOnly"><%=sParCode%></td>
													</tr>
													<!--tr>
															<td class="FieldCell" width="115" valign="top">Select Location</td>
															<td class="FieldsubCell">
															<Select name="SelLocName" class="FormElem" onChange="DispLocDet()">
															<Option Value="0">Select Location Name </Option>
															</td>
													</tr-->
														<tr>
															<td class=FieldCell width="115"> Location Name</td>
															<td class='FieldCell'><input type="text" name="txtLocationName" size="52" value="" maxlength="50" class="Formelem"></td>
														</tr>
													</table>
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell valign="top" align="left">
                                                    <table border="0" cellpadding="0" cellspacing="0">
                                                      <tr>
                                                        <td>
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="60"><p align="center">Address
                                                            </td>

															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable valign="top">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td class=MiddlePack colspan="2"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub width="65"> Address</td>
															<td class='FieldCell'><input type="text" name="txtAddress1" size="32" maxlength="50" class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub width="65"> </td>
															<td class='FieldCell'><input type="text" name="txtAddress2" size="32" maxlength="50" class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub width="65"> City</td>
															<td class='FieldCell' width="210">
															<input type="text"  name="txtCity" size="27" value="" maxlength="30" class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub width="65"> State</td>
															<td class='FieldCell' width="210"><input type="text" value="" name="txtState" size="27" maxlength="30" class="Formelem"></td>
														</tr>
														<tr>
															<td class=FieldCellSub width="65"> Country</td>
															<td class='FieldCell' width="210"><input type="text" value="" name="txtCountry" size="20" maxlength="30" class="Formelem"></td>
														</tr>

													</table>
                                                            </td>
														</tr>
													</table>
                                                        </td>
                                                        <td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                                                        </td>
                                                        <td valign="top">
                                                            <table border="0" cellspacing="0" cellpadding="0">
                                                        <tr>
                                                    <td class="MiddlePack" colspan="2">&nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="MiddlePack" colspan="2">&nbsp;&nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="175">Excise
                                                      ECC Number</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtEccNo" value="<%=sECNo%>" size="18" class="Formelem"> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="175">Sales
                                                      Tax Number - Local</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtSalesLocal"  value="<%=sLocalTax%>" size="18" class="Formelem"> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="175">Sales
                                                      Tax Number - Central</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtSalesCenteral"  value="<%=sCentralTax%>" size="18" class="Formelem"> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="175">IT
                                                      PAN Number</td>
                                                    <td class="FieldCellSub"> <input type="text" name="txtPanNo"  value="<%=sPanNo%>" size="18" class="Formelem"> </td>
                                                        </tr>
                                                        <tr>
														<td class="FieldCellSub" width="175">TIN Number
														  </td>
														<td class="FieldCellSub"> <input type="text" name="txtTinNo"  value="<%=sTinNo%>" size="18" class="FormelemRead" readonly> </td>
                                                        </tr>
                                                            </table>
                                                        </td>
                                                      </tr>
                                                    </table>
                                                            </td>
														</tr>
													</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                               <input type="button" value="Add Next" name="btnAdd" class="ActionButton" onClick="addEntry('A')"   >
                                                                <input type="button" value="Update" name="btnUpdate" class="ActionButton" onClick="updateEntry()" disabled   >
                                                               <input type="button" value="Delete Location" name="btnDel" onClick="DelEntry()" class="ActionButtonX" disabled>
                                                                <input type="button" value="Save" name="btnNext" onClick="addEntry('S')" class="ActionButton"  >
                                                                <input type="button" value="Cancel" name="btnCancel" onClick="window.close()"  class="ActionButton" >
                                                                <input type="Button" value="Reset" name="B1" onClick="clearXML()" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
												<DIV class=frmBody id=frm1 style="width: 585; height:160;">
                                                <table border="0" id="tblBin" cellspacing="1" class="ExcelTable" width="569">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="20"><p align="center">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="10"><p align="center">&nbsp;</td>
                                        <td class="ExcelHeaderCell" align="center" width="150">Location Name</td>
                                        <td class="ExcelHeaderCell" align="center" width="249"> Address</td>
                                        <td class="ExcelHeaderCell" align="center" width="150"></td>

                                            </tr>

                                                </table>
												</div>
								</td>
								<td align="center">
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
<%
	Function CheckLoc(iParCode,iLocCode)
		Dim sCheckType
		sQuery = "Select Count(1) From Sal_T_OrdersShipTo Where PartyCode = "&iParCode&" and PartyLocation = "&iLocCode&" "
		Objrs1.Open sQuery,Con
		IF Not Objrs1.EOF Then
			sCheckType = Objrs1(0)
		End IF
		Objrs1.Close

		IF CStr(sCheckType) <> "0" Then
			CheckLoc = "0"
		Else
			sQuery = "Select Count(1) From Sal_T_OCShipTo Where PartyCode = "&iParCode&" and PartyLocation = "&iLocCode&" "
			Objrs1.Open sQuery,Con
			IF Not Objrs1.EOF Then
				sCheckType = Objrs1(0)
			End IF
			Objrs1.Close
			IF CStr(sCheckType) <> "0" Then
				CheckLoc = "0"
			Else
				CheckLoc = "1"
			End IF
		End IF
	End Function
%>
