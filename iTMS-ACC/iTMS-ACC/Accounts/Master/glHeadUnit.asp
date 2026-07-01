<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	glHeadUnit.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 10, 2002
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
<!--#include file="../../include/accpopulate.asp"-->
<%
dim sGlHeadName,objRs
'XML DOM Variables
Dim oDOM,nodHeader,Root,nodBook,nodUnit,nodTemp,sGroupName,bSubLedger

Set objRs = Server.CreateObject("ADODB.RecordSet")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.Load server.MapPath("../temp/master/glEntry_"&Session.SessionID&".xml")

Set Root = oDOM.documentElement
For Each nodHeader In Root.childNodes

	if StrComp(nodHeader.nodeName,"Description") = 0 then
		sGlHeadName=nodHeader.text
	end if
	if StrComp(nodHeader.nodeName,"GroupCode") = 0 then
		sGroupName=nodHeader.Attributes.Item(0).nodeValue
	end if
	if StrComp(nodHeader.nodeName,"SubLedger") = 0 then
		bSubLedger=nodHeader.Attributes.Item(0).nodeValue
	end if
	if StrComp(nodHeader.nodeName,"Units") = 0 then
		set nodUnit=nodHeader
	end if
next

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript">
window.__itmsPopupCompat = { type: "glHeadUnit" };
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="glHeadUnitUpdate.asp">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">GL Account Head Creation
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
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Main
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="95">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Cost Center
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="120">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Analytical Head</td>
										</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Books</td>
										</tr>
								  </table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
									<tr>
										<td align="center">Unit
										</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="left">
									<table border="0" cellpadding="0" cellspacing="0" class="TabTableEnd">
										<tr>
											<td valign="bottom">
												<p align="center"><font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font></p>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" rowspan="6" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="145">GL Account Head Name</td>
                            <td>
                            <span class="DataOnly"><%=sGlHeadName%>  </span>
                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="145" valign="top">GL Classification</td>
                            <td><span class="DataOnly"><%=sGroupName%></span>
                            </td>
                                </tr>
                                  <tr>
                            <td class="FieldCell" width="145" valign="top">Opening MonthYear</td>
                            <td>

                            <input type="text" value="<%=getFromFinYear%>" name="txtOpenYear" maxlength="6" size="7" class="FormElem">

                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" rowspan="6" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td valign="top" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td valign="top" align="center">
												<DIV class=frmBody id=frm1 style="width: 585; height:240;">
<%if bSubLedger="1" then%>
                                                <table border="0" cellspacing="1" class="ExcelTable" width="569">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" valign="Top" rowspan="2" width="25"><p align="center">S.No.</td>
                                        <td class="ExcelHeaderCell" align="left" colspan="2" valign="Top">&nbsp;&nbsp;&nbsp;Unit Name</td>
                                        </tr>
                                        <tr>
                                        <td class="ExcelHeaderCell" width="150" align="center">Opening Balance</td>
                                        <td class="ExcelHeaderCell" align="center"> Party Sub Type</td>

                                            </tr>
<%
dim sQuery
dim sUnitId,sUnitName,iSno,bFirst,iParTypeCount
iSno=0
	For Each nodTemp In nodUnit.childNodes
		iSno=iSno+1
		sUnitId=trim(nodTemp.Attributes.Item(0).nodeValue)
		sUnitName=nodTemp.Attributes.Item(1).nodeValue
		iParTypeCount=0
			sQuery="select PartyType, PartySubType, SubTypeShortName FROM APP_M_PartyTypes where "&_
						"PartyType+ltrim(str(PartySubType)) NOT in (select PartyType+ltrim(str(PartySubType))"&_
						" from Acc_R_OrgPartyType where OUDefinitionID='"&sUnitId&"')"

			'Response.Write sQuery & "<br><br>"


			with objRs
				.CursorLocation =3
				.CursorType =3
				.Source =sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs.ActiveConnection=nothing
			iParTypeCount=objRs.RecordCount
%>
									<tr>
                                        <td class="ExcelSerial"  width="25" align="center" rowspan="2" valign="top"><%=iSno%></td>
                                        <td class="ExcelDisplayCell" colspan="2" valign="top"><b><%=sUnitName%></b></td>
                                        </tr>
                                        <tr>
                                        <td class="ExcelFieldCell" width="170" valign="top">
                                        <table width="100%"><tr>
                                        <td class="ExcelFieldCell"><input type="text" value="0"  name="txtOpenBal<%=Trim(sUnitId)%>" maxlength="13" size="15" class="FormElem" style="text-align: Right"></td>
                                        <td class="ExcelFieldCell"><input type="radio" value="D" name="optOpenCD<%=sUnitId%>" class="FormElem"> Dr </td>
                                        <td class="ExcelFieldCell"><input type="radio" value="C" name="optOpenCD<%=sUnitId%>" checked class="FormElem"> Cr</td>
                                        </tr></table>
                                        </td>
										<td class="ExcelDisplayCell" valign="top">
	<%
			do while not objRs.EOF
	%>
									<input type="checkbox" name="selPartyType"
									value="<%=sUnitId&"?"& objrs(0) &"?" & objrs(1)%>"> <%=objrs(0) &"-"& objrs(2)%><br>
	<%
				objRs.MoveNext
			loop
			objRs.Close
	%>
	</td>

                                       </tr>

<%
	next

%>



                                                </table>
<%else%>
                                                <table border="0" cellspacing="1" class="ExcelTable" width="569">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" valign="Top" width="25"><p align="center">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center"  valign="Top">Unit Name</td>
                                        <td class="ExcelHeaderCell" width="150" align="center">Opening Balance</td>
										</tr>
<%

iSno=0
	For Each nodTemp In nodUnit.childNodes
		iSno=iSno+1
		sUnitId=trim(nodTemp.Attributes.Item(0).nodeValue)
		sUnitName=nodTemp.Attributes.Item(1).nodeValue
%>
									<tr>
                                        <td class="ExcelSerial"  width="25" align="center"  valign="top"><%=iSno%></td>
                                        <td class="ExcelDisplayCell"  valign="top"><%=sUnitName%></td>
                                        <td class="ExcelFieldCell" width="170" valign="top">
                                        <table width="100%"><tr>
                                        <td class="ExcelFieldCell"><input type="text" value="0"  name="txtOpenBal<%=Trim(sUnitId)%>" maxlength="13" size="15" class="FormElem" style="text-align: Right"></td>
                                        <td class="ExcelFieldCell"><input type="radio" value="D" name="optOpenCD<%=sUnitId%>" class="FormElem"> Dr </td>
                                        <td class="ExcelFieldCell"><input type="radio" value="C" name="optOpenCD<%=sUnitId%>" checked class="FormElem"> Cr</td>
                                        </tr></table>

                                       </tr>

<%
	next
%>                                              </table>

<%end if%>
</div>
								</td>
							</tr>
							<tr>
								<td valign="top" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td valign="top" class="BottomPack">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Done" name="B2" class="ActionButton" tabindex="3" onClick="PageSubmit()">
																<input type="reset" value="Reset" name="B1" class="ActionButton" tabindex="4" >
														</td>
													</tr>
												</table>
								</td>
							</tr>
							<tr>
								<td valign="top" class="BottomPack">
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
