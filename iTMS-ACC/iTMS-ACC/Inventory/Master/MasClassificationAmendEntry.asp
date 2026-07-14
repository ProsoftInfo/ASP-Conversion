<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MasClassificationAmendEntry.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 22, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	MasClassificationNameUpdate.asp
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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Classification Amendment</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" id="Data" data-itms-xml-island="1">
<Root/>
</script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ClassificationAmend.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="" target="body">
<input type=hidden name="pGroup" value="">
<input type=hidden name="pName" value="">
<input type=hidden name="gPath" value="">
<input type=hidden name="hPara" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Master Amendment
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
								<td class="TabCell" valign="bottom" width="75">
								  <span style="cursor: hand">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="MasCategoryAmendEntry.asp">
											<td width="100%" align="center">Category
											</td></a>
										</tr>
									</table>
								  </span>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="50">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="MasUoMAmendEntry.asp">
											<td width="100%" align="center" height="13">UoM
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="100">
								  <span style="cursor: hand">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
								   <tr>
									  <td width="100%" align="center">Classification</td>
									</tr>
								  </table>
								  </span>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
									<table border="0" cellpadding="0" cellspacing="0" width="20" class="TabTableEnd">
										<tr>
											<td width="100%" valign="bottom">
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
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
										<tr>
											<td width="5">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
											<td class="TableOutlineOnly" width="48%">
                                                <IFRAME NAME=main FRAMEBORDER=0 SCROLLING=NO SRC="comClassificationTree.asp" NORESIZE="RESIZE" STYLE="WIDTH=99.5%; HEIGHT=375"></IFRAME>
											</td>
											<td width="5">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
											<td class="TableOutlineOnly" valign=top width="50%" rowspan="2">
                                                <IFRAME NAME=body FRAMEBORDER=0 SCROLLING=NO SRC="MasClassificationNameEntry.asp" NORESIZE="RESIZE" STYLE="WIDTH=100%; HEIGHT=400"></IFRAME>
											</td>
											<td width="5">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
										<tr>
											<td width="5">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
											<td class="TableOutlineOnly">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
												            <p align="center">
												             <input type="button" value="Refresh" class="ActionButton" onClick="Refresh()">
												             <input type="button" value="Amend" class="ActionButton" onClick="Amend('A')">
												             <input type="button" value="Delete" class="ActionButton" onClick="Delete('C')">
														</td>
													</tr>
												</table>
											</td>
											<td width="5">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
                                    </table>
								</td>
							</tr>
                            <tr>
								<td align="center" class="BottomPack" width="100%" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
						</table>
					</td>
					<td align="center">
                        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
				</tr>
				<tr>
					<td class=MiddlePack colspan="2"> </td>
				</tr>
				<tr>
					<td class=ActionCell colspan="2"> <p align="center">
						<input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</html>
