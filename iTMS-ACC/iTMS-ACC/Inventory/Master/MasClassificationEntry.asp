<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MasClassificationEntry.asp
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 16, 2002
	'Modified By                :   Ragavendran R
	'Modified On				:   Feb 01,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	MasClassificationNameEntry.asp
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
<HTML><HEAD><TITLE>Classification</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data">
<Root/>
</script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ClassificationCreate.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ClassificationAmend.js"></SCRIPT>
<style type="text/css">
html,
body {
	min-height: 100%;
}
.classification-tree-frame {
	width: 99.5%;
	height: calc(100vh - 190px);
	min-height: 375px;
}
.classification-detail-frame {
	width: 100%;
	height: calc(100vh - 165px);
	min-height: 400px;
}
</style>
<script language="javascript">
function Help() {
	window.open("../HelpFiles/ItemCatalog.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px,left=10,top=10");
	return false;
}
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="" target="body">
<input type=hidden name="pGroup" value="">
<input type=hidden name="pName" value="">
<input type=hidden name="gPath" value="">
<input type=hidden name="hPara" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td>
	         <table>
	            <tr>
	                <td class="PageTitle" >
	                    Master Creation
	                </td>
	                <td class="PageTitle" >
	                    <a style="text-decoration:none;font:color:black" href="#" onclick="Help()">Help</a>
	                </td>
	            </tr>
	        </table>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<!--<td class="TabCell" valign="bottom" width="75">
								  <span style="cursor: hand">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="MasCategoryEntry.asp">
											<td width="100%" align="center">Category
											</td></a>
										</tr>
									</table>
								  </span>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="50">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="MasUoMEntry.asp">
											<td width="100%" align="center" height="13">UoM
											</td></a>
										</tr>
									</table>
								</td>-->
								<td class="TabCurrentCell" valign="bottom" align="center" width="100">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
									   <tr>
										  <td width="100%" align="center">Classification</td>
										</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
									<table border="0" cellpadding="0" cellspacing="0" width="20" class="TabTableEnd">
										<tr>
											<td width="100%" valign="bottom">
												<font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class="TabBody">
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
                                                <IFRAME NAME="main" FRAMEBORDER="0" SCROLLING="NO" SRC="comClassificationTree.asp" NORESIZE="RESIZE" class="classification-tree-frame"></IFRAME>
											</td>
											<td width="5">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
											<td class="TableOutlineOnly" valign=top width="50%" rowspan="2">
                                                <IFRAME NAME="body" FRAMEBORDER="0" SCROLLING="NO" SRC="MasClassificationNameEntry.asp" NORESIZE="RESIZE" class="classification-detail-frame"></IFRAME>
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
												             <input type="button" value="Refresh" class="ActionButton" onClick="Refresh()">
												             <input type="button" value="New Group" class="ActionButton" onClick="NewGroup()">
												             <input type="button" value="Amend" class="ActionButton" onClick="Amend('N')" id=button1 name=button1>
												             <input type="button" value="Delete" class="ActionButton" onClick="Delete('C')" id=button2 name=button2>
												             <!--input type="button" value="Attributes" class="ActionButton" onClick="Attributes()"-->
														</td>
													</tr>
												</table>
											</td>
											<td width="5">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
					<td class=ActionCell colspan="2">
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
