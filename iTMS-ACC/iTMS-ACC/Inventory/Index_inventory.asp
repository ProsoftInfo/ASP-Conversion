<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	Index_inventory.asp
	'Module Name				:	Inventory (Index)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 26, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Jan 05,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!-- #include File="../include/DatabaseConnection.asp" -->
<!-- #include File="../include/GetOrganization.asp" -->
<%
	Dim sOrgName,sOrgID,dcrs,sQuery
	set dcrs = Server.CreateObject("ADODB.Recordset")
	sOrgID = Session("organizationcode")
	sQuery = "Select OrgUnitShortDescription from DCS_OrganizationUnitDefinitions  where OUDefinitionID = '"& sOrgID  & "'"
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sQuery
		.Open
	end with
	if not dcrs.EOF then
		sOrgName = dcrs(0)
	end if
	dcrs.Close
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<Title>Inventory</Title>
<LINK REL="STYLESHEET" HREF="../assets/styles/Standard.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../scripts/checkwidth.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/redirectToPage.js"></SCRIPT>
<script src="../scripts/itms-modern-compat.js"></script>
<script>
function Help() {
	window.open("../Inventory/HelpFiles/InvHelp.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px,left=10,top=10");
}
</script>
</HEAD>
<BODY class="MainBack" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
	<FORM id="form1id" method="POST" name="form1name">
		<table border="0" id="main" cellpadding="0" cellspacing="0" width=100% height=100%>
			<TR>
				<TD VALIGN="top" align=center height="45" >
					<table border="0" cellpadding="0" cellspacing="0" width="100%" class="MainTitle">
						<tr>
							<td height="40"><b><%=GetOrganization()%></td>
							<td align="right"><b>Inventory</b>&nbsp;&nbsp;
							</td>
							<td align="right">
								<img src="../assets/images/Prosoft_Logo.gif" width="35" height="35">&nbsp;
							</td>
						</tr>
					</table>
				</TD>
			</TR>
			<TR>
				<TD VALIGN="TOP">
					<table border="0" id="tblBody" cellpadding="0" cellspacing="0" width=100% HEIGHT="100%" class="MainBack">
						<TR>
							<TD VALIGN="TOP" width="20%" class="NavCell">
								<table border="0" cellpadding=0 id="tblMenuHead" class="NavTitle" cellspacing=0  height="100%" width="100%" STYLE="border-bottom:solid black;border-width=1;" color="#ffffff" bgcolor="#FFFFFF">
									<TR>
										<td class="NavTitleText">&nbsp;Menu</td>
										<td align="right" class="NavTitleImg"><span style="cursor: hand"><IMG id=imgEC onclick="javaScript:Home()" Title=Collapse src="../assets/images/CollapseButton.gif" border=0 width="17" height="14" style="background-color: #ffffff; border: 2 solid #999999" ></span></td>
									</TR>
									<TR>
										<TD COLSPAN="2" STYLE="LEFT:0PX;BORDER-LEFT:SOLID BLACK ;BORDER-RIGHT:SOLID BLACK ; BORDER-TOP:SOLID BLACK;BORDER-WIDTH:1;">
											<DIV ID="Menu" CLASS="MenuTop" valign="top" style="height: 700px;">
												<IFRAME NAME="menuFrame" ID=IFrame1 FRAMEBORDER=0 SCROLLING=AUTO SRC="MenuInventory.asp"  NORESIZE="RESIZE" height="100%" width="100%"></IFRAME>
											</DIV>
										</TD>
									</TR>
								</TABLE>

							</TD>
							<TD  valign="top" class="LeftMenuBorder">
								<table border="0" cellspacing="0" width="100%" class="MainToolBarTable" cellpadding="0">
									<tr>
										<td width="60" align="center" class="MainToolBarCell">
										<%
											' Function Call to Insert Date Picker
											Response.Write InsertMenu(Request.QueryString("AppCode"))
										%>
										</td>
										<td width="50" align="center" class="MainToolBarCell">Tools</td>
										<td width="40" align="center" class="MainToolBarCell"><a style="text-decoration:none;font:color:black" href="#" onclick="Help()">Help</a></td>
										<td align="center" class="MainToolBarCell"><span style="text-decoration:none;font:bold;color:red"><%=sOrgName%>&nbsp;[&nbsp;<%=Session("FinPeriod")%>&nbsp;]</span></td>
										<td width="135" align="right" class="MainToolBarCell">&nbsp;Login: <span style="text-decoration:none;font:color:black"><%=Session("username")%>&nbsp;</span></td>
										<td align="center" class="MainToolBarCell"><a style="text-decoration:none;font:color:black" href="../itmslogout.asp">Logout</a></td>
									</tr>
								</table>
								<IFRAME NAME="bodyFrame" ID=IFrame2 FRAMEBORDER=0 SCROLLING=AUTO SRC="./welcome_Inventory.asp" NORESIZE="RESIZE" width="100%" HEIGHT="690px"></IFRAME>
							</TD>
						</TR>
					</TABLE>
				</TD>

			</TR>
		</TABLE>
	</FORM>

</BODY>
</HTML>
