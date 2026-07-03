<%@ Language=VBScript %>
<% option explicit %>
<%
	'Program Name				:	Index_admin.asp
	'Module Name				:	Inventory (Admin)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 26, 2003
	'Modified On				:
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
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<Title>Admin Module</Title>
<LINK REL="STYLESHEET" HREF="../assets/styles/Standard.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../scripts/checkwidth.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/redirectToPage.js"></SCRIPT>
</HEAD>
<BODY class="MainBack" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
	<FORM id="form1id" method="POST" name="form1name">
		<table border="0" id="main" cellpadding="0" cellspacing="0" width=100% height=100%>
			<TR>
				<TD VALIGN="top" align=center height="45" >
					<table border="0" cellpadding="0" cellspacing="0" width="100%" class="MainTitle">
						<tr>
							<td height="40"><b><%=GetOrganization()%></td>
							<td style="border-right: 1 solid #31659C" align="right">Admin&nbsp;&nbsp;<span id="timer"></span></td>
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
										<TD COLSPAN="2" STYLE="LEFT:0PX;BORDER-LEFT:SOLID BLACK ;BORDER-RIGHT:SOLID BLACK ; BORDER-TOP:SOLID BLACK;BORDER-WIDTH=1;">
											<DIV ID="Menu" CLASS="MenuTop" valign="top" style="height: 700px;">
												<IFRAME NAME="menuFrame" ID=IFrame1 FRAMEBORDER=0 SCROLLING=AUTO SRC="Menuadmin.html"  NORESIZE="RESIZE" height="100%" width="100%"></IFRAME>
											</DIV>
										</TD>
									</TR>
								</TABLE>

							</TD>
							<TD  valign="top" class="LeftMenuBorder">
								<table border="0" cellspacing="0" width="100%" class="MainToolBarTable" cellpadding="0">
									<tr>
										<td width="75" align="center" class="MainToolBarCell">
										<%
													' Function Call to Insert Menu
													Response.Write InsertMenu("Accmenu")
										%>
										</td>
										<td width="75" align="center" class="MainToolBarCell">Tools</td>
										<td width="75" align="center" class="MainToolBarCell">Help</td>
										<td class="MainToolBarCell" align="center">Login: <span style="text-decoration:none;font:color:black"><%=Session("username")%></span></td>
										<td class="MainToolBarCell" align="center"><span style="text-decoration:none;font:bold;color:red">&nbsp;&nbsp;Financial Year - <%=Session("FinPeriod")%>&nbsp;</span></td>
										<td width="75" align="center" class="MainToolBarCell"><A style="text-decoration:none;font:color:black" HREF="../itmslogout.asp" ALT="Logout">Logout</A></td>
									</tr>
								</table>
								<IFRAME NAME="bodyFrame" ID=IFrame2 FRAMEBORDER=0 SCROLLING=AUTO SRC="./welcome_admin.asp" NORESIZE="RESIZE" width="100%" HEIGHT="95%"></IFRAME>
							</TD>
						</TR>
					</TABLE>
				</TD>

			</TR>
		</TABLE>
	</FORM>

</BODY>
</HTML>
