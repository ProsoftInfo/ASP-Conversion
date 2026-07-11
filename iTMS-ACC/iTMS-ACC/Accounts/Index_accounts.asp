<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	Index_account.asp
	'Module Name				:	Accounts (Index)
	'Author Name				:	Accounts
	'Created On					:	July 15, 2003
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
<Title>Accounts</Title>
<LINK REL="SHORTCUT ICON" HREF="../assets/images/home.gif">
<link REL="stylesheet" href="../assets/styles/Standard.css" type="text/css">
<SCRIPT SRC="../scripts/redirectToPage.js"></SCRIPT>
<Script>
function hideMenu() {
	var i,divlen
	var aReturn= window.frames[1].document.body.getElementsByTagName("DIV");
	//alert(aReturn(1).id);
	//for (i=0;i<aReturn.length;i++)
	//{
	//	if (aReturn[i].id!=null && aReturn[i].id!="")
	//	{
	//		divlen = new String(aReturn[i].style.width).replace('px','');
	//		divlen = Math.round((parseInt(divlen) /0.809));
	//		aReturn[i].style.width=divlen+'px';
	//	}
	//}
	tblMenuHead.deleteRow(0);
	oRow = tblMenuHead.insertRow(0);
	headerCell=oRow.insertCell();
	headerCell.height=10;
	headerCell.width=20;
	headerCell.colSpan=2;

	tblMenuHead.rows[1].cells[0].width=20;
	tblMenuHead.rows[1].cells[0].bgColor = "#cccccc"

	tblBody.rows[0].cells[0].width="10";
	tblBody.rows[0].cells[1].width="100%";

	headerCell.innerHTML="<span style=\"cursor: pointer\"><IMG id=\"imgEC\" onclick=\"AccHome()\" Title=Expand src=\"../assets/images/ExpandButton.gif\"  border=2 width=\"17\" height=\"14\" style=\"border-style: solid; border-color: #999999; background-color: #ffffff;\"></span>";

	Menu.style.visibility="hidden";
}


</Script>

<Script>
function showMenu() {
	var i,divlen
	var aReturn= window.frames[1].document.body.getElementsByTagName("DIV");


//-----------------
	var divarr, divwidth, divid, temp;
	divarr = window.frames[1].document.body.getElementsByTagName('DIV');
	divarrlength = divarr.length;

	for(i=0;i<divarrlength;i++){
		if(divarr[i].id != ""){
			//alert( "index is " + i + " "+ divarr[i].id + " - " + divarr[i].style.width );

			temp=divarr[i].style.width;
			// width of the div	alert(temp.substring(0,temp.length-2));
			divwidth = temp.substring(0,temp.length-2);
			//alert(divwidth);

			if(divwidth > 585)
			{

				//for (i=0;i<aReturn.length;i++)
				//{
				//	if (aReturn[i].id!=null && aReturn[i].id!="")
				//	{
				//		divlen = new String(aReturn[i].style.width).replace('px','');
				//		divlen = Math.round((parseInt(divlen) * 0.809));
				//		aReturn[i].style.width=divlen+'px';
				//	}
				//}
			}
			else if(divwidth < 574)
			{
				//for (i=0;i<aReturn.length;i++)
				//{
				//	if (aReturn[i].id!=null && aReturn[i].id!="")
				//	{
				//		divlen = new String(aReturn[i].style.width).replace('px','');
				//		divlen = Math.round((parseInt(divlen) * 0.809));
				//		aReturn[i].style.width=divlen+'px';
				//	}
				//}
			}
			else
			{
				//for (i=0;i<aReturn.length;i++)
				//{
				//	if (aReturn[i].id!=null && aReturn[i].id!="")
				//	{
				//		divlen = new String(aReturn[i].style.width).replace('px','');
				//		divlen = Math.round((parseInt(divlen) * 1.0));
				//		aReturn[i].style.width=divlen+'px';
				//	}
				//}
			}
		}
	}
//-----------------

	tblMenuHead.deleteRow(0);
	oRow = tblMenuHead.insertRow(0);
	headerCell=oRow.insertCell();

	headerCell.innerHTML="&nbsp;Menu";
	headerCell.className="NavTitleText";
	headerCell.width="50%";

	headerCell=oRow.insertCell();
	headerCell.className="NavTitleImg";
	headerCell.width="50%";
	headerCell.align="right"
	headerCell.innerHTML="<span style=\"cursor: pointer\"><IMG id=\"imgEC\" onclick=\"AccHome()\" Title=Collapse src=\"../assets/images/CollapseButton.gif\"  border=2 width=\"17\" height=\"14\" style=\"border-style: solid; border-color: #999999; background-color: #ffffff;\"></span>";

	tblMenuHead.rows[1].cells[0].bgColor = "#ffffff"

	//alert(tblBody.rows[0].width);
	//alert("Hello");

	tblBody.rows[0].cells[0].width="20%";
	tblBody.rows[0].cells[1].width="80%";



	Menu.style.visibility="visible";
}
</Script>
<script src="../scripts/AccountsIndexCompat.js"></script>

</HEAD>
<BODY class="MainBack" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 onLoad="DispErr()">
	<FORM id="form1id" method="POST" name="form1name">
	<Input type="hidden" name="hccVal" Value="0">
		<table border="1" id="main" cellpadding="0" cellspacing="0" width="100%" height="100%">
			<TR>
				<TD VALIGN="top" align=center height="45" >
					<table border="0" cellpadding="0" cellspacing="0" width="100%" class="MainTitle">
						<tr>
							<td height="40"><b><%=GetOrganization()%></td>
							<td style="border-right: 1 solid #31659C" align="right"><b>Accounts&nbsp;&nbsp;&nbsp;</b></td>
						</tr>
					</table>
				</TD>
			</TR>
			<TR>
				<TD VALIGN="TOP">
					<table border="0" id="tblBody" cellpadding="0" cellspacing="0" width="100%" HEIGHT="100%" class="MainBack">
						<TR>
							<TD VALIGN="TOP" width="20%" class="NavCell">
								<table border="0" cellpadding=0 id="tblMenuHead" class="NavTitle" cellspacing=0  height="100%" width="100%" STYLE="border-bottom:solid black;border-width=1px;" color="#ffffff" bgcolor="#FFFFFF">
									<TR>
										<td class="NavTitleText">&nbsp;Menu</td>
										<td align="right" class="NavTitleImg"><span style="cursor: pointer"><IMG id=imgEC onclick="AccHome()" Title=Collapse src="../assets/images/CollapseButton.gif" border=0 width="17" height="14" style="background-color: #ffffff; border: 2 solid #999999" ></span></td>
									</TR>
									<TR>
										<TD COLSPAN="2" style="left:0px;border-left:solid black ;border-right:solid black ; border-top:solid black;border-width=1px;">
											<DIV ID="Menu" valign="top" style="height: 700px;">
												<IFRAME NAME="menuFrame" ID="IFrame1" FRAMEBORDER="0" SCROLLING="AUTO" SRC="MenuAccounts.asp"  NORESIZE="RESIZE" height="100%" width="100%"></IFRAME>
											</DIV>
										</TD>
									</TR>
								</TABLE>

							</TD>
							<TD  valign="top" class="LeftMenuBorder">
								<table border="0" cellspacing="0" width="100%" class="MainToolBarTable" cellpadding="0">
									<tr>
										<td width="60px" align="center" class="MainToolBarCell">
										<%
													' Function Call to Insert Menu
													'Response.Write InsertMenu("Accmenu")
													'Response.write InsertSplitButtonMenu("Accmenu")
										%>
										<%
											Response.Write InsertMenu(Request.QueryString("AppCode"))
										%>
											<!--table>
												<tr>
													<td>
														<fieldset id="splitbuttonsfrommarkup">
															<input type="submit" id="splitbutton1" name="splitbutton1_button" value="Module">
															<select id="splitbutton1select" name="splitbutton1select" multiple>
																<option value="1">Accounts</option>
																<option value="2">Inventory</option>
																<option value="3">Purchase</option>
																<option value="4">Sales</option>
																<option value="5">Production</option>
															</select>
														</fieldset>
													</td>
												</tr>
											</table-->
										</td>
										<td width="50" align="center" class="MainToolBarCell">Tools</td>
										<td width="40" align="center" class="MainToolBarCell">
										<a style="text-decoration:none;font:color:black" href="#" onclick="Help(); return false;">Help</a>
										</td>
										<td align="center" class="MainToolBarCell"><span style="text-decoration:none;font:bold;color:red"><%=sOrgName%>&nbsp;[&nbsp;<%=Session("FinPeriod")%>&nbsp;]</span></td>
										<td width="135" align="right" class="MainToolBarCell">&nbsp;Login: <span style="text-decoration:none;font:color:black"><%=Session("username")%>&nbsp;</span></td>
										<td align="center" class="MainToolBarCell"><A style="text-decoration:none;font:color:black" HREF="../itmslogout.asp">Logout</A></td>
									</tr>
								</table>
								<IFRAME NAME="bodyFrame" ID="IFrame2" FRAMEBORDER="0" SCROLLING="AUTO" SRC="AccountsHome.asp" NORESIZE="RESIZE" width="100%" height="690px"></IFRAME>
							</TD>
						</TR>
					</TABLE>
				</TD>

			</TR>
		</TABLE>
	</FORM>
</BODY>
</HTML>
