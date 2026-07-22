<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AttributeWiseLotDetailSelection.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	UmaMaheswari S
	'Created On					:	June 08, 2011
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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Storage Lot Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/attributeWiseLotDetailSelection.js"></SCRIPT>
<%
Dim nOptValue,sItemTypeID
nOptValue = Request.QueryString("Data")
sItemTypeID = Request.QueryString("sType")
'Response.Write "<p><Font color=red>data="&nOptValue
%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="FnInit()">

<form method="POST" name="formname" action="">
<input type="hidden" name="hOptValue" value="<%=nOptValue%>">
<input type="hidden" name="hIType" value="<%=sItemTypeID%>">
<input type="hidden" name="hCnt" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot Details
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
								<td>
									<table border="0" cellpadding="0" cellspacing="0" width="100%">

									</table>
								</td>
								<td >
									<table border="0" cellpadding="0" cellspacing="0" width="100%" >

									</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    <p align="center"><font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font>
								</td>
							</tr>
						</table>
					</td>
				</tr >
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" class="MiddlePack">
                                    <table border="0" cellspacing="1" Id ="tblBin" name="tblBin" class="ExcelTable" width="350"></table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
													<!--<input type="button" value="Add New" name="B4" class="ActionButtonX" onClick="AddNew()">-->
                                                    <input type="button" value="Done" name="B2" class="ActionButton" onClick="CheckSubmit()">
													<!--<input type="button" value="Close" name="B3" class="ActionButton" onClick="window.close()">-->
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" colspan="3" class="BottomPack">
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
