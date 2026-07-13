
<%@ Language="VBScript" %>
<% option explicit %>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl="no-cache"
%>
<%
	'Program Name				:	ITEMIMPORTEXPORT.ASP
	'Module Name				:	INVENTORY
	'Author Name				:	Kalaiselvi R
	'Created On					:	July 01,2011
	'Modified By				:
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
    Dim rsObj
    Dim sOrgID,sOrgName,sCreatedBy,sPackingName,sQuery,sNumberingType
    Dim iSlNo
    set rsObj = Server.CreateObject("ADODB.Recordset")
    sOrgID = Session("organizationcode")
    sOrgName = Session("OrgShortName")
	sCreatedBy = Session("userid")
	sPackingName = Request.QueryString("PackName")
	if trim(Request.QueryString)="" then
	    sPackingName = Request.Form("hPackName")
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Paking Types</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<script>
function ImportData(sModule) {
	document.formname.action = "DataImportModuleWiseInsert.asp?Module=" + sModule;
	document.formname.submit();
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" >
	<form method="POST" name="formname" action="">

		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr><td height="1px"></td></tr>
			<tr>
				<td class="PageTitle">
				Item Data Import / Export
				</td>
			</tr>
			<tr>
				<td align="center" class="TopPack">
				</td>
			</tr>
			<tr>
				<td valign="top">
					<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
							<td class="TabCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="ItemListEntryForEdit.asp">
												<td align="center">List
												</td></a>
											</tr>
										</table>
									</td>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmEditEntry.asp">
											<td align="center">Basic
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmDetailedDefnAmd.asp">
											<td align="center">Purch. & Sales
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmInvDetAmd.asp">
											<td align="center">Inventory
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
								    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
									    <tr><a href="ItmManufactureAmd.ASP">
										    <td align="center">Manufacturing
										    </td>
									    </tr>
								    </table>
							    </td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="120">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Import/Export Item
											</td>
										</tr>
									</table>
								</td>

								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
                </tr>
						<tr>
							<td class="TabBody">
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td align="center" colspan="3" class="MiddlePack" height="7">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
										</td>
									</tr>

						            <tr>
								        <td align="center">
								        </td>
								        <td valign="top" width="100%" align="left">
                                            <table border="0" cellpadding="0" cellspacing="0" width=100%>
                                                <tr>
											        <td>

									                </td>
									            </tr>
									            <tr>
							                        <td>
                                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
					                                    <tr>
					                                    <td valign="middle" class="ActionCell">
						                                    <input type="button" value="Inventory_Module"	name="btnInventory"		class="ActionButtonX" onclick="ImportData('INV')">
						                                    <input type="button" value="Purchase_Module"	name="btnPurchase"		class="ActionButtonX" onclick="ImportData('PUR')">
						                                    <input type="button" value="Production_Module"	name="btnProduction"	class="ActionButtonX" onclick="ImportData('PRD')">
						                                    <input type="button" value="Sales_Module"		name="btnSales"			class="ActionButtonX" onclick="ImportData('SAL')" disabled>
					                                    </td>
					                                    </tr>
					                                    </table>
					                                </td>
                                                 </tr>
									        </table>
								        </td>
								        <td align="center"></td>
                                    </tr>
                            		<tr>
										<td align="center" class="MiddlePack" colspan="3"></td>
									</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
			</tr>
		</table>
	</form>
</body>
</html>
