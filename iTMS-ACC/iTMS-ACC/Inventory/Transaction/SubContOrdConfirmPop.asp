<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	SubConProcessDetails.asp
	'Module Name				:	Inventory
	'Author Name				:	Ragavendran R
	'Created On					:	Jun 13,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/purpopulate.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS- Subcontract </TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Root Done="N">
</Root>
</script>
<script type="application/xml" data-itms-xml-island="1" id="Data">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="DRGData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="ItemAddData"><Root></Root></script>

<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
function saveXML()
Dim newElem,SCProcess,MatRecdAs
set rtData = OutData.documentElement
    if document.formname.radSelection(0).checked = true then
	    rtData.setAttribute "Done","Y"
	else
	    rtData.setAttribute "Done","N"
	end if 'if document.formname.radSelection(0).checked = true then
	window.close 
End Function
'-------------------------------------------------------------------------------------------
Function window_onunload()
	Set window.returnvalue= outData.documentElement
End Function
'-------------------------------------------------------------------------------------------
</Script>

<%
Dim objRs,sDrgNo,sItemType,sDesc,OrderFor,sUnit,sSupp,sSql,iClassCode,iItemCode,sSupplier,bFlag
Dim sClassDesc,sItemDesc,saTemp,sDrawVerNo,sSource,sPRNo,sPRNoStr,iItemRecdAs,iSCProcess
Dim indrwgStoreNo,iItemRecdAt,sInstruct,iTempItemCode,sDNo
Set objRs = server.CreateObject("Adodb.recordset")

sUnit = Session("organizationcode")

	Dim oDOM,objFS,RootQuote,oNodTemp,oNodItem
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objFS = Server.CreateObject("Scripting.FileSystemObject")
	sItemType = Request("ItemType")
	OrderFor = "C"
		
%>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<input type="hidden" name="hOrderFor" value="<%=OrderFor%>">
<input type="hidden" name="hDrawNo" value="<%=sDrgNo%>">
<input type="hidden" name="hSource" value="<%=bFlag%>">
<input type="hidden" name="hUnit" value="<%=sUnit%>">
<input type="hidden" name="cmbMatRecdAs" value="0">
<input type="hidden" name="hAddMatAs" value="">
<input type="hidden" name="hAddAttribute" value="">
<input type="hidden" name="hAttribute" value="">


<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
    <tr>
		<td align="center" class="TopPack">
		</td>
    </tr>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		Sub Contract Order
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
								<div id=DivSubSelection>
								<table>
                                    <tr>
                                        <td align="center">
						                </td>
								        <td Class="FieldCell">
								            Do you want to create Sub Contract Order?
								        </td>
						            </tr>
						            <tr>
						                <td align="center">
						                </td>
								        <td Class="FieldCell">
								            <input type=radio name=radSelection value="Y">Now
								        </td>
						            </tr>
						            <tr>
						                <td align="center">
						                </td>
								        <td Class="FieldCell">
								            <input type=radio name=radSelection value="N">Later
								        </td>
						            </tr>
						         </table>
						         </div>
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
                                            <input type="button" value="Done" name="B4" onClick="saveXML()" class="ActionButton" tabindex="3">
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
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</Html>



