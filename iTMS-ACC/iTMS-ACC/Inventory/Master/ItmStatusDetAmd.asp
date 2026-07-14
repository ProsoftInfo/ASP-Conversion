<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmStatusDet.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	Ragavendran R
	'Created On					:	July 16,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:
	'Procedures/Functions Used	:	populateInterUnit,populateStLocation
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
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/ItemDisplay.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Control Definition - Inventory</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" id="storageData" data-itms-xml-island="1" data-src="../xmldata/Storage.xml"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT>
function CheckBack() {
	window.location.href = "ItmInvDetAmd.asp?ItemCode=" + document.formname.hItmCode.value + "&ClassCode=" + document.formname.hClassCode.value;
}

function CheckSubmit() {
	var radios = document.formname.radStatus;
	var status = "";
	for (var i = 0; i < radios.length; i += 1) {
		if (radios[i].checked) {
			status = radios[i].value;
			break;
		}
	}
	document.formname.action = "ItemStatusUpdate.asp?ItemCode=" + document.formname.hItmCode.value + "&ClassCode=" + document.formname.hClassCode.value + "&Status=" + status;
	document.formname.submit();
}
</SCRIPT>
</HEAD>
<%
	'XML DOM Variables
	Dim oDOM,Root,objfs,PGNode,rsTemp

	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")

	dim sOrgName,sClassName,sClassCode,iItmCode,sItmDescr,sOrgCode,sSTUoM,sMAUoM,sSAUoM,sPUUoM
	dim schkSal, schkPur, schkMan,sQuery,sStatus,sItemActive,sItemHold,sDeadStock
	Response.Write "<font color=red>"

	iItmCode = Request.QueryString("ItemCode")
    sClassCode = Request.QueryString("ClassCode")
    sOrgCode = Session("organizationcode")
    sOrgName = Session("OrgShortName")

    sQuery = "Select ItemDescription,(Select GroupName from INV_M_Classification where GroupCode = "&_
             " V.ClassificationCode),StoresUOM,PurchaseUOM,ManufacturingUOM,SalesUOM,PurchaseEligible,"&_
             " ManufactureEligible,SalesEligible from VwItem V where ItemCode="& iItmCode &" and ClassificationCode = "& sClassCode
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        sItmDescr = trim(rsTemp(0))
        sClassName = trim(rsTemp(1))
        sSTUoM = trim(rsTemp(2))
        sPUUoM = trim(rsTemp(3))
		sMAUoM = trim(rsTemp(4))
		sSAUoM = trim(rsTemp(5))
		schkPur = trim(rsTemp(6))
		schkMan = trim(rsTemp(7))
		schkSal = trim(rsTemp(8))
    end if
    rsTemp.Close

		sPUUoM = DisplayUoM(sPUUoM)

		sQuery= "Select ITEMACTIVE,ITEMONHOLD,isNull(DeadStock,'N') from INV_M_ITEMMASTER where ItemCode="& iItmCode &" and ClassificationCode="& sClassCode &" and OrganisationCode = "& Pack(sOrgCode)
		'Response.Write sQuery
		rsTemp.Open sQuery,con
		if not rsTemp.EOF then
		    sItemActive = rsTemp(0)
		    sItemHold = rsTemp(1)
		    sDeadStock = rsTemp(2)
		end if
		rsTemp.Close
		
		if sDeadStock = "Y" then
		    sStatus = "DS"
		else
		    if trim(sItemActive)="Y" and trim(sItemHold) = "0" then
		        sStatus = "AC"
		    elseif trim(sItemActive)="Y" and trim(sItemHold) = "1" then
		        sStatus = "OH"
	        elseif trim(sItemActive)="N" and trim(sItemHold) = "1" then
	            sStatus = "OH"
	        elseif trim(sItemActive)="N" and trim(sItemHold) = "0" then
	            sStatus = "IA"
	        end if
		end if
		
		
		if trim(sStatus)="" then sStatus = "AC"
		
		'Response.Write "sStatus = "& sStatus 

%>
<script type="application/xml" id="ItemData" data-itms-xml-island="1" data-src="../Temp/Master/<%=iItmCode%>_DetailedItem.xml"></script>
<BODY leftMargin=0 topMargin=0>
<form method="POST" name="formname" action="">
<INPUT TYPE=HIDDEN NAME="hClassName" VALUE="<%=sClassName%>">
<INPUT TYPE=HIDDEN NAME="hOrgName" VALUE="<%=sOrgName%>">
<INPUT TYPE=HIDDEN NAME="hItmName" VALUE="<%=sItmDescr%>">
<INPUT TYPE=HIDDEN NAME="hClassCode" VALUE="<%=sClassCode%>">
<INPUT TYPE=HIDDEN NAME="hOrgCode" VALUE="<%=sOrgCode%>">
<INPUT TYPE=HIDDEN NAME="hItmCode" VALUE="<%=iItmCode%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Control Definition
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
                <tr>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmEditEntry.asp?hItemCode=<%=iItmCode%>">
											<td align="center">Basic
											</td></a>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmDetailedDefnAmd.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=sClassCode%>">
											<td align="center">Purch. & Sales
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmInvDetAmd.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=sClassCode%>">
											<td align="center">Inventory
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="120">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Status Control
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmManufactureAmd.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=sClassCode%>">
											<td align="center">Manufacturing
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
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly">
									    <tr>
											<td class="FieldCellSub" width="80">Item Name</td>
											<td>
											<span class="DataOnly"><%=sItmDescr%>&nbsp;</span>
											</td>
											<td class="FieldCell" width="15"></td>
											<td class="FieldCell" width="82">Classification</td>
											<td>
											<span class="DataOnly"><%=sClassName%>&nbsp;</span>
											&nbsp;</td>
											<td></td>
									    </tr>
                                    </table>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                        <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
								    <table border=0 cellpadding=0 cellspacing =0>
								        <tr>
								            <td class="FieldCellSub">Item Status
								            </td>
								            <td class="FieldCellSub">&nbsp;&nbsp;
								                <input type=radio name=radStatus value="AC" <%if trim(sStatus)="AC" then response.write "Checked" %> >Active&nbsp;&nbsp;
								                <input type=radio name=radStatus value="OH" <%if trim(sStatus)="OH" then response.write "Checked" %>>On Hold&nbsp;&nbsp;
								                <input type=radio name=radStatus value="IA" <%if trim(sStatus)="IA" then response.write "Checked" %>>In Active&nbsp;&nbsp;
								                <input type=radio name=radStatus value="NS">Not for Sale&nbsp;&nbsp;
								                <input type=radio name=radStatus value="DS" <%if trim(sStatus)="DS" then response.write "Checked"%>>Dead Stock&nbsp;&nbsp;
								            </td>
								        </tr>
								    </table>
                        		</td>
								<td align="center">
								</td>
                        </tr>
                        <tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Back" name="B5" class="ActionButton" onClick="CheckBack()" >
                                                    <input type="button" value="Save" name="B4" class="ActionButton" onClick="CheckSubmit()" >
													<input type="button" value="Cancel" name="B1" class="ActionButton">
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
</HTML>
<%
	con.close
	set con = nothing
%>

<%
	' Function to populate UoM
	Function DisplayUoM(sUoM)
		' Declaration of variables
		Dim dcrs,sUoMDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = '" & sUoM & "'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMDesc = dcrs(0)

		if Not dcrs.EOF then
			DisplayUoM = sUoMDesc
		else
			DisplayUoM = "N/A"
		end if
		dcrs.Close
	End Function
%>

