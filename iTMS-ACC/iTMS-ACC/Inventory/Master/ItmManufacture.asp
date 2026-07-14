<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmManufacture.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	Ragavendran R
	'Created On					:	Feb 20,2013
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
<%
	'XML DOM Variables
	Dim oDOM,Root,objfs,PGNode,rsTemp

	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")

	dim sOrgName,sClassName,sClassCode,iItmCode,sItmDescr,sOrgCode,sSTUoM,sMAUoM,sSAUoM,sPUUoM
	dim schkSal, schkPur, schkMan,sQuery

	iItmCode = Request("ItemCode")
    sClassCode = Request("ClassCode")
    sOrgCode = Session("organizationcode")
    sOrgName = Session("OrgShortName")

if trim(iItmCode)<>"" then
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
end if 'if trim(iItemCode)<>"" then

		sPUUoM = DisplayUoM(sPUUoM)

%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Control Definition - Inventory</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="storageData" data-src="../xmldata/Storage.xml"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="TempData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root ItemCode="<%=iItmCode%>" ClassCode="<%=sClassCode%>" OrgCode="<%=sOrgCode%>"></Root></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmManufacture.js"></SCRIPT>
</HEAD>
<script type="application/xml" data-itms-xml-island="1" id="ItemData" data-src="../Temp/Master/<%=iItmCode%>_DetailedItem.xml"><Root/></script>
<BODY leftMargin=0 topMargin=0>
<form method="POST" name="formname" action="">
<INPUT TYPE=HIDDEN NAME="hClassName" VALUE="<%=sClassName%>">
<INPUT TYPE=HIDDEN NAME="hOrgName" VALUE="<%=sOrgName%>">
<INPUT TYPE=HIDDEN NAME="hItmName" VALUE="<%=sItmDescr%>">
<INPUT TYPE=HIDDEN NAME="hClassCode" VALUE="<%=sClassCode%>">
<INPUT TYPE=HIDDEN NAME="hOrgCode" VALUE="<%=sOrgCode%>">
<INPUT TYPE=HIDDEN NAME="hItmCode" VALUE="<%=iItmCode%>">
<input type="hidden" name="hPartyCode" value="">

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
							    <td class="TabCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="ItemListEntryForCreate.asp">
												<td align="center">List
												</td></a>
											</tr>
										</table>
									</td>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center"><a href="ItmCreationDefinitionEntry.asp">Basic</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmDetailedDefn.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=sClassCode%>">
											<td align="center">Purch. & Sales
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr><a href="ItmInvDet.asp?ItemCode=<%=iItmCode%>&ClassCode=<%=sClassCode%>">
											<td align="center">Inventory
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="120">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
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
								            <td class="FieldCellSub">
								                Customer
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtCustomer" class="formelemread" size="50" readonly="true">
								                <img id="Img1" border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" onclick="SelectCustomer()" align="center" alt="Customer Selection Selection" width="10" height="11" style="cursor:hand">
								            </td>
								            <td class="FieldCellSub">
								                Alias
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtAlias" class="formelem">
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Castting Drawing No
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtCDN" class="formelem">
								            </td>
								            <td class="FieldCellSub">
								                Machining Drawing No
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtMDN" class="formelem">
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Grade
								            </td>
								            <td class="FieldCellSub">
								                <select name="selGrade" class="formelem">
								                    <option value="S">Select</option>
								                </select>
								            </td>
								            <td class="FieldCellSub">
								                Moulding Process
								            </td>
								            <td class="FieldCellSub">
								                <select name="selModelProcess" class="formelem">
								                    <option value="S">Select</option>
								                </select>
								            </td>
								        </tr>
								         <tr>
								            <td class="FieldCellSub">
								                No of Cavities
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtNoofCavities" class="formelem">
								            </td>
								            <td class="FieldCellSub">
								                Match Plate No
								            </td>
								            <td class="FieldCellSub">
								                <select name="selMatchPlateNo" class="formelem">
								                    <option value="0">Select</option>
								                </select>
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Item Weight
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtItemWeight" class="formelem">
								            </td>
								            <td class="FieldCellSub">
								                Base Value
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtBaseValue" class="formelem">
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Item Rate
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtItemRate" class="formelem">
								            </td>
								            <td class="FieldCellSub">
								                Export Rate
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtExportRate" class="formelem">
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Currency
								            </td>
								            <td class="FieldCellSub">
								                 <select name="selCurrency" class="formelem">
								                    <option value="S">Select</option>
								                    <%
								                        sQuery = "Select CurrencyCode,CurrencyShortName from MS_CurrencyMaster"
								                        rsTemp.open sQuery,con
								                        if not rsTemp.eof then
								                            do while not rsTemp.eof
								                                response.write "<option value="&trim(rsTemp(0))&">"&trim(rsTemp(1))&"</option>"
								                                rsTemp.movenext
								                            loop
								                        end if
								                        rsTemp.close
								                    %>
								                </select>
								            </td>
								            <td class="FieldCellSub">
								                Pattern Material
								            </td>
								            <td class="FieldCellSub">
								                 <select name="selPatternMaterial" class="formelem">
								                    <option value="S">Select</option>
								                </select>
								            </td>
								        </tr>
								        
								        <tr>
								            <td class="FieldCellSub">
								                Pattern Owner
								            </td>
								            <td class="FieldCellSub">
								                 <select name="selPatternOwner" class="formelem">
								                    <option value="S">Select</option>
								                </select>
								            </td>
								            <td class="FieldCellSub">
								                Pattern Availability
								            </td>
								            <td class="FieldCellSub">
								                 <select name="selPatternAvailability" class="formelem">
								                    <option value="S">Select</option>
								                </select>
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


