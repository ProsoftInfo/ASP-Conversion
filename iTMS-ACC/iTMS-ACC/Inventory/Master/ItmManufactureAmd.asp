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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
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
    
    if trim(iItmCode)="" then
	%>
	    <script>
	        alert("Please Select the Item in List Tab")
	        window.history.back(-1)
	    </script>
	<%
	end if

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
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmManufacture.js"></SCRIPT>
</HEAD>
<script type="application/xml" data-itms-xml-island="1" id="ItemData" data-src="../Temp/Master/<%=iItmCode%>_DetailedItem.xml"><Root/></script>
<%
Dim rsManu
Dim PartyCode,Alias,CasttingDrawNo,MachineDrawNo,Grade,MouldingProcess,NoofCavities
Dim MatchPlateNo,ItemWeight,BaseValue,ItemRate,ExportRate,sCurrency,PatternMaterial
Dim PatternOwner,PatternAvailability,ItemCode,ClassificationCode,OrganisationCode,PartyName

set rsManu = Server.createObject("ADODB.Recordset")
set rsTemp = Server.CreateObject("ADODB.Recordset")
sQuery = "Select IsNull(PartyCode,0),isNull(Alias,''),IsNull(CasttingDrawNo,''),IsNull(MachineDrawNo,''),IsNull(Grade,''),IsNull(MouldingProcess,''),IsNull(NoofCavities,0),"
sQuery = sQuery & " IsNull(MatchPlateNo,0),IsNull(ItemWeight,0),IsNull(BaseValue,0),IsNull(ItemRate,0),IsNull(ExportRate,0),IsNull(Currency,''),IsNull(PatternMaterial,''),"
sQuery = sQuery & " IsNull(PatternOwner,''),IsNull(PatternAvailability,'') from INV_M_ItemOrgManufacturing where "
sQuery = sQuery & " ItemCode="& iItmCode &" and ClassificationCode="& sClassCode &" and OrganisationCode="& pack(sOrgCode)
rsManu.open sQuery,con
if not rsManu.eof then
    PartyCode = rsManu(0)
    Alias = rsManu(1)
    CasttingDrawNo = rsManu(2)
    MachineDrawNo = rsManu(3)
    Grade = rsManu(4)
    MouldingProcess = rsManu(5)
    NoofCavities = rsManu(6)
    MatchPlateNo = rsManu(7)
    ItemWeight= rsManu(8)
    BaseValue = rsManu(9)
    ItemRate =rsManu(10)
    ExportRate= rsManu(11)
    sCurrency = rsManu(12)
    PatternMaterial = rsManu(13)
    PatternOwner = rsManu(14)
    PatternAvailability = rsManu(15)
end if
rsManu.close
if trim(PartyCode)<>"" then
    sQuery = "Select PartyName from APP_M_PartyMaster where PartyCode = "& PartyCode
    rsManu.open sQuery,con
    if not rsManu.eof then
        PartyName = trim(rsManu(0))
    end if
    rsManu.close
end if 'if trim(PartyCode)<>"" then

%>
<BODY leftMargin=0 topMargin=0>
<form method="POST" name="formname" action="">
<INPUT TYPE=HIDDEN NAME="hClassName" VALUE="<%=sClassName%>">
<INPUT TYPE=HIDDEN NAME="hOrgName" VALUE="<%=sOrgName%>">
<INPUT TYPE=HIDDEN NAME="hItmName" VALUE="<%=sItmDescr%>">
<INPUT TYPE=HIDDEN NAME="hClassCode" VALUE="<%=sClassCode%>">
<INPUT TYPE=HIDDEN NAME="hOrgCode" VALUE="<%=sOrgCode%>">
<INPUT TYPE=HIDDEN NAME="hItmCode" VALUE="<%=iItmCode%>">
<input type="hidden" name="hPartyCode" value="<%=PartyCode%>">

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
											<tr><a href="ItemListEntryForEdit.asp">
												<td align="center">List
												</td></a>
											</tr>
										</table>
									</td>
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
								<td class="TabCurrentCell" valign="bottom" width="145">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Manufacturing
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="145">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="ITEMIMPORTEXPORT.ASP">
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
								                <input type="text" name="txtCustomer" class="formelemread" size="50" readonly="true" value="<%=PartyName%>">
								                <img id="Img1" border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" onclick="SelectCustomer()" align="center" alt="Customer Selection Selection" width="10" height="11" style="cursor:hand">
								            </td>
								            <td class="FieldCellSub">
								                Alias
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtAlias" class="formelem" value="<%=Alias%>">
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Castting Drawing No
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtCDN" class="formelem" value="<%=CasttingDrawNo%>">
								            </td>
								            <td class="FieldCellSub">
								                Machining Drawing No
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtMDN" class="formelem" value="<%=MachineDrawNo%>">
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Grade
								            </td>
								            <td class="FieldCellSub">
								                <select name="selGrade" class="formelem" >
								                    <option value="S" <%if Grade="S" then response.write "Selected" %>>Select</option>
								                </select>
								            </td>
								            <td class="FieldCellSub">
								                Moulding Process
								            </td>
								            <td class="FieldCellSub">
								                <select name="selModelProcess" class="formelem">
								                    <option value="S" <%if MouldingProcess="S" then response.write "Selected" %>>Select</option>
								                </select>
								            </td>
								        </tr>
								         <tr>
								            <td class="FieldCellSub">
								                No of Cavities
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtNoofCavities" class="formelem" value="<%=NoofCavities%>">
								            </td>
								            <td class="FieldCellSub">
								                Match Plate No
								            </td>
								            <td class="FieldCellSub">
								                <select name="selMatchPlateNo" class="formelem">
								                    <option value="0" <%if MatchPlateNo="0" then response.write "Selected" %>>Select</option>
								                </select>
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Item Weight
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtItemWeight" class="formelem" value="<%=ItemWeight%>">
								            </td>
								            <td class="FieldCellSub">
								                Base Value
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtBaseValue" class="formelem" value="<%=BaseValue%>">
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Item Rate
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtItemRate" class="formelem" value="<%=ItemRate%>">
								            </td>
								            <td class="FieldCellSub">
								                Export Rate
								            </td>
								            <td class="FieldCellSub">
								                <input type="text" name="txtExportRate" class="formelem" value="<%=ExportRate%>">
								            </td>
								        </tr>
								        <tr>
								            <td class="FieldCellSub">
								                Currency
								            </td>
								            <td class="FieldCellSub">
								                 <select name="selCurrency" class="formelem">
								                    <option value="S" <%if sCurrency="S" then response.write "Selected" %>>Select</option>
								                    <%
								                        sQuery = "Select IsNull(CurrencyCode,0),CurrencyShortName from MS_CurrencyMaster"
								                        rsTemp.open sQuery,con
								                        if not rsTemp.eof then
								                            do while not rsTemp.eof
								                                if trim(sCurrency) = trim(rsTemp(0)) then
								                                    response.write "<option value="&trim(rsTemp(0))&" selected>"&trim(rsTemp(1))&"</option>"
								                                else
								                                    response.write "<option value="&trim(rsTemp(0))&">"&trim(rsTemp(1))&"</option>"
								                                end if
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
								                    <option value="S" <%if PatternMaterial="S" then response.write "Selected" %>>Select</option>
								                </select>
								            </td>
								        </tr>
								        
								        <tr>
								            <td class="FieldCellSub">
								                Pattern Owner
								            </td>
								            <td class="FieldCellSub">
								                 <select name="selPatternOwner" class="formelem">
								                    <option value="S" <%if PatternOwner="S" then response.write "Selected" %>>Select</option>
								                </select>
								            </td>
								            <td class="FieldCellSub">
								                Pattern Availability
								            </td>
								            <td class="FieldCellSub">
								                 <select name="selPatternAvailability" class="formelem">
								                    <option value="S" <%if PatternAvailability="S" then response.write "Selected" %>>Select</option>
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


