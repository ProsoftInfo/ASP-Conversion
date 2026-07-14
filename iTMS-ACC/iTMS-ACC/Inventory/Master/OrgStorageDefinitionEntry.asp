<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgStorageDefinitionEntry.asp
	'Module Name				:	Inventory (Storage Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 14, 2002
	'Modified By				:	Ragavendran R
	'Modified On				:	Jan 07,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	OrgStorageDefinitionInsert.asp
	'Procedures/Functions Used	:	populateUnit
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
<!--#include virtual="/Inventory/Master/NewStorageXML.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Storage Location</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../xmlData/Storage.xml"%>"></script>
<script type="application/xml" data-itms-xml-island="1" id="NewData" data-src="<%="../Temp/Master/StorageNew"&Session.SessionID&".xml"%>"></script>
<script type="application/xml" data-itms-xml-island="1" id="TempData"></script>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/orgStorageCreate.js"></SCRIPT>

<%
Dim oDOM,Root,Root1,sExp,oDOM1,objfs,Elem,Elem1,OrgNode ,Rt,rsObj
Dim sOrgID,sQry ,iCtr ,sFlag,sOrgName,sQuery
dim sLocCode,sLocName,sAppFor,iTyFree,iTyBin,iUseArea,iNoOfBins
dim iBinNo,sBinCode,sBinName,sBinArea,BinElem
Set rs = Server.CreateObject("ADODB.RecordSet")
Set objfs = CreateObject("Scripting.FileSystemObject")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set oDOM1 =  Server.CreateObject("Microsoft.XMLDOM")
set rsObj = Server.CreateObject("ADODB.Recordset")

dim iLocNo,sTemp,sPara

sOrgID = Session("organizationcode")
sQuery = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = "& sOrgID
rsObj.Open sQuery,con
if not rsObj.EOF then
	sOrgName = trim(rsObj(0))
end if
rsObj.Close


sPara = Request("hPara")
' Response.Write "iLocNo="&sPara
 	if objfs.FileExists(server.MapPath("../xmlData/Storage.xml")) then
		oDOM.load  server.MapPath("../xmlData/Storage.xml")
	end if
	Set Root = oDOM.documentElement
	sExp = "//Organization"
	Set OrgNode = Root.SelectNodes(sExp)
	IF OrgNode.length <> 0 then
		'sOrgID = OrgNode.item(0).Attributes.getNamedItem("OUDEFINITIONID").value
	End IF
	Set Root1 = oDOM1.documentElement
	Set Rt = oDOM1.createElement("Root")
	oDOM1.appendchild Rt

	set Elem = oDOM1.createElement("Organization")
	Elem.setAttribute "OUDEFINITIONID",sOrgID

	Rt.appendchild Elem

	IF trim(sPara) <> "" then 'Amend-------------
		sTemp = split(sPara,":")
		'sOrgID = sTemp(0)
		iLocNo = sTemp(0)
		sQry = "Select Distinct LocationNumber from Inv_M_ItemStorage where LocationNumber = "&iLocNo&" "
		'Response.Write sQry  &"<BR>"
		rs.Open sQry,con
		If not rs.EOF then
			sFlag = True
		Else
			sFlag = False
		End IF
		rs.close
		sFlag = False
		'Response.Write sFlag
		sQry  = "Select LocationCode,LocationName,ApplicableFor,StorageTypeFree,StorageTypeBins,UsableFreeArea,NumberOfBins "&_
				"From Inv_M_Storage where OUDefinitionID = '"& sOrgID &"' and LocationNumber = "& iLocNo &" "
		'Response.Write sQry
		rs.Open sQry,con

		If not rs.EOF then

			sLocCode  = rs(0)
			sLocName  = rs(1)
			sAppFor   = rs(2)
			iTyFree   = rs(3)
			iTyBin    = rs(4)
			iUseArea  = rs(5)
			iNoOfBins = rs(6)
		End If
		rs.Close

		set Elem1 = oDOM1.createElement("Storage")
		Elem1.setAttribute "LOCATIONNUMBER",iLocNo
		Elem1.setAttribute "LOCATIONCODE",sLocCode
		Elem1.setAttribute "LOCATIONNAME",sLocName
		Elem1.setAttribute "APPLICABLEFOR",sAppFor
		Elem1.setAttribute "STORAGETYPEFREE",iTyFree
		Elem1.setAttribute "STORAGETYPEBINS",iTyBin
		Elem1.setAttribute "USABLEFREEAREA",iUseArea
		Elem1.setAttribute "NUMBEROFBINS",iNoOfBins
		Elem.appendchild Elem1

		'Bin Details...
		IF trim(iTyBin) <> "0" and trim(iTyFree) = "0" then
			sQry = "Select BinNumber,BinCode,BinName,BinArea from Inv_M_StoreBinDetails where  "&_
					"OUDefinitionID = '"& sOrgID &"' and LocationNumber = "& iLocNo &" "
		'Response.Write sQry
			rs.Open sQry,con
			iCtr = 1
			do while not rs.EOF
				iBinNo	  = rs(0)
				sBinCode  = rs(1)
				sBinName  = rs(2)
				sBinArea   = rs(3)

				set BinElem	 = oDOM1.createElement("Bin")
				BinElem.setAttribute "SLNO",iCtr
				BinElem.setAttribute "BINNUMBER",iBinNo
				BinElem.setAttribute "BINCODE",sBinCode
				BinElem.setAttribute "BINNAME",sBinName
				BinElem.setAttribute "BINAREA",sBinArea
				Elem1.appendchild BinElem
				iCtr = iCtr + 1
				rs.MoveNext
			loop
			rs.Close

		End If

	End IF

oDOM1.save server.MapPath("../Temp/Master/StorageNew"&Session.SessionID&".xml")

%>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="FnInit()">

<form method="POST" name="formname" action="">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hPara" value="<%=sPara%>">
<input type="hidden" name="hFlag" value="<%=sFlag%>">
<input type="hidden" name="hLocNo" value="<%=iLocNo%>">
<input type="hidden" name="hOrgID" value="<%=sOrgID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Storage Location Creation
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
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
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
									<table cellpadding="0" cellspacing="0">
										<!--<tr>
											<td class=FieldCell> Organization Unit</td>
											<td class='FieldCellSub'>
												<select size="1" name="selOrgUnit" class="FormElem">

												<%	'Calling the Function which populates the Units list
													populateUnit
												%>
												</select>
                                            </td>
										</tr>-->
										<tr>
											<td class=FieldCell> Location Name</td>
											<td class='FieldCellSub'><input type="text" name="txtLocationName" size="45" maxlength=40 class="Formelem"></td>
										</tr>
										<tr>
											<td class=FieldCell> Location Code</td>
											<td class='FieldCellSub'><input type="text" name="txtLocationCode" size="15" maxlength=10 class="Formelem"></td>
										</tr>
									</table>
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
                                        <tr>
											<td align="center" class="MiddlePack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
										<tr>
											<td class="FieldCell" width="100%">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="95"><p align="center">Applicable
                                                              For
                                                            </td>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable>
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td class=MiddlePack colspan="4"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub>
                                                            <input type="radio" value="PU" name="App" class="FormElem" > Purchase</td>
															<td class='FieldCellSub'>
                                                            <input type="radio" value="OI" name="App" class="FormElem"> Inspection-Outorder</td>
															<td class='FieldCellSub'>
                                                            <input type="radio" value="POI" name="App" class="FormElem"> Inspection-Preorder</td>
															<td class='FieldCellSub'>
                                                            <input type="radio" value="SA" name="App" class="FormElem"> Sales</td>
														</tr>
														<tr>
															<td class=FieldCellSub>
                                                            <input type="radio" value="IN" name="App" class="FormElem"> Inventory</td>
															<td class='FieldCellSub'>
                                                            <input type="radio" value="PI" name="App" class="FormElem"> Inspection Process</td>
															<td class='FieldCellSub'>
                                                            <input type="radio" value="PSI" name="App" class="FormElem"> Post-Sale</td>
															<td class='FieldCellSub'>
                                                            <input type="radio" value="MA" name="App" class="FormElem"> Manufacturing</td>
														</tr>
													</table>
                                                            </td>
														</tr>
													</table>
											</td>
										</tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%">
											</td>
										</tr>
										<tr>
											<td width="100%">
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td colspan="2">
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="90"><p align="center">Storage
                                                              Type
                                                            </td>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
                                                    <tr>
															<td class=GroupTable>
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td class=MiddlePack width="100%"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub width="100%">
                                                            <input type="radio" name="ST" value="F" onClick="document.formname.txtUsable.value = '';SetBinEnable('F')" class="FormElem"   > Free / Open Area</td>
														</tr>
														<tr>
															<td class=FieldCellSub width="100%"> &nbsp;Usable Free Area&nbsp; <input type="text" name="txtUsable" size="11" maxlength=10 class="Formelem"> Sq. Ft.</td>
														</tr>
													</table>
                                                            </td>
															<td class=GroupTableCell>
													<table cellpadding="0" cellspacing="0">
														<tr>
															<td class=MiddlePack width="100%"> </td>
														</tr>
														<tr>
															<td class=FieldCellSub width="100%">
                                                            <input type="radio" name="ST" value="B" class="FormElem"  onClick="SetBinEnable('B')"> Specific Location / Bins</td>
														</tr>
														<tr>
															<td class=FieldCellSub width="100%"> &nbsp;Number of Bins&nbsp; <input type="text" name="txtBins" size="3" maxlength=2 class="Formelem" disabled >&nbsp;
															<img id="Img2" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" onclick="popBinSelect()" style="cursor: hand" disabled></td>
														</tr>
													</table>
                                                            </td>
                                                    </tr>
													</table>
											</td>
										</tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<tr>
											<td width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Save" name="B1" class="ActionButton" onClick="XmlUpDate();JavaScript:checkSubmit()">
																<input type="reset" value="Reset" name="B1" class="ActionButton">
																<input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
														</td>
													</tr>
												</table>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="BottomPack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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

