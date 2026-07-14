<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PackingType.asp
	'Module Name				:	INVENTORY (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 21, 2011
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
Dim objDOM,Root,newElem,objRs,objRs1,objRs2,iSno,sType,iCtr
dim sOrgId,sQuery,sTemp,nPackCode
Dim sGrossPackLabel,sTarePack,sTarePackLabel

sTemp = Split(Trim(Request.QueryString("Type")),":")
sType = sTemp(0)
nPackCode = sTemp(1)

set objRs   = server.CreateObject("adodb.recordset")
set objRs1  = server.CreateObject("adodb.recordset")
set objDOM  = server.CreateObject("Microsoft.XMLDOM")

set Root = objDOM.CreateElement("Root")
objDOM.appendchild Root

If sType = "E" Then
	
	iCtr = 1
	
	sQuery = " Select PackingCode,IsNull(PackingShortName,''),isNull(PackingName,''),isNull(AlternateName,''),isNull(NumberingType,''),"&_
			 " isNull(ManualLotNumbering,''),isNull(ManualSerialNumbering,''),isNull(SerialWithinLot,''),isNull(NoOfSubLevels,0),isNull(Enforce,'E'),IsNull(GrossPerPackLabel,'Gross/Pack'),IsNull(TarePerPack,'N'),IsNull(TarePerPackLabel,'Tare/Pack') "&_
			 " From APP_M_PackingType where PackingCode="& nPackCode &" "
	
	objrs.Open sQuery,con
	
	If Not objrs.EOF Then
		Root.setAttribute("Type"),sType
		Root.setAttribute("PackCode"),nPackCode
		Root.setAttribute("ShortName"),objRs(1)
		Root.setAttribute("Name"),objRs(2)
		Root.setAttribute("AltLabel"),objRs(3)
		Root.setAttribute("ShowBothChk"),""
		Root.setAttribute("ReceiptNumbering"),objRs(4)
	
		Root.setAttribute("LotNoSelection"),objRs(5)
		Root.setAttribute("LotNoEnforceCheck"),objRs(9)
	
		Root.setAttribute("SerialNoSelection"),objRs(6)
		Root.setAttribute("SerialNoWithinLotCheck"),objRs(7)
		Root.setAttribute("NoOfSubLevel"),objRs(8)
		Root.setAttribute "GrossLabel",objrs(10)
        Root.setAttribute "Tare",objrs(11)
        Root.setAttribute "TareLabel",objrs(12)
		
		sQuery = "Select SubLevelId,isNull(SubLevelName,'') From APP_M_PackingTypeSubLevel where PackingCode="& nPackCode&" "
		objRs1.Open sQuery,con
	
		Do While Not objRs1.EOF 
			
			set newElem = objDOM.createElement("SubLevelDetails")
			newElem.setAttribute("LevelNo"),objrs1(0)
			newElem.setAttribute("LevelLabel"),objrs1(1)
			Root.appendchild newElem
			
			iCtr = iCtr + 1
			
			objRs1.MoveNext 
		Loop
		
	End IF
	objRs.Close 
	
	objDOM.save Server.MapPath("../Temp/Master/PackingType"&Session.SessionID&".XML")
		 
End IF	'If sType = "E" Then

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%If sType = "C" Then%>
	<script type="application/xml" data-itms-xml-island="1" id="PackingData"><Root/></script>
<%Else%>
	<script type="application/xml" data-itms-xml-island="1" id="PackingData" data-src="<%="../Temp/Master/PackingType"&Session.SessionID&".XML"%>"><Root/></script>
<%End IF%>
<script type="application/xml" data-itms-xml-island="1" id="RetData"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/packingTypeEntry.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="ShowData('<%=sType%>')">
<form method="POST" name="formname" action="">
	
	<Input type="hidden" name="hType" value="<%=sType%>">
	<Input type="hidden" name="hPackCode" value="<%=nPackCode%>">
	
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Packing Type
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>

							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
                                &nbsp;
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" class="TableOutlineOnly" width="100%">
                                <tr>
                                    <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Short Name </td>
									<td class="FieldCellSub">
										<Input type="textbox" Name="txtShortName" class="FormElem" value="">
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Name</td>
									<td class="FieldCellSub">
										<Input type="textbox" Name="txtName" class="FormElem" value="">
									</td>
									<td class="FieldCellsub">Alternate Label</td>
									<td class="FieldCellSub">
										<Input type="textbox" Name="txtLabel" class="FormElem" value="">
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Show Both</td>
									<td class="FieldCellSub">
										<Input type="Checkbox" Name="chkShowBoth">
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Receipt Numbering</td>
									<td class="FieldCellSub" colspan="3">
										<Input type="Radio" Name="RadRecNo" value="N" checked>None
										<Input type="Radio" Name="RadRecNo" value="L">Lot
										<Input type="Radio" Name="RadRecNo" value="LS">Lot & Serial
										<Input type="Radio" Name="RadRecNo" value="S">Serial
									</td>
                                </tr>
                                <tr>
                                     <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                            
                            <tr>
								<td></td>
								<td style="width:100%">
									<Table border="0" cellspacing="1" class="TableOutlineOnly" width="100%">
										<tr>
											<td class=ExcelDisplaycell align="Left" width=50%><b>Number Generation</b></td>
											<td class=ExcelDisplaycell align="Left" colspan="2"><b>Number Of Sublevels</b> &nbsp;&nbsp;
											<Input type="text" name="txtNoOfSubLevel" class="FormElem" size="4" onChange="AddSubLevel()" ></td>
										</tr>
										<tr>
											<td width="50%">
												<table border="0" cellspacing="1" class="Exceltable" width="100%">
													
													<tr>
														<td class="ExcelHeaderCell" align="center">Lot No</td>
														<td class="ExcelHeaderCell" align="center">Serial No</td>
													</tr>
													<tr>
														<td class="ExcelDisplaycell" align="Left">
															<Input type="Radio" name="RadLotNo" Value="M" Checked>Manual<br>
															<Input type="Radio" name="RadLotNo" Value="A">Auto<br>
															<!--<Input type="Checkbox" name="chkLotNo" Value="E">Enforce-->
														</td>
										
														<td class="ExcelDisplaycell" align="LEft">
															<Input type="Radio" name="RadSerNo" Value="M" checked>Manual<br>
															<Input type="Radio" name="RadSerNo" Value="A">Auto<br>
															<!--<Input type="Checkbox" name="chkSerNo" Value="W">Within Lot-->
														</td>
													</tr>
													<tr>
														<td class="ExcelDisplaycell" align="LEft">
															<Input type="Checkbox" name="chkLotNo">Enforce</tD>
														<td class="ExcelDisplaycell" align="LEft">
														<Input type="Checkbox" name="chkSerNo">Within Lot</td>
													</tr>
												</table>
											</td>
											<td width="50%" valign="Top">
												<table border="0"  cellspacing="1" class="ExcelTable" ID="tabNoOfLevel" width="100%">
													<!--<tr>
														<td class=ExcelDisplaycell align="Left" colspan="2"><b>Number Of Sublevels</b> &nbsp;&nbsp;
														<Input type="text" name="txtNoOfSubLevel" class="FormElem" size="4" onChange="AddSubLevel()" ></td>
													</tr>-->
													
												    <tr>
												        <td class="ExcelHeaderCell" align="center">Level No</td>
												        <td class="ExcelHeaderCell" align="center">Level Label</td>
												    </tr>
												</table>
											</td>
										</tr>
										
									</Table>
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                            <tr>
                                <td align="center" width="5" class="ClearPixel">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" style="width:100%">
								    <table cellpadding="0" cellspacing="0" width="100%">
								        <tr>
                                            <td>
								                <table cellpadding="0" cellspacing="0" width="100%">
								                    <tr>
								                        <td class="GroupTitleLeft" style="width:10">&nbsp;</td>
								                        <td class="GroupTitle" style="width:55">Template</td>
								                        <td class="GroupTitleRight">&nbsp;</td>
								                    </tr>
								                </table>
								            </td>
								        </tr>
								        <tr>
							            <td class="GroupTable">
							                <table cellpadding="0" cellspacing="0" width="100%" border="0">
							                    <tr>
								                    <td align="center" class="TopPack" colspan="3">
								                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="FieldCell">
                                                        <table width="100%">
                                                            <tr>
                                                                <td class="FieldCellSub">
                                                                    <input type="text" name="txtGrossPack" class="FormElem" value="Gross/Pack" size="12">
                                                                </td>
                                                                <td class="FieldCellSub">
                                                                    <input type="checkbox" name="chkTare"></td>
                                                                <td class="FieldCellSub">
                                                                    <input type="text" name="txtTarePack" class="FormElem" value="Tare/Pack" size="12">
                                                                </td>
                                                                <td class="FieldCellSub">
                                                                    <input type="button" name="btnPreview" value="Preview" class="AddButton" onclick="PopSample()">
                                                                </td>
                                                            </tr>
                                                             <tr>
                                                                <td class="FieldCellSub" colspan="4">
                                                                    <font color="red">If Tare/Pack is not selected, then Gross/Pack will be treated as Nett/Pack</font>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td colspan="4">
                                                                    <div id="divPreview" style="display:none">
                                                                        <table>
                                                                            <tr>
                                                                                <td id="tdGross" class="FieldCellSub"></td>
                                                                                <td class="FieldCellSub">
                                                                                    <input type="text" id="txtGross" class="Formelem" size="5">&nbsp;UOM
                                                                                </td>
                                                                                <td id="tdTare" class="FieldCellSub"></td>
                                                                                <td class="FieldCellSub" id="divTare" style="display:none">
                                                                                    <input type="text" id="txtTare" class="FormElem" size="5">&nbsp;UOM
                                                                                </td>
                                                                                <td class="FieldCellSub" id="divCone">
                                                                                    with
																	                <input type="text" size="3" name="txtNoofCone" class="FormElem">&nbsp;<span id="spanSellingForm">Cone</span>&nbsp;of 
																    	            <input type="text" name="txtWeight" size="3" class="FormElem" value="0">&nbsp;UOM&nbsp;each
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </div>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                                <tr>
								                    <td align="center" class="BottomPack" colspan="3">
								                    </td>
                                                </tr>
							                </table>    
							            </td>
								    </table>
								</td>
							</tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center">
													<input type="button" value="Save" name="B3" class="ActionButton" onclick="Save()" >
                                                    <input type="button" value="Close" name="B2" class="ActionButton" onclick="window.close()" >
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="BottomPack" colspan="3">
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
</html>

