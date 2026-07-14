<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PackingForEntry.asp
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
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<%
Dim objRs,objRs1,rsTemp
Dim sOrgId,iSno,sLastLotNo,sQuery,sSubLevelName

set objRs  = server.CreateObject("adodb.recordset")
set objRs1  = server.CreateObject("adodb.recordset")
set rsTemp  = server.CreateObject("adodb.recordset")

sLastLotNo =""
sQuery = "select top 1 ReceiptNumber,MillLotNo from RCV_T_ActualRcptItemLot where MillLotNo is not null order by ReceiptNumber desc"
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sQuery
	.Open
End With
Set rsTemp.ActiveConnection = Nothing

if not rsTemp.EOF then
	sLastLotNo = rsTemp(1)
end if 
rsTemp.Close 


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="SubLevelQty"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="PACKFORMDATA">
<Root>
<%
Dim sCode,sName
'sSql = "Select PackingCode,Packingname from APP_M_PackingType where packingcode in (Select Packingcode from SAL_R_itemtypepack where ItemtypeID = '"&sItemType&"')"
sQuery = "Select PackingCode,Packingname from APP_M_PackingType "
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sQuery
	.Open
End With
Set rsTemp.ActiveConnection = Nothing

Set sCode = rsTemp(0)
Set sName = rsTemp(1)

If not rsTemp.EOF then
	Do while Not rsTemp.eof
		Response.Write("<PACK VALUE="""&trim(sCode)&""" NAME="""&trim(sName)&""" />" &vbcrlf)
		rsTemp.MoveNext
	Loop
End if
rsTemp.close
%>
</Root>
</script>
<script type="application/xml" data-itms-xml-island="1" id="LotData"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/packingForEntry.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
	<Input type="hidden" name="hLastUsedLotNo" value="<%=sLastLotNo%>">
	<Input type="hidden" name="hCnt" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot Entry
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
									<td class="FieldCellsub">Item Description </td>
									<td class="FieldCellSub">
										<span Id="dataonly">xxxx</span>
									</td>
									<td class="FieldCellsub">Date Of Arrival / Packing</td>
									<td class="FieldCellSub">
										<%Response.Write InsertDatePicker("ctlDate")%>
									</td>
                                </tr>
                                
                                <tr>
									<td class="FieldCellsub">Packing Type</td>
									<td class="FieldCellSub" colspan="3">
										<select class="FormElem" Name="selPackType" onchange=CheckLotSerial(this)>
											<option value="S">Select</option>
											<%
											sQuery = "SELECT PACKINGCODE,PACKINGSHORTNAME,PACKINGNAME,isNULL(NUMBERINGTYPE,''),isNull(ManualLotNumbering,''),isNull(ManualSerialNumbering,'') FROM APP_M_PACKINGTYPE ORDER BY PACKINGCODE"
											With objRs 
												.CursorLocation = 3
												.CursorType = 3
												.ActiveConnection = con
												.Source = sQuery
												.Open
											End With
											Set objRs.ActiveConnection = Nothing
											
											Do while Not objrs.EOF
												sSubLevelName = ""
												sQuery = "Select isNull(SubLevelName,'') From APP_M_PackingTypeSubLevel Where PackingCode="& objRs(0) &" "
												rsTemp.Open sQuery,con
												If Not rsTemp.EOF Then
													sSubLevelName = rsTemp(0)
												End IF
												rsTemp.Close 
												%>
												<option value="<%=objrs(0)%>:<%=objrs(3)%>:<%=sSubLevelName%>:<%=objrs(4)%>:<%=objrs(5)%>"><%=objrs(2)%></option>
												<%
												objrs.MoveNext 
											Loop
											objrs.Close 
											%>
										</select>
										&nbsp;No Of Packs&nbsp;
										<Input type="text" Name="txtNoOfPack" size="6" class="FormElem" value="">
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub">Gross/Pack</td>
									<td class="FieldCellSub" colspan="4">
										<Input type="text" Name="txtuniSerQty" size="10" class="FormElem" value="">
										<span id="Dataonly">KGS</span>
										&nbsp;Tare/Pack
										<Input type="text" Name="txtTareQty" size="10" class="FormElem" value="">
										<span id="Dataonly">KGS</span>
										&nbsp;with
										<input type="text" Name="txtSubLevelQty" size="5" class="FormElem" value="">
										<span ID="Data1">Cones</span>&nbsp;of
										<input type="text" Name="txtSubLevelwt" size="5" class="FormElem" value="">
										<span id="Dataonly">KGS</span>&nbsp;each
									</td>
                                </tr>
                                <tr>
									<td class="FieldCellsub"><span ID="Data3">Lot</span>&nbsp;No</td>
									
									
									<td class="FieldCellSub" ID="DivMain" style="display:block">
									    <div id="DivAuto" style="display:block">
										    <select class="FormElem" Name="selLotNo"  Disabled>
											    <option value="S"> Select </option>
												<option value="<%=sLastLotNo%>"> <%=sLastLotNo%></option>
											    <option value="N">New&nbsp;</option>
										    </select>
										</div>
										<div id="DivManual" style="display:none">
										<input type="text" Name="txtLotNo" size="5" class="FormElem" value="">
										</div>
									</td>
									
									<td class="FieldCellSub" ID="DivMain1" style="display:block">
										<div id="DivLastUsed" style="display:block">
										[Last Used&nbsp;<font color="red"><b><span Id="LastUsedData"></span></b></font>]&nbsp;
										<!--<input type="text" Name="LotNo" size="5" class="FormElem" value="">-->
										</div>
									</Td>
									
									<!--<td class="FieldCellSub" ID="SerialNo" style="visibility:hidden">
										<Input type="checkbox" Name="chkAutoSerial" class="FormElem" value="">
										Auto Serial&nbsp;[Last Used&nbsp;<font color="red"><b>XXX</b></font>]&nbsp;
									</td>-->
                                </tr>
                                <tr>
									<td class="FieldCellSub">Applicable Mix</td>
									<td class="FieldCellSub" colspan="2">
										<select class="FormElem" Name="selAppMix">
											<option value="S">Mix Name / code</option>
										</select>
									</td>
									<td class="FieldcellSub">
										<Input type="Button" name="btnAdd" class="ActionButton" value="Add" onclick="AddData()">
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
								<td Width="100%">
									<table border="0" cellspacing="1" class="ExcelTable" width="100%" id="tblLotDetail">
											<tr>
												<td class="ExcelSerial" align=center>&nbsp;</td>
												<td class="ExcelSerial" align=center>&nbsp;</td>
												<td class="ExcelHeaderCell" align="left" colspan="4">Lot No - Qty [Rate / KGS]</td>
												<td class="ExcelHeaderCell" align="center" colspan="2">Party</td>
											</tr>
											<tr>
												<td class="ExcelHeaderCell" align="center">Serial</td>
												<td class="ExcelHeaderCell" align="center" >
													<a href="#"><img style="cursor:hand" border="0" width="11" height="11" src="../../assets/images/iTMS Icons/DeleteIcon.gif" alt="Delete" onclick="DeleteData()"></a>
												</td>
												<td class="ExcelHeaderCell" align="center">Packing Type</td>
												<td class="ExcelHeaderCell" align="center">Gross Qty</td>
												<td class="ExcelHeaderCell" align="center">Nett Qty</td>
												<td class="ExcelHeaderCell" align="center">No Of <span Id="Data2"><b>Cones</b></span></td>
												<td class="ExcelHeaderCell" align="center">Serial</td>
												<td class="ExcelHeaderCell" align="center">Qty</td>
											</tr>
											<!--<tr>
												<td class="ExcelSerial" align=center>&nbsp;</td>
												<td class="ExcelSerial" align=center>&nbsp;</td>
												<td class=ExcelDisplaycell align="Left" colspan="4"><b>1 - 312.5 KGS [Rs.125.00]</b></td>
												<td class=ExcelDisplaycell align="Left" colspan="2"><b>1 - 312.5 KGS</b></td>
											</tr>
											<tr>
												<td class=ExcelSerial align=center>1</td>
												<td class=ExcelDisplaycell align=center>
													<Input type="Checkbox" name="chkbox" value="">
												</td>
												<td class="ExcelInputcell" align="center">
													<select class="FormElem" Name="selPackType">
														<option value="S">Yarn Bag</option>
													</select>
												</td>
												<td class="ExcelInputcell" align="Left">
													<Input type="text" name="txtSubLevelQty" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
													<a href="#"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Sub Level Qty" onclick="ShowSubLevelQty()"></a>
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtQty" value="" class="FormElem" size="10">
												</td>
											</tr>
											<tr>
												<td class="ExcelSerial" align="center">2</td>
												<td class="ExcelDisplaycell" align="center">
													<Input type="Checkbox" name="chkbox" value="">
												</td>
												<td class="ExcelInputcell" align="center">
													<select class="FormElem" Name="selPackType">
														<option value="S">Yarn Bag</option>
													</select>
												</td>
												<td class="ExcelInputcell" align="Left">
													<Input type="text" name="txtSubLevelQty" value="" class="FormElem" size=10>
												</td><td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
													<a href="#"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Sub Level Qty" onclick="ShowSubLevelQty()"></a>
												</td>
												<td class="ExcelInputcell" align="center">
													<Input type="text" name="txtSerial" value="" class="FormElem" size="10">
												</td>
												<td class="ExcelInputcell" align=center>
													<Input type="text" name="txtQty" value="" class="FormElem" size="10">
												</td>
											</tr>
											<tr>
												<td class=ExcelDisplayCell align=right colspan="3"><b>Total</b></td>
												<td class=ExcelDisplayCell align=right><b>xxxx</b></td>
												<td class=ExcelDisplayCell align=right><b>xxxx</b></td>
												<td class=ExcelDisplayCell align=right><b>xxxx</b></td>
												<td class=ExcelDisplayCell align=right colspan="2"></td>
											</tr>-->
									</Table>
								</td>
                            </tr>

                <!--<div class="frmBody" id="frm2" style="width: 50%; height:150;">
             	</div>-->

								</td>
								<td align="center" class="ClearPixel" width="5">
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

