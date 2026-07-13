<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	itmCodeCreate.asp
	'Module Name				:	Inventory (Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 20, 2003
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
<!--#include file="../../include/populate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Code Creation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
dim objfs

set objfs = Server.CreateObject("Scripting.FileSystemObject")
if objfs.FileExists(Server.MapPath("..\Temp\Master\ItemCodeCreate" & Session.SessionID & ".xml")) then
%>
<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="..\Temp\Master\ItemCodeCreate" & Session.SessionID & ".xml"%>"></script>
<%	else %>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><ROOT TYPE=""/></script>
<%	end if %>
<script type="application/xml" data-itms-xml-island="1" id="Data"><ROOT/></script>
<%
	Dim oDom,Root,HeaderNode,newElem,iLen,iFindLen
	Dim dcrs,dcrs1,sCode,sCodeName,iCount,sItmType,arrCode,iRecCount
	dim sItemCode,iStart,sItemDesc,iValue

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	sItmType = trim(Request.QueryString("sTemp"))
	sItemCode = trim(Request.QueryString("ItemCode"))
	sItemDesc = trim(Request.QueryString("ItemDesc"))

	if sItmType = "FAB" then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CODELENGTH FROM APP_M_CODETYPES WHERE ITEMTYPEID = " & Pack(sItmType) & " AND DISPLAYORDER = 1"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then iLen = trim(dcrs(0))
		dcrs.Close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CODELENGTH FROM APP_M_CODETYPES WHERE ITEMTYPEID = " & Pack(sItmType) & " AND DISPLAYORDER = 2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then iFindLen = trim(dcrs(0))
		dcrs.Close

		Set Root = oDOM.createElement("FABRIC")
		oDOM.appendChild Root

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CODE FROM APP_M_CODEMASTER WHERE CODETYPE=(SELECT CODETYPE FROM APP_M_CODETYPES WHERE ITEMTYPEID = " & Pack(sItmType) & " AND DISPLAYORDER = 1)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		do while not dcrs.eof

			iValue = 0
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT COMPANYITEMCODE FROM INV_M_ITEMMASTER WHERE ITEMCODE IN (SELECT DISTINCT ITEMCODE FROM INV_M_ITEMGROUP WHERE CLASSIFICATIONCODE IN (SELECT GROUPCODE FROM INV_M_CLASSIFICATION WHERE ITEMTYPEID = " & Pack(sItmType) & ")) AND LEFT(COMPANYITEMCODE," & iLen & ")=" & Pack(trim(dcrs(0)))
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			Do While Not dcrs1.EOF
				if cint(iValue) < cint(trim(mid(dcrs1(0),iLen+1,iFindLen))) then iValue = cint(trim(mid(dcrs1(0),iLen+1,iFindLen)))
				dcrs1.MoveNext
			loop
			dcrs1.Close

			iValue = iValue + 1
			Set newElem = oDOM.createElement("DETAILS")
			newElem.setAttribute "CODE",trim(dcrs(0))
			newElem.setAttribute "VALUE",GetValue(ivalue,iFindLen)
			Root.appendChild newElem

		dcrs.MoveNext
		Loop
		dcrs.close

		oDOM.Save server.MapPath("../temp/master/ItemCreateValue" & Session.SessionID & ".xml")
%>
<script type="application/xml" data-itms-xml-island="1" id="MaxData" data-src="<%="..\Temp\Master\ItemCreateValue" & Session.SessionID & ".xml"%>"></script>
<%	else %>
<script type="application/xml" data-itms-xml-island="1" id="MaxData"><ROOT/></script>
<%	end if %>
<script type="application/xml" data-itms-xml-island="1" id="ItemData"><ROOT/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmCodeCreate.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<%

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT CODETYPE,CODETYPENAME,CODELENGTH FROM APP_M_CODETYPES WHERE ITEMTYPEID = " & Pack(sItmType) & " ORDER BY DISPLAYORDER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	iRecCount = cint(dcrs.recordCount) - 1

	if not dcrs.EOF then
		arrCode = dcrs.getRows()
	end if
	dcrs.Close
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init('<%=sItmType%>')">
<form method="POST" name="formname">
<input type="hidden" name="hCount" value="<%=iRecCount%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr>
		<td align="center" class="TopPack">
		</td>
    </tr>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Code Creation
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center"></td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
											<td class="FieldCell" width="90">Product Name</td>
											<td class="FieldCell" colspan="3">
												<input TYPE="TEXT" NAME="txtProductname" VALUE="<%=sItemDesc%>" SIZE="65" maxlength="60" class="formelem">
											</td>
										</tr>
 									</table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan=3 class="MiddlePack">
							        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
						<%
							if sItmType = "FAB" then
						%>
							<tr>
								<td align="center"></td>
								<td valign="top" class="FieldCell" width="100%">
                                   <table border="0" cellpadding="0" cellspacing="1" class="TableOutlineOnly">
                                        <tr>
                                            <td class="FieldCellSub">Weave</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtWeave" VALUE="" SIZE="15" maxlength="50" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">No of ends / dent</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtDent" VALUE="" SIZE="15" maxlength="4" onkeypress="DoKeyPress('N',4,0)" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Width</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtWidth" VALUE="" SIZE="15" maxlength="9" onkeypress="DoKeyPress('Y',6,3)" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">No of ends - Total</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtEnds" VALUE="" SIZE="15" maxlength="5" onkeypress="DoKeyPress('N',5,0)" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Reed Count</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtReedCount" VALUE="" SIZE="15" maxlength="9" onkeypress="DoKeyPress('Y',6,3)" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">No of ends / inch(avg)</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtEndsInch" VALUE="" SIZE="15" maxlength="9" onkeypress="DoKeyPress('Y',6,3)" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Reed Space</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtReedSpace" VALUE="" SIZE="15" maxlength="50" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">No of Picks / inch(avg)</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtPicksInch" VALUE="" SIZE="15" maxlength="9" onkeypress="DoKeyPress('Y',6,3)" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Weight (lbs/yrds)</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtWeight" VALUE="" SIZE="15" maxlength="9" onkeypress="DoKeyPress('Y',6,3)" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">Average Warp Count</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtAvgWrap" VALUE="" SIZE="15" maxlength="10" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Variety</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtVariety" VALUE="" SIZE="15" maxlength="50" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub">Tape Length (%)</td>
                                            <td class="FieldCell">
												<input TYPE="TEXT" NAME="txtTapeLne" VALUE="" SIZE="4" maxlength="5" onkeypress="DoKeyPress('Y',2,2)" class="formelem">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Warp Yarn</td>
                                            <td class="FieldCell" colspan=4>
												<select size="5" name="selWarp" class="FormElem" multiple>
											<%	'Calling the Function which populates Yarn List
												populateYarn
											%>
												</select>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>

                                        <tr>
                                            <td class="FieldCellSub">Weft Yarn</td>
                                            <td class="FieldCell" colspan=4>
												<select size="5" name="selWeft" class="FormElem" multiple>
											<%	'Calling the Function which populates Yarn List
												populateYarn
											%>
												</select>
												<input type="button" value="Check" name="B7" class="AddButtonX" onClick="CheckData('<%=sItmType%>')">
                                            </td>
                                            <td class="FieldCellSub"></td>
                                        </tr>
									</table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan=3 class="MiddlePack">
							        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
						<%
							end if
						%>

							<tr>
								<td align="center"></td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
									<%
										iStart = 1
										for iCount = 0 to iRecCount
									%>
                                        <tr>
                                            <td class="FieldCell"><%=trim(arrCode(1,iCount))%></td>
                                            <td class="FieldCell">
                                            <%	if iCount+1 = 1 and sItmType = "FAB" then %>
												<select size="1" name="sel<%=iCount+1%>" class="FormElem" onChange="LoadMaxData()">
											<%	else %>
												<select size="1" name="sel<%=iCount+1%>" class="FormElem">
											<%	end if %>
													<option value="select">Select</option>
											<%	'Calling the Function which populates Codes List
												if sItemCode = "" then
													populateCode arrCode(0,iCount),0
												else
													populateCode arrCode(0,iCount),mid(sItemCode,iStart,arrcode(2,icount))
													iStart = iStart + cint(arrcode(2,icount))
												end if
												iCount = iCount + 1
											%>
												</select>
                                            </td>
											<%
												if 	iCount > iRecCount then exit for
											%>
                                            <!--td class="FieldCellSub"></td-->
                                            <td class="FieldCell"><%=trim(arrCode(1,iCount))%></td>
                                            <td class="FieldCell">
                                            <%	if iCount+1 = 2 and sItmType = "FAB" then %>
													<input type="radio" value="A" name="radSort" class="formelem"  onClick="document.formname.txtManual.value = '';LoadMaxData()">Automatic
													<input TYPE="TEXT" NAME="txtAutomatic" VALUE="" SIZE="5" maxlength="3" class="FormElemRead" READONLY>
													<input type="radio" value="M" name="radSort" class="FormElem" onClick="document.formname.txtAutomatic.value = ''">Manual
													<input TYPE="TEXT" NAME="txtManual" VALUE="" SIZE="5" maxlength="<%=arrcode(2,icount)%>" onkeypress="DoKeyPress('N',<%=arrcode(2,icount)%>,0)" class="formelem">
													<input TYPE="hidden" NAME="hSort" VALUE="<%=mid(sItemCode,iStart,arrcode(2,icount))%>" >
                                            <%
												iStart = iStart + cint(arrcode(2,icount))
												else	%>
													<select size="1" name="sel<%=iCount+1%>" class="FormElem">
														<option value="select">Select</option>
												<%	'Calling the Function which populates Codes List
													if sItemCode = "" then
														populateCode arrCode(0,iCount),0
													else
														populateCode arrCode(0,iCount),mid(sItemCode,iStart,arrcode(2,icount))
														iStart = iStart + cint(arrcode(2,icount))
													end if
												%>
													</select>
											<%	end if %>
                                            </td>
                                            <!--td class="FieldCellSub"></td-->
                                        </tr>
									<%
										next
									%>
                                    </table>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3"></td>
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
                                                    <input type="button" value="Save" name="B4" class="ActionButton" onClick="CheckSubmit('<%=sItmType%>')">
                                                    <input type="reset" value="Reset" name="B5" class="ActionButton" >
                                                    <input type="button" value="Close" name="B6" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="BottomPack" colspan="3">
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
</HTML>
<%
	' Function to populate Codes List
	Function populateCode(iID, iValue)
		' Declaration of variables
		Dim dcrs,iOptVal,sOptName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CODE,CODENAME FROM APP_M_CODEMASTER WHERE CODETYPE = '" & iID & "'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set iOptVal = dcrs(0)
		set sOptName = dcrs(1)

		Do While Not dcrs.EOF
			if trim(iOptVal) = trim(iValue) then
				Response.Write("<OPTION VALUE="""&trim(iOptVal)&""" SELECTED>"&trim(sOptName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(iOptVal)&""">"&trim(sOptName)&"</OPTION>" &vbcrlf)
			end if
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>

<%
	' Function to populate Yarn List
	Function populateYarn()
		' Declaration of variables
		Dim dcrs,iYarn,sYarnName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE,ITEMDESCRIPTION FROM INV_M_ITEMMASTER WHERE ITEMCODE IN (SELECT ITEMCODE FROM INV_M_ITEMORGMASTER WHERE CLASSIFICATIONCODE IN (SELECT GROUPCODE FROM INV_M_CLASSIFICATION WHERE ITEMTYPEID = 'YRN'))"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set iYarn = dcrs(0)
		set sYarnName = dcrs(1)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(iYarn)&""">"&trim(sYarnName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>

<%
	Function GetValue(iValue,iLen)
		dim str,i,k

		i = len(cstr(iValue))
		if i < iLen then
			for k = 1 to iLen - i
				str = str & "0"
			next
		end if
		str = str & iValue
		GetValue = str
	End Function
%>

