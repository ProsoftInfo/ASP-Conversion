<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtSMEntry.asp
	'Module Name				:	Inventory (Stock Management Status Management)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 02, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Dec 21,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	stkMgmtSMInsert.asp
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock Management - Status Management</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Output/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="OutData1">
<Output/>
</script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/stockStatusManagementEntryModern.js"></SCRIPT>
</head>

<%
	dim iCtr,arrTemp,sTemp,arrValue,sOrgID,iClass,arrTempName,sTempName
	dim sOrgName,sClassName,rsTemp
	
	set rsTemp = server.CreateObject("ADODB.Recordset")
	
	'sOrgName = trim(Request.Form("hOrgName"))
	sClassName = trim(Request.Form("hClassName"))
	'sOrgID = trim(Request.Form("selUnit"))
	sOrgID = Session("organizationcode")
	iClass = trim(Request.Form("selClass"))
	sTemp  = trim(Request.Form("hSelectedValue"))
	sTempName = trim(Request.Form("hItemNames"))
	
	'Response.Write "<p>sTemp="&sTemp
	
	if sTempName  = "" then
		with rsTemp
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source= "Select ItemDescription from Inv_M_ItemMaster where ItemCode = " & mid(sTemp,1,len(sTemp)-1) 
			.Open
		end with
		
		if not rsTemp.EOF then
			sTempName = trim(rsTemp(0)) & "|"
		end if 
		rsTemp.Close 	
	end if 'if sTempName  = "" then
	
	if trim(sClassName) = "" then
		with rsTemp
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source= "Select GroupName from Inv_M_Classification where GroupCode = " & iClass
			.Open
		end with
		
		if not rsTemp.EOF then
			sClassName = trim(rsTemp(0))
		end if 
		rsTemp.Close
	end if 'if trim(sClassName) = "" then
	
	arrTempName = split(mid(sTempName,1,len(sTempName)-1),"|")
	arrTemp = split(mid(sTemp,1,len(sTemp)-1),"|")
	
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="FnInit('<%= mid(sTemp,1,len(sTemp)-1)%>')">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hClass" value="<%=iClass%>">
<input type=hidden name="hItemCode" value="<%= mid(sTemp,1,len(sTemp)-1)%>">
<input type=hidden name="hItem" value="">
<input type=hidden name="hRcptNumbering" value="">
<input type=hidden name="hLoc" value="">
<input type=hidden name="hBin" value="">

<input type="hidden" name="hCallFrom" value="<%=Request.Form("hCallFrom")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Status Management
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<td height="20" valign="bottom">
						<!--<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td width="100%" align="center">Header</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">Control</td>
										</tr>
									</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    <font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font>
								</td>
							</tr>
						</table>-->
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
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <!--<td class="FieldCell">Organization</td>
                                            <td class="FieldCellSub">
	                                            <span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                                            </td>
                                            <td class="FieldCellSub"></td>-->
                                            <!--<td class="FieldCell">Classification</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idClass"><%=sClassName%>&nbsp;</span>
                                            </td>-->
                                        </tr>
                                        <!--<tr>
                                            <td class="FieldCell">Item Name</td>
                                            <td class="FieldCellSub" colspan=4>
												<select size="1" name="selItem" class="FormElem" onChange="GetXML()">
													<option value="select">Select</option>
												<%
													for iCtr = 0 to UBound(arrTempName)
												%>
													<option value="<%=arrTemp(iCtr)%>"><%=arrTempName(iCtr)%></option>
												<%
													next
												%>
												</select>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">UoM</td>
                                            <td class="FieldCellSub">
	                                            <span class="DataOnly" id="idUoM"></span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Store</td>
                                            <td class="FieldCellSub">
												<select size="1" name="selStore" class="FormElem" onChange="DisplayDetails(this.value)">
													<option value="select">Select</option>
												</select>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCell"></td>
                                            <td class="FieldCellSub"></td>
                                        </tr>-->
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <div class="frmBody" id="frm2" style="width: 575; height:300;">
                                        <table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">Item Name</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">Store</td>
                                                <td class="ExcelHeaderCell" align="center" colspan="2">Existing Stock Information</td>
                                                <td class="ExcelHeaderCell" align="center" colspan="3">Status Management</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">UoM</td>
                                            </tr>
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center">Lot Number</td>
                                                <td class="ExcelHeaderCell" align="center">Quantity</td>
                                                <td class="ExcelHeaderCell" align="center">Rejected</td>
                                                <td class="ExcelHeaderCell" align="center">On Hold</td>
                                                <td class="ExcelHeaderCell" align="center">Reserved</td>
                                            </tr>
                                        </table>
                                    </div>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
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
                                                <input type="button" value="Save" name="B3" class="ActionButton" onClick="CheckSubmit()">
                                                <input type="reset" value="Reset" name="B1" class="ActionButton">
                                                <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="Cancel('stkMgmtEntry.asp')">
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
