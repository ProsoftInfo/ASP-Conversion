<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	StkMergeEntry.asp
	'Module Name				:	Inventory (Stock Management Status Management)
	'Author Name				:	Ragavendran R
	'Created On					:	May 25, 2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
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
<HTML><HEAD><TITLE>Stock Management - Stock Merge</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Output/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="OutData1">
<Output/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="ItemSelectData">
    <Root />
</script>
<script type="application/xml" data-itms-xml-island="1" id="LotData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempItemData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemTypeData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="IssueData"><ISSTYPE></ISSTYPE></script>
<script type="application/xml" data-itms-xml-island="1" id="IntReceipt"><ROOT></ROOT></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script language="javascript" src="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/stockMergeEntryModern.js"></SCRIPT>
</head>

<%
	dim iCtr,arrTemp,sTemp,arrValue,sOrgID,iClass,arrTempName,sTempName
	dim sOrgName,sClassName,rsTemp,iCreatedBy,dCreatedOn
	
	set rsTemp = server.CreateObject("ADODB.Recordset")
	'sOrgName = trim(Request.Form("hOrgName"))
	sClassName = trim(Request.Form("hClassName"))
	sOrgID = Session("organizationcode")
	iClass = trim(Request.Form("selClass"))
	sTemp  = trim(Request.Form("hSelectedValue"))
	sTempName = trim(Request.Form("hItemNames"))
	
	'Response.Write "<p>sTemp="&sTemp
	iCreatedBy = Session("userid")
	dCreatedOn = FormatDate(Date())
	
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
<input type="hidden" name="hItemType" value="">
<input type=hidden name="hUserID" value="<%=iCreatedBy%>">
<input type=hidden name="hCreatedOn" value="<%=dCreatedOn%>">
<input type=hidden name="hItemRow" value="">
<input type=hidden name="hFromItemType" value="">
<input type=hidden name="hReceiptNum" value="">
<input type=hidden name="hFromRcptNum" value="">

<input type="hidden" name="hCallFrom" value="<%=Request.Form("hCallFrom")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Stock Merge
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
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <div class="frmBody" id="frm2" style="width: 700; height:150;">
                                        <table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10" rowspan=2 >S.No.</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan=2>Item To Merge</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan=2>Store</td>
                                                <td class="ExcelHeaderCell" align="center" colspan=2>Quantity</td>
                                            </tr>
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" >Stock</td>
                                                <td class="ExcelHeaderCell" align="center" >Merge</td>
                                            </tr>
                                        </table>
                                    </div>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <div class="frmBody" id="Div1" style="width: 700; height:150;">
                                        <table border="0" cellspacing="1" id="tblToItemData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10" >S.No.</td>
                                                <td class="ExcelHeaderCell" align="center" >Item Merged With</td>
                                                <td class="ExcelHeaderCell" align="center" >Store</td>
                                                <td class="ExcelHeaderCell" align="center" >Stock</td>
                                                <td class="ExcelHeaderCell" align="center" >Merged Qty</td>
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
