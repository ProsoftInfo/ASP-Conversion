<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	DirectIssuePackingEntry.asp
	'Module Name				:	Inventory (Issue Additional Details)
	'Author Name				:	TAJUDEEN S
	'Created On					:	December 03, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
	' Declaration of variables
	Dim dcrs,iCtr,arrPONo
	'Declaration of Objects
	iCtr = 0
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	dim iItem,iClass,sItemName,sUsage,sClassName
	dim arrTemp,sOrgID,sOrgName,iQty,arrUoM,sUoMCode,sUoM,sPONO

	arrTemp = split(trim(Request.QueryString("sTemp")),"|")
	iClass = arrTemp(0)
	iItem = arrTemp(1)
	sOrgID = arrTemp(2)
	sUsage = arrTemp(3)
	iQty = arrTemp(4)
	sPONO = arrTemp(5)

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT GROUPNAME,ITEMDESCRIPTION,ORGUNITSHORTDESCRIPTION FROM VWITEM WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		sClassName = trim(dcrs(0))
		sItemName =  trim(dcrs(1))
		sOrgName = trim(dcrs(2))
	end if
	dcrs.Close

	sItemName = ItemDisplay(iItem,iClass)

	arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
	sUoMCode = arrUoM(0)
	sUoM = arrUoM(1)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Additional Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/DirectIssuePackingDetails.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=sOrgID%>','<%=iClass%>','<%=iItem%>','<%=sUsage%>')">
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Additional Details
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
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item Name</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Unit Name</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sOrgName%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idQty"><%=iQty%>&nbsp;</span>&nbsp;
												<span class="DataOnly"><%=sUoM%>&nbsp;</span>
											</td>
										</tr>
									</table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td width="100%">
									<div class="frmBody" id="frm2" style="width: 100%; height:230;">
										<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Production Order No</td>
												<td class="ExcelHeaderCell" align="center">Quantity</td>
											</tr>
										<%

											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT DISTINCT PRODUCTIONORDERNO FROM PRD_T_PRODUCTDETAILS WHERE PRODUCTIONORDERNO IN ( " & sPONo & ") AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass
												.ActiveConnection = con
												.Open
											end with
											set dcrs.ActiveConnection = nothing

											do while not dcrs.EOF
												iCtr = iCtr + 1

										%>
											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell" align="left"><%=dcrs(0)%></td>
												<td class="ExcelInputCell" align="right" width="80">
													<input type=text name="txtQty<%=iCtr%>" size="13" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" class="FormElem" style="text-align=right">
													<input type=hidden name="hPONO<%=iCtr%>" value="<%=dcrs(0)%>">
												</td>
											</tr>
										<%
												dcrs.moveNext
											loop
											dcrs.Close

										%>

										</table>
									</div>
								</td>
								<td align="center"></td>
								<input type=hidden name=hiCtr value="<%=iCtr%>">
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="window.close()">
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
	' Function to populate Store
	Function DisplayUoM(sOrgID,iClass,iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>
