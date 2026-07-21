<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsIssueSTPoP.asp
	'Module Name				:	Inventory (MRS Approval)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 17, 2003
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
dim oDom,Root,PageNode,HeaderNode,PGNode,objfs

dim iItem,iClass
Dim dcrs,dcrs1,dcrs2,sLocOrgID,sLocOrgName
dim iCtr,iQty,arrTemp,iMRSNo,sOrgID,sOrgName,sItemName,sUnitName,iQtyTransfer
dim arrUoM,sUoMDesc,sUoMCode,sTemp
dim iEntNo,sOptName

iCtr = 0
'Declaration of Objects
set dcrs = server.CreateObject("Adodb.recordset")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

set oDom = server.CreateObject("Microsoft.xmlDom")
Set objfs = CreateObject("Scripting.FileSystemObject")
'Response.Write "Request="&trim(Request.QueryString("sTemp"))
arrTemp = split(trim(Request.QueryString("sTemp")),":")
iClass	= arrTemp(1)
iItem	= arrTemp(0)
iMRSNo = arrTemp(2)
'sUnitName = arrTemp(3)
iEntNo = arrTemp(4)
sOptName = arrTemp(5)
sOrgID =  arrTemp(6)
Response.Write "iEntNo="&iEntNo
'sItemName = ItemDisplay(iItem,iClass)
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ITEMDESCRIPTION FROM VWITEM WHERE ITEMCODE = " & iItem & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	sItemName = trim(dcrs(0))
end if
dcrs.close
sItemName = sItemName & sOptName



Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
if objfs.FileExists(Server.MapPath("../temp/transaction/MRS"&iMRSNo&".xml")) then
	oDOM.Load server.MapPath("../temp/transaction/MRS"&iMRSNo&".xml")
	Set Root = oDOM.documentElement

	if Root.HaschildNodes() then

		For Each HeaderNode In Root.childNodes

			if StrComp(HeaderNode.nodeName,"MRSHeader") = 0 then
				'iMRSNo = HeaderNode.Attributes.Item(0).nodeValue & "&nbsp;"
				sOrgID = HeaderNode.Attributes.Item(2).nodeValue
				sOrgName = HeaderNode.Attributes.Item(3).nodeValue & "&nbsp;"
			end if
			if StrComp(HeaderNode.nodeName,"ITEM") = 0 then

				if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = iItem and HeaderNode.Attributes.Item(2).nodeValue = iClass then
					'sItemName = HeaderNode.Attributes.Item(2).nodeValue & "&nbsp;"
					sUoMDesc = HeaderNode.Attributes.Item(5).nodeValue
					iQty = HeaderNode.Attributes.Item(6).nodeValue
				end if
			end if
		next
	end if
else
IF iMRSNo <> "" then
	with dcrs2
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT MRSFORUNIT,ORGUNITSHORTDESCRIPTION FROM VWMRSLIST WHERE MRSNUMBER = " & iMRSNo & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
		.ActiveConnection = con
		.Open
	end with
	set dcrs2.ActiveConnection = nothing

	if not dcrs2.EOF then
		sOrgID = trim(dcrs2(0))
		sOrgName = trim(dcrs2(1))
	end if
	dcrs2.Close
	with dcrs2
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT ISNULL(QUANTITYAPPROVED,0) FROM VWMRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ISNULL(ICOUNTER,0) = " & iEntNo & " "
		.ActiveConnection = con
		.Open
	end with
	set dcrs2.ActiveConnection = nothing
'Response.Write dcrs2.Source
	if not dcrs2.EOF then
		iQty = trim(dcrs2(0))
	end if
	dcrs2.Close
End IF

end if

with dcrs2
	.CursorLocation = 3
	.CursorType = 3
	IF iMRSNo <> ""  then
		.Source = "SELECT ISNULL(QUANTITYREQUESTED,0),ISNULL(QUANTITYAPPROVED,0),ISNULL(QUANTITYISSUED,0),ISNULL(QUANTITYTOPURCHASE,0),ISNULL(QUANTITYFORTRANSFER,0) FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND ISNULL(ICOUNTER,0) = " & iEntNo & ""
	Else
		.Source = "SELECT ISNULL(QUANTITYREQUESTED,0),ISNULL(QUANTITYAPPROVED,0),ISNULL(QUANTITYISSUED,0),ISNULL(QUANTITYTOPURCHASE,0),ISNULL(QUANTITYFORTRANSFER,0) FROM INV_T_MRSITEMDETAILS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND ISNULL(ICOUNTER,0) = " & iEntNo & ""
	End IF
	.ActiveConnection = con
	.Open
end with
set dcrs2.ActiveConnection = nothing
'Response.Write dcrs2.Source

if not dcrs2.EOF then
	'Response.Write dcrs2(1) &" - "& dcrs2(2)  &" + "&  dcrs2(3)  &" + "&  dcrs2(4)
	iQty = cdbl(trim(dcrs2(1))) - (cdbl(trim(dcrs2(2))) + (cdbl(trim(dcrs2(3)))+cdbl(trim(dcrs2(4)))))
end if
dcrs2.Close
'sTemp  = DisplayUoM(sOrgID,iClass,iItem)
'if sTemp <> "" then
'	arrUoM = split(sTemp,":")
'
'	sUoMCode = arrUoM(0)
'	sUoMDesc = arrUoM(1)
'else
'	sUoMCode = ""
'	sUoMDesc = ""

'end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Stock Details for Transfer</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsIssueST.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0"  onLoad="fnInit('<%=iItem%>','<%=iClass%>','<%=cdbl(iQty)%>','<%=iEntNo%>')">

<form method="POST" name="formname" action="">
<Input type="hidden" name="hMrsNo" value="<%=iMRSNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20">
          <p align="center">Stock Transfer Details
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
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
                                    <div align="left">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td class=FieldCell> Item Description</td>
												<td class='FieldCellSub'>
                                                  <span class="DataOnly"><%=sItemName%></span>
                                                </td>
											</tr>
											<% IF iMRSNo <> "" then %>
											<tr>
												<td class=FieldCell> Quantity Pending</td>
												<td class='FieldCellSub'>
													<span class="DataOnly"><%=iQty%>&nbsp;</span>
													<span class="DataOnly"><%=sUoMDesc%>&nbsp;</span>
                                                </td>
											</tr>
											<% End IF %>
										</table>
                                    </div>
								</td>
								<td align="center"></td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
                            <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
									<div class="frmBody" id="frm2" style="width: 385; height:150;">
                                        <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
												<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
												<td class="ExcelHeaderCell" align="center" rowspan="2">Unit Name</td>
												<td class="ExcelHeaderCell" align="center" colspan="2"> Quantity</td>
                                            </tr>
                                            <tr>
												<td class="ExcelHeaderCell" align="center" width="80">In Stock</td>
												<td class="ExcelHeaderCell" align="center">To Transfer</td>
                                            </tr>
									<%'if 1 = 2 then
									'	with dcrs2
									'		.CursorLocation = 3
									'		.CursorType = 3
									'		.Source = "SELECT DISTINCT OTHERORGANISATIONCODE FROM INV_M_ITEMORGTRANSFERUNITS WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
									'		.ActiveConnection = con
									'		.Open
									'	end with
									'	set dcrs2.ActiveConnection = nothing

									'	if not dcrs2.EOF then
									'		Do While Not dcrs2.EOF
									'			sLocOrgID = trim(dcrs2(0))


												with dcrs
													.CursorLocation = 3
													.CursorType = 3
													'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sLocOrgID) & " AND STR(LOCATIONNUMBER)+STR(ISNULL(BINNUMBER,0)) IN (SELECT DISTINCT STR(IM.LOCATIONNUMBER)+STR(ISNULL(BINNUMBER,0)) FROM INV_M_ORGSTORAGE IC,INV_M_ITEMORGSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sLocOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' AND ALLOWTRANSFERS = 1) GROUP BY ORGANISATIONCODE"
													'old qry
													'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM VWSTOCKTRANSFERSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sLocOrgID) & " GROUP BY ORGANISATIONCODE"
													'old qry
													'.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM VWSTOCKTRANSFERSTOCK WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " GROUP BY ORGANISATIONCODE"
													.Source = "SELECT ISNULL(SUM(YEARCLOSINGSTOCK),0),ORGANISATIONCODE FROM Inv_T_ItemYearlyStock WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " GROUP BY ORGANISATIONCODE"
													.ActiveConnection = con
													.Open
												end with
												'set dcrs.ActiveConnection = nothing
												'  Response.Write "<BR><BR>" &dcrs.Source
												'Response.Write dcrs.EOF
											if not dcrs.EOF then
												Do While Not dcrs.EOF
													iCtr = iCtr + 1
													'Response.Write "<P>dcrs(0)="&dcrs(0)
													with dcrs1
														.CursorLocation = 3
														.CursorType = 3
														.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID = " & Pack(trim(dcrs(1))) & ""
														.ActiveConnection = con
														.Open
													end with
													set dcrs1.ActiveConnection = nothing
													if not dcrs1.EOF then
														sLocOrgName = trim(dcrs1(0))
													end if
													dcrs1.Close

													'with dcrs1
													'	.CursorLocation = 3
													'	.CursorType = 3
													'	.Source = "SELECT QUANTITYREQUESTEDTR FROM INV_T_MRSSTOCKTRANSFER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND MRSNUMBER = " & iMRSNo & " AND TRANSFERFROMUNIT = " & Pack(trim(dcrs(1))) & ""
													'	.ActiveConnection = con
													'	.Open
													'end with

													'set dcrs1.ActiveConnection = nothing
													'if not dcrs1.EOF then
													'	iQtyTransfer = trim(dcrs1(0))
													'end if
													'dcrs1.Close
									%>
                                            <tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell"><%=sLocOrgName%></td>
												<td class="ExcelDisplayCell" align="center"><p align="right"><%=trim(dcrs(0))%></td>
												<td class="ExcelInputCell" align="center" width="10">
													<input type="hidden" name="hST<%=iCtr%>" value="<%=trim(dcrs(0))%>">
													<input type="text" name="txtST<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" value="0" class="Formelem">
													<input type="hidden" name="hOrgID<%=iCtr%>" value="<%=trim(dcrs(1))%>">
												</td>
                                            </tr>
									<%
												dcrs.MoveNext
												Loop
												dcrs.Close
											'dcrs2.MoveNext
											'Loop
										else
									%>
                                            <tr>
												<td colspan=4 class="ExcelDisplayCell" align="center"><B>No Units Available</B></td>
                                            </tr>

									<%
										'end if
										'dcrs2.Close
										end if 'if not dcrs.EOF then
									%>
                                        </table>
                           			</div>
								</td>
								<td align="center"></td>
								<input type=hidden name="hiCtr" value="<%=iCtr%>">
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="OK" name="B3" class="ActionButton" onClick="CheckSubmit()">
													<input type="reset" value="Reset" name="B4" class="ActionButton">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td colspan="3" class="BottomPack">
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
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
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
