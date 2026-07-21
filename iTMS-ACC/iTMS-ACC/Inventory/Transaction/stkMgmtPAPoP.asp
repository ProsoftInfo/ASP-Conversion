<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	stkMgmtPAPoP.asp
	'Module Name				:	Inventory (Stock Management Physical Adjustment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 30, 2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	stkMgmtPAEntry.asp
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
<!-- #include file="../../include/CommonFunctions.asp" -->
<%
' Declaration of variables
Dim dcrs2,iCtr
iCtr = 0
'Declaration of Objects
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

Dim iItem,iClass,sItemName,iLot,sBin,sLoc
Dim arrTemp,sOrgID,iQty,iInvRecNo,arr,sCheck

arrTemp = split(trim(Request.QueryString("sTemp")),"AAAA")

sOrgID = arrTemp(1)
iItem = arrTemp(2)
iClass = arrTemp(3)
iLot = arrTemp(4)
iQty = arrTemp(5)
sLoc = arrTemp(6)
sBin = arrTemp(7)
iInvRecNo = arrTemp(8)
sCheck = arrTemp(9)

'Response.write "Request.QueryString(sValue)="& Request.QueryString("sValue")
arr = split(trim(Request.QueryString("sValue")),"`")
sItemName = arr(0)
iQty = arr(1)

sItemName = GetItemName(iItem,iClass)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : Stock Management - Physical Adjustment</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Dim objTemp,Root,newElem
Dim iClass,iItem,iQty,ii,sLot,sBin,sStore,sOrgID,iInvRec

Function fnInit(obj)
	Set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement

	'if cdbl(document.formname.hiCtr.value) = 0 then exit function

	arrTemp = split(obj,"AAAA")
	sOrgID = arrTemp(1)
	iItem = arrTemp(2)
	iClass = arrTemp(3)
	sLot = arrTemp(4)

	if sLot = "0" then sLot = "NULL"

	sStore = arrTemp(6)
	sBin = arrTemp(7)
	iInvRec = arrTemp(8)

	For Each Node In Root.childNodes
		For Each HeaderNode in Node.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sStore and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				For Each HNode In HeaderNode.childNodes
					if HNode.Attributes.Item(0).nodeValue = sLot and HNode.Attributes.Item(1).nodeValue = iInvRec then
						if HNode.HaschildNodes() then
							For Each INode In HNode.childNodes
								if StrComp(Trim(INode.NodeName),"SERIALDETAILS") = 0 then
									ii = INode.Attributes.Item(3).nodeValue
									set Q = eval("document.formname.txtQty"&ii)
									Q.value = INode.Attributes.Item(2).nodeValue
								end if
							next
						end if
						exit function
					end if
				next
			end if
		end if
		Next
	Next
end Function

Function CheckSubmit()
	dim ictr,objQ,iQtyTot,objSTQ,objSerial

	ictr = document.formname.hiCtr.value
	if ictr = "" then exit function

	for i=1 to ictr
		set objQ = eval("document.formname.txtQty"&i)
		set objSTQ = eval("document.formname.txtStQty"&i)

		if trim(objQ.value) = "" then
			msgbox "Enter Quantity",0,"Quantity"
			objQ.select()
			exit function
		else
			iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
		end if

	next

	if (cdbl(iQtyTot) <> cdbl(idQty.innerText)) then
		alert("Total Adjust Quantity should be equal to Quantity (" & idQty.innerText & ")")
		exit function
	end if

	For Each Node In Root.childNodes
		For Each HeaderNode In Node.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sStore and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				For Each HNode In HeaderNode.childNodes
					if HNode.Attributes.Item(0).nodeValue = sLot and HNode.Attributes.Item(1).nodeValue = iInvRec then
						if HNode.HaschildNodes() then
							For Each INode In HNode.childNodes
								if StrComp(Trim(INode.NodeName),"SERIALDETAILS") = 0 then
									set a = HNode.removeChild(INode)
								end if
							next
							exit for
						end if
					end if
				next
			end if
		end if
		Next
	Next

	For Each Node In Root.childNodes
		For Each HeaderNode In Node.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sStore and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				For Each HNode In HeaderNode.childNodes
					if HNode.Attributes.Item(0).nodeValue = sLot and HNode.Attributes.Item(1).nodeValue = iInvRec then
						for i=1 to ictr
							objQty = trim(eval("document.formname.txtQty"&i&".value"))
							if objQty <> "" then
								if cdbl(objQty) <> 0 then
									set objSerial = eval("document.formname.txtSerial"&i)
									Set newElem = objTemp.createElement("SERIALDETAILS")
									newElem.setAttribute "SERIALNO", trim(objSerial.value)
									newElem.setAttribute "STKQTY", trim(eval("document.formname.txtStQty"&i&".value"))
									newElem.setAttribute "QTY", trim(eval("document.formname.txtQty"&i&".value"))
									newElem.setAttribute "ENTRYNO", i
									HNode.appendChild newElem
								end if
							end if
						next
					end if
				next
			end if
		end if
		Next
	Next

	For Each Node In Root.childNodes
		For Each HeaderNode In Node.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sStore and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				For Each HNode In HeaderNode.childNodes
					if HNode.Attributes.Item(0).nodeValue = sLot and HNode.Attributes.Item(1).nodeValue = iInvRec then
						HNode.setAttribute "QTYISS", cdbl(iQtyTot)
						window.close
						exit function
					end if
				next
			end if
		end if
		Next
	Next

end Function

Function window_onunload()
	set window.returnValue = objTemp.documentElement
	window.close()
end Function


</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>')">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot / Serial Details
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item &nbsp;</td>
                                            <td class="FieldCellSub" colspan="4">
												<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idQty"><%=iQty%>&nbsp;</span>
                                            </td>
                                            <td class="FieldCellSub">Lot No. <span class="DataOnly"><%=replace(iLot,"NULL","N/A")%>&nbsp;</span></td>
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
								<td>
									<div class="frmBody" id="frm2" style="width: 352; height:240;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Serial / Packing Number</td>
												<td class="ExcelHeaderCell" align="center">Stock</td>
												<td class="ExcelHeaderCell" align="center">Quantity Adjust</td>
											</tr>
										<%
											with dcrs2
												.CursorLocation = 3
												.CursorType = 3
												if iLot = "NULL"  or iLot="" then
													'.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0),PACKINGNUMBER FROM INV_T_RECEIPTLOTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 ORDER BY 1"
													.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0),PACKINGNUMBER FROM INV_T_LocationLot WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 ORDER BY 1"
												else
													'.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0),PACKINGNUMBER FROM INV_T_RECEIPTLOTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 ORDER BY 1"
													.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0),PACKINGNUMBER FROM INV_T_LocationLot WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 ORDER BY 1"
												end if
												.ActiveConnection = con
												.Open
											end with
											
											set dcrs2.ActiveConnection = nothing
											
											if not dcrs2.EOF then
												Do While Not dcrs2.EOF
													iCtr = iCtr + 1
													if not (cdbl(trim(dcrs2(1))) - cdbl(trim(dcrs2(2)))) = 0 then
										%>

											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell">
													<input type="text" name="txtSerialD<%=iCtr%>" size="28" maxlength=10 value="<%=trim(dcrs2(3))%>" class="FormElemRead" READONLY>
													<input type="hidden" name="txtSerial<%=iCtr%>" value="<%=trim(dcrs2(0))%>">
												</td>
												<td class="ExcelDisplayCell" width="10">
													<input type="text" name="txtStQty<%=iCtr%>" size="11" maxlength=10 value="<%=cdbl(trim(dcrs2(1))) - cdbl(trim(dcrs2(2)))%>" class="FormElemRead" READONLY style="text-align=right">
												</td>
												<td class="ExcelInputCell" width="10">
													<input type="text" name="txtQty<%=iCtr%>" size="11" onkeypress="DoKeyPressHypen('<%=sCheck%>',7,3)" value="0" class="FormElem" style="text-align=right">
												</td>
											</tr>
										<%
													end if
												dcrs2.MoveNext
												Loop
											end if
											dcrs2.Close
										%>
										</table>
										
										<input type=hidden name="hiCtr" value="<%=iCtr%>">
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="reset" value="Reset" name="B2" class="ActionButton">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack"></td>
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