<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	stkMgmtSMPoP.asp
	'Module Name				:	Inventory (Stock Management Status Management)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 03, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	stkMgmtSMEntry.asp
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
<%
' Declaration of variables
Dim dcrs2,dcrs3,iCtr
iCtr = 0
'Declaration of Objects
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
Set dcrs3 = Server.CreateObject("ADODB.RecordSet")

dim iItem,iClass,sItemName,iLot,iStockNo,sRemarks,sBin,sLoc
dim arrTemp,sOrgID,iQty,iInvRecNo,arr,iRejQty,iOhoQty,iResQty,sCheck

arrTemp = split(trim(Request.QueryString("sTemp")),"AAAA")
sOrgID = arrTemp(1)
iItem = arrTemp(2)
iClass = arrTemp(3)
iLot = arrTemp(4)
iQty = arrTemp(5)
sLoc = arrTemp(6)
sBin = arrTemp(7)
iInvRecNo = arrTemp(8)
iStockNo = arrTemp(9)
sCheck = arrTemp(10)

arr = split(trim(Request.QueryString("sValue")),"`")
sItemName = arr(0)
iQty = arr(1)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : Stock Management - Status Management</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data">
<Data>
<%
	with dcrs2
		.CursorLocation = 3
		.CursorType = 3
		if iLot = "NULL" then
			.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0) FROM INV_T_RECEIPTLOTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 ORDER BY 1"
		else
			.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0) FROM INV_T_RECEIPTLOTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 ORDER BY 1"
		end if
		.ActiveConnection = con
		.Open
	end with
	set dcrs2.ActiveConnection = nothing

	if not dcrs2.EOF then
		Do While Not dcrs2.EOF
			if not (cdbl(trim(dcrs2(1))) - cdbl(trim(dcrs2(2)))) = 0 then
				with dcrs3
					.CursorLocation = 3
					.CursorType = 3
					if iLot = "NULL" then
						.Source = "SELECT ISNULL(CLEAN,0),ISNULL(RESERVED,0),ISNULL(ONHOLD,0),ISNULL(REJECTED,0) FROM INV_M_STOCKSTATUSDETAILS WHERE STOCKNO = " & iStockNo & " AND LOTNO IS NULL AND SERIALNO = " & trim(dcrs2(0)) & ""
					else
						.Source = "SELECT ISNULL(CLEAN,0),ISNULL(RESERVED,0),ISNULL(ONHOLD,0),ISNULL(REJECTED,0) FROM INV_M_STOCKSTATUSDETAILS WHERE STOCKNO = " & iStockNo & " AND LOTNO = " & Pack(iLot) & " AND SERIALNO = " & trim(dcrs2(0)) & ""
					end if
					.ActiveConnection = con
					.Open
				end with
				set dcrs3.ActiveConnection = nothing

				if not dcrs3.EOF then
					iResQty = trim(dcrs3(1))
					iOhoQty = trim(dcrs3(2))
					iRejQty = trim(dcrs3(3))
				else
					iResQty = "0"
					iOhoQty = "0"
					iRejQty = "0"
				end if
				Response.Write "<SERIAL SNO="""&trim(dcrs2(0))&""" RESQTY="""&trim(iResQty)&""" OHOQTY="""&trim(iOhoQty)&""" REJQTY="""&trim(iRejQty)&""" />"

				dcrs3.Close
			end if
		dcrs2.MoveNext
		Loop
	end if
	dcrs2.Close
%>
</Data>
</script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim objTemp,Root,newElem
dim iClass,iItem,iQty,ii,sLot,sBin,sStore,sOrgID,iInvRec

Function fnInit(obj)
	arrTemp = split(obj,"AAAA")
	sOrgID = arrTemp(1)
	iItem = arrTemp(2)
	iClass = arrTemp(3)
	sLot = arrTemp(4)
	if sLot = "0" then sLot = "NULL"
	sStore = arrTemp(6)
	sBin = arrTemp(7)
	iInvRec = arrTemp(8)

	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement

	if document.formname.hiCtr.value = 0 then exit function

	For Each HeaderNode In Root.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sStore and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				For Each HNode In HeaderNode.childNodes
					if HNode.Attributes.Item(0).nodeValue = sLot and HNode.Attributes.Item(1).nodeValue = iInvRec then
						if HNode.HaschildNodes() then
							For Each INode In HNode.childNodes
								if StrComp(Trim(INode.NodeName),"SERIALDETAILS") = 0 then
								ii = ii + 1
								set Q = eval("document.formname.txtRejQty"&ii)
								Q.value = INode.Attributes.Item(3).nodeValue
								set Q = eval("document.formname.txtOhOQty"&ii)
								Q.value = INode.Attributes.Item(4).nodeValue
								set Q = eval("document.formname.txtResQty"&ii)
								Q.value = INode.Attributes.Item(5).nodeValue
								set Q = eval("document.formname.txtRem"&ii)
								Q.value = INode.Attributes.Item(6).nodeValue
								end if
							next
						end if
						exit function
					end if
				next
			end if
		end if
	Next
end Function

Function CheckSubmit()
	'MsgBox Data.xml
	dim ictr,objQ,iQtyTot,objSTQ,objSerial

	ictr = document.formname.hiCtr.value

	if ictr = "" then exit function
	i = 0
	Set RootO = Data.documentElement

	For Each HeaderONode In RootO.childNodes
		i = i + 1
		set obj = eval("document.formname.txtStQty"&i)
		set objQ = eval("document.formname.txtQty"&i)
		set objRJQ = eval("document.formname.txtRejQty"&i)
		set objOHQ = eval("document.formname.txtOhOQty"&i)
		set objREQ = eval("document.formname.txtResQty"&i)

		if trim(objRJQ.value) = "" then
			msgbox "Enter Rejected Stock",0,"Rejected Stock"
			objRJQ.select()
			exit function
		elseif trim(objOHQ.value) = "" then
			msgbox "Enter On Hold Stock",0,"On Hold Stock"
			objOHQ.select()
			exit function
		elseif trim(objREQ.value) = "" then
			msgbox "Enter On Reserved Stock",0,"On Reserved Stock"
			objREQ.select()
			exit function
		else
			iCQty = 0
			iCQty = cdbl(objRJQ.value) + cdbl(objOHQ.value) + cdbl(objREQ.value)
			iPQty = 0
			iPQty = cdbl(HeaderONode.Attributes.Item(1).nodeValue) + cdbl(HeaderONode.Attributes.Item(2).nodeValue) + cdbl(HeaderONode.Attributes.Item(3).nodeValue)

			iTempQty = cdbl(iCQty) - cdbl(iPQty)
			objQ.value = iTempQty

			if trim(eval("document.formname.txtRem"&i).value) = "" and (cdbl(objQ.value) <> 0) then
				msgbox "Enter Remarks",0,"Remarks"
				eval("document.formname.txtRem"&i).select()
				exit function
			end if

			iQtyTot = cdbl(iQtyTot) + cdbl(objRJQ.value) + cdbl(objOHQ.value) + cdbl(objREQ.value)
			iQtyCLTot = cdbl(iQtyCLTot) + cdbl(objQ.value)
			iQtyRJTot = cdbl(iQtyRJTot) + cdbl(objRJQ.value)
			iQtyOHTot = cdbl(iQtyOHTot) + cdbl(objOHQ.value)
			iQtyRSTot = cdbl(iQtyRSTot) + cdbl(objREQ.value)
		end if

	next

	For Each HeaderNode In Root.childNodes
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

	For Each HeaderNode In Root.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sStore and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				For Each HNode In HeaderNode.childNodes
					if HNode.Attributes.Item(0).nodeValue = sLot and HNode.Attributes.Item(1).nodeValue = iInvRec then
						for i=1 to ictr
							set objSerial = eval("document.formname.txtSerial"&i)
							Set newElem = objTemp.createElement("SERIALDETAILS")
							newElem.setAttribute "SERIALNO", trim(objSerial.value)
							newElem.setAttribute "STKQTY", trim(eval("document.formname.txtStQty"&i&".value"))
							newElem.setAttribute "QTY", trim(eval("document.formname.txtQty"&i&".value"))
							newElem.setAttribute "QTYREJ", trim(eval("document.formname.txtRejQty"&i&".value"))
							newElem.setAttribute "QTYOHO", trim(eval("document.formname.txtOhOQty"&i&".value"))
							newElem.setAttribute "QTYRES", trim(eval("document.formname.txtResQty"&i&".value"))
							newElem.setAttribute "REMARKS", trim(eval("document.formname.txtRem"&i&".value"))
							HNode.appendChild newElem
						next
					end if
				next
			end if
		end if
	Next

	For Each HeaderNode In Root.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sStore and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				For Each HNode In HeaderNode.childNodes
					if HNode.Attributes.Item(0).nodeValue = sLot and HNode.Attributes.Item(1).nodeValue = iInvRec then
						HNode.setAttribute "QTYCLE", cdbl(iQtyCLTot)
						HNode.setAttribute "QTYREJ", cdbl(iQtyRJTot)
						HNode.setAttribute "QTYOHO", cdbl(iQtyOHTot)
						HNode.setAttribute "QTYRES", cdbl(iQtyRSTot)
						window.close
						exit function
					end if
				next
			end if
		end if
	Next

end Function

Function window_onunload()
	set window.returnValue = objTemp.documentElement
	window.close()
end Function


</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>')">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot Serial Details
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
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCell">Lot No.<span class="DataOnly"><%=replace(iLot,"NULL","N/A")%>&nbsp;</span></td>
                                            <td></td>
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
									<div class="frmBody" id="frm2" style="width: 555; height:140;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Serial Number</td>
												<td class="ExcelHeaderCell" align="center">Stock</td>
												<td class="ExcelHeaderCell" align="center">Clean</td>
												<td class="ExcelHeaderCell" align="center">Rejected</td>
												<td class="ExcelHeaderCell" align="center">On Hold</td>
												<td class="ExcelHeaderCell" align="center">Reserved</td>
												<td class="ExcelHeaderCell" align="center">Remarks</td>
											</tr>
										<%
											with dcrs2
												.CursorLocation = 3
												.CursorType = 3
												if iLot = "NULL" then
													.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0),PACKINGNUMBER FROM INV_T_RECEIPTLOTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER IS NULL AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER NOT IN (SELECT SERIALNO FROM INV_T_MRSISSUEPICKSERIAL) ORDER BY 1"
												else
													.Source = "SELECT SERIALNUMBER,ISNULL(LOTQUANTITYNETT,0),ISNULL(QUANTITYISSUED,0),PACKINGNUMBER FROM INV_T_RECEIPTLOTDETAILS WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOTNUMBER = " & Pack(iLot) & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND (ISNULL(LOTQUANTITYNETT,0) - ISNULL(QUANTITYISSUED,0)) > 0 AND SERIALNUMBER NOT IN (SELECT SERIALNO FROM INV_T_MRSISSUEPICKSERIAL) ORDER BY 1"
												end if
												.ActiveConnection = con
												.Open
											end with
											set dcrs2.ActiveConnection = nothing

											if not dcrs2.EOF then
												Do While Not dcrs2.EOF
													iCtr = iCtr + 1
													if not (cdbl(trim(dcrs2(1))) - cdbl(trim(dcrs2(2)))) = 0 then
														with dcrs3
															.CursorLocation = 3
															.CursorType = 3
															if iLot = "NULL" then
																.Source = "SELECT ISNULL(CLEAN,0),ISNULL(RESERVED,0),ISNULL(ONHOLD,0),ISNULL(REJECTED,0) FROM INV_M_STOCKSTATUSDETAILS WHERE STOCKNO = " & iStockNo & " AND LOTNO IS NULL AND SERIALNO = " & trim(dcrs2(0)) & ""
															else
																.Source = "SELECT ISNULL(CLEAN,0),ISNULL(RESERVED,0),ISNULL(ONHOLD,0),ISNULL(REJECTED,0) FROM INV_M_STOCKSTATUSDETAILS WHERE STOCKNO = " & iStockNo & " AND LOTNO = " & Pack(iLot) & " AND SERIALNO = " & trim(dcrs2(0)) & ""
															end if
															.ActiveConnection = con
															.Open
														end with
														set dcrs3.ActiveConnection = nothing

														if not dcrs3.EOF then
															iResQty = trim(dcrs3(1))
															iOhoQty = trim(dcrs3(2))
															iRejQty = trim(dcrs3(3))
															'sRemarks = trim(dcrs3(4))
														else
															iResQty = "0"
															iOhoQty = "0"
															iRejQty = "0"
															sRemarks = ""
														end if
														dcrs3.Close
										%>

											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell">
													<input type="text" name="txtSerial<%=iCtr%>" size="11" maxlength=10 value="<%=trim(dcrs2(3))%>" class="FormElemRead" READONLY>
												</td>
												<td class="ExcelDisplayCell" width="10">
													<input type="text" name="txtStQty<%=iCtr%>" size="11" maxlength=10 value="<%=cdbl(trim(dcrs2(1))) - cdbl(trim(dcrs2(2)))%>" class="FormElemRead" READONLY style="text-align=right">
												</td>
												<td class="ExcelInputCell" width="10">
													<input type="text" name="txtQty<%=iCtr%>" size="11" maxlength=10 value="0" class="FormElem" READONLY style="text-align=right">
												</td>
												<td class="ExcelInputCell" width="10">
													<input type="text" name="txtRejQty<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=sCheck%>',7,3)" value="<%=iRejQty%>" class="FormElem" style="text-align=right">
												</td>
												<td class="ExcelInputCell" width="10">
													<input type="text" name="txtOhOQty<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=sCheck%>',7,3)" value="<%=iOhoQty%>" class="FormElem" style="text-align=right">
												</td>
												<td class="ExcelInputCell" width="10">
													<input type="text" name="txtResQty<%=iCtr%>" size="11" onkeypress="DoKeyPress('<%=sCheck%>',7,3)" value="<%=iResQty%>" class="FormElem" style="text-align=right">
												</td>
												<td class="ExcelInputCell" width="10">
													<input type="text" name="txtRem<%=iCtr%>" size="11" maxlength=50 value="<%=sRemarks%>" class="FormElem" style="text-align=right">
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