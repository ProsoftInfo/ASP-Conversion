<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassServiceAccountPop.asp
	'Module Name				:	Gate Pass - Service
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	April 10, 2006
	'Modified By				:	KUMAR K A
	'Modified By				:	Ragavendran R
	'Modified On				:	September 02,2010
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
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
' Declaration of variables
Dim sForUnit,sItemDesc,sQuery ,sMaterialRcvd
dim iGPNo, iEntryNo,iCtr

Dim dcrs,rsTemp

'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set rsTemp = Server.CreateObject("ADODB.RecordSet")

iGPNo = trim(Request.QueryString("iGPNo"))
sForUnit = trim(Request.QueryString("ForUnit"))

sQuery = "Select isNull(MaterialRcvd,'N') from FORGATEPASSDETAILS where GatePassNo = "& iGPNo
rsTemp.Open sQuery,con
if not rsTemp.EOF then
	sMaterialRcvd = trim(rsTemp(0))
end if 'if not rsTemp.EOF then
rsTemp.Close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : Gate Pass Service</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Output/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="InvData">
<Root></Root>
</script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Function CheckSubmit()

	dim nItemCtr,iTempCtr,ItemNode,sErrorOccur,iDataCtr
	Dim ndRoot
	sToDaysDate = document.formname.hToDaysDate.value
	nItemCtr = document.formname.hItemCtr.value

	if nItemCtr = "0" then exit function

	Set RootO = OutData.documentElement
	RootO.setAttribute "GPNO", trim(document.formname.hGPNo.value)

	sErrorOccur = "N"

	iDataCtr = 0

	For iTempCtr = 1 to nItemCtr
		set Obj = eval("document.formname.hEntryNo" + trim(iTempCtr))
		iEntryNo = obj.value

		set Obj = eval("document.formname.RecdON" + trim(iTempCtr))
		dtRecdON = Obj.value


		if trim(dtRecdON) <> "" then


			if not IsDate(dtRecdON) then
				alert("Enter valid date")
				obj.focus
				sErrorOccur = "Y"
				exit for
			end if

			if ( not checkValidDate(dtRecdON,sToDaysDate,0)) then
				MsgBox "Invalid Date",0,"Invalid Date"
				obj.focus
				sErrorOccur = "Y"
				exit for
			end if


			iDataCtr = iDataCtr + 1

			set ItemNode = OutData.CreateElement("Item")
			ItemNode.setAttribute "EntryNo", iEntryNo
			ItemNode.setAttribute "ReceivedOn",dtRecdON
			RootO.appendChild ItemNode
		end if 'if trim(dtRecdON) <> "" then

	Next

	set ndRoot = InvData.documentElement
	if trim(document.formname.hMaterialRcvd.value)="N" then
		if document.formname.chkInvoice.checked = true then
			set ndInv = InvData.createElement("WithInv")
				'ndInv.setAttribute "Rate",document.formname.radRate.value
			if document.formname.radRate(0).checked = true then
				ndInv.setAttribute "Rate",document.formname.radRate(0).value
			elseif document.formname.radRate(1).checked = true then
				ndInv.setAttribute "Rate",document.formname.radRate(1).value
			end if 'if document.formname.radRate(0).checked = true then
			ndRoot.appendChild ndInv
		end if
	end if ' if trim(document.formname.hMaterialRcvd.value)="N" then


	if sErrorOccur = "Y" then exit function
	'alert(OutData.xml)
	'exit function

	if cint(nItemCtr) = cint(iDataCtr) then
		Set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","GatePassServiceAccountInsert.asp", false
		objhttp.send OutData.XMLDocument

		'alert(objhttp.responseText)
		'exit function
		if objhttp.responseText = "" then
			alert("Gate Pass for Service has been accounted.")
			window.close
		else
			alert(objhttp.responseText)
		end if
	else
		window.close
	end if
end Function
'******************************************************************
Function chkInvoice_onClick()
	Dim nRow,iCnt
	if document.formname.chkInvoice.checked = true then
	'	document.formname.radRate.disabled = false
		document.formname.radRate(0).disabled = false
		document.formname.radRate(1).disabled = false
	else
	'	document.formname.radRate.disabled = true
		document.formname.radRate(0).disabled =	true
		document.formname.radRate(1).disabled = true
	end if
End Function
'******************************************************************
Function CalculateTotal()
	Dim nRow,iCnt
	Dim iItemTotal,iInvoiceTotal,iItemQty,iItemRate
	nRow = document.formname.hItemCtr.value

	for iCnt = 1 to nRow
		if trim(eval("document.formname.txtRate"&iCnt).value) ="" then
			alert("Enter the Rate")
			eval("document.formname.txtRate"&iCnt).value=""
			eval("document.formname.txtRate"&iCnt).focus
			exit function
		elseif trim(eval("document.formname.txtRate"&iCnt).value) ="0" then
			alert("Item Rate must be greater that zero")
			eval("document.formname.txtRate"&iCnt).value=""
			eval("document.formname.txtRate"&iCnt).focus
			exit function
		elseif not IsNumeric(eval("document.formname.txtRate"&iCnt).value) then
			alert("Enter Numeric Value")
			eval("document.formname.txtRate"&iCnt).value=""
			eval("document.formname.txtRate"&iCnt).focus
			exit function
		end if
	next


	for iCnt = 1 to nRow
		iItemQty = eval("document.formname.txtQuantity"&iCnt).value
		iItemRate = eval("document.formname.txtRate"&iCnt).value
		iItemTotal = cdbl(iItemQty)*cdbl(iItemRate)
		iInvoiceTotal = cdbl(iInvoiceTotal) + cdbl(iItemTotal)
		eval("document.formname.txtValue"&iCnt).value = iItemTotal
	next
	document.formname.txtTotalValue.value = iInvoiceTotal
End Function
'******************************************************************
Function window_onunload()
	set window.returnvalue = InvData.documentElement
End Function
'******************************************************************
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type=hidden name="hGPNo" value="<%=iGPNo%>">
<input type=hidden name="hUnit" value="<%=sForUnit%>">
<input type=hidden name="hToDaysDate" value="<%=FormatDate(date())%>" >
<input type=hidden name="hMaterialRcvd" value="<%=sMaterialRcvd%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Gate Pass - Service Details
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
                                    <!--table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Gate Pass Number &nbsp;</td>
                                            <td class="FieldCellSub" colspan="4">
												<span class="DataOnly"><%=iGPNo%>&nbsp;</span>
											</td>
                                        </tr>
                                    </table-->
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <% if sMaterialRcvd="N" then %>
                            <tr>
								<td colspan=3 width=100%>
									<table border=0 cellspacing=0 cellpadding=0 width=100%>
										<tr>
											<td class=FieldCellSub colspan=3><input type=Checkbox name=chkInvoice onClick=chkInvoice_onClick()>With Invoice &nbsp;&nbsp;[<input type=radio name=radRate value="I" checked disabled >Item Wise  &nbsp;&nbsp;
											<input type=radio name=radRate value="C" disabled  >Consolidated&nbsp;]
											</td>
										</tr>
									</table>
								</td>
                            </tr>
                            <% end if 'if sMaterialRcvd="N" then %>
                            <tr>
								<td align="center"></td>
								<td>
									<div class="frmBody" id="frm2" style="width: 550; height:200px;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" >Description</td>
												<td class="ExcelHeaderCell" align="center" >Quantity</td>
												<td class="ExcelHeaderCell" align="center" >UOM</td>
												<td class="ExcelHeaderCell" align="center" >Recd. On</td>
											</tr>
										<%

											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT isNull(DESCRIPTION,''),isNull(QUANTITY,0),isNull(INVOICEDUOM,''),convert(varchar,MaterialRcvdOn,103),isNull(ITEMCODE,0),isNull(CLASSIFICATIONCODE,0),isNull(EntryNo,0) FROM FORGATEPASSDETAILS WHERE GATEPASSNO = " & iGPNo & " order by EntryNo"
												.ActiveConnection = con
												.Open
											end with
											set dcrs.ActiveConnection = nothing

											if not dcrs.EOF then
												Do While Not dcrs.EOF
													iCtr = iCtr + 1

													sItemDesc = ""
													if trim(dcrs(4)) <> "0" then
														With rsTemp
															.CursorLocation = 3
															.CursorType = 3
															.Source = "Select ItemDescription from inv_M_ItemMaster where ItemCode = " & dcrs(4) & " and ClassificationCode = " & dcrs(5) & " and OrganisationCode = '" & sForUnit & "'"
															.ActiveConnection = con
															.Open
														End With
														If Not rsTemp.EOF Then
															sItemDesc = rsTemp(0)
														End If
														rsTemp.Close
													end if 'if trim(dcrs(4)) <> "0" then

													if trim(dcrs(0)) <> "" then
														if trim(sItemDesc) <> "" then
															sItemDesc = sItemDesc & " - " &  trim(dcrs(0))
														else
															sItemDesc = trim(dcrs(0))
														end if
													end if


										%>

											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell"><%=sItemDesc%></td>
												<td class="ExcelDisplayCell" align="right">
													<input type="text" name="txtQuantity<%=iCtr%>" value="<%=trim(dcrs(1))%>" class="FormElemRead" size=7 style="text-align:right">
												</td>
												<td class="ExcelDisplayCell" width="20"><%=trim(dcrs(2))%></td>
												<td class="ExcelInputCell" width="30">
													<Input type="text" name="RecdON<%=iCtr%>" value="<%=trim(dcrs(3))%>" maxlength="10" class="FormElem" size="11" >
													<input type=hidden name="hEntryNo<%=iCtr%>" value="<%=dcrs(6)%>">
												</td>
											</tr>
										<%
												dcrs.MoveNext
												Loop
											end if
											dcrs.Close
										%>
										</table>
										<input type=hidden name="hItemCtr" value="<%=iCtr%>">
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
                                                    <input type="button" value="Account" name="B1" class="ActionButton" onClick="CheckSubmit()">
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