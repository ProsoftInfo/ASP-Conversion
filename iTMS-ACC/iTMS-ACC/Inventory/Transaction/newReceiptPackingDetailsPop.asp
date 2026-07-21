
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	newReceiptPackingDetailsPop.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	October 06, 2004
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Packing Detailss</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim objTemp,Root,iItem,iClass

Function fnInit(iClassP,iItemP)
	dim sTemp,iValue
	
	if document.formname.hiCtr.value = 0 then exit function

	set objTemp = window.dialogArguments

	Set Root = objTemp.documentElement
	
	iClass = iClassP
	iItem = iItemP
	
	'alert Root.xml
	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.getNamedItem("ICODE").Value = iItem and HeaderNode.Attributes.getNamedItem("CCODE").Value = iClass then
			For Each SHNode In HeaderNode.childNodes
				if StrComp(Trim(SHNode.NodeName),"STAGE") = 0 then
					ii = ii + 1
					if ii <= document.formname.hiCtr.value then
						set Q = eval("document.formname.txtA"&ii)
						iTemp = trim(SHNode.Attributes.getNamedItem("IVALUE").Value)
						Q.value = iTemp
					end if
				end if
			next
		end if
	Next
end Function

Function CheckSubmit()
	ictr = document.formname.hiCtr.value

	if ictr = "0" then exit function 

	for i=1 to ictr
		set objQ = eval("document.formname.txtA"&i)

		if trim(objQ.value) = "" then
			alert("Enter Value")
			objQ.select()
			exit function
		end if
	next
	ii = 0
	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.getNamedItem("ICODE").Value = iItem and HeaderNode.Attributes.getNamedItem("CCODE").Value = iClass then
			For Each SHNode In HeaderNode.childNodes
				if StrComp(Trim(SHNode.NodeName),"STAGE") = 0 then
					ii = ii + 1
					if ii <= document.formname.hiCtr.value then
						set Q = eval("document.formname.txtA"&ii)
						SHNode.Attributes.getNamedItem("IVALUE").Value = Q.value
						set oQ = eval("idStageQty"&ii)
						SHNode.Attributes.getNamedItem("IQTY").Value = oQ.innerText
					end if
				end if
			next
		end if
	Next

	window.close
	exit function
end Function

Function window_onunload() 
	set window.returnValue = objTemp.documentElement
	'alert(objTemp.xml)
	window.close()
end Function
	

</SCRIPT>

</head>
<%
	dim dcrs,dcrs1
	dim iItem,iClass,sOrgID,arrTemp,iCtr
	dim sItemName,sOrgName,iRecNo,iQty,sUoM,iStage,sStageName,iValue

	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

	arrTemp = split(trim(Request.QueryString("sTemp")),":")

	iRecNo = arrTemp(1)
	iClass = trim(arrTemp(2))
	iItem = trim(arrTemp(3))
	sOrgID = trim(arrTemp(4))
	iQty = arrTemp(5)
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT STORESUOM,ORGUNITDESCRIPTION FROM VWITEM WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sUoM = trim(dcrs(0))
		sOrgName = trim(dcrs(1))
	end if
	dcrs.close

	sItemName = ItemDisplay(iItem,iClass)
%>

<body leftmargin="0" topmargin="0" margin"0" marginwidth="0" onLoad="fnInit('<%=iClass%>','<%=iItem%>')">

	<form method="POST" name="formname" action>
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Packing Details
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" >
					<tr>
						<td class="TabBodywithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" "5">
									</td>
								</tr>

								<tr>
									<td align="center">
									</td>
									<td valign="top">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td class="FieldCell">Item Name</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=sItemName%></span>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Quantity
												</td>
												<td class="FieldCellSub">
													<span class="DataOnly"><%=iQty%>&nbsp;</span>
 													<span class="DataOnly"><%=sUoM%></span>
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
									</td>
									<td valign="top" width="100%">
										<div class="frmBody" id="frm1" style="width: 320; ">
											<table border="0" cellspacing="1" class="ExcelTable" width="100%">
												<tr>
													<td class="ExcelHeaderCell" rowspan="2" width="10">S.No.</td>
													<td class="ExcelHeaderCell">&nbsp;Quality Type</td>
													<td class="ExcelHeaderCell" align="center">Total Qty</td>
													<td class="ExcelHeaderCell" align="center">Value</td>
												</tr>
												<tr>
													<td class="ExcelHeaderCell" align="right">Packing Number</td>
													<td class="ExcelHeaderCell" align="center">Quantity</td>
													<td class="ExcelHeaderCell" align="center"></td>
												</tr>
										<%
											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = "SELECT DISTINCT ISNULL(STAGEID,'0'),ISNULL(SUM(QUANTITYRETURN),0) FROM APP_T_INTERNALRECEIPTDETAILS WHERE INTERNALRECEIPTNO = " & iRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " GROUP BY STAGEID"
												.ActiveConnection = con
												.Open
											end with
											'Response.Write dcrs.Source 
											set dcrs.ActiveConnection = nothing
											do while not dcrs.EOF
												iStage = trim(dcrs(0))
												iValue = "0"

												if iStage = "0" then
													iStage = "0"
												else
													iStage = mid(iStage,1,InStr(1,iStage,":")-1)
												end if

												with dcrs1
													.CursorLocation = 3
													.CursorType = 3
													.Source = "SELECT STAGENAME FROM INV_M_STAGE WHERE STAGEID = " & iStage & ""
													.ActiveConnection = con
													.Open
												end with
												set dcrs1.ActiveConnection = nothing
												if not dcrs1.EOF then
													sStageName = trim(dcrs1(0))
												end if
												dcrs1.close

												with dcrs1
													.CursorLocation = 3
													.CursorType = 3
													.Source = "SELECT RATE FROM INV_M_PACKINGQUALITYRATES WHERE STAGEID = " & iStage & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & ""
													.ActiveConnection = con
													.Open
												end with
												set dcrs1.ActiveConnection = nothing
												if not dcrs1.EOF then
													iValue = trim(dcrs1(0))
												else
													iValue = "0"
												end if
												dcrs1.close

												iCtr = iCtr + 1
										%>
												<tr>
													<td class="ExcelSerial" align="center"><%=iCtr%></td>
													<td class="ExcelDisplayCell"><%=sStageName%></td>
													<td class="ExcelDisplayCell" align="right"><span class="DataOnly" id="idStageQty<%=iCtr%>"><%=trim(dcrs(1))%></span></td>
													<td class="ExcelInputCell" width="10">
														<input type="text" name="txtA<%=iCtr%>" size="12" value="<%=cdbl(iValue) * cdbl(dcrs(1))%>" class="FormElem" style="text-align=right" onkeypress="DoKeyPress('Y',7,3)">
														<input type=hidden name="hStage<%=iCtr%>" value="<%=iStage%>">
													</td>
												</tr>
										<%
												iStage = trim(dcrs(0))

												if iStage = "0" then
													iStage = "0"
												else
													iStage = Pack(iStage)
												end if

												with dcrs1
													.CursorLocation = 3
													.CursorType = 3
													.Source = "SELECT PACKINGCODE,PACKINGNUM,QUANTITYRETURN FROM APP_T_INTERNALRECEIPTDETAILS WHERE INTERNALRECEIPTNO = " & iRecNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND isNull(STAGEID,0) = " & iStage & " ORDER BY INTERNALRECEIPTNO,SERIALNO"
													.ActiveConnection = con
													.Open
												end with
												'Response.Write "<p>aa = " & dcrs1.Source 
												
												set dcrs1.ActiveConnection = nothing
												do while not dcrs1.EOF

										%>
												<tr>
													
													<td class="ExcelSerial" align="center"></td>
													<td class="ExcelDisplayCell" align="right"><%=trim(dcrs1(1))%></td>
													<td class="ExcelDisplayCell" align="right"><%=trim(dcrs1(2))%></td>
													<td class="ExcelDisplayCell"></td>
												</tr>
										<%
												dcrs1.MoveNext
												loop
												dcrs1.close
												
											dcrs.MoveNext
											loop
											dcrs.close
										%>
											</table>
										</div>
									</td>
									<td align="center">
									</td>
								</tr>
								<input type=hidden name=hiCtr value="<%=iCtr%>">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" "31">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" "5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<input type="button" value="Done" name="B2" class="ActionButton" onclick="CheckSubmit()">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
												</td>
											</tr>

										</table>
									</td>
									<td align="center" "31">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" "5">
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
</body>
</html>