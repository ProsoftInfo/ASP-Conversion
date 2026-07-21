<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtArrangeBinEntry.asp
	'Module Name				:	Inventory (Stock Management Arrange Bin Entry)
	'Author Name				:	UmaMaheswari S
	'Created On					:	May 31, 2011
	'Modified By				:	
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock Management - Arrange Bin Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Output/></script>
<script type="application/xml" data-itms-xml-island="1" id="IssueData"><ISSTYPE></ISSTYPE></script>
<script type="application/xml" data-itms-xml-island="1" id="IntReceipt"><ROOT></ROOT></script>
<script type="application/xml" data-itms-xml-island="1" id="NewData"><ROOT/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,OutDataValue
dim sorgID,iClass,sStore,iInvRec,sLot,sBin

Function FnInit(sItemCode)
	'document.formname.selItem.value = sItemCode
	GetXML()
End Function

Function GetXML()
	ClearTable
	clearXML

'	set obj = document.formname.selItem

'	sorgID = trim(document.formname.hOrgID.value)
'	iClass = trim(document.formname.hClass.value)

'	if obj.value <> "select" then
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		'objhttp.Open "GET","itmStoreXMLSelect.asp", false
		'objhttp.Open "GET","itmStoreXMLSelectNew.asp", false
		objhttp.Open "GET","ArrangeBinDetXML.asp", false
		objhttp.send
		'alert(objhttp.responseText)
		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			DisplayDetails()
			popClaDisplay
		else
			clearXML
		end if
'	end if
End Function

Function CheckLot(obj)
	'sTemp = document.formname.selItem(document.formname.selItem.selectedIndex).text & " -- " & trim(idClass.innerText)

	set OutDataValue = showModalDialog("stkMgmtSTPoP.asp?sTemp="&obj.name&"&sValue="&sTemp,OutData,"dialogHeight:310px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	
	arrTemp = split(obj.name,"AAAA")
	sOrgID = arrTemp(1)
	iItem = arrTemp(2)
	iClass = arrTemp(3)
	sLot = arrTemp(4)
	if sLot = "0" then sLot = "NULL"
	sStore = arrTemp(6)
	sBin = arrTemp(7)
	iInvRec = arrTemp(8)

	Set RootO = OutData.documentElement
	For Each HeaderONode In RootO.childNodes
		if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
			if (HeaderONode.Attributes.Item(0).nodeValue = sStore and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
				For Each PickNode In HeaderONode.childNodes
					i = i + 1
					set objQ = eval("document.formname.txtQtyA"&i)
					'objQ.value = "0"

					if PickNode.Attributes.Item(0).nodeValue = sLot and PickNode.Attributes.Item(1).nodeValue = iInvRec then
						objQ.value = PickNode.Attributes.Item(3).nodeValue
					end if
					'if objQ.value = "" then objQ.value = "0"
				next
				exit for
			end if
		end if
	next
End Function

Function DisplayDetails()
	Dim sRecNumStatus
	ClearTable
	j = 0
	Set Root = OutData.documentElement
	Set RootO = OutData.documentElement
	
	'For Each ndItem In RootO.childNodes
	'	For each HeaderONode in ndItem.childNodes
	'		if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
	'			For Each PickNode In HeaderONode.childNodes
	'				PickNode.setAttribute "QTYISS", ""
	'				PickNode.setAttribute "STORE", ""
	'			next
	'		end if
	'		if StrComp(Trim(HeaderONode.NodeName),"UOM") = 0 then
	'			sCheck = HeaderONode.Attributes.Item(2).nodeValue
	'		end if
	'	Next
	'Next
	i = 1
	
	
	For Each ndItem In Root.childNodes
		sItemName = ndItem.getAttribute("IName")
		iItem = ndItem.getAttribute("ICode")
		iClass = ndItem.getAttribute("CCode")
		sorgID = ndItem.getAttribute("Unit")
		sRecNumStatus = ndItem.getAttribute("RecNumStatus")

		For Each HeaderNode in ndItem.childNodes
			sStoreName = HeaderNode.getAttribute("LOCNAME")
			if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
				sLoc = HeaderNode.getAttribute("LOC")
				nStkQty = HeaderNode.getAttribute("STOCK")
				sBinStatus = HeaderNode.getAttribute("BINSTATUS")
				'sBin = HeaderNode.getAttribute("BIN")
					'For Each ItemNode In HeaderNode.childNodes
				
						j = j + 1
						'set oRow = document.all.tblData.insertRow(j+1)
						set oRow = document.all.tblData.insertRow(j)

						set headerCell=oRow.insertCell()
						headerCell.innerHTML=j
						headerCell.className="ExcelSerial"
						headerCell.align="center"

						set headerCell=oRow.insertCell()
						headerCell.innerHTML=sItemName
						headerCell.className="ExcelDisplayCell"
						headerCell.align="Left"

						set headerCell=oRow.insertCell()
						headerCell.innerHTML=sStoreName
						headerCell.className="ExcelDisplayCell"
						headerCell.align="Left"
						
						
						set headerCell=oRow.insertCell()
						set oText = document.createElement("<input type=""text"" name=""txtStk"&CStr(j)&""" size=""12"" maxlength=""10"" value="""&nStkQty&""" READONLY class=""FormelemRead"" style=""text-align=right"">" )
						headerCell.appendChild(oText)
						headerCell.className="ExcelDisplayCell"
						'headerCell.width = "10"
						
						If InStr(1,sStoreName,"&") Then
							sStoreName = Replace(sStoreName,"&","AND")
						End IF
						sItemName = Replace(trim(sItemName),chr(39),"~~")
						sItemName = Replace(trim(sItemName),Chr(34),"``")
						
						sTemp=iItem & ":" & iClass & ":" & sItemName & ":" & sStoreName & ":" & nStkQty & ":" & sLoc & ":" & sBinStatus
						
						set headerCell=oRow.insertCell()
						set oText = document.createElement("<a href=""#"">" )

						set oText1 = document.createElement("<img name=""btn"" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" width=""15"" height=""15"" value="""& sTemp&""" onClick=""BinStockDet(this)"">")
						oText.appendChild(oText1)

						headerCell.appendChild(oText)
						headerCell.className="ExcelDisplayCell"
						headerCell.align="center"
						'headerCell.width = "10"
						
					'next
			end if
		Next
	next
end Function

Function CheckQty(nStkQty,nVal)
	nIssueQty = cdbl(Eval("document.formname.txtQtyA"&nval).value)
	If nIssueQty > cdbl(nStkQty) Then
		alert("Enter Issue Qty which is less than Stock Qty")
		Eval("document.formname.txtQtyA"&nval).value = "0"
		Exit Function
	End IF
End Function

Function BinStockDet(sPassValue)
	Dim sTempArr,RootN

	sTempArr = Split(sPassValue.value,":")
	nItemCode = trim(sTempArr(0))
	nClassCode = trim(sTempArr(1))
	nBinNo = sTempArr(6)
	
	If nBinNo = "N" Then
		alert("Create the Bin No for Selected Item")
		Exit Function
	End IF
	set Root  = OutData.documentElement
	set RootN = NewData.documentElement
	
	If Root.haschildNodes Then
		For Each Node in Root.childNodes
			If Node.nodeName = "Item" Then
				If Node.getAttribute("ICode") = nItemCode and Node.getAttribute("CCode") = nClassCode Then
					RootN.appendchild Node
				End IF
			End IF
		Next 
	End IF
	
	Outvalue = showModalDialog("ArrangeBinDetailsEntry.asp?Data="&trim(sPassValue.value),NewData,"","dialogHeight:350px;dialogWidth:400px;center:Yes;status:No")
	If Outvalue = "Done" Then
	End IF
	'window.open("ArrangeBinDetailsEntry.asp")
End Function

Function popClaDisplay()

	'document.formname.selStore.options.length = 1
	Set Root = OutData.documentElement

	For Each Node In Root.childNodes
		For Each HeaderNode in Node.childNodes
			if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then

				if not HeaderNode.Attributes.Item(1).nodeValue = "0" then
				'	document.formname.selStore.length = document.formname.selStore.length+1
				'	document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)&" -- "&trim(HeaderNode.Attributes.Item(3).nodeValue)
				'	document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)&":"&trim(HeaderNode.Attributes.Item(1).nodeValue)
				else
				'	document.formname.selStore.length = document.formname.selStore.length+1
				'	document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)
				'	document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)
				end if
			end if
			if StrComp(Trim(HeaderNode.NodeName),"UOM") = 0 then
				idUoM.innerHTML = trim(HeaderNode.Attributes.Item(1).nodeValue) & "&nbsp;"
			end if
		Next
	next

end Function

Function clearXML()
	Set Root = OutData.documentElement
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			set a=Root.removeChild(HeaderNode)
		next
	end if
	'document.formname.selStore.options.length = 1
end Function

Function ClearTable()
	dim i
	for i=2 to document.all.tblData.rows.length - 1
		document.all.tblData.deleteRow(2)
	next
end Function

Function CheckSubmit()
	dim objQ,iQtyTot
	Set Root = OutData.documentElement
	if j = 0 then exit function
	'alert(Root.xml)

	iLen =  document.formname.selSTA1.length

'	alert(iLen)
	if cint(iLen) < 1 then exit function
'
'	iSTNo = trim(document.formname.selStore.value)
'
'	arrTemp1 = split(iSTNo,":")
'	if ubound(arrTemp1) > 0 then
'		sLoc = arrTemp1(0)
'		sBin = arrTemp1(1)
'	else
'		sLoc = arrTemp1(0)
'		sBin = "0"
'	end if

	For Each ndItem In Root.childNodes
		For Each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
				if (HeaderONode.Attributes.Item(0).nodeValue = sLoc and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
					For Each PickNode In HeaderONode.childNodes
						i = i + 1
						set objQ = eval("document.formname.txtQtyA"&i)
						if trim(objQ.value) = "" then
							alert("Enter Quantity")
							objQ.select()
							exit function
						else
							set obj = eval("document.formname.txtStk"&cstr(i))

							if (cdbl(objQ.value) > cdbl(obj.value)) then
								alert("Transfer Quantity should be equal to or less than Stock Quantity")
								objQ.select()
								exit function
							end if

							iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)
						end if
					next
				end if
			end if
		Next
	Next

	i = 1

	Set RootO = OutData.documentElement
	For Each ndItem In RootO.childNodes
		For Each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
			'	if (HeaderONode.Attributes.Item(0).nodeValue = sLoc and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
					For Each PickNode In HeaderONode.childNodes
						set objQ = eval("document.formname.txtQtyA"&i)
						PickNode.setAttribute "QTYISS", cdbl(objQ.value)

						set objQ = eval("document.formname.selSTA"&cstr(i))
						PickNode.setAttribute "STORE", objQ.value
						i = i + 1
					next
			'	else
			'		set a = RootO.removeChild(HeaderONode)
			'	end if
			end if
		Next
	Next

'	sorgID = trim(document.formname.hOrgID.value)
'	iClass = trim(document.formname.hClass.value)

'	RootO.setAttribute "ITEM", trim(document.formname.selItem.value)
'	RootO.setAttribute "CLASS", iClass
'	RootO.setAttribute "ORG", sorgID
	alert(OutData.xml)
	exit function
	document.formname.B7.disabled = True
	
	GenIssueXML()
	GenReceiptXML()
	

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","stkMgmtSTInsert.asp", false
	objhttp.send OutData.XMLDocument

	'alert objhttp.responseText
	'exit function

	if objhttp.responseText = "" then
		Msgbox ("Stock Transfer has done")
		'document.formname.B7.disabled = true
		'if document.formname.hCallFrom.value = "ItemList" then
			window.location.href = "../master/ITEMLISTENTRY.ASP"
		'else
		'	window.location.href = "stkMgmtSTEntry.asp"	'"stkMgmtEntry.asp"
		'end if
	else
		alert(objhttp.responseText)
		'document.formname.B7.disabled = False
	end if

end Function


Function Submit()
	document.formname.action = "../MASTER/ITEMLISTENTRY.ASP?ACTN=M"
	document.formname.submit 
End Function

</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>

<%
	dim iCtr,arrTemp,sTemp,arrValue,sOrgID,iClass,arrTempName,sTempName
	dim sOrgName,sClassName,rsTemp


	set rsTemp = server.CreateObject("ADODB.Recordset")

	sOrgName = trim(Request.Form("hOrgName"))
	sClassName = trim(Request.Form("hClassName"))
	'sOrgID = trim(Request.Form("selUnit"))
	sOrgID =session("organizationcode")
	iClass = trim(Request.Form("selClass"))
	sTemp = trim(Request.Form("hSelectedValue"))
	sTempName = trim(Request.Form("hItemNames"))
	
	
	
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
	
	'Response.Write "<p><font color=red>sTempName="&sTempName & "-"& sTemp 
	
	arrTempName = split(mid(sTempName,1,len(sTempName)-1),"|")
	arrTemp = split(mid(sTemp,1,len(sTemp)-1),"|")

%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="FnInit('<%= mid(sTemp,1,len(sTemp)-1)%>')">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hClass" value="<%=iClass%>">

<input type="hidden" name="hCallFrom" value="<%=Request.Form("hCallFrom")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Arrange Bin Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<!--<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
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
						</table>
					</td>
				</tr>-->
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
                                        <!--<tr>
                                            <td class="FieldCell">Organization</td>
                                            <td class="FieldCellSub">
	                                            <span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                                            </td>
                                        </tr>-->
                                        <tr>
                                            <td class="FieldCell">Classification</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idClass"><%=sClassName%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <!--<tr>
                                            <td class="FieldCell">Item </td>
                                            <td class="FieldCellSub" colspan=4>
												<select size="1" name="selItem" class="FormElem" onChange="GetXML()">
													<option value="select">Select</option>
												<%
													for iCtr = 0 to UBound(arrTempName)
														If iCtr = 0 Then
													%>
														<option value="<%=arrTemp(iCtr)%>" selected><%=arrTempName(iCtr)%></option>
													<%
														Else
													%>
														<option value="<%=arrTemp(iCtr)%>"><%=arrTempName(iCtr)%></option>
													<%	End IF
													next
												%>
												</select>
                                            </td>
                                        </tr>-->
                                        <tr>
                                            <td class="FieldCell">UoM</td>
                                            <td class="FieldCellSub">
	                                            <span class="DataOnly" id="idUoM"></span>
                                            </td>
                                        </tr>
                                        <!--<tr>
                                            <td class="FieldCell">Store -- Bin</td>
                                            <!--<td class="FieldCellSub">-->
												<!--<select size="1" name="selStore" class="FormElem" onChange="DisplayDetails(this.value)">-->
												<!--<select size="1" name="selStore" class="FormElem">
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
                                    <div class="frmBody" id="frm2" style="width: 580; height:250;">
                                        <table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
                                                <td class="ExcelHeaderCell" align="center">Item Name</td>
                                                <td class="ExcelHeaderCell" align="center">Store</td>
                                                <td class="ExcelHeaderCell" align="center">Stock Qty</td>
                                                <td class="ExcelHeaderCell" align="center"></td>
                                            </tr>
                                            
                                        </table>
                                    </div>
								</td>
								<td align="center"></td>
							</tr>
							<tr>
								<td align="left" colspan="3" class="FieldCell">&nbsp; <B>Note: The above stock shown excludes the quantity already reserved for Issue</B></td>
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
                                                    <!--<input type="button" value="Transfer" name="B7" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="reset" value="Reset" name="B1" class="ActionButton">
                                                    <input type="button" value="Cancel" name="B1" class="ActionButton" onClick="Cancel('stkMgmtEntry.asp')">-->
                                                    <input type="button" value="Done" name="B7" class="ActionButton" onClick="Submit()">
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
