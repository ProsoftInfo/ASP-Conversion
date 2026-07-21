<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MatConLotSerPop.asp
	'Module Name				:	Inventory (Receipt Creation)
	'Author Name				:	Ragavendran
	'Created On					:	Jun 20,2011
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
	dim dcrs,dcrs1,iItem,iClass,sOrgID,iMRSNo,arrTemp,iIssNo
	dim sItmName,sClassName,iQty,sTemp,sAlt,iLot
	dim sRead,iCtr,dIssDate,sTitle,sVar,iLineNo
	dim arrUoM,sUoMDesc,sUoMCode,sType,iDINo,sSQL,iAccHead,sRecptNum,sAttributeList,sCallFrom
	iAccHead = 0
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")

	iCtr = 0
	Response.Write "<font color=red>"
	'Response.Write Request.QueryString
	arrTemp = split(trim(Request.QueryString("sTemp")),":")
	sType   = arrTemp(0)
	iIssNo  = arrTemp(1)
	iItem	= arrTemp(2)
	iClass	= arrTemp(3)
	sOrgID	= arrTemp(4)
	sTitle  = arrTemp(5)
	sVar    = arrTemp(6)
	sAttributeList = arrTemp(7)
	if UBound(arrTemp)>7 then
	    sCallFrom = arrTemp(9)
'	    Response.Write "CallFrom="& sCallFrom
	end if
	if trim(sCallFrom)="ITEM" then
	    sSQL = "Select ReceiptNumbering from VwItem where ItemCode ="& iItem &" and ClassificationCode = "& iClass
	    'Response.Write sSQL
	    dcrs.open sSQL,con
	    if not dcrs.eof then
	        sRecptNum = trim(dcrs(0))
	    end if
	    dcrs.close
	end if 'if trim(sCallFrom)="ITEM" then

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - <%=sTitle%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="ItemData">
<Root>
<%
    if ucase(trim(sCallFrom))="ITEM" then
      sSQL = "Select SUM(IsNull(QuantityIssued,0)-(IsNull(QuantityConsumed,0)+"&_
               " IsNull(QuantityReturned,0))),isNull(LOTNO,'NULL'),Convert(Varchar,IssueDate,103)"&_
               " from INV_T_MaterialIssueDetails D,Inv_T_MaterialIssueHeader H "&_
               " where D.IssueEntryNo = H.IssueEntryNo and D.IssueEntryNo = "& iIssNo &" and D.ItemCode = "& iItem &" and D.ClassificationCode = "& iClass &" and "&_
               " (IsNull(QuantityIssued,0)-(IsNull(QuantityConsumed,0) + IsNull(QuantityReturned,0)))>0 "&_
               " Group by Convert(varchar,IssueDate,103),LotNo Order By 3,2"
    else
      sSQL = "Select SUM(IsNull(QuantityIssued,0)-(IsNull(QuantityConsumed,0)+"&_
               " IsNull(QuantityReturned,0))),isNull(LOTNO,'NULL'),Convert(Varchar,IssueDate,103)"&_
               " from INV_T_MaterialIssueDetails D,Inv_T_MaterialIssueHeader H "&_
               " where D.IssueEntryNo = H.IssueEntryNo and D.IssueEntryNo = "& iIssNo &" and "&_
               " (IsNull(QuantityIssued,0)-(IsNull(QuantityConsumed,0) + IsNull(QuantityReturned,0)))>0 "&_
               " Group by Convert(varchar,IssueDate,103),LotNo Order By 3,2"
    end if 'if trim(sCallFrom)="ISSNO" then
    'response.write "<p>"&sSQL
    dcrs.open sSQL,con
	if not dcrs.EOF then
		do while not dcrs.EOF
			iCtr = iCtr + 1

			iQty = dcrs(0)
			iLot = dcrs(1)
			dIssDate = dcrs(2)
			if iLot="" or iLot = "N/A" or iLot="NULL" then iLot = "0"

			Response.Write "<LotDet LotNo="""&trim(iLot)&""" IssueDate="""&trim(dIssDate)&""" QtyRem="""&iQty&""" QtyRet="""" LineNo="""" AccHead=""""/>"
		dcrs.MoveNext
		loop
	end if
	dcrs.Close
%>
</Root>
</script>
<%
    if trim(sCallFrom)="ISSNO" then
        sSQL ="Select GroupName,isNull(ShortDescription,ItemDescription),(IsNull(SUM(QuantityIssued),0)- "&_
          " (IsNull(SUM(QuantityConsumed),0)+IsNull(SUM(QuantityReturned),0))) from INV_T_MaterialIssueDetails D,"&_
          " VwItem V where D.ItemCode = V.ItemCode and D.ClassificationCode = V.ClassificationCode "&_
          " and D.OrganisationCode = "& Pack(sorgID) &" and"&_
          " IssueEntryNo = "& iIssNo
        sSQL = sSQL & " Group by GroupName,ShortDescription,ItemDescription"
    else
        sSQL ="Select GroupName,isNull(ShortDescription,ItemDescription),(IsNull(SUM(QuantityIssued),0)- "&_
          " (IsNull(SUM(QuantityConsumed),0)+IsNull(SUM(QuantityReturned),0))) from INV_T_MaterialIssueDetails D,"&_
          " VwItem V where D.ItemCode = V.ItemCode and D.ClassificationCode = V.ClassificationCode "&_
          " and D.ItemCode ="& iItem &" and D.ClassificationCode ="& iClass &" and D.OrganisationCode = "& Pack(sorgID) &" and"&_
          " IssueEntryNo = "& iIssNo
        sSQL = sSQL & " Group by GroupName,ShortDescription,ItemDescription"
    end if'if trim(sCallFrom)="ISSNO" then

'    Response.write "<textarea>"& sSQL &"</textarea>"
	dcrs.open sSQL,con
	if not dcrs.EOF then
		sItmName = sItmName &","& trim(dcrs(1))
		iQty = cdbl(iQty) + cdbl(dcrs(2))
	end if
	dcrs.Close
	if trim(sItmName)<>"" then
	    sItmName = mid(sItmName,2)
	end if


	if trim(sCallFrom)="ITEM" then
	arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
	'Response.Write  sOrgID&iClass
	sUoMCode = arrUoM(0)
	sUoMDesc = arrUoM(1)
	end if 'if trim(sCallFrom)="ITEM" then

%>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim objTemp,Root,newElem,iQtyTot
dim iClass,iItem,iIssNo,iMRSNo,sOrgID,sType,iDINo

Function checkNumbers(val)
	dim valid,temp,i
	valid = "0123456789."
	for i=1 to len(val)
		temp = mid(val,i,1)
		if Instr(1,valid,temp) > 0 then
			checkNumbers = true
		else
			checkNumbers = false
			exit for
		end if
	next
end Function

'*************
    Function fnInit(obj)
	arrTemp = split(obj,":")
	sType   = arrTemp(0)
	iIssNo  = arrTemp(1)
	iItem	= arrTemp(2)
	iClass	= arrTemp(3)
	sOrgID	= arrTemp(4)

	set objTemp = window.dialogArguments
	Set Root = objTemp.documentElement

	set RootO = ItemData.documentElement
	if document.formname.hiCtr.value = 0 then exit function
	For Each HeaderNode In Root.childNodes
		if HeaderNode.HaschildNodes() then
			For Each HNode In HeaderNode.childNodes
				if StrComp(Trim(HNode.NodeName),"LotDet") = 0 then
					ii = ii + 1
					set Q = eval("document.formname.txtQty"&ii)
					Q.value = HNode.Attributes.Item(3).nodeValue
					for Each Node in RootO.ChildNodes
						if Node.Attributes.getNamedItem("LotNo").Value = HNode.Attributes.getNamedItem("LotNo").Value and Node.Attributes.getNamedItem("IssueDate").Value = HNode.Attributes.getNamedItem("IssueDate").Value then
							for Each SNode in HNode.ChildNodes
								Node.appendChild SNode
							next
							Node.attributes.getNamedItem("QtyRet").value = HNode.attributes.getNamedItem("QtyRet").value
						end if
					next
				end if
			next
		end if
	Next

end Function
'*************************************************
Function CheckLot(obj,ictr)
'alert(obj.name)
	sTempValues = obj.name
	set OutDataValue = showModalDialog("MatConSerialPop.asp?sTemp="&sTempValues,ItemData,"dialogHeight:400px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(obj.name,"`")
	sType  = arrTemp(1)
	sOrgID = arrTemp(2)
	iItem = arrTemp(3)
	iClass = arrTemp(4)
	sLot = arrTemp(5)
	if trim(sType) = "M" then
		iMRSNo = arrTemp(6)
	else
		iDINo = arrTemp(6)
	end if
	iQty = arrTemp(7)
	dIssDate = arrTemp(9)

	set Q = eval("document.formname.txtQty"&ictr)
	Set ndRoot = ItemData.documentElement
	if ndRoot.hasChildNodes() then
	    for each ndLotDet in ndRoot.childNodes
	        if ndLotDet.nodeName="LotDet" then
	            if trim(ndLotDet.getAttribute("LotNo"))=trim(sLot) and trim(ndLotDet.getAttribute("IssueDate"))=trim(dIssDate) then
	            iSerQty = 0
	                for each ndSer in ndLotDet.childNodes
	                    iSerQty = cdbl(iSerQty)+cdbl(ndSer.getAttribute("Qty"))
	                next
	            end if
	        end if
	    next
	end if
	Q.value = iSerQty
End Function

Function CheckSubmit()
	dim ictr,objQ,objSTQ,objSerial

	ictr = document.formname.hiCtr.value

	if ictr = "" then exit function

	for i=1 to ictr
		'alert i
		set objQ = eval("document.formname.txtQty"&i)
		set objRemQ = eval("document.formname.txtRemQty"&i)

		if trim(objQ.value) = "" then
			msgbox "Enter Quantity to Return",0,"Quantity"
			objQ.select()
			exit function
		elseif not checkNumbers(objQ.value) then
			msgbox "Enter Numerals Only",0,"Numerals"
			objQ.select()
			exit function
		else

			if (cdbl(objQ.value) > cdbl(objRemQ.value)) or (cdbl(objQ.value) > cDbl(document.formname.idQty.value))  then
				msgbox "Return Quantity should be equal to or less than Remaining Quantity",0,"Quantity"
				objQ.select()
				exit function
			end if

			iQtyTot = cdbl(iQtyTot) + cdbl(objQ.value)

		end if

	next

	ictr = 0


	Set Root = ItemData.documentElement
	Set RootO = objTemp.documentElement


	objIss = document.formname.hObjVal.value
    arrTemp = split(objIss,":")
    sType   = arrTemp(0)
    iIssNo  = arrTemp(1)
    iItem	= arrTemp(2)
    iClass	= arrTemp(3)
    sOrgID	= arrTemp(4)

    if RootO.hasChildNodes() then
        for each ndIss in RootO.childNodes
            if ndIss.getAttribute("Item")=trim(iItem) and ndIss.getAttribute("Class")=trim(iClass) and ndIss.getAttribute("IssEntryNo")=trim(iIssNo) then
                set ndIssNode = ndIss
            end if
        next
    end if

	For Each HeaderNode In Root.childNodes
		ictr = ictr + 1
		HeaderNode.setAttribute "QtyRet",trim(eval("document.formname.txtQty"&ictr&".value"))
		ndIssNode.appendChild HeaderNode
	Next

	window.close
end Function

Function window_onunload()
	set window.returnValue = objTemp.documentElement
	window.close()
end Function

</script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0  onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>')">

<form method="POST" name="formname" action="">
<input type=hidden name="hCallFrom" value="<%=sCallFrom%>">
<input type="hidden" name="hObjVal" value="<%=Request.QueryString("sTemp")%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Lot / Serial Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%">
				<TR>
					<TD class="TabBodyWithTopLine">
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
                                            <td class="FieldCell">Item Description</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idItemName"><%=sItmName%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity Remaining</td>
                                            <td class="FieldCellSub">
												<input type="text" name="idQty" size="5" value="<%=iQty%>" class="FormElemRead" READONLY style="text-align:right">
                                                <span class="DataOnly"><%=sUoMDesc%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Issue Date</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=dIssDate%>&nbsp;</span>
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
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<DIV class=frmBody id=frm6 style="width: 100%; height:150;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<!--<td class="ExcelHeaderCell" align="center">Issue Date</td>-->
												<%if sRecptNum<>"N" then %>
												    <td class="ExcelHeaderCell" align="center">Lot No.</td>
												<%end if %>
												<td class="ExcelHeaderCell" align="center">Quantity</td>
												<%if sRecptNum<>"N" then %>
												<td class="ExcelHeaderCell" align="center">Serial</td>
												<%end if %>
												<td class="ExcelHeaderCell" align="center">Quantity <%=sVar%></td>
											</tr>
											<%
													sSQL = "SELECT IssueEntryNo FROM INV_T_MaterialIssueDetails WHERE IssueEntryNo = "& iIssNo &" "&_
													       " AND (ISNULL(QUANTITYISSUED,0) - (ISNULL(QUANTITYCONSUMED,0) + ISNULL(QUANTITYRETURNED,0))) > 0 AND "&_
													       " SERIALNO IS NOT NULL"
													dcrs.open sSQL,con
													if not dcrs.EOF then
														sTemp = ""
													else
														sTemp = " DISABLED "
													end if
													dcrs.close

													iCtr = 0
													if trim(sCallFrom)="ISSNO" then
													        sSQL = "SELECT SUM(ISNULL(QUANTITYISSUED,0) - (ISNULL(QUANTITYCONSUMED,0) + ISNULL(QUANTITYRETURNED,0))),ISNULL(LOTNO,'NULL'),CONVERT(CHAR,ISSUEDATE,103)"&_
													    "  FROM INV_T_MaterialIssueDetails D,INV_T_MaterialIssueHeader H WHERE D.IssueEntryNo = H.IssueEntryNo and D.IssueEntryNo = "& iIssNo &" and "&_
													    "  (ISNULL(QUANTITYISSUED,0) - (ISNULL(QUANTITYCONSUMED,0) + ISNULL(QUANTITYRETURNED,0))) > 0 GROUP BY CONVERT(CHAR,ISSUEDATE,103),LOTNO ORDER BY 3,2"
													else
													    sSQL = "SELECT SUM(ISNULL(QUANTITYISSUED,0) - (ISNULL(QUANTITYCONSUMED,0) + ISNULL(QUANTITYRETURNED,0))),ISNULL(LOTNO,'NULL'),CONVERT(CHAR,ISSUEDATE,103)"&_
													    "  FROM INV_T_MaterialIssueDetails D,INV_T_MaterialIssueHeader H WHERE D.IssueEntryNo = H.IssueEntryNo and D.IssueEntryNo = "& iIssNo &" and "&_
													    " D.ItemCode ="& iItem &" and D.ClassificationCode = "& iClass &" AND (ISNULL(QUANTITYISSUED,0) - (ISNULL(QUANTITYCONSUMED,0) + ISNULL(QUANTITYRETURNED,0))) > 0 GROUP BY CONVERT(CHAR,ISSUEDATE,103),LOTNO ORDER BY 3,2"
													end if 'if trim(sCallFrom)="ISSNO" then
													'Response.Write sSQL
													dcrs.open sSQL,con
													if not dcrs.EOF then
														do while not dcrs.EOF
															iCtr = iCtr + 1
															iQty = dcrs(0)
															iLot = dcrs(1)
															dIssDate = dcrs(2)
															if iLot = "NULL" then
																iLot = "N/A"
																'sTemp = " DISABLED "
																sAlt = ""
																sRead = ""
															else
																'sTemp = ""
																sAlt = "Serial Details"
																sRead = " READONLY "
															end if

															if sRecptNum<>"N" then
															    sRead =" READONLY"
															else
															    sRead = ""
															end if

											%>
											<tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<!--<td class="ExcelDisplayCell" align="center" width="10"><%=dIssDate%></td>-->
												<%if sRecptNum<>"N" then %>
												    <td class="ExcelDisplayCell" align="center"><%=iLot%></td>
												<%end if %>
												<td class="ExcelDisplayCell" width="10">
													<input type="text" name="txtRemQty<%=iCtr%>" size="10" value="<%=iQty%>" class="FormElemRead" READONLY style="text-align:right">
												</td>
												<%if sRecptNum<>"N" then %>
												    <td class="ExcelDisplayCell" align="right">
													    <a href="#">
														    <img name="btn`<%=sType%>`<%=sOrgID%>`<%=iItem%>`<%=iClass%>`<%=iLot%>`<%=iDINo%>`<%=iQty%>`<%=iIssNo%>`<%=dIssDate%>" <%=sTemp%> border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" width="15" height="15" alt="<%=sAlt%>" onClick="CheckLot(this,'<%=iCtr%>')">
													    </a>
												    </td>
												<%end if %>
												<td class="ExcelInputCell" align="right">
												        <input type="text" name="txtQty<%=iCtr%>" size="11" value="0" class="FormElem" <%=sRead%> style="text-align:right" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)">
												</td>
											</tr>
											<%			dcrs.MoveNext
														loop
													end if
													dcrs.Close
											%>
										</table>
									</div>
								</td>
								<input type="hidden" name="hiCtr" value="<%=iCtr%>">
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Close" name="B2" class="ActionButton" onclick="window.close()">
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
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			.ActiveConnection = con
			.Open
		end with
		'Response.Write dcrs.source
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Account Head list
	Function populateAccountHead()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT DISTINCT ACCOUNTHEAD,ACCOUNTHEADCODE FROM VWORGGLHEADS WHERE OUDEFINITIONID = " & Pack(sOrgID) & " AND ACCOUNTHEAD IN (SELECT ACCOUNTHEAD FROM ACC_R_GLACCAPPLICATIONS WHERE AVAILABLEINAPPLN IN (4,5,6) AND OUDEFINITIONID = " & Pack(sOrgID) & ") ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				stypID = dcrs(0)
				stypName = dcrs(1)
				if cint(iAccHead) = cint(stypID) then
					Response.Write("<OPTION VALUE="""&trim(stypID)&""" SELECTED>"&trim(stypName)&"</OPTION>" &vbcrlf)
				else
					Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				end if
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>
