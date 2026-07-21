<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	mrsIssuePickItemList.asp
	'Module Name				:	Inventory (Pick Issue)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 21,2012
	'Modified By                :   
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
<!-- #include File="../../include/Populate.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>MR Pick Issue - Item List</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<%
	dim dcrs,dcrs1,iCtr,sMRNo,sOrg,sItem,sClass,iQtyToIssue,iQtyIssued,sClassName,sItemName
	dim iSchQty,sReqBy,sReqValue,dMRDate,bFlag,iTotalMarkedQty,iTotalIssuedQty
	dim sIssEntryNo,iIssueEntNo,iItemEntNo
	dim sIssueCode,issueForCode,sQuery,sIssType,sMarkPackFlag
	dim sUnit,sUnitName,sType,sDate,sDept,sUsage,sRemarks,sAttID,sOptName,iPickNo,sScheduleNo,sScheduleDate
	Dim IssueToCode,sIssueToString,sIssueToSubCode,sIssueToType

    Response.Write "<font color=red>"
	
	bFlag = false
	iCtr = 0
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")	
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")	
	
	sOrg = session("organizationcode")
	sIssEntryNo = trim(Request("IssueNo"))
	sScheduleNo = trim(Request("ScheduleNo"))
	
	if trim(sScheduleNo)="" or isNull(sScheduleNo) then sScheduleNo = ""
	
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
'	sQuery ="Select Convert(varchar,IssueDate,103),IssuedForDescription,OrgUnitShortDescription,isNull(Remarks,''),"
'	sQuery = sQuery &" IssueType,MarkPackFlag,OrganisationCode,IsNull(IssueEntryCode,IssueEntryNo),M.IssuedToCode "
'	sQuery = sQuery &" from INV_T_MaterialIssueHeader M,INV_M_IssuedFor I,DCS_OrganizationUnitDefinitions U "
'	sQuery = sQuery &" where M.OrganisationCode = U.OUDefinitionID and M.IssuedToCode = I.IssuedForCode and IssueEntryNo="& sIssEntryNo
	
    sQuery ="Select Convert(varchar,IssueDate,103),OrgUnitShortDescription,isNull(Remarks,''),"
	sQuery = sQuery &" IssueType,MarkPackFlag,OrganisationCode,IsNull(IssueEntryCode,IssueEntryNo),M.IssuedToCode,M.IssuedToSubCode,M.IssuedToType "
	sQuery = sQuery &" from INV_T_MaterialIssueHeader M,DCS_OrganizationUnitDefinitions U "
	sQuery = sQuery &" where M.OrganisationCode = U.OUDefinitionID and IssueEntryNo="& sIssEntryNo
	
	'response.write sQuery
	dcrs.open sQuery,con
	if not dcrs.EOF then
		sDate = trim(dcrs(0))
		sUnitName = trim(dcrs(1))
		sRemarks = trim(dcrs(2))
		sIssType = trim(dcrs(3))
		sMarkPackFlag = trim(dcrs(4))
		sUnit = trim(dcrs(5))
		sIssueCode = trim(dcrs(6))
		IssueToCode = trim(dcrs(7))
		sIssueToSubCode = trim(dcrs(8))
		sIssueToType = trim(dcrs(9))
	end if
	dcrs.close
	
	
	sIssueToString = IssuedToString(sIssueToType,IssueToCode,sIssueToSubCode)
	
	if trim(sScheduleNo) <>"" then
	    SQuery = "Select Convert(varchar,ScheduledOn,103) from Inv_T_IssueForPickSchedule where IssueEntryNo = "& sIssEntryNo &" and ScheduleNo ="&  sScheduleNo
	    dcrs.open sQuery,con
	    if not dcrs.eof then
	        sScheduleDate = trim(dcrs(0))
	    end if
	    dcrs.close
	end if
%>

 
<script type="application/xml" data-itms-xml-island="1" id="ItemData">
<Pick UNIT="<%=sOrg%>" TOT="" ISSUENO="<%=sIssEntryNo%>">
<%

sQuery = "Select ClassificationCode,ItemCode,IsNull(SUM(QuantityForPick),0),IsNull(SUM(QuantityPicked),0),"
sQuery = sQuery & "Convert(varchar,IssueDate,103),ItemAttributes,IssueEntryNo,ItemEntryNo from VW_INV_IssuedForPick where IssueEntryNo = "& sIssEntryNo
sQuery = sQuery & " GROUP BY CLASSIFICATIONCODE,ITEMCODE,CONVERT(varchar,IssueDate,103),ItemAttributes,IssueEntryNo,ItemEntryNo"

dcrs.open sQuery,con
if not dcrs.EOF then
	Do While Not dcrs.EOF
	    sAttID = dcrs(5)
	    iPickNo = dcrs(6)
		if Trim(sIssEntryNo) <> "" then
			iSchQty = cdbl(dcrs(2)) - cdbl(dcrs(3))
			if iSchQty > 0 then
				Response.Write "<ITM CLACODE="""&trim(dcrs(0))&""" ITMCODE="""&trim(dcrs(1))&""" QTYFORISS="""&iSchQty&""" ISSQTY="""" ATTID="""&sAttID&""" PICKNO="""&iPickNo&""" ItemEntNo="""&trim(dcrs(7))&""" />" & vbCrLf
			end if
			
			
				 
		else
			sQuery = "SELECT REQUIREDBY,REQUIREDVALUE FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & sMRNo & " AND CLASSIFICATIONCODE = " & trim(dcrs(0)) & " AND ITEMCODE = " & trim(dcrs(1)) & ""
			if trim(sAttID)<>"0" and trim(sAttID)<>"" then
				sQuery = sQuery &" and ItemAttributes = "& sAttID
			end if
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				sReqBy = trim(dcrs1(0))
				sReqValue = trim(dcrs1(1))
			end if
			dcrs1.Close
		
			if sReqBy = "S" then
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ISNULL(SUM(MARKEDQTY),0),ISNULL(SUM(ISSUEDQTY),0) FROM INV_T_MRSITEMSCHEDULES WHERE SCHEDULETYPE = 'D' AND MRSNUMBER = " & sMRNo & " AND CLASSIFICATIONCODE = " & trim(dcrs(0)) & " AND ITEMCODE = " & trim(dcrs(1)) & " AND CONVERT(DATETIME,SCHEDULEDON,103) <= CONVERT(DATETIME," & Pack(FormatDate(date())) & ",103)"
					.ActiveConnection = con
					.Open
				end with

				set dcrs1.ActiveConnection = nothing
				if not dcrs1.EOF then
					iSchQty = cdbl(dcrs1(0)) - cdbl(dcrs1(1))
				end if
				dcrs1.Close
				if iSchQty > 0 then
					Response.Write "<ITM CLACODE="""&trim(dcrs(0))&""" ITMCODE="""&trim(dcrs(1))&""" QTYFORISS="""&iSchQty&""" ISSQTY="""" ATTID="""&sAttID&""" PICKNO="""&iPickNo&""" ItemEntNo="""&trim(dcrs(7))&"""/>" & vbCrLf
				end if
			elseif sReqBy = "I" or sReqBy = "D" then
				if DateDiff("d",FormatDate(sReqValue),date()) >= 0 then
					iSchQty = cdbl(dcrs(2)) - cdbl(dcrs(3))
					if iSchQty > 0 then
						Response.Write "<ITM CLACODE="""&trim(dcrs(0))&""" ITMCODE="""&trim(dcrs(1))&""" QTYFORISS="""&iSchQty&""" ISSQTY="""" ATTID="""&sAttID&""" PICKNO="""&iPickNo&""" ItemEntNo="""&trim(dcrs(7))&"""/>" & vbCrLf
					end if
				end if
			elseif sReqBy = "W" then
				if DateDiff("d",DateAdd("d",cdbl(sReqValue),FormatDate(dcrs(4))),date()) >= 0 then
					iSchQty = cdbl(dcrs(2)) - cdbl(dcrs(3))
					if iSchQty > 0 then
						Response.Write "<ITM CLACODE="""&trim(dcrs(0))&""" ITMCODE="""&trim(dcrs(1))&""" QTYFORISS="""&iSchQty&""" ISSQTY="""" ATTID="""&sAttID&""" PICKNO="""&iPickNo&""" ItemEntNo="""&trim(dcrs(7))&"""/>" & vbCrLf
					end if
				end if
			end if		
			
		end if
	dcrs.MoveNext
	Loop
end if
dcrs.Close

%>
</Pick>
</script>
<script type="application/xml" data-itms-xml-island="1" id="SalesInvoice"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="POrder"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="GatePass"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="ConfData"><Root/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
	Function GetPick(sClass,sItem,iQty,sIssEntryNo,sAttID,sItemEntNo)
		sTempValues = sClass&"|"&sItem&"|"&iQty&"|"&sIssEntryNo&"|"&sAttID&"|"&sItemEntNo
		
		set OutValue = showModalDialog("mrsIssuePickDetailsEntry.asp?sTemp="&sTempValues,ItemData,"dialogHeight:420px;dialogWidth:680px;center:Yes;help:No;resizable:No;status:No")
		'window.open "mrsIssuePickDetailsEntry.asp?sTemp="&sTempValues,"ItemData","",""
		sExp ="//ITM [ @CLACODE = """&sClass&""" and @ITMCODE = """&sItem&""" and @ItemEntNo = """&sItemEntNo&""" ]/PickDet"
		Set RootO = ItemData.documentElement
		Set ItemNode = RootO.Selectnodes(sExp)
	'	alert(RootO.xml)
	'	alert(ItemNode.length)

		if ItemNode.Length > 0 then
			set Q = eval("document.formname.txtQtyPA"&sClass&"A"&sItem&"A"&sIssEntryNo&"A"&sItemEntNo)
			Q.value = ItemNode.Item(0).Attributes.getNamedItem("TOT").Value
		else
			set Q = eval("document.formname.txtQtyPA"&sClass&"A"&sItem&"A"&sIssEntryNo&"A"&sItemEntNo)
			Q.value = "0"
		end if		
	end Function
	
	Function CheckSubmit()
		dim iCounter, sTempValues,sOrgCode
		document.formname.IssDate.value = document.formname.ctlPickDate.GetDate()  
		document.formname.RecBy.value = document.formname.txtRecdBy.value      
		sOrgCode =document.formname.hUnit.value
		
		sTempValues = ""
		sTempValues = document.formname.IssDate.value & ":" & document.formname.RecBy.value
		sExp ="//ITM"
		Set RootO = ItemData.documentElement
		Set ItemNode = RootO.Selectnodes(sExp)

		For iCounter = 0 to ItemNode.Length - 1
			sClass = ItemNode.Item(iCounter).Attributes.getNamedItem("CLACODE").Value 
			sItem = ItemNode.Item(iCounter).Attributes.getNamedItem("ITMCODE").Value 
			sPickNo = ItemNode.Item(iCounter).Attributes.getNamedItem("PICKNO").Value
			iItemEntNo = ItemNode.Item(iCounter).Attributes.getNamedItem("ItemEntNo").Value
			set Itm = eval("document.formname.txtQtyPA"&sClass&"A"&sItem&"A"&sPickNo&"A"&iItemEntNo)
			ItemNode.Item(iCounter).Attributes.getNamedItem("ISSQTY").Value = Itm.value
			'ItemNode.Item(iCounter).setAttribute "ISSDATE", document.formname.IssDate.value			
		next

		'alert ItemData.xml
		
		RootO.setAttribute "POConfirm","N"
		RootO.setAttribute "SInvConfirm","N"
		RootO.setAttribute "Invoice","A"
		RootO.setAttribute "GPConfirm","N"
		RootO.setAttribute "ProConfirm","N"
		RootO.setAttribute "PickDate",document.formname.ctlPickDate.GetDate()
	
	if 1=2 then
	
		if trim(document.formname.hIssForCode.value)="SUB" then
		
			set outValue = showModalDialog("../../Inventory/Transaction/IssSubConPurDetPop.asp?sHead=Purchase Order Creation&OrgCode="&sOrgCode,POrder,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			POrder.LoadXML(outValue.xml)
			
			set RootSub = POrder.documentElement
			if RootSub.hasChildNodes() then
				for each subNode in RootSub.childNodes
					if strcomp(subNode.nodeName,"PURACC")=0 then
						RootO.appendChild subNode
					end if
				next
			end if 'if RootSub.hasChildNodes() then
			
			set outValue = showModalDialog("IssSubConSalnvPop.asp?sHead=Proforma Invoice Creation&OrgCode="&sOrgCode,SalesInvoice,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			SalesInvoice.LoadXML(outValue.xml)
			
			set RootSub = SalesInvoice.documentElement
			if RootSub.hasChildNodes() then
				for each subNode in RootSub.childNodes
					if strcomp(subNode.nodeName,"SALINV")=0 then
						RootO.appendChild subNode
						sInvType = Subnode.getAttribute("InvType")
					end if
				next
			end if
			
			set outValue = showModalDialog("CommonConfirmPop.asp?sHead=Proforma Invoice Creation&CallFrom=SUB","","dialogHeight:180px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No")
			ConfData.LoadXML(outValue.xml)
			ProformaConfirm = outValue.getAttribute("ProInv")
			
			set RootSub = ConfData.documentElement
			RootO.setAttribute "Invoice","P"
			RootO.setAttribute "POConfirm",RootSub.getAttribute("Confirm")
			RootO.setAttribute "ProConfirm",ProformaConfirm
			
			sPOConfirm = RootSub.getAttribute("Confirm")
			
			
		elseif trim(document.formname.hIssForCode.value)="DIS" then
			set outValue = showModalDialog("CommonConfirmPop.asp?sHead=Sales Invoice Creation&CallFrom=DIS","","dialogHeight:230px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No")
			'alert(outValue.xml)
			SalesInvoice.LoadXML(outValue.xml)
			
			set RootSub = SalesInvoice.documentElement
			RootO.setAttribute "SInvConfirm",RootSub.getAttribute("Confirm")
			RootO.setAttribute "Invoice",RootSub.getAttribute("Invoice")
			sInvoiceSelected = RootSub.getAttribute("Invoice")
			sPOConfirm = RootSub.getAttribute("Confirm")
			
			set outValue = showModalDialog("../../Inventory/Transaction/IssSubConSalnvPop.asp?sHead=Sales Invoice Creation&OrgCode="&sOrgCode,SalesInvoice,"dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			SalesInvoice.LoadXML(outValue.xml)
			'alert(outValue.xml)
			set RootSub = SalesInvoice.documentElement
		    if RootSub.hasChildNodes() then
				for each subNode in RootSub.childNodes
					if strcomp(subNode.nodeName,"SALINV")=0 then
						RootO.appendChild subNode
						sInvType = Subnode.getAttribute("InvType")
					end if
				next
			end if
		elseif trim(document.formname.hIssForCode.value)="SER" then
			set outValue = showModalDialog("CommonConfirmPop.asp?sHead=Gate Pass Creation&CallFrom=SER","","dialogHeight:180px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No")
			'alert(outValue.xml)
			GatePass.LoadXML(outValue.xml)
			
			set RootSub = GatePass.documentElement
			RootO.setAttribute "Confirm","Y"
			RootO.setAttribute "POConfirm","N"
			RootO.setAttribute "SInvConfirm","N"
			RootO.setAttribute "GPConfirm",RootSub.getAttribute("Confirm")
			sPOConfirm = RootSub.getAttribute("Confirm")
			
			set ndServices = GatePass.createElement("SERVICES")
			ndServices.setAttribute "Transport",""
			ndServices.setAttribute "TakenBy",""
			ndServices.setAttribute "DelivertyBy",""
			ndServices.setAttribute "Remarks",""
			RootO.appendChild ndServices
			
		end if
		
	end if 'if 1=2 then
		
		Set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","XMLSave.asp?SessionFlag=true&Name=IssuePick_",false
		objhttp.send ItemData.XMLDocument
		
		sScheduleNo = document.formname.hScheduleNo.value
		
		document.formname.action = "mrsIssuePickInsert.asp?TEMPVALUES="&sTempValues&"&ScheduleNo="&sScheduleNo
		document.formname.submit
		
	'	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	'	objhttp.Open "POST","mrsIssuePickInsert.asp?TEMPVALUES="&sTempValues , false
	'	objhttp.send ItemData.XMLDocument
	'	sTempArr = trim(objhttp.responseText)
	 '   sTempValue = split(sTempArr,"@")
	
	
	
	if 1= 2 and trim(objhttp.responseText)<>"" then
		if len(sTempValue(0))>0 then
			alert(objhttp.responseText)
			'document.formname.issue.disabled = true
			exit function
		end if
	
		sTempInvArr = split(sTempValue(1),":")
		sForInvNo = sTempInvArr(1)
		if trim(sForINvNo) <>"" then
			if trim(document.formname.hIssForCode.value)="SUB" then
				if trim(sPOConfirm)="Y" then
					set objhttp2 = CreateObject("Microsoft.XMLHTTP")
					objhttp2.Open "POST","../../Purchase/Transaction/POGenerationInsert.asp",false
					objhttp2.send 
					if trim(objhttp2.responseText)<>"" then
						alert(objhttp2.responseText)
						exit function
					else
						if trim(ProformaConfirm)="Y" then
							window.location.href ="../../Sales/Transaction/SalTrProInvoice.asp?hInvNo="&sForINvNo&"&InvType="&sInvType&"&CallFrom=MRISSUE&AppRefNo="&sTempInvArr(2)&"&AppRefType=12"
						else
							window.location.href = "ISSUEMGMT.asp"
						end if
					end if
				end if
			elseif trim(document.formname.hIssForCode.value)="DIS" then
				if trim(sPOConfirm)="Y" then
					if trim(sInvoiceSelected)="A" then
						window.location.href ="../../Sales/Transaction/SalInvoiceEntry.asp?hInvNo="&sForInvNo&"&InvType="&sInvType&"&CallFrom=MRISSUE&AppRefNo="&sTempInvArr(2)&"&AppRefType=12"
					elseif trim(sInvoiceSelected)="P" then
						window.location.href ="../../Sales/Transaction/SalTrProInvoice.asp?hInvNo="&sForInvNo&"&InvType="&sInvType&"&CallFrom=MRISSUE&AppRefNo="&sTempInvArr(2)&"&AppRefType=12"
					end if
					exit function
				end if
			elseif trim(document.formname.hIssForCode.value)="SER" then
				if trim(sPOConfirm)="Y" then
					window.location.href ="GatePassServiceEntryAmd.asp?GatePassNo="&sForInvNo
					exit function
				end if
			end if
		else
			Msgbox ("Issue Pick has be done")
			window.location.href = "ISSUEMGMT.asp"
		end if ' if trim(sForINvNo) <>"" then
'	else
'		Msgbox ("Issue Pick has be done")
'			window.location.href = "ISSUEMGMT.asp"
	end if
	end Function
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action>
<input type="hidden" name="IssDate" value="">
<input type="hidden" name="RecBy" value="">
<input type="hidden" name="hIssForCode" value="<%=issueForCode%>">
<input type="hidden" name="hUnit" value="<%=sUnit%>">
<input type="hidden" name="hScheduleNo" value="<%=sScheduleNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class="PageTitle" height="20">
			<p align="center">Pick Item List
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
					<td class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">						
						    <tr>
						        <td align="center">
						        </td>
								<td width="100%" colspan="4">
								    <div align="left">
									    <table border="0" cellspacing="0" cellpadding="0" >
								            <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
							                        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
						                        </td>
					                        </tr>
									        <tr>
											    <td class="FieldCellSub">Issue Entry Number&nbsp;</td>
											    <td class="FieldCell"><span class="DataOnly"><%=sIssueCode%>&nbsp;</span></td>
											    <td class="FieldCellSub">Issue Date</td>
											    <td class="FieldCell"><span class="DataOnly"><%=sDate%>&nbsp;</span></td>
											</tr>
											<tr>
											    <td class="FieldCellSub">Requisition by Unit&nbsp;</td>
											    <td class="FieldCell"><span class="DataOnly"><%=sUnitName%>&nbsp;</span></td>
											    <td class="FieldCellSub">Issued To</td>
											    <td class="FieldCell"><span class="DataOnly"><%=sIssueToString%>&nbsp;</span></td>
											</tr>
											<tr>
											    <td class="FieldCellSub" width="45">Date</td>
											    <td class="FieldCell" width="45">
												    <%
													    ' Function Call to Insert Date Picker
													    Response.Write InsertDatePicker("ctlPickDate")
												    %>
											    </td>												
										    </tr>																						
										    <tr>
											    <td class="FieldCellSub">Material Received By</td>
											    <td class="FieldCellSub" width="85">
												    <input type="text" name="txtRecdBy" size="40" value="" class="FormElem" maxlength=30 style="text-align:Left">
											    </td>																		
										    </tr>	
										    <%if trim(sScheduleNo) <>"" then %>
										     <tr>
											    <td class="FieldCellSub">Schedule On</td>
											    <td class="FieldCellSub" width="85">
												    <span id="spanScheduleDate" class="dataonly"><%=sScheduleDate%></span>
											    </td>																		
										    </tr>							
										    <%end if 'if trim(sScheduleNo) <>"" then %>															
									    </table>
									</div>
								</td>
							</tr>
							<tr>
								<td align="center" colspan="5" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr></tr>
							<tr>
								<td align="center">
								</td>
								<td width="100%" colspan="4">
								    <div style="width:90%;">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%"> 
										    <tr>
											    <td>
												    <div class="frmBody" id="frm3" style="width: 100%; height:390">
													    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
														    <tr>
															    <td class="ExcelHeaderCell" align="center" width="30">S.No.</td>
															    <td class="ExcelHeaderCell" align="center">Item Description</td>
															    <td class="ExcelHeaderCell" align="center">Quantity Marked</td>
															    <td class="ExcelHeaderCell" align="center">Quantity Issued</td>
															    <td class="ExcelHeaderCell" align="center">Quantity Available for Pick</td>
															    <td class="ExcelHeaderCell" align="center">Quantity Picked</td>
														    </tr>
													    <%
													        Response.Write "<font color=red>"
														    iSchQty = "0"
    														if trim(sScheduleNo)<>"" then
    														    sQuery = "Select ClassificationCode,ItemCode,GroupName,ItemDescription,IsNull(SUM(ScheduledQty),0), "
														        sQuery = sQuery & "IsNull(SUM(PickedQty),0),Convert(Varchar,IssueDate,103),IsNull(ItemAttributes,''),IssueEntryNo,ItemEntryNo "
														        sQuery = sQuery & " from VW_INV_IssuedForPickSchedule where (ISNULL(ScheduledQty,0) - ISNULL(PickedQty,0) > 0) and IssueEntryNo = "& sIssEntryNo &" and ScheduleNo = "& sScheduleNO
														        sQuery = sQuery & "  GROUP BY CLASSIFICATIONCODE,ITEMCODE,GROUPNAME,ItemDescription,CONVERT(Varchar,IssueDate,103),ItemAttributes,IssueEntryNo,ItemEntryNo"
    														else
														        sQuery = "Select ClassificationCode,ItemCode,GroupName,ItemDescription,IsNull(SUM(QuantityForPick),0), "
														        sQuery = sQuery & "IsNull(SUM(QuantityPicked),0),Convert(Varchar,IssueDate,103),IsNull(ItemAttributes,''),IssueEntryNo,ItemEntryNo "
														        sQuery = sQuery & " from VW_INV_IssuedForPick where (ISNULL(QuantityForPick,0) - ISNULL(QuantityPicked,0) > 0) and IssueEntryNo = "& sIssEntryNo
														        sQuery = sQuery & "  GROUP BY CLASSIFICATIONCODE,ITEMCODE,GROUPNAME,ItemDescription,CONVERT(Varchar,IssueDate,103),ItemAttributes,IssueEntryNo,ItemEntryNo"
														    end if
														    dcrs.open sQuery,con
														    if not dcrs.EOF then
															    set sClass = dcrs(0)
															    set sItem = dcrs(1)
															    set sClassName = dcrs(2)
															    set sItemName = dcrs(3)  
															    set iQtyToIssue = dcrs(4)
															    set iQtyIssued = dcrs(5)
															    set dMRDate = dcrs(6)
															    set sAttID = dcrs(7)
															    set iIssueEntNo = dcrs(8)
															    set iItemEntNo =  dcrs(9)
															    Do While Not dcrs.EOF
																    bFlag = false
																    sItemName = ItemDisplay(trim(sItem),trim(sClass))
    																
																    sQuery = "Select IsNull(SUM(QuantityForPick),0),IsNull(SUM(QuantityPicked),0) from VW_INV_IssuedForPick "
																    sQuery = sQuery & " where IssueEntryNo = "& iIssueEntNo &" and ClassificationCode = " & trim(sClass) & " and ItemCode = " & trim(sItem) & ""
																    if Trim(sAttID)<>"" and Trim(sAttID)<>"0" then
																        sQuery = sQuery &" and ItemAttributes in ("& sAttID &")"
																    end if
																    'response.write "<textarea>"&sQuery&"</textarea>"
																    dcrs1.open sQuery,con
																    if not dcrs1.EOF then
																	    iTotalMarkedQty = trim(dcrs1(0))
																	    iTotalIssuedQty = trim(dcrs1(1))
																    end if
																    dcrs1.Close
																    
																    'Response.write "<textarea>"&iTotalMarkedQty&"</textarea>"
    																
																    if trim(sAttID)<>"" and Trim(sAttID)<>"0" then
																        sQuery = "Select OptionName from INV_M_ITEMTYPEOPTIONS where OptionValue = "& sAttID
																        dcrs1.Open sQuery,con
																        if not dcrs1.EOF then
																            sOptName = " ["&trim(dcrs1(0))&"]"
																        end if
																        dcrs1.Close 
																    end if
																    if trim(sOptName)<>"" then
																        sItemName = sItemName & sOptName
																    end if

																    if Trim(sIssEntryNo) <> "" then
																	    sReqBy = "I"
																	    iSchQty = cdbl(iQtyToIssue) - cdbl(iQtyIssued)
																	    bFlag = true
																    else
    																
																	    with dcrs1
																		    .CursorLocation = 3
																		    .CursorType = 3
																		    .Source = "SELECT REQUIREDBY,REQUIREDVALUE FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & sMRNo & " AND CLASSIFICATIONCODE = " & trim(dcrs(0)) & " AND ITEMCODE = " & trim(dcrs(1)) & ""
																		    .ActiveConnection = con
																		    .Open
																	    end with

																	    set dcrs1.ActiveConnection = nothing
																	    if not dcrs1.EOF then
																		    sReqBy = trim(dcrs1(0))
																		    sReqValue = trim(dcrs1(1))
																	    end if
																	    dcrs1.Close
																    end if
    																		
																    if sReqBy = "S" then
																	    with dcrs1
																		    .CursorLocation = 3
																		    .CursorType = 3
																		    .Source = "SELECT ISNULL(SUM(MARKEDQTY),0),ISNULL(SUM(ISSUEDQTY),0) FROM INV_T_MRSITEMSCHEDULES WHERE SCHEDULETYPE = 'D' AND MRSNUMBER = " & sMRNo & " AND CLASSIFICATIONCODE = " & trim(sClass) & " AND ITEMCODE = " & trim(sItem) & " AND CONVERT(DATETIME,SCHEDULEDON,103) <= CONVERT(DATETIME," & Pack(FormatDate(date())) & ",103)"
																		    .ActiveConnection = con
																		    .Open
																	    end with
																	    set dcrs1.ActiveConnection = nothing
																	    if not dcrs1.EOF then
																		    iSchQty = cdbl(dcrs1(0)) - cdbl(dcrs1(1))
																		    bFlag = true
																	    end if
																	    dcrs1.Close
																    elseif sReqBy = "I" or sReqBy = "D" then
																	    'Response.Write sReqValue & "<BR>"
																	    'Response.Write date() & "<BR>"
																	    'Response.Write DateDiff("d",FormatDate(sReqValue),date()) & "<BR>"
																	    if DateDiff("d",FormatDate(sReqValue),date()) >= 0 then
																		    iSchQty = cdbl(iQtyToIssue) - cdbl(iQtyIssued)
																		    bFlag = true
																	    end if
																    elseif sReqBy = "W" then
																	    if DateDiff("d",DateAdd("d",cdbl(sReqValue),FormatDate(dMRDate)),date()) >= 0 then
																		    iSchQty = cdbl(iQtyToIssue) - cdbl(iQtyIssued)
																		    bFlag = true
																	    end if
																    end if	
																    if cdbl(iSchQty) > 0 then
																	    if bFlag then	
																		    iCtr = iCtr + 1

													    %>

														    <tr>
															    <td class="ExcelSerial" align="center" width="30"><%=iCtr%></td>
															    <td class="ExcelDisplayCell"><%=sItemName%></td>
															    <td class="ExcelDisplayCell" align="right" width="10">
																    <input type="text" name="txtQtyTMA<%=trim(sClass)%>A<%=trim(sItem)%>" size="12" value="<%=iTotalMarkedQty%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
															    </td>
															    <td class="ExcelDisplayCell" align="right" width="10">
																    <input type="text" name="txtQtyIA<%=trim(sClass)%>A<%=trim(sItem)%>" size="12" value="<%=iTotalIssuedQty%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
															    </td>
															    <td class="ExcelDisplayCell" align="center" width="90">
																    <input type="text" name="txtQtyMA<%=trim(sClass)%>A<%=trim(sItem)%>" size="12" value="<%=iSchQty%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
																    <a href="#">
																	    <img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" alt="Enter Pick Details for <%=trim(sItemName)%>" width="11" height="12" style="cursor: hand" onClick="GetPick('<%=trim(sClass)%>','<%=trim(sItem)%>','<%=iSchQty%>','<%=sIssEntryNo%>','<%=sAttID%>','<%=iItemEntNo%>')">
																    </a>
															    </td>
															    <td class="ExcelDisplayCell" align="right" width="10">
																    <input type="text" name="txtQtyPA<%=trim(sClass)%>A<%=trim(sItem)%>A<%=iIssueEntNo%>A<%=iItemEntNo%>" size="12" value="0" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
															    </td>
														    </tr>
													    <%	
																	    end if
																    end if
															    dcrs.MoveNext
															    loop
														    end if
														    dcrs.Close
													    %>

													    </table>
												    </div>
											    </td>
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
												<input type="button" value="Issue" name="issue" class="ActionButton" onClick="CheckSubmit()">
 												<input type="button" value="Cancel" name="cancel" class="ActionButton" onClick="Cancel('mrsIssuePickListEntry.asp')">
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
</body>
</html>
